import JSZip from "jszip";
import * as mammoth from "mammoth";
import { LossItem, scanLosses } from "./losses";
import { buildStyleMap } from "./styles";
import { countMatches, decodeXml } from "./xml";

export const IMAGE_TOKEN_PREFIX = "sun-daoru-image-";

export interface PendingImage {
  token: string;
  data: ArrayBuffer;
  contentType: string;
  extension: string;
  previewable: boolean;
  altText?: string;
  link?: string;
}

export interface DocxConversion {
  html: string;
  images: PendingImage[];
  warnings: string[];
  losses: LossItem[];
  /** 正文 XML，导入结束后还要拿它做逐段文字对账 */
  documentXml: string;
  styleMapEntries: number;
  headingLevelsShifted: boolean;
  stats: {
    headings: number;
    paragraphs: number;
    tables: number;
    images: number;
  };
}

export interface DocxConvertOptions {
  normalizeHeadingLevels?: boolean;
}

const INTERNAL_ANCHOR = /href="#([^"]*)"/g;
/** mammoth 自己给脚注/尾注生成的锚点，不是文档书签 */
const MAMMOTH_ANCHOR = /^(?:footnote|endnote)(?:-ref)?-\d+$/;

export async function convertDocx(
  data: ArrayBuffer,
  options: DocxConvertOptions = {},
): Promise<DocxConversion> {
  assertZip(data);

  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(data);
  } catch {
    throw new Error("无法读取文件：它可能已加密、损坏，或不是真正的 .docx 文件。");
  }

  const documentFile = zip.file("word/document.xml");
  if (!documentFile) {
    throw new Error("这个文件里没有 Word 正文，可能不是有效的 .docx 文档。");
  }

  const documentXml = await documentFile.async("string");
  const stylesFile = zip.file("word/styles.xml");
  const stylesXml = stylesFile ? await stylesFile.async("string") : "";

  const styleMap = buildStyleMap(
    stylesXml,
    documentXml,
    Boolean(options.normalizeHeadingLevels),
  );
  const images: PendingImage[] = [];
  const warnings: string[] = [];
  const imageReadFailures: number[] = [];

  let imageIndex = 0;
  const result = await mammoth.convertToHtml(
    nodeOrBrowserInput(data),
    {
      styleMap: styleMap.entries,
      ignoreEmptyParagraphs: true,
      convertImage: mammoth.images.imgElement(async (image) => {
        const index = imageIndex++;
        const token = `${IMAGE_TOKEN_PREFIX}${index}`;
        try {
          const buffer = await image.readAsArrayBuffer();
          const format = imageFormat(image.contentType);
          // mammoth 的声明里没有 altText，但运行时确实会把 wp:docPr 的 descr 传进来
          const altText = String((image as { altText?: unknown }).altText || "").trim();
          images.push({
            token,
            data: buffer,
            contentType: image.contentType || "application/octet-stream",
            extension: format.extension,
            previewable: format.previewable,
            altText: altText || undefined,
          });
          return { src: token };
        } catch {
          imageReadFailures.push(index + 1);
          return { src: "" };
        }
      }),
    },
  );

  if (!result.value.trim()) {
    throw new Error("未能从文档中读取正文内容。");
  }

  const html = result.value;

  const internalLinks = Array.from(html.matchAll(INTERNAL_ANCHOR))
    .map((match) => decodeXml(match[1]))
    .filter((name) => name && !MAMMOTH_ANCHOR.test(name)).length;

  const losses = await scanLosses(zip, documentXml, { html, internalLinks });

  // 链接到文件的图片已经在丢失清单里单独报了，这里只报剩下的读取失败，避免重复
  const externalImages =
    losses.find((loss) => loss.label === "链接到文件的图片")?.count ?? 0;
  const extraFailures = Math.max(0, imageReadFailures.length - externalImages);
  if (extraFailures > 0) {
    warnings.push(`有 ${extraFailures} 张图片读不出内容，已跳过。`);
  }

  const seenMessages = new Set<string>();
  for (const message of result.messages) {
    const text = String(message.message || "").trim();
    if (!text || seenMessages.has(text)) continue;
    if (/unrecognised element was ignored/i.test(text)) continue;
    if (!/image|picture|drawing/i.test(text)) continue;
    const unlikely = /Image of type ([^ ]+) is unlikely to display/i.exec(text);
    if (unlikely && imageFormat(unlikely[1]).previewable) continue;
    seenMessages.add(text);
    warnings.push(text);
  }

  return {
    html,
    images,
    warnings,
    losses,
    documentXml,
    styleMapEntries: styleMap.entries.length,
    headingLevelsShifted: styleMap.shifted,
    stats: {
      headings: countMatches(html, /<h[1-6]\b/g),
      paragraphs: countMatches(html, /<p\b/g),
      tables: countMatches(html, /<table\b/g),
      images: images.length,
    },
  };
}

function assertZip(data: ArrayBuffer): void {
  const bytes = new Uint8Array(data);
  if (bytes.length < 4) {
    throw new Error("文件为空或内容不完整。");
  }
  const isZip = bytes[0] === 0x50 && bytes[1] === 0x4b;
  if (isZip) return;

  const isOle =
    bytes[0] === 0xd0 &&
    bytes[1] === 0xcf &&
    bytes[2] === 0x11 &&
    bytes[3] === 0xe0;
  if (isOle) {
    throw new Error(
      "无法读取：文件可能已加密，或是旧版 .doc 格式。请取消密码并在 Word/WPS 中另存为 .docx 后重试。",
    );
  }
  throw new Error("无法读取：这不是有效的 .docx 文件。");
}

function nodeOrBrowserInput(
  data: ArrayBuffer,
): { buffer: Buffer } | { arrayBuffer: ArrayBuffer } {
  if (typeof window === "undefined" && typeof Buffer !== "undefined") {
    return { buffer: Buffer.from(data) };
  }
  return { arrayBuffer: data };
}

function imageFormat(contentType: string): {
  extension: string;
  previewable: boolean;
} {
  const type = (contentType || "").toLowerCase().split(";")[0].trim();
  switch (type) {
    case "image/png":
    case "image/x-png":
      return { extension: "png", previewable: true };
    case "image/jpeg":
    case "image/jpg":
    case "image/pjpeg":
      return { extension: "jpg", previewable: true };
    case "image/gif":
      return { extension: "gif", previewable: true };
    case "image/webp":
      return { extension: "webp", previewable: true };
    case "image/svg+xml":
      return { extension: "svg", previewable: true };
    case "image/bmp":
    case "image/x-ms-bmp":
      return { extension: "bmp", previewable: true };
    case "image/tiff":
    case "image/x-tiff":
      return { extension: "tiff", previewable: false };
    case "image/x-emf":
    case "image/emf":
      return { extension: "emf", previewable: false };
    case "image/x-wmf":
    case "image/wmf":
      return { extension: "wmf", previewable: false };
    default: {
      const subtype = type.split("/")[1]?.replace(/[^a-z0-9]+/g, "") || "bin";
      return { extension: subtype.slice(0, 10) || "bin", previewable: false };
    }
  }
}
