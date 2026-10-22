import JSZip from "jszip";
import * as mammoth from "mammoth";

export const IMAGE_TOKEN_PREFIX = "sun-daoru-image-";

const EMU_PER_PIXEL = 9525;
const MAX_IMAGE_WIDTH_PX = 1600;

export interface PendingImage {
  token: string;
  data: ArrayBuffer;
  contentType: string;
  extension: string;
  previewable: boolean;
  widthPx?: number;
  link?: string;
}

export interface DocxConversion {
  html: string;
  images: PendingImage[];
  warnings: string[];
  styleMapEntries: number;
  stats: {
    headings: number;
    paragraphs: number;
    tables: number;
    images: number;
  };
}

interface StyleDef {
  id: string;
  name: string;
  outline?: number;
  basedOn?: string;
}

export async function convertDocx(data: ArrayBuffer): Promise<DocxConversion> {
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

  const styleMap = buildStyleMap(stylesXml, documentXml);
  const extents = extractImageExtents(documentXml);
  const images: PendingImage[] = [];
  const warnings: string[] = [];
  const imageReadFailures: number[] = [];

  let imageIndex = 0;
  const result = await mammoth.convertToHtml(
    nodeOrBrowserInput(data),
    {
      styleMap,
      ignoreEmptyParagraphs: true,
      convertImage: mammoth.images.imgElement(async (image) => {
        const index = imageIndex++;
        const token = `${IMAGE_TOKEN_PREFIX}${index}`;
        try {
          const buffer = await image.readAsArrayBuffer();
          const format = imageFormat(image.contentType);
          images.push({
            token,
            data: buffer,
            contentType: image.contentType || "application/octet-stream",
            extension: format.extension,
            previewable: format.previewable,
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

  if (imageReadFailures.length > 0) {
    warnings.push(
      `有 ${imageReadFailures.length} 张图片无法读取，已跳过（第 ${imageReadFailures.slice(0, 8).join("、")} 张${imageReadFailures.length > 8 ? "等" : ""}）。`,
    );
  }

  if (extents.length === images.length) {
    images.forEach((image, index) => {
      const width = Math.round(extents[index].cx / EMU_PER_PIXEL);
      if (width >= 16) image.widthPx = Math.min(width, MAX_IMAGE_WIDTH_PX);
    });
  }

  const seenMessages = new Set<string>();
  for (const message of result.messages) {
    const text = String(message.message || "").trim();
    if (!text || seenMessages.has(text)) continue;
    if (/unrecognised element was ignored/i.test(text)) continue;
    if (!/image|picture|drawing/i.test(text)) continue;
    seenMessages.add(text);
    warnings.push(text);
  }

  const html = result.value;
  return {
    html,
    images,
    warnings,
    styleMapEntries: styleMap.length,
    stats: {
      headings: (html.match(/<h[1-6]\b/g) || []).length,
      paragraphs: (html.match(/<p\b/g) || []).length,
      tables: (html.match(/<table\b/g) || []).length,
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

function buildStyleMap(stylesXml: string, documentXml: string): string[] {
  const styles = parseStyles(stylesXml);
  const usedIds = new Set<string>();
  for (const match of documentXml.matchAll(/<w:pStyle\b[^>]*\bw:val="([^"]*)"/g)) {
    usedIds.add(decodeXml(match[1]));
  }

  const entries: string[] = [];
  const seenNames = new Set<string>();
  const addEntry = (name: string, level: number) => {
    const cleanName = name.trim();
    if (!cleanName || cleanName.includes("'") || cleanName.includes("\\")) return;
    const key = cleanName.toLowerCase();
    if (seenNames.has(key)) return;
    seenNames.add(key);
    entries.push(`p[style-name='${cleanName}'] => h${level}:fresh`);
  };

  const candidates = usedIds.size > 0 ? Array.from(usedIds) : Array.from(styles.keys());
  for (const id of candidates) {
    const style = styles.get(id);
    if (!style) continue;
    const level = resolveLevel(style, styles, new Set());
    if (level) addEntry(style.name, level);
  }

  for (const [name, level] of FALLBACK_HEADING_STYLES) {
    addEntry(name, level);
  }

  return entries;
}

function parseStyles(xml: string): Map<string, StyleDef> {
  const styles = new Map<string, StyleDef>();
  if (!xml) return styles;

  for (const match of xml.matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
    const attrs = match[1];
    const body = match[2];
    const type = attrs.match(/\bw:type="([^"]*)"/)?.[1];
    if (type && type !== "paragraph") continue;

    const id = decodeXml(attrs.match(/\bw:styleId="([^"]*)"/)?.[1] || "");
    if (!id) continue;

    const name = decodeXml(xmlValue(body, "w:name") || "");
    const outlineRaw = xmlValue(body, "w:outlineLvl");
    const outline = outlineRaw === undefined ? undefined : Number(outlineRaw);
    const basedOn = decodeXml(xmlValue(body, "w:basedOn") || "") || undefined;

    styles.set(id, {
      id,
      name,
      outline: Number.isFinite(outline) ? outline : undefined,
      basedOn,
    });
  }

  return styles;
}

function xmlValue(xml: string, tag: string): string | undefined {
  const escapedTag = tag.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return new RegExp(`<${escapedTag}\\b[^>]*\\bw:val="([^"]*)"`).exec(xml)?.[1];
}

function resolveLevel(
  style: StyleDef,
  styles: Map<string, StyleDef>,
  visited: Set<string>,
): number | undefined {
  if (visited.has(style.id)) return undefined;
  visited.add(style.id);

  const byName = levelFromName(style.name);
  if (byName) return byName;

  if (style.outline !== undefined && style.outline >= 0 && style.outline <= 5) {
    return style.outline + 1;
  }

  if (style.basedOn) {
    const parent = styles.get(style.basedOn);
    if (parent) return resolveLevel(parent, styles, visited);
  }

  return undefined;
}

function levelFromName(name: string): number | undefined {
  const trimmed = name.trim();
  if (!trimmed) return undefined;
  if (/^(?:toc|table of contents|目录)/i.test(trimmed)) return undefined;
  if (/^(?:subtitle|副标题)$/i.test(trimmed)) return undefined;

  const heading = /^(?:heading|标题)\s*([1-6])(?=$|[_\-\s.])/i.exec(trimmed);
  if (heading) return Number(heading[1]);

  if (/^(?:title|文件标题)(?=$|[_\-\s.])/i.test(trimmed)) return 1;

  const chinese = /^(一级|二级|三级|四级|五级|六级)标题$/.exec(trimmed);
  if (chinese) {
    return { 一级: 1, 二级: 2, 三级: 3, 四级: 4, 五级: 5, 六级: 6 }[chinese[1]];
  }

  return undefined;
}

const FALLBACK_HEADING_STYLES: Array<[string, number]> = [
  ["Heading 1", 1],
  ["Heading 2", 2],
  ["Heading 3", 3],
  ["Heading 4", 4],
  ["Heading 5", 5],
  ["Heading 6", 6],
  ["Title", 1],
  ["文件标题", 1],
  ["标题 1", 1],
  ["标题 2", 2],
  ["标题 3", 3],
  ["标题 4", 4],
  ["标题 5", 5],
  ["标题 6", 6],
  ["标题1", 1],
  ["标题2", 2],
  ["标题3", 3],
  ["标题4", 4],
  ["标题5", 5],
  ["标题6", 6],
  ["一级标题", 1],
  ["二级标题", 2],
  ["三级标题", 3],
  ["四级标题", 4],
  ["五级标题", 5],
  ["六级标题", 6],
];

function extractImageExtents(xml: string): Array<{ cx: number; cy: number }> {
  const extents: Array<{ cx: number; cy: number }> = [];
  for (const match of xml.matchAll(/<wp:extent\b[^>]*>/g)) {
    const cx = Number(/\bcx="(\d+)"/.exec(match[0])?.[1]);
    const cy = Number(/\bcy="(\d+)"/.exec(match[0])?.[1]);
    if (Number.isFinite(cx) && cx > 0) extents.push({ cx, cy });
  }
  return extents;
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

function decodeXml(value: string): string {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}
