import JSZip from "jszip";

export type LossKind = "lost" | "downgraded" | "info";

export interface LossItem {
  kind: LossKind;
  label: string;
  count: number;
  detail: string;
}

export const FORMULA_PLACEHOLDER = "【公式未导入】";

export interface LossContext {
  /** mammoth 产出的 HTML，用来验证文本框文字是否真的进了正文 */
  html: string;
  /** 正文里指向文档内书签/标题的跳转数量 */
  internalLinks: number;
  /** 脚注/尾注正文；提前读取，避免转换前插入的公式标记影响清点 */
  noteParts: string[];
  /** 实际出现在笔记里的公式标记数量 */
  markedFormulas: number;
}

const TEXT_TAG = /<w:t(?:\s[^>]*)?>/g;
const TEXT_RUN = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
const DELETED_TEXT_RUN = /<w:delText(?:\s[^>]*)?>([\s\S]*?)<\/w:delText>/g;
const MATH_TEXT_RUN = /<m:t(?:\s[^>]*)?>([\s\S]*?)<\/m:t>/g;
const HEADER_FOOTER_PART = /^word\/(?:header|footer)\d*\.xml$/;

export async function scanLosses(
  zip: JSZip,
  documentXml: string,
  ctx: LossContext,
): Promise<LossItem[]> {
  const losses: LossItem[] = [];

  const bodyParts: string[] = [documentXml, ...ctx.noteParts];
  let headerFooterParts = 0;
  let headerFooterTexts = 0;
  let headerFooterImages = 0;

  for (const entry of Object.values(zip.files)) {
    if (entry.dir) continue;
    if (entry.name === "word/comments.xml") continue;
    if (HEADER_FOOTER_PART.test(entry.name)) {
      const xml = await entry.async("string");
      headerFooterParts++;
      headerFooterTexts += count(xml, TEXT_TAG);
      headerFooterImages += count(xml, /<a:blip\b/g);
    }
  }

  const body = bodyParts.join("\n");
  // 公式标记会插进文本框文字的中间，比对前先去掉，否则会把已导入的文本框误报成丢失
  const plainHtml = stripTags(ctx.html).split(FORMULA_PLACEHOLDER).join("");

  const commentsFile = zip.file("word/comments.xml");
  if (commentsFile) {
    const xml = await commentsFile.async("string");
    const total = count(xml, /<w:comment\b/g);
    if (total > 0) {
      losses.push({
        kind: "lost",
        label: "批注内容",
        count: total,
        detail: `批注框里的文字不导入（被批注的正文已保留）。例如：${preview(sample(xml, TEXT_RUN))}`,
      });
    }
  }

  const deleted = count(body, /<w:delText\b/g);
  if (deleted > 0) {
    losses.push({
      kind: "lost",
      label: "修订里被删掉的文字",
      count: deleted,
      detail: `文档还有未接受的修订，删除的部分不导入。例如：${preview(sample(body, DELETED_TEXT_RUN))}`,
    });
  }

  const formulas = count(body, /<m:oMath\b/g);
  if (formulas > 0) {
    const visibleBody = body.replace(
      /<w:(?:del|moveFrom)\b[\s\S]*?<\/w:(?:del|moveFrom)>/g,
      "",
    );
    const deletedFormulas = formulas - count(visibleBody, /<m:oMath\b/g);
    const markable = Math.max(0, formulas - deletedFormulas);
    const marked = Math.min(ctx.markedFormulas, markable);
    const markerNotes: string[] = [];
    if (marked > 0) {
      markerNotes.push(
        `已在笔记中对应位置标出 ${marked} 处（搜索「公式未导入」可定位）。`,
      );
    }
    if (markable > marked) {
      markerNotes.push(
        `另有 ${markable - marked} 处无法在正文中定位（可能在文本框或嵌入对象里）。`,
      );
    }
    if (deletedFormulas > 0) {
      markerNotes.push(
        `${deletedFormulas} 处位于已删除的修订里，不会出现在笔记中。`,
      );
    }
    losses.push({
      kind: "lost",
      label: "公式",
      count: formulas,
      detail: `公式对象整体不导入，含其中的文字。${markerNotes.join("")}例如：${preview(sample(body, MATH_TEXT_RUN))}`,
    });
  }

  const textboxLost: string[] = [];
  let textboxTotal = 0;
  for (const match of documentXml.matchAll(/<w:txbxContent\b[\s\S]*?<\/w:txbxContent>/g)) {
    textboxTotal++;
    const text = sample(match[0], TEXT_RUN, 80);
    if (text && !plainHtml.includes(text.slice(0, 20))) textboxLost.push(text);
  }
  if (textboxLost.length > 0) {
    losses.push({
      kind: "lost",
      label: "文本框文字",
      count: textboxLost.length,
      detail: `共 ${textboxTotal} 个文本框，其中 ${textboxLost.length} 个的文字没进正文。例如：${preview(
        textboxLost[0],
      )}`,
    });
  }

  if (headerFooterParts > 0) {
    losses.push({
      kind: "lost",
      label: "页眉/页脚",
      count: headerFooterParts,
      detail: `${headerFooterParts} 个页眉页脚部件不导入，含 ${headerFooterTexts} 段文字、${headerFooterImages} 张图片。`,
    });
  }

  const fieldNames = new Map<string, number>();
  for (const match of body.matchAll(/<w:instrText(?:\s[^>]*)?>([\s\S]*?)<\/w:instrText>/g)) {
    const words = decodeXml(match[1]).toUpperCase().match(/[A-Z][A-Z0-9]*/g);
    const name = words?.[0];
    if (!name || name === "HYPERLINK") continue;
    fieldNames.set(name, (fieldNames.get(name) || 0) + 1);
  }
  if (fieldNames.size > 0) {
    const summary = Array.from(fieldNames)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 4)
      .map(([name, total]) => `${name}×${total}`)
      .join("、");
    losses.push({
      kind: "downgraded",
      label: "域代码结果",
      count: count(body, /<w:instrText\b/g),
      detail: `页码、目录、交叉引用等域的计算结果不完整（${summary}）。`,
    });
  }

  if (ctx.internalLinks > 0) {
    losses.push({
      kind: "lost",
      label: "文档内跳转",
      count: ctx.internalLinks,
      detail: "目录条目、书签和交叉引用点不动了，文字已保留。",
    });
  }

  const relTargets = new Map<string, string>();
  const externalRids = new Set<string>();
  const relsFile = zip.file("word/_rels/document.xml.rels");
  if (relsFile) {
    const rels = await relsFile.async("string");
    for (const match of rels.matchAll(/<Relationship\b[^>]*>/g)) {
      const id = /\bId="([^"]*)"/.exec(match[0])?.[1];
      const target = /\bTarget="([^"]*)"/.exec(match[0])?.[1];
      if (id && target) relTargets.set(id, target);
      if (id && /\bTargetMode="External"/.test(match[0])) externalRids.add(id);
    }
  }

  const blipIds = Array.from(
    documentXml.matchAll(/<a:blip\b[^>]*\br:(?:embed|link)="([^"]+)"/g),
  ).map((match) => decodeXml(match[1]));
  const externalImages = blipIds.filter((id) => externalRids.has(id)).length;
  if (externalImages > 0) {
    losses.push({
      kind: "lost",
      label: "链接到文件的图片",
      count: externalImages,
      detail:
        "Word 里以「链接到文件」方式插入的图，图片本体不在文档内，无法导入。",
    });
  }

  const embeds = blipIds.map((id) => relTargets.get(id) || id);
  const duplicates = embeds.length - new Set(embeds).size;
  if (duplicates > 0) {
    losses.push({
      kind: "info",
      label: "被多次引用的同一图片",
      count: duplicates,
      detail: `${embeds.length} 处图片引用里有 ${duplicates} 处是重复图片，附件会各存一份。`,
    });
  }

  const embedded =
    count(body, /<w:object\b/g) +
    count(body, /\buri="[^"]*\/chart"/g) +
    count(body, /\buri="[^"]*diagram(?:Layout|Data|Colors|Text|QuickStyle)?"/g) +
    count(body, /<w:altChunk\b/g);
  if (embedded > 0) {
    losses.push({
      kind: "lost",
      label: "图表与嵌入对象",
      count: embedded,
      detail: "图表、SmartArt、Excel 等嵌入对象不导入（Word 里看到的是它们的预览图）。",
    });
  }

  return losses;
}

function count(xml: string, pattern: RegExp): number {
  let total = 0;
  for (const _ of xml.matchAll(pattern)) total++;
  return total;
}

/**
 * 兜底自检：逐段核对原文有没有进笔记，并比对出现次数。
 * 只管"文字"这一维，抓的是我们没预料到的丢失；已知类别由上面的清单负责。
 * 文本框和嵌入对象各有专门条目，这里跳过以免同一件事报两遍。
 */
const FIDELITY_SKIP = ["w:object", "w:txbxContent"];
const MIN_SEGMENT = 4;

export function checkTextFidelity(
  documentXml: string,
  markdown: string,
): LossItem | null {
  const haystack = plainText(markdown);
  if (!haystack) return null;

  let region = documentXml;
  for (const tag of FIDELITY_SKIP) {
    region = region.replace(new RegExp(`<${tag}\\b[\\s\\S]*?</${tag}>`, "g"), "");
  }

  const segments = new Map<string, { expected: number; sample: string }>();
  for (const match of region.matchAll(TEXT_RUN)) {
    const raw = decodeXml(match[1]);
    const key = plainText(raw);
    if (key.length < MIN_SEGMENT) continue;
    const known = segments.get(key);
    if (known) known.expected++;
    else segments.set(key, { expected: 1, sample: raw.trim() });
  }
  if (segments.size === 0) return null;

  const missing: string[] = [];
  let missingCount = 0;
  let shrunkCount = 0;
  for (const [key, entry] of segments) {
    const found = occurrences(haystack, key, entry.expected);
    if (found === 0) {
      missingCount++;
      if (missing.length < 3) missing.push(entry.sample.slice(0, 30));
    } else if (found < entry.expected) {
      shrunkCount++;
    }
  }

  if (missingCount === 0 && shrunkCount === 0) return null;
  const parts: string[] = [];
  if (missingCount > 0) parts.push(`${missingCount} 段文字在笔记里找不到`);
  if (shrunkCount > 0) parts.push(`${shrunkCount} 段文字出现次数比原文档少`);
  return {
    kind: "lost",
    label: "疑似未识别的内容丢失",
    count: missingCount + shrunkCount,
    detail: `${parts.join("，")}。${missing.length ? `例如：${missing.map((text) => `「${text}」`).join(" ")}` : ""}这类差异不在已知清单里，请把这份文件发给开发者排查。`,
  };
}

function occurrences(haystack: string, needle: string, atLeast: number): number {
  let count = 0;
  let from = 0;
  while (count < atLeast) {
    const index = haystack.indexOf(needle, from);
    if (index === -1) break;
    count++;
    from = index + needle.length;
  }
  return count;
}

/** 抹掉 Markdown 语法痕迹，只留可对比的正文字符 */
function plainText(value: string): string {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, digits: string) => String.fromCodePoint(Number(digits)))
    .replace(/&#x([0-9a-f]+);/gi, (_, digits: string) =>
      String.fromCodePoint(parseInt(digits, 16)),
    )
    .replace(/&amp;/g, "&")
    .replace(/\\(.)/g, "$1")
    .replace(/<\/?[a-z][^>]*>/gi, "")
    .replace(/[![\]()<>#*_|~`-]/g, "")
    .replace(/[\s　]+/g, "");
}

function sample(xml: string, pattern: RegExp, maxLength = 60): string {
  const parts: string[] = [];
  let length = 0;
  for (const match of xml.matchAll(pattern)) {
    const text = decodeXml(match[1]).replace(/\s+/g, " ").trim();
    if (!text) continue;
    parts.push(text);
    length += text.length;
    if (length >= maxLength) break;
  }
  return parts.join("");
}

function preview(text: string): string {
  const characters = Array.from(text);
  return characters.length > 40 ? `${characters.slice(0, 40).join("")}…` : text;
}

function stripTags(html: string): string {
  return html
    .replace(/<[^>]*>/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function decodeXml(value: string): string {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, digits: string) => String.fromCodePoint(Number(digits)))
    .replace(/&#x([0-9a-f]+);/gi, (_, digits: string) =>
      String.fromCodePoint(parseInt(digits, 16)),
    )
    .replace(/&amp;/g, "&");
}
