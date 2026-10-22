import { decodeXml, xmlValue } from "./xml";

interface StyleDef {
  id: string;
  name: string;
  outline?: number;
  basedOn?: string;
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

export function buildStyleMap(
  stylesXml: string,
  documentXml: string,
  normalizeHeadingLevels: boolean,
): { entries: string[]; shifted: boolean } {
  const styles = parseStyles(stylesXml);
  const usedIds = new Set<string>();
  for (const match of documentXml.matchAll(/<w:pStyle\b[^>]*\bw:val="([^"]*)"/g)) {
    usedIds.add(decodeXml(match[1]));
  }

  const usedStyles: StyleDef[] = [];
  for (const id of usedIds) {
    const style = styles.get(id);
    if (style) usedStyles.push(style);
  }
  const hasFileTitle = usedStyles.some((style) => isFileTitleStyle(style.name));
  const shift = normalizeHeadingLevels && hasFileTitle;

  const entries: string[] = [];
  const seenNames = new Set<string>();
  const addEntry = (name: string, level: number) => {
    const cleanName = name.trim();
    if (!cleanName || cleanName.includes("'") || cleanName.includes("\\")) return;
    const key = cleanName.toLowerCase();
    if (seenNames.has(key)) return;
    seenNames.add(key);
    const effectiveLevel = isFileTitleStyle(cleanName)
      ? 1
      : shift
        ? Math.min(level + 1, 6)
        : level;
    entries.push(`p[style-name='${cleanName}'] => h${effectiveLevel}:fresh`);
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

  return { entries, shifted: shift };
}

/** 段落样式 ID → 标题层级，供书签定位使用 */
export function headingLevelsByStyleId(stylesXml: string): Map<string, number> {
  const styles = parseStyles(stylesXml);
  const levels = new Map<string, number>();
  for (const style of styles.values()) {
    const level = resolveLevel(style, styles, new Set());
    if (level) levels.set(style.id, level);
  }
  return levels;
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

  if (isFileTitleStyle(trimmed)) return 1;

  const chinese = /^(一级|二级|三级|四级|五级|六级)标题$/.exec(trimmed);
  if (chinese) {
    return { 一级: 1, 二级: 2, 三级: 3, 四级: 4, 五级: 5, 六级: 6 }[chinese[1]];
  }

  return undefined;
}

export function isFileTitleStyle(name: string): boolean {
  return /^(?:title|文件标题)(?=$|[_\-\s.])/i.test(name.trim());
}
