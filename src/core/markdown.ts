import TurndownService from "turndown";
import { strikethrough } from "turndown-plugin-gfm";
import { IMAGE_TOKEN_PREFIX, PendingImage } from "./docx";
import { LossItem } from "./losses";

export interface MarkdownConversion {
  markdown: string;
  simpleTables: number;
  complexTables: number;
  warnings: string[];
  losses: LossItem[];
}

export function htmlToMarkdown(
  html: string,
  images: PendingImage[],
): MarkdownConversion {
  const warnings: string[] = [];
  const losses: LossItem[] = [];
  const imageByToken = new Map(images.map((image) => [image.token, image]));
  let simpleTables = 0;
  let complexTables = 0;
  let tableImages = 0;
  let unresolvedImages = 0;

  const service = new TurndownService({
    headingStyle: "atx",
    hr: "---",
    bulletListMarker: "-",
    codeBlockStyle: "fenced",
    emDelimiter: "*",
    strongDelimiter: "**",
    linkStyle: "inlined",
  });

  service.use(strikethrough);
  // Word 正文里字面写出的 <b>文字</b> 之类会被 Markdown 渲染器当 HTML 吞掉，
  // 文本节点里的 < 一律转义（代码内容 turndown 不走 escape，不受影响）
  const defaultEscape = service.escape.bind(service);
  service.escape = (text: string) => defaultEscape(text).replace(/</g, "&lt;");

  service.remove([
    "script",
    "style",
    "noscript",
    "iframe",
    "object",
    "embed",
    "form",
    "input",
    "button",
    "textarea",
    "select",
    "option",
  ]);

  service.addRule("sunDaoruLink", {
    filter: "a",
    replacement: (content, node) => {
      const href = (node.getAttribute("href") || "").trim();
      const label = content.replace(/\s+/g, " ").trim();
      // 文档内跳转（#锚点）不做转换：标准 Markdown 的锚点规则和 Word 对不上，只保留文字
      if (!href || href.startsWith("#") || isUnsafeHref(href)) return content;
      if (!label) return content;
      if (label.startsWith("[") && label.endsWith(")")) return content;
      return `[${label}](${escapeHref(href)})`;
    },
  });

  service.addRule("sunDaoruImage", {
    filter: (node) =>
      node.nodeName === "IMG" &&
      (node.getAttribute("src") || "").startsWith(IMAGE_TOKEN_PREFIX),
    replacement: (_content, node) => {
      const token = node.getAttribute("src") || "";
      const image = imageByToken.get(token);
      if (!image || !image.link) {
        unresolvedImages++;
        return "";
      }
      const alt = cleanAlt(node.getAttribute("alt") || image.altText || "");
      if (!image.previewable) return `[${alt || "原图"}](${destination(image.link)})`;
      return imageEmbed(image.link, alt);
    },
  });

  service.addRule("sunDaoruTable", {
    filter: "table",
    replacement: (_content, node) => {
      const table = node as HTMLTableElement;
      if (isSimpleTable(table)) {
        const markdown = simpleTableToMarkdown(table, service);
        if (markdown) {
          simpleTables++;
          return `\n\n${markdown}\n\n`;
        }
      }
      complexTables++;
      const result = complexTableToHtml(table, imageByToken);
      tableImages += result.images;
      return `\n\n${result.html}\n\n`;
    },
  });

  service.addRule("sunDaoruLineBreak", {
    filter: "br",
    replacement: () => "<br>\n",
  });

  service.addRule("sunDaoruCodeBlock", {
    filter: "pre",
    replacement: (_content, node) => {
      const text = (node.textContent || "").replace(/\s+$/, "");
      if (!text.trim()) return "";
      const fence = text.includes("```") ? "````" : "```";
      return `\n\n${fence}\n${text}\n${fence}\n\n`;
    },
  });

  service.addRule("sunDaoruUnderline", {
    filter: "u",
    replacement: (content) => `<u>${content}</u>`,
  });

  service.addRule("sunDaoruSuperscript", {
    filter: "sup",
    replacement: (content) => `<sup>${content}</sup>`,
  });

  service.addRule("sunDaoruSubscript", {
    filter: "sub",
    replacement: (content) => `<sub>${content}</sub>`,
  });

  let markdown = service.turndown(html);
  markdown = cleanMarkdown(markdown);

  if (tableImages > 0) {
    losses.push({
      kind: "info",
      label: "HTML 表格里的图片",
      count: tableImages,
      detail: "表格是 HTML 块，标准图片语法进不去，用 <img> 标签保留。",
    });
  }
  if (unresolvedImages > 0) {
    losses.push({
      kind: "lost",
      label: "找不到附件的图片引用",
      count: unresolvedImages,
      detail: "这些图片没有对应的图片数据，已从正文移除。",
    });
  }
  if (markdown.includes(IMAGE_TOKEN_PREFIX)) {
    warnings.push("正文中残留了图片占位符，请把文档发给开发者排查。");
  }

  return { markdown, simpleTables, complexTables, warnings, losses };
}

function imageEmbed(link: string, alt: string): string {
  // Obsidian 会把纯数字的 alt 当成显示尺寸，只能退回 <img> 才保得住这行文字
  if (/^\d+(?:x\d+)?$/.test(alt)) return htmlImage(link, alt);
  return `![${alt}](${destination(link)})`;
}

/** 标准 Markdown 的链接目标：空格和括号要用尖括号包住，尖括号本身转义 */
function destination(path: string): string {
  const clean = path.replace(/</g, "%3C").replace(/>/g, "%3E");
  return /[\s()]/.test(clean) ? `<${clean}>` : clean;
}

function htmlImage(link: string, alt: string): string {
  return `<img src="${escapeAttribute(link)}" alt="${escapeAttribute(alt)}">`;
}

function isSimpleTable(table: HTMLTableElement): boolean {
  if (table.querySelector("table")) return false;
  if (table.querySelector("caption")) return false;
  if (table.querySelector("[colspan],[rowspan]")) return false;
  if (table.querySelector("ol,ul,li,pre,blockquote,figure,div,section,article,dl")) {
    return false;
  }

  const rows = Array.from(table.querySelectorAll("tr"));
  if (rows.length === 0) return false;

  const columnCount = directCells(rows[0]).length;
  if (columnCount === 0) return false;
  if (!rows.every((row) => directCells(row).length === columnCount)) return false;

  const header = findHeaderRow(table, rows);
  if (!header) return false;

  const ths = Array.from(table.querySelectorAll("th"));
  if (ths.some((th) => th.parentElement !== header)) return false;

  return true;
}

function findHeaderRow(
  table: HTMLTableElement,
  rows: Element[],
): HTMLTableRowElement | null {
  const thead = table.querySelector("thead");
  if (thead) {
    const headRows = Array.from(thead.querySelectorAll("tr"));
    if (headRows.length !== 1) return null;
    return headRows[0] as HTMLTableRowElement;
  }
  return rows[0] as HTMLTableRowElement;
}

function simpleTableToMarkdown(
  table: HTMLTableElement,
  service: TurndownService,
): string {
  const rows = Array.from(table.querySelectorAll("tr"));
  const header = findHeaderRow(table, rows);
  if (!header) return "";

  const headerCells = directCells(header);
  const bodyRows = rows.filter((row) => row !== header);
  const lines: string[] = [];

  lines.push(
    `| ${headerCells.map((cell) => cellMarkdown(cell, service) || " ").join(" | ")} |`,
  );
  lines.push(`| ${headerCells.map(() => "---").join(" | ")} |`);
  for (const row of bodyRows) {
    const cells = directCells(row);
    if (cells.length !== headerCells.length) return "";
    lines.push(
      `| ${cells.map((cell) => cellMarkdown(cell, service) || " ").join(" | ")} |`,
    );
  }

  return lines.join("\n");
}

function cellMarkdown(cell: Element, service: TurndownService): string {
  let markdown = service.turndown(cell.innerHTML);
  markdown = markdown
    .replace(/\r/g, "")
    .replace(/<br\s*\/?>/gi, "<br>")
    .replace(/\n+/g, "<br>")
    .replace(/(<br>)+/g, "<br>")
    .replace(/\|/g, "\\|")
    .trim();
  return markdown;
}

function complexTableToHtml(
  table: HTMLTableElement,
  imageByToken: Map<string, PendingImage>,
): { html: string; images: number } {
  const clone = table.cloneNode(true) as HTMLTableElement;
  const images = replaceTableImages(clone, imageByToken);
  sanitizeTable(clone);
  return { html: clone.outerHTML, images };
}

function replaceTableImages(
  root: HTMLTableElement,
  imageByToken: Map<string, PendingImage>,
): number {
  let count = 0;

  for (const image of Array.from(root.querySelectorAll("img"))) {
    const token = image.getAttribute("src") || "";
    const pending = imageByToken.get(token);
    const parent = image.parentNode;
    if (!parent) continue;

    const doc = image.ownerDocument;
    const alt = cleanAlt(image.getAttribute("alt") || pending?.altText || "");
    let text: Node;

    if (pending?.link && pending.previewable) {
      const element = doc.createElement("img");
      element.setAttribute("src", pending.link);
      element.setAttribute("alt", alt);
      text = element;
      count++;
    } else if (pending?.link) {
      text = doc.createTextNode(
        `${alt || "图片"}（${pending.link}，这种格式无法预览，已保存原文件）`,
      );
      count++;
    } else {
      text = doc.createTextNode("（图片缺失）");
    }

    parent.insertBefore(doc.createTextNode(" "), image);
    parent.insertBefore(doc.createElement("br"), image);
    parent.insertBefore(text, image);
    parent.insertBefore(doc.createElement("br"), image);
    parent.removeChild(image);
  }
  return count;
}

const ALLOWED_TABLE_TAGS = new Set([
  "table",
  "thead",
  "tbody",
  "tfoot",
  "tr",
  "th",
  "td",
  "caption",
  "p",
  "br",
  "strong",
  "b",
  "em",
  "i",
  "u",
  "s",
  "strike",
  "del",
  "sup",
  "sub",
  "a",
  "img",
  "ol",
  "ul",
  "li",
  "mark",
  "code",
  "pre",
  "blockquote",
]);

const DANGEROUS_TABLE_TAGS = new Set([
  "script",
  "style",
  "iframe",
  "object",
  "embed",
  "form",
  "input",
  "button",
  "textarea",
  "select",
  "option",
  "link",
  "meta",
]);

function sanitizeTable(root: HTMLTableElement): void {
  for (const attr of Array.from(root.attributes)) {
    root.removeAttribute(attr.name);
  }
  for (const element of Array.from(root.querySelectorAll("*"))) {
    const tag = element.tagName.toLowerCase();
    if (DANGEROUS_TABLE_TAGS.has(tag)) {
      element.remove();
      continue;
    }
    if (!ALLOWED_TABLE_TAGS.has(tag)) {
      unwrap(element);
      continue;
    }

    const href = tag === "a" ? element.getAttribute("href") : null;
    const src = tag === "img" ? element.getAttribute("src") : null;
    const alt = tag === "img" ? element.getAttribute("alt") : null;
    const colspan = element.getAttribute("colspan");
    const rowspan = element.getAttribute("rowspan");
    for (const attr of Array.from(element.attributes)) {
      element.removeAttribute(attr.name);
    }

    if (tag === "a" && href && !isUnsafeHref(href)) {
      element.setAttribute("href", href.trim());
    }
    if (tag === "img") {
      if (src) element.setAttribute("src", src.trim());
      if (alt) element.setAttribute("alt", alt);
    }
    if (tag === "td" || tag === "th") {
      if (colspan && /^\d+$/.test(colspan) && Number(colspan) > 1) {
        element.setAttribute("colspan", colspan);
      }
      if (rowspan && /^\d+$/.test(rowspan) && Number(rowspan) > 1) {
        element.setAttribute("rowspan", rowspan);
      }
    }
  }
}

function unwrap(element: Element): void {
  const parent = element.parentNode;
  if (!parent) return;
  while (element.firstChild) parent.insertBefore(element.firstChild, element);
  parent.removeChild(element);
}

function directCells(row: Element): HTMLTableCellElement[] {
  return Array.from(row.children).filter(
    (child) => child.tagName === "TD" || child.tagName === "TH",
  ) as HTMLTableCellElement[];
}

function isUnsafeHref(href: string | null): boolean {
  if (!href) return false;
  const value = href.trim().toLowerCase().replace(/[\u0000-\u001f]/g, "");
  if (!value) return false;
  return /^(?:javascript|vbscript|data|file):/.test(value);
}

function cleanAlt(value: string): string {
  return value
    .replace(/[|[\]]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function escapeHref(href: string): string {
  return href.replace(
    /[ ()<>"`]/g,
    (character) => `%${character.charCodeAt(0).toString(16).toUpperCase().padStart(2, "0")}`,
  );
}

function escapeAttribute(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function cleanMarkdown(markdown: string): string {
  return (
    markdown
      .replace(/\u00a0/g, " ")
      .replace(/[\u200b-\u200d\u2060\ufeff]/g, "")
      .replace(/[ \t]+\n/g, "\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim() + "\n"
  );
}
