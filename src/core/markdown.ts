import TurndownService from "turndown";
import { strikethrough } from "turndown-plugin-gfm";
import { IMAGE_TOKEN_PREFIX, PendingImage } from "./docx";

export interface MarkdownConversion {
  markdown: string;
  simpleTables: number;
  complexTables: number;
  warnings: string[];
}

export function htmlToMarkdown(
  html: string,
  images: PendingImage[],
): MarkdownConversion {
  const warnings: string[] = [];
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

  service.addRule("sunDaoruUnsafeLink", {
    filter: (node) =>
      node.nodeName === "A" && isUnsafeHref(node.getAttribute("href")),
    replacement: (content) => content,
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
      if (image.previewable) {
        const width = image.widthPx ? `|${image.widthPx}` : "";
        return `![[${image.link}${width}]]`;
      }
      const alt = cleanAlt(node.getAttribute("alt") || "原图");
      return `[[${image.link}|${alt}]]`;
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

  service.addRule("sunDaoruHighlight", {
    filter: "mark",
    replacement: (content) => `==${content}==`,
  });

  let markdown = service.turndown(html);
  markdown = cleanMarkdown(markdown);

  if (tableImages > 0) {
    warnings.push(
      `复杂表格内有 ${tableImages} 处图片，已保留为嵌入语法文本；Obsidian 不一定会直接在 HTML 表格中渲染。`,
    );
  }
  if (unresolvedImages > 0) {
    warnings.push(`有 ${unresolvedImages} 张图片未能匹配到附件，已从正文中移除。`);
  }
  if (markdown.includes(IMAGE_TOKEN_PREFIX)) {
    warnings.push("正文中残留了图片占位符，请把文档发给开发者排查。");
  }

  return { markdown, simpleTables, complexTables, warnings };
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
    const blank = doc.createTextNode(" ");
    const breakBefore = doc.createElement("br");
    const breakAfter = doc.createElement("br");
    let text: Text;

    if (pending && pending.link) {
      const alt = cleanAlt(image.getAttribute("alt") || "原图");
      text = doc.createTextNode(
        pending.previewable ? `![[${pending.link}]]` : `[[${pending.link}|${alt}]]`,
      );
      count++;
    } else {
      text = doc.createTextNode("（图片缺失）");
    }

    parent.insertBefore(blank, image);
    parent.insertBefore(breakBefore, image);
    parent.insertBefore(text, image);
    parent.insertBefore(breakAfter, image);
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
    const colspan = element.getAttribute("colspan");
    const rowspan = element.getAttribute("rowspan");
    for (const attr of Array.from(element.attributes)) {
      element.removeAttribute(attr.name);
    }

    if (tag === "a" && href && !isUnsafeHref(href)) {
      element.setAttribute("href", href.trim());
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
  return /^(?:javascript|vbscript|data|file):/.test(value) || value.startsWith("#");
}

function cleanAlt(value: string): string {
  const clean = value
    .replace(/[|[\]]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return clean || "原图";
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
