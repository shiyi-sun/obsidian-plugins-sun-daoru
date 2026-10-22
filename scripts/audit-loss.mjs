/*
静默丢失审计：不测"导入对不对"，专测"丢了有没有说"。

做法：把 docx 里所有内容性元素列成一张清单（正文文本、图片、链接、脚注、
批注、公式、页眉页脚、修订、图表/嵌入对象、域代码、文本框……），
每一项按三种结局归类：
  已导入  —— 内容出现在生成的 Markdown 里
  已报告  —— 没导入，但丢失清单里有对应条目
  静默丢失 —— 既没导入也没报告  ← 只有这一格是问题

用法：
  node scripts/audit-loss.mjs                     # 跑仓库根目录的样本 + DAORU_EXTRA_DOCX
  DAORU_EXTRA_DOCX=/path/a.docx,b.docx node scripts/audit-loss.mjs
*/
import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import esbuild from "esbuild";
import JSZip from "jszip";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const bundlePath = path.join(root, ".dist-test", "core.cjs");

await esbuild.build({
  stdin: {
    contents: [
      'export { convertDocx } from "./src/core/docx";',
      'export { htmlToMarkdown } from "./src/core/markdown";',
    ].join("\n"),
    resolveDir: root,
    sourcefile: "audit-entry.ts",
    loader: "ts",
  },
  bundle: true,
  platform: "node",
  format: "cjs",
  target: "node18",
  outfile: bundlePath,
  logLevel: "silent",
});

const require = createRequire(import.meta.url);
const { convertDocx, htmlToMarkdown } = require(bundlePath);

const toArrayBuffer = (buffer) =>
  buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);

/** 抹掉 Markdown 语法痕迹，只留可比对的正文字符 */
function normalize(value) {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, digits) => String.fromCodePoint(Number(digits)))
    .replace(/&#x([0-9a-f]+);/gi, (_, digits) => String.fromCodePoint(parseInt(digits, 16)))
    .replace(/\\(.)/g, "$1")
    .replace(/<\/?(?:br|sup|sub|u|b|i|em|strong|s|del|strike|mark|code|pre|kbd|table|thead|tbody|tfoot|tr|td|th|caption|p|ol|ul|li|div|span|blockquote|h[1-6]|a|img)\b[^>]*>/gi, "")
    .replace(/[![\]()<>#*_|~-]/g, "")
    .replace(/[\s\u00a0\u200b-\u200d\ufeff]+/g, "");
}

const textRuns = (xml) =>
  Array.from(xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g))
    .map((match) => decode(match[1]))
    .filter((text) => text.trim().length > 0);

function decode(value) {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

const count = (xml, pattern) => Array.from(xml.matchAll(pattern)).length;
const blank = (xml, tag) =>
  xml.replace(new RegExp(`<${tag}\\b[\\s\\S]*?</${tag}>`, "g"), "<!--blanked-->");

async function audit(file) {
  const buffer = fs.readFileSync(file);
  const zip = await JSZip.loadAsync(buffer);
  const read = async (name) => (zip.file(name) ? await zip.file(name).async("string") : "");

  const documentXml = await read("word/document.xml");
  const footnotesXml = await read("word/footnotes.xml");
  const endnotesXml = await read("word/endnotes.xml");
  const commentsXml = await read("word/comments.xml");
  const relsXml = await read("word/_rels/document.xml.rels");
  const partNames = Object.keys(zip.files);
  const headerFooterNames = partNames.filter((name) =>
    /^word\/(?:header|footer)\d*\.xml$/.test(name),
  );
  let headerFooterXml = "";
  for (const name of headerFooterNames) headerFooterXml += await read(name);

  /* 生成 Markdown（附件路径用假数据，只关心文本层面） */
  const conversion = await convertDocx(toArrayBuffer(buffer));
  conversion.images.forEach((image, index) => {
    image.link = `附件/审计-${String(index + 1).padStart(3, "0")}.${image.extension}`;
  });
  const rendered = htmlToMarkdown(conversion.html, conversion.images);
  const markdown = normalize(rendered.markdown);
  const lossLabels = [...conversion.losses, ...rendered.losses].map((loss) => loss.label);
  const reported = (label) => lossLabels.includes(label);
  const reportedCount = (label) =>
    [...conversion.losses, ...rendered.losses].find((loss) => loss.label === label)?.count ?? 0;

  const rows = [];
  const add = (类别, 文档里有, 结局, 说明) =>
    rows.push({ 类别, 文档里有, 结局, 说明 });

  /* —— 正文文本：逐 run 比对 —— */
  const textboxXml = Array.from(
    documentXml.matchAll(/<w:txbxContent\b[\s\S]*?<\/w:txbxContent>/g),
  ).map((match) => match[0]).join("");
  const objectXml = Array.from(
    documentXml.matchAll(/<w:object\b[\s\S]*?<\/w:object>/g),
  ).map((match) => match[0]).join("");
  const bodyXml = blank(blank(documentXml, "w:txbxContent"), "w:object");

  const bodyRuns = textRuns(bodyXml).filter((text) => normalize(text).length >= 4);
  const missingRuns = bodyRuns.filter((text) => !markdown.includes(normalize(text)));
  add(
    "正文文字",
    `${bodyRuns.length} 段（≥4 字）`,
    missingRuns.length === 0 ? "已导入" : "静默丢失",
    missingRuns.length === 0
      ? "全部逐段出现在 Markdown 里"
      : `未出现：${missingRuns.slice(0, 3).map((t) => `「${t.slice(0, 30)}」`).join(" ")}`,
  );

  /* —— 脚注 / 尾注：只算正文真正引用到的定义（孤儿定义 Word 也不显示） —— */
  const referenced = (xml, tag) =>
    new Set(
      Array.from(xml.matchAll(new RegExp(`<w:${tag}Reference\\b[^>]*\\bw:id="([^"]+)"`, "g")))
        .map((match) => match[1]),
    );
  const definitions = (xml, tag, ids) =>
    Array.from(xml.matchAll(new RegExp(`<w:${tag}\\b([^>]*)>([\\s\\S]*?)</w:${tag}>`, "g")))
      .filter((match) => {
        const attrs = match[1];
        if (/w:type="(separator|continuationSeparator)"/.test(attrs)) return false;
        const id = /\bw:id="([^"]*)"/.exec(attrs)?.[1];
        return id === undefined || ids.has(id);
      })
      .map((match) => match[2])
      .join("");

  const footnoteRuns = textRuns(
    definitions(footnotesXml, "footnote", referenced(documentXml, "footnote")),
  ).filter((t) => normalize(t).length >= 4);
  const footnoteMissing = footnoteRuns.filter((t) => !markdown.includes(normalize(t)));
  add("脚注文字", `${footnoteRuns.length} 段`, footnoteMissing.length ? "静默丢失" : footnoteRuns.length ? "已导入" : "无", footnoteMissing.map((t) => t.slice(0, 24)).join(" "));
  const endnoteRuns = textRuns(
    definitions(endnotesXml, "endnote", referenced(documentXml, "endnote")),
  ).filter((t) => normalize(t).length >= 4);
  const endnoteMissing = endnoteRuns.filter((t) => !markdown.includes(normalize(t)));
  add("尾注文字", `${endnoteRuns.length} 段`, endnoteMissing.length ? "静默丢失" : endnoteRuns.length ? "已导入" : "无", endnoteMissing.map((t) => t.slice(0, 24)).join(" "));

  /* —— 批注 —— */
  const commentRuns = textRuns(commentsXml).filter((t) => normalize(t).length >= 2);
  const commentImported = commentRuns.filter((t) => markdown.includes(normalize(t))).length;
  add(
    "批注文字",
    `${commentRuns.length} 段`,
    commentRuns.length === 0
      ? "无"
      : reported("批注内容")
        ? "已报告"
        : commentImported === commentRuns.length
          ? "已导入"
          : "静默丢失",
    `导入 ${commentImported}，报告 ${reportedCount("批注内容")} 条`,
  );

  /* —— 页眉页脚 —— */
  const hfRuns = textRuns(headerFooterXml).filter((t) => normalize(t).length >= 2);
  const hfImported = hfRuns.filter((t) => markdown.includes(normalize(t))).length;
  add(
    "页眉页脚文字",
    `${hfRuns.length} 段`,
    hfRuns.length === 0 ? "无" : reported("页眉/页脚") ? "已报告" : hfImported === hfRuns.length ? "已导入" : "静默丢失",
    `导入 ${hfImported}，报告 ${reportedCount("页眉/页脚")} 个部件`,
  );

  /* —— 修订里被删掉的文字 —— */
  const deletedRuns = Array.from(
    [documentXml, footnotesXml, endnotesXml, commentsXml].join("\n").matchAll(
      /<w:delText(?:\s[^>]*)?>([\s\S]*?)<\/w:delText>/g,
    ),
  ).map((match) => decode(match[1])).filter((t) => normalize(t).length >= 2);
  add(
    "修订被删文字",
    `${deletedRuns.length} 段`,
    deletedRuns.length === 0 ? "无" : reported("修订里被删掉的文字") ? "已报告" : "静默丢失",
    `报告 ${reportedCount("修订里被删掉的文字")} 处`,
  );

  /* —— 公式 —— */
  const mathCount = count(documentXml, /<m:oMath\b/g);
  const mathTexts = Array.from(documentXml.matchAll(/<m:t(?:\s[^>]*)?>([\s\S]*?)<\/m:t>/g)).map((m) => decode(m[1]));
  const mathImported = mathTexts.filter((t) => normalize(t) && markdown.includes(normalize(t))).length;
  add(
    "公式",
    `${mathCount} 个（含 ${mathTexts.length} 段文字）`,
    mathCount === 0 ? "无" : reported("公式") ? "已报告" : mathImported === mathTexts.length ? "已导入" : "静默丢失",
    `文字导入 ${mathImported}/${mathTexts.length}，报告 ${reportedCount("公式")} 个`,
  );

  /* —— 文本框 —— */
  const textboxRuns = textRuns(textboxXml).filter((t) => normalize(t).length >= 4);
  const textboxImported = textboxRuns.filter((t) => markdown.includes(normalize(t))).length;
  add(
    "文本框文字",
    `${textboxRuns.length} 段`,
    textboxRuns.length === 0
      ? "无"
      : textboxImported === textboxRuns.length
        ? "已导入"
        : reported("文本框文字")
          ? "已报告"
          : "静默丢失",
    `导入 ${textboxImported}/${textboxRuns.length}`,
  );

  /* —— 图表 / OLE / SmartArt —— */
  const embedded =
    count(documentXml, /<w:object\b/g) +
    count(documentXml, /\buri="[^"]*\/chart"/g) +
    count(documentXml, /\buri="[^"]*diagram(?:Layout|Data|Colors|Text|QuickStyle)?"/g) +
    count(documentXml, /<w:altChunk\b/g);
  add(
    "图表/嵌入对象",
    `${embedded} 个`,
    embedded === 0 ? "无" : reported("图表与嵌入对象") ? "已报告" : "静默丢失",
    `报告 ${reportedCount("图表与嵌入对象")} 个`,
  );

  /* —— 图片 —— */
  const rels = new Map(
    Array.from(relsXml.matchAll(/<Relationship\b[^>]*>/g)).map((match) => {
      const id = /\bId="([^"]*)"/.exec(match[0])?.[1];
      const external = /\bTargetMode="External"/.test(match[0]);
      return [id, { target: decode(/\bTarget="([^"]*)"/.exec(match[0])?.[1] || ""), external }];
    }),
  );
  const blips = Array.from(
    documentXml.matchAll(/<a:blip\b[^>]*\br:(embed|link)="([^"]+)"/g),
  ).map((match) => ({ kind: match[1], id: match[2], external: rels.get(match[2])?.external ?? false }));
  const internalBlips = blips.filter((blip) => !blip.external && blip.kind === "embed");
  const externalBlips = blips.filter((blip) => blip.external || blip.kind === "link");
  const vmlImages = count(documentXml, /<v:imagedata\b/g);
  const extracted = conversion.images.length;
  add(
    "内嵌图片",
    `${internalBlips.length} 处引用`,
    internalBlips.length === 0 ? "无" : extracted >= internalBlips.length ? "已导入" : "静默丢失",
    `提取到 ${extracted} 张${vmlImages ? ` · 另有 VML 图片引用 ${vmlImages} 处（多在 w:object 内）` : ""}`,
  );
  add(
    "链接到文件的图片",
    `${externalBlips.length} 处`,
    externalBlips.length === 0 ? "无" : reported("链接到文件的图片") ? "已报告" : "静默丢失",
    `报告 ${reportedCount("链接到文件的图片")} 条`,
  );

  /* —— 链接 —— */
  const hyperlinkIds = Array.from(
    documentXml.matchAll(/<w:hyperlink\b[^>]*\br:id="([^"]+)"/g),
  ).map((match) => match[1]);
  const httpTargets = hyperlinkIds.map((id) => rels.get(id)?.target || "").filter((t) => /^https?:|^mailto:/i.test(t));
  const httpMissing = httpTargets.filter((target) => !rendered.markdown.includes(target) && !normalize(rendered.markdown).includes(normalize(target)));
  add(
    "外部链接",
    `${httpTargets.length} 条`,
    httpTargets.length === 0 ? "无" : httpMissing.length === 0 ? "已导入" : "静默丢失",
    httpMissing.length ? `缺失：${httpMissing.slice(0, 2).join(" ")}` : "",
  );
  const anchorLinks = count(documentXml, /<w:hyperlink\b[^>]*\bw:anchor="/g);
  add(
    "文档内跳转",
    `${anchorLinks} 条`,
    anchorLinks === 0 ? "无" : reported("文档内跳转") ? "已报告" : "静默丢失",
    `报告 ${reportedCount("文档内跳转")} 条`,
  );

  /* —— 域代码 —— */
  const fieldNames = new Map();
  for (const match of documentXml.matchAll(/<w:instrText(?:\s[^>]*)?>([\s\S]*?)<\/w:instrText>/g)) {
    const name = decode(match[1]).toUpperCase().match(/[A-Z][A-Z0-9]*/g)?.[0];
    if (name && name !== "HYPERLINK") fieldNames.set(name, (fieldNames.get(name) || 0) + 1);
  }
  const fieldTotal = [...fieldNames.values()].reduce((a, b) => a + b, 0);
  add(
    "域代码（页码/目录/交叉引用）",
    fieldTotal ? `${fieldTotal} 处（${[...fieldNames.keys()].join("、")}）` : "0",
    fieldTotal === 0 ? "无" : reported("域代码结果") ? "已报告" : "静默丢失",
    `报告 ${reportedCount("域代码结果")} 处`,
  );

  /* —— 其他可能被忽略的结构 —— */
  const shapes = count(documentXml, /<w:pict\b/g) + count(documentXml, /<mc:AlternateContent\b/g);
  add("图形/兼容块", `${shapes} 处`, shapes === 0 ? "无" : "未单独归类", "含文本框、WordArt、形状；其文字已并入正文/文本框检查");
  const noBreakHyphen = count(documentXml, /<w:noBreakHyphen\/?>/g);
  const tabs = count(documentXml, /<w:tab\/?>/g);
  add("连字符/制表符", `${noBreakHyphen} 个连字符、${tabs} 个制表符`, noBreakHyphen + tabs === 0 ? "无" : "未导入未报告", "Word 里可见的分隔符，Markdown 里没有对应字符");

  return { 文件: path.basename(file), 行: rows, 丢失清单: lossLabels };
}

const targets = fs
  .readdirSync(root)
  .filter((name) => name.toLowerCase().endsWith(".docx") && !name.startsWith("~$"))
  .map((name) => path.join(root, name));
if (process.env.DAORU_EXTRA_DOCX) {
  targets.push(...process.env.DAORU_EXTRA_DOCX.split(path.delimiter).filter((p) => fs.existsSync(p)));
}

let silentTotal = 0;
for (const target of targets) {
  const result = await audit(target);
  console.log(`\n════ ${result.文件} ════`);
  for (const row of result.行) {
    const mark = row.结局 === "静默丢失" ? "❗" : row.结局 === "未导入未报告" ? "⚠️" : row.结局 === "未单独归类" ? "·" : " ";
    console.log(`${mark} ${row.类别.padEnd(18)} 有:${String(row.文档里有).padEnd(22)} → ${row.结局}${row.说明 ? `  （${row.说明}）` : ""}`);
    if (row.结局 === "静默丢失") silentTotal++;
  }
}
console.log(`\n════ 合计静默丢失条目：${silentTotal} ════`);
