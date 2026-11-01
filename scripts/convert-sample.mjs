import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import esbuild from "esbuild";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outRoot = path.join(root, ".dist-test");
const bundlePath = path.join(outRoot, "core.cjs");
const outDir = path.join(outRoot, "out");

fs.mkdirSync(outDir, { recursive: true });

await esbuild.build({
  stdin: {
    contents: [
      'export { convertDocx } from "./src/core/docx";',
      'export { htmlToMarkdown } from "./src/core/markdown";',
      'export { relativePathFrom } from "./src/core/relative-path";',
      'export { describeError } from "./src/core/errors";',
      'export { checkTextFidelity } from "./src/core/losses";',
    ].join("\n"),
    resolveDir: root,
    sourcefile: "test-entry.ts",
    loader: "ts",
  },
  bundle: true,
  platform: "node",
  format: "cjs",
  target: "node18",
  alias: {
    jszip: path.join(root, "node_modules/jszip/lib/index.js"),
    immediate: path.join(root, "src/shims/immediate.cjs"),
    setimmediate: path.join(root, "src/shims/setimmediate.ts"),
  },
  outfile: bundlePath,
  logLevel: "silent",
});

const require = createRequire(import.meta.url);
const { convertDocx, htmlToMarkdown, relativePathFrom, describeError, checkTextFidelity } =
  require(bundlePath);
const commonmark = require("commonmark");
const JSZip = require("jszip");

const samples = fs
  .readdirSync(root)
  .filter((name) => name.toLowerCase().endsWith(".docx"))
  .filter((name) => !name.startsWith("~$")) // Word 打开文档时生成的锁文件，不是真文档
  .sort();

if (samples.length === 0) {
  throw new Error("项目根目录没有 .docx 样本。");
}

const failures = [];
const check = (condition, message) => {
  if (!condition) failures.push(message);
};

function toArrayBuffer(buffer) {
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

function image(token, link, previewable = true) {
  return {
    token,
    data: new ArrayBuffer(0),
    contentType: "image/png",
    extension: "png",
    previewable,
    link,
  };
}

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

/** 用 CommonMark 参考实现解析输出，证明语法是标准的、不依赖 Obsidian */
function parseMarkdown(markdown) {
  const document = new commonmark.Parser().parse(markdown);
  const images = [];
  const links = [];
  const literal = [];
  // 规范实现会把地址里的中文百分号编码，比对前先还原
  const decode = (value) => {
    try {
      return decodeURIComponent(value || "");
    } catch {
      return value || "";
    }
  };
  let event;
  const walker = document.walker();
  while ((event = walker.next())) {
    const node = event.node;
    if (!event.entering) continue;
    if (node.type === "image") images.push({ destination: decode(node.destination), title: node.title });
    if (node.type === "link") links.push({ destination: decode(node.destination) });
    if (node.type === "text" && /\[\[|\]\]/.test(node.literal || "")) {
      literal.push(node.literal);
    }
  }
  return { images, links, literal };
}

for (const fileName of samples) {
  const source = path.join(root, fileName);
  const data = toArrayBuffer(fs.readFileSync(source));
  const base = fileName.replace(/\.(docx|docm)$/i, "");
  const started = performance.now();
  const conversion = await convertDocx(data);
  const afterConvert = performance.now();
  conversion.images.forEach((image, index) => {
    image.link = `附件/${base}-${String(index + 1).padStart(3, "0")}.${image.extension}`;
  });
  const markdownResult = htmlToMarkdown(conversion.html, conversion.images);
  const finished = performance.now();

  const markdownPath = path.join(outDir, `${base}.md`);
  fs.writeFileSync(markdownPath, markdownResult.markdown, "utf8");

  const parsed = parseMarkdown(markdownResult.markdown);
  const headings = (markdownResult.markdown.match(/^#{1,6} /gm) || []).length;
  const withAlt = conversion.images.filter((image) => image.altText).length;
  const nonPreviewable = conversion.images.filter((image) => !image.previewable).length;

  console.log(`\n=== ${fileName} ===`);
  console.log(
    `解析 ${(afterConvert - started).toFixed(0)}ms · 转 Markdown ${(finished - afterConvert).toFixed(0)}ms · 输出 ${formatBytes(Buffer.byteLength(markdownResult.markdown))}`,
  );
  console.log(
    `标题 ${headings} · 图片 ${conversion.images.length}（带替代文字 ${withAlt}，不可预览 ${nonPreviewable}） · 表格 简单 ${markdownResult.simpleTables} / HTML ${markdownResult.complexTables} · 丢失清单 ${conversion.losses.length + markdownResult.losses.length} 项`,
  );
  console.log(
    `CommonMark 解析出图片 ${parsed.images.length} 个、链接 ${parsed.links.length} 个、残留 wikilink 文本 ${parsed.literal.length} 处`,
  );
  const guardStarted = performance.now();
  const guard = checkTextFidelity(conversion.documentXml, markdownResult.markdown);
  const guardMs = performance.now() - guardStarted;
  console.log(
    `兜底自检 ${guard ? "报警" : "未报警"} · 逐段对账耗时 ${guardMs.toFixed(0)}ms`,
  );
  check(
    !guard,
    `${fileName}: 兜底自检误报（${guard?.detail || ""}）`,
  );
  console.log(`警告 ${conversion.warnings.length + markdownResult.warnings.length}`);
  for (const warning of [...conversion.warnings, ...markdownResult.warnings]) {
    console.log(`  - ${warning}`);
  }
  console.log("开头：");
  console.log(
    markdownResult.markdown
      .split("\n")
      .slice(0, 12)
      .join("\n"),
  );

  check(conversion.html.length > 0, `${fileName}: HTML 为空`);
  check(markdownResult.markdown.length > 100, `${fileName}: Markdown 太短`);
  check(
    !markdownResult.markdown.includes("sun-daoru-image-"),
    `${fileName}: 仍有图片占位符`,
  );
  check(
    !/<img[^>]+sun-daoru-image-/.test(markdownResult.markdown),
    `${fileName}: HTML 表格内仍有图片占位符`,
  );
  check(
    parsed.literal.length === 0,
    `${fileName}: 正文里残留 Obsidian 专有 wikilink 语法`,
    parsed.literal.slice(0, 3).join(" | "),
  );
  check(
    !markdownResult.markdown.includes("[[") &&
      !markdownResult.markdown.includes("]]"),
    `${fileName}: 输出里不应再有 [[ 或 ]]`,
  );

  if (fileName.includes("功能营养素手册")) {
    check(headings > 1500, `${fileName}: 动态标题映射疑似失效（标题太少）`);
    check(conversion.images.length === 149, `${fileName}: 图片数量不符（${conversion.images.length}）`);
    check(withAlt === 102, `${fileName}: 图片替代文字数量不符（${withAlt}）`);
    check(
      markdownResult.markdown.includes(
        "![维生素D核受体的分布](附件/功能营养素手册-007.png)",
      ),
      `${fileName}: 替代文字没写进标准图片语法`,
    );
    check(
      parsed.images.length === 149,
      `${fileName}: CommonMark 只解析出 ${parsed.images.length} 个图片节点`,
    );
    check(
      parsed.images.every((image) =>
        /^附件\/功能营养素手册-\d{3}\.[a-z0-9]+$/i.test(image.destination || ""),
      ),
      `${fileName}: 有图片地址解析结果不对`,
    );
    check(
      !/!\[\[[^\]|]*\|\d+\]\]/.test(markdownResult.markdown),
      `${fileName}: 仍在输出图片宽度`,
    );
    check(
      markdownResult.markdown.startsWith("# 营养成分手册使用须知\n\n# "),
      `${fileName}: 文件标题与第一章被粘连`,
    );
    check(conversion.stats.tables === 1, `${fileName}: 表格数量不符`);

    const normalized = await convertDocx(data.slice(0), {
      normalizeHeadingLevels: true,
    });
    normalized.images.forEach((image, index) => {
      image.link = `附件/规范-${index + 1}.${image.extension}`;
    });
    const normalizedMd = htmlToMarkdown(
      normalized.html,
      normalized.images,
    ).markdown;
    const normalizedH1 = (normalizedMd.match(/^# /gm) || []).length;
    check(normalized.headingLevelsShifted, `${fileName}: 未标记标题层级已调整`);
    check(
      normalizedMd.startsWith(
        "# 营养成分手册使用须知\n\n## 一、关于参考摄入量DRIs",
      ),
      `${fileName}: 标题下移后章节层级不正确`,
    );
    check(
      normalizedH1 > 0 && normalizedH1 < 150,
      `${fileName}: 标题下移后 H1 数量异常（${normalizedH1}）`,
    );
  }

  if (fileName.includes("妊娠期糖尿病")) {
    check(markdownResult.complexTables === 9, `${fileName}: 复杂表格数量应为 9`);
    check(markdownResult.simpleTables === 1, `${fileName}: 简单表格数量应为 1`);
    check(
      markdownResult.markdown.includes("<table>") &&
        markdownResult.markdown.includes('colspan="4"'),
      `${fileName}: 复杂表格未保留 HTML/合并单元格`,
    );
    check(
      markdownResult.markdown.includes("| **分类** | **优选（绿灯）食物**"),
      `${fileName}: 简单表格未转成 GFM`,
    );
    check(headings >= 14, `${fileName}: 标题数量过少`);
    check(
      markdownResult.markdown.startsWith("# 妊娠期糖尿病营养干预方案\n\n# 一、体重监控"),
      `${fileName}: 文件标题与第一章被粘连`,
    );

    const normalized = await convertDocx(data.slice(0), {
      normalizeHeadingLevels: true,
    });
    const normalizedMd = htmlToMarkdown(
      normalized.html,
      normalized.images,
    ).markdown;
    check(normalized.headingLevelsShifted, `${fileName}: 未标记标题层级已调整`);
    check(
      normalizedMd.startsWith(
        "# 妊娠期糖尿病营养干预方案\n\n## 一、体重监控",
      ),
      `${fileName}: 标题下移后章节层级不正确`,
    );
    check(
      (normalizedMd.match(/^# /gm) || []).length === 1,
      `${fileName}: 标题下移后 H1 数量应只剩文件标题`,
    );
  }

  if (fileName.startsWith("3我热天")) {
    check(markdownResult.complexTables === 2, `${fileName}: 复杂表格数量应为 2`);
    check(markdownResult.simpleTables === 1, `${fileName}: 简单表格数量应为 1`);
    check(
      markdownResult.markdown.includes("| 孕前体重分类 | BMI（kg/m<sup>2</sup>） |"),
      `${fileName}: 简单表格未转成 GFM`,
    );
    check(
      markdownResult.markdown.includes("<ol>") &&
        markdownResult.markdown.includes("<table>"),
      `${fileName}: 表格内的嵌套列表未保留`,
    );
    check(headings === 18, `${fileName}: 标题数量不符（${headings}）`);
  }
}

const simpleTable = htmlToMarkdown(
  "<table><thead><tr><th>姓名</th><th>数量</th></tr></thead><tbody><tr><td>甲</td><td>a|b</td></tr></tbody></table>",
  [],
);
check(
  simpleTable.markdown.includes("| 姓名 | 数量 |") &&
    simpleTable.markdown.includes("| --- | --- |") &&
    simpleTable.markdown.includes("a\\|b"),
  "简单表格未正确转成 GFM 表格或未转义竖线",
);
check(simpleTable.simpleTables === 1, "简单表格计数不正确");

const complexTable = htmlToMarkdown(
  '<table><tr><td colspan="2">标题</td></tr><tr><td>甲</td><td>乙</td></tr></table>',
  [],
);
check(
  complexTable.markdown.includes("<table>") &&
    complexTable.markdown.includes('colspan="2"'),
  "合并单元格表格未保留为 HTML",
);
check(complexTable.complexTables === 1, "复杂表格计数不正确");

const tableImage = htmlToMarkdown(
  '<table><tr><td colspan="2"><img src="sun-daoru-image-0" alt="表格内 PNG"></td></tr><tr><td>甲</td><td>乙</td></tr></table>',
  [
    {
      token: "sun-daoru-image-0",
      data: new ArrayBuffer(0),
      contentType: "image/png",
      extension: "png",
      previewable: true,
      link: "附件/示例.png",
    },
  ],
);
check(
  tableImage.markdown.includes('<img src="附件/示例.png" alt="表格内 PNG">') &&
    tableImage.markdown.includes("<table>"),
  "复杂表格内的图片未改成可渲染的 img 标签",
);
check(
  tableImage.losses.some((loss) => loss.label === "HTML 表格里的图片"),
  "复杂表格图片未写进丢失清单",
);

const altImage = htmlToMarkdown(
  '<p><img src="sun-daoru-image-0" alt="维生素D核受体的分布"></p>',
  [image("sun-daoru-image-0", "功能营养素手册-附件/007.png", true)],
);
check(
  altImage.markdown.trim() === "![维生素D核受体的分布](功能营养素手册-附件/007.png)",
  `替代文字没写进标准图片语法：${altImage.markdown.trim()}`,
);

const numericAlt = htmlToMarkdown(
  '<p><img src="sun-daoru-image-0" alt="120"></p>',
  [image("sun-daoru-image-0", "附件/示例.png", true)],
);
check(
  numericAlt.markdown.includes('<img src="附件/示例.png" alt="120">'),
  "纯数字替代文字没退回 img 标签",
);

const noAlt = htmlToMarkdown('<p><img src="sun-daoru-image-0"></p>', [
  image("sun-daoru-image-0", "附件/示例.png", true),
]);
check(noAlt.markdown.trim() === "![](附件/示例.png)", `无替代文字时输出不对：${noAlt.markdown.trim()}`);

const spacedImage = htmlToMarkdown(
  '<p><img src="sun-daoru-image-0" alt="图注"></p>',
  [image("sun-daoru-image-0", "附件/导入/我的 图.png", true)],
);
check(
  spacedImage.markdown.trim() === "![图注](<附件/导入/我的 图.png>)",
  `带空格路径没包尖括号：${spacedImage.markdown.trim()}`,
);
check(
  parseMarkdown(spacedImage.markdown).images[0]?.destination === "附件/导入/我的 图.png",
  "尖括号地址未被 CommonMark 正确解析",
);

const nestedLink = htmlToMarkdown(
  '<p><a href="https://example.com"><a href="https://example.com"><img src="sun-daoru-image-0" alt="超链接图片"></a></a></p>',
  [image("sun-daoru-image-0", "附件/示例.png", true)],
);
check(
  nestedLink.markdown.trim() === "[![超链接图片](附件/示例.png)](https://example.com)",
  `图片外层重复链接未合并：${nestedLink.markdown.trim()}`,
);
check(
  parseMarkdown(nestedLink.markdown).images.length === 1,
  "链接里的图片没被 CommonMark 解析成图片节点",
);

const internalLink = htmlToMarkdown(
  '<p>见 <a href="#bm1">跳回<sup>1</sup></a> 和 <a href="#missing">失效跳转</a></p>',
  [],
);
check(
  internalLink.markdown.includes("跳回") && internalLink.markdown.includes("失效跳转"),
  `文档内跳转的文字丢了：${internalLink.markdown.trim()}`,
);
check(
  !internalLink.markdown.includes("](#"),
  "文档内跳转仍留下了锚点地址",
);
check(
  parseMarkdown(internalLink.markdown).links.length === 0,
  "文档内跳转被解析成了链接节点",
);

check(relativePathFrom("笔记/子目录", "笔记/子目录/附件/001.png") === "附件/001.png", "同目录相对路径算错");
check(relativePathFrom("笔记/子目录", "笔记/其他/001.png") === "../其他/001.png", "上级相对路径算错");
check(relativePathFrom("", "附件/导入/001.png") === "附件/导入/001.png", "库根相对路径算错");
check(relativePathFrom("a/b/c", "x/y/001.png") === "../../../x/y/001.png", "跨分支相对路径算错");

const spacedHref = htmlToMarkdown(
  '<p><a href="https://example.com/a b(c)">带空格</a></p>',
  [],
);
check(
  spacedHref.markdown.includes("[带空格](https://example.com/a%20b%28c%29)"),
  "含空格和括号的链接地址未转义",
);

const unsafe = htmlToMarkdown(
  '<p><a href="javascript:alert(1)">点我</a> <a href="https://example.com">正常</a></p>',
  [],
);
check(!unsafe.markdown.includes("javascript:"), "危险链接未被过滤");
check(unsafe.markdown.includes("点我"), "危险链接的文字丢失");
check(unsafe.markdown.includes("[正常](https://example.com)"), "正常链接丢失");

const formatting = htmlToMarkdown(
  "<p><u>下划线</u> x<sup>2</sup> H<sub>2</sub>O</p>",
  [],
);
check(formatting.markdown.includes("<u>下划线</u>"), "下划线丢失");
check(formatting.markdown.includes("<sup>2</sup>"), "上标丢失");
check(formatting.markdown.includes("<sub>2</sub>"), "下标丢失");

const highlight = htmlToMarkdown("<p><mark>底色</mark></p>", []);
check(
  !highlight.markdown.includes("=="),
  `不该再输出 Obsidian 专有的 ==高亮==：${highlight.markdown.trim()}`,
);

const escaping = htmlToMarkdown(
  "<h1>1 &lt; 2 &amp; 3</h1><p>a &lt; b 和 <code>x &lt; y</code> 和 <b>粗</b></p>",
  [],
);
check(
  escaping.markdown.includes("1 &lt; 2") && escaping.markdown.includes("a &lt; b"),
  `正文里的 < 未转义：${escaping.markdown.trim()}`,
);
check(
  escaping.markdown.includes("`x < y`"),
  `代码里的 < 不该被转义：${escaping.markdown.trim()}`,
);
check(
  escaping.markdown.includes("<b>粗</b>") || escaping.markdown.includes("**粗**"),
  "标签本身被错误转义",
);

const codeBlock = htmlToMarkdown("<pre>let x = 1;\nif (a &lt; b) {}</pre>", []);
check(
  codeBlock.markdown.includes("```") &&
    codeBlock.markdown.includes("let x = 1;"),
  `pre 段落没转成围栏代码块：${JSON.stringify(codeBlock.markdown)}`,
);

/** 造一个最小 docx，用来验证丢失清单里的新条目 */
async function minimalDocx(bodyXml, extraRels = "", extraFiles = {}) {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><w:body>${bodyXml}<w:sectPr/></w:body></w:document>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${extraRels}</Relationships>`,
  );
  for (const [name, content] of Object.entries(extraFiles)) zip.file(name, content);
  return toArrayBuffer(await zip.generateAsync({ type: "nodebuffer" }));
}

const drawing = (inner) =>
  `<w:p><w:r><w:t>正文</w:t></w:r><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="图1" descr="链接图"/>${inner}</wp:inline></w:drawing></w:r></w:p>`;

const linkedImage = await convertDocx(
  await minimalDocx(
    drawing(
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:link="rId9"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>',
    ),
    '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="C:/pics/x.png" TargetMode="External"/>',
  ),
);
check(
  linkedImage.losses.some(
    (loss) => loss.label === "链接到文件的图片" && loss.count === 1 && loss.kind === "lost",
  ),
  `外部链接图片没进丢失清单：${JSON.stringify(linkedImage.losses.map((l) => l.label))}`,
);
check(
  !linkedImage.warnings.some((warning) => warning.includes("读不出内容")),
  `外部链接图片被重复报成读取失败：${JSON.stringify(linkedImage.warnings)}`,
);

const png = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DABQAM/wDf0IvLDQAAAABJRU5ErkJggg==",
  "base64",
);

const embedded = await convertDocx(
  await minimalDocx(
    `<w:p><w:r><w:t>正文</w:t></w:r><w:r><w:object><v:shape><v:imagedata r:id="rId10"/></v:shape><o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" r:id="rId11"/></w:object></w:r></w:p>` +
      drawing(
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId12"/></a:graphicData></a:graphic>',
      ),
    '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/><Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="C:/tmp/book.xlsx" TargetMode="External"/><Relationship Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>',
    { "word/media/image1.png": png },
  ),
);
check(
  embedded.losses.some(
    (loss) => loss.label === "图表与嵌入对象" && loss.count >= 2,
  ),
  `图表/嵌入对象没进丢失清单：${JSON.stringify(embedded.losses.map((l) => [l.label, l.count]))}`,
);

// 引用悬空（图片关系指向不存在的内容）时，底层会抛英文 TypeError，必须翻译成人话
let danglingError = "";
try {
  await convertDocx(
    await minimalDocx(
      drawing(
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdMissing"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>',
      ),
    ),
  );
} catch (error) {
  danglingError = describeError(error);
}
check(Boolean(danglingError), "悬空引用没有抛错，样本失效了");
check(
  /[一-鿿]/.test(danglingError) && !/^Cannot read/.test(danglingError),
  `底层报错没被翻译：${danglingError}`,
);

const fidelityDoc = `<w:p><w:r><w:t>第一段完整的文字。</w:t></w:r></w:p><w:p><w:r><w:t>第二段会被弄丢的文字。</w:t></w:r></w:p><w:p><w:r><w:t>重复出现的句子。</w:t></w:r></w:p><w:p><w:r><w:t>重复出现的句子。</w:t></w:r></w:p>`;
const fidelityAll = "第一段完整的文字。\n\n第二段会被弄丢的文字。\n\n重复出现的句子。\n\n重复出现的句子。";
check(
  checkTextFidelity(fidelityDoc, fidelityAll) === null,
  "文字齐全时兜底自检误报",
);
const fidelityLost = checkTextFidelity(
  fidelityDoc,
  fidelityAll.replace("第二段会被弄丢的文字。", ""),
);
check(
  fidelityLost?.label === "疑似未识别的内容丢失" &&
    fidelityLost.detail.includes("第二段会被弄丢的文字。"),
  `段落缺失时没报或没指出是哪段：${JSON.stringify(fidelityLost)}`,
);
const fidelityShrunk = checkTextFidelity(
  fidelityDoc,
  fidelityAll.replace("\n\n重复出现的句子。", ""),
);
check(
  fidelityShrunk?.detail.includes("出现次数比原文档少"),
  `重复次数变少没被抓到：${JSON.stringify(fidelityShrunk)}`,
);

// VML 文本框里的文字：实测 mammoth 会读出来，这里当正向断言守住
const longBody = "这是正文内容。".repeat(60);
const boxText = "这是文本框里的文字。".repeat(120);
const textboxHeavy = await convertDocx(
  await minimalDocx(
    `<w:p><w:r><w:t>${longBody}</w:t></w:r></w:p>` +
      `<w:p><w:r><w:pict><v:shape id="s1" type="#_x0000_t202" style="width:100pt;height:50pt"><v:textbox inset="0,0,0,0"><w:txbxContent><w:p><w:r><w:t>${boxText}</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>`,
  ),
);
textboxHeavy.images.forEach((image, index) => {
  image.link = `附件/重-${index}.${image.extension}`;
});
const heavyMarkdown = htmlToMarkdown(textboxHeavy.html, textboxHeavy.images);
check(
  heavyMarkdown.markdown.includes("这是文本框里的文字"),
  "文本框里的文字没进正文",
);
check(
  checkTextFidelity(textboxHeavy.documentXml, heavyMarkdown.markdown) === null,
  "文本框文字已导入，兜底自检不该响",
);

if (failures.length > 0) {
  console.error("\n回归失败：");
  for (const failure of failures) console.error(`  - ${failure}`);
  process.exit(1);
}

console.log(`\n回归通过：${samples.length} 份样本 + 16 个合成用例 + 4 条相对路径检查 + 标题下移检查。输出目录：${outDir}`);
