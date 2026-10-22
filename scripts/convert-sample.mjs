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
    ].join("\n"),
    resolveDir: root,
    sourcefile: "test-entry.ts",
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

const samples = fs
  .readdirSync(root)
  .filter((name) => name.toLowerCase().endsWith(".docx"))
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

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
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

  const headings = (markdownResult.markdown.match(/^#{1,6} /gm) || []).length;
  const widths = conversion.images.filter((image) => image.widthPx).length;
  const nonPreviewable = conversion.images.filter((image) => !image.previewable).length;

  console.log(`\n=== ${fileName} ===`);
  console.log(
    `解析 ${(afterConvert - started).toFixed(0)}ms · 转 Markdown ${(finished - afterConvert).toFixed(0)}ms · 输出 ${formatBytes(Buffer.byteLength(markdownResult.markdown))}`,
  );
  console.log(
    `标题 ${headings} · 图片 ${conversion.images.length}（带宽度 ${widths}，不可预览 ${nonPreviewable}） · 表格 简单 ${markdownResult.simpleTables} / HTML ${markdownResult.complexTables}`,
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

  if (fileName.includes("功能营养素手册")) {
    check(headings > 1500, `${fileName}: 动态标题映射疑似失效（标题太少）`);
    check(conversion.images.length === 149, `${fileName}: 图片数量不符（${conversion.images.length}）`);
    check(widths === 149, `${fileName}: 图片显示宽度未全部读取`);
    check(
      /!\[\[附件\/[^|\]]+\|698\]\]/.test(markdownResult.markdown),
      `${fileName}: 未输出图片宽度`,
    );
    check(
      markdownResult.markdown.startsWith("# 营养成分手册使用须知\n\n# "),
      `${fileName}: 文件标题与第一章被粘连`,
    );
    check(conversion.stats.tables === 1, `${fileName}: 表格数量不符`);
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
  '<table><tr><td colspan="2"><img src="sun-daoru-image-0"></td></tr><tr><td>甲</td><td>乙</td></tr></table>',
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
  tableImage.markdown.includes("![[附件/示例.png]]") &&
    tableImage.markdown.includes("<table>"),
  "复杂表格内的图片未保留为嵌入语法",
);
check(
  tableImage.warnings.some((warning) => warning.includes("复杂表格内")),
  "复杂表格图片未产生提示",
);

const unsafe = htmlToMarkdown(
  '<p><a href="javascript:alert(1)">点我</a> <a href="https://example.com">正常</a></p>',
  [],
);
check(!unsafe.markdown.includes("javascript:"), "危险链接未被过滤");
check(unsafe.markdown.includes("点我"), "危险链接的文字丢失");
check(unsafe.markdown.includes("[正常](https://example.com)"), "正常链接丢失");

const formatting = htmlToMarkdown(
  "<p><u>下划线</u> x<sup>2</sup> H<sub>2</sub>O ==高亮==</p>",
  [],
);
check(formatting.markdown.includes("<u>下划线</u>"), "下划线丢失");
check(formatting.markdown.includes("<sup>2</sup>"), "上标丢失");
check(formatting.markdown.includes("<sub>2</sub>"), "下标丢失");
check(formatting.markdown.includes("==高亮=="), "高亮丢失");

if (failures.length > 0) {
  console.error("\n回归失败：");
  for (const failure of failures) console.error(`  - ${failure}`);
  process.exit(1);
}

console.log(`\n回归通过：${samples.length} 份样本 + 5 个合成用例。输出目录：${outDir}`);
