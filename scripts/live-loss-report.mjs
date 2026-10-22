/*
丢失清单与链接/表格图片的实机回归。

需要一个「什么都带一点」的烤测文档（批注、公式、页眉页脚、书签跳转、
超链接图片、合并单元格里的图片）。默认路径在系统临时目录，可用环境变量
DAORU_AUDIT_FIXTURE 指定；文件不存在时整个用例跳过，不算失败。

  DAORU_AUDIT_FIXTURE=/path/to/doc.docx node ../sun-home/scripts/cdp/run.mjs scripts/live-loss-report.mjs

切阅读视图必须用真鼠标点标题栏那个模式按钮：markdown:toggle-preview 命令
和 view.setMode()/toggleMode() 在 1.14.4 里都会静默失败或把视图弄坏。
*/
import fs from "node:fs";
import path from "node:path";

const DEFAULT_FIXTURE = path.join(
  process.env.LOCALAPPDATA || "",
  "Temp",
  "daoru-audit",
  "audit-comprehensive.docx",
);

export default async function (ctx) {
  const fixture = process.env.DAORU_AUDIT_FIXTURE || DEFAULT_FIXTURE;
  if (!fs.existsSync(fixture)) {
    return { 汇总: `跳过：找不到烤测文档 ${fixture}`, checks: [] };
  }

  const imported = await ctx.page(async (fixturePath) => {
    const checks = [];
    const check = (ok, name, detail = "") =>
      checks.push({ 检查: name, 结论: ok ? "PASS" : "FAIL", 详情: String(detail).slice(0, 200) });
    const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
    const waitFor = async (predicate, timeout, label) => {
      const started = Date.now();
      while (Date.now() - started < timeout) {
        const value = await predicate();
        if (value) return value;
        await sleep(150);
      }
      throw new Error(`等待超时：${label}`);
    };

    const app = window.app;
    const plugin = app.plugins.plugins["sun-daoru"];
    const notePath = "__daoru自测__/audit-comprehensive.md";

    // 上一轮留下的结果弹窗会一直堆在 DOM 里，必须先点掉，否则会读到旧内容
    const closeStaleModals = async () => {
      for (let i = 0; i < 10; i++) {
        const button = Array.from(
          document.querySelectorAll(".sun-daoru-result-modal button"),
        ).find((item) => item.textContent?.trim() === "完成");
        if (!button) return;
        button.click();
        await sleep(200);
      }
    };
    const latestResultModal = () => {
      const modals = document.querySelectorAll(".sun-daoru-result-modal");
      return modals.length ? modals[modals.length - 1] : null;
    };

    await closeStaleModals();
    check(
      document.querySelectorAll(".sun-daoru-result-modal").length === 0,
      "测试开始前没有残留弹窗",
      String(document.querySelectorAll(".sun-daoru-result-modal").length),
    );
    await waitFor(async () => !plugin.importing, 120000, "上一批导入结束");

    const oldSandbox = app.vault.getAbstractFileByPath("__daoru自测__");
    if (oldSandbox) await app.vault.delete(oldSandbox, true);
    await app.vault.createFolder("__daoru自测__");
    await app.vault.create("__daoru自测__/导入测试.md", "测试用笔记。\n");
    const sandboxNote = app.vault.getAbstractFileByPath("__daoru自测__/导入测试.md");
    await app.workspace.getLeaf(false).openFile(sandboxNote);
    await sleep(300);

    plugin.settings.groupAttachmentsByNote = true;
    plugin.settings.openAfterImport = false;
    plugin.settings.addSourceFrontmatter = false;
    plugin.settings.normalizeHeadingLevels = false;
    plugin.settings.defaultTarget = "new";
    await plugin.saveSettings();

    window.__daoruLive = {
      origClick: HTMLInputElement.prototype.click,
      files: [{ path: fixturePath, name: "audit-comprehensive.docx" }],
    };
    HTMLInputElement.prototype.click = function () {
      if (this.type !== "file") return window.__daoruLive.origClick.call(this);
      const transfer = new DataTransfer();
      for (const item of window.__daoruLive.files) {
        const buffer = require("fs").readFileSync(item.path);
        transfer.items.add(
          new File([buffer], item.name, {
            type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          }),
        );
      }
      this.files = transfer.files;
      this.dispatchEvent(new Event("change", { bubbles: true }));
    };

    let resultText = "";
    try {
      app.commands.executeCommandById("sun-daoru:import-word");
      const modal = await waitFor(
        () => document.querySelector(".sun-daoru-modal"),
        30000,
        "导入面板",
      );
      const start = Array.from(modal.querySelectorAll("button")).find((button) =>
        button.textContent?.includes("开始导入"),
      );
      check(Boolean(start), "找到开始导入按钮");
      const before = document.querySelectorAll(".sun-daoru-result-modal").length;
      start?.click();
      check(
        await waitFor(async () => plugin.importing, 15000, "导入任务启动"),
        "导入任务已启动",
      );
      check(
        await waitFor(async () => !plugin.importing, 120000, "导入任务结束"),
        "导入任务结束（未卡住）",
      );
      const resultModal = await waitFor(
        () =>
          document.querySelectorAll(".sun-daoru-result-modal").length > before
            ? latestResultModal()
            : null,
        30000,
        "结果弹窗",
      );
      resultText = resultModal.textContent || "";
      check(
        document.querySelectorAll(".sun-daoru-result-modal").length === before + 1,
        "只出现一个新的结果弹窗",
        String(document.querySelectorAll(".sun-daoru-result-modal").length),
      );
    } finally {
      if (window.__daoruLive?.origClick) {
        HTMLInputElement.prototype.click = window.__daoruLive.origClick;
      }
      delete window.__daoruLive;
    }

    check(resultText.includes("成功导入 1 个文件"), "烤测文档导入成功", resultText);
    check(/有 \d+ 类内容没能导入/.test(resultText), "顶部汇总提示有内容丢失", resultText);
    check(/丢失\s*批注内容/.test(resultText), "结果面板列出批注丢失");
    check(/丢失\s*公式/.test(resultText), "结果面板列出公式丢失");
    check(/丢失\s*页眉\/页脚/.test(resultText), "结果面板列出页眉页脚丢失");
    check(/丢失\s*修订里被删掉的文字/.test(resultText), "结果面板列出被删修订文字");
    check(/丢失\s*文档内跳转/.test(resultText), "结果面板把文档内跳报告为丢失");
    check(/图片替代文字 ×\d+/.test(resultText), "结果面板统计替代文字数量");

    // 结果弹窗不关掉，后面切阅读视图会失败
    await closeStaleModals();

    const note = app.vault.getAbstractFileByPath(notePath);
    const sandboxFiles = app.vault
      .getFiles()
      .filter((file) => file.path.startsWith("__daoru自测__/"))
      .map((file) => file.path);
    let content = "";
    if (note) content = await app.vault.read(note);
    check(Boolean(note), "生成笔记", notePath);

    check(
      content.includes("[![超链接图片](audit-comprehensive-附件/009.png)](https://example.com)"),
      "超链接图片合成一个标准链接",
      (content.match(/.*009\.png.*/) || [""])[0],
    );
    check(!/\]\([^)]*\)\[!\[/.test(content), "不再出现重复嵌套的图片链接");
    check(
      /<img src="audit-comprehensive-附件\/010\.png" alt="表格内 PNG">/.test(content),
      "表格内图片是相对路径的 img 标签",
      (content.match(/<table>.*?<\/table>/s) || [""])[0].slice(0, 160),
    );
    check(
      content.includes("跳回书签") && content.includes("一、普通标题 Heading 1"),
      "文档内跳转的文字保留",
      (content.match(/.*跳回书签.*/) || [""])[0],
    );
    check(
      !content.includes("[["),
      "正文不含 Obsidian 专有链接语法",
      (content.match(/.*\[\[.*/) || [""])[0],
    );
    check(
      content.includes("![PNG audit image](audit-comprehensive-附件/001.png)"),
      "图片替代文字写进 alt",
    );
    check(!content.includes("这是批注内容"), "批注文字确实没混进正文");

    const leaf = app.workspace.getLeaf(false);
    await leaf.openFile(note);
    await sleep(600);
    const view = leaf.view;
    check(view.getMode?.() === "source", "笔记先在编辑视图打开", String(view.getMode?.()));
    const rect = view.modeButtonEl?.getBoundingClientRect?.();
    check(
      Boolean(rect && rect.width > 0),
      "找到阅读视图切换按钮",
      rect ? `${Math.round(rect.x)},${Math.round(rect.y)} ${Math.round(rect.width)}x${Math.round(rect.height)}` : "无",
    );

    return {
      checks,
      sandboxFiles,
      resultText: resultText.slice(0, 700),
      button: rect
        ? { x: Math.round(rect.x + rect.width / 2), y: Math.round(rect.y + rect.height / 2) }
        : null,
    };
  }, fixture);

  if (!imported.button) {
    return {
      checks: imported.checks,
      汇总: "失败：没能定位到阅读视图切换按钮",
      沙箱文件: imported.sandboxFiles,
    };
  }

  await ctx.click(imported.button.x, imported.button.y);
  await ctx.sleep(1400);

  const rendered = await ctx.page(async () => {
    const checks = [];
    const check = (ok, name, detail = "") =>
      checks.push({ 检查: name, 结论: ok ? "PASS" : "FAIL", 详情: String(detail).slice(0, 200) });
    const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

    const app = window.app;
    const plugin = app.plugins.plugins["sun-daoru"];

    const visible = Array.from(document.querySelectorAll(".markdown-preview-view"))
      .filter((el) => el.getBoundingClientRect().height > 0)
      .pop();
    check(Boolean(visible), "阅读视图已显示", String(document.querySelectorAll(".markdown-preview-view").length));

    let scroller = visible;
    while (
      scroller &&
      scroller.parentElement &&
      scroller.scrollHeight <= scroller.clientHeight + 20
    ) {
      scroller = scroller.parentElement;
    }
    const 滚动容器 = {
      类名: String(scroller?.className).slice(0, 70),
      scrollHeight: scroller?.scrollHeight,
      clientHeight: scroller?.clientHeight,
    };
    const isOurs = (image) =>
      (image.getAttribute("src") || "").includes("audit-comprehensive");
    const step = Math.max(200, Math.floor((scroller.clientHeight || 400) * 0.6));
    // 滚回顶部后离屏的嵌入会被卸载，所以边滚边按 src 累积记录
    const seen = new Map();
    const seenLinks = new Map();
    const seenExternal = new Map();
    for (let y = 0; y <= scroller.scrollHeight; y += step) {
      scroller.scrollTop = y;
      await sleep(350);
      for (const image of scroller.querySelectorAll("img")) {
        if (isOurs(image)) seen.set(image.getAttribute("src"), image);
      }
      for (const link of scroller.querySelectorAll("a.internal-link")) {
        seenLinks.set(`${link.getAttribute("href")}|${link.textContent}`, link);
      }
      for (const link of scroller.querySelectorAll("a[href]")) {
        const href = link.getAttribute("href") || "";
        if (href.startsWith("https://example.com")) seenExternal.set(href, link);
      }
    }
    const images = Array.from(seen.values());
    const internalLinks = Array.from(seenLinks.values());
    await Promise.all(images.map((image) => image.decode?.().catch(() => undefined)));
    await sleep(300);
    const ok = images.filter((image) => image.complete && image.naturalWidth > 0);
    const tableImage = images.find((image) =>
      (image.getAttribute("alt") || "").includes("表格内"),
    );
    const linkedImage = images.find((image) =>
      (image.getAttribute("alt") || "").includes("超链接"),
    );

    check(images.length === 10, "10 张图片都进入渲染", String(images.length));
    check(ok.length === images.length && ok.length >= 10, "阅读视图全部出图", `${ok.length}/${images.length}`);
    check(
      Boolean(tableImage && tableImage.closest("td") && tableImage.naturalWidth > 0),
      "HTML 表格里的图片能直接显示",
      tableImage ? `${tableImage.naturalWidth}x${tableImage.naturalHeight}` : "未找到",
    );
    check(
      Boolean(linkedImage?.closest("a")?.getAttribute("href")?.startsWith("https://example.com")),
      "超链接图片可点击跳转",
      linkedImage?.closest("a")?.getAttribute("href") || "未找到",
    );
    check(
      internalLinks.length === 0,
      "正文不再产生 Obsidian 专有内部链接",
      internalLinks.map((link) => link.textContent.trim()).slice(0, 4).join(" / "),
    );
    const externalLinks = Array.from(seenExternal.values());
    check(
      externalLinks.length >= 2,
      "外部链接与图片链接都渲染为可点击地址",
      `${externalLinks.length} 个`,
    );

    const sandbox = app.vault.getAbstractFileByPath("__daoru自测__");
    if (sandbox) await app.vault.delete(sandbox, true);
    plugin.settings.openAfterImport = true;
    await plugin.saveSettings();
    const welcome = app.vault.getAbstractFileByPath("欢迎.md");
    if (welcome) await app.workspace.getLeaf(false).openFile(welcome);
    await sleep(300);

    return { checks, 滚动容器, 图片总数: images.length, 渲染成功: ok.length, remaining: app.vault.getFiles().map((file) => file.path) };
  });

  const checks = [...imported.checks, ...rendered.checks];
  const failed = checks.filter((item) => item.结论 === "FAIL");
  return {
    checks,
    汇总: failed.length === 0 ? `全部通过（${checks.length} 项）` : `失败 ${failed.length} 项`,
    图片总数: rendered.图片总数,
    滚动容器: rendered.滚动容器,
    渲染成功: rendered.渲染成功,
    沙箱文件: imported.sandboxFiles,
    清理后库内文件: rendered.remaining,
  };
}
