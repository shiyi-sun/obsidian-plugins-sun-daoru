import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const big = {
  path: path.join(root, "功能营养素手册.docx"),
  name: "功能营养素手册.docx",
};
const small = {
  path: path.join(root, "妊娠期糖尿病营养干预方案.docx"),
  name: "妊娠期糖尿病营养干预方案.docx",
};
const tiny = {
  path: path.join(root, "3我热天热顺丰到付到付电风扇电风扇发射点手动阀手动阀.docx"),
  name: "3我热天热顺丰到付到付电风扇电风扇发射点手动阀手动阀.docx",
};
const tinyBase = "3我热天热顺丰到付到付电风扇电风扇发射点手动阀手动阀";

const invalidDocx = path.join(os.tmpdir(), "daoru-invalid.docx");
const oldDoc = path.join(os.tmpdir(), "daoru-old.doc");
fs.writeFileSync(invalidDocx, "this is not a zip");
fs.writeFileSync(oldDoc, "this is not a legacy word file");

const optionsShot = path.join(os.tmpdir(), "daoru-import-options.png");
const resultShot = path.join(os.tmpdir(), "daoru-import-result.png");

export default async function (ctx) {
  const checks = [];
  const screenshots = [];

  const addChecks = (items) => checks.push(...items);
  const checkAll = () => {
    const failed = checks.filter((item) => item.结论 === "FAIL");
    if (failed.length > 0) {
      throw new Error(
        "实机断言失败：" + failed.map((item) => item.检查).join("；"),
      );
    }
  };

  try {
    ctx.note({ step: "setup+big import" });
    const step1 = await ctx.page(
      async (bigPath, bigName) => {
        const checks = [];
        const check = (ok, name, detail = "") =>
          checks.push({ 检查: name, 结论: ok ? "PASS" : "FAIL", 详情: detail });
        const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
        const waitFor = async (selector, timeout) => {
          const started = Date.now();
          while (Date.now() - started < timeout) {
            const element = document.querySelector(selector);
            if (element) return element;
            await sleep(100);
          }
          throw new Error(`等待超时：${selector}`);
        };
        const closeResult = () => {
          const done = Array.from(
            document.querySelectorAll(".sun-daoru-result-modal button"),
          ).find((button) => button.textContent?.trim() === "完成");
          done?.click();
        };

        const app = window.app;
        const plugin = app.plugins.plugins["sun-daoru"];
        const oldSandbox = app.vault.getAbstractFileByPath("__daoru自测__");
        if (oldSandbox) await app.vault.delete(oldSandbox, true);
        await app.vault.createFolder("__daoru自测__");
        await app.vault.create("__daoru自测__/导入测试.md", "测试用笔记，导入完成后删除。\n");
        await app.vault.create("__daoru自测__/插入目标.md", "插入前\n");
        const testNote = app.vault.getAbstractFileByPath("__daoru自测__/导入测试.md");
        await app.workspace.getLeaf(false).openFile(testNote);
        for (let i = 0; i < 20 && app.workspace.getActiveFile()?.path !== testNote.path; i++) {
          await sleep(100);
        }

        plugin.settings.attachmentMode = "obsidian";
        plugin.settings.customAttachmentFolder = "附件/导入";
        plugin.settings.groupAttachmentsByNote = true;
        plugin.settings.normalizeHeadingLevels = false;
        plugin.settings.openAfterImport = false;
        plugin.settings.defaultTarget = "new";
        plugin.settings.addSourceFrontmatter = false;
        await plugin.saveSettings();

        window.__daoruLive = {
          origClick: HTMLInputElement.prototype.click,
          files: [{ path: bigPath, name: bigName }],
        };
        HTMLInputElement.prototype.click = function () {
          if (this.type !== "file") return window.__daoruLive.origClick.call(this);
          const fs = require("fs");
          const transfer = new DataTransfer();
          for (const file of window.__daoruLive.files) {
            const buffer = fs.readFileSync(file.path);
            transfer.items.add(
              new File([buffer], file.name, {
                type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              }),
            );
          }
          this.files = transfer.files;
          this.dispatchEvent(new Event("change", { bubbles: true }));
        };

        app.commands.executeCommandById("sun-daoru:import-word");
        const modal = await waitFor(".sun-daoru-modal", 30000);
        await sleep(200);
        const text = modal.textContent || "";
        check(text.includes(bigName), "导入面板显示文件名", bigName);
        check(
          text.includes("__daoru自测__"),
          "默认目标文件夹是当前笔记所在文件夹",
        );
        check(
          text.includes("附件位置：笔记同目录的 <文档名>-附件/"),
          "导入面板显示按文档分组的附件位置",
        );
        check(
          app.workspace.getActiveFile()?.path === "__daoru自测__/导入测试.md",
          "导入前活动笔记正确",
          app.workspace.getActiveFile()?.path || "",
        );
        return { checks, modalText: text.slice(0, 800) };
      },
      big.path,
      big.name,
    );
    addChecks(step1.checks);
    screenshots.push(await ctx.screenshot(optionsShot));

    ctx.note({ step: "verify big import" });
    const step2 = await ctx.page(
      async () => {
        const checks = [];
        const check = (ok, name, detail = "") =>
          checks.push({ 检查: name, 结论: ok ? "PASS" : "FAIL", 详情: detail });
        const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
        const waitFor = async (selector, timeout) => {
          const started = Date.now();
          while (Date.now() - started < timeout) {
            const element = document.querySelector(selector);
            if (element) return element;
            await sleep(100);
          }
          throw new Error(`等待超时：${selector}`);
        };

        const app = window.app;
        const start = Array.from(
          document.querySelectorAll(".sun-daoru-modal button"),
        ).find((button) => button.textContent?.includes("开始导入"));
        check(Boolean(start), "找到开始导入按钮");
        if (!start) throw new Error("没有开始导入按钮");
        start.click();

        const resultModal = await waitFor(".sun-daoru-result-modal", 120000);
        await sleep(300);
        const resultText = resultModal.textContent || "";
        check(resultText.includes("成功"), "结果面板显示成功");
        check(resultText.includes("149 张图片"), "结果面板显示 149 张图片");

        const notePath = "__daoru自测__/功能营养素手册.md";
        const note = app.vault.getAbstractFileByPath(notePath);
        check(Boolean(note), "生成 Markdown 笔记", notePath);
        let content = "";
        if (note) content = await app.vault.read(note);
        check(
          content.startsWith("# 营养成分手册使用须知\n\n# 一、关于参考摄入量DRIs"),
          "文件标题与第一章正确拆分",
        );
        check(
          (content.match(/^# /gm) || []).length > 500,
          "动态标题映射生成了足够的一级标题",
          String((content.match(/^# /gm) || []).length),
        );
        check(
          content.includes(
            "![](功能营养素手册-附件/001.png)",
          ),
          "图片按相对路径用标准语法引用",
        );
        check(
          content.includes(
            "![维生素D核受体的分布](功能营养素手册-附件/007.png)",
          ),
          "Word 图片替代文字写进了 alt",
        );
        check(
          !content.includes("[[") && !content.includes("]]"),
          "正文不含 Obsidian 专有链接语法",
        );
        check(!content.includes("sun-daoru-image-"), "正文无图片占位符残留");
        const attachments = app.vault
          .getFiles()
          .filter((file) =>
            file.path.startsWith("__daoru自测__/功能营养素手册-附件/"),
          );
        check(
          attachments.length === 149,
          "149 张图片全部写入附件目录",
          String(attachments.length),
        );
        check(
          app.workspace.getActiveFile()?.path === "__daoru自测__/导入测试.md",
          "测试设置下未自动打开导入笔记",
          app.workspace.getActiveFile()?.path || "",
        );
        return {
          checks,
          resultText: resultText.slice(0, 1000),
          contentLength: content.length,
          attachments: attachments.length,
        };
      },
    );
    addChecks(step2.checks);
    screenshots.push(await ctx.screenshot(resultShot));

    ctx.note({ step: "small import with frontmatter" });
    const step3 = await ctx.page(
      async (smallPath, smallName) => {
        const checks = [];
        const check = (ok, name, detail = "") =>
          checks.push({ 检查: name, 结论: ok ? "PASS" : "FAIL", 详情: detail });
        const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
        const waitFor = async (selector, timeout) => {
          const started = Date.now();
          while (Date.now() - started < timeout) {
            const element = document.querySelector(selector);
            if (element) return element;
            await sleep(100);
          }
          throw new Error(`等待超时：${selector}`);
        };
        const closeResult = () => {
          const done = Array.from(
            document.querySelectorAll(".sun-daoru-result-modal button"),
          ).find((button) => button.textContent?.trim() === "完成");
          done?.click();
        };

        closeResult();
        await sleep(300);
        const app = window.app;
        window.__daoruLive.files = [{ path: smallPath, name: smallName }];
        app.commands.executeCommandById("sun-daoru:import-word");
        const modal = await waitFor(".sun-daoru-modal", 30000);
        const sourceSetting = Array.from(
          modal.querySelectorAll(".setting-item"),
        ).find((item) => item.textContent?.includes("写入来源信息"));
        const checkbox = sourceSetting?.querySelector('input[type="checkbox"]');
        check(Boolean(checkbox), "找到写入来源信息开关");
        if (checkbox && !checkbox.checked) {
          checkbox.click();
          await sleep(120);
        }
        const start = Array.from(modal.querySelectorAll("button")).find((button) =>
          button.textContent?.includes("开始导入"),
        );
        start?.click();
        const resultModal = await waitFor(".sun-daoru-result-modal", 60000);
        const resultText = resultModal.textContent || "";
        check(resultText.includes("成功"), "小文档导入成功");

        const notePath = "__daoru自测__/妊娠期糖尿病营养干预方案.md";
        const note = app.vault.getAbstractFileByPath(notePath);
        let content = "";
        if (note) content = await app.vault.read(note);
        check(Boolean(note), "生成小文档笔记", notePath);
        check(
          content.startsWith(
            '---\nsource: "妊娠期糖尿病营养干预方案.docx"\nimported: ',
          ),
          "frontmatter 写入 source/imported",
          content.slice(0, 90),
        );
        check(
          content.includes("---\n\n# 妊娠期糖尿病营养干预方案"),
          "frontmatter 与正文之间有空行",
        );
        check(
          content.includes("# 妊娠期糖尿病营养干预方案\n\n# 一、体重监控"),
          "标题与第一章正确拆分",
        );
        check(
          content.includes("| **分类** | **优选（绿灯）食物**"),
          "简单表格转成 GFM",
        );
        check(
          content.includes("<table>") && content.includes('colspan="4"'),
          "合并单元格表格保留 HTML",
        );
        closeResult();
        await sleep(200);
        return { checks, contentSnippet: content.slice(0, 400) };
      },
      small.path,
      small.name,
    );
    addChecks(step3.checks);

    ctx.note({ step: "insert into current note" });
    const step4 = await ctx.page(
      async (tinyPath, tinyName, tinyBase) => {
        const checks = [];
        const check = (ok, name, detail = "") =>
          checks.push({ 检查: name, 结论: ok ? "PASS" : "FAIL", 详情: detail });
        const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
        const waitFor = async (selector, timeout) => {
          const started = Date.now();
          while (Date.now() - started < timeout) {
            const element = document.querySelector(selector);
            if (element) return element;
            await sleep(100);
          }
          throw new Error(`等待超时：${selector}`);
        };
        const closeResult = () => {
          const done = Array.from(
            document.querySelectorAll(".sun-daoru-result-modal button"),
          ).find((button) => button.textContent?.trim() === "完成");
          done?.click();
        };

        closeResult();
        await sleep(200);
        const app = window.app;
        const target = app.vault.getAbstractFileByPath("__daoru自测__/插入目标.md");
        await app.workspace.getLeaf(false).openFile(target);
        for (let i = 0; i < 20 && app.workspace.getActiveFile()?.path !== target.path; i++) {
          await sleep(100);
        }
        const view = app.workspace.activeLeaf?.view;
        const editor = view?.editor;
        check(Boolean(editor), "插入目标笔记可编辑");
        editor.setCursor({
          line: editor.lineCount() - 1,
          ch: editor.getLine(editor.lineCount() - 1).length,
        });
        editor.focus();

        window.__daoruLive.files = [{ path: tinyPath, name: tinyName }];
        app.commands.executeCommandById("sun-daoru:import-word");
        const modal = await waitFor(".sun-daoru-modal", 30000);
        const select = modal.querySelector("select");
        check(Boolean(select), "找到导入方式下拉框");
        select.value = "insert";
        select.dispatchEvent(new Event("change", { bubbles: true }));
        await sleep(150);
        const start = Array.from(modal.querySelectorAll("button")).find((button) =>
          button.textContent?.includes("开始导入"),
        );
        start?.click();

        const resultModal = await waitFor(".sun-daoru-result-modal", 60000);
        const resultText = resultModal.textContent || "";
        check(resultText.includes("已插入"), "结果面板显示已插入");
        const newNotes = app.vault
          .getFiles()
          .filter(
            (file) =>
              file.extension === "md" &&
              file.path.startsWith(`__daoru自测__/${tinyBase}`),
          );
        check(newNotes.length === 0, "插入模式没有新建同名笔记", String(newNotes.length));
        closeResult();
        await sleep(300);

        const content = editor.getValue();
        check(content.startsWith("插入前"), "插入前内容保留");
        check(content.includes("# 孕前及孕期保健指南"), "插入内容包含标题");
        check(
          content.includes("<table>") && content.includes("<ol>"),
          "插入内容包含 HTML 表格与嵌套列表",
        );
        check(!content.startsWith("---"), "插入模式不写 frontmatter");
        return { checks, contentLength: content.length, resultText: resultText.slice(0, 600) };
      },
      tiny.path,
      tiny.name,
      tinyBase,
    );
    addChecks(step4.checks);

    ctx.note({ step: "batch partial failure" });
    const step5 = await ctx.page(
      async (validPath, validName, invalidDocxPath, oldDocPath, tinyBase) => {
        const checks = [];
        const check = (ok, name, detail = "") =>
          checks.push({ 检查: name, 结论: ok ? "PASS" : "FAIL", 详情: detail });
        const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
        const waitFor = async (selector, timeout) => {
          const started = Date.now();
          while (Date.now() - started < timeout) {
            const element = document.querySelector(selector);
            if (element) return element;
            await sleep(100);
          }
          throw new Error(`等待超时：${selector}`);
        };
        const closeResult = () => {
          const done = Array.from(
            document.querySelectorAll(".sun-daoru-result-modal button"),
          ).find((button) => button.textContent?.trim() === "完成");
          done?.click();
        };

        closeResult();
        await sleep(200);
        const app = window.app;
        window.__daoruLive.files = [
          { path: validPath, name: validName },
          { path: invalidDocxPath, name: "坏文档.docx" },
          { path: oldDocPath, name: "旧文档.doc" },
        ];
        app.commands.executeCommandById("sun-daoru:import-word");
        const modal = await waitFor(".sun-daoru-modal", 30000);
        const modalText = modal.textContent || "";
        check(
          modalText.includes("坏文档.docx") && modalText.includes("旧文档.doc"),
          "批量面板显示全部文件",
        );
        const select = modal.querySelector("select");
        const hasInsert = select
          ? Array.from(select.options).some((option) => option.value === "insert")
          : false;
        check(!hasInsert, "批量导入不提供插入选项");
        const start = Array.from(modal.querySelectorAll("button")).find((button) =>
          button.textContent?.includes("开始导入"),
        );
        start?.click();
        const resultModal = await waitFor(".sun-daoru-result-modal", 120000);
        const resultText = resultModal.textContent || "";
        check(resultText.includes("成功 1 个，失败 2 个"), "批量结果 1 成功 2 失败");
        check(resultText.includes("旧版 .doc 不支持"), "报告旧版 .doc 错误");
        check(
          resultText.includes("不是有效的 .docx"),
          "报告损坏 .docx 错误",
        );
        const notePath = `__daoru自测__/${tinyBase}.md`;
        check(
          Boolean(app.vault.getAbstractFileByPath(notePath)),
          "批次中的有效文件仍生成笔记",
          notePath,
        );
        closeResult();
        await sleep(200);
        return { checks, resultText: resultText.slice(0, 1200) };
      },
      tiny.path,
      tiny.name,
      invalidDocx,
      oldDoc,
      tinyBase,
    );
    addChecks(step5.checks);

    checkAll();
  } finally {
    ctx.note({ step: "cleanup" });
    try {
      const cleanup = await ctx.page(async () => {
        const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
        const app = window.app;
        const plugin = app.plugins.plugins["sun-daoru"];

        for (const close of Array.from(
          document.querySelectorAll(
            ".modal-container .modal-close-button, .sun-daoru-modal .modal-header-button",
          ),
        )) {
          if (close instanceof HTMLElement) close.click();
        }
        await sleep(200);

        if (window.__daoruLive?.origClick) {
          HTMLInputElement.prototype.click = window.__daoruLive.origClick;
        }
        delete window.__daoruLive;

        plugin.settings.attachmentMode = "obsidian";
        plugin.settings.customAttachmentFolder = "附件/导入";
        plugin.settings.groupAttachmentsByNote = false;
        plugin.settings.normalizeHeadingLevels = false;
        plugin.settings.openAfterImport = true;
        plugin.settings.defaultTarget = "new";
        plugin.settings.addSourceFrontmatter = false;
        await plugin.saveSettings();

        const folder = app.vault.getAbstractFileByPath("__daoru自测__");
        if (folder) await app.vault.delete(folder, true);
        const welcome = app.vault.getAbstractFileByPath("欢迎.md");
        if (welcome) await app.workspace.getLeaf(false).openFile(welcome);
        await sleep(300);

        return {
          files: app.vault.getFiles().map((file) => file.path),
          settings: plugin.settings,
          modals: document.querySelectorAll(".modal-container").length,
        };
      });
      checks.push({
        检查: "清理测试数据并恢复默认设置",
        结论:
          cleanup.files.every((file) => !file.startsWith("__daoru自测__")) &&
          cleanup.modals === 0
            ? "PASS"
            : "FAIL",
        详情: cleanup,
      });
    } catch (error) {
      checks.push({
        检查: "清理测试数据并恢复默认设置",
        结论: "FAIL",
        详情: String(error?.message || error),
      });
    }
  }

  const failed = checks.filter((item) => item.结论 === "FAIL");
  return {
    checks,
    汇总:
      failed.length === 0
        ? `全部通过（${checks.length} 项）`
        : `失败 ${failed.length} 项`,
    screenshots,
  };
}
