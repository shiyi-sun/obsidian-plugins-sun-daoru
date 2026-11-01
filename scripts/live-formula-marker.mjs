/*
公式原位标记的实机回归。

用真实小文档在临时目录生成一份带公式的 .docx，导入测试库的 __daoru自测__，
检查笔记里的【公式未导入】位置、结果面板文案和清理情况。

需要 Obsidian 带 --remote-debugging-port=9223 启动：
  node ../sun-home/scripts/cdp/run.mjs scripts/live-formula-marker.mjs
*/
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);
const JSZip = require("jszip");
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const fixturePath = path.join(os.tmpdir(), "daoru-formula-live.docx");

const T = (text) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const F = (text) => `<m:oMath><m:r><m:t>${text}</m:t></m:r></m:oMath>`;
const P = (inner) => `<w:p>${inner}</w:p>`;

async function buildFixture() {
  const zip = await JSZip.loadAsync(
    fs.readFileSync(path.join(root, "妊娠期糖尿病营养干预方案.docx")),
  );
  let xml = await zip.file("word/document.xml").async("string");
  const boundary = xml.indexOf("<w:body>") + "<w:body>".length;
  const injected =
    P(T("公式前文") + F("LIVEINLINE") + T("公式后文")) +
    P(`<m:oMathPara><m:oMath><m:r><m:t>LIVEDISPLAY</m:t></m:r></m:oMath></m:oMathPara>`) +
    `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr/>${P(
      T("表格前") + F("LIVECELL") + T("表格后"),
    )}</w:tc></w:tr></w:tbl>`;
  xml = xml.slice(0, boundary) + injected + xml.slice(boundary);
  zip.file("word/document.xml", xml);
  fs.writeFileSync(
    fixturePath,
    await zip.generateAsync({ type: "nodebuffer", compression: "STORE" }),
  );
}

export default async function (ctx) {
  await buildFixture();
  try {
    const result = await ctx.page(async (filePath) => {
      const checks = [];
      const check = (ok, name, detail = "") =>
        checks.push({
          检查: name,
          结论: ok ? "PASS" : "FAIL",
          详情: String(detail).slice(0, 400),
        });
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
      const closeModals = async () => {
        for (let i = 0; i < 10; i++) {
          const done = Array.from(
            document.querySelectorAll(".sun-daoru-result-modal button"),
          ).find((button) => button.textContent?.trim() === "完成");
          if (!done) break;
          done.click();
          await sleep(200);
        }
      };

      const app = window.app;
      const plugin = app.plugins.plugins["sun-daoru"];
      await closeModals();
      await waitFor(
        async () => !plugin.importing,
        120000,
        "上一批导入结束",
      );

      const oldSandbox = app.vault.getAbstractFileByPath("__daoru自测__");
      if (oldSandbox) await app.vault.delete(oldSandbox, true);
      await app.vault.createFolder("__daoru自测__");
      await app.vault.create("__daoru自测__/导入测试.md", "测试用笔记。\n");
      const sandboxNote = app.vault.getAbstractFileByPath("__daoru自测__/导入测试.md");
      await app.workspace.getLeaf(false).openFile(sandboxNote);
      await sleep(300);

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
        filePath,
      };
      HTMLInputElement.prototype.click = function () {
        if (this.type !== "file") return window.__daoruLive.origClick.call(this);
        const buffer = require("fs").readFileSync(window.__daoruLive.filePath);
        const transfer = new DataTransfer();
        transfer.items.add(
          new File([buffer], "daoru-formula-live.docx", {
            type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          }),
        );
        this.files = transfer.files;
        this.dispatchEvent(new Event("change", { bubbles: true }));
      };

      let resultText = "";
      let content = "";
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
        if (!start) throw new Error("没有开始导入按钮");
        const before = document.querySelectorAll(".sun-daoru-result-modal").length;
        start.click();
        check(
          await waitFor(async () => plugin.importing, 15000, "导入任务启动"),
          "导入任务已启动",
        );
        check(
          await waitFor(async () => !plugin.importing, 120000, "导入任务结束"),
          "导入任务结束（未卡住）",
        );
        const resultModal = await waitFor(() => {
          const modals = document.querySelectorAll(".sun-daoru-result-modal");
          return modals.length > before ? modals[modals.length - 1] : null;
        }, 30000, "结果弹窗");
        resultText = resultModal.textContent || "";

        const notePath = "__daoru自测__/daoru-formula-live.md";
        const note = app.vault.getAbstractFileByPath(notePath);
        check(Boolean(note), "生成公式测试笔记", notePath);
        if (note) content = await app.vault.read(note);
      } finally {
        if (window.__daoruLive?.origClick) {
          HTMLInputElement.prototype.click = window.__daoruLive.origClick;
        }
        delete window.__daoruLive;
        await closeModals();
        plugin.settings.attachmentMode = "obsidian";
        plugin.settings.customAttachmentFolder = "附件/导入";
        plugin.settings.groupAttachmentsByNote = true;
        plugin.settings.normalizeHeadingLevels = false;
        plugin.settings.openAfterImport = true;
        plugin.settings.defaultTarget = "new";
        plugin.settings.addSourceFrontmatter = false;
        await plugin.saveSettings();
        const sandbox = app.vault.getAbstractFileByPath("__daoru自测__");
        if (sandbox) await app.vault.delete(sandbox, true);
        const welcome = app.vault.getAbstractFileByPath("欢迎.md");
        if (welcome) await app.workspace.getLeaf(false).openFile(welcome);
        await sleep(300);
      }

      check(resultText.includes("成功导入 1 个文件"), "公式文档导入成功", resultText);
      check(/丢失\s*公式 ×3/.test(resultText), "结果面板统计 3 个公式", resultText);
      check(
        resultText.includes("已在笔记中对应位置标出 3 处"),
        "结果面板说明已标出 3 处",
        resultText,
      );
      check(
        resultText.includes("搜索「公式未导入」可定位"),
        "结果面板给出搜索定位提示",
        resultText,
      );
      check(
        content.includes("公式前文【公式未导入】公式后文"),
        "句中公式标记在原位置",
        (content.match(/.*公式前文.*/) || [""])[0],
      );
      check(
        /\n\n【公式未导入】\n\n/.test(content),
        "独立成行的公式单独标记",
        content.slice(0, 200),
      );
      check(
        content.includes("表格前【公式未导入】表格后"),
        "表格单元格里的公式标记在原位置",
        (content.match(/.*表格前.*/) || [""])[0],
      );
      check(
        (content.match(/【公式未导入】/g) || []).length === 3,
        "笔记里恰好 3 个公式标记",
        String((content.match(/【公式未导入】/g) || []).length),
      );
      check(
        !content.includes("SUNDAORUFORMULAMARK"),
        "内部占位 token 不残留",
        "",
      );

      return {
        checks,
        resultText: resultText.slice(0, 900),
        contentSnippet: content.slice(0, 500),
        remaining: app.vault.getFiles().map((file) => file.path),
      };
    }, fixturePath);

    const failed = result.checks.filter((item) => item.结论 === "FAIL");
    return {
      checks: result.checks,
      汇总:
        failed.length === 0
          ? `全部通过（${result.checks.length} 项）`
          : `失败 ${failed.length} 项`,
      结果面板: result.resultText,
      笔记开头: result.contentSnippet,
      清理后库内文件: result.remaining,
    };
  } finally {
    fs.rmSync(fixturePath, { force: true });
  }
}
