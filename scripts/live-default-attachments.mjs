import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const big = {
  path: path.join(root, "功能营养素手册.docx"),
  name: "功能营养素手册.docx",
};

export default async function (ctx) {
  const report = await ctx.page(
    async (bigPath, bigName) => {
      const checks = [];
      const check = (ok, name, detail = "") =>
        checks.push({ 检查: name, 结论: ok ? "PASS" : "FAIL", 详情: detail });
      const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
      const app = window.app;
      const plugin = app.plugins.plugins["sun-daoru"];
      const imagePattern =
        /(?:^|\/)功能营养素手册-\d+\.(png|jpg|jpeg|gif|webp|svg|bmp|tiff|emf|wmf)$/i;

      const waitFor = async (selector, timeout) => {
        const started = Date.now();
        while (Date.now() - started < timeout) {
          const element = document.querySelector(selector);
          if (element) return element;
          await sleep(100);
        }
        throw new Error(`等待超时：${selector}`);
      };
      const waitForIdle = async (timeout = 180000) => {
        const started = Date.now();
        while (plugin?.importing) {
          if (Date.now() - started > timeout) {
            throw new Error("导入未在限时内结束");
          }
          await sleep(500);
        }
      };
      const closeAllModals = async () => {
        for (let i = 0; i < 10; i++) {
          const close = document.querySelector(
            ".modal-container .modal-close-button, .sun-daoru-modal .modal-header-button",
          );
          if (!close) break;
          if (close instanceof HTMLElement) close.click();
          await sleep(200);
        }
        await sleep(300);
      };
      const closeResult = () => {
        const done = Array.from(
          document.querySelectorAll(".sun-daoru-result-modal button"),
        ).find((button) => button.textContent?.trim() === "完成");
        done?.click();
      };

      const welcomeBefore = Boolean(app.vault.getAbstractFileByPath("欢迎.md"));
      let imageFolders = [];
      let newFiles = [];
      let resultText = "";
      let beforeFiles = new Set();
      let snapshotTaken = false;
      let idleForCleanup = true;

      try {
        await waitForIdle();
        await closeAllModals();
        await waitForIdle();

        const oldSandbox = app.vault.getAbstractFileByPath("__daoru自测__");
        if (oldSandbox) await app.vault.delete(oldSandbox, true);
        await app.vault.createFolder("__daoru自测__");
        await app.vault.create(
          "__daoru自测__/导入测试.md",
          "测试用笔记，导入完成后删除。\n",
        );
        const testNote = app.vault.getAbstractFileByPath(
          "__daoru自测__/导入测试.md",
        );
        await app.workspace.getLeaf(false).openFile(testNote);
        for (
          let i = 0;
          i < 20 && app.workspace.getActiveFile()?.path !== testNote.path;
          i++
        ) {
          await sleep(100);
        }
        check(
          app.workspace.getActiveFile()?.path === testNote.path,
          "导入前活动笔记正确",
          app.workspace.getActiveFile()?.path || "",
        );
        if (app.workspace.getActiveFile()?.path !== testNote.path) {
          throw new Error("测试笔记未激活，无法验证默认目标文件夹");
        }

        plugin.settings.attachmentMode = "obsidian";
        plugin.settings.groupAttachmentsByNote = false;
        plugin.settings.normalizeHeadingLevels = false;
        plugin.settings.openAfterImport = false;
        plugin.settings.defaultTarget = "new";
        plugin.settings.addSourceFrontmatter = false;
        await plugin.saveSettings();

        beforeFiles = new Set(app.vault.getFiles().map((file) => file.path));
        snapshotTaken = true;

        window.__daoruLive = {
          origClick: HTMLInputElement.prototype.click,
          files: [{ path: bigPath, name: bigName }],
        };
        HTMLInputElement.prototype.click = function () {
          if (this.type !== "file") {
            return window.__daoruLive.origClick.call(this);
          }
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
        check(
          (modal.textContent || "").includes("跟随 Obsidian"),
          "面板显示跟随 Obsidian 附件设置",
        );
        const start = Array.from(modal.querySelectorAll("button")).find((button) =>
          button.textContent?.includes("开始导入"),
        );
        check(Boolean(start), "找到开始导入按钮");
        start?.click();

        await waitForIdle(180000);
        await sleep(500);
        const resultModal = document.querySelector(".sun-daoru-result-modal");
        resultText = resultModal ? resultModal.textContent || "" : "";

        const notePath = "__daoru自测__/功能营养素手册.md";
        const note = app.vault.getAbstractFileByPath(notePath);
        let content = "";
        if (note) content = await app.vault.read(note);
        check(Boolean(note), "生成笔记", notePath);

        const afterFiles = app.vault.getFiles();
        newFiles = afterFiles.filter((file) => !beforeFiles.has(file.path));
        const imageFiles = afterFiles.filter(
          (file) => !beforeFiles.has(file.path) && imagePattern.test(file.path),
        );
        check(
          imageFiles.length === 149,
          "默认附件模式写入 149 张图片",
          String(imageFiles.length),
        );
        if (imageFiles[0]) {
          const name = imageFiles[0].path.split("/").pop();
          const pattern = new RegExp(
            `!\\[[^\\]]*\\]\\(<?[^)\\n]*${name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}>?\\)`,
          );
          check(
            pattern.test(content),
            "笔记图片链接指向实际附件文件",
            `${imageFiles[0].path} · ${(content.match(/!\[[^\]]*\]\([^)]*\)/) || ["无"])[0]}`,
          );
        }
        imageFolders = Array.from(
          new Set(
            imageFiles.map((file) =>
              file.path.includes("/")
                ? file.path.slice(0, file.path.lastIndexOf("/"))
                : "（库根目录）",
            ),
          ),
        );

        closeResult();
        await sleep(300);
      } finally {
        try {
          await waitForIdle(60000);
        } catch {
          idleForCleanup = false;
        }
        await closeAllModals();

        if (idleForCleanup) {
          const leftoverImages = snapshotTaken
            ? app.vault
                .getFiles()
                .filter(
                  (file) =>
                    !beforeFiles.has(file.path) && imagePattern.test(file.path),
                )
            : [];
          for (const file of leftoverImages) {
            try {
              await app.vault.delete(file, true);
            } catch {
              // 清理尽力而为
            }
          }

          const sandbox = app.vault.getAbstractFileByPath("__daoru自测__");
          if (sandbox) await app.vault.delete(sandbox, true);

          if (window.__daoruLive?.origClick) {
            HTMLInputElement.prototype.click = window.__daoruLive.origClick;
          }
          delete window.__daoruLive;

          plugin.settings.attachmentMode = "obsidian";
          plugin.settings.customAttachmentFolder = "附件/导入";
          plugin.settings.groupAttachmentsByNote = true;
          plugin.settings.normalizeHeadingLevels = false;
          plugin.settings.openAfterImport = true;
          plugin.settings.defaultTarget = "new";
          plugin.settings.addSourceFrontmatter = false;
          await plugin.saveSettings();

          const welcome = app.vault.getAbstractFileByPath("欢迎.md");
          if (welcome) await app.workspace.getLeaf(false).openFile(welcome);
          await sleep(300);
        }

        const welcomeAfter = Boolean(app.vault.getAbstractFileByPath("欢迎.md"));
        check(
          !welcomeBefore || welcomeAfter,
          "测试清理没有误删欢迎.md",
          welcomeBefore ? "导入前存在" : "导入前不存在",
        );
        check(idleForCleanup, "导入在清理前已结束");
      }

      return {
        checks,
        imageFolders,
        newFiles: newFiles.map((file) => file.path),
        remainingFiles: app.vault.getFiles().map((file) => file.path),
        resultText: resultText.slice(0, 600),
      };
    },
    big.path,
    big.name,
  );

  const failed = report.checks.filter((item) => item.结论 === "FAIL");
  return {
    checks: report.checks,
    汇总:
      failed.length === 0
        ? `全部通过（${report.checks.length} 项）`
        : `失败 ${failed.length} 项`,
    imageFolders: report.imageFolders,
    newFiles: report.newFiles,
    remainingFiles: report.remainingFiles,
  };
}
