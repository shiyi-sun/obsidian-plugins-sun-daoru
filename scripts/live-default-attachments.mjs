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
      const testNote = app.vault.getAbstractFileByPath("__daoru自测__/导入测试.md");
      await app.workspace.getLeaf(false).openFile(testNote);
      await sleep(300);

      plugin.settings.attachmentMode = "obsidian";
      plugin.settings.openAfterImport = false;
      plugin.settings.defaultTarget = "new";
      plugin.settings.addSourceFrontmatter = false;
      await plugin.saveSettings();

      const beforeFiles = new Set(app.vault.getFiles().map((file) => file.path));
      const beforeFolders = new Set(
        app.vault
          .getAllLoadedFiles()
          .filter((file) => Array.isArray(file.children))
          .map((folder) => folder.path),
      );

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
      check(
        (modal.textContent || "").includes("跟随 Obsidian"),
        "面板显示跟随 Obsidian 附件设置",
      );
      const start = Array.from(modal.querySelectorAll("button")).find((button) =>
        button.textContent?.includes("开始导入"),
      );
      start?.click();
      const resultModal = await waitFor(".sun-daoru-result-modal", 120000);
      const resultText = resultModal.textContent || "";
      check(resultText.includes("成功"), "默认附件模式下导入成功");
      check(resultText.includes("149 张图片"), "默认附件模式下保存 149 张图片");

      const notePath = "__daoru自测__/功能营养素手册.md";
      const note = app.vault.getAbstractFileByPath(notePath);
      let content = "";
      if (note) content = await app.vault.read(note);
      check(Boolean(note), "生成笔记", notePath);

      const afterFiles = app.vault.getFiles();
      const newFiles = afterFiles.filter((file) => !beforeFiles.has(file.path));
      const imageFiles = newFiles.filter(
        (file) =>
          file.path !== notePath &&
          /功能营养素手册-\d+\.(png|jpg|jpeg|gif|webp|svg|bmp|tiff|emf|wmf)$/i.test(
            file.path,
          ),
      );
      check(imageFiles.length === 149, "默认附件模式写入 149 张图片", String(imageFiles.length));
      if (imageFiles[0]) {
        check(
          content.includes(`![[${imageFiles[0].path}|`),
          "笔记图片链接指向实际附件路径",
          imageFiles[0].path,
        );
      }
      const imageFolders = Array.from(
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

      for (const file of newFiles) {
        try {
          await app.vault.delete(file, true);
        } catch {
          // 清理尽力而为
        }
      }
      const afterFolders = app.vault
        .getAllLoadedFiles()
        .filter((file) => Array.isArray(file.children))
        .map((folder) => folder.path)
        .filter((folderPath) => !beforeFolders.has(folderPath))
        .sort((a, b) => b.length - a.length);
      for (const folderPath of afterFolders) {
        const folder = app.vault.getAbstractFileByPath(folderPath);
        if (folder && Array.isArray(folder.children) && folder.children.length === 0) {
          try {
            await app.vault.delete(folder, true);
          } catch {
            // 清理尽力而为
          }
        }
      }

      if (window.__daoruLive?.origClick) {
        HTMLInputElement.prototype.click = window.__daoruLive.origClick;
      }
      delete window.__daoruLive;

      plugin.settings.attachmentMode = "obsidian";
      plugin.settings.customAttachmentFolder = "附件/Word 导入";
      plugin.settings.openAfterImport = true;
      plugin.settings.defaultTarget = "new";
      plugin.settings.addSourceFrontmatter = false;
      await plugin.saveSettings();

      const sandbox = app.vault.getAbstractFileByPath("__daoru自测__");
      if (sandbox) await app.vault.delete(sandbox, true);
      const welcome = app.vault.getAbstractFileByPath("欢迎.md");
      if (welcome) await app.workspace.getLeaf(false).openFile(welcome);
      await sleep(300);

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
    汇总: failed.length === 0 ? `全部通过（${report.checks.length} 项）` : `失败 ${failed.length} 项`,
    imageFolders: report.imageFolders,
    newFiles: report.newFiles,
    remainingFiles: report.remainingFiles,
  };
}
