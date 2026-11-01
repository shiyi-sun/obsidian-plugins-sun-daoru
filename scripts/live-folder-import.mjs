import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const small = {
  path: path.join(root, "妊娠期糖尿病营养干预方案.docx"),
  name: "妊娠期糖尿病营养干预方案.docx",
};
const tiny = {
  path: path.join(
    root,
    "3我热天热顺丰到付到付电风扇电风扇发射点手动阀手动阀.docx",
  ),
  name: "3我热天热顺丰到付到付电风扇电风扇发射点手动阀手动阀.docx",
};
const tinyBase = "3我热天热顺丰到付到付电风扇电风扇发射点手动阀手动阀";
const imageDoc = {
  path: path.join(root, "scripts", "fixtures", "one-image.docx"),
  name: "单图报告.docx",
};
if (!fs.existsSync(imageDoc.path)) {
  throw new Error(`缺少带图测试文档 ${imageDoc.path}`);
}

const invalidDocx = path.join(os.tmpdir(), "daoru-folder-invalid.docx");
const oldDoc = path.join(os.tmpdir(), "daoru-folder-old.doc");
const textFile = path.join(os.tmpdir(), "daoru-folder-note.txt");
const tempDocx = path.join(os.tmpdir(), "daoru-folder-temp.docx");
const batchDoc = path.join(os.tmpdir(), "daoru-folder-batch.docx");
fs.writeFileSync(invalidDocx, "not a zip");
fs.writeFileSync(oldDoc, "not a legacy word file");
fs.writeFileSync(textFile, "not a word file");
fs.writeFileSync(tempDocx, "word lock file");
fs.copyFileSync(small.path, batchDoc);

export default async function (ctx) {
  const report = await ctx.page(
    async (paths) => {
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
      const tinyBase = paths.tinyName.replace(/\.docx$/i, "");

      try {
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

        plugin.settings.attachmentMode = "obsidian";
        plugin.settings.customAttachmentFolder = "附件/导入";
        plugin.settings.groupAttachmentsByNote = true;
        plugin.settings.normalizeHeadingLevels = true;
        plugin.settings.openAfterImport = false;
        plugin.settings.defaultTarget = "new";
        plugin.settings.addSourceFrontmatter = true;
        await plugin.saveSettings();

        window.__daoruLive = {
          origClick: HTMLInputElement.prototype.click,
          files: [],
          folderFiles: [
            {
              path: paths.smallPath,
              name: paths.smallName,
              relativePath: `营养资料/一级/${paths.smallName}`,
            },
            {
              path: paths.tinyPath,
              name: paths.tinyName,
              relativePath: `营养资料/二级/${paths.tinyName}`,
            },
            {
              path: paths.invalidDocx,
              name: "坏文档.docx",
              relativePath: "营养资料/坏文档.docx",
            },
            {
              path: paths.oldDoc,
              name: "旧文档.doc",
              relativePath: "营养资料/二级/旧文档.doc",
            },
            {
              path: paths.textFile,
              name: "说明.txt",
              relativePath: "营养资料/说明.txt",
            },
            {
              path: paths.tempDocx,
              name: "~$临时.docx",
              relativePath: "营养资料/~$临时.docx",
            },
            {
              path: paths.imagePath,
              name: paths.imageName,
              relativePath: `营养资料/图片/${paths.imageName}`,
            },
            {
              path: paths.imagePath,
              name: paths.imageName,
              relativePath: `营养资料/备份/${paths.imageName}`,
            },
            {
              path: paths.batchDoc,
              name: "批量一.docx",
              relativePath: "营养资料/批量/批量一.docx",
            },
            {
              path: paths.batchDoc,
              name: "批量二.docx",
              relativePath: "营养资料/批量/批量二.docx",
            },
            {
              path: paths.batchDoc,
              name: "批量三.docx",
              relativePath: "营养资料/批量/嵌套/批量三.docx",
            },
          ],
        };

        HTMLInputElement.prototype.click = function () {
          if (this.type !== "file") {
            return window.__daoruLive.origClick.call(this);
          }
          const isFolder = this.hasAttribute("webkitdirectory");
          const source = isFolder
            ? window.__daoruLive.folderFiles
            : window.__daoruLive.files;
          const fs = require("fs");
          const transfer = new DataTransfer();
          for (const item of source) {
            const buffer = fs.readFileSync(item.path);
            const file = new File([buffer], item.name, {
              type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            });
            if (item.relativePath) {
              Object.defineProperty(file, "webkitRelativePath", {
                value: item.relativePath,
                configurable: true,
              });
            }
            transfer.items.add(file);
          }
          this.files = transfer.files;
          this.dispatchEvent(new Event("change", { bubbles: true }));
        };

        app.commands.executeCommandById("sun-daoru:import-word-folder");
        const modal = await waitFor(".sun-daoru-modal", 30000);
        const modalText = modal.textContent || "";
        check(modalText.includes("导入 Word 文件夹"), "显示文件夹导入标题");
        check(
          modalText.includes("已选择 9 个 Word 文件"),
          "只统计 Word 文件，跳过 txt 和 ~$ 临时文件",
          modalText.slice(0, 120),
        );
        check(
          modalText.includes("营养资料/一级/妊娠期糖尿病营养干预方案.docx"),
          "面板显示源相对路径",
        );
        check(
          !modalText.includes("说明.txt") && !modalText.includes("~$临时.docx"),
          "非 Word 和临时文件未进入导入列表",
        );
        check(modalText.includes("会保留源文件夹"), "提示会保留源文件夹层级");
        check(modalText.includes("__daoru自测__"), "默认目标为当前笔记所在文件夹");
        const select = modal.querySelector("select");
        const hasInsert = select
          ? Array.from(select.options).some((option) => option.value === "insert")
          : false;
        check(!hasInsert, "文件夹导入不提供插入当前笔记");
        check(
          modalText.includes("附件位置：笔记同目录的 <文档名>-附件/"),
          "面板显示按文档分组的附件位置",
        );

        const start = Array.from(modal.querySelectorAll("button")).find((button) =>
          button.textContent?.includes("开始导入"),
        );
        check(Boolean(start), "找到开始导入按钮");
        start?.click();

        const resultModal = await waitFor(".sun-daoru-result-modal", 120000);
        await sleep(300);
        const resultText = resultModal.textContent || "";
        check(resultText.includes("成功 7 个，失败 2 个"), "结果 7 成功 2 失败", resultText.slice(0, 80));
        check(resultText.includes("旧版 .doc 不支持"), "报告旧版 .doc 错误");
        check(resultText.includes("不是有效的 .docx"), "报告坏 docx 错误");
        check(
          resultText.includes("营养资料/一级/妊娠期糖尿病营养干预方案.docx"),
          "结果面板保留源相对路径",
        );

        const firstNotePath =
          "__daoru自测__/营养资料/一级/妊娠期糖尿病营养干预方案.md";
        const secondNotePath = `__daoru自测__/营养资料/二级/${tinyBase}.md`;
        const firstNote = app.vault.getAbstractFileByPath(firstNotePath);
        const secondNote = app.vault.getAbstractFileByPath(secondNotePath);
        check(Boolean(firstNote), "一级目录生成笔记", firstNotePath);
        check(Boolean(secondNote), "二级目录生成笔记", secondNotePath);

        let firstContent = "";
        if (firstNote) firstContent = await app.vault.read(firstNote);
        check(
          firstContent.startsWith(
            '---\nsource: "营养资料/一级/妊娠期糖尿病营养干预方案.docx"\nimported: ',
          ),
          "frontmatter source 使用源相对路径",
          firstContent.slice(0, 120),
        );
        check(
          firstContent.includes(
            "# 妊娠期糖尿病营养干预方案\n\n## 一、体重监控",
          ),
          "标题层级整理生效：文件标题 H1，章节 H2",
        );
        check(
          (firstContent.match(/^# /gm) || []).length === 1,
          "整理后只保留一个 H1",
          String((firstContent.match(/^# /gm) || []).length),
        );

        let secondContent = "";
        if (secondNote) secondContent = await app.vault.read(secondNote);
        check(
          secondContent.includes("# 孕前及孕期保健指南") &&
            secondContent.includes("## 一、健康教育及指导"),
          "第二个文档也按文件标题下移一级",
        );

        /* —— 文件夹导入里带文档的硬骨头：附件必须落在笔记自己的子目录里 —— */
        const imageNotePath = "__daoru自测__/营养资料/图片/单图报告.md";
        const backupNotePath = "__daoru自测__/营养资料/备份/单图报告.md";
        const imageNote = app.vault.getAbstractFileByPath(imageNotePath);
        const backupNote = app.vault.getAbstractFileByPath(backupNotePath);
        check(Boolean(imageNote), "图片文档在子目录里生成笔记", imageNotePath);
        check(Boolean(backupNote), "同名文档在别的子目录不被覆盖");

        const attached = app.vault
          .getFiles()
          .filter((file) => file.path.includes("单图报告-附件/"));
        check(
          attached.length === 2,
          "两份同名笔记各有自己的附件目录",
          attached.map((f) => f.path).join(" | "),
        );
        check(
          attached.some((file) =>
            file.path.startsWith("__daoru自测__/营养资料/图片/单图报告-附件/"),
          ) &&
            attached.some((file) =>
              file.path.startsWith("__daoru自测__/营养资料/备份/单图报告-附件/"),
            ),
          "附件目录建在各自笔记旁边",
          attached.map((f) => f.path).join(" | "),
        );

        let imageContent = "";
        if (imageNote) imageContent = await app.vault.read(imageNote);
        check(
          /!\[[^\]]*\]\(单图报告-附件\/001\.png\)/.test(imageContent),
          "笔记里用的是指向同目录附件的相对路径",
          (imageContent.match(/!\[[^\]]*\]\([^)]*\)/) || ["无"])[0],
        );

        // Obsidian 能不能真的解析这个相对路径（等价于图片显示得出来）
        let resolved = null;
        for (let i = 0; i < 40 && !resolved; i++) {
          const links = app.metadataCache.resolvedLinks[imageNotePath];
          if (links && Object.keys(links).length > 0) resolved = links;
          else await sleep(150);
        }
        check(
          Boolean(
            resolved &&
              Object.keys(resolved).some((target) =>
                target.startsWith("__daoru自测__/营养资料/图片/单图报告-附件/"),
              ),
          ),
          "相对路径被 Obsidian 解析到实际附件",
          resolved ? Object.keys(resolved).join(" | ") : "超时未解析",
        );

        /* —— 多文件与更深层级 —— */
        const batchNotes = app.vault
          .getFiles()
          .filter((file) => file.path.startsWith("__daoru自测__/营养资料/批量/"))
          .map((file) => file.path);
        check(
          batchNotes.length === 3 &&
            batchNotes.some((p) => p.includes("批量/嵌套/批量三.md")),
          "一次导入 3 份且保留更深层级",
          batchNotes.join(" | "),
        );

        /* —— 取消文件夹选择 —— */
        const filesBeforeCancel = app.vault.getFiles().length;
        window.__daoruLive.folderFiles = [];
        app.commands.executeCommandById("sun-daoru:import-word-folder");
        await sleep(700);
        check(
          !document.querySelector(".sun-daoru-modal") &&
            app.vault.getFiles().length === filesBeforeCancel &&
            plugin.importing === false,
          "取消选择：不弹面板、不写文件、导入锁释放",
          `面板 ${Boolean(document.querySelector(".sun-daoru-modal"))} · 文件 ${app.vault.getFiles().length}/${filesBeforeCancel} · importing ${plugin.importing}`,
        );

        closeResult();
        await sleep(300);

        return {
          checks,
          notePaths: [firstNotePath, secondNotePath],
          firstSnippet: firstContent.slice(0, 300),
          resultText: resultText.slice(0, 1200),
        };
      } finally {
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
        plugin.settings.groupAttachmentsByNote = true;
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
      }
    },
    {
      smallPath: small.path,
      smallName: small.name,
      tinyPath: tiny.path,
      tinyName: tiny.name,
      imagePath: imageDoc.path,
      imageName: imageDoc.name,
      batchDoc,
      invalidDocx,
      oldDoc,
      textFile,
      tempDocx,
    },
  );

  const failed = report.checks.filter((item) => item.结论 === "FAIL");
  return {
    checks: report.checks,
    汇总:
      failed.length === 0
        ? `全部通过（${report.checks.length} 项）`
        : `失败 ${failed.length} 项`,
    notePaths: report.notePaths,
    firstSnippet: report.firstSnippet,
  };
}
