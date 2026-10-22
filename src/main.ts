import { MarkdownView, Notice, Plugin } from "obsidian";
import { pickDocxFiles } from "./file-picker";
import { importFiles, ImportRequest } from "./importer";
import { parentFolderOf } from "./paths";
import {
  DEFAULT_SETTINGS,
  SunDaoruSettingTab,
  SunDaoruSettings,
} from "./settings";
import { ImportOptionsModal } from "./ui/import-options-modal";
import { ImportResultModal } from "./ui/result-modal";

export default class SunDaoruPlugin extends Plugin {
  settings: SunDaoruSettings;
  private importing = false;

  async onload(): Promise<void> {
    await this.loadSettings();

    this.addCommand({
      id: "import-word",
      name: "导入 Word 文档…",
      callback: () => {
        void this.chooseAndImport();
      },
    });

    this.addRibbonIcon("import", "导入 Word 文档", () => {
      void this.chooseAndImport();
    });

    this.addSettingTab(new SunDaoruSettingTab(this.app, this));
  }

  async loadSettings(): Promise<void> {
    this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
  }

  async saveSettings(): Promise<void> {
    await this.saveData(this.settings);
  }

  private async chooseAndImport(): Promise<void> {
    if (this.importing) {
      new Notice("正在导入，请稍候。");
      return;
    }
    this.importing = true;

    try {
      const activeView = this.app.workspace.getActiveViewOfType(MarkdownView);
      const editor = activeView?.editor ?? null;
      const activeFile = activeView?.file ?? null;
      const canInsert = Boolean(editor && activeFile);

      let defaultFolder = "";
      if (activeFile) {
        defaultFolder = parentFolderOf(activeFile.path);
      }

      const files = await pickDocxFiles();
      if (files.length === 0) return;

      const options = await new ImportOptionsModal(this.app, files, {
        canInsert,
        defaultTarget: this.settings.defaultTarget,
        defaultFolder,
        addSourceFrontmatter: this.settings.addSourceFrontmatter,
        insertTargetPath: activeFile?.path || "",
        attachmentDescription:
          this.settings.attachmentMode === "custom"
            ? `附件位置：${this.settings.customAttachmentFolder}`
            : "附件位置：跟随 Obsidian 的「文件与链接」设置",
      }).openAndGet();
      if (!options) return;

      this.settings.defaultTarget = options.target;
      this.settings.addSourceFrontmatter = options.addSourceFrontmatter;
      await this.saveSettings();

      const request: ImportRequest = {
        target: options.target,
        folder: options.folder,
        addSourceFrontmatter: options.addSourceFrontmatter,
        editor,
        activeFilePath: activeFile?.path || "",
      };

      const notice = new Notice(
        files.length === 1
          ? `正在导入 ${files[0].name}…`
          : `正在导入 ${files.length} 个文件…`,
        0,
      );
      try {
        const results = await importFiles(
          this.app,
          files,
          request,
          this.settings,
          (message) => notice.setMessage(message),
        );
        notice.hide();
        new ImportResultModal(
          this.app,
          results,
          this.settings.openAfterImport,
        ).open();
      } catch (error) {
        notice.hide();
        const message = error instanceof Error ? error.message : String(error);
        new Notice(`导入失败：${message}`, 10000);
      }
    } finally {
      this.importing = false;
    }
  }
}
