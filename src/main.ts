import { MarkdownView, Notice, Plugin } from "obsidian";
import { pickDocxFiles, pickDocxFolder } from "./file-picker";
import { importFiles, importItems, ImportItem, ImportRequest } from "./importer";
import { parentFolderOf } from "./paths";
import {
  DEFAULT_SETTINGS,
  SunDaoruSettingTab,
  SunDaoruSettings,
} from "./settings";
import { ImportOptionsModal } from "./ui/import-options-modal";
import { ImportResultModal } from "./ui/result-modal";

type ImportMode = "files" | "folder";

const SKIP_FOLDER_SEGMENTS = new Set(["node_modules", "__MACOSX"]);

export default class SunDaoruPlugin extends Plugin {
  settings: SunDaoruSettings;
  private importing = false;

  async onload(): Promise<void> {
    await this.loadSettings();

    this.addCommand({
      id: "import-word",
      name: "导入 Word 文档…",
      callback: () => {
        void this.chooseAndImport("files");
      },
    });

    this.addCommand({
      id: "import-word-folder",
      name: "导入 Word 文件夹…",
      callback: () => {
        void this.chooseAndImport("folder");
      },
    });

    this.addRibbonIcon("import", "导入 Word 文档", () => {
      void this.chooseAndImport("files");
    });

    this.addRibbonIcon("folder-open", "导入 Word 文件夹", () => {
      void this.chooseAndImport("folder");
    });

    this.addSettingTab(new SunDaoruSettingTab(this.app, this));
  }

  async loadSettings(): Promise<void> {
    this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
  }

  async saveSettings(): Promise<void> {
    await this.saveData(this.settings);
  }

  private async chooseAndImport(mode: ImportMode): Promise<void> {
    if (this.importing) {
      new Notice("正在导入，请稍候。");
      return;
    }
    this.importing = true;

    try {
      const activeView = this.app.workspace.getActiveViewOfType(MarkdownView);
      const editor = activeView?.editor ?? null;
      const activeFile = activeView?.file ?? null;
      // 阅读视图下 editor.replaceSelection 的改动不落盘（实测），所以只在编辑视图允许插入
      const inEditMode = activeView ? activeView.getMode() === "source" : false;
      const canInsert = Boolean(editor && activeFile) && inEditMode;
      const insertUnavailable = !activeFile
        ? "当前没有打开的 Markdown 笔记，只能新建笔记。"
        : !inEditMode
          ? "当前笔记处于阅读视图，插入不可用；请切到编辑视图，或改用新建笔记。"
          : "";
      const defaultFolder = activeFile ? parentFolderOf(activeFile.path) : "";

      const picked = mode === "folder" ? await pickDocxFolder() : await pickDocxFiles();
      if (picked.length === 0) return;

      let files: File[];
      let items: ImportItem[];
      let sourceRootName = "";

      if (mode === "folder") {
        files = picked.filter(isFolderWordFile);
        if (files.length === 0) {
          new Notice("所选文件夹里没有可导入的 Word 文件。");
          return;
        }
        items = files.map((file) => ({
          file,
          relativePath: normalizeRelativePath(file.webkitRelativePath || file.name),
        }));
        sourceRootName =
          items[0]?.relativePath?.split("/").filter(Boolean)[0] || "所选文件夹";
      } else {
        files = picked;
        items = files.map((file) => ({ file }));
      }

      const displayNames =
        mode === "folder"
          ? items.map((item) => item.relativePath || item.file.name)
          : files.map((file) => file.name);

      const options = await new ImportOptionsModal(
        this.app,
        files,
        displayNames,
        {
          canInsert: mode === "files" ? canInsert : false,
          insertUnavailable:
            mode === "files" ? insertUnavailable : "文件夹导入只能新建笔记。",
          defaultTarget: mode === "folder" ? "new" : this.settings.defaultTarget,
          defaultFolder,
          addSourceFrontmatter: this.settings.addSourceFrontmatter,
          insertTargetPath: activeFile?.path || "",
          attachmentDescription: this.attachmentDescription(),
          folderMode: mode === "folder",
          sourceRootName,
          title: mode === "folder" ? "导入 Word 文件夹" : "导入 Word",
        },
      ).openAndGet();
      if (!options) return;

      const target = mode === "folder" ? "new" : options.target;
      if (mode === "files") this.settings.defaultTarget = options.target;
      this.settings.addSourceFrontmatter = options.addSourceFrontmatter;
      await this.saveSettings();

      const request: ImportRequest = {
        target,
        folder: options.folder,
        addSourceFrontmatter: options.addSourceFrontmatter,
        editor,
        activeFilePath: activeFile?.path || "",
      };

      const notice = new Notice(
        mode === "folder"
          ? `正在导入 ${files.length} 个 Word 文件…`
          : files.length === 1
            ? `正在导入 ${files[0].name}…`
            : `正在导入 ${files.length} 个文件…`,
        0,
      );
      try {
        const results =
          mode === "folder"
            ? await importItems(
                this.app,
                items,
                request,
                this.settings,
                (message) => notice.setMessage(message),
              )
            : await importFiles(
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

  private attachmentDescription(): string {
    if (this.settings.groupAttachmentsByNote) {
      return "附件位置：笔记同目录的 <文档名>-附件/（可在设置中关闭）";
    }
    return this.settings.attachmentMode === "custom"
      ? `附件位置：${this.settings.customAttachmentFolder}`
      : "附件位置：跟随 Obsidian 的「文件与链接」设置";
  }
}

function normalizeRelativePath(value: string): string {
  return value.replace(/\\/g, "/");
}

function isFolderWordFile(file: File): boolean {
  const relative = normalizeRelativePath(file.webkitRelativePath || file.name);
  const segments = relative.split("/").filter(Boolean);
  if (segments.some((segment) => segment.startsWith("~$") || segment.startsWith("."))) {
    return false;
  }
  if (segments.some((segment) => SKIP_FOLDER_SEGMENTS.has(segment))) {
    return false;
  }
  return /\.(docx|docm|doc)$/i.test(file.name);
}
