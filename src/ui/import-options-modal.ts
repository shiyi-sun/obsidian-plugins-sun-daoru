import {
  App,
  FuzzySuggestModal,
  Modal,
  Setting,
  TFolder,
} from "obsidian";
import type { ImportTarget } from "../settings";

export interface ImportDialogResult {
  target: ImportTarget;
  folder: string;
  addSourceFrontmatter: boolean;
}

interface ImportDialogConfig {
  canInsert: boolean;
  defaultTarget: ImportTarget;
  defaultFolder: string;
  addSourceFrontmatter: boolean;
  attachmentDescription: string;
  insertTargetPath: string;
}

export class ImportOptionsModal extends Modal {
  private readonly files: File[];
  private readonly config: ImportDialogConfig;
  private readonly canInsert: boolean;
  private target: ImportTarget;
  private folder: string;
  private addSourceFrontmatter: boolean;
  private resolveResult: ((value: ImportDialogResult | null) => void) | null = null;
  private settled = false;
  private folderSetting: Setting | null = null;
  private sourceSetting: Setting | null = null;

  constructor(app: App, files: File[], config: ImportDialogConfig) {
    super(app);
    this.files = files;
    this.config = config;
    this.canInsert = config.canInsert && files.length === 1;
    this.target =
      config.defaultTarget === "insert" && this.canInsert ? "insert" : "new";
    this.folder = config.defaultFolder;
    this.addSourceFrontmatter = config.addSourceFrontmatter;
  }

  openAndGet(): Promise<ImportDialogResult | null> {
    return new Promise((resolve) => {
      this.resolveResult = resolve;
      this.open();
    });
  }

  onOpen(): void {
    const { contentEl } = this;
    this.modalEl.addClass("sun-daoru-modal");
    contentEl.empty();

    contentEl.createEl("h2", { text: "导入 Word" });
    const totalBytes = this.files.reduce((sum, file) => sum + file.size, 0);
    contentEl.createEl("p", {
      cls: "sun-daoru-modal-summary",
      text:
        this.files.length === 1
          ? `已选择 1 个文件（${formatBytes(totalBytes)}）`
          : `已选择 ${this.files.length} 个文件（共 ${formatBytes(totalBytes)}）`,
    });

    const fileList = contentEl.createDiv({ cls: "sun-daoru-file-list" });
    for (const file of this.files.slice(0, 6)) {
      const row = fileList.createDiv({ cls: "sun-daoru-file-row" });
      row.createSpan({ cls: "sun-daoru-file-name", text: file.name });
      row.createSpan({ cls: "sun-daoru-file-size", text: formatBytes(file.size) });
    }
    if (this.files.length > 6) {
      fileList.createEl("p", {
        cls: "sun-daoru-muted",
        text: `还有 ${this.files.length - 6} 个文件未显示。`,
      });
    }

    const oldDocCount = this.files.filter((file) => /\.doc$/i.test(file.name)).length;
    if (oldDocCount > 0) {
      contentEl.createEl("p", {
        cls: "sun-daoru-warning-text",
        text: `其中 ${oldDocCount} 个是旧版 .doc，无法转换；请在 Word 或 WPS 中另存为 .docx。`,
      });
    }

    new Setting(contentEl)
      .setName("导入方式")
      .setDesc(
        this.canInsert
          ? `插入到当前笔记：${this.config.insertTargetPath || "当前笔记"}。批量导入只能新建笔记。`
          : "当前没有可插入的 Markdown 笔记，或一次选择了多个文件，因此只能新建笔记。",
      )
      .addDropdown((dropdown) => {
        dropdown.addOption("new", "新建笔记");
        if (this.canInsert) dropdown.addOption("insert", "插入当前笔记");
        dropdown.setValue(this.target).onChange((value) => {
          this.target = value as ImportTarget;
          this.syncDisabledState();
        });
      });

    this.folderSetting = new Setting(contentEl)
      .setName("目标文件夹")
      .setDesc(this.folder || "库根目录")
      .addButton((button) =>
        button.setButtonText("选择").onClick(() => {
          new FolderSuggestModal(this.app, (folder) => {
            this.folder = folder;
            this.folderSetting?.setDesc(this.folder || "库根目录");
          }).open();
        }),
      );

    this.sourceSetting = new Setting(contentEl)
      .setName("写入来源信息")
      .setDesc("在笔记开头写入 source 和 imported。只对新建笔记生效。")
      .addToggle((toggle) =>
        toggle.setValue(this.addSourceFrontmatter).onChange((value) => {
          this.addSourceFrontmatter = value;
        }),
      );

    contentEl.createEl("p", {
      cls: "sun-daoru-modal-note",
      text: this.config.attachmentDescription,
    });
    contentEl.createEl("p", {
      cls: "sun-daoru-modal-note",
      text: "字体、字号、颜色、对齐、页眉页脚等 Word 排版不会导入；合并单元格或含嵌套内容的表格会保留为 HTML。",
    });

    const buttons = contentEl.createDiv({ cls: "sun-daoru-modal-buttons" });
    const cancel = buttons.createEl("button", { text: "取消" });
    cancel.addEventListener("click", () => this.finish(null));
    const start = buttons.createEl("button", {
      text: "开始导入",
      cls: "mod-cta",
    });
    start.addEventListener("click", () =>
      this.finish({
        target: this.target,
        folder: this.folder,
        addSourceFrontmatter: this.addSourceFrontmatter,
      }),
    );

    this.syncDisabledState();
  }

  private syncDisabledState(): void {
    const disabled = this.target === "insert";
    this.folderSetting?.setDisabled(disabled);
    this.sourceSetting?.setDisabled(disabled);
  }

  private finish(result: ImportDialogResult | null): void {
    if (this.settled) return;
    this.settled = true;
    this.resolveResult?.(result);
    this.close();
  }

  onClose(): void {
    this.contentEl.empty();
    if (!this.settled) {
      this.settled = true;
      this.resolveResult?.(null);
    }
  }
}

class FolderSuggestModal extends FuzzySuggestModal<TFolder> {
  constructor(
    app: App,
    private readonly onChoose: (folder: string) => void,
  ) {
    super(app);
    this.setPlaceholder("选择目标文件夹");
  }

  getItems(): TFolder[] {
    return this.app.vault
      .getAllLoadedFiles()
      .filter(
        (file): file is TFolder =>
          file instanceof TFolder &&
          (file.path === "/" ||
            !file.path.split("/").some((segment) => segment.startsWith("."))),
      )
      .sort((a, b) => a.path.localeCompare(b.path, "zh"));
  }

  getItemText(folder: TFolder): string {
    return folder.path === "/" ? "库根目录" : folder.path;
  }

  onChooseItem(folder: TFolder): void {
    this.onChoose(folder.path === "/" ? "" : folder.path);
  }
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}
