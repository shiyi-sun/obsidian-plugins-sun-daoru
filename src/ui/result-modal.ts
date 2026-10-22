import { App, Modal, Notice, TFile } from "obsidian";
import type { ImportFileResult } from "../importer";

export class ImportResultModal extends Modal {
  constructor(
    app: App,
    private readonly results: ImportFileResult[],
    private readonly openAfterImport: boolean,
  ) {
    super(app);
  }

  onOpen(): void {
    const { contentEl } = this;
    this.modalEl.addClass("sun-daoru-result-modal");
    contentEl.empty();

    const success = this.results.filter((result) => !result.error).length;
    const failed = this.results.length - success;

    contentEl.createEl("h2", { text: "导入结果" });
    contentEl.createEl("p", {
      cls: "sun-daoru-modal-summary",
      text:
        failed === 0
          ? `成功导入 ${success} 个文件。`
          : `成功 ${success} 个，失败 ${failed} 个。`,
    });

    const list = contentEl.createDiv({ cls: "sun-daoru-result-list" });
    for (const result of this.results) {
      this.renderResult(list, result);
    }

    contentEl.createEl("p", {
      cls: "sun-daoru-modal-note",
      text: "提示：Word 的字体、字号、颜色、对齐、页边距、页眉页脚不会导入；标题、列表、表格、链接和图片会尽量保留。",
    });

    const buttons = contentEl.createDiv({ cls: "sun-daoru-modal-buttons" });
    const done = buttons.createEl("button", { text: "完成", cls: "mod-cta" });
    done.addEventListener("click", () => this.close());

    if (this.openAfterImport) {
      const first = this.results.find(
        (result) => !result.error && result.notePath && !result.inserted,
      );
      if (first?.notePath) this.openNote(first.notePath, false);
    }
  }

  private renderResult(container: HTMLElement, result: ImportFileResult): void {
    const row = container.createDiv({ cls: "sun-daoru-result-row" });
    const header = row.createDiv({ cls: "sun-daoru-result-header" });
    header.createSpan({ cls: "sun-daoru-result-name", text: result.fileName });
    header.createSpan({
      cls: result.error ? "sun-daoru-status-error" : "sun-daoru-status-ok",
      text: result.error ? "失败" : result.inserted ? "已插入" : "成功",
    });

    if (result.error) {
      row.createEl("p", { cls: "sun-daoru-error-text", text: result.error });
      return;
    }

    const meta: string[] = [];
    if (result.images > 0) meta.push(`${result.images} 张图片`);
    if (result.simpleTables > 0) meta.push(`${result.simpleTables} 个简单表格`);
    if (result.complexTables > 0) meta.push(`${result.complexTables} 个复杂表格（HTML）`);
    row.createEl("p", {
      cls: "sun-daoru-result-meta",
      text: meta.length > 0 ? meta.join(" · ") : "纯文字内容",
    });

    if (result.inserted) {
      row.createEl("p", {
        cls: "sun-daoru-muted",
        text: `已插入：${result.notePath || "当前笔记"}`,
      });
    } else if (result.notePath) {
      const actions = row.createDiv({ cls: "sun-daoru-result-actions" });
      const open = actions.createEl("button", { text: "打开笔记" });
      open.addEventListener("click", () => {
        if (result.notePath) this.openNote(result.notePath, true);
      });
      actions.createSpan({ cls: "sun-daoru-muted", text: result.notePath });
    }

    if (result.warnings.length > 0) {
      const warnings = row.createEl("ul", { cls: "sun-daoru-warnings" });
      for (const warning of result.warnings) {
        warnings.createEl("li", { text: warning });
      }
    } else {
      row.createEl("p", { cls: "sun-daoru-muted", text: "未发现图片/表格等文件级问题。" });
    }
  }

  private openNote(path: string, closeModal: boolean): void {
    const file = this.app.vault.getAbstractFileByPath(path);
    if (!(file instanceof TFile)) {
      new Notice(`找不到笔记：${path}`);
      return;
    }
    void this.app.workspace.getLeaf(false).openFile(file);
    if (closeModal) this.close();
  }

  onClose(): void {
    this.contentEl.empty();
  }
}
