import { App, Modal, Notice, TFile } from "obsidian";
import type { ImportFileResult } from "../importer";
import type { LossItem } from "../core/losses";

const KIND_LABEL: Record<LossItem["kind"], string> = {
  lost: "丢失",
  downgraded: "降级",
  info: "说明",
};

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
    const lostTotal = this.results.reduce(
      (sum, result) => sum + result.losses.filter((loss) => loss.kind === "lost").length,
      0,
    );

    contentEl.createEl("h2", { text: "导入结果" });
    contentEl.createEl("p", {
      cls: "sun-daoru-modal-summary",
      text:
        failed === 0
          ? `成功导入 ${success} 个文件。`
          : `成功 ${success} 个，失败 ${failed} 个。`,
    });
    if (lostTotal > 0) {
      contentEl.createEl("p", {
        cls: "sun-daoru-loss-total",
        text: `有 ${lostTotal} 类内容没能导入，详见下面的「丢失」。`,
      });
    }

    const list = contentEl.createDiv({ cls: "sun-daoru-result-list" });
    for (const result of this.results) {
      this.renderResult(list, result);
    }

    contentEl.createEl("p", {
      cls: "sun-daoru-modal-note",
      text: "文字、标题、列表、表格、链接、图片和图片替代文字都会保留；字体、字号、颜色、对齐、页边距等排版不导入。没能导入的内容已在上面逐项列出。",
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

    this.renderLosses(row, result.losses);

    for (const warning of result.warnings) {
      row.createEl("p", { cls: "sun-daoru-warning-text", text: `提示：${warning}` });
    }
  }

  private renderLosses(row: HTMLElement, losses: LossItem[]): void {
    if (losses.length === 0) {
      row.createEl("p", { cls: "sun-daoru-loss-none", text: "未发现内容丢失。" });
      return;
    }

    const order: LossItem["kind"][] = ["lost", "downgraded", "info"];
    for (const kind of order) {
      for (const loss of losses.filter((item) => item.kind === kind)) {
        const line = row.createDiv({ cls: `sun-daoru-loss sun-daoru-loss-${kind}` });
        line.createSpan({ cls: "sun-daoru-loss-tag", text: KIND_LABEL[kind] });
        line.createSpan({
          cls: "sun-daoru-loss-label",
          text: `${loss.label} ×${loss.count}`,
        });
        line.createSpan({ cls: "sun-daoru-loss-detail", text: loss.detail });
      }
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
