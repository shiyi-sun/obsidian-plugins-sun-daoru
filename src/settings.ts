import { App, PluginSettingTab, Setting } from "obsidian";
import type SunDaoruPlugin from "./main";

export type ImportTarget = "new" | "insert";
export type AttachmentMode = "obsidian" | "custom";

export interface SunDaoruSettings {
  defaultTarget: ImportTarget;
  attachmentMode: AttachmentMode;
  customAttachmentFolder: string;
  addSourceFrontmatter: boolean;
  openAfterImport: boolean;
}

export const DEFAULT_SETTINGS: SunDaoruSettings = {
  defaultTarget: "new",
  attachmentMode: "obsidian",
  customAttachmentFolder: "附件/Word 导入",
  addSourceFrontmatter: false,
  openAfterImport: true,
};

export class SunDaoruSettingTab extends PluginSettingTab {
  constructor(
    app: App,
    private plugin: SunDaoruPlugin,
  ) {
    super(app, plugin);
  }

  display(): void {
    const { containerEl } = this;
    containerEl.empty();

    new Setting(containerEl).setName("导入").setHeading();

    new Setting(containerEl)
      .setName("默认导入方式")
      .setDesc("新建笔记，或插入当前打开的 Markdown 笔记。批量导入时始终新建笔记。")
      .addDropdown((dropdown) =>
        dropdown
          .addOption("new", "新建笔记")
          .addOption("insert", "插入当前笔记")
          .setValue(this.plugin.settings.defaultTarget)
          .onChange(async (value) => {
            this.plugin.settings.defaultTarget = value as ImportTarget;
            await this.plugin.saveSettings();
          }),
      );

    new Setting(containerEl)
      .setName("附件位置")
      .setDesc("图片等附件放在哪里。跟随 Obsidian 设置时，会使用「文件与链接」中的附件保存位置。")
      .addDropdown((dropdown) =>
        dropdown
          .addOption("obsidian", "跟随 Obsidian 设置")
          .addOption("custom", "固定文件夹")
          .setValue(this.plugin.settings.attachmentMode)
          .onChange(async (value) => {
            this.plugin.settings.attachmentMode = value as AttachmentMode;
            await this.plugin.saveSettings();
            this.display();
          }),
      );

    if (this.plugin.settings.attachmentMode === "custom") {
      new Setting(containerEl)
        .setName("附件文件夹")
        .setDesc("相对于库根目录，例如「附件/Word 导入」。不存在的文件夹会自动创建。")
        .addText((text) =>
          text
            .setPlaceholder("附件/Word 导入")
            .setValue(this.plugin.settings.customAttachmentFolder)
            .onChange(async (value) => {
              this.plugin.settings.customAttachmentFolder =
                value.trim() || DEFAULT_SETTINGS.customAttachmentFolder;
              await this.plugin.saveSettings();
            }),
        );
    }

    new Setting(containerEl)
      .setName("新建笔记时写入来源信息")
      .setDesc("在笔记开头写入 source（原文件名）和 imported（导入日期）。")
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.addSourceFrontmatter)
          .onChange(async (value) => {
            this.plugin.settings.addSourceFrontmatter = value;
            await this.plugin.saveSettings();
          }),
      );

    new Setting(containerEl)
      .setName("导入后打开笔记")
      .setDesc("导入结束后自动打开第一篇成功生成的笔记。")
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.openAfterImport)
          .onChange(async (value) => {
            this.plugin.settings.openAfterImport = value;
            await this.plugin.saveSettings();
          }),
      );

    const note = containerEl.createDiv({ cls: "sun-daoru-settings-note" });
    note.createEl("p", {
      text: "仅支持 .docx / .docm。旧版 .doc 请先在 Word 或 WPS 中另存为 .docx。",
    });
    note.createEl("p", {
      text: "字体、字号、颜色、对齐、页边距、页眉页脚等排版信息不会导入；复杂表格会保留为 HTML。",
    });
  }
}
