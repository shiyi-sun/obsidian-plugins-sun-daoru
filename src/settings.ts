import { App, PluginSettingTab, Setting } from "obsidian";
import type SunDaoruPlugin from "./main";

export type ImportTarget = "new" | "insert";
export type AttachmentMode = "obsidian" | "custom";

export interface SunDaoruSettings {
  defaultTarget: ImportTarget;
  attachmentMode: AttachmentMode;
  customAttachmentFolder: string;
  groupAttachmentsByNote: boolean;
  normalizeHeadingLevels: boolean;
  addSourceFrontmatter: boolean;
  openAfterImport: boolean;
}

export const DEFAULT_SETTINGS: SunDaoruSettings = {
  defaultTarget: "new",
  attachmentMode: "obsidian",
  customAttachmentFolder: "附件/导入",
  groupAttachmentsByNote: false,
  normalizeHeadingLevels: false,
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
      .setDesc("新建笔记，或插入当前打开的 Markdown 笔记。批量导入和文件夹导入始终新建笔记。")
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
      .setName("按文档建附件子文件夹")
      .setDesc(
        "开启时图片放到笔记同目录的「<文档名>-附件/」；插入当前笔记时放到当前笔记目录下。关闭后改用下面的附件位置设置（默认关闭）。",
      )
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.groupAttachmentsByNote)
          .onChange(async (value) => {
            this.plugin.settings.groupAttachmentsByNote = value;
            await this.plugin.saveSettings();
            this.display();
          }),
      );

    new Setting(containerEl)
      .setName("附件位置")
      .setDesc(
        this.plugin.settings.groupAttachmentsByNote
          ? "当前由「按文档建附件子文件夹」接管；关闭该开关后此设置生效。"
          : "图片等附件放在哪里。跟随 Obsidian 设置时，会使用「文件与链接」中的附件保存位置。",
      )
      .setDisabled(this.plugin.settings.groupAttachmentsByNote)
      .addDropdown((dropdown) =>
        dropdown
          .addOption("obsidian", "跟随 Obsidian 设置")
          .addOption("custom", "固定文件夹")
          .setValue(this.plugin.settings.attachmentMode)
          .onChange(async (value) => {
            this.plugin.settings.attachmentMode = value as AttachmentMode;
            await this.plugin.saveSettings();
            this.display();
          })
          .setDisabled(this.plugin.settings.groupAttachmentsByNote),
      );

    if (this.plugin.settings.attachmentMode === "custom") {
      new Setting(containerEl)
        .setName("附件文件夹")
        .setDesc("相对于库根目录，例如「附件/导入」。不存在的文件夹会自动创建。")
        .setDisabled(this.plugin.settings.groupAttachmentsByNote)
        .addText((text) =>
          text
            .setPlaceholder("附件/导入")
            .setValue(this.plugin.settings.customAttachmentFolder)
            .onChange(async (value) => {
              this.plugin.settings.customAttachmentFolder =
                value.trim() || DEFAULT_SETTINGS.customAttachmentFolder;
              await this.plugin.saveSettings();
            })
            .setDisabled(this.plugin.settings.groupAttachmentsByNote),
        );
    }

    new Setting(containerEl)
      .setName("整理标题层级")
      .setDesc(
        "检测到「文件标题 / Title」样式时，把其余标题整体下移一级。只对“一个文件标题 + 多级章节”的文档有效，默认关闭。",
      )
      .addToggle((toggle) =>
        toggle
          .setValue(this.plugin.settings.normalizeHeadingLevels)
          .onChange(async (value) => {
            this.plugin.settings.normalizeHeadingLevels = value;
            await this.plugin.saveSettings();
          }),
      );

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
      text:
        "字体、字号、颜色、对齐、页边距等排版信息不会导入；复杂表格保留为 HTML。页眉页脚、批注、公式等真正丢掉的内容会在导入结果里逐项列出，公式还会在正文原位置标出。",
    });
    note.createEl("p", {
      text: "文件夹导入会保留源文件夹名和内部层级，空文件夹与非 Word 文件跳过。",
    });
  }
}
