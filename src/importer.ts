import { App, Editor, TFile, normalizePath } from "obsidian";
import { describeError } from "./core/errors";
import { convertDocx } from "./core/docx";
import { htmlToMarkdown } from "./core/markdown";
import { checkTextFidelity } from "./core/losses";
import type { LossItem } from "./core/losses";
import { relativePathFrom } from "./core/relative-path";
import {
  ensureFolder,
  formatLocalDate,
  joinVaultPath,
  parentFolderOf,
  sanitizeBaseName,
  sanitizeRelativeFolder,
  splitAttachmentName,
  uniqueFilePath,
} from "./paths";
import type { ImportTarget, SunDaoruSettings } from "./settings";

export interface ImportItem {
  file: File;
  /** 文件夹导入时的源相对路径，含选中的根文件夹名 */
  relativePath?: string;
}

export interface ImportRequest {
  target: ImportTarget;
  folder: string;
  addSourceFrontmatter: boolean;
  editor: Editor | null;
  activeFilePath: string;
}

export interface ImportFileResult {
  fileName: string;
  notePath?: string;
  inserted: boolean;
  images: number;
  simpleTables: number;
  complexTables: number;
  warnings: string[];
  losses: LossItem[];
  error?: string;
}

type ProgressCallback = (message: string) => void;

export async function importFiles(
  app: App,
  files: File[],
  request: ImportRequest,
  settings: SunDaoruSettings,
  onProgress: ProgressCallback,
): Promise<ImportFileResult[]> {
  return importItems(
    app,
    files.map((file) => ({ file })),
    request,
    settings,
    onProgress,
  );
}

export async function importItems(
  app: App,
  items: ImportItem[],
  request: ImportRequest,
  settings: SunDaoruSettings,
  onProgress: ProgressCallback,
): Promise<ImportFileResult[]> {
  if (request.target === "insert" && items.length > 1) {
    throw new Error("插入当前笔记一次只能处理一个文件。");
  }

  const results: ImportFileResult[] = [];
  for (let index = 0; index < items.length; index++) {
    const prefix = items.length > 1 ? `(${index + 1}/${items.length}) ` : "";
    results.push(
      await importSingleFile(
        app,
        items[index],
        request,
        settings,
        onProgress,
        prefix,
      ),
    );
  }
  return results;
}

async function importSingleFile(
  app: App,
  item: ImportItem,
  request: ImportRequest,
  settings: SunDaoruSettings,
  onProgress: ProgressCallback,
  prefix: string,
): Promise<ImportFileResult> {
  const file = item.file;
  const fileName = file.name;
  const displayName = item.relativePath || fileName;
  const result: ImportFileResult = {
    fileName: displayName,
    inserted: false,
    images: 0,
    simpleTables: 0,
    complexTables: 0,
    warnings: [],
    losses: [],
  };

  const createdAttachments: TFile[] = [];
  try {
    assertSupportedFile(fileName);

    onProgress(`${prefix}正在读取 ${displayName}…`);
    const data = await file.arrayBuffer();

    onProgress(`${prefix}正在解析 Word（大文档可能需要一会儿）…`);
    const conversion = await convertDocx(data, {
      normalizeHeadingLevels: settings.normalizeHeadingLevels,
    });
    if (
      settings.normalizeHeadingLevels &&
      !conversion.headingLevelsShifted
    ) {
      conversion.warnings.push(
        "已开启标题层级整理，但文档里没有检测到「文件标题 / Title」样式，标题层级未调整。",
      );
    }

    const baseName = sanitizeBaseName(fileName);
    let targetFolder = normalizePath(request.folder || "");
    if (item.relativePath) {
      const relativeDir = item.relativePath
        .replace(/\\/g, "/")
        .split("/")
        .slice(0, -1)
        .join("/");
      const safeRelative = sanitizeRelativeFolder(relativeDir);
      if (safeRelative) targetFolder = joinVaultPath(targetFolder, safeRelative);
    }

    let notePath: string | undefined;
    if (request.target === "new") {
      await ensureFolder(app, targetFolder);
      notePath = await uniqueFilePath(app, targetFolder, baseName, ".md");
      result.notePath = notePath;
    }

    const attachmentSourcePath =
      request.target === "new" ? notePath || "" : request.activeFilePath;
    const imageBase = baseName.slice(0, 60).replace(/[. ]+$/g, "") || "导入";
    const noteBase = notePath
      ? notePath
          .slice(notePath.lastIndexOf("/") + 1)
          .replace(/\.md$/i, "")
      : "";
    const groupBase =
      (request.target === "new" ? noteBase : imageBase)
        .slice(0, 60)
        .replace(/[. ]+$/g, "") || imageBase;
    const totalImages = conversion.images.length;

    const noteFolder =
      request.target === "new"
        ? parentFolderOf(notePath || "")
        : parentFolderOf(request.activeFilePath);

    let groupingFolder = "";
    if (settings.groupAttachmentsByNote && totalImages > 0) {
      groupingFolder = joinVaultPath(noteFolder, `${groupBase}-附件`);
      await ensureFolder(app, groupingFolder);
    }

    for (let index = 0; index < totalImages; index++) {
      const image = conversion.images[index];
      const ordinal = String(index + 1).padStart(3, "0");
      let attachmentPath: string;

      if (groupingFolder) {
        attachmentPath = await uniqueFilePath(
          app,
          groupingFolder,
          ordinal,
          image.extension,
        );
      } else {
        const attachmentName = `${imageBase}-${ordinal}.${image.extension}`;
        attachmentPath = await resolveAttachmentPath(
          app,
          attachmentName,
          attachmentSourcePath,
          settings,
        );
        await ensureFolder(app, parentFolderOf(attachmentPath));
      }

      const created = await app.vault.createBinary(attachmentPath, image.data);
      image.link = relativePathFrom(noteFolder, created.path);
      createdAttachments.push(created);

      if (index === 0 || (index + 1) % 10 === 0 || index + 1 === totalImages) {
        onProgress(`${prefix}正在保存图片 ${index + 1}/${totalImages}…`);
      }
    }

    const markdownResult = htmlToMarkdown(conversion.html, conversion.images);
    const markdown =
      request.target === "new" && request.addSourceFrontmatter
        ? withSourceFrontmatter(markdownResult.markdown, displayName)
        : markdownResult.markdown;

    if (request.target === "insert") {
      if (!request.editor) {
        throw new Error("当前没有打开的 Markdown 笔记，无法插入。");
      }
      onProgress(`${prefix}正在插入当前笔记…`);
      request.editor.replaceSelection(markdown);
      result.inserted = true;
      result.notePath = request.activeFilePath;
    } else {
      onProgress(`${prefix}正在写入笔记…`);
      const note = await app.vault.create(notePath!, markdown);
      result.notePath = note.path;
    }

    result.images = conversion.images.length;
    result.simpleTables = markdownResult.simpleTables;
    result.complexTables = markdownResult.complexTables;
    result.warnings.push(...conversion.warnings, ...markdownResult.warnings);

    const nonPreviewable = conversion.images.filter(
      (image) => !image.previewable,
    ).length;
    const withAlt = conversion.images.filter((image) => image.altText).length;
    const unknownLoss = checkTextFidelity(
      conversion.documentXml,
      markdownResult.markdown,
    );

    result.losses = [
      ...conversion.losses,
      ...markdownResult.losses,
      ...(unknownLoss ? [unknownLoss] : []),
      ...(nonPreviewable > 0
        ? [
            {
              kind: "downgraded" as const,
              label: "Obsidian 显示不了的图片格式",
              count: nonPreviewable,
              detail: "EMF/WMF/TIFF 等原文件已保存，正文里改成指向附件的链接。",
            },
          ]
        : []),
      ...(withAlt > 0
        ? [
            {
              kind: "info" as const,
              label: "图片替代文字",
              count: withAlt,
              detail: "已作为图片的替代文字（alt）写出，页面上不显示。",
            },
          ]
        : []),
    ];

    return result;
  } catch (error) {
    await cleanupAttachments(app, createdAttachments);
    result.error = describeError(error);
    result.warnings = [];
    result.losses = [];
    return result;
  }
}

function assertSupportedFile(fileName: string): void {
  if (/\.doc$/i.test(fileName)) {
    throw new Error("旧版 .doc 不支持，请在 Word 或 WPS 中另存为 .docx 后重试。");
  }
  if (!/\.(docx|docm)$/i.test(fileName)) {
    throw new Error("只支持 .docx / .docm 文件。");
  }
}

async function resolveAttachmentPath(
  app: App,
  fileName: string,
  sourcePath: string,
  settings: SunDaoruSettings,
): Promise<string> {
  const { base, extension } = splitAttachmentName(fileName);

  if (settings.attachmentMode === "custom") {
    const folder = normalizePath(
      settings.customAttachmentFolder.trim() || "附件/导入",
    );
    await ensureFolder(app, folder);
    return uniqueFilePath(app, folder, base, extension);
  }

  try {
    const path = await app.fileManager.getAvailablePathForAttachment(
      fileName,
      sourcePath || undefined,
    );
    return normalizePath(path);
  } catch {
    const folder = parentFolderOf(sourcePath);
    return uniqueFilePath(app, folder, base, extension);
  }
}

function withSourceFrontmatter(markdown: string, sourceName: string): string {
  const frontmatter = [
    "---",
    `source: ${JSON.stringify(sourceName)}`,
    `imported: ${formatLocalDate()}`,
    "---",
  ].join("\n");
  return `${frontmatter}\n\n${markdown}`;
}

async function cleanupAttachments(app: App, files: TFile[]): Promise<void> {
  for (const file of files.reverse()) {
    try {
      await app.vault.delete(file, true);
    } catch {
      // 回滚是尽力而为；留下附件不会破坏笔记数据。
    }
  }
}
