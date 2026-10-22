import { App, Editor, TFile, normalizePath } from "obsidian";
import { convertDocx } from "./core/docx";
import { htmlToMarkdown } from "./core/markdown";
import {
  ensureFolder,
  formatLocalDate,
  parentFolderOf,
  sanitizeBaseName,
  splitAttachmentName,
  uniqueFilePath,
} from "./paths";
import type { ImportTarget, SunDaoruSettings } from "./settings";

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
  if (request.target === "insert" && files.length > 1) {
    throw new Error("插入当前笔记一次只能处理一个文件。");
  }

  const results: ImportFileResult[] = [];
  for (let index = 0; index < files.length; index++) {
    const file = files[index];
    const prefix = files.length > 1 ? `(${index + 1}/${files.length}) ` : "";
    results.push(
      await importSingleFile(app, file, request, settings, onProgress, prefix),
    );
  }
  return results;
}

async function importSingleFile(
  app: App,
  file: File,
  request: ImportRequest,
  settings: SunDaoruSettings,
  onProgress: ProgressCallback,
  prefix: string,
): Promise<ImportFileResult> {
  const fileName = file.name;
  const result: ImportFileResult = {
    fileName,
    inserted: false,
    images: 0,
    simpleTables: 0,
    complexTables: 0,
    warnings: [],
  };

  const createdAttachments: TFile[] = [];
  try {
    assertSupportedFile(fileName);

    onProgress(`${prefix}正在读取 ${fileName}…`);
    const data = await file.arrayBuffer();

    onProgress(`${prefix}正在解析 Word（大文档可能需要一会儿）…`);
    const conversion = await convertDocx(data);

    const baseName = sanitizeBaseName(fileName);
    let notePath: string | undefined;
    if (request.target === "new") {
      const folder = normalizePath(request.folder || "");
      await ensureFolder(app, folder);
      notePath = await uniqueFilePath(app, folder, baseName, ".md");
      result.notePath = notePath;
    }

    const attachmentSourcePath =
      request.target === "new" ? notePath || "" : request.activeFilePath;
    const imageBase = baseName.slice(0, 60).replace(/[. ]+$/g, "") || "导入";
    const totalImages = conversion.images.length;

    for (let index = 0; index < totalImages; index++) {
      const image = conversion.images[index];
      const ordinal = String(index + 1).padStart(3, "0");
      const attachmentName = `${imageBase}-${ordinal}.${image.extension}`;
      const attachmentPath = await resolveAttachmentPath(
        app,
        attachmentName,
        attachmentSourcePath,
        settings,
      );
      await ensureFolder(app, parentFolderOf(attachmentPath));
      const created = await app.vault.createBinary(attachmentPath, image.data);
      image.link = created.path;
      createdAttachments.push(created);

      if (index === 0 || (index + 1) % 10 === 0 || index + 1 === totalImages) {
        onProgress(`${prefix}正在保存图片 ${index + 1}/${totalImages}…`);
      }
    }

    const markdownResult = htmlToMarkdown(conversion.html, conversion.images);
    const markdown = request.target === "new" && request.addSourceFrontmatter
      ? withSourceFrontmatter(markdownResult.markdown, fileName)
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

    const nonPreviewable = conversion.images.filter((image) => !image.previewable).length;
    if (nonPreviewable > 0) {
      result.warnings.push(
        `${nonPreviewable} 张图片（如 EMF/WMF/TIFF）Obsidian 无法预览，已保存原文件并以链接形式插入。`,
      );
    }

    return result;
  } catch (error) {
    await cleanupAttachments(app, createdAttachments);
    result.error = errorMessage(error);
    result.warnings = [];
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
      settings.customAttachmentFolder.trim() || "附件/Word 导入",
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

function withSourceFrontmatter(markdown: string, fileName: string): string {
  const frontmatter = [
    "---",
    `source: ${JSON.stringify(fileName)}`,
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

function errorMessage(error: unknown): string {
  if (error instanceof Error && error.message) return error.message;
  return String(error) || "未知错误";
}
