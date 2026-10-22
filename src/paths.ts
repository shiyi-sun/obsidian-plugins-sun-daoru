import { App, normalizePath, TFolder } from "obsidian";

export function sanitizeBaseName(raw: string): string {
  let base = raw
    .replace(/\.(docx|docm)$/i, "")
    .replace(/[\\/:*?"<>|#^\[\]]/g, "-")
    .replace(/[\u0000-\u001f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/[. ]+$/g, "");

  if (!base) base = "导入";
  const characters = Array.from(base);
  if (characters.length > 80) base = characters.slice(0, 80).join("").trim();
  if (/^(con|prn|aux|nul|com[1-9]|lpt[1-9])$/i.test(base)) {
    base = `导入-${base}`;
  }
  return base;
}

export function sanitizeFolderName(raw: string): string {
  let name = raw
    .replace(/[\\/:*?"<>|#^\[\]]/g, "-")
    .replace(/[\u0000-\u001f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/[. ]+$/g, "");

  if (!name) name = "未命名";
  const characters = Array.from(name);
  if (characters.length > 80) name = characters.slice(0, 80).join("").trim();
  if (/^(con|prn|aux|nul|com[1-9]|lpt[1-9])$/i.test(name)) {
    name = `导入-${name}`;
  }
  return name;
}

export function sanitizeRelativeFolder(relative: string): string {
  return relative
    .replace(/\\/g, "/")
    .split("/")
    .filter(Boolean)
    .map(sanitizeFolderName)
    .join("/");
}

export function joinVaultPath(folder: string, name: string): string {
  const cleanFolder = normalizePath(folder || "").replace(/^\/+|\/+$/g, "");
  return normalizePath(cleanFolder ? `${cleanFolder}/${name}` : name);
}

export function parentFolderOf(path: string): string {
  const normalized = normalizePath(path);
  const index = normalized.lastIndexOf("/");
  return index === -1 ? "" : normalized.slice(0, index);
}

export function splitAttachmentName(fileName: string): {
  base: string;
  extension: string;
} {
  const index = fileName.lastIndexOf(".");
  if (index <= 0) return { base: fileName, extension: "" };
  return {
    base: fileName.slice(0, index),
    extension: fileName.slice(index + 1),
  };
}

export async function ensureFolder(app: App, folder: string): Promise<void> {
  const normalized = normalizePath(folder || "").replace(/^\/+|\/+$/g, "");
  if (!normalized || normalized === ".") return;

  const segments = normalized.split("/").filter(Boolean);
  let current = "";
  for (const segment of segments) {
    current = current ? `${current}/${segment}` : segment;
    const existing = app.vault.getAbstractFileByPath(current);
    if (existing) {
      if (!(existing instanceof TFolder)) {
        throw new Error(`路径已被同名文件占用：${current}`);
      }
      continue;
    }
    try {
      await app.vault.createFolder(current);
    } catch (error) {
      if (!(app.vault.getAbstractFileByPath(current) instanceof TFolder)) {
        throw error;
      }
    }
  }
}

export async function uniqueFilePath(
  app: App,
  folder: string,
  base: string,
  extension: string,
): Promise<string> {
  const ext = extension.startsWith(".") ? extension : `.${extension}`;
  for (let index = 1; index < 1000; index++) {
    const name = index === 1 ? `${base}${ext}` : `${base} ${index}${ext}`;
    const path = joinVaultPath(folder, name);
    if (!app.vault.getAbstractFileByPath(path)) return path;
  }
  throw new Error(`同名文件太多，无法为「${base}」生成新文件名。`);
}

export function formatLocalDate(date = new Date()): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}
