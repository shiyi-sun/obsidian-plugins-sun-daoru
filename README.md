# Daoru

English | [简体中文](README.zh-CN.md)

Daoru is a desktop-only Obsidian plugin that imports Microsoft Word documents (`.docx` and `.docm`) as Markdown notes. It converts the document structure, saves embedded images into your vault, and ends every import with a report of what was lost or downgraded, so you can see exactly what did not make it into the note.

Requires Obsidian 1.14.4 or later. The plugin interface is currently in Chinese; the display name in Obsidian is **Daoru** and the plugin ID is `sun-daoru`.

## Features

- Import one document, several documents at once, or a whole folder.
- Folder import keeps the source folder name and internal hierarchy.
- Create new notes or insert the converted content into the current note.
- Keep headings, lists, tables, links, images, and image alt text.
- Convert simple tables to Markdown; keep complex tables as HTML.
- Save images in Obsidian's configured attachment location, beside the note, or in a fixed folder.
- Report lost and downgraded content for every file.
- Run an extra text check and flag paragraphs that cannot be found in the generated note.
- Optionally write `source` and `imported` frontmatter, and normalize heading levels for documents built around a Title style.

## Usage

### Import documents

1. Click the import icon in the ribbon, or run `导入 Word 文档…` from the command palette.
2. Select one or more `.docx` or `.docm` files.
3. Choose whether to create a new note or insert into the current note, select a destination folder, and start the import.
4. Review the import result.

Multiple files are always imported as new notes. Inserting into the current note requires exactly one file and an open Markdown note in edit mode.

### Import a folder

Click the folder-import icon in the ribbon, or run `导入 Word 文件夹…` from the command palette. Daoru recreates the selected folder and its subfolders under the destination folder, then imports every Word file it finds.

Empty folders, non-Word files, hidden items (`~$` temporary files or names starting with a dot), `node_modules`, and `__MACOSX` are skipped. Folder import always creates new notes.

## What is converted

| Word content | Result in the note |
| --- | --- |
| Headings | Markdown headings |
| Bulleted and numbered lists | Markdown lists |
| Simple tables | Markdown tables |
| Complex tables | HTML tables, including merged cells and nested content |
| External links | Markdown links |
| Internal jumps | Link removed, link text kept |
| Images | Extracted as attachments; Word alt text becomes image alt text |
| Bold, italic, underline, strikethrough, superscript, subscript | Matching Markdown or HTML inline formatting |
| Text in text boxes | Kept when it can be read; otherwise listed in the import report |

## What is not imported

- Page layout and character formatting: fonts, font sizes, colors, alignment, spacing, indents, and margins.
- Headers and footers.
- Comments. The commented text itself is kept.
- Text deleted by unaccepted tracked changes. Inserted text is kept.
- Word equations (OMML). They are not converted; each equation in the visible body leaves a `【公式未导入】` marker at its original position and is counted in the report.
- Charts, SmartArt, and embedded OLE objects such as Excel worksheets.
- Images inserted as "link to file", because the image data is not stored inside the document.
- Field results such as page numbers, tables of contents, and cross-references. Some fields may be incomplete.
- Internal document jumps such as bookmarks, table-of-contents entries, and cross-reference links. The visible text is kept without a working link.

EMF, WMF, TIFF, and other image formats that Obsidian cannot preview are saved as files and shown as attachment links. The original is not discarded.

## Import report

Every import ends with a per-file report:

- **丢失 (lost)** — content that has no equivalent in the note.
- **降级 (downgraded)** — content kept in a reduced form, such as field results or non-previewable images.
- **说明 (info)** — useful notes, such as duplicate image references.

Lost items include comments, headers and footers, equations, charts and embedded objects, externally linked images, internal jumps, text from text boxes that could not be read, and deleted tracked changes. Each item shows a count and, where possible, a short sample of the affected content. Visible equations are also marked in the note at their original positions; searching for `公式未导入` jumps to each one.

Daoru also compares text segments from the document XML with the generated Markdown and reports paragraphs it cannot find. This catches content loss outside the known categories. If that warning appears, keep the original document for troubleshooting.

## Attachments

By default, attachments follow Obsidian's "Files & Links" attachment setting. You can enable "Per-document attachment folder" to save images to `<note name>-附件/` next to the note, or use a fixed folder relative to the vault root.

When the per-document folder is on, inserting into the current note places attachments next to that note.

## Settings

| Setting | Default | Description |
| --- | --- | --- |
| Default import action | New note | Create a new note or insert into the current note. Batch and folder imports always create new notes. |
| Per-document attachment folder | Off | Save images to `<note name>-附件/` beside the note. |
| Attachment location | Follow Obsidian | Used when the per-document folder is off. Follow Obsidian's Files & Links setting or use a fixed folder. |
| Attachment folder | `附件/导入` | Custom folder relative to the vault root. Created automatically when needed. |
| Normalize heading levels | Off | If the document uses a Title style, move other headings down one level. |
| Add source information on new notes | Off | Add `source` (original name or relative path) and `imported` (date) frontmatter. |
| Open note after import | On | Open the first successfully created note. |

The import dialog also lets you choose the destination folder and whether to add source information for the current batch.

## Commands and ribbon buttons

| Action | Where |
| --- | --- |
| Import Word documents | Ribbon import icon; command palette `导入 Word 文档…` |
| Import Word folder | Ribbon folder icon; command palette `导入 Word 文件夹…` |

## Installation

### From a release

1. Download the latest plugin package from [GitHub Releases](https://github.com/shiyi-sun/obsidian-plugins-sun-daoru/releases) or [AtomGit Releases](https://atomgit.com/shiyi-sun/obsidian-plugins-sun-daoru/releases).
2. Unzip it. The archive contains a `sun-daoru/` folder with `main.js`, `manifest.json`, and `styles.css`.
3. Copy that folder to `<vault>/.obsidian/plugins/sun-daoru/`.
4. Reload Obsidian, then enable **Daoru** under Settings → Community plugins.

When upgrading an existing installation, keep the folder name `sun-daoru` and replace the three files inside it.

### From source

```bash
git clone https://github.com/shiyi-sun/obsidian-plugins-sun-daoru.git
cd obsidian-plugins-sun-daoru
npm install
npm run build
```

Copy `main.js`, `manifest.json`, and `styles.css` into `<vault>/.obsidian/plugins/sun-daoru/`.

## Requirements and limitations

- Obsidian 1.14.4 or later, desktop only.
- `.docx` and `.docm` are supported. Legacy `.doc` files must be saved as `.docx` first.
- Encrypted or password-protected documents cannot be read.
- The interface is currently Chinese only.
- Everything runs locally. Documents are not uploaded anywhere.

## Feedback

Daoru is made by **sun**. WeChat official account: **十一营养师**.

Source and releases:

- [GitHub](https://github.com/shiyi-sun/obsidian-plugins-sun-daoru)
- [AtomGit](https://atomgit.com/shiyi-sun/obsidian-plugins-sun-daoru)

## License

MIT
