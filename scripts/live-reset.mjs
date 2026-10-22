export default async function (ctx) {
  const result = await ctx.page(async () => {
    const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
    const app = window.app;
    const plugin = app.plugins.plugins["sun-daoru"];
    const started = Date.now();
    while (plugin?.importing && Date.now() - started < 180000) {
      await sleep(500);
    }
    for (let i = 0; i < 20; i++) {
      const close = document.querySelector(".modal-container .modal-close-button");
      if (!close) break;
      if (close instanceof HTMLElement) close.click();
      await sleep(200);
    }
    await sleep(300);

    if (plugin) {
      plugin.settings.attachmentMode = "obsidian";
      plugin.settings.customAttachmentFolder = "附件/导入";
      plugin.settings.groupAttachmentsByNote = true;
      plugin.settings.normalizeHeadingLevels = false;
      plugin.settings.openAfterImport = true;
      plugin.settings.defaultTarget = "new";
      plugin.settings.addSourceFrontmatter = false;
      await plugin.saveSettings();
    }

    const folder = app.vault.getAbstractFileByPath("__daoru自测__");
    if (folder) await app.vault.delete(folder, true);
    const welcome = app.vault.getAbstractFileByPath("欢迎.md");
    if (welcome) await app.workspace.getLeaf(false).openFile(welcome);
    await sleep(500);

    return {
      idle: !plugin?.importing,
      files: app.vault.getFiles().map((file) => file.path),
      modals: document.querySelectorAll(".modal-container").length,
      settings: plugin?.settings ?? null,
    };
  });
  return result;
}
