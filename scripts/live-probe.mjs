export default async function (ctx) {
  const info = await ctx.page(async () => {
    const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
    const started = Date.now();
    while (Date.now() - started < 30000) {
      if (window.app?.plugins?.plugins?.["sun-daoru"]) break;
      await sleep(300);
    }
    const plugin = window.app?.plugins?.plugins?.["sun-daoru"];
    const file = window.app?.vault?.getAbstractFileByPath("__daoru自测__/导入测试.md");
    if (file) {
      await window.app.workspace.getLeaf(false).openFile(file);
      await sleep(500);
    }
    return {
      vaultName: window.app?.vault?.getName?.(),
      vaultPath: window.app?.vault?.adapter?.getBasePath?.(),
      appVersion: window.app?.version ?? null,
      pluginLoaded: Boolean(plugin),
      pluginManifest: plugin?.manifest ?? null,
      commandExists: Boolean(window.app?.commands?.commands?.["sun-daoru:import-word"]),
      requireType: typeof require,
      windowRequireType: typeof window.require,
      processType: typeof process,
      activeFile: window.app?.workspace?.getActiveFile?.()?.path ?? null,
      plugins: Object.keys(window.app?.plugins?.plugins ?? {}),
      settings: plugin?.settings ?? null,
    };
  });

  ctx.note({ info });
  return { info };
}
