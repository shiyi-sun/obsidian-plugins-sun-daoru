/*
阅读视图下的「插入当前笔记」守卫：实测过 editor.replaceSelection 在阅读视图不落盘，
所以面板必须收掉插入选项并说明原因。切视图要用真鼠标点模式按钮。
*/
export default async function (ctx) {
  const fixture = {
    path: "C:/Users/shiyi/Desktop/kaifa/obsidian/sun-daoru/妊娠期糖尿病营养干预方案.docx",
    name: "妊娠期糖尿病营养干预方案.docx",
  };

  const opened = await ctx.page(async () => {
    const app = window.app;
    const plugin = app.plugins.plugins["sun-daoru"];
    const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
    const old = app.vault.getAbstractFileByPath("__daoru自测__");
    if (old) await app.vault.delete(old, true);
    await app.vault.createFolder("__daoru自测__");
    const note = await app.vault.create("__daoru自测__/插入守卫.md", "原有内容。\n");
    const leaf = app.workspace.getLeaf(false);
    await leaf.openFile(note);
    await sleep(500);
    plugin.settings.openAfterImport = false;
    plugin.settings.defaultTarget = "insert";
    await plugin.saveSettings();
    const rect = leaf.view.modeButtonEl?.getBoundingClientRect?.();
    return {
      modeBefore: leaf.view.getMode?.() ?? null,
      button: rect
        ? { x: Math.round(rect.x + rect.width / 2), y: Math.round(rect.y + rect.height / 2) }
        : null,
    };
  });

  const checks = [];
  const check = (ok, name, detail = "") =>
    checks.push({ 检查: name, 结论: ok ? "PASS" : "FAIL", 详情: String(detail).slice(0, 160) });

  check(opened.modeBefore === "source", "起点是编辑视图", String(opened.modeBefore));
  if (!opened.button) {
    return { checks, 汇总: "失败：找不到模式切换按钮" };
  }

  let mode = opened.modeBefore;
  for (let attempt = 0; attempt < 3 && mode !== "preview"; attempt++) {
    await ctx.click(opened.button.x, opened.button.y);
    await ctx.sleep(1000);
    mode = await ctx.page(async () => window.app.workspace.getLeaf(false).view.getMode?.() ?? null);
  }
  check(mode === "preview", "已切到阅读视图", String(mode));

  const readingPanel = await ctx.page(async (fixtureJson) => {
    const item = JSON.parse(fixtureJson);
    const app = window.app;
    const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
    const wait = async (predicate, timeout = 30000) => {
      const started = Date.now();
      while (Date.now() - started < timeout) {
        const value = await predicate();
        if (value) return value;
        await sleep(120);
      }
      return null;
    };

    window.__guard = { origClick: HTMLInputElement.prototype.click, item };
    HTMLInputElement.prototype.click = function () {
      if (this.type !== "file") return window.__guard.origClick.call(this);
      const fsx = require("fs");
      const transfer = new DataTransfer();
      const buffer = fsx.readFileSync(window.__guard.item.path);
      transfer.items.add(
        new File([buffer], window.__guard.item.name, {
          type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        }),
      );
      this.files = transfer.files;
      this.dispatchEvent(new Event("change", { bubbles: true }));
    };
    app.commands.executeCommandById("sun-daoru:import-word");
    const modal = await wait(() => document.querySelector(".sun-daoru-modal"));
    const select = modal?.querySelector("select");
    const options = Array.from(select?.querySelectorAll("option") || []).map((o) => o.value);
    const desc = Array.from(modal?.querySelectorAll(".setting-item-description") || [])
      .map((el) => el.textContent?.trim())
      .find((text) => text && text.length > 4);
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    await sleep(300);
    if (window.__guard.origClick) HTMLInputElement.prototype.click = window.__guard.origClick;
    delete window.__guard;
    return { options, desc };
  }, JSON.stringify(fixture));

  check(
    !readingPanel.options.includes("insert"),
    "阅读视图下面板不提供插入选项",
    readingPanel.options.join(","),
  );
  check(
    (readingPanel.desc || "").includes("阅读视图"),
    "面板说明了不可插入的原因",
    readingPanel.desc || "",
  );

  const restored = await ctx.page(async (fixtureJson) => {
    const item = JSON.parse(fixtureJson);
    const app = window.app;
    const plugin = app.plugins.plugins["sun-daoru"];
    const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
    const wait = async (predicate, timeout = 30000) => {
      const started = Date.now();
      while (Date.now() - started < timeout) {
        const value = await predicate();
        if (value) return value;
        await sleep(120);
      }
      return null;
    };
    const leaf = app.workspace.getLeaf(false);

    window.__guard = { origClick: HTMLInputElement.prototype.click, item };
    HTMLInputElement.prototype.click = function () {
      if (this.type !== "file") return window.__guard.origClick.call(this);
      const fsx = require("fs");
      const transfer = new DataTransfer();
      const buffer = fsx.readFileSync(window.__guard.item.path);
      transfer.items.add(
        new File([buffer], window.__guard.item.name, {
          type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        }),
      );
      this.files = transfer.files;
      this.dispatchEvent(new Event("change", { bubbles: true }));
    };
    app.commands.executeCommandById("sun-daoru:import-word");
    const modal = await wait(() => document.querySelector(".sun-daoru-modal"));
    const select = modal?.querySelector("select");
    const options = Array.from(select?.querySelectorAll("option") || []).map((o) => o.value);
    Array.from(modal?.querySelectorAll("button") || [])
      .find((button) => button.textContent?.includes("开始导入"))
      ?.click();
    const result = await wait(() => document.querySelector(".sun-daoru-result-modal"), 60000);
    const resultText = (result?.textContent || "").slice(0, 80);
    if (window.__guard.origClick) HTMLInputElement.prototype.click = window.__guard.origClick;
    delete window.__guard;
    Array.from(document.querySelectorAll(".sun-daoru-result-modal button"))
      .find((button) => button.textContent?.trim() === "完成")
      ?.click();
    await sleep(400);

    const created = app.vault.getAbstractFileByPath(
      "__daoru自测__/妊娠期糖尿病营养干预方案.md",
    );
    const content = created ? await app.vault.read(created) : "";
    const original = app.vault.getAbstractFileByPath("__daoru自测__/插入守卫.md");
    const originalContent = original ? await app.vault.read(original) : "";
    const sandbox = app.vault.getAbstractFileByPath("__daoru自测__");
    if (sandbox) await app.vault.delete(sandbox, true);
    plugin.settings.defaultTarget = "new";
    plugin.settings.openAfterImport = true;
    await plugin.saveSettings();
    const welcome = app.vault.getAbstractFileByPath("欢迎.md");
    if (welcome) await app.workspace.getLeaf(false).openFile(welcome);
    await sleep(300);
    return {
      结果: resultText,
      新建笔记产出: Boolean(created) && content.includes("体重监控"),
      原笔记未被偷偷改写: originalContent === "原有内容。\n",
      清理后: app.vault.getFiles().map((f) => f.path),
    };
  }, JSON.stringify(fixture));

  check(restored.新建笔记产出, "阅读视图下改用新建笔记仍正常产出", restored.结果);
  check(restored.原笔记未被偷偷改写, "原笔记内容保持原样");
  check(restored.清理后.every((p) => !p.includes("__daoru自测__")), "沙箱已清理", restored.清理后.join(" | "));

  const failed = checks.filter((item) => item.结论 === "FAIL");
  return {
    checks,
    汇总: failed.length === 0 ? `全部通过（${checks.length} 项）` : `失败 ${failed.length} 项`,
    清理后库内文件: restored.清理后,
  };
}
