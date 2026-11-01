/*
导入锁与「取消」路径实机回归。

覆盖用户真的会做的取消动作：面板开在半路按 Esc、点右上角 X、点面板外的遮罩、
文件对话框取消、文件对话框既没 change 也没 cancel（靠窗口聚焦兜底）。
每一种都要满足：不产生笔记、不残留弹窗、plugin.importing 回到 false，
最后再跑一次真实导入证明锁没被锁死。

fixture 默认用仓库里的 scripts/fixtures/one-image.docx（含 1 张图）。
*/
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const FIXTURE = path.join(root, "scripts", "fixtures", "one-image.docx");
const SANDBOX = "__daoru自测__";
const NOTE = `${SANDBOX}/锁测试.md`;

export default async function (ctx) {
  if (!fs.existsSync(FIXTURE)) return { 汇总: `跳过：找不到 ${FIXTURE}`, checks: [] };
  const checks = [];
  const check = (name, ok, detail = "") =>
    checks.push({ 检查: name, 结论: ok ? "PASS" : "FAIL", 详情: String(detail).slice(0, 220) });

  const inPage = async (fn, ...args) => {
    try {
      return await ctx.page(fn, ...args);
    } catch (error) {
      return { __error: String(error).slice(0, 300) };
    }
  };

  // 1. 准备沙箱 + 把真实文件读进页面，并给 input.click 装一个可切换行为的桩
  const setup = await inPage(async (fixturePath, fixtureBytes, sandbox, note) => {
    const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
    const app = window.app;
    const plugin = app.plugins.plugins["sun-daoru"];
    // 上一轮崩在中途时可能留下没结束的选取框或没关的面板：先把导入锁放开
    window.dispatchEvent(new FocusEvent("focus"));
    const releaseStart = Date.now();
    while (plugin.importing && Date.now() - releaseStart < 8000) await sleep(150);
    for (let i = 0; i < 10; i++) {
      const close = document.querySelector(
        ".modal-container .modal-close-button, .sun-daoru-modal .modal-header-button",
      );
      if (!close) break;
      close.click();
      await sleep(150);
    }
    for (let i = 0; i < 5; i++) {
      if (!document.querySelector(".sun-daoru-modal")) break;
      document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
      await sleep(200);
    }
    const old = app.vault.getAbstractFileByPath(sandbox);
    if (old) await app.vault.delete(old, true);
    await app.vault.createFolder(sandbox);
    await app.vault.create(note, "锁测试。\n");
    await app.workspace.getLeaf(false).openFile(app.vault.getAbstractFileByPath(note));
    await sleep(300);

    plugin.settings.groupAttachmentsByNote = true;
    plugin.settings.defaultTarget = "new";
    plugin.settings.openAfterImport = false;
    await plugin.saveSettings();

    const buffer = Uint8Array.from(atob(fixtureBytes), (c) => c.charCodeAt(0));
    window.__lockFixture = { name: "锁测试单图.docx", buffer };
    window.__lockMode = "files"; // files | cancel | silent
    if (!window.__lockPatched) {
      window.__lockOriginalClick = HTMLInputElement.prototype.click;
      HTMLInputElement.prototype.click = function () {
        const input = this;
        if (window.__lockMode === "cancel") {
          setTimeout(() => input.dispatchEvent(new Event("cancel")), 30);
          return;
        }
        if (window.__lockMode === "silent") return; // 既不 change 也不 cancel
        const transfer = new DataTransfer();
        transfer.items.add(
          new File([window.__lockFixture.buffer.slice(0)], window.__lockFixture.name, {
            type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          }),
        );
        input.files = transfer.files;
        setTimeout(() => input.dispatchEvent(new Event("change", { bubbles: true })), 20);
      };
      window.__lockPatched = true;
    }
    return { importing: plugin.importing, files: app.vault.getFiles().length };
  }, FIXTURE, fs.readFileSync(FIXTURE).toString("base64"), SANDBOX, NOTE);
  if (setup.__error) return { 汇总: `准备失败：${setup.__error}`, checks };
  check("测试开始前导入锁是空闲", setup.importing === false, `importing=${setup.importing}`);

  const countFiles = () =>
    inPage(async (sandbox) => window.app.vault.getFiles().filter((f) => f.path.startsWith(sandbox)).length, SANDBOX);
  const openModalRect = (selector) =>
    inPage(async (selector) => {
      const all = document.querySelectorAll(selector);
      const el = all.length ? all[all.length - 1] : null;
      if (!el) return null;
      const r = el.getBoundingClientRect();
      return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) };
    }, selector);

  const startImport = () =>
    inPage(async () => {
      const before = document.querySelectorAll(".sun-daoru-modal").length;
      window.__lockMode = "files";
      window.app.commands.executeCommandById("sun-daoru:import-word");
      const started = Date.now();
      while (Date.now() - started < 15000) {
        const now = document.querySelectorAll(".sun-daoru-modal").length;
        if (now > before) return { modal: true };
        await new Promise((r) => setTimeout(r, 120));
      }
      return { modal: false, before };
    });

  const waitIdle = (ms) =>
    inPage(async (timeout) => {
      const started = Date.now();
      while (Date.now() - started < timeout) {
        if (!window.app.plugins.plugins["sun-daoru"].importing) return { idle: true };
        await new Promise((r) => setTimeout(r, 150));
      }
      return { idle: false, importing: window.app.plugins.plugins["sun-daoru"].importing };
    }, ms);

  const modalCount = () =>
    inPage(async () => document.querySelectorAll(".modal-container .modal").length);

  // 2. Esc 取消面板
  const a1 = await countFiles();
  const a2 = await startImport();
  check("Esc 用例：导入面板打开", a2.modal === true);
  const hasCloseButton = await inPage(async () =>
    !!document.querySelector(
      ".sun-daoru-modal .modal-close-button, .sun-daoru-modal .modal-header-button",
    ),
  );
  check(
    "导入面板提供右上角关闭按钮",
    hasCloseButton === true,
    hasCloseButton ? "有" : "面板 DOM 里没有关闭按钮，用户只能 Esc 或点遮罩关闭",
  );
  await ctx.key("Escape", "Escape", 27);
  const a3 = await waitIdle(4000);
  const a4 = await countFiles();
  const a5 = await modalCount();
  check("Esc 取消面板后导入锁释放", a3.idle === true, JSON.stringify(a3));
  check("Esc 取消面板后没有生成笔记", a4 === a1, `${a1} → ${a4}`);
  check("Esc 后面板真的关掉", a5 === 0, `残留 ${a5} 个弹窗`);

  // 3. 点面板右上角 X（若存在）
  const b1 = await countFiles();
  const b2 = await startImport();
  check("X 用例：导入面板打开", b2.modal === true);
  const closeRect = await openModalRect(
    ".sun-daoru-modal .modal-close-button, .sun-daoru-modal .modal-header-button",
  );
  let b3 = { idle: true, skipped: "没有 X 按钮，已用 Esc 收场" };
  if (closeRect) {
    await ctx.click(closeRect.x, closeRect.y);
    b3 = await waitIdle(4000);
  } else {
    await ctx.key("Escape", "Escape", 27);
    b3 = await waitIdle(4000);
  }
  const b4 = await countFiles();
  check("关闭面板后导入锁释放", b3.idle === true, JSON.stringify(b3));
  check("关面板后没有生成笔记", b4 === b1, `${b1} → ${b4}`);

  // 4. 点面板外的遮罩
  const c1 = await countFiles();
  const c2 = await startImport();
  check("遮罩用例：导入面板打开", c2.modal === true);
  const viewport = await inPage(async () => ({ w: window.innerWidth, h: window.innerHeight }));
  await ctx.click(6, Math.round(viewport.h / 2));
  const c3 = await waitIdle(4000);
  const c4 = await countFiles();
  check("点遮罩关闭面板后导入锁释放", c3.idle === true, JSON.stringify(c3));
  check("点遮罩后没有生成笔记", c4 === c1, `${c1} → ${c4}`);

  // 5. 文件对话框取消（派发 cancel 事件，验证插件里的 cancel 分支）
  const d1 = await countFiles();
  await inPage(async () => {
    window.__lockMode = "cancel";
    window.app.commands.executeCommandById("sun-daoru:import-word");
    await new Promise((r) => setTimeout(r, 1200));
    return document.querySelectorAll(".modal-container .modal").length;
  });
  const d2 = await waitIdle(4000);
  const d3 = await countFiles();
  const d4 = await modalCount();
  check("对话框 cancel 后不弹导入面板", d4 === 0, `弹窗 ${d4}`);
  check("对话框 cancel 后导入锁释放", d2.idle === true, JSON.stringify(d2));
  check("对话框 cancel 后没有生成笔记", d3 === d1, `${d1} → ${d3}`);

  // 6. 对话框既不 change 也不 cancel：靠窗口聚焦兜底
  await inPage(async () => {
    window.__lockMode = "silent";
    window.app.commands.executeCommandById("sun-daoru:import-word");
    await new Promise((r) => setTimeout(r, 600));
    return window.app.plugins.plugins["sun-daoru"].importing;
  });
  const silent = await inPage(async () => {
    window.dispatchEvent(new FocusEvent("focus"));
    const started = Date.now();
    while (Date.now() - started < 6000) {
      if (!window.app.plugins.plugins["sun-daoru"].importing) return { idle: true };
      await new Promise((r) => setTimeout(r, 150));
    }
    return { idle: false };
  });
  check("没有任何事件的选取框也能靠聚焦兜底解锁", silent.idle === true, JSON.stringify(silent));

  // 6b. 系统文件框还开着（导入锁已占用、Obsidian 面板未弹出）时，用真鼠标再点功能区图标
  await inPage(async () => {
    window.__lockMode = "silent";
    window.app.commands.executeCommandById("sun-daoru:import-word");
    await new Promise((r) => setTimeout(r, 500));
    return window.app.plugins.plugins["sun-daoru"].importing;
  });
  const ribbonRect = await inPage(async () => {
    const icon = Array.from(document.querySelectorAll(".side-dock-ribbon-action")).find((el) =>
      (el.getAttribute("aria-label") || "").includes("导入 Word 文档"),
    );
    if (!icon) return null;
    const r = icon.getBoundingClientRect();
    return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) };
  });
  if (ribbonRect) await ctx.click(ribbonRect.x, ribbonRect.y);
  const g1 = await inPage(async () => {
    // 连续采样 2 秒：只要「正在导入」提示在任何一刻出现过就算数
    const seen = new Set();
    let panels = 0;
    const started = Date.now();
    while (Date.now() - started < 2000) {
      panels = Math.max(panels, document.querySelectorAll(".sun-daoru-modal").length);
      for (const el of document.querySelectorAll(".notice"))
        seen.add((el.textContent || "").replace(/\s+/g, " ").slice(0, 40));
      await new Promise((r) => setTimeout(r, 60));
    }
    return { panels, notices: Array.from(seen) };
  });
  check("重复点击用例：找到功能区图标并真实点击", !!ribbonRect, JSON.stringify(ribbonRect));
  check("忙碌时重复点功能区不会开出第二个面板", g1.panels === 0, `面板数 ${g1.panels}`);
  check(
    "忙碌时重复点功能区有「正在导入」提示",
    g1.notices.some((text) => text.includes("正在导入")),
    `2 秒内出现的提示：${JSON.stringify(g1.notices)}`,
  );
  const g3 = await inPage(async () => {
    window.dispatchEvent(new FocusEvent("focus"));
    const started = Date.now();
    while (Date.now() - started < 6000) {
      if (!window.app.plugins.plugins["sun-daoru"].importing) return { idle: true };
      await new Promise((r) => setTimeout(r, 150));
    }
    return { idle: false };
  });
  check("重复点击用例收尾后锁已释放", g3.idle === true, JSON.stringify(g3));

  // 7. 连续取消后，真导入必须还能成功（证明锁没被锁死）
  const e1 = await countFiles();
  const e2 = await startImport();
  check("反复取消后仍能重新打开面板", e2.modal === true);
  const finish = await inPage(async () => {
    const all = document.querySelectorAll(".sun-daoru-modal");
    const modal = all[all.length - 1];
    const button = Array.from(modal.querySelectorAll("button")).find(
      (item) => (item.textContent || "").trim() === "开始导入",
    );
    button.click();
    const started = Date.now();
    while (Date.now() - started < 40000) {
      const result = document.querySelector(".sun-daoru-result-modal");
      if (result) return { result: (result.textContent || "").replace(/\s+/g, " ").slice(0, 200) };
      await new Promise((r) => setTimeout(r, 150));
    }
    return { result: null, importing: window.app.plugins.plugins["sun-daoru"].importing };
  });
  check("反复取消后真实导入成功", /成功|已导入|1/.test(finish.result || ""), JSON.stringify(finish).slice(0, 200));
  const e3 = await countFiles();
  check("真实导入确实写入了文件", e3 > e1, `${e1} → ${e3}`);
  const afterRun = await waitIdle(30000);
  check("导入完成后导入锁释放", afterRun.idle === true, JSON.stringify(afterRun));

  // 8. 收尾：还原桩、删沙箱、还原设置
  const cleanup = await inPage(async (sandbox) => {
    const app = window.app;
    const plugin = app.plugins.plugins["sun-daoru"];
    for (let i = 0; i < 10; i++) {
      const button = Array.from(
        document.querySelectorAll(".sun-daoru-result-modal button"),
      ).find((item) => item.textContent?.trim() === "完成");
      if (!button) break;
      button.click();
      await new Promise((r) => setTimeout(r, 200));
    }
    if (window.__lockPatched) {
      HTMLInputElement.prototype.click = window.__lockOriginalClick;
      window.__lockPatched = false;
    }
    delete window.__lockFixture;
    delete window.__lockMode;
    plugin.settings.groupAttachmentsByNote = false;
    plugin.settings.defaultTarget = "new";
    plugin.settings.openAfterImport = true;
    await plugin.saveSettings();
    const folder = app.vault.getAbstractFileByPath(sandbox);
    if (folder) await app.vault.delete(folder, true);
    return {
      sandbox: app.vault.getAbstractFileByPath(sandbox) ? "还在" : "已删",
      importing: plugin.importing,
      modals: document.querySelectorAll(".modal-container .modal").length,
      strays: app.vault.getFiles().filter((f) => !f.path.startsWith(".")).map((f) => f.path),
    };
  }, SANDBOX);
  check("沙箱清理干净且无残留弹窗", cleanup.sandbox === "已删" && cleanup.modals === 0, JSON.stringify(cleanup).slice(0, 260));

  const failed = checks.filter((c) => c.结论 !== "PASS");
  return {
    汇总: failed.length ? `${failed.length} 项失败 / ${checks.length} 项` : `全部通过（${checks.length} 项）`,
    checks,
    残留: cleanup.strays,
  };
}
