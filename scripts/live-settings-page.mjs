/*
设置页实机检查（独立脚本）。
Obsidian 1.14.4 把设置开在独立窗口：主窗口 target 有 app 但没有设置 DOM，
设置窗口 target 有设置 DOM 但没有 app。这里同时连两个窗口，
用真实鼠标点设置项，再从主窗口读回插件设置值，验证 UI ↔ 数据双向生效。
*/
import { writeFileSync } from "node:fs";

const PORT = 9223;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function attach(match) {
  const list = await (await fetch(`http://127.0.0.1:${PORT}/json/list`)).json();
  const target = list.find(match);
  if (!target) throw new Error("找不到 target");
  const ws = new WebSocket(target.webSocketDebuggerUrl);
  await new Promise((resolve, reject) => {
    ws.addEventListener("open", resolve, { once: true });
    ws.addEventListener("error", reject, { once: true });
  });
  let id = 0;
  const send = (method, params) =>
    new Promise((resolve, reject) => {
      const mid = ++id;
      const onMessage = (event) => {
        const msg = JSON.parse(event.data);
        if (msg.id !== mid) return;
        ws.removeEventListener("message", onMessage);
        if (msg.error) reject(new Error(msg.error.message));
        else resolve(msg.result);
      };
      ws.addEventListener("message", onMessage);
      ws.send(JSON.stringify({ id: mid, method, params }));
    });
  await send("Runtime.enable", {});
  return {
    async eval(body) {
      const r = await send("Runtime.evaluate", {
        expression: `(async () => { ${body} })()`,
        awaitPromise: true,
        returnByValue: true,
      });
      if (r.exceptionDetails)
        throw new Error(r.exceptionDetails.exception?.description || JSON.stringify(r.exceptionDetails));
      return r.result.value;
    },
    async click(x, y) {
      for (const [type, buttons] of [
        ["mousePressed", 1],
        ["mouseReleased", 0],
      ]) {
        await send("Input.dispatchMouseEvent", { type, x, y, button: "left", clickCount: 1, buttons });
        await sleep(40);
      }
    },
    async screenshot(file) {
      await send("Page.enable", {});
      const r = await send("Page.captureScreenshot", { format: "png" });
      writeFileSync(file, Buffer.from(r.data, "base64"));
      return file;
    },
    close: () => ws.close(),
  };
}

const results = [];
const check = (label, ok, detail = "") => results.push({ label, ok: !!ok, detail: String(detail).slice(0, 240) });

const main = await attach((t) => t.type === "page" && t.url.includes("index.html"));
await main.eval(`
  for (let i = 0; i < 10; i++) {
    const close = document.querySelector(
      ".modal-container .modal-close-button, .sun-daoru-modal .modal-header-button",
    );
    if (!close) break;
    close.click();
    await new Promise((r) => setTimeout(r, 150));
  }
  app.setting.open();
  await new Promise((r) => setTimeout(r, 800));
  app.setting.openTabById("sun-daoru");
  await new Promise((r) => setTimeout(r, 800));
  return true;
`);
const win = await attach((t) => t.type === "page" && t.title.startsWith("设置"));

const readPage = () =>
  win.eval(`
    const content = document.querySelector(".vertical-tab-content");
    const items = Array.from(content.querySelectorAll(".setting-item")).map((el) => ({
      name: (el.querySelector(".setting-item-name")?.textContent || "").trim(),
      desc: (el.querySelector(".setting-item-description")?.textContent || "").trim(),
      heading: el.classList.contains("setting-item-heading"),
      control: el.querySelector(".checkbox-container") ? "toggle"
        : el.querySelector("select") ? "select"
        : el.querySelector("input[type=text]") ? "text"
        : el.querySelector("button") ? "button" : "无",
      checked: el.querySelector("input[type=checkbox]")?.checked ?? null,
      isOn: el.querySelector(".checkbox-container")?.classList.contains("is-enabled") ?? null,
      boxClass: el.querySelector(".checkbox-container")?.className ?? null,
      value: el.querySelector("input[type=text]")?.value ?? null,
      selectText: el.querySelector("select") ? Array.from(el.querySelector("select").options).find((o) => o.selected)?.textContent.trim() : null,
      options: el.querySelector("select") ? Array.from(el.querySelector("select").options).map((o) => o.textContent.trim()) : null,
      box: (() => { const t = el.querySelector(".checkbox-container") || el.querySelector("select") || el.querySelector("input[type=text]"); if (!t) return null;
        const r = t.getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2), top: Math.round(r.top), bottom: Math.round(r.bottom) }; })(),
    }));
    const select = content.querySelector("select");
    return {
      items,
      text: content.textContent || "",
      viewport: { w: window.innerWidth, h: window.innerHeight },

    };
  `);

const view0 = await readPage();
const view = view0;
console.log("调试：每行控件", JSON.stringify(view.items.map((i) => ({ n: i.name, c: i.control, checked: i.checked, opts: i.options })), null, 0));
const settingsDump = () =>
  main.eval(`const s = app.plugins.plugins["sun-daoru"].settings; return {
    defaultTarget: s.defaultTarget, groupAttachmentsByNote: s.groupAttachmentsByNote,
    attachmentMode: s.attachmentMode, customAttachmentFolder: s.customAttachmentFolder,
    normalizeHeadingLevels: s.normalizeHeadingLevels, addSourceFrontmatter: s.addSourceFrontmatter,
    openAfterImport: s.openAfterImport };`);

const before = await settingsDump();
console.log("调试：视口", JSON.stringify(view.viewport), "下拉选项", JSON.stringify(view.defaultOptionTexts));
console.log("调试：开关行", JSON.stringify(view.items.filter((i) => i.control === "toggle").map((i) => ({ n: i.name, c: i.checked, top: i.box.top, bottom: i.box.bottom }))));
check("设置页第一项是「导入」标题", view.items[0]?.name === "导入" && view.items[0]?.heading, view.items.map((i) => i.name).join(" / "));
const expected = [
  "默认导入方式",
  "按文档建附件子文件夹",
  "附件位置",
  "整理标题层级",
  "新建笔记时写入来源信息",
  "导入后打开笔记",
];
const missing = expected.filter((name) => !view.items.some((i) => i.name === name));
check("六个设置项都在", missing.length === 0, missing.length ? `缺少 ${missing.join("、")}` : `${view.items.length} 行`);
check("说明里没有 Obsidian 专有语法残留", !/\[\[|==[^=]/.test(view.text));
check("说明里承诺会列出丢失内容", view.text.includes("逐项列出"));
const mismatches = view.items
  .filter((i) => i.control === "toggle")
  .map((i) => {
    const key = i.name.includes("按文档建")
      ? "groupAttachmentsByNote"
      : i.name.includes("整理标题")
        ? "normalizeHeadingLevels"
        : i.name.includes("来源")
          ? "addSourceFrontmatter"
          : i.name.includes("打开笔记")
            ? "openAfterImport"
            : null;
    return key && (i.isOn ?? i.checked) !== before[key] ? `${i.name} 界面=${i.isOn}/${i.checked} 数据=${before[key]}` : null;
  })
  .filter(Boolean);
check("开关默认状态与设置一致", mismatches.length === 0, mismatches.join("；") || "4 个开关都一致");

const scrollRowIntoView = (name) =>
  win.eval(`
    const el = Array.from(document.querySelectorAll(".vertical-tab-content .setting-item"))
      .find((row) => (row.querySelector(".setting-item-name")?.textContent || "").trim() === ${JSON.stringify(name)});
    if (!el) return false;
    el.scrollIntoView({ block: "center" });
    await new Promise((r) => setTimeout(r, 250));
    return true;
  `);

// 真鼠标逐个点开关（先滚到可视区）→ 主窗口读回设置值 → 再点回去
const flipped = [];
for (const item of view.items.filter((i) => i.control === "toggle")) {
  await scrollRowIntoView(item.name);
  const fresh = (await readPage()).items.find((i) => i.name === item.name);
  await win.click(fresh.box.x, fresh.box.y);
  await sleep(300);
  const after = await settingsDump();
  const now = (await readPage()).items.find((i) => i.name === item.name);
  const key = Object.keys(after).find((k) => {
    const name = item.name;
    if (k === "groupAttachmentsByNote") return name.includes("按文档建");
    if (k === "normalizeHeadingLevels") return name.includes("整理标题");
    if (k === "addSourceFrontmatter") return name.includes("来源");
    if (k === "openAfterImport") return name.includes("打开笔记");
    return false;
  });
  flipped.push({
    设置项: item.name,
    点前: before[key],
    点后: after[key],
    界面: now.isOn ?? now.checked,
    点击处: [fresh.box.x, fresh.box.y],
    视口高: fresh.bottom,
  });
  await scrollRowIntoView(item.name);
  const back = (await readPage()).items.find((i) => i.name === item.name);
  await win.click(back.box.x, back.box.y);
  await sleep(300);
}
check(
  "真鼠标点开关后设置值翻转",
  flipped.length === 4 && flipped.every((f) => f.点后 === !f.点前 && f.界面 === f.点后),
  JSON.stringify(flipped)
);
const restored = await settingsDump();
check("再点一次全部还原", JSON.stringify(restored) === JSON.stringify(before), JSON.stringify(restored));

// 附件位置：分组开着时应该置灰；关掉分组后能切「固定文件夹」并出现文件夹输入框
const setGroup = async (wanted) => {
  if ((await settingsDump()).groupAttachmentsByNote === wanted) return;
  await scrollRowIntoView("按文档建附件子文件夹");
  const row = (await readPage()).items.find((i) => i.name === "按文档建附件子文件夹");
  await win.click(row.box.x, row.box.y);
  await sleep(400);
};
await setGroup(true);
const disabledWithGroup = await win.eval(`
  const el = Array.from(document.querySelectorAll(".vertical-tab-content .setting-item"))
    .find((row) => (row.querySelector(".setting-item-name")?.textContent || "").trim() === "附件位置");
  const select = el.querySelector("select");
  return { selectDisabled: select.disabled, desc: (el.querySelector(".setting-item-description")?.textContent || "").trim() };
`);
check(
  "开启「按文档建附件子文件夹」时，附件位置置灰并解释原因",
  disabledWithGroup.selectDisabled === true && disabledWithGroup.desc.includes("接管"),
  JSON.stringify(disabledWithGroup)
);

await setGroup(false);
const enabledAfterUngroup = await win.eval(`
  const el = Array.from(document.querySelectorAll(".vertical-tab-content .setting-item"))
    .find((row) => (row.querySelector(".setting-item-name")?.textContent || "").trim() === "附件位置");
  return el.querySelector("select").disabled;
`);
check("关掉分组开关后附件位置可选", enabledAfterUngroup === false, `disabled=${enabledAfterUngroup}`);

const customOption = await win.eval(`
  const el = Array.from(document.querySelectorAll(".vertical-tab-content .setting-item"))
    .find((row) => (row.querySelector(".setting-item-name")?.textContent || "").trim() === "附件位置");
  const select = el.querySelector("select");
  const texts = Array.from(select.options).map((o) => o.textContent.trim());
  const index = texts.findIndex((t) => t.includes("固定文件夹"));
  if (index < 0) return { texts, picked: null };
  select.selectedIndex = index;
  select.dispatchEvent(new Event("change", { bubbles: true }));
  await new Promise((r) => setTimeout(r, 600));
  return { texts, picked: texts[index] };
`);
const afterMode = await readPage();
const folderItem = afterMode.items.find((i) => i.name === "附件文件夹");
check(
  "附件位置能切到「固定文件夹」",
  !!customOption.picked && (await settingsDump()).attachmentMode === "custom",
  JSON.stringify(customOption)
);
check("切到固定文件夹后出现附件文件夹输入框", !!folderItem, afterMode.items.map((i) => i.name).join(" / "));
check(
  "附件文件夹默认值是 附件/导入",
  folderItem?.value === "附件/导入",
  `输入框实际值：${folderItem?.value}`
);
check("附件文件夹说明写明相对库根", (folderItem?.desc || "").includes("相对于库根目录"), folderItem?.desc);

// 还原：固定文件夹 → 跟随 Obsidian，并把分组开关恢复到测试前状态
await win.eval(`
  const el = Array.from(document.querySelectorAll(".vertical-tab-content .setting-item"))
    .find((row) => (row.querySelector(".setting-item-name")?.textContent || "").trim() === "附件位置");
  const select = el.querySelector("select");
  select.selectedIndex = Array.from(select.options).findIndex((o) => o.value === "obsidian");
  select.dispatchEvent(new Event("change", { bubbles: true }));
  await new Promise((r) => setTimeout(r, 500));
  return true;
`);
await setGroup(before.groupAttachmentsByNote);

const finalDump = await settingsDump();
check(
  "收尾：附件位置与分组恢复测试前设置",
  finalDump.attachmentMode === before.attachmentMode &&
    finalDump.groupAttachmentsByNote === before.groupAttachmentsByNote,
  JSON.stringify(finalDump),
);

await win.screenshot(`${process.env.TEMP}\\daoru-settings-page.png`).catch(() => {});
await main.eval(`app.setting.close(); return true;`);
win.close();
main.close();

const failed = results.filter((r) => !r.ok);
for (const r of results) console.log(`${r.ok ? "✓" : "✗"} ${r.label}${r.detail ? "  · " + r.detail : ""}`);
console.log(`\n设置页 ${results.length - failed.length}/${results.length} 项通过`);
if (failed.length) process.exitCode = 1;
