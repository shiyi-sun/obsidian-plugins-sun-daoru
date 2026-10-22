/** 纯函数版相对路径计算：不依赖 Obsidian，方便 Node 侧回归直接验证 */
export function relativePathFrom(fromFolder: string, toPath: string): string {
  const from = clean(fromFolder);
  const to = clean(toPath);
  const fromSegments = from ? from.split("/") : [];
  const toSegments = to.split("/");

  let common = 0;
  while (
    common < fromSegments.length &&
    common < toSegments.length - 1 &&
    fromSegments[common] === toSegments[common]
  ) {
    common++;
  }

  const up = fromSegments.length - common;
  return "../".repeat(up) + toSegments.slice(common).join("/");
}

function clean(path: string): string {
  return path
    .replace(/\\/g, "/")
    .replace(/\/{2,}/g, "/")
    .replace(/^\/+|\/+$/g, "");
}
