/** 把各种来源的异常翻译成用户看得懂的话 */
export function describeError(error: unknown): string {
  const raw = (error instanceof Error ? error.message : String(error)).trim();
  if (!raw) return "导入失败：未知错误。";
  // 我们自己抛的提示都是中文，直接显示；第三方库和 Obsidian 的英文报错要包一层人话
  if (/[一-鿿]/.test(raw)) return raw;
  if (/end with a dot or a space|outside the vault|invalid (?:path|name|file name)/i.test(raw)) {
    return `附件文件夹路径不合法（${raw}）。请在设置里换一个正常的文件夹名，别用「..」、纯点号或结尾留空格。`;
  }
  if (/reading '[^']*' of undefined|charAt|undefined is not a function/i.test(raw)) {
    return `这份文档结构异常，内部有指向不存在内容的引用，无法解析（${raw}）。如果它在 Word 里能正常打开，可以把文件发给开发者排查。`;
  }
  return `这份文档解析失败：${raw}。如果它在 Word 里能正常打开，可以把文件发给开发者排查。`;
}
