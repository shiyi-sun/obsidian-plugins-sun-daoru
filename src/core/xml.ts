export function decodeXml(value: string): string {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, digits: string) => String.fromCodePoint(Number(digits)))
    .replace(/&#x([0-9a-f]+);/gi, (_, digits: string) =>
      String.fromCodePoint(parseInt(digits, 16)),
    )
    .replace(/&amp;/g, "&");
}

export function xmlValue(xml: string, tag: string): string | undefined {
  const escapedTag = tag.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return new RegExp(`<${escapedTag}\\b[^>]*\\bw:val="([^"]*)"`).exec(xml)?.[1];
}

export function countMatches(xml: string, pattern: RegExp): number {
  let total = 0;
  for (const _ of xml.matchAll(pattern)) total++;
  return total;
}
