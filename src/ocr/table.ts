export function parseTableFromOCRText(text: string): string[][] {
  const rows = text.trim().split(/\r?\n/);
  return rows.map((r) => r.split(/\s{2,}|\t/).filter((c) => c.trim().length > 0));
}
