/* Fully offline Office-JS helper */
export async function write(text: string): Promise<void> {
  await Excel.run(async (ctx) => {
    const rng = ctx.workbook.getSelectedRange();
    rng.values = [[text]];
    rng.format.autofitColumns();
    await ctx.sync();
  });
}

export async function writeTick(): Promise<void> {
  await write("✓");
}
export async function writeCross(): Promise<void> {
  await write("✗");
}

/* Persistent metadata in hidden sheet */
export async function logSnip(record: {
  cell: string;
  page: number;
  rect: string;
  mode: string;
  text: string;
}): Promise<void> {
  await Excel.run(async (ctx) => {
    let sheet = ctx.workbook.worksheets.getItemOrNullObject("_Snips");
    await ctx.sync();
    if (sheet.isNullObject) {
      sheet = ctx.workbook.worksheets.add("_Snips");
      sheet.visibility = Excel.SheetVisibility.hidden;
    }
    const last = sheet.getRange("A:A").getUsedRange().rowCount + 1;
    sheet.getRange(`A${last}:E${last}`).values = [
      [record.cell, record.page, record.rect, record.mode, record.text]
    ];
    await ctx.sync();
  });
}

export async function getCurrentCellAddress(): Promise<string> {
  return await Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange();
    range.load("address");
    await ctx.sync();
    return range.address;
  });
}

export async function writeTableToCell(tableData: string[][], startCell?: string): Promise<void> {
  await Excel.run(async (ctx) => {
    let range;

    if (startCell) {
      range = ctx.workbook.worksheets.getActiveWorksheet().getRange(startCell);
    } else {
      range = ctx.workbook.getSelectedRange();
    }

    const targetRange = range.getResizedRange(tableData.length - 1, tableData[0].length - 1);
    targetRange.values = tableData;
    targetRange.format.autofitColumns();
    await ctx.sync();
  });
}

export interface SnipRecord {
  cell: string;
  page: number;
  rect: string; // JSON string {x,y,width,height}
  mode: string;
  text: string;
}

export async function findSnipByCell(cellAddress: string): Promise<SnipRecord | null> {
  return await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItemOrNullObject("_Snips");
    await ctx.sync();

    if (sheet.isNullObject) return null;

    const used = sheet.getUsedRange();
    used.load(["values", "rowCount"]);
    await ctx.sync();

    for (const row of used.values as string[][]) {
      if (row[0] === cellAddress) {
        const [cell, page, rect, mode, text] = row;
        return {
          cell,
          page: Number(page),
          rect,
          mode,
          text
        } as SnipRecord;
      }
    }
    return null;
  });
}
