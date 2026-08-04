import type { WalnutContext } from './walnut';
import ExcelJS from 'exceljs';

/** @walnut_method
 * name: Get Excel Cell Value by Column Header
 * description: Open workbook ${filePath} on sheet ${sheetName} and fetch row ${rowNumber} value under column header ${columnHeader} and store in $[cellValue]
 * actionType: custom_get_excel_cell_value
 * context: shared
 * needsLocator: false
 * category: Data Processing
 */
export async function getExcelCellValue(ctx: WalnutContext) {
  // ctx.args resolved in order of ${...} and $[...] placeholders in description:
  // args[0] = filePath      (from ${filePath})
  // args[1] = sheetName     (from ${sheetName})
  // args[2] = rowNumber     (from ${rowNumber})  — 1-based data row, excluding the header row
  // args[3] = columnHeader  (from ${columnHeader})
  // args[4] = "cellValue"   (from $[cellValue])  — runtime variable name to store result

  const filePath     = ctx.args[0];
  const sheetName    = ctx.args[1];
  const rowNumber    = parseInt(ctx.args[2], 10);
  const columnHeader = ctx.args[3];
  const outputVar    = ctx.args[4];

  ctx.log(`Opening workbook: ${filePath}`);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  // Validate sheet exists
  const sheet = workbook.getWorksheet(sheetName);
  if (!sheet) {
    const available = workbook.worksheets.map((ws: ExcelJS.Worksheet) => ws.name).join(', ');
    throw new Error(`Sheet "${sheetName}" not found. Available sheets: ${available}`);
  }

  // Row 1 = header row in ExcelJS (1-based)
  const headerRow = sheet.getRow(1);
  let colNumber = -1;

  headerRow.eachCell({ includeEmpty: false }, (cell: ExcelJS.Cell, colIdx: number) => {
    if (String(cell.value ?? '').trim().toLowerCase() === columnHeader.trim().toLowerCase()) {
      colNumber = colIdx;
    }
  });

  if (colNumber === -1) {
    const headers: string[] = [];
    headerRow.eachCell({ includeEmpty: false }, (cell: ExcelJS.Cell) => headers.push(String(cell.value ?? '')));
    throw new Error(
      `Column header "${columnHeader}" not found. Available headers: ${headers.join(', ')}`
    );
  }

  // rowNumber is 1-based data row → Excel row = rowNumber + 1 (offset by header)
  const excelRowNumber = rowNumber + 1;
  const totalRows = sheet.rowCount;

  if (excelRowNumber > totalRows) {
    throw new Error(
      `Row ${rowNumber} is out of range. Sheet has ${totalRows - 1} data row(s).`
    );
  }

  const dataRow = sheet.getRow(excelRowNumber);
  const cell = dataRow.getCell(colNumber);

  // Resolve the display value (handles formulas, rich text, dates, plain values)
  let cellValue: string;
  const raw = cell.value;

  if (raw === null || raw === undefined) {
    cellValue = '';
  } else if (typeof raw === 'object' && 'result' in (raw as any)) {
    // Formula cell — use the cached result
    cellValue = String((raw as ExcelJS.CellFormulaValue).result ?? '');
  } else if (typeof raw === 'object' && 'richText' in (raw as any)) {
    // Rich text cell
    cellValue = (raw as ExcelJS.CellRichTextValue).richText.map((r) => r.text).join('');
  } else if (raw instanceof Date) {
    cellValue = raw.toISOString().slice(0, 10);
  } else {
    cellValue = String(raw);
  }

  ctx.log(
    `Fetched value at sheet="${sheetName}", column="${columnHeader}", row=${rowNumber}: ${cellValue}`
  );

  ctx.setVariable(outputVar, cellValue);
}
