/** @walnut_method
 * name: Export Submitted Form Data to Excel
 * description: Export submitted form data to XLSX file ${filePath} worksheet ${sheetName}
 * actionType: custom_export_submitted_form_to_excel
 * context: web
 * needsLocator: false
 * category: Data Processing
 * modules: exceljs
 */
import type { WalnutWebContext } from './walnut';
import ExcelJS from 'exceljs';
import fs from 'node:fs';
import path from 'node:path';

export async function exportSubmittedFormToExcel(ctx: WalnutWebContext): Promise<void> {
  const filePath = ctx.args[0]?.trim();
  const sheetName = ctx.args[1]?.trim();

  if (!filePath) throw new Error('An XLSX output file path is required.');
  if (!sheetName) throw new Error('An XLSX worksheet name is required.');

  const entries = await ctx.page.evaluate(() => {
    const rows = Array.from(document.querySelectorAll('.modal-content .table tbody tr'));
    return rows
      .map((row) => {
        const cells = row.querySelectorAll('td');
        return {
          field: (cells[0]?.textContent ?? '').trim(),
          value: (cells[1]?.textContent ?? '').trim(),
        };
      })
      .filter((entry) => entry.field.length > 0);
  });

  if (entries.length === 0) {
    throw new Error('No submitted-form summary was found. Submit the form before exporting.');
  }

  const outputPath = path.resolve(filePath);
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });

  const workbook = new ExcelJS.Workbook();
  if (fs.existsSync(outputPath)) await workbook.xlsx.readFile(outputPath);

  let worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) worksheet = workbook.addWorksheet(sheetName);

  const timestampHeader = 'Exported At';
  const requiredHeaders = [timestampHeader, ...entries.map((entry) => entry.field)];
  const headerRow = worksheet.getRow(1);
  const headers: string[] = [];
  headerRow.eachCell({ includeEmpty: false }, (cell) => {
    headers[cell.col - 1] = String(cell.value ?? '').trim();
  });

  for (const header of requiredHeaders) {
    if (!headers.includes(header)) {
      headers.push(header);
      headerRow.getCell(headers.length).value = header;
    }
  }

  if (headerRow.cellCount > 0) {
    headerRow.font = { bold: true };
    headerRow.commit();
  }

  const record: Record<string, string> = {
    [timestampHeader]: new Date().toISOString(),
  };
  for (const entry of entries) record[entry.field] = entry.value;

  worksheet.addRow(headers.map((header) => record[header] ?? '')).commit();
  worksheet.columns = headers.map((header) => ({
    width: Math.min(40, Math.max(14, header.length + 4)),
  }));

  await workbook.xlsx.writeFile(outputPath);
  ctx.log(`Exported ${entries.length} form field(s) to ${outputPath} (${sheetName}).`);
}
