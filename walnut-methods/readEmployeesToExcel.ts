import type { WalnutWebContext } from './walnut';
import ExcelJS from 'exceljs';
import fs from 'node:fs';
import path from 'node:path';

/** @walnut_method
 * name: Read Employees and Export to Excel
 * description: Login with ${username} and ${password}, then read all employee-list pages and write the roster to ${filePath} worksheet ${sheetName}
 * actionType: custom_read_employees_to_excel
 * context: web
 * needsLocator: false
 * category: Data Processing
 * modules: exceljs
 */
export async function readEmployeesToExcel(ctx: WalnutWebContext): Promise<void> {
  const [username, password, filePath, sheetName] = ctx.args.map((value) => value?.trim() ?? '');
  if (!username) throw new Error('A username is required.');
  if (!password) throw new Error('A password is required.');
  if (!filePath) throw new Error('An XLSX output file path is required.');
  if (!sheetName) throw new Error('An XLSX worksheet name is required.');
  if (!filePath.toLowerCase().endsWith('.xlsx')) throw new Error('The output file must use the .xlsx extension.');

  const page = ctx.page;
  const loginUrl = new URL('/web/index.php/auth/login', ctx.testBaseUrl).toString();
  await ctx.navigate(loginUrl);

  const usernameField = page.getByRole('textbox', { name: 'Username' });
  if (await usernameField.isVisible().catch(() => false)) {
    await usernameField.fill(username);
    await page.getByRole('textbox', { name: 'Password' }).fill(password);
    await page.getByRole('button', { name: 'Login' }).click();
  }

  const pimLink = page.getByRole('link', { name: 'PIM', exact: true });
  await pimLink.waitFor({ state: 'visible', timeout: 30000 });
  await pimLink.click();

  const rowSelector = '.oxd-table-body .oxd-table-row';
  const rowsLocator = page.locator(rowSelector);
  await rowsLocator.first().waitFor({ state: 'visible', timeout: 30000 });

  const bodyText = await page.locator('body').innerText();
  const recordMatch = bodyText.match(/\(([\d,]+)\)\s*Records Found/);
  if (!recordMatch) throw new Error('Could not read the employee-list record count.');
  const expectedCount = Number(recordMatch[1].replace(/,/g, ''));
  if (!Number.isFinite(expectedCount) || expectedCount < 1) throw new Error('The employee list contains no records.');

  const headerCells = await page.locator('.oxd-table-header [role="columnheader"]').allInnerTexts();
  const columns = headerCells.map((text: string) => text.trim()).slice(1, -1);
  if (columns.length !== 7) {
    throw new Error(`Expected 7 roster columns, found ${columns.length}: ${columns.join(', ')}`);
  }

  const records: string[][] = [];
  const visitedFirstRows = new Set<string>();
  let pageNumber = 1;

  while (records.length < expectedCount) {
    const pageRows = await rowsLocator.count();
    if (pageRows < 1) throw new Error(`No employee rows found on page ${pageNumber}.`);

    const firstRowSignature = (await rowsLocator.first().innerText()).trim();
    if (visitedFirstRows.has(firstRowSignature)) {
      throw new Error(`Pagination repeated a page before all ${expectedCount} records were read.`);
    }
    visitedFirstRows.add(firstRowSignature);

    for (let i = 0; i < pageRows; i++) {
      const cells = (await rowsLocator.nth(i).locator('.oxd-table-cell').allInnerTexts()).map((text: string) => text.trim());
      const record = cells.slice(1, -1);
      if (record.length !== columns.length) {
        throw new Error(`Employee row ${records.length + 1} has ${record.length} values; expected ${columns.length}.`);
      }
      records.push(record);
    }

    if (records.length > expectedCount) {
      throw new Error(`Read ${records.length} rows, exceeding the application's count of ${expectedCount}.`);
    }
    if (records.length === expectedCount) break;

    const paginationButtons = page.locator('button.oxd-pagination-page-item--previous-next');
    let nextButton: any = null;
    for (let i = 0; i < await paginationButtons.count(); i++) {
      const candidate = paginationButtons.nth(i);
      if (await candidate.locator('.bi-chevron-right').count()) {
        nextButton = candidate;
        break;
      }
    }
    if (!nextButton || await nextButton.isDisabled()) {
      throw new Error(`Pagination ended after ${records.length} of ${expectedCount} employee records.`);
    }

    const selectedPage = page.locator('.oxd-pagination-page-item--page-selected');
    const previousPageText = (await selectedPage.innerText()).trim();
    await nextButton.click();
    await page.waitForFunction(
      ({ pageText, rowSignature }: { pageText: string; rowSignature: string }) => {
        const currentPage = document.querySelector('.oxd-pagination-page-item--page-selected');
        const firstRow = document.querySelector('.oxd-table-body .oxd-table-row');
        return (currentPage?.textContent ?? '').trim() !== pageText &&
          (firstRow?.innerText ?? '').trim() !== rowSignature;
      },
      { pageText: previousPageText, rowSignature: firstRowSignature },
      { timeout: 30000 },
    );
    pageNumber += 1;
  }

  const outputPath = path.resolve(filePath);
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });

  const workbook = new ExcelJS.Workbook();
  if (fs.existsSync(outputPath)) await workbook.xlsx.readFile(outputPath);
  const existingSheet = workbook.getWorksheet(sheetName);
  if (existingSheet) workbook.removeWorksheet(existingSheet.id);
  const worksheet = workbook.addWorksheet(sheetName);

  worksheet.addRow(columns);
  for (const record of records) worksheet.addRow(record);
  worksheet.views = [{ state: 'frozen', ySplit: 1 }];
  worksheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: records.length + 1, column: columns.length },
  };

  const header = worksheet.getRow(1);
  header.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
  header.commit();

  for (let columnIndex = 0; columnIndex < columns.length; columnIndex++) {
    const maxValueLength = records.reduce(
      (max, row) => Math.max(max, (row[columnIndex] ?? '').length),
      columns[columnIndex].length,
    );
    worksheet.getColumn(columnIndex + 1).width = Math.min(36, Math.max(14, maxValueLength + 2));
  }

  await workbook.xlsx.writeFile(outputPath);
  ctx.log(`Exported ${records.length} employee roster row(s) across ${pageNumber} page(s) to ${outputPath} (${sheetName}).`);
}
