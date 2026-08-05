import type { WalnutWebContext } from './walnut';
import ExcelJS from 'exceljs';
import { Worker, isMainThread, parentPort, workerData } from 'worker_threads';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';



/** @walnut_method
 * name: Parallel Search IDs from Excel and Write Results (2 Workers × 2 Tabs)
 * description: Open workbook ${filePath} sheet ${inputSheetName} column ${idColumnName} search each ID in parallel using 2 browser workers with 2 tabs each from date ${fromDate} to date ${toDate} and write results into sheet ${outputSheetName}
 * actionType: custom_parallel_search_and_export_to_excel
 * context: web
 * needsLocator: false
 * category: Data Processing
 */
export async function parallelSearchAndExportToExcel(ctx: WalnutWebContext) {
  const filePath        = ctx.args[0];
  const inputSheetName  = ctx.args[1];
  const idColumnName    = ctx.args[2];
  const fromDate        = ctx.args[3];
  const toDate          = ctx.args[4];
  const outputSheetName = (ctx.args[5] ?? '').trim() || `Output_${new Date().toISOString().slice(0,19).replace(/[:T]/g,'-')}`;

  // ── LOCATORS — identical to searchAndExportToExcel ──────────────────────────
  const LOCATORS = {
    searchInputHost:  'ui5-input#searchFieldInput',
    searchInputInner: 'input#inner',
    tabTimeExpense:   '//a[@id="tab_timeAndExpense"]',
    fromDateInput:    '//input[@id="filterStartDate"]',
    toDateInput:      '//input[@id="filterEndDate"]',
    applyFilterBtn:   '//input[@id="timeAndExpenseFitlerBtn"]',
    resultRow:        'div[role="row"][id*="timeSheet_workOrder_list_byWorkerId"]',
    columns: [
      { index: 0,  header: 'Status',        child: 'span.fd-object-status__text'  },
      { index: 1,  header: 'Timesheet ID',  child: 'a.archiveLink'                },
      { index: 2,  header: 'Start Date',    child: 'div.jqx-grid-cell-left-align' },
      { index: 3,  header: 'End Date',      child: 'div.jqx-grid-cell-left-align' },
      { index: 4,  header: 'Approved Date', child: 'div.jqx-grid-cell-left-align' },
      { index: 5,  header: 'ST Hours',      child: 'div.jqx-grid-cell-left-align' },
      { index: 6,  header: 'OT Hours',      child: 'div.jqx-grid-cell-left-align' },
      { index: 7,  header: 'DT Hours',      child: 'div.jqx-grid-cell-left-align' },
      { index: 8,  header: 'Others Hours',  child: 'div.jqx-grid-cell-left-align' },
      { index: 9,  header: 'NB Hours',      child: 'div.jqx-grid-cell-left-align' },
      { index: 10, header: 'Amount (INR)',   child: 'div.jqx-grid-cell-left-align' },
    ],
    workerName:     'h1[data-help-id="TITLE_270"] span.titlePrimary',
    pageReadyProbe: '//a[@id="tab_timeAndExpense"]',
    homeBtn:        '//li[@id="homeMenuTitle"]//a[@title="Home"]',
    homeReadyProbe: 'ui5-input#searchFieldInput',
  } as const;

  // ── Helpers shared by main thread ────────────────────────────────────────────
  function cellText(value: ExcelJS.CellValue): string {
    if (value === null || value === undefined) return '';
    if (typeof value === 'object') {
      if ('result' in value) return String((value as ExcelJS.CellFormulaValue).result ?? '');
      if ('richText' in value)
        return (value as ExcelJS.CellRichTextValue).richText.map((r) => r.text).join('');
      if ('text' in value) return (value as ExcelJS.CellHyperlinkValue).text;
      if (value instanceof Date) return value.toISOString().slice(0, 10);
    }
    return String(value);
  }

  function styleHeaderRow(row: ExcelJS.Row, colCount: number): void {
    for (let c = 1; c <= colCount; c++) {
      const cell = row.getCell(c);
      cell.font      = { bold: true };
      cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
      cell.border    = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' },
      };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    }
    row.commit();
  }

  function styleDataRow(row: ExcelJS.Row, colCount: number, dataIndex: number): void {
    const bgColor = dataIndex % 2 === 0 ? 'FFFFFFFF' : 'FFF2F2F2';
    for (let c = 1; c <= colCount; c++) {
      const cell = row.getCell(c);
      cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
      cell.border    = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' },
      };
      cell.alignment = { vertical: 'middle', wrapText: true };
    }
    row.commit();
  }

  // ── Generic chunk splitter ───────────────────────────────────────────────────
  function splitIntoChunks<T>(arr: T[], chunkCount: number): T[][] {
    const chunks: T[][] = Array.from({ length: chunkCount }, () => []);
    arr.forEach((item, i) => chunks[i % chunkCount].push(item));
    return chunks;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // STEP 1 — Read workbook ONCE, extract IDs, then close workbook
  // ═══════════════════════════════════════════════════════════════════════════
  ctx.log('Reading Workbook...');

  const inputWorkbook = new ExcelJS.Workbook();
  await inputWorkbook.xlsx.readFile(filePath);

  const inSheet = inputWorkbook.getWorksheet(inputSheetName);
  if (!inSheet) {
    const available = inputWorkbook.worksheets.map((ws: ExcelJS.Worksheet) => ws.name).join(', ');
    throw new Error(`Input sheet "${inputSheetName}" not found. Available: ${available}`);
  }

  // Locate the ID column by header name
  const headerRow = inSheet.getRow(1);
  let idColIndex   = -1;
  headerRow.eachCell({ includeEmpty: false }, (cell: ExcelJS.Cell, colIdx: number) => {
    if (String(cell.value ?? '').trim().toLowerCase() === idColumnName.trim().toLowerCase()) {
      idColIndex = colIdx;
    }
  });
  if (idColIndex === -1) {
    const found: string[] = [];
    headerRow.eachCell({ includeEmpty: false }, (cell: ExcelJS.Cell) =>
      found.push(String(cell.value ?? '')),
    );
    throw new Error(
      `Column "${idColumnName}" not found in sheet "${inputSheetName}". ` +
      `Found headers: ${found.join(', ')}`,
    );
  }

  const ids: string[] = [];
  inSheet.eachRow({ includeEmpty: false }, (row: ExcelJS.Row, rowNumber: number) => {
    if (rowNumber === 1) return;
    const val = cellText(row.getCell(idColIndex).value).trim();
    if (val !== '') ids.push(val);
  });

  if (ids.length === 0)
    throw new Error(`No IDs found in column "${idColumnName}" of sheet "${inputSheetName}".`);

  ctx.log(`${ids.length} IDs Loaded`);

  // Determine output column order before releasing the workbook
  const defaultColOrder: string[] = ['WOID', 'Worker Name', ...LOCATORS.columns.map((c) => c.header)];
  let colOrder: string[] = defaultColOrder;
  {
    const existingSheet = inputWorkbook.getWorksheet(outputSheetName);
    if (existingSheet) {
      const hdrs: string[] = [];
      existingSheet.getRow(1).eachCell({ includeEmpty: false }, (cell: ExcelJS.Cell) =>
        hdrs.push(String(cell.value ?? '')),
      );
      if (hdrs.length > 0) colOrder = hdrs;
    }
  }

  // Release workbook from memory — workers must NOT touch the file
  // @ts-ignore
  inputWorkbook._worksheets = [];

  // ═══════════════════════════════════════════════════════════════════════════
  // STEP 2 — Split IDs into 4 equal chunks
  // ═══════════════════════════════════════════════════════════════════════════
  ctx.log('Splitting into 4 Chunks...');
  const chunks = splitIntoChunks(ids, 4);
  chunks.forEach((chunk, i) =>
    ctx.log(`  Chunk${i + 1}: ${chunk.length} IDs`),
  );

  // ═══════════════════════════════════════════════════════════════════════════
  // STEP 3 — Worker thread script (runs inside Worker, no ctx available)
  //
  //   Each worker:
  //     1. Launches one Chromium browser
  //     2. Navigates to baseUrl and assumes login is already active
  //        (session cookies / storage are supplied via storageState)
  //     3. Opens two tabs (pages) concurrently, one per chunk
  //     4. Processes each chunk, returns combined result array
  //     5. Closes the browser
  //
  //   Communication: workerData (input) → parentPort.postMessage (output)
  // ═══════════════════════════════════════════════════════════════════════════

  const workerScript = /* javascript */ `
const { workerData, parentPort } = require('worker_threads');
const { chromium } = require(workerData.playwrightPath);

const {
  workerId,
  baseUrl,
  authStorageState,   // serialised Playwright storage state (cookies + localStorage)
  chunk1,
  chunk2,
  fromDate,
  toDate,
  locators,
} = workerData;

function log(msg) {
  parentPort.postMessage({ type: 'log', msg: \`[Worker\${workerId}] \${msg}\` });
}

// ── Wait until the grid has at least one real Timesheet ID ───────────────────
async function waitForGridReady(page, timeoutMs = 90000) {
  try {
    await page.waitForFunction(
      ({ rowSel }) => {
        const rows = document.querySelectorAll(rowSel);
        for (const row of Array.from(rows)) {
          const cell  = row.querySelector('div[columnindex="1"][role="gridcell"]');
          const title = cell?.getAttribute('title') ?? '';
          if (title.trim().length > 3) return true;
        }
        return false;
      },
      { rowSel: locators.resultRow },
      { timeout: timeoutMs },
    );
    return true;
  } catch {
    return false;
  }
}

// ── Scrape all result rows from the current grid ─────────────────────────────
async function scrapeResultRows(page, sourceId) {
  const cols = locators.columns;
  const rawRows = await page.evaluate(
    ({ rowSel, columns }) => {
      const rowEls  = document.querySelectorAll(rowSel);
      const results = [];
      rowEls.forEach((rowEl) => {
        const record = {};
        columns.forEach(({ index, header, child }) => {
          const cell  = rowEl.querySelector(\`div[columnindex="\${index}"][role="gridcell"]\`);
          if (!cell) { record[header] = ''; return; }
          const title = (cell.getAttribute('title') ?? '').trim();
          if (title !== '') { record[header] = title; return; }
          const childEl = cell.querySelector(child);
          record[header] = (childEl?.textContent ?? '').trim();
        });
        results.push(record);
      });
      return results;
    },
    { rowSel: locators.resultRow, columns: cols },
  );
  return rawRows
    .filter((r) => (r['Timesheet ID'] ?? '').trim().length > 3)
    .map((r) => ({ WOID: sourceId, ...r }));
}

// ── Process one tab's chunk of IDs ───────────────────────────────────────────
async function processChunk(page, chunk, tabLabel) {
  const results = [];

  log(\`\${tabLabel} Processing \${chunk.length} IDs\`);

  for (let i = 0; i < chunk.length; i++) {
    const id = chunk[i];
    log(\`\${tabLabel} [\${i + 1}/\${chunk.length}] ID: \${id}\`);

    try {
      // Search
      const searchHost  = page.locator(locators.searchInputHost);
      const searchInner = searchHost.locator(locators.searchInputInner);
      await searchInner.waitFor({ state: 'visible' });
      await searchInner.click();
      await searchInner.click({ clickCount: 3 });
      await searchInner.fill(id);
      await searchInner.press('Enter');

      await page.locator(locators.pageReadyProbe).first().waitFor({ state: 'visible', timeout: 60000 });

      // Worker name
      const workerName = await page
        .locator(locators.workerName)
        .first()
        .textContent()
        .then((t) => t?.trim() ?? '')
        .catch(() => '');

      // Open Time & Expense tab
      await page.locator(locators.tabTimeExpense).click();

      // Set date filters
      await page.locator(locators.fromDateInput).first().waitFor({ state: 'visible' });
      await page.locator(locators.fromDateInput).first().clear();
      await page.locator(locators.fromDateInput).first().type(fromDate);
      await page.keyboard.press('Tab');

      await page.locator(locators.toDateInput).first().waitFor({ state: 'visible' });
      await page.locator(locators.toDateInput).first().clear();
      await page.locator(locators.toDateInput).first().type(toDate);
      await page.keyboard.press('Tab');

      await page.locator(locators.applyFilterBtn).first().waitFor({ state: 'visible' });
      await page.locator(locators.applyFilterBtn).first().click();

      const hasResults = await waitForGridReady(page, 90000);

      if (!hasResults) {
        log(\`\${tabLabel} No results for ID: \${id}\`);
        results.push({ WOID: id, 'Worker Name': workerName, Status: 'No Results' });
      } else {
        const rows = await scrapeResultRows(page, id);
        rows.forEach((r) => { r['Worker Name'] = workerName; });
        log(\`\${tabLabel} \${rows.length} row(s) scraped for ID: \${id}\`);
        results.push(...rows);
      }

      // Navigate home (skip after last ID)
      if (i < chunk.length - 1) {
        await page.locator(locators.homeBtn).click();
        await page.locator(locators.homeReadyProbe).waitFor({ state: 'visible', timeout: 30000 });
        await new Promise((r) => setTimeout(r, 300));
      }

    } catch (err) {
      log(\`\${tabLabel} ERROR on ID "\${id}": \${String(err)}\`);
      results.push({ WOID: id, Status: 'FAILED', reason: String(err) });

      try {
        await page.locator(locators.homeBtn).click();
        await page.locator(locators.homeReadyProbe).waitFor({ state: 'visible', timeout: 20000 });
      } catch { /* best-effort recovery */ }
    }
  }

  log(\`\${tabLabel} Completed — \${results.length} record(s)\`);
  return results;
}

// ── Main worker entry point ───────────────────────────────────────────────────
(async () => {
  log('Started');

  let browser;
  try {
    // Launch Chromium — restore auth state so login is already active
    const launchOpts = { headless: false };
    browser = await chromium.launch(launchOpts);

    const contextOpts = authStorageState
      ? { storageState: authStorageState }
      : {};

    const browserCtx = await browser.newContext(contextOpts);

    // Open two pages concurrently
    const [page1, page2] = await Promise.all([
      browserCtx.newPage(),
      browserCtx.newPage(),
    ]);

    // Navigate both tabs to the base URL so each has an active session
    await Promise.all([
      page1.goto(baseUrl, { waitUntil: 'domcontentloaded' }),
      page2.goto(baseUrl, { waitUntil: 'domcontentloaded' }),
    ]);

    // Wait for both tabs to show the search input (app is ready)
    await Promise.all([
      page1.locator(locators.homeReadyProbe).waitFor({ state: 'visible', timeout: 60000 }),
      page2.locator(locators.homeReadyProbe).waitFor({ state: 'visible', timeout: 60000 }),
    ]);

    log('Both tabs ready — starting parallel chunk processing');

    // Process both chunks concurrently
    const [results1, results2] = await Promise.all([
      processChunk(page1, chunk1, 'Tab1'),
      processChunk(page2, chunk2, 'Tab2'),
    ]);

    const combined = [...results1, ...results2];
    log(\`Completed — \${combined.length} total record(s) collected\`);

    parentPort.postMessage({ type: 'result', data: combined });

  } catch (err) {
    log(\`Worker-level error: \${String(err)}\`);
    parentPort.postMessage({ type: 'result', data: [] });
  } finally {
    if (browser) await browser.close().catch(() => {});
    log('Browser closed');
  }
})();
`;

  // ═══════════════════════════════════════════════════════════════════════════
  // STEP 4 — Snapshot the already-logged-in Walnut browser session and write
  //           it to a temp file so both worker browsers can restore it without
  //           performing any login themselves.
  // ═══════════════════════════════════════════════════════════════════════════
  // Playwright is bundled inside the Walnut Agent at a fixed known path
  const playwrightPath = '/Applications/Walnut Agent.app/Contents/Resources/node_modules/playwright/index.js';
  ctx.log('Capturing session storage state from active browser...');
  const storageState = await ctx.page.context().storageState();
  const authStorageState = path.join(os.tmpdir(), `walnut_auth_${Date.now()}.json`);
  fs.writeFileSync(authStorageState, JSON.stringify(storageState), 'utf-8');
  ctx.log(`Session state saved to: ${authStorageState}`);
  // ctx.testBaseUrl is empty in this runtime — read the URL from the live page instead
  const rawUrl = ctx.page.url() as string;
  const baseUrl: string = rawUrl && rawUrl !== 'about:blank' ? rawUrl : ctx.testBaseUrl;

  // ═══════════════════════════════════════════════════════════════════════════
  // STEP 5 — Launch 2 workers, each receives 2 chunks
  // ═══════════════════════════════════════════════════════════════════════════

  type ResultRecord = Record<string, string>;

  function spawnWorker(
    workerId: number,
    chunk1: string[],
    chunk2: string[],
  ): Promise<ResultRecord[]> {
    return new Promise((resolve) => {
      const worker = new Worker(workerScript, {
        eval: true,
        workerData: {
          workerId,
          baseUrl,
          authStorageState,
          playwrightPath,
          chunk1,
          chunk2,
          fromDate,
          toDate,
          locators: {
            searchInputHost:  LOCATORS.searchInputHost,
            searchInputInner: LOCATORS.searchInputInner,
            tabTimeExpense:   LOCATORS.tabTimeExpense,
            fromDateInput:    LOCATORS.fromDateInput,
            toDateInput:      LOCATORS.toDateInput,
            applyFilterBtn:   LOCATORS.applyFilterBtn,
            resultRow:        LOCATORS.resultRow,
            columns:          LOCATORS.columns.map((c) => ({ ...c })),
            workerName:       LOCATORS.workerName,
            pageReadyProbe:   LOCATORS.pageReadyProbe,
            homeBtn:          LOCATORS.homeBtn,
            homeReadyProbe:   LOCATORS.homeReadyProbe,
          },
        },
      });

      let collected: ResultRecord[] = [];

      worker.on('message', (msg: { type: string; msg?: string; data?: ResultRecord[] }) => {
        if (msg.type === 'log') {
          ctx.log(msg.msg ?? '');
        } else if (msg.type === 'result') {
          collected = msg.data ?? [];
        }
      });

      worker.on('error', (err) => {
        ctx.log(`[Worker${workerId}] Thread error: ${String(err)}`);
      });

      worker.on('exit', () => {
        ctx.log(`Worker${workerId} Completed`);
        resolve(collected);
      });
    });
  }

  ctx.log('Worker1 Started');
  ctx.log('Worker2 Started');
  ctx.log(`Worker1 Tab1 Processing ${chunks[0].length} IDs`);
  ctx.log(`Worker1 Tab2 Processing ${chunks[1].length} IDs`);
  ctx.log(`Worker2 Tab1 Processing ${chunks[2].length} IDs`);
  ctx.log(`Worker2 Tab2 Processing ${chunks[3].length} IDs`);

  // Both workers execute concurrently
  const [worker1Results, worker2Results] = await Promise.all([
    spawnWorker(1, chunks[0], chunks[1]),
    spawnWorker(2, chunks[2], chunks[3]),
  ]);

  // ═══════════════════════════════════════════════════════════════════════════
  // STEP 6 — Merge results from all workers
  // ═══════════════════════════════════════════════════════════════════════════
  ctx.log('Merging Results...');
  const allResults: ResultRecord[] = [...worker1Results, ...worker2Results];
  ctx.log(`  Worker1: ${worker1Results.length} record(s)`);
  ctx.log(`  Worker2: ${worker2Results.length} record(s)`);
  ctx.log(`  Total:   ${allResults.length} record(s)`);

  if (allResults.length === 0) {
    ctx.log('No results collected — skipping workbook write.');
    return;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // STEP 7 — Write ALL results to workbook exactly once
  // ═══════════════════════════════════════════════════════════════════════════
  ctx.log('Writing Workbook...');

  const wb   = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);

  let sheet = wb.getWorksheet(outputSheetName);
  if (!sheet) {
    sheet = wb.addWorksheet(outputSheetName);
    sheet.columns = colOrder.map((h) => ({
      width: Math.min(40, Math.max(14, h.length + 4)),
    }));
    ctx.log(`Created output sheet "${outputSheetName}".`);
  }

  // Write header if missing
  const cell1     = sheet.getRow(1).getCell(1);
  const hasHeader = cell1.value !== null && cell1.value !== undefined && cell1.value !== '';
  if (!hasHeader) {
    const hRow = sheet.getRow(1);
    colOrder.forEach((h, idx) => { hRow.getCell(idx + 1).value = h; });
    styleHeaderRow(hRow, colOrder.length);
    ctx.log(`Written header row to sheet "${outputSheetName}".`);
  }

  // Determine starting data row index (for alternating colour)
  const existingDataRows = Math.max(0, sheet.rowCount - 1);
  const nextRowNum        = sheet.rowCount + 1;

  allResults.forEach((record, idx) => {
    const row = sheet!.getRow(nextRowNum + idx);
    colOrder.forEach((h, colIdx) => {
      row.getCell(colIdx + 1).value = record[h] ?? '';
    });
    styleDataRow(row, colOrder.length, existingDataRows + idx);
    row.commit();
  });

  await wb.xlsx.writeFile(filePath);

  const fileName = filePath.replace(/\\/g, '/').split('/').pop() ?? filePath;
  ctx.log(`Completed`);
  ctx.log(
    `${allResults.length} row(s) written to sheet "${outputSheetName}" in ${fileName}.`,
  );
}
