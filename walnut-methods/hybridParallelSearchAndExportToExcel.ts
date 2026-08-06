import type { WalnutWebContext } from './walnut';
import ExcelJS from 'exceljs';
import { Worker } from 'worker_threads';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

/** @walnut_method
 * name: Hybrid Parallel Search IDs from Excel and Write Results
 * description: Open workbook ${filePath} sheet ${inputSheetName} column ${idColumnName} search each ID using ${workerCount} browser workers with ${tabsPerWorker} tabs each from date ${fromDate} to date ${toDate} and write results into sheet ${outputSheetName}
 * actionType: custom_hybrid_parallel_search_and_export_to_excel
 * context: web
 * needsLocator: false
 * category: Data Processing
 */
export async function hybridParallelSearchAndExportToExcel(ctx: WalnutWebContext) {
  const filePath        = ctx.args[0];
  const inputSheetName  = ctx.args[1];
  const idColumnName    = ctx.args[2];
  const workerCount     = Math.max(1, parseInt(ctx.args[3], 10) || 2);
  const tabsPerWorker   = Math.max(1, parseInt(ctx.args[4], 10) || 2);
  const fromDate        = ctx.args[5];
  const toDate          = ctx.args[6];
  const outputSheetName = (ctx.args[7] ?? '').trim() || `Output_${new Date().toISOString().slice(0,19).replace(/[:T]/g,'-')}`;

  const totalTabs = workerCount * tabsPerWorker;

  ctx.log(`Configuration: ${workerCount} worker(s) × ${tabsPerWorker} tab(s) = ${totalTabs} total tabs`);

  // ── LOCATORS ─────────────────────────────────────────────────────────────────
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
    workerNameSel:  'h1[data-help-id="TITLE_270"] span.titlePrimary',
    pageReadyProbe: '//a[@id="tab_timeAndExpense"]',
    homeBtn:        '//li[@id="homeMenuTitle"]//a[@title="Home"]',
    homeReadyProbe: 'ui5-input#searchFieldInput',
  } as const;

  type ResultRecord = Record<string, string>;

  // ── Helpers ───────────────────────────────────────────────────────────────────
  function cellText(value: ExcelJS.CellValue): string {
    if (value === null || value === undefined) return '';
    if (typeof value === 'object') {
      if ('result' in value) return String((value as ExcelJS.CellFormulaValue).result ?? '');
      if ('richText' in value) return (value as ExcelJS.CellRichTextValue).richText.map((r) => r.text).join('');
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
      cell.border    = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    }
    row.commit();
  }

  function styleDataRow(row: ExcelJS.Row, colCount: number, dataIndex: number): void {
    const bgColor = dataIndex % 2 === 0 ? 'FFFFFFFF' : 'FFF2F2F2';
    for (let c = 1; c <= colCount; c++) {
      const cell = row.getCell(c);
      cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
      cell.border    = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      cell.alignment = { vertical: 'middle', wrapText: true };
    }
    row.commit();
  }

  function splitIntoChunks<T>(arr: T[], chunkCount: number): T[][] {
    const chunks: T[][] = Array.from({ length: chunkCount }, () => []);
    arr.forEach((item, i) => chunks[i % chunkCount].push(item));
    return chunks;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // STEP 1 — Read workbook ONCE
  // ═══════════════════════════════════════════════════════════════════════════
  ctx.log('Reading Workbook...');
  const inputWorkbook = new ExcelJS.Workbook();
  await inputWorkbook.xlsx.readFile(filePath);

  const inSheet = inputWorkbook.getWorksheet(inputSheetName);
  if (!inSheet) {
    const available = inputWorkbook.worksheets.map((ws: ExcelJS.Worksheet) => ws.name).join(', ');
    throw new Error(`Input sheet "${inputSheetName}" not found. Available: ${available}`);
  }

  const headerRow = inSheet.getRow(1);
  let idColIndex = -1;
  headerRow.eachCell({ includeEmpty: false }, (cell: ExcelJS.Cell, colIdx: number) => {
    if (String(cell.value ?? '').trim().toLowerCase() === idColumnName.trim().toLowerCase())
      idColIndex = colIdx;
  });
  if (idColIndex === -1) {
    const found: string[] = [];
    headerRow.eachCell({ includeEmpty: false }, (cell: ExcelJS.Cell) => found.push(String(cell.value ?? '')));
    throw new Error(`Column "${idColumnName}" not found. Found: ${found.join(', ')}`);
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

  const defaultColOrder: string[] = ['WOID', 'Worker Name', ...LOCATORS.columns.map((c) => c.header)];
  let colOrder: string[] = defaultColOrder;
  {
    const existingSheet = inputWorkbook.getWorksheet(outputSheetName);
    if (existingSheet) {
      const hdrs: string[] = [];
      existingSheet.getRow(1).eachCell({ includeEmpty: false }, (cell: ExcelJS.Cell) => hdrs.push(String(cell.value ?? '')));
      if (hdrs.length > 0) colOrder = hdrs;
    }
  }
  // @ts-ignore
  inputWorkbook._worksheets = [];

  // ═══════════════════════════════════════════════════════════════════════════
  // STEP 2 — Split IDs into totalTabs chunks
  // ═══════════════════════════════════════════════════════════════════════════
  ctx.log(`Splitting into ${totalTabs} Chunks (${workerCount} worker(s) × ${tabsPerWorker} tab(s))...`);
  const allChunks = splitIntoChunks(ids, totalTabs);
  allChunks.forEach((chunk, i) => ctx.log(`  Chunk${i + 1}: ${chunk.length} IDs`));

  // ═══════════════════════════════════════════════════════════════════════════
  // STEP 3 — Capture session state from the live Walnut browser
  // ═══════════════════════════════════════════════════════════════════════════
  ctx.log('Capturing session storage state from active browser...');
  const storageState = await ctx.page.context().storageState();
  const authStorageStatePath = path.join(os.tmpdir(), `walnut_auth_${Date.now()}.json`);
  fs.writeFileSync(authStorageStatePath, JSON.stringify(storageState), 'utf-8');
  ctx.log(`Session state saved to: ${authStorageStatePath}`);

  const rawUrl = ctx.page.url() as string;
  const baseUrl: string = rawUrl && rawUrl !== 'about:blank' ? rawUrl : ctx.testBaseUrl;
  ctx.log(`Base URL: ${baseUrl}`);

  const playwrightPath = '/Applications/Walnut Agent.app/Contents/Resources/node_modules/playwright/index.js';

  // ═══════════════════════════════════════════════════════════════════════════
  // STEP 4 — Worker thread script (used by extra worker processes)
  //          Accepts N chunks (tabsPerWorker) and processes them concurrently
  // ═══════════════════════════════════════════════════════════════════════════
  const workerScript = /* javascript */ `
const { workerData, parentPort } = require('worker_threads');
const { chromium } = require(workerData.playwrightPath);

const { workerId, baseUrl, authStorageState, chunks, fromDate, toDate, locators } = workerData;

function log(msg) {
  parentPort.postMessage({ type: 'log', msg: \`[Worker\${workerId}] \${msg}\` });
}

async function waitForGridReady(page, timeoutMs = 90000) {
  try {
    await page.waitForFunction(
      ({ rowSel }) => {
        const rows = document.querySelectorAll(rowSel);
        for (const row of Array.from(rows)) {
          const cell = row.querySelector('div[columnindex="1"][role="gridcell"]');
          if ((cell?.getAttribute('title') ?? '').trim().length > 3) return true;
        }
        return false;
      },
      { rowSel: locators.resultRow },
      { timeout: timeoutMs },
    );
    return true;
  } catch { return false; }
}

async function scrapeResultRows(page, sourceId) {
  const rawRows = await page.evaluate(
    ({ rowSel, columns }) => {
      const rowEls = document.querySelectorAll(rowSel);
      const results = [];
      rowEls.forEach((rowEl) => {
        const record = {};
        columns.forEach(({ index, header, child }) => {
          const cell = rowEl.querySelector(\`div[columnindex="\${index}"][role="gridcell"]\`);
          if (!cell) { record[header] = ''; return; }
          const title = (cell.getAttribute('title') ?? '').trim();
          record[header] = title !== '' ? title : (cell.querySelector(child)?.textContent ?? '').trim();
        });
        results.push(record);
      });
      return results;
    },
    { rowSel: locators.resultRow, columns: locators.columns },
  );
  return rawRows
    .filter((r) => (r['Timesheet ID'] ?? '').trim().length > 3)
    .map((r) => ({ WOID: sourceId, ...r }));
}

async function processChunk(page, chunk, tabLabel) {
  log(\`\${tabLabel} Processing \${chunk.length} IDs\`);
  const results = [];
  for (let i = 0; i < chunk.length; i++) {
    const id = chunk[i];
    log(\`\${tabLabel} [\${i + 1}/\${chunk.length}] ID: \${id}\`);
    try {
      const searchInner = page.locator(locators.searchInputHost).locator(locators.searchInputInner);
      await searchInner.waitFor({ state: 'visible' });
      await searchInner.click();
      await searchInner.click({ clickCount: 3 });
      await searchInner.fill(id);
      await searchInner.press('Enter');
      await page.locator(locators.pageReadyProbe).first().waitFor({ state: 'visible', timeout: 60000 });

      const workerName = await page.locator(locators.workerNameSel).first().textContent().then(t => t?.trim() ?? '').catch(() => '');

      await page.locator(locators.tabTimeExpense).click();
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
        log(\`\${tabLabel} No results for \${id}\`);
        results.push({ WOID: id, 'Worker Name': workerName, Status: 'No Results' });
      } else {
        const rows = await scrapeResultRows(page, id);
        rows.forEach(r => { r['Worker Name'] = workerName; });
        log(\`\${tabLabel} \${rows.length} row(s) scraped for \${id}\`);
        results.push(...rows);
      }

      if (i < chunk.length - 1) {
        await page.locator(locators.homeBtn).click();
        await page.locator(locators.homeReadyProbe).waitFor({ state: 'visible', timeout: 30000 });
        await new Promise(r => setTimeout(r, 300));
      }
    } catch (err) {
      log(\`\${tabLabel} ERROR on "\${id}": \${String(err)}\`);
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

(async () => {
  log('Started');
  let browser;
  try {
    browser = await chromium.launch({ headless: false });
    const browserCtx = await browser.newContext({ storageState: authStorageState });

    // Open tabsPerWorker pages concurrently
    const pages = await Promise.all(chunks.map(() => browserCtx.newPage()));

    // Navigate all tabs to baseUrl concurrently
    await Promise.all(pages.map(p => p.goto(baseUrl, { waitUntil: 'domcontentloaded' })));

    // Wait for all tabs to be ready
    await Promise.all(pages.map(p =>
      p.locator(locators.homeReadyProbe).waitFor({ state: 'visible', timeout: 60000 })
    ));

    log('All tabs ready — starting parallel processing');

    // Process all chunks concurrently
    const tabResults = await Promise.all(
      chunks.map((chunk, i) => processChunk(pages[i], chunk, \`Tab\${i + 1}\`))
    );

    const combined = tabResults.flat();
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
  // STEP 5 — Main-thread tab processor (reuses the Walnut browser directly)
  //          Worker 0 = the already-logged-in Walnut browser
  // ═══════════════════════════════════════════════════════════════════════════
  async function waitForGridReadyMain(page: import('playwright').Page, timeoutMs = 90000): Promise<boolean> {
    try {
      await page.waitForFunction(
        ({ rowSel }: { rowSel: string }) => {
          const rows = document.querySelectorAll(rowSel);
          for (const row of Array.from(rows)) {
            const cell = row.querySelector('div[columnindex="1"][role="gridcell"]');
            if ((cell?.getAttribute('title') ?? '').trim().length > 3) return true;
          }
          return false;
        },
        { rowSel: LOCATORS.resultRow },
        { timeout: timeoutMs },
      );
      return true;
    } catch { return false; }
  }

  async function scrapeResultRowsMain(page: import('playwright').Page, sourceId: string): Promise<ResultRecord[]> {
    const rawRows = await page.evaluate(
      ({ rowSel, columns }: { rowSel: string; columns: typeof LOCATORS.columns }) => {
        const rowEls = document.querySelectorAll(rowSel);
        const results: ResultRecord[] = [];
        rowEls.forEach((rowEl) => {
          const record: ResultRecord = {};
          columns.forEach(({ index, header, child }) => {
            const cell = rowEl.querySelector(`div[columnindex="${index}"][role="gridcell"]`);
            if (!cell) { record[header] = ''; return; }
            const title = (cell.getAttribute('title') ?? '').trim();
            record[header] = title !== '' ? title : ((cell.querySelector(child) as HTMLElement)?.textContent ?? '').trim();
          });
          results.push(record);
        });
        return results;
      },
      { rowSel: LOCATORS.resultRow, columns: LOCATORS.columns as any },
    );
    return (rawRows as ResultRecord[])
      .filter((r) => (r['Timesheet ID'] ?? '').trim().length > 3)
      .map((r) => ({ WOID: sourceId, ...r }));
  }

  async function processChunkMain(
    page: import('playwright').Page,
    chunk: string[],
    tabLabel: string,
  ): Promise<ResultRecord[]> {
    ctx.log(`${tabLabel} Processing ${chunk.length} IDs`);
    const results: ResultRecord[] = [];

    for (let i = 0; i < chunk.length; i++) {
      const id = chunk[i];
      ctx.log(`${tabLabel} [${i + 1}/${chunk.length}] ID: ${id}`);
      try {
        const searchHost  = page.locator(LOCATORS.searchInputHost);
        const searchInner = searchHost.locator(LOCATORS.searchInputInner);
        await searchInner.waitFor({ state: 'visible' });
        await searchInner.click();
        await searchInner.click({ clickCount: 3 });
        await searchInner.fill(id);
        await searchInner.press('Enter');
        await page.locator(LOCATORS.pageReadyProbe).first().waitFor({ state: 'visible', timeout: 60000 });

        const workerName = await page.locator(LOCATORS.workerNameSel).first().textContent()
          .then((t) => t?.trim() ?? '').catch(() => '');

        await page.locator(LOCATORS.tabTimeExpense).click();
        await page.locator(LOCATORS.fromDateInput).first().waitFor({ state: 'visible' });
        await page.locator(LOCATORS.fromDateInput).first().clear();
        await page.locator(LOCATORS.fromDateInput).first().type(fromDate);
        await page.keyboard.press('Tab');
        await page.locator(LOCATORS.toDateInput).first().waitFor({ state: 'visible' });
        await page.locator(LOCATORS.toDateInput).first().clear();
        await page.locator(LOCATORS.toDateInput).first().type(toDate);
        await page.keyboard.press('Tab');
        await page.locator(LOCATORS.applyFilterBtn).first().waitFor({ state: 'visible' });
        await page.locator(LOCATORS.applyFilterBtn).first().click();

        const hasResults = await waitForGridReadyMain(page, 90000);
        if (!hasResults) {
          ctx.log(`${tabLabel} No results for ${id}`);
          results.push({ WOID: id, 'Worker Name': workerName, Status: 'No Results' });
        } else {
          const rows = await scrapeResultRowsMain(page, id);
          rows.forEach((r) => { r['Worker Name'] = workerName; });
          ctx.log(`${tabLabel} ${rows.length} row(s) scraped for ${id}`);
          results.push(...rows);
        }

        if (i < chunk.length - 1) {
          await page.locator(LOCATORS.homeBtn).click();
          await page.locator(LOCATORS.homeReadyProbe).waitFor({ state: 'visible', timeout: 30000 });
          await new Promise((r) => setTimeout(r, 300));
        }
      } catch (err) {
        ctx.log(`${tabLabel} ERROR on "${id}": ${String(err)}`);
        results.push({ WOID: id, Status: 'FAILED', reason: String(err) });
        try {
          await page.locator(LOCATORS.homeBtn).click();
          await page.locator(LOCATORS.homeReadyProbe).waitFor({ state: 'visible', timeout: 20000 });
        } catch { /* best-effort */ }
      }
    }
    ctx.log(`${tabLabel} Completed — ${results.length} record(s)`);
    return results;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // STEP 6 — Orchestrate
  //   Worker 0  = Walnut main browser  → tabsPerWorker tabs on main thread
  //   Worker 1+ = spawned Node threads → tabsPerWorker tabs each
  // ═══════════════════════════════════════════════════════════════════════════

  // --- Worker 0: open extra tabs in the EXISTING Walnut browser context ------
  function runMainBrowserWorker(workerChunks: string[][]): Promise<ResultRecord[]> {
    return new Promise(async (resolve) => {
      ctx.log('Worker0 (Main Browser) Started');
      try {
        const browserCtx = ctx.page.context();

        // Open tabsPerWorker pages; first tab reuses ctx.page, rest are new
        const pages: import('playwright').Page[] = [ctx.page];
        for (let t = 1; t < workerChunks.length; t++) {
          const p = await browserCtx.newPage();
          await p.goto(baseUrl, { waitUntil: 'domcontentloaded' });
          await p.locator(LOCATORS.homeReadyProbe).waitFor({ state: 'visible', timeout: 60000 });
          pages.push(p);
        }

        ctx.log(`Worker0 All ${pages.length} tab(s) ready`);

        const tabResults = await Promise.all(
          workerChunks.map((chunk, i) =>
            processChunkMain(pages[i], chunk, `Worker0-Tab${i + 1}`),
          ),
        );

        // Close extra pages (not ctx.page)
        for (let t = 1; t < pages.length; t++) {
          await pages[t].close().catch(() => {});
        }

        const combined = tabResults.flat();
        ctx.log(`Worker0 Completed — ${combined.length} record(s)`);
        resolve(combined);
      } catch (err) {
        ctx.log(`Worker0 error: ${String(err)}`);
        resolve([]);
      }
    });
  }

  // --- Worker 1+: spawn Node worker_threads --------------------------------
  function spawnWorker(workerId: number, workerChunks: string[][]): Promise<ResultRecord[]> {
    return new Promise((resolve) => {
      const worker = new Worker(workerScript, {
        eval: true,
        workerData: {
          workerId,
          baseUrl,
          authStorageState: authStorageStatePath,
          playwrightPath,
          chunks: workerChunks,
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
            workerNameSel:    LOCATORS.workerNameSel,
            pageReadyProbe:   LOCATORS.pageReadyProbe,
            homeBtn:          LOCATORS.homeBtn,
            homeReadyProbe:   LOCATORS.homeReadyProbe,
          },
        },
      });

      let collected: ResultRecord[] = [];

      worker.on('message', (msg: { type: string; msg?: string; data?: ResultRecord[] }) => {
        if (msg.type === 'log') ctx.log(msg.msg ?? '');
        else if (msg.type === 'result') collected = msg.data ?? [];
      });

      worker.on('error', (err) => ctx.log(`[Worker${workerId}] Thread error: ${String(err)}`));

      worker.on('exit', () => {
        ctx.log(`Worker${workerId} Completed`);
        resolve(collected);
      });
    });
  }

  // --- Distribute chunks across workers ------------------------------------
  // allChunks = [c0,c1,c2,...,c(totalTabs-1)]
  // Worker 0 gets indices 0..(tabsPerWorker-1)
  // Worker 1 gets indices tabsPerWorker..(2*tabsPerWorker-1)  etc.
  const workerChunkGroups: string[][][] = [];
  for (let w = 0; w < workerCount; w++) {
    workerChunkGroups.push(allChunks.slice(w * tabsPerWorker, (w + 1) * tabsPerWorker));
  }

  ctx.log(`Launching ${workerCount} worker(s)...`);
  workerChunkGroups.forEach((group, w) => {
    const label = w === 0 ? 'Worker0 (Main Browser)' : `Worker${w}`;
    group.forEach((chunk, t) => ctx.log(`${label} Tab${t + 1} Processing ${chunk.length} IDs`));
  });

  // Worker 0 runs on the main thread; workers 1+ are spawned threads
  const allPromises: Promise<ResultRecord[]>[] = [
    runMainBrowserWorker(workerChunkGroups[0]),
    ...workerChunkGroups.slice(1).map((group, idx) => spawnWorker(idx + 1, group)),
  ];

  const workerResults = await Promise.all(allPromises);

  // ═══════════════════════════════════════════════════════════════════════════
  // STEP 7 — Merge
  // ═══════════════════════════════════════════════════════════════════════════
  ctx.log('Merging Results...');
  const allResults: ResultRecord[] = workerResults.flat();
  workerResults.forEach((r, i) => ctx.log(`  Worker${i}: ${r.length} record(s)`));
  ctx.log(`  Total: ${allResults.length} record(s)`);

  if (allResults.length === 0) {
    ctx.log('No results collected — skipping workbook write.');
    return;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // STEP 8 — Write workbook exactly once
  // ═══════════════════════════════════════════════════════════════════════════
  ctx.log('Writing Workbook...');
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);

  let sheet = wb.getWorksheet(outputSheetName);
  if (!sheet) {
    sheet = wb.addWorksheet(outputSheetName);
    sheet.columns = colOrder.map((h) => ({ width: Math.min(40, Math.max(14, h.length + 4)) }));
    ctx.log(`Created output sheet "${outputSheetName}".`);
  }

  const cell1 = sheet.getRow(1).getCell(1);
  const hasHeader = cell1.value !== null && cell1.value !== undefined && cell1.value !== '';
  if (!hasHeader) {
    const hRow = sheet.getRow(1);
    colOrder.forEach((h, idx) => { hRow.getCell(idx + 1).value = h; });
    styleHeaderRow(hRow, colOrder.length);
    ctx.log(`Written header row to sheet "${outputSheetName}".`);
  }

  const existingDataRows = Math.max(0, sheet.rowCount - 1);
  const nextRowNum = sheet.rowCount + 1;

  allResults.forEach((record, idx) => {
    const row = sheet!.getRow(nextRowNum + idx);
    colOrder.forEach((h, colIdx) => { row.getCell(colIdx + 1).value = record[h] ?? ''; });
    styleDataRow(row, colOrder.length, existingDataRows + idx);
    row.commit();
  });

  await wb.xlsx.writeFile(filePath);

  const fileName = filePath.replace(/\\/g, '/').split('/').pop() ?? filePath;
  ctx.log(`Completed`);
  ctx.log(`${allResults.length} row(s) written to sheet "${outputSheetName}" in ${fileName}.`);
}
