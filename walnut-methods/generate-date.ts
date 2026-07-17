import type { WalnutContext } from './walnut';

/** @walnut_method
 * name: custom Generate Date
 * description: Generate date in format ${format} with offset ${offset} days and store in $[date]
 * actionType: custom_generate_date
 * context: shared
 * needsLocator: false
 * category: Data Generation
 */
export async function generateDate(ctx: WalnutContext) {
  // ctx.args[0] = format string  (from ${format})  — see supported tokens below
  // ctx.args[1] = offset in days (from ${offset})  — e.g. "0" = today, "1" = tomorrow, "-1" = yesterday
  // ctx.args[2] = "date"         (from $[date])    — runtime variable name to store output

  const format    = ctx.args[0];
  const offsetRaw = ctx.args[1];
  const outputVar = ctx.args[2];

  if (!format) {
    throw new Error('Format argument is required. Example formats: YYYY-MM-DD, DD/MM/YYYY, MM-DD-YYYY HH:mm:ss');
  }

  const offsetDays = offsetRaw !== undefined && offsetRaw !== '' ? parseInt(offsetRaw, 10) : 0;

  if (isNaN(offsetDays)) {
    throw new Error(`Invalid offset: "${offsetRaw}" must be an integer (e.g. 0, 1, -7)`);
  }

  // ── Build the target date ──────────────────────────────────────────────────
  const now = new Date();
  now.setDate(now.getDate() + offsetDays);

  // ── Named preset aliases ───────────────────────────────────────────────────
  const PRESETS: Record<string, string> = {
    // ISO / standard
    ISO8601:          'YYYY-MM-DDTHH:mm:ss.SSSZ',
    ISO8601_DATE:     'YYYY-MM-DD',
    ISO8601_TIME:     'HH:mm:ss',
    ISO8601_DATETIME: 'YYYY-MM-DDTHH:mm:ss',

    // Regional
    US:               'MM/DD/YYYY',
    US_DATETIME:      'MM/DD/YYYY hh:mm:ss A',
    EU:               'DD/MM/YYYY',
    EU_DATETIME:      'DD/MM/YYYY HH:mm:ss',
    UK:               'DD/MM/YYYY',
    DE:               'DD.MM.YYYY',
    JP:               'YYYY/MM/DD',
    CN:               'YYYY年MM月DD日',

    // Human-readable
    LONG:             'MMMM DD, YYYY',
    LONG_DATETIME:    'MMMM DD, YYYY HH:mm:ss',
    FULL:             'dddd, MMMM DD, YYYY',
    SHORT:            'MMM DD, YYYY',
    SHORT_DATETIME:   'MMM DD, YYYY HH:mm:ss',

    // Time only
    TIME_12:          'hh:mm:ss A',
    TIME_24:          'HH:mm:ss',
    TIME_24_MS:       'HH:mm:ss.SSS',

    // Timestamps
    UNIX:             'X',
    UNIX_MS:          'x',

    // Ordinal / other
    ORDINAL:          'Do MMMM YYYY',
    MONTH_YEAR:       'MMMM YYYY',
    YEAR_MONTH:       'YYYY-MM',
    YYYYMMDD:         'YYYYMMDD',
    DDMMYYYY:         'DDMMYYYY',
    MMDDYYYY:         'MMDDYYYY',
  };

  const resolvedFormat = PRESETS[format.toUpperCase()] ?? format;

  // ── Token helpers ──────────────────────────────────────────────────────────
  const MONTHS_FULL  = ['January','February','March','April','May','June',
                        'July','August','September','October','November','December'];
  const MONTHS_SHORT = ['Jan','Feb','Mar','Apr','May','Jun',
                        'Jul','Aug','Sep','Oct','Nov','Dec'];
  const DAYS_FULL    = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  const DAYS_SHORT   = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];

  const pad  = (n: number, w = 2) => String(n).padStart(w, '0');

  const YYYY = String(now.getFullYear());
  const YY   = YYYY.slice(-2);
  const M_   = now.getMonth() + 1;
  const D_   = now.getDate();
  const H_   = now.getHours();
  const h_   = H_ % 12 === 0 ? 12 : H_ % 12;
  const min_ = now.getMinutes();
  const s_   = now.getSeconds();
  const ms_  = now.getMilliseconds();

  // Timezone offset  (+05:30 style)
  const tzOffset  = -now.getTimezoneOffset();
  const tzSign    = tzOffset >= 0 ? '+' : '-';
  const tzH       = pad(Math.floor(Math.abs(tzOffset) / 60));
  const tzM       = pad(Math.abs(tzOffset) % 60);
  const TZ_COLON  = `${tzSign}${tzH}:${tzM}`;
  const TZ_PLAIN  = `${tzSign}${tzH}${tzM}`;

  // Ordinal suffix  (1st, 2nd, 3rd, 4th …)
  const ordinal = (n: number): string => {
    const s = ['th','st','nd','rd'];
    const v = n % 100;
    return n + (s[(v - 20) % 10] || s[v] || s[0]);
  };

  // Unix timestamps
  const unixSec = String(Math.floor(now.getTime() / 1000));
  const unixMs  = String(now.getTime());

  // ── Token replacement (longest tokens first to avoid partial matches) ──────
  const tokens: [RegExp, string][] = [
    // 4-char tokens
    [/YYYY/g,  YYYY],
    [/MMMM/g,  MONTHS_FULL[M_ - 1]],
    [/dddd/g,  DAYS_FULL[now.getDay()]],
    // 3-char tokens
    [/MMM/g,   MONTHS_SHORT[M_ - 1]],
    [/ddd/g,   DAYS_SHORT[now.getDay()]],
    [/SSS/g,   pad(ms_, 3)],
    // 2-char tokens
    [/YY/g,    YY],
    [/MM/g,    pad(M_)],
    [/DD/g,    pad(D_)],
    [/HH/g,    pad(H_)],
    [/hh/g,    pad(h_)],
    [/mm/g,    pad(min_)],
    [/ss/g,    pad(s_)],
    [/ZZ/g,    TZ_PLAIN],
    // 1-char tokens (after all 2-char variants are resolved)
    [/\bM\b/g, String(M_)],
    [/\bD\b/g, String(D_)],
    [/\bH\b/g, String(H_)],
    [/\bh\b/g, String(h_)],
    [/Do/g,    ordinal(D_)],
    [/\bA\b/g, H_ < 12 ? 'AM' : 'PM'],
    [/\ba\b/g, H_ < 12 ? 'am' : 'pm'],
    [/\bZ\b/g, TZ_COLON],
    [/\bX\b/g, unixSec],
    [/\bx\b/g, unixMs],
  ];

  let result = resolvedFormat;
  for (const [pattern, value] of tokens) {
    result = result.replace(pattern, value);
  }

  ctx.setVariable(outputVar, result);
  ctx.log(`Generated date: "${result}" (format: "${resolvedFormat}", offset: ${offsetDays}d) → stored in $[${outputVar}]`);
}
