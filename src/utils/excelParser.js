import * as XLSX from 'xlsx-js-style';

const REQUIRED_COLS = ['key', 'fix version/s'];

const SD_PREFIXES = ['HELP'];
const JIRA_BASE_URL = 'https://jira.ors-aero.ru/browse/';
const SD_BASE_URL   = 'https://support.ors-aero.ru/browse/';

function fallbackUrl(key) {
  const upper = key.toUpperCase();
  const isSD = SD_PREFIXES.some((p) => upper.startsWith(p + '-'));
  return (isSD ? SD_BASE_URL : JIRA_BASE_URL) + key;
}

// Normalise a raw comma-separated cell value (may have lots of internal whitespace)
function normCsv(raw) {
  if (!raw) return '';
  return raw.split(',').map((v) => v.trim()).filter(Boolean).join(', ');
}

// ── HTML path ────────────────────────────────────────────────────────────────
// Jira HTML exports mark rows with class="issuerow" and cells with class="issuekey",
// class="fixVersions", class="issuelinks". Parse directly via DOMParser.

function isHtmlBuffer(arrayBuffer) {
  const head = new TextDecoder('utf-8', { fatal: false }).decode(new Uint8Array(arrayBuffer, 0, 20));
  return head.trimStart().toLowerCase().startsWith('<');
}

function parseHtmlContent(arrayBuffer) {
  const text = new TextDecoder('utf-8', { fatal: false }).decode(new Uint8Array(arrayBuffer));
  const doc = new DOMParser().parseFromString(text, 'text/html');
  const rows = doc.querySelectorAll('tr.issuerow');

  if (rows.length === 0) {
    return { error: 'В HTML-файле не найдены строки с данными (ожидались элементы <tr class="issuerow">).' };
  }

  const data = [];
  for (const row of rows) {
    const keyCell   = row.querySelector('td.issuekey');
    const fvCell    = row.querySelector('td.fixVersions');
    const liCell    = row.querySelector('td.issuelinks');

    if (!keyCell) continue;

    const anchor = keyCell.querySelector('a');
    const key    = (anchor?.textContent ?? keyCell.textContent).trim();
    if (!key) continue;

    const url          = anchor?.href || fallbackUrl(key);
    const fixVersions  = fvCell  ? normCsv(fvCell.textContent)  : '';
    const linkedIssues = liCell  ? normCsv(liCell.textContent)  : '';

    data.push({ key, url, fixVersions, linkedIssues });
  }

  if (data.length === 0) {
    return { error: 'HTML-файл не содержит строк с ключами задач.' };
  }
  return { data };
}

// ── XLS/XLSX path ────────────────────────────────────────────────────────────
// .xls files from Jira are actually HTML-in-XLS format; SheetJS reads the
// table data but loses hyperlinks, so we also extract hrefs via regex.

function extractLinkMap(arrayBuffer) {
  const text = new TextDecoder('utf-8', { fatal: false }).decode(new Uint8Array(arrayBuffer));
  const map  = new Map();
  const re   = /href="([^"]+)"[^>]*>([A-Z][A-Z0-9]*-\d+)<\/a>/gi;
  let m;
  while ((m = re.exec(text)) !== null) {
    const [, url, key] = m;
    if (!map.has(key)) map.set(key, url);
  }
  return map;
}

function normCol(name) {
  return name ? name.toString().trim().toLowerCase() : '';
}

function findDataSheet(workbook) {
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows  = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    if (rows.length > 0 && rows[0].some((col) => normCol(col) === 'key')) {
      return sheetName;
    }
  }
  return null;
}

function parseXlsContent(arrayBuffer) {
  const linkMap  = extractLinkMap(arrayBuffer);
  const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });

  const sheetName = findDataSheet(workbook);
  if (!sheetName) {
    return { error: 'Не найден лист с колонкой "Key". Убедитесь, что файл содержит правильные данные.' };
  }

  const sheet   = workbook.Sheets[sheetName];
  const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });

  if (rawRows.length < 2) {
    return { error: 'Файл не содержит строк данных (только заголовок или пустой).' };
  }

  const headers = rawRows[0].map((h) => h.toString());

  for (const required of REQUIRED_COLS) {
    if (!headers.some((h) => normCol(h) === required)) {
      const displayName = required === 'key' ? 'Key' : 'Fix Version/s';
      return { error: `Не найдена обязательная колонка "${displayName}".` };
    }
  }

  const colMap = {};
  headers.forEach((h) => { const n = normCol(h); if (!colMap[n]) colMap[n] = h; });

  const keyCol        = colMap['key'];
  const fixVersionCol = colMap['fix version/s'];
  const linkedIssuesCol = colMap['linked issues'] || null;

  const data = [];
  for (let i = 1; i < rawRows.length; i++) {
    const rawRow = rawRows[i];
    const rowObj = {};
    headers.forEach((h, idx) => { rowObj[h] = rawRow[idx] != null ? rawRow[idx].toString() : ''; });

    const key = rowObj[keyCol]?.trim() ?? '';
    if (!key) continue;

    const fixVersions  = rowObj[fixVersionCol]?.trim() ?? '';
    const linkedIssues = linkedIssuesCol ? (rowObj[linkedIssuesCol]?.trim() ?? '') : '';

    data.push({ key, fixVersions, linkedIssues, url: linkMap.get(key) || fallbackUrl(key) });
  }

  return { data };
}

// ── Public API ────────────────────────────────────────────────────────────────
export function parseExcelFile(arrayBuffer) {
  return isHtmlBuffer(arrayBuffer)
    ? parseHtmlContent(arrayBuffer)
    : parseXlsContent(arrayBuffer);
}
