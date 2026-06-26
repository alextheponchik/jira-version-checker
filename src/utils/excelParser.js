import * as XLSX from 'xlsx-js-style';

const REQUIRED_COLS = ['key', 'fix version/s'];

// Extract key→url mapping from the raw HTML content of Jira-exported .xls files.
// Jira exports HTML with anchors like: href="https://…/browse/SIG-123">SIG-123</a>
function extractLinkMap(arrayBuffer) {
  const text = new TextDecoder('utf-8', { fatal: false }).decode(new Uint8Array(arrayBuffer));
  const map = new Map();
  // Matches: href="URL" ... >KEY</a>  where KEY looks like ALPHA-digits
  const re = /href="([^"]+)"[^>]*>([A-Z][A-Z0-9]*-\d+)<\/a>/gi;
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
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    if (rows.length > 0) {
      const headerRow = rows[0];
      if (headerRow.some((col) => normCol(col) === 'key')) {
        return sheetName;
      }
    }
  }
  return null;
}

export function parseExcelFile(arrayBuffer) {
  const linkMap = extractLinkMap(arrayBuffer);
  const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });

  const sheetName = findDataSheet(workbook);
  if (!sheetName) {
    return { error: 'Не найден лист с колонкой "Key". Убедитесь, что файл содержит правильные данные.' };
  }

  const sheet = workbook.Sheets[sheetName];
  const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });

  if (rawRows.length < 2) {
    return { error: 'Файл не содержит строк данных (только заголовок или пустой).' };
  }

  const headers = rawRows[0].map((h) => h.toString());

  // Validate required columns
  for (const required of REQUIRED_COLS) {
    if (!headers.some((h) => normCol(h) === required)) {
      const displayName = required === 'key' ? 'Key' : 'Fix Version/s';
      return { error: `Не найдена обязательная колонка "${displayName}".` };
    }
  }

  // Map header names to their actual casing in the file
  const colMap = {};
  headers.forEach((h) => {
    const norm = normCol(h);
    if (!colMap[norm]) colMap[norm] = h;
  });

  const keyCol = colMap['key'];
  const fixVersionCol = colMap['fix version/s'];
  const linkedIssuesCol = colMap['linked issues'] || null;

  const data = [];
  for (let i = 1; i < rawRows.length; i++) {
    const rawRow = rawRows[i];
    const rowObj = {};
    headers.forEach((h, idx) => {
      rowObj[h] = rawRow[idx] != null ? rawRow[idx].toString() : '';
    });

    const key = rowObj[keyCol]?.trim() ?? '';
    if (!key) continue;

    const fixVersions = rowObj[fixVersionCol]?.trim() ?? '';
    const linkedIssues = linkedIssuesCol ? (rowObj[linkedIssuesCol]?.trim() ?? '') : '';

    data.push({ key, fixVersions, linkedIssues, url: linkMap.get(key) || null });
  }

  return { data };
}
