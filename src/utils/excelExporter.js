import * as XLSX from 'xlsx-js-style';

const GREEN = { fgColor: { rgb: 'C3E6CB' } };
const RED   = { fgColor: { rgb: 'F5C6CB' } };
const YELLOW = { fgColor: { rgb: 'FFEEBA' } };

const HEADER_FILL = { fgColor: { rgb: 'D9D9D9' } };
const HEADER_FONT = { bold: true };

function makeCell(value, fill) {
  return {
    v: value ?? '',
    t: 's',
    s: {
      fill: { patternType: 'solid', ...fill },
      font: {},
      alignment: { wrapText: true },
    },
  };
}

function makeHeaderCell(value) {
  return {
    v: value,
    t: 's',
    s: {
      fill: { patternType: 'solid', ...HEADER_FILL },
      font: HEADER_FONT,
    },
  };
}

function arrToSheet(rows) {
  const ws = {};
  let maxCol = 0;

  rows.forEach((row, r) => {
    row.forEach((cell, c) => {
      const addr = XLSX.utils.encode_cell({ r, c });
      ws[addr] = cell;
      if (c > maxCol) maxCol = c;
    });
  });

  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows.length - 1, c: maxCol } });
  ws['!cols'] = Array.from({ length: maxCol + 1 }, () => ({ wch: 20 }));
  return ws;
}

export function exportToExcel(matches, unlinked) {
  // Sheet 1: Matches
  const matchHeaders = [
    makeHeaderCell('Задача Jira'),
    makeHeaderCell('Задача Service Desk'),
    makeHeaderCell('Fix Version/s (Jira)'),
    makeHeaderCell('Fix Version/s (Service Desk)'),
    makeHeaderCell('Статус'),
  ];

  const matchRows = [matchHeaders];
  for (const m of matches) {
    const fill = m.matched ? GREEN : RED;
    matchRows.push([
      makeCell(m.jiraKey, fill),
      makeCell(m.helpKey, fill),
      makeCell(m.jiraFixVersions, fill),
      makeCell(m.helpFixVersions, fill),
      makeCell(m.status, fill),
    ]);
  }

  // Sheet 2: Unlinked
  const unlinkedHeaders = [
    makeHeaderCell('Ключ задачи'),
    makeHeaderCell('Система'),
    makeHeaderCell('Fix Version/s'),
    makeHeaderCell('Linked Issues'),
  ];

  const unlinkedRows = [unlinkedHeaders];
  for (const u of unlinked) {
    unlinkedRows.push([
      makeCell(u.key, YELLOW),
      makeCell(u.system, YELLOW),
      makeCell(u.fixVersions, YELLOW),
      makeCell(u.linkedIssues, YELLOW),
    ]);
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, arrToSheet(matchRows), 'Сопоставления');
  XLSX.utils.book_append_sheet(wb, arrToSheet(unlinkedRows), 'Без связей');

  XLSX.writeFile(wb, 'jira-version-check.xlsx');
}
