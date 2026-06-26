import { useState, useCallback, useRef, useEffect } from 'react';
import { parseExcelFile } from './utils/excelParser';
import { processData } from './utils/dataProcessor';
import { exportToExcel } from './utils/excelExporter';
import './App.css';

function readFileAsBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(e.target.result);
    reader.onerror = () => reject(new Error('Ошибка чтения файла'));
    reader.readAsArrayBuffer(file);
  });
}

function FileInput({ label, onChange, fileName, error }) {
  return (
    <div className="file-input-group">
      <label className="file-label">{label}</label>
      <input
        type="file"
        accept=".xlsx,.xls,.html,.htm"
        onChange={(e) => onChange(e.target.files[0] || null)}
        className="file-input"
      />
      {fileName && <span className="file-name">{fileName}</span>}
      {error && <span className="error-msg">{error}</span>}
    </div>
  );
}

function Badge({ count, color }) {
  return <span className={`badge badge-${color}`}>{count}</span>;
}

function KeyLink({ keyText, url }) {
  if (url) {
    return (
      <a href={url} target="_blank" rel="noopener noreferrer" className="key-link">
        {keyText}
      </a>
    );
  }
  return <span>{keyText}</span>;
}

// ── Column filter ────────────────────────────────────────────────────────────

function FilterHeader({ label, allValues, selected, onToggle, onSelectAll }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);

  useEffect(() => {
    const handler = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  const isActive = selected !== null;
  const allChecked = selected === null;

  return (
    <th ref={ref} className={`th-filterable${isActive ? ' th-active' : ''}`} onClick={() => setOpen((o) => !o)}>
      <span className="th-label">{label}</span>
      <span className="th-arrow">{isActive ? '▼' : '⌄'}</span>
      {open && (
        <div className="filter-dropdown" onClick={(e) => e.stopPropagation()}>
          <label className="filter-option filter-all">
            <input
              type="checkbox"
              checked={allChecked}
              onChange={() => { onSelectAll(); setOpen(false); }}
            />
            <span>Все</span>
          </label>
          <div className="filter-divider" />
          {allValues.map((v) => {
            const checked = selected === null || selected.has(v);
            return (
              <label key={v} className="filter-option">
                <input type="checkbox" checked={checked} onChange={() => onToggle(v)} />
                <span>{v || '—'}</span>
              </label>
            );
          })}
        </div>
      )}
    </th>
  );
}

// selected: null = show all; Set<string> = show only these values
function useColumnFilter(allValues) {
  const [selected, setSelected] = useState(null);

  const toggle = useCallback((value) => {
    setSelected((prev) => {
      const base = prev ?? new Set(allValues);
      const next = new Set(base);
      if (next.has(value)) next.delete(value); else next.add(value);
      return next.size === allValues.length ? null : next;
    });
  }, [allValues]);

  const selectAll = useCallback(() => setSelected(null), []);
  const passes = useCallback((value) => selected === null || selected.has(value), [selected]);

  return { selected, toggle, selectAll, passes };
}

// ─────────────────────────────────────────────────────────────────────────────

function MatchTable({ matches }) {
  const [matchFilter, setMatchFilter] = useState('all');

  const allJiraTypes    = [...new Set(matches.map((m) => m.jiraIssueType || ''))].sort();
  const allJiraStatuses = [...new Set(matches.map((m) => m.jiraStatus    || ''))].sort();
  const allHelpTypes    = [...new Set(matches.map((m) => m.helpIssueType || ''))].sort();
  const allHelpStatuses = [...new Set(matches.map((m) => m.helpStatus    || ''))].sort();

  const jiraTypeFilter   = useColumnFilter(allJiraTypes);
  const jiraStatusFilter = useColumnFilter(allJiraStatuses);
  const helpTypeFilter   = useColumnFilter(allHelpTypes);
  const helpStatusFilter = useColumnFilter(allHelpStatuses);

  const filtered = matches.filter((m) => {
    if (matchFilter === 'match'    && !m.matched) return false;
    if (matchFilter === 'mismatch' &&  m.matched) return false;
    if (!jiraTypeFilter.passes(m.jiraIssueType || ''))   return false;
    if (!jiraStatusFilter.passes(m.jiraStatus  || ''))   return false;
    if (!helpTypeFilter.passes(m.helpIssueType || ''))   return false;
    if (!helpStatusFilter.passes(m.helpStatus  || ''))   return false;
    return true;
  });

  const matchCount    = matches.filter((m) => m.matched).length;
  const mismatchCount = matches.filter((m) => !m.matched).length;

  return (
    <section className="section">
      <div className="section-header">
        <h2>Сравнение версий релиза в Jira и SD</h2>
        <div className="section-stats">
          <Badge count={matchCount} color="green" /> совпадают&nbsp;&nbsp;
          <Badge count={mismatchCount} color="red" /> не совпадают
          {filtered.length !== matches.length && (
            <span className="filter-hint">&nbsp;(показано {filtered.length} из {matches.length})</span>
          )}
        </div>
        <div className="filter-buttons">
          <button className={matchFilter === 'all'      ? 'active' : ''} onClick={() => setMatchFilter('all')}>Все</button>
          <button className={matchFilter === 'mismatch' ? 'active' : ''} onClick={() => setMatchFilter('mismatch')}>Не совпадают</button>
          <button className={matchFilter === 'match'    ? 'active' : ''} onClick={() => setMatchFilter('match')}>Совпадают</button>
        </div>
      </div>
      <div className="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Задача Jira</th>
              <FilterHeader label="Тип (Jira)"    allValues={allJiraTypes}    selected={jiraTypeFilter.selected}   onToggle={jiraTypeFilter.toggle}   onSelectAll={jiraTypeFilter.selectAll} />
              <FilterHeader label="Статус (Jira)" allValues={allJiraStatuses} selected={jiraStatusFilter.selected} onToggle={jiraStatusFilter.toggle} onSelectAll={jiraStatusFilter.selectAll} />
              <th>Задача Service Desk</th>
              <FilterHeader label="Тип (SD)"      allValues={allHelpTypes}    selected={helpTypeFilter.selected}   onToggle={helpTypeFilter.toggle}   onSelectAll={helpTypeFilter.selectAll} />
              <FilterHeader label="Статус (SD)"   allValues={allHelpStatuses} selected={helpStatusFilter.selected} onToggle={helpStatusFilter.toggle} onSelectAll={helpStatusFilter.selectAll} />
              <th>Fix Version/s (Jira)</th>
              <th>Fix Version/s (SD)</th>
              <th>Релиз</th>
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 && (
              <tr><td colSpan={9} className="empty-row">Нет данных</td></tr>
            )}
            {filtered.map((m, i) => (
              <tr key={i} className={m.matched ? 'row-green' : 'row-red'}>
                <td className="key-cell"><KeyLink keyText={m.jiraKey} url={m.jiraUrl} /></td>
                <td className="type-cell">{m.jiraIssueType || '—'}</td>
                <td><span className="issue-status">{m.jiraStatus || '—'}</span></td>
                <td className="key-cell"><KeyLink keyText={m.helpKey} url={m.helpUrl} /></td>
                <td className="type-cell">{m.helpIssueType || '—'}</td>
                <td><span className="issue-status">{m.helpStatus || '—'}</span></td>
                <td>{m.jiraFixVersions || '—'}</td>
                <td>{m.helpFixVersions || '—'}</td>
                <td>
                  <span className={`status-badge ${m.matched ? 'status-green' : 'status-red'}`}>
                    {m.status}
                  </span>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
}

function UnlinkedTable({ unlinked }) {
  const allTypes    = [...new Set(unlinked.map((u) => u.issueType || ''))].sort();
  const allStatuses = [...new Set(unlinked.map((u) => u.status    || ''))].sort();

  const typeFilter   = useColumnFilter(allTypes);
  const statusFilter = useColumnFilter(allStatuses);

  const filtered = unlinked.filter(
    (u) => typeFilter.passes(u.issueType || '') && statusFilter.passes(u.status || ''),
  );

  const sdCount = unlinked.length;

  return (
    <section className="section">
      <div className="section-header">
        <h2>Задачи SD без связи с производственной в Jira</h2>
        <div className="section-stats">
          <Badge count={filtered.length} color="yellow" />
          {filtered.length !== sdCount && <span className="filter-hint"> из {sdCount}</span>}
        </div>
      </div>
      <div className="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Ключ задачи</th>
              <th>Система</th>
              <FilterHeader
                label="Тип"
                allValues={allTypes}
                selected={typeFilter.selected}
                onToggle={typeFilter.toggle}
                onSelectAll={typeFilter.selectAll}
              />
              <FilterHeader
                label="Статус"
                allValues={allStatuses}
                selected={statusFilter.selected}
                onToggle={statusFilter.toggle}
                onSelectAll={statusFilter.selectAll}
              />
              <th>Fix Version/s</th>
              <th>Linked Issues</th>
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 && (
              <tr><td colSpan={6} className="empty-row">Нет задач, соответствующих фильтрам</td></tr>
            )}
            {filtered.map((u, i) => (
              <tr key={i} className="row-yellow">
                <td className="key-cell"><KeyLink keyText={u.key} url={u.url} /></td>
                <td>{u.system}</td>
                <td className="type-cell">{u.issueType || '—'}</td>
                <td><span className="issue-status">{u.status || '—'}</span></td>
                <td>{u.fixVersions || '—'}</td>
                <td className="linked-cell">{u.linkedIssues || '—'}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
}

export default function App() {
  const [jiraFile, setJiraFile] = useState(null);
  const [sdFile, setSdFile] = useState(null);
  const [jiraError, setJiraError] = useState('');
  const [sdError, setSdError] = useState('');
  const [result, setResult] = useState(null);
  const [loading, setLoading] = useState(false);
  const [globalError, setGlobalError] = useState('');

  const handleJiraFile = useCallback((file) => {
    setJiraFile(file);
    setJiraError('');
    setResult(null);
    setGlobalError('');
  }, []);

  const handleSdFile = useCallback((file) => {
    setSdFile(file);
    setSdError('');
    setResult(null);
    setGlobalError('');
  }, []);

  const handleCompare = useCallback(async () => {
    setJiraError('');
    setSdError('');
    setGlobalError('');

    if (!jiraFile) { setJiraError('Выберите файл Jira'); return; }
    if (!sdFile) { setSdError('Выберите файл Service Desk'); return; }

    setLoading(true);
    try {
      const [jiraBuf, sdBuf] = await Promise.all([
        readFileAsBuffer(jiraFile),
        readFileAsBuffer(sdFile),
      ]);

      const jiraParsed = parseExcelFile(jiraBuf);
      if (jiraParsed.error) { setJiraError(jiraParsed.error); setLoading(false); return; }

      const sdParsed = parseExcelFile(sdBuf);
      if (sdParsed.error) { setSdError(sdParsed.error); setLoading(false); return; }

      const processed = processData(jiraParsed.data, sdParsed.data);
      setResult(processed);
    } catch (err) {
      setGlobalError('Ошибка обработки файлов: ' + err.message);
    } finally {
      setLoading(false);
    }
  }, [jiraFile, sdFile]);

  const handleExport = useCallback(() => {
    if (!result) return;
    exportToExcel(result.matches, result.unlinked);
  }, [result]);

  return (
    <div className="app">
      <header className="app-header">
        <h1>Сверка версий Jira ↔ Service Desk</h1>
        <p className="subtitle">
          Сравнение Fix Version/s между связанными задачами Jira (SIG/TTSH/TSH) и Jira Service Desk (HELP)
        </p>
      </header>

      <div className="upload-card">
        <div className="upload-row">
          <FileInput
            label="Файл Jira (SIG / TTSH / TSH)"
            onChange={handleJiraFile}
            fileName={jiraFile?.name}
            error={jiraError}
          />
          <FileInput
            label="Файл Service Desk (HELP)"
            onChange={handleSdFile}
            fileName={sdFile?.name}
            error={sdError}
          />
        </div>

        <div className="upload-actions">
          <button
            className="btn btn-primary"
            onClick={handleCompare}
            disabled={loading}
          >
            {loading ? 'Обработка...' : 'Сравнить'}
          </button>

          {result && (
            <button className="btn btn-export" onClick={handleExport}>
              Скачать Excel
            </button>
          )}
        </div>

        {globalError && <div className="global-error">{globalError}</div>}
      </div>

      {result && (
        <div className="results">
          <div className="summary-bar">
            Найдено пар: <strong>{result.pairedCount}</strong>&nbsp;|&nbsp;
            Без связей: <strong>{result.unlinked.length}</strong>
          </div>
          <MatchTable matches={result.matches} />
          <UnlinkedTable unlinked={result.unlinked} />
        </div>
      )}
    </div>
  );
}
