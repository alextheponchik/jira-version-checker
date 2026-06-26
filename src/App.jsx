import { useState, useCallback } from 'react';
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
        accept=".xlsx,.xls"
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

function MatchTable({ matches }) {
  const [filter, setFilter] = useState('all');

  const filtered = matches.filter((m) => {
    if (filter === 'match') return m.matched;
    if (filter === 'mismatch') return !m.matched;
    return true;
  });

  const matchCount = matches.filter((m) => m.matched).length;
  const mismatchCount = matches.filter((m) => !m.matched).length;

  return (
    <section className="section">
      <div className="section-header">
        <h2>Сопоставления</h2>
        <div className="section-stats">
          <Badge count={matchCount} color="green" /> совпадают&nbsp;&nbsp;
          <Badge count={mismatchCount} color="red" /> не совпадают
        </div>
        <div className="filter-buttons">
          <button className={filter === 'all' ? 'active' : ''} onClick={() => setFilter('all')}>Все</button>
          <button className={filter === 'mismatch' ? 'active' : ''} onClick={() => setFilter('mismatch')}>Не совпадают</button>
          <button className={filter === 'match' ? 'active' : ''} onClick={() => setFilter('match')}>Совпадают</button>
        </div>
      </div>
      <div className="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Задача Jira</th>
              <th>Задача Service Desk</th>
              <th>Fix Version/s (Jira)</th>
              <th>Fix Version/s (Service Desk)</th>
              <th>Статус</th>
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 && (
              <tr><td colSpan={5} className="empty-row">Нет данных</td></tr>
            )}
            {filtered.map((m, i) => (
              <tr key={i} className={m.matched ? 'row-green' : 'row-red'}>
                <td className="key-cell"><KeyLink keyText={m.jiraKey} url={m.jiraUrl} /></td>
                <td className="key-cell"><KeyLink keyText={m.helpKey} url={m.helpUrl} /></td>
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
  const jiraCount = unlinked.filter((u) => u.system === 'Jira').length;
  const sdCount = unlinked.filter((u) => u.system === 'Service Desk').length;

  return (
    <section className="section">
      <div className="section-header">
        <h2>Задачи без межсистемных связей</h2>
        <div className="section-stats">
          <Badge count={jiraCount} color="yellow" /> Jira&nbsp;&nbsp;
          <Badge count={sdCount} color="yellow" /> Service Desk
        </div>
      </div>
      <div className="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Ключ задачи</th>
              <th>Система</th>
              <th>Fix Version/s</th>
              <th>Linked Issues</th>
            </tr>
          </thead>
          <tbody>
            {unlinked.length === 0 && (
              <tr><td colSpan={4} className="empty-row">Все задачи имеют межсистемные связи</td></tr>
            )}
            {unlinked.map((u, i) => (
              <tr key={i} className="row-yellow">
                <td className="key-cell"><KeyLink keyText={u.key} url={u.url} /></td>
                <td>{u.system}</td>
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
