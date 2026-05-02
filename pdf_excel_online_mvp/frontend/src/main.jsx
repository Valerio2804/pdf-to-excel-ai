import React, { useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import {
  Upload,
  Download,
  Search,
  FileText,
  ChevronRight,
  ChevronDown,
  CheckCircle2,
  AlertTriangle,
  Loader2,
  Trash2,
  ExternalLink,
  RefreshCw,
  Database,
  Package,
  Warehouse,
  FileSpreadsheet,
} from 'lucide-react';
import './style.css';

const API_BASE = import.meta.env.VITE_API_BASE || 'https://pdf-to-excel-ai.onrender.com';

function value(v, fallback = '—') {
  return v !== undefined && v !== null && v !== '' ? v : fallback;
}

function normalizeStatus(doc) {
  if (doc.status === 'error') return 'errore';
  if (doc.status === 'processing') return 'in lavorazione';
  const warnings = (doc.campi_da_verificare || []).length + (doc.errori || []).length;
  return warnings > 0 ? 'da verificare' : 'elaborato';
}

function downloadBlob(filename, content) {
  const blob = new Blob([content], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function App() {
  const fileInputRef = useRef(null);
  const [files, setFiles] = useState([]);
  const [documents, setDocuments] = useState([]);
  const [expanded, setExpanded] = useState({});
  const [query, setQuery] = useState('');
  const [loading, setLoading] = useState(false);
  const [globalError, setGlobalError] = useState('');

  const filteredDocs = useMemo(() => {
    const q = query.toLowerCase().trim();
    if (!q) return documents;
    return documents.filter((doc) => {
      const s = doc.summary || {};
      return [
        doc.fileName,
        s.numero_documento,
        s.data_documento,
        s.mittente,
        s.destinatario,
        s.indirizzo_destinatario,
        normalizeStatus(doc),
      ]
        .join(' ')
        .toLowerCase()
        .includes(q);
    });
  }, [documents, query]);

  const stats = useMemo(() => {
    const ddtTotali = documents.length;
    const righeTotali = documents.reduce((acc, doc) => acc + (doc.righe?.length || 0), 0);
    const daVerificare = documents.filter((doc) => normalizeStatus(doc) === 'da verificare').length;
    const magazzini = new Set(
      documents
        .map((doc) => doc.summary?.destinatario || doc.summary?.mittente)
        .filter(Boolean)
    ).size;
    return { ddtTotali, righeTotali, daVerificare, magazzini };
  }, [documents]);

  function toggleExpanded(id) {
    setExpanded((prev) => ({ ...prev, [id]: !prev[id] }));
  }

  function removeDoc(id) {
    setDocuments((prev) => prev.filter((doc) => doc.id !== id));
  }

  function exportAllRowsCsv() {
    const headers = [
      'File',
      'Numero Documento',
      'Data Documento',
      'Mittente',
      'Destinatario',
      'Codice',
      'Descrizione',
      'EAN',
      'Quantita',
      'Confidenza',
      'Stato',
    ];

    const rows = [];
    documents.forEach((doc) => {
      const s = doc.summary || {};
      (doc.righe || []).forEach((r) => {
        rows.push([
          doc.fileName,
          s.numero_documento,
          s.data_documento,
          s.mittente,
          s.destinatario,
          r.codice,
          r.descrizione,
          r.ean,
          r.quantita,
          r.confidenza,
          normalizeStatus(doc),
        ]);
      });
    });

    const csv = [headers, ...rows]
      .map((line) => line.map((cell) => `"${String(cell ?? '').replace(/"/g, '""')}"`).join(';'))
      .join('\n');

    downloadBlob('ddt_manager_export.csv', csv);
  }

  async function handleUpload(selectedFiles) {
    const pdfs = Array.from(selectedFiles || []).filter((f) => f.type === 'application/pdf' || f.name.toLowerCase().endsWith('.pdf'));

    if (!pdfs.length) {
      setGlobalError('Seleziona almeno un file PDF.');
      return;
    }

    setFiles(pdfs);
    setLoading(true);
    setGlobalError('');

    const pendingDocs = pdfs.map((file) => ({
      id: crypto.randomUUID(),
      fileName: file.name,
      status: 'processing',
      createdAt: new Date().toLocaleString('it-IT'),
      summary: {},
      righe: [],
      campi_da_verificare: [],
      errori: [],
      download_url: null,
    }));

    setDocuments((prev) => [...pendingDocs, ...prev]);

    const formData = new FormData();
    pdfs.forEach((file) => formData.append('files', file));

    try {
      const response = await fetch(`${API_BASE}/api/convert`, {
        method: 'POST',
        body: formData,
      });

      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.detail || 'Errore durante la conversione.');
      }

      const completedDocs = (data.documents || []).map((doc, index) => ({
        id: pendingDocs[index]?.id || doc.job_id || crypto.randomUUID(),
        fileName: doc.file || pendingDocs[index]?.fileName || 'Documento PDF',
        status: doc.status || 'completed',
        createdAt: pendingDocs[index]?.createdAt || new Date().toLocaleString('it-IT'),
        summary: doc.summary || {},
        righe: doc.righe || [],
        campi_da_verificare: doc.campi_da_verificare || [],
        errori: doc.errori || [],
        download_url: doc.download_url || null,
        error: doc.error || '',
      }));

      setDocuments((prev) => {
        const pendingIds = new Set(pendingDocs.map((d) => d.id));
        const oldDocs = prev.filter((d) => !pendingIds.has(d.id));
        return [...completedDocs, ...oldDocs];
      });
    } catch (err) {
      setGlobalError(err.message);
      setDocuments((prev) =>
        prev.map((doc) =>
          pendingDocs.some((p) => p.id === doc.id)
            ? { ...doc, status: 'error', error: err.message }
            : doc
        )
      );
    } finally {
      setLoading(false);
      setFiles([]);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  }

  return (
    <main className="ddt-app">
      <header className="topbar">
        <div className="brand">
          <div className="brand-icon"><Package size={24} /></div>
          <div>
            <h1>DDT Manager</h1>
            <p>Documenti di Trasporto</p>
          </div>
        </div>

        <div className="top-actions">
          <button className="secondary-button" onClick={exportAllRowsCsv} disabled={!documents.length}>
            <Download size={18} /> Esporta Excel
          </button>
          <button className="primary-button" onClick={() => fileInputRef.current?.click()} disabled={loading}>
            {loading ? <Loader2 className="spin" size={18} /> : <Upload size={18} />}
            Carica PDF
          </button>
          <input
            ref={fileInputRef}
            type="file"
            multiple
            accept="application/pdf"
            hidden
            onChange={(e) => handleUpload(e.target.files)}
          />
        </div>
      </header>

      <section className="kpi-row">
        <div className="kpi-card"><strong>{stats.ddtTotali}</strong><span>DDT Totali</span></div>
        <div className="kpi-card"><strong>{stats.magazzini}</strong><span>Magazzini</span></div>
        <div className="kpi-card"><strong>{stats.daVerificare}</strong><span>Da Verificare</span></div>
        <div className="kpi-card"><strong>{stats.righeTotali}</strong><span>Righe Totali</span></div>
      </section>

      {globalError && (
        <section className="alert-box">
          <AlertTriangle size={18} /> {globalError}
        </section>
      )}

      <section className="documents-panel">
        <div className="panel-header">
          <h2>{filteredDocs.length} documenti</h2>
          <div className="search-box">
            <Search size={18} />
            <input
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="Cerca per numero, mittente, destinatario..."
            />
          </div>
        </div>

        <div className="table">
          <div className="table-head">
            <span></span>
            <span>Numero</span>
            <span>Data</span>
            <span>Mittente</span>
            <span>Destinatario</span>
            <span>Destinazione</span>
            <span>Stato</span>
            <span>Azioni</span>
          </div>

          {filteredDocs.length === 0 && (
            <div className="empty-state">
              <FileSpreadsheet size={36} />
              <h3>Nessun DDT caricato</h3>
              <p>Carica uno o più PDF per iniziare l’elaborazione.</p>
            </div>
          )}

          {filteredDocs.map((doc) => {
            const s = doc.summary || {};
            const isExpanded = Boolean(expanded[doc.id]);
            const status = normalizeStatus(doc);
            return (
              <div className="doc-block" key={doc.id}>
                <div className="table-row">
                  <button className="icon-button" onClick={() => toggleExpanded(doc.id)}>
                    {isExpanded ? <ChevronDown size={18} /> : <ChevronRight size={18} />}
                  </button>
                  <span>{value(s.numero_documento)}</span>
                  <span>{value(s.data_documento)}</span>
                  <span>{value(s.mittente)}</span>
                  <span>{value(s.destinatario)}</span>
                  <span>{value(s.indirizzo_destinatario, 'Non specificato')}</span>
                  <span><em className={`pill ${status.replace(' ', '-')}`}>{status}</em></span>
                  <span className="row-actions">
                    {doc.status === 'processing' && <Loader2 className="spin" size={17} />}
                    {doc.download_url && (
                      <a title="Scarica Excel" href={`${API_BASE}${doc.download_url}`}>
                        <Download size={17} />
                      </a>
                    )}
                    <button title="Rielabora" onClick={() => fileInputRef.current?.click()}><RefreshCw size={17} /></button>
                    <button title="Elimina" onClick={() => removeDoc(doc.id)}><Trash2 size={17} /></button>
                  </span>
                </div>

                {isExpanded && (
                  <div className="details">
                    <div className="detail-summary">
                      <div><span>File</span><strong>{doc.fileName}</strong></div>
                      <div><span>Totale pezzi</span><strong>{value(s.totale_pezzi, 0)}</strong></div>
                      <div><span>Confidenza</span><strong>{value(s.confidenza, 'N/D')}</strong></div>
                      <div><span>Righe</span><strong>{doc.righe?.length || 0}</strong></div>
                    </div>

                    {doc.error && <div className="alert-box small"><AlertTriangle size={16} />{doc.error}</div>}

                    <div className="subtable-wrap">
                      <table className="subtable">
                        <thead>
                          <tr>
                            <th>Codice</th>
                            <th>Descrizione</th>
                            <th>EAN</th>
                            <th>Quantità</th>
                            <th>Confidenza</th>
                          </tr>
                        </thead>
                        <tbody>
                          {(doc.righe || []).length === 0 ? (
                            <tr><td colSpan="5">Nessuna riga rilevata.</td></tr>
                          ) : (
                            doc.righe.map((r, idx) => (
                              <tr key={idx}>
                                <td>{value(r.codice)}</td>
                                <td>{value(r.descrizione)}</td>
                                <td>{value(r.ean)}</td>
                                <td>{value(r.quantita, 0)}</td>
                                <td>{value(r.confidenza, 'N/D')}</td>
                              </tr>
                            ))
                          )}
                        </tbody>
                      </table>
                    </div>

                    <div className="warning-grid">
                      <div>
                        <h4><AlertTriangle size={16} /> Campi da verificare</h4>
                        {(doc.campi_da_verificare || []).length ? (
                          <ul>{doc.campi_da_verificare.map((x, i) => <li key={i}>{String(x)}</li>)}</ul>
                        ) : <p>Nessun campo critico.</p>}
                      </div>
                      <div>
                        <h4><CheckCircle2 size={16} /> Errori</h4>
                        {(doc.errori || []).length ? (
                          <ul>{doc.errori.map((x, i) => <li key={i}>{String(x)}</li>)}</ul>
                        ) : <p>Nessun errore rilevato.</p>}
                      </div>
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </section>
    </main>
  );
}

createRoot(document.getElementById('root')).render(<App />);
