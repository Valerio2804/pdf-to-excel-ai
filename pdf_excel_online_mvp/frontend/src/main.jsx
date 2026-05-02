import React, { useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import {
  Upload,
  FileSpreadsheet,
  CheckCircle2,
  AlertCircle,
  Loader2,
  FileText,
  Table,
  ShieldCheck,
  Download,
  RefreshCw
} from 'lucide-react';
import './style.css';

const API_BASE = import.meta.env.VITE_API_BASE || 'https://pdf-to-excel-ai.onrender.com';

function safe(value, fallback = 'Da verificare') {
  return value !== undefined && value !== null && value !== '' ? value : fallback;
}

function App() {
  const [files, setFiles] = useState([]);
  const [documents, setDocuments] = useState([]);
  const [selectedId, setSelectedId] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  const selected = useMemo(
    () => documents.find((doc) => doc.id === selectedId) || documents[0] || null,
    [documents, selectedId]
  );

  async function handleSubmit(e) {
    e.preventDefault();

    if (!files.length) {
      setError('Seleziona almeno un PDF prima di continuare.');
      return;
    }

    setLoading(true);
    setError('');

    const pendingDocs = files.map((file) => ({
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
    setSelectedId(pendingDocs[0].id);

    const formData = new FormData();

    files.forEach((file) => {
      formData.append('files', file);
    });

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
        fileName: doc.file || pendingDocs[index]?.fileName || 'Documento',
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
        const pendingIds = new Set(pendingDocs.map((doc) => doc.id));
        const oldDocs = prev.filter((doc) => !pendingIds.has(doc.id));
        return [...completedDocs, ...oldDocs];
      });

      setSelectedId(completedDocs[0]?.id || null);
      setFiles([]);
    } catch (err) {
      setError(err.message);

      setDocuments((prev) =>
        prev.map((doc) =>
          pendingDocs.some((p) => p.id === doc.id)
            ? { ...doc, status: 'error', error: err.message }
            : doc
        )
      );
    } finally {
      setLoading(false);
    }
  }

  return (
    <main className="page enterprise-page">
      <section className="hero enterprise-hero">
        <div className="badge">Enterprise DDT AI</div>
        <h1>Da PDF scansionati a Excel gestionale</h1>
        <p>
          Carica uno o più DDT, bolle e documenti logistici. L’AI Vision estrae
          testata, righe articoli, quantità, campi incerti e genera Excel professionali.
        </p>
      </section>

      <section className="enterprise-layout">
        <aside className="sidebar card">
          <div className="panel-title">
            <FileText size={20} />
            <h2>Documenti</h2>
          </div>

          {documents.length === 0 && (
            <p className="muted">Nessun documento elaborato in questa sessione.</p>
          )}

          <div className="document-list">
            {documents.map((doc) => (
              <button
                key={doc.id}
                className={`document-item ${selected?.id === doc.id ? 'active' : ''}`}
                onClick={() => setSelectedId(doc.id)}
              >
                <strong>{doc.fileName}</strong>
                <span>{doc.createdAt}</span>
                <em className={`status ${doc.status}`}>
                  {doc.status === 'processing' && 'In elaborazione'}
                  {doc.status === 'completed' && 'Completato'}
                  {doc.status === 'error' && 'Errore'}
                </em>
              </button>
            ))}
          </div>
        </aside>

        <section className="main-panel">
          <section className="card upload-card">
            <form onSubmit={handleSubmit}>
              <label className="dropzone">
                <Upload size={42} />
                <strong>
                  {files.length
                    ? `${files.length} PDF selezionati`
                    : 'Carica PDF scansionati'}
                </strong>
                <span>Formato supportato: .pdf — puoi selezionare più file</span>
                <input
                  type="file"
                  accept="application/pdf"
                  multiple
                  onChange={(e) => setFiles(Array.from(e.target.files || []))}
                />
              </label>

              {files.length > 0 && (
                <div className="selected-files">
                  {files.map((file, index) => (
                    <span key={`${file.name}-${index}`}>{file.name}</span>
                  ))}
                </div>
              )}

              <button className="button" disabled={loading}>
                {loading ? <Loader2 className="spin" size={20} /> : <FileSpreadsheet size={20} />}
                {loading ? 'Elaborazione AI in corso...' : 'Converti PDF in Excel'}
              </button>
            </form>
          </section>

          {error && (
            <section className="message error">
              <AlertCircle />
              <span>{error}</span>
            </section>
          )}

          {selected && (
            <section className="card result enterprise-result">
              <div className="success-title">
                {selected.status === 'processing' ? (
                  <Loader2 className="spin" />
                ) : selected.status === 'error' ? (
                  <AlertCircle />
                ) : (
                  <CheckCircle2 />
                )}

                <h2>
                  {selected.status === 'processing' && 'Documento in elaborazione'}
                  {selected.status === 'completed' && 'Documento elaborato'}
                  {selected.status === 'error' && 'Errore documento'}
                </h2>
              </div>

              {selected.status === 'error' && (
                <section className="message error">
                  <AlertCircle />
                  <span>{selected.error || 'Errore durante elaborazione.'}</span>
                </section>
              )}

              {selected.status === 'completed' && (
                <>
                  <div className="kpi-grid">
                    <div className="kpi">
                      <span>Tipo documento</span>
                      <strong>{safe(selected.summary.tipo_documento, 'DDT')}</strong>
                    </div>
                    <div className="kpi">
                      <span>Numero</span>
                      <strong>{safe(selected.summary.numero_documento)}</strong>
                    </div>
                    <div className="kpi">
                      <span>Data</span>
                      <strong>{safe(selected.summary.data_documento)}</strong>
                    </div>
                    <div className="kpi">
                      <span>Totale pezzi</span>
                      <strong>{safe(selected.summary.totale_pezzi, 0)}</strong>
                    </div>
                    <div className="kpi">
                      <span>Righe</span>
                      <strong>{selected.righe?.length || 0}</strong>
                    </div>
                    <div className="kpi">
                      <span>Confidenza</span>
                      <strong>{safe(selected.summary.confidenza, 'N/D')}</strong>
                    </div>
                  </div>

                  <div className="summary-grid">
                    <div>
                      <span>Mittente</span>
                      <strong>{safe(selected.summary.mittente)}</strong>
                    </div>
                    <div>
                      <span>Destinatario</span>
                      <strong>{safe(selected.summary.destinatario)}</strong>
                    </div>
                    <div>
                      <span>Indirizzo destinatario</span>
                      <strong>{safe(selected.summary.indirizzo_destinatario)}</strong>
                    </div>
                  </div>

                  <div className="section-header">
                    <Table size={20} />
                    <h3>Righe DDT</h3>
                  </div>

                  <div className="table-wrap">
                    <table>
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
                        {(selected.righe || []).length === 0 ? (
                          <tr>
                            <td colSpan="5">Nessuna riga rilevata.</td>
                          </tr>
                        ) : (
                          selected.righe.map((riga, index) => (
                            <tr key={index}>
                              <td>{safe(riga.codice, '-')}</td>
                              <td>{safe(riga.descrizione, '-')}</td>
                              <td>{safe(riga.ean, '-')}</td>
                              <td>{safe(riga.quantita, 0)}</td>
                              <td>{safe(riga.confidenza, 'N/D')}</td>
                            </tr>
                          ))
                        )}
                      </tbody>
                    </table>
                  </div>

                  <div className="enterprise-bottom">
                    <section className="mini-panel">
                      <div className="section-header">
                        <ShieldCheck size={18} />
                        <h3>Campi da verificare</h3>
                      </div>

                      {(selected.campi_da_verificare || []).length === 0 ? (
                        <p className="muted">Nessun campo critico segnalato.</p>
                      ) : (
                        <ul>
                          {selected.campi_da_verificare.map((item, index) => (
                            <li key={index}>
                              {typeof item === 'string' ? item : JSON.stringify(item)}
                            </li>
                          ))}
                        </ul>
                      )}
                    </section>

                    <section className="mini-panel">
                      <div className="section-header">
                        <AlertCircle size={18} />
                        <h3>Errori</h3>
                      </div>

                      {(selected.errori || []).length === 0 ? (
                        <p className="muted">Nessun errore rilevato.</p>
                      ) : (
                        <ul>
                          {selected.errori.map((item, index) => (
                            <li key={index}>
                              {typeof item === 'string' ? item : JSON.stringify(item)}
                            </li>
                          ))}
                        </ul>
                      )}
                    </section>
                  </div>

                  <div className="actions">
                    {selected.download_url && (
                      <a className="download" href={`${API_BASE}${selected.download_url}`}>
                        <Download size={18} />
                        Scarica Excel
                      </a>
                    )}

                    <button
                      className="secondary-button"
                      type="button"
                      onClick={() => {
                        setFiles([]);
                        setError('');
                      }}
                    >
                      <RefreshCw size={18} />
                      Nuovo caricamento
                    </button>
                  </div>
                </>
              )}
            </section>
          )}
        </section>
      </section>

      <section className="note">
        <strong>Enterprise mode:</strong> AI Vision, caricamento multiplo, JSON strutturato,
        validazione campi, righe DDT, storico sessione e download Excel gestionale.
      </section>
    </main>
  );
}

createRoot(document.getElementById('root')).render(<App />);
