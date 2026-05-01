import React, { useState } from 'react';
import { createRoot } from 'react-dom/client';
import { Upload, FileSpreadsheet, CheckCircle2, AlertCircle, Loader2 } from 'lucide-react';
import './style.css';

const API_BASE = import.meta.env.VITE_API_BASE || 'http://localhost:8000';

function App() {
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [error, setError] = useState('');

  async function handleSubmit(e) {
    e.preventDefault();
    if (!file) {
      setError('Seleziona un PDF prima di continuare.');
      return;
    }
    setLoading(true);
    setError('');
    setResult(null);

    const formData = new FormData();
    formData.append('file', file);

    try {
      const response = await fetch(`${API_BASE}/api/convert`, {
        method: 'POST',
        body: formData,
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.detail || 'Errore durante la conversione.');
      setResult(data);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  }

  return (
    <main className="page">
      <section className="hero">
        <div className="badge">MVP Online</div>
        <h1>Da PDF scansionato a Excel</h1>
        <p>Carica una scansione PDF, il sistema esegue OCR e genera un file Excel strutturato.</p>
      </section>

      <section className="card">
        <form onSubmit={handleSubmit}>
          <label className="dropzone">
            <Upload size={42} />
            <strong>{file ? file.name : 'Carica PDF scansionato'}</strong>
            <span>Formato supportato: .pdf</span>
            <input
              type="file"
              accept="application/pdf"
              onChange={(e) => setFile(e.target.files?.[0] || null)}
            />
          </label>

          <button className="button" disabled={loading}>
            {loading ? <Loader2 className="spin" size={20} /> : <FileSpreadsheet size={20} />}
            {loading ? 'Elaborazione in corso...' : 'Converti in Excel'}
          </button>
        </form>
      </section>

      {error && (
        <section className="message error">
          <AlertCircle />
          <span>{error}</span>
        </section>
      )}

      {result && (
        <section className="card result">
          <div className="success-title">
            <CheckCircle2 />
            <h2>Excel pronto</h2>
          </div>
          <div className="grid">
            <div><span>Numero documento</span><strong>{result.summary.numero_documento || 'Da verificare'}</strong></div>
            <div><span>Data documento</span><strong>{result.summary.data_documento || 'Da verificare'}</strong></div>
            <div><span>Totale probabile</span><strong>{result.summary.totale_probabile || 'Da verificare'}</strong></div>
            <div><span>Righe rilevate</span><strong>{result.summary.righe_rilevate}</strong></div>
          </div>
          <a className="download" href={`${API_BASE}${result.download_url}`}>
            Scarica Excel
          </a>
        </section>
      )}

      <section className="note">
        <strong>Nota:</strong> questa versione usa OCR open-source. Per qualità massima su scansioni difficili, si può collegare Azure Document Intelligence o Google Document AI.
      </section>
    </main>
  );
}

createRoot(document.getElementById('root')).render(<App />);
