import { useState, useEffect, useCallback } from 'react'
import axios from 'axios'
import * as XLSX from 'xlsx'
import './App.css'

const API_URL = 'http://localhost:3001/api'

function App() {
  const [view, setView] = useState('scraper')
  const [urls, setUrls] = useState('')
  const [results, setResults] = useState([])
  const [loading, setLoading] = useState(false)
  const [progress, setProgress] = useState({ current: 0, total: 0 })
  const [usePuppeteer, setUsePuppeteer] = useState(false)
  const [activeTab, setActiveTab] = useState('all')
  const [searchTerm, setSearchTerm] = useState('')
  const [scrapings, setScrapings] = useState([])
  const [selectedScraping, setSelectedScraping] = useState(null)
  const [scrapingResults, setScrapingResults] = useState([])
  const [importedFileName, setImportedFileName] = useState('')

  useEffect(() => {
    if (view === 'history') fetchScrapings()
  }, [view])

  const fetchScrapings = async () => {
    try {
      const res = await axios.get(`${API_URL}/scrapings`)
      setScrapings(res.data.scrapings)
    } catch (e) {
      console.error('Error fetching scrapings:', e)
    }
  }

  const parseUrls = (text) =>
    text.split(/[\n,;]+/).map(u => u.trim()).filter(u => u.length > 0)

  const handleScrape = async () => {
    const urlList = parseUrls(urls)
    if (urlList.length === 0) return

    setLoading(true)
    setProgress({ current: 0, total: urlList.length })
    setResults([])
    setActiveTab('all')
    setSearchTerm('')

    try {
      const response = await fetch(`${API_URL}/scrape`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ urls: urlList, usePuppeteer, fileName: importedFileName || 'Manual Input' })
      })

      const reader = response.body.getReader()
      const decoder = new TextDecoder()
      const allResults = []

      while (true) {
        const { done, value } = await reader.read()
        if (done) break
        const text = decoder.decode(value)
        const lines = text.trim().split('\n')
        for (const line of lines) {
          if (line) {
            const data = JSON.parse(line)
            if (data.type === 'progress') {
              allResults.push(data.result)
              setProgress({ current: data.current, total: data.total })
              setResults([...allResults])
            }
          }
        }
      }
    } catch (error) {
      console.error('Scraping error:', error)
    } finally {
      setLoading(false)
    }
  }

  const handleImportExcel = async (e) => {
    const file = e.target.files[0]
    if (!file) return
    setImportedFileName(file.name.replace(/\.[^/.]+$/, ''))

    const reader = new FileReader()
    reader.onload = (event) => {
      const workbook = XLSX.read(event.target.result, { type: 'binary' })
      const allUrls = []
      workbook.SheetNames.forEach(sheetName => {
        const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 })
        data.forEach(row => {
          if (row[0] && typeof row[0] === 'string' && row[0].startsWith('http')) allUrls.push(row[0])
        })
      })
      setUrls(allUrls.join('\n'))
    }
    reader.readAsBinaryString(file)
  }

  const handleExportExcel = (data, filename = 'scraping-results.xlsx') => {
    if (data.length === 0) return
    const exportData = data.map(r => ({
      URL: r.url,
      Status: r.status || 'Error',
      Emails: (r.emails || []).join('; '),
      WhatsApp: (r.whatsapp || []).join('; '),
      LinkedIn: (r.linkedin || []).join('; '),
      Phone: (r.phoneNumbers || []).join('; '),
      Error: r.error || ''
    }))
    const worksheet = XLSX.utils.json_to_sheet(exportData)
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Results')
    XLSX.writeFile(workbook, filename)
  }

  const handleExportAll = async () => {
    try {
      const res = await axios.get(`${API_URL}/contacts/all`)
      handleExportExcel(res.data.results, 'all-contacts.xlsx')
    } catch (e) {
      console.error('Error exporting all:', e)
    }
  }

  const viewScraping = async (id) => {
    try {
      const res = await axios.get(`${API_URL}/scrapings/${id}`)
      setSelectedScraping(res.data.scraping)
      setScrapingResults(res.data.results)
      setView('scraping-detail')
      setActiveTab('all')
      setSearchTerm('')
    } catch (e) {
      console.error('Error fetching scraping:', e)
    }
  }

  const deleteScraping = async (id) => {
    if (!confirm('Supprimer ce scraping et tous ses résultats ?')) return
    try {
      await axios.delete(`${API_URL}/scrapings/${id}`)
      if (selectedScraping?.id === id) { setSelectedScraping(null); setScrapingResults([]) }
      fetchScrapings()
    } catch (e) { console.error(e) }
  }

  const deleteAllScrapings = async () => {
    if (!confirm('Supprimer TOUS les scrapings et résultats ?')) return
    try {
      await axios.delete(`${API_URL}/scrapings`)
      setSelectedScraping(null)
      setScrapingResults([])
      fetchScrapings()
    } catch (e) { console.error(e) }
  }

  const filterData = useCallback((data) => {
    return data.filter(r => {
      if (searchTerm && !r.url.toLowerCase().includes(searchTerm.toLowerCase())) return false
      if (activeTab === 'email') return r.emails?.length > 0
      if (activeTab === 'whatsapp') return r.whatsapp?.length > 0
      if (activeTab === 'linkedin') return r.linkedin?.length > 0
      if (activeTab === 'error') return r.error
      return true
    })
  }, [searchTerm, activeTab])

  const computeStats = (data) => ({
    total: data.length,
    emails: data.reduce((a, r) => a + (r.emails?.length || 0), 0),
    whatsapp: data.reduce((a, r) => a + (r.whatsapp?.length || 0), 0),
    linkedin: data.reduce((a, r) => a + (r.linkedin?.length || 0), 0),
    errors: data.filter(r => r.error).length
  })

  const currentData = view === 'scraping-detail' ? scrapingResults : results
  const filteredData = filterData(currentData)
  const stats = computeStats(currentData)

  const DataTable = ({ data, onExport }) => {
    const filtered = filterData(data)
    const s = computeStats(data)

    return (
      <>
        <div className="stats">
          <div className="stat-card"><span className="stat-value">{s.total}</span><span className="stat-label">Sites Scraped</span></div>
          <div className="stat-card stat-email"><span className="stat-value">{s.emails}</span><span className="stat-label">Emails Found</span></div>
          <div className="stat-card stat-whatsapp"><span className="stat-value">{s.whatsapp}</span><span className="stat-label">WhatsApp Found</span></div>
          <div className="stat-card stat-linkedin"><span className="stat-value">{s.linkedin}</span><span className="stat-label">LinkedIn Found</span></div>
          <div className="stat-card stat-error"><span className="stat-value">{s.errors}</span><span className="stat-label">Errors</span></div>
        </div>

        <div className="results-header">
          <div className="tabs">
            <button className={activeTab === 'all' ? 'active' : ''} onClick={() => setActiveTab('all')}>All ({data.length})</button>
            <button className={activeTab === 'email' ? 'active' : ''} onClick={() => setActiveTab('email')}>Emails ({data.filter(r => r.emails?.length > 0).length})</button>
            <button className={activeTab === 'whatsapp' ? 'active' : ''} onClick={() => setActiveTab('whatsapp')}>WhatsApp ({data.filter(r => r.whatsapp?.length > 0).length})</button>
            <button className={activeTab === 'linkedin' ? 'active' : ''} onClick={() => setActiveTab('linkedin')}>LinkedIn ({data.filter(r => r.linkedin?.length > 0).length})</button>
            <button className={activeTab === 'error' ? 'active' : ''} onClick={() => setActiveTab('error')}>Errors ({s.errors})</button>
          </div>
          <div className="results-actions">
            <input type="text" placeholder="Search URLs..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} className="search-input" />
            <button className="btn btn-secondary" onClick={() => onExport(filtered)}>Export Excel</button>
          </div>
        </div>

        <div className="table-container">
          <table className="results-table">
            <thead>
              <tr>
                <th>#</th>
                <th>URL</th>
                <th>Status</th>
                <th>Emails</th>
                <th>WhatsApp</th>
                <th>LinkedIn</th>
                <th>Phone</th>
                <th>Error</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((r, i) => (
                <tr key={i} className={r.error ? 'row-error' : ''}>
                  <td>{i + 1}</td>
                  <td><a href={r.url} target="_blank" rel="noopener noreferrer" className="table-url" title={r.url}>{r.url.length > 40 ? r.url.substring(0, 40) + '...' : r.url}</a></td>
                  <td><span className={`status-badge ${r.status === 200 ? 'success' : r.error ? 'error' : 'warning'}`}>{r.status || 'Err'}</span></td>
                   <td className="cell-values">{r.emails?.length > 0 ? r.emails.map((e, j) => <div key={j}><a href={`mailto:${e}`} className="inline-link">{e}</a></div>) : '-'}</td>
                   <td className="cell-values">{r.whatsapp?.length > 0 ? r.whatsapp.map((w, j) => <div key={j}><a href={`https://wa.me/${w.replace(/\D/g, '')}`} target="_blank" rel="noopener noreferrer" className="inline-link">{w}</a></div>) : '-'}</td>
                   <td className="cell-values">{r.linkedin?.length > 0 ? r.linkedin.map((l, j) => <div key={j}><a href={l} target="_blank" rel="noopener noreferrer" className="inline-link">{l.length > 40 ? l.substring(0, 40) + '...' : l}</a></div>) : '-'}</td>
                   <td className="cell-values">{r.phoneNumbers?.length > 0 ? r.phoneNumbers.map((p, j) => <div key={j}><a href={`tel:${p}`} className="inline-link">{p}</a></div>) : '-'}</td>
                  <td className="error-cell">{r.error ? r.error.substring(0, 30) + '...' : '-'}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

      </>
    )
  }

  return (
    <div className="app">
      <header className="header">
        <h1>Web Scraper - Contacts</h1>
        <p>Extract emails, WhatsApp numbers, and LinkedIn profiles</p>
        <nav className="nav-tabs">
          <button className={view === 'scraper' ? 'active' : ''} onClick={() => { setView('scraper'); setSelectedScraping(null); }}>Scraper</button>
          <button className={view === 'history' ? 'active' : ''} onClick={() => setView('history')}>Historique ({scrapings.length})</button>
          <button className="btn-export-all" onClick={handleExportAll}>Exporter Tout</button>
        </nav>
      </header>

      <main className="main">
        {view === 'scraper' && !selectedScraping && (
          <>
            <section className="input-section">
              <div className="card">
                <h2>Input URLs</h2>
                <div className="import-area">
                  <label className="file-upload">
                    <input type="file" accept=".xlsx,.xls,.csv" onChange={handleImportExcel} />
                    Import Excel File
                  </label>
                  <span className="import-hint">
                    Supports: .xlsx, .xls, .csv
                    {importedFileName && <span className="file-name"> • {importedFileName}</span>}
                  </span>
                </div>
                <textarea value={urls} onChange={(e) => setUrls(e.target.value)} placeholder="Enter URLs (one per line or separated by commas)&#10;&#10;Example:&#10;https://example.com&#10;https://example2.com/contact" rows={10} />
                <div className="options">
                  <label className="checkbox-label">
                    <input type="checkbox" checked={usePuppeteer} onChange={(e) => setUsePuppeteer(e.target.checked)} />
                    Use Puppeteer (slower but more accurate for JS-rendered pages)
                  </label>
                </div>
                <button className="btn btn-primary" onClick={handleScrape} disabled={loading || parseUrls(urls).length === 0}>
                  {loading ? `Scraping... (${progress.current}/${progress.total})` : 'Start Scraping'}
                </button>
              </div>

              {loading && (
                <div className="progress-bar">
                  <div className="progress-fill" style={{ width: `${(progress.current / progress.total) * 100}%` }} />
                  <span>{Math.round((progress.current / progress.total) * 100)}%</span>
                </div>
              )}
            </section>

            {results.length > 0 && (
              <section className="results-section">
                <DataTable data={results} onExport={(d) => handleExportExcel(d, 'scraping-results.xlsx')} />
              </section>
            )}
          </>
        )}

        {view === 'history' && (
          <section className="history-section">
            <div className="history-header">
              <h2>Historique des Scrapings</h2>
              {scrapings.length > 0 && <button className="btn btn-danger" onClick={deleteAllScrapings}>Supprimer Tout</button>}
            </div>

            {scrapings.length === 0 ? (
              <div className="empty-state">
                <p>Aucun scraping enregistré pour le moment.</p>
                <button className="btn btn-primary" onClick={() => setView('scraper')}>Commencer un scraping</button>
              </div>
            ) : (
              <div className="history-list">
                {scrapings.map(s => (
                  <div key={s.id} className="history-card" onClick={() => viewScraping(s.id)}>
                    <div className="history-info">
                      <h3>{s.file_name}</h3>
                      <span className="history-date">
                        {new Date(s.created_at).toLocaleDateString('fr-FR', { day: '2-digit', month: 'long', year: 'numeric', hour: '2-digit', minute: '2-digit' })}
                      </span>
                    </div>
                    <div className="history-stats">
                      <span className="history-stat">{s.total_urls} sites</span>
                      <span className="history-stat stat-email">{s.total_emails} emails</span>
                      <span className="history-stat stat-whatsapp">{s.total_whatsapp} whatsapp</span>
                      <span className="history-stat stat-linkedin">{s.total_linkedin} linkedin</span>
                      {s.total_errors > 0 && <span className="history-stat stat-error">{s.total_errors} erreurs</span>}
                    </div>
                    <button className="btn-delete" onClick={(e) => { e.stopPropagation(); deleteScraping(s.id); }}>✕</button>
                  </div>
                ))}
              </div>
            )}
          </section>
        )}

        {view === 'scraping-detail' && selectedScraping && (
          <section className="results-section">
            <div className="detail-header">
              <button className="btn btn-back" onClick={() => setView('history')}>← Retour à l'historique</button>
              <div>
                <h2>{selectedScraping.file_name}</h2>
                <span className="detail-date">
                  {new Date(selectedScraping.created_at).toLocaleDateString('fr-FR', { day: '2-digit', month: 'long', year: 'numeric', hour: '2-digit', minute: '2-digit' })}
                </span>
              </div>
            </div>
            <DataTable data={scrapingResults} onExport={(d) => handleExportExcel(d, `${selectedScraping.file_name}-results.xlsx`)} />
          </section>
        )}
      </main>
    </div>
  )
}

export default App
