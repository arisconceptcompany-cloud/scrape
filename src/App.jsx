import { useState, useEffect, useCallback } from 'react'
import axios from 'axios'
import * as XLSX from 'xlsx'
import './App.css'

const API_URL = 'https://apiscrape.aris-cc.com/api'

function App() {
  const [view, setView] = useState('extraction')
  const [urls, setUrls] = useState('')
  const [results, setResults] = useState([])
  const [loading, setLoading] = useState(false)
  const [progress, setProgress] = useState({ current: 0, total: 0 })
  const [usePuppeteer, setUsePuppeteer] = useState(false)
  const [deepScan, setDeepScan] = useState(true)
  const [currentSessionId, setCurrentSessionId] = useState(null)
  const [activeTab, setActiveTab] = useState('all')
  const [searchTerm, setSearchTerm] = useState('')
  const [extractions, setExtractions] = useState([])
  const [selectedExtraction, setSelectedExtraction] = useState(null)
  const [extractionResults, setExtractionResults] = useState([])
  const [importedFileName, setImportedFileName] = useState('')
  const [loadingHistory, setLoadingHistory] = useState(false)
  const [remainingUrls, setRemainingUrls] = useState([])
  const [allUrlList, setAllUrlList] = useState([])

  useEffect(() => {
    if (view === 'history') {
      fetchExtractions()
    }
  }, [view])

  // Charger les derniers résultats depuis localStorage au démarrage
  useEffect(() => {
    try {
      const savedData = localStorage.getItem('lastExtraction')
      if (savedData) {
        const { results: savedResults, date, fileName } = JSON.parse(savedData)
        if (savedResults && savedResults.length > 0 && results.length === 0) {
          setResults(savedResults)
          if (fileName) setImportedFileName(fileName)
        }
      }
    } catch (e) {
      console.error('Error loading from localStorage:', e)
    }
  }, [])

  // Sauvegarder les résultats dans localStorage à chaque changement
  useEffect(() => {
    if (results.length > 0) {
      const saveData = {
        results: results,
        date: new Date().toISOString(),
        fileName: importedFileName || 'Saisie manuelle'
      }
      localStorage.setItem('lastExtraction', JSON.stringify(saveData))
    }
  }, [results, importedFileName])

  const fetchExtractions = async () => {
    setLoadingHistory(true)
    try {
      const res = await axios.get(`${API_URL}/scrapings`)
      console.log('Extractions response:', res.data)
      
      let extractionsData = []
      if (Array.isArray(res.data)) {
        extractionsData = res.data
      } else if (res.data.scrapings && Array.isArray(res.data.scrapings)) {
        extractionsData = res.data.scrapings
      } else if (res.data.data && Array.isArray(res.data.data)) {
        extractionsData = res.data.data
      }
      
      setExtractions(extractionsData)
      console.log('Extractions set:', extractionsData.length)
    } catch (e) {
      console.error('Error fetching extractions:', e)
      setExtractions([])
    } finally {
      setLoadingHistory(false)
    }
  }

  const parseUrls = (text) =>
    text.split(/[\n,;]+/).map(u => u.trim()).filter(u => u.length > 0)

  const handleExtract = async (urlsToExtract, append = false) => {
    const urlList = urlsToExtract || parseUrls(urls)
    if (urlList.length === 0) return

    setLoading(true)
    setProgress({ current: 0, total: urlList.length })
    if (!append) {
      setResults([])
      setAllUrlList(urlList)
    }
    setRemainingUrls([])
    setActiveTab('all')
    setSearchTerm('')

    try {
      const response = await fetch(`${API_URL}/scrape`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ urls: urlList, usePuppeteer, deepScan, fileName: importedFileName || 'Saisie manuelle' })
      })

      const reader = response.body.getReader()
      const decoder = new TextDecoder()
      const newResults = []
      let buffer = ''

      while (true) {
        const { done, value } = await reader.read()
        if (done) break
        buffer += decoder.decode(value, { stream: true })

        const lines = buffer.split('\n')
        buffer = lines.pop()

        for (const line of lines) {
          if (!line) continue
          try {
            const data = JSON.parse(line)
            if (data.type === 'session') {
              setCurrentSessionId(data.sessionId)
            } else if (data.type === 'progress') {
              newResults.push(data.result)
              setProgress({ current: data.current, total: data.total })
              const remaining = urlList.slice(data.current)
              setRemainingUrls(remaining)
              setResults(prev => append ? [...prev, ...newResults] : [...newResults])
            }
          } catch (e) {
            console.error('Error parsing line:', e)
          }
        }
      }

      if (buffer.trim()) {
        try {
          const data = JSON.parse(buffer)
          if (data.type === 'progress') {
            newResults.push(data.result)
            setResults(prev => append ? [...prev, ...newResults] : [...newResults])
          }
        } catch (e) {
          console.error('Error parsing final buffer:', e)
        }
      }

      setRemainingUrls([])

      const allResults = append ? [...results, ...newResults] : newResults
      const saveData = {
        results: allResults,
        date: new Date().toISOString(),
        fileName: importedFileName || 'Saisie manuelle'
      }
      localStorage.setItem('lastExtraction', JSON.stringify(saveData))

      await fetchExtractions()
    } catch (error) {
      console.error('Extraction error:', error)
    } finally {
      setLoading(false)
      setCurrentSessionId(null)
    }
  }

  const handleStop = async () => {
    if (!currentSessionId) return
    try {
      await fetch(`${API_URL}/scrape/${currentSessionId}/stop`, {
        method: 'POST'
      })
      setLoading(false)
      setCurrentSessionId(null)
      console.log('Extraction arrêtée')
    } catch (error) {
      console.error('Erreur lors de l\'arrêt:', error)
    }
  }

  const handleImportExcel = async (e) => {
    const file = e.target.files[0]
    if (!file) return
    // Définir le nom par défaut depuis le fichier, mais modifiable
    setImportedFileName(file.name.replace(/\.[^/.]+$/, ''))

    const reader = new FileReader()
    reader.onload = (event) => {
      try {
        const workbook = XLSX.read(event.target.result, { type: 'binary' })
        const allUrls = []
        workbook.SheetNames.forEach(sheetName => {
          const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 })
          data.forEach(row => {
            if (row[0] && typeof row[0] === 'string' && row[0].startsWith('http')) allUrls.push(row[0])
          })
        })
        setUrls(allUrls.join('\n'))
      } catch (err) {
        console.error('Error reading Excel file:', err)
        alert('Erreur lors de la lecture du fichier Excel')
      }
    }
    reader.readAsBinaryString(file)
  }

  const handleExportExcel = (data, filename = 'extraction-results.xlsx') => {
    if (data.length === 0) {
      alert('Aucune donnée à exporter')
      return
    }
    try {
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
    } catch (err) {
      console.error('Error exporting Excel:', err)
      alert('Erreur lors de l\'export')
    }
  }

  const handleExportAll = async () => {
    try {
      const res = await axios.get(`${API_URL}/contacts/all`)
      const data = res.data.results || res.data.data || res.data
      handleExportExcel(Array.isArray(data) ? data : [], 'all-contacts.xlsx')
    } catch (e) {
      console.error('Error exporting all:', e)
      alert('Erreur lors de l\'export de tous les contacts')
    }
  }

  const viewExtraction = async (id) => {
    try {
      const res = await axios.get(`${API_URL}/scrapings/${id}`)
      console.log('Extraction detail response:', res.data)
      
      const extraction = res.data.scraping || res.data.data || res.data
      const results = res.data.results || res.data.contacts || []
      
      setSelectedExtraction(extraction)
      setExtractionResults(Array.isArray(results) ? results : [])
      setView('extraction-detail')
      setActiveTab('all')
      setSearchTerm('')
    } catch (e) {
      console.error('Error fetching extraction:', e)
      alert('Erreur lors du chargement de l\'extraction')
    }
  }

  const deleteExtraction = async (id) => {
    if (!confirm('Supprimer cette extraction et tous ses résultats ?')) return
    try {
      await axios.delete(`${API_URL}/scrapings/${id}`)
      if (selectedExtraction?.id === id) { 
        setSelectedExtraction(null)
        setExtractionResults([]) 
      }
      await fetchExtractions()
    } catch (e) { 
      console.error('Error deleting extraction:', e)
      alert('Erreur lors de la suppression')
    }
  }

  const deleteAllExtractions = async () => {
    if (!confirm('Supprimer TOUTES les extractions et résultats ?')) return
    try {
      await axios.delete(`${API_URL}/scrapings`)
      setSelectedExtraction(null)
      setExtractionResults([])
      await fetchExtractions()
    } catch (e) { 
      console.error('Error deleting all extractions:', e)
      alert('Erreur lors de la suppression')
    }
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

  const currentData = view === 'extraction-detail' ? extractionResults : results
  const filteredData = filterData(currentData)
  const stats = computeStats(currentData)

  const DataTable = ({ data, onExport }) => {
    const filtered = filterData(data)
    const s = computeStats(data)

    return (
      <>
        <div className="stats">
          <div className="stat-card"><span className="stat-value">{s.total}</span><span className="stat-label">Sites traités</span></div>
          <div className="stat-card stat-email"><span className="stat-value">{s.emails}</span><span className="stat-label">Emails trouvés</span></div>
          <div className="stat-card stat-whatsapp"><span className="stat-value">{s.whatsapp}</span><span className="stat-label">WhatsApp trouvés</span></div>
          <div className="stat-card stat-linkedin"><span className="stat-value">{s.linkedin}</span><span className="stat-label">LinkedIn trouvés</span></div>
          <div className="stat-card stat-error"><span className="stat-value">{s.errors}</span><span className="stat-label">Erreurs</span></div>
        </div>

        <div className="results-header">
          <div className="tabs">
            <button className={activeTab === 'all' ? 'active' : ''} onClick={() => setActiveTab('all')}>Tout ({data.length})</button>
            <button className={activeTab === 'email' ? 'active' : ''} onClick={() => setActiveTab('email')}>Emails ({data.filter(r => r.emails?.length > 0).length})</button>
            <button className={activeTab === 'whatsapp' ? 'active' : ''} onClick={() => setActiveTab('whatsapp')}>WhatsApp ({data.filter(r => r.whatsapp?.length > 0).length})</button>
            <button className={activeTab === 'linkedin' ? 'active' : ''} onClick={() => setActiveTab('linkedin')}>LinkedIn ({data.filter(r => r.linkedin?.length > 0).length})</button>
            <button className={activeTab === 'error' ? 'active' : ''} onClick={() => setActiveTab('error')}>Erreurs ({s.errors})</button>
          </div>
          <div className="results-actions">
            <input type="text" placeholder="Rechercher des URLs..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} className="search-input" />
            <button className="btn btn-secondary" onClick={() => onExport(filtered)}>Exporter Excel</button>
          </div>
        </div>

        <div className="table-container">
          <table className="results-table">
            <thead>
              <tr>
                <th>#</th>
                <th>URL</th>
                <th>Statut</th>
                <th>Emails</th>
                <th>WhatsApp</th>
                <th>LinkedIn</th>
                <th>Téléphone</th>
                <th>Erreur</th>
              </tr>
            </thead>
            <tbody>
              {filtered.length > 0 ? (
                filtered.map((r, i) => (
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
                ))
              ) : (
                <tr>
                  <td colSpan="8" style={{ textAlign: 'center', padding: '20px', color: '#999' }}>Aucune donnée trouvée</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </>
    )
  }

  return (
    <div className="app">
      <header className="header">
        <div className="header-content">
          <div className="header-text">
            <h1>Extracteur de Contacts</h1>
            <p>Extrait les emails, numéros WhatsApp et profils LinkedIn de sites web</p>
          </div>
        </div>
        <nav className="nav-tabs">
          <button className={view === 'extraction' ? 'active' : ''} onClick={() => { setView('extraction'); setSelectedExtraction(null); }}>Nouvelle extraction</button>
          <button className={view === 'history' ? 'active' : ''} onClick={() => setView('history')}>
            Historique {extractions.length > 0 && `(${extractions.length})`}
          </button>
          <button className="btn-export-all" onClick={handleExportAll}>Exporter tout</button>
        </nav>
      </header>

      <main className="main">
        {view === 'extraction' && !selectedExtraction && (
          <>
            <section className="input-section">
              <div className="card">
                <h2>Saisir les URLs</h2>
                <div className="import-area">
                  <label className="file-upload">
                    <input type="file" accept=".xlsx,.xls,.csv" onChange={handleImportExcel} />
                    Importer un fichier Excel
                  </label>
                  <span className="import-hint">
                    Formats supportés: .xlsx, .xls, .csv
                    {importedFileName && <span className="file-name"> • {importedFileName}</span>}
                  </span>
                </div>
                <div className="name-input">
                  <label>Nom de l'extraction :</label>
                  <input 
                    type="text" 
                    value={importedFileName} 
                    onChange={(e) => setImportedFileName(e.target.value)}
                    placeholder="Ex: Clients Paris, Site officiel..."
                    className="name-field"
                  />
                </div>

                <textarea 
                  value={urls} 
                  onChange={(e) => setUrls(e.target.value)} 
                  placeholder="Entrez les URLs (une par ligne ou séparées par des virgules)&#10;&#10;Exemple:&#10;https://exemple.com&#10;https://exemple2.com/contact" 
                  rows={10} 
                />
                <div className="options">
                  <label className="checkbox-label">
                    <input type="checkbox" checked={usePuppeteer} onChange={(e) => setUsePuppeteer(e.target.checked)} />
                    Mode avancé - Plus lent mais plus précis (pages dynamiques)
                  </label>
                  <label className="checkbox-label">
                    <input type="checkbox" checked={deepScan} onChange={(e) => setDeepScan(e.target.checked)} />
                    Exploration complète - Parcourt tous les menus et sous-pages
                  </label>
                </div>
                <div className="button-group">
                  {loading ? (
                    <>
                      <div className="btn btn-loading">
                        <span className="spinner" />
                        Extraction en cours... ({progress.current}/{progress.total})
                      </div>
                      <button className="btn btn-danger" onClick={handleStop}>
                        Arrêter
                      </button>
                    </>
                  ) : (
                    <>
                      <button 
                        className="btn btn-primary" 
                        onClick={() => handleExtract(parseUrls(urls), false)} 
                        disabled={parseUrls(urls).length === 0}
                      >
                        Lancer l'extraction
                      </button>
                      {remainingUrls.length > 0 && (
                        <button 
                          className="btn btn-resume" 
                          onClick={() => handleExtract(remainingUrls, true)}
                        >
                          Reprendre ({remainingUrls.length} restantes)
                        </button>
                      )}
                      {localStorage.getItem('lastExtraction') && (
                        <button 
                          className="btn btn-secondary" 
                          onClick={() => {
                            try {
                              const savedData = JSON.parse(localStorage.getItem('lastExtraction'))
                              if (savedData.results && savedData.results.length > 0) {
                                setResults(savedData.results)
                                setView('extraction')
                              }
                            } catch (e) {
                              console.error('Error loading saved results:', e)
                            }
                          }}
                        >
                          Derniers résultats
                        </button>
                      )}
                    </>
                  )}
                </div>
              </div>

              {loading && (
                <div className="progress-section">
                  <div className="progress-header">
                    <span className="progress-label">Progression</span>
                    <span className="progress-count">{progress.current}/{progress.total}</span>
                  </div>
                  <div className="progress-track">
                    <div className="progress-fill" style={{ width: `${(progress.current / progress.total) * 100}%` }} />
                    <div className="progress-glow" style={{ width: `${(progress.current / progress.total) * 100}%` }} />
                  </div>
                </div>
              )}
            </section>

            {results.length > 0 && (
              <section className="results-section">
                {remainingUrls.length > 0 && !loading && (
                  <div className="resume-banner">
                    <div className="resume-banner-icon">⏸</div>
                    <div className="resume-banner-text">
                      <strong>Extraction interrompue</strong> — {remainingUrls.length} URL{remainingUrls.length > 1 ? 's' : ''} non traitée{remainingUrls.length > 1 ? 's' : ''}
                    </div>
                    <button className="btn btn-resume" onClick={() => handleExtract(remainingUrls, true)}>
                      Reprendre
                    </button>
                  </div>
                )}
                <DataTable data={results} onExport={(d) => handleExportExcel(d, 'extraction-results.xlsx')} />
              </section>
            )}
          </>
        )}

        {view === 'history' && (
          <section className="history-section">
            <div className="history-header">
              <h2>Historique des extractions</h2>
              {extractions.length > 0 && <button className="btn btn-danger" onClick={deleteAllExtractions}>Tout supprimer</button>}
            </div>

            {loadingHistory ? (
              <div className="empty-state">
                <p>Chargement des extractions...</p>
              </div>
            ) : extractions.length === 0 ? (
              <div className="empty-state">
                <p>Aucune extraction enregistrée pour le moment.</p>
                <button className="btn btn-primary" onClick={() => setView('extraction')}>Commencer une extraction</button>
              </div>
            ) : (
              <div className="history-list">
                {extractions.map(s => (
                  <div key={s.id} className="history-card" onClick={() => viewExtraction(s.id)}>
                    <div className="history-info">
                      <h3>{s.file_name || s.fileName || 'Sans titre'}</h3>
                      <span className="history-date">
                        {new Date(s.created_at || s.createdAt).toLocaleDateString('fr-FR', { 
                          day: '2-digit', 
                          month: 'long', 
                          year: 'numeric', 
                          hour: '2-digit', 
                          minute: '2-digit' 
                        })}
                      </span>
                    </div>
                    <div className="history-stats">
                      <span className="history-stat">{s.total_urls || s.totalUrls || 0} sites</span>
                      <span className="history-stat stat-email">{s.total_emails || s.totalEmails || 0} emails</span>
                      <span className="history-stat stat-whatsapp">{s.total_whatsapp || s.totalWhatsapp || 0} whatsapp</span>
                      <span className="history-stat stat-linkedin">{s.total_linkedin || s.totalLinkedin || 0} linkedin</span>
                      {(s.total_errors || s.totalErrors || 0) > 0 && <span className="history-stat stat-error">{s.total_errors || s.totalErrors} erreurs</span>}
                    </div>
                    <button className="btn-delete" onClick={(e) => { e.stopPropagation(); deleteExtraction(s.id); }}>✕</button>
                  </div>
                ))}
              </div>
            )}
          </section>
        )}

        {view === 'extraction-detail' && selectedExtraction && (
          <section className="results-section">
            <div className="detail-header">
              <button className="btn btn-back" onClick={() => setView('history')}>← Retour à l'historique</button>
              <div>
                <h2>{selectedExtraction.file_name || selectedExtraction.fileName || 'Sans titre'}</h2>
                <span className="detail-date">
                  {new Date(selectedExtraction.created_at || selectedExtraction.createdAt).toLocaleDateString('fr-FR', { 
                    day: '2-digit', 
                    month: 'long', 
                    year: 'numeric', 
                    hour: '2-digit', 
                    minute: '2-digit' 
                  })}
                </span>
              </div>
            </div>
            <DataTable data={extractionResults} onExport={(d) => handleExportExcel(d, `${selectedExtraction.file_name || selectedExtraction.fileName || 'results'}-results.xlsx`)} />
          </section>
        )}
      </main>
    </div>
  )
}

export default App
