import { useState, useEffect } from 'react'

function Receipt({ user }) {
  const [documents, setDocuments] = useState([])
  const [loading, setLoading] = useState(true)
  const [status, setStatus] = useState('Подключение к БД...')
  const [detailDoc, setDetailDoc] = useState(null)
  const [detailData, setDetailData] = useState([])
  const [detailFilter, setDetailFilter] = useState('all')
  const [showModal, setShowModal] = useState(false)

  const getApiBase = () => 'http://localhost:3000'
  const getAuthToken = () => localStorage.getItem('authToken')

  const fetchDocuments = async () => {
    setLoading(true)
    setStatus('Загрузка данных...')
    try {
      const token = getAuthToken()
      const response = await fetch(`${getApiBase()}/api/receipt-documents`, {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json'
        }
      })
      
      if (!response.ok) {
        throw new Error('Ошибка загрузки документов')
      }
      
      const data = await response.json()
      console.log('Documents from API:', data)
      setDocuments(data)
      setStatus('')
    } catch (err) {
      console.error('Error fetching documents:', err)
      setStatus('Ошибка: ' + err.message)
    } finally {
      setLoading(false)
    }
  }

  const fetchDocumentDetails = async (docNum) => {
    try {
      const token = getAuthToken()
      const response = await fetch(`${getApiBase()}/api/receipt-detail?docNumber=${encodeURIComponent(docNum)}`, {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json'
        }
      })
      
      if (!response.ok) {
        throw new Error('Ошибка загрузки деталей')
      }
      
      const data = await response.json()
      console.log('Details from API:', data)
      setDetailData(data)
      setDetailDoc(docNum)
      setShowModal(true)
    } catch (err) {
      console.error('Error fetching details:', err)
      alert(err.message)
    }
  }

  const closeModal = () => {
    setShowModal(false)
    setDetailDoc(null)
    setDetailData([])
  }

  const filteredDetailData = detailFilter === 'all' 
    ? detailData 
    : detailData.filter(item => item.palletState === detailFilter)

  useEffect(() => {
    fetchDocuments()
  }, [])

  return (
    <div className="card">
      <div className="card-header">
        <div>
          <h2 className="card-title">Приемка</h2>
          <p className="card-subtitle">Мониторинг документов приемки в реальном времени</p>
        </div>
        <button onClick={fetchDocuments} className="btn btn-secondary">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/>
            <polyline points="9 11 12 14 22 4"/>
          </svg>
          Обновить
        </button>
      </div>

      <div style={{ padding: '20px' }}>
        <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: '10px', padding: '16px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '14px' }}>
            <div style={{ width: '4px', height: '18px', background: 'var(--primary)', borderRadius: '2px' }}></div>
            <h3 style={{ fontSize: '14px', fontWeight: 700, margin: 0, color: 'var(--text-primary)' }}>Документы приемки</h3>
            <span style={{ fontSize: '11px', color: 'var(--text-secondary)', marginLeft: 'auto' }}>{status}</span>
          </div>
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
              <thead>
                <tr style={{ background: 'var(--border)' }}>
                  <th style={{ padding: '10px', textAlign: 'left', border: '1px solid var(--border)' }}>Номер документа</th>
                  <th style={{ padding: '10px', textAlign: 'left', border: '1px solid var(--border)' }}>Время начала</th>
                  <th style={{ padding: '10px', textAlign: 'left', border: '1px solid var(--border)' }}>Время окончания</th>
                </tr>
              </thead>
              <tbody>
                {loading ? (
                  <tr>
                    <td colSpan="3" style={{ padding: '20px', textAlign: 'center', color: 'var(--text-secondary)' }}>
                      Загрузка данных...
                    </td>
                  </tr>
                ) : documents.length === 0 ? (
                  <tr>
                    <td colSpan="3" style={{ padding: '20px', textAlign: 'center', color: 'var(--text-secondary)' }}>
                      Нет данных
                    </td>
                  </tr>
                ) : (
                  documents.map((doc, index) => (
                    <tr 
                      key={index} 
                      style={{ cursor: 'pointer' }}
                      onClick={() => fetchDocumentDetails(doc.docNumber || doc.docNum)}
                      onMouseEnter={(e) => e.currentTarget.style.background = 'var(--background)'}
                      onMouseLeave={(e) => e.currentTarget.style.background = 'transparent'}
                    >
                      <td style={{ padding: '10px', border: '1px solid var(--border)' }}>{doc.docNumber || doc.docNum}</td>
                      <td style={{ padding: '10px', border: '1px solid var(--border)' }}>{doc.startTime}</td>
                      <td style={{ padding: '10px', border: '1px solid var(--border)' }}>{doc.endTime}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      {showModal && (
        <div style={{ 
          display: 'flex', 
          position: 'fixed', 
          top: 0, 
          left: 0, 
          width: '100%', 
          height: '100%', 
          background: 'rgba(0,0,0,0.5)', 
          zIndex: 1000, 
          alignItems: 'center', 
          justifyContent: 'center' 
        }}>
          <div style={{ 
            background: 'var(--surface)', 
            borderRadius: '10px', 
            maxWidth: '90%', 
            maxHeight: '90%', 
            overflow: 'auto', 
            padding: '20px' 
          }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
              <h3 style={{ fontSize: '16px', fontWeight: 700, margin: 0, color: 'var(--text-primary)' }}>
                Детали документа: {detailDoc}
              </h3>
              <button 
                onClick={closeModal} 
                style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: '24px', color: 'var(--text-secondary)' }}
              >
                ×
              </button>
            </div>

            <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: '10px', padding: '12px', marginBottom: '16px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '10px' }}>
                <div style={{ width: '4px', height: '16px', background: 'var(--primary)', borderRadius: '2px' }}></div>
                <h4 style={{ fontSize: '13px', fontWeight: 700, margin: 0, color: 'var(--text-primary)' }}>Фильтр по состоянию паллеты</h4>
              </div>
              <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
                <label style={{ display: 'flex', alignItems: 'center', gap: '6px', cursor: 'pointer', fontSize: '12px' }}>
                  <input 
                    type="radio" 
                    name="palletStateFilter" 
                    value="all" 
                    checked={detailFilter === 'all'} 
                    onChange={(e) => setDetailFilter(e.target.value)}
                  />
                  <span>Все</span>
                </label>
                <label style={{ display: 'flex', alignItems: 'center', gap: '6px', cursor: 'pointer', fontSize: '12px' }}>
                  <input 
                    type="radio" 
                    name="palletStateFilter" 
                    value="Импорт" 
                    checked={detailFilter === 'Импорт'} 
                    onChange={(e) => setDetailFilter(e.target.value)}
                  />
                  <span>Импорт</span>
                </label>
                <label style={{ display: 'flex', alignItems: 'center', gap: '6px', cursor: 'pointer', fontSize: '12px' }}>
                  <input 
                    type="radio" 
                    name="palletStateFilter" 
                    value="Получен" 
                    checked={detailFilter === 'Получен'} 
                    onChange={(e) => setDetailFilter(e.target.value)}
                  />
                  <span>Получен</span>
                </label>
                <label style={{ display: 'flex', alignItems: 'center', gap: '6px', cursor: 'pointer', fontSize: '12px' }}>
                  <input 
                    type="radio" 
                    name="palletStateFilter" 
                    value="Выставлено" 
                    checked={detailFilter === 'Выставлено'} 
                    onChange={(e) => setDetailFilter(e.target.value)}
                  />
                  <span>Выставлено</span>
                </label>
              </div>
            </div>

            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px' }}>
                <thead>
                  <tr style={{ background: 'var(--border)' }}>
                    <th style={{ padding: '8px', textAlign: 'left', border: '1px solid var(--border)' }}>Номер контейнера</th>
                    <th style={{ padding: '8px', textAlign: 'left', border: '1px solid var(--border)' }}>Количество</th>
                    <th style={{ padding: '8px', textAlign: 'left', border: '1px solid var(--border)' }}>Номер внешнего документа</th>
                    <th style={{ padding: '8px', textAlign: 'left', border: '1px solid var(--border)' }}>Артикул</th>
                    <th style={{ padding: '8px', textAlign: 'left', border: '1px solid var(--border)' }}>Номер партии</th>
                    <th style={{ padding: '8px', textAlign: 'left', border: '1px solid var(--border)' }}>Состояние паллеты</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredDetailData.map((item, index) => (
                    <tr key={index}>
                      <td style={{ padding: '8px', border: '1px solid var(--border)' }}>{item.containerNumber || item.container}</td>
                      <td style={{ padding: '8px', border: '1px solid var(--border)' }}>{item.quantity || item.qty}</td>
                      <td style={{ padding: '8px', border: '1px solid var(--border)' }}>{item.externalDocNumber || item.externalDocNum}</td>
                      <td style={{ padding: '8px', border: '1px solid var(--border)' }}>{item.sku}</td>
                      <td style={{ padding: '8px', border: '1px solid var(--border)' }}>{item.batchNumber || item.batch}</td>
                      <td style={{ padding: '8px', border: '1px solid var(--border)' }}>{item.palletState}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

export default Receipt
