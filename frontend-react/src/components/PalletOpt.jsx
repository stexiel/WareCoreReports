import { useState } from 'react'
import ExcelJS from 'exceljs'

function PalletOpt() {
  const [mode, setMode] = useState('file')
  const [sourceFile, setSourceFile] = useState(null)
  const [templateFile, setTemplateFile] = useState(null)
  const [sourceStatus, setSourceStatus] = useState('')
  const [templateStatus, setTemplateStatus] = useState('')
  const [manualRows, setManualRows] = useState([])
  const [manualInput, setManualInput] = useState({
    container: '',
    sku: '',
    prodDate: '',
    expDate: '',
    batch: '',
    qty: ''
  })
  const [reference1, setReference1] = useState('')
  const [generateStatus, setGenerateStatus] = useState('')
  const [showPreview, setShowPreview] = useState(false)
  const [previewData, setPreviewData] = useState([])

  const switchMode = (newMode) => {
    setMode(newMode)
  }

  const handleSourceFile = (e) => {
    const file = e.target.files[0]
    if (file) {
      setSourceFile(file)
      setSourceStatus(`Выбран: ${file.name}`)
      loadSourceData(file)
    }
  }

  const loadSourceData = async (file) => {
    try {
      const reader = new FileReader()
      reader.onload = async (e) => {
        const buffer = e.target.result
        const wb = new ExcelJS.Workbook()
        await wb.xlsx.load(buffer)
        const ws = wb.worksheets[0]
        
        const data = []
        ws.eachRow((row, rowNumber) => {
          if (rowNumber > 1) {
            const values = row.values
            data.push(values)
          }
        })
        
        setPreviewData(data)
        setShowPreview(true)
      }
      reader.readAsArrayBuffer(file)
    } catch (error) {
      setSourceStatus('Ошибка загрузки файла')
    }
  }

  const handleTemplateFile = (e) => {
    const file = e.target.files[0]
    if (file) {
      setTemplateFile(file)
      setTemplateStatus(`Выбран: ${file.name}`)
    }
  }

  const addManualRow = () => {
    if (!manualInput.container) {
      alert('Введите контейнер')
      return
    }
    setManualRows([...manualRows, { 
      container: manualInput.container,
      art: manualInput.sku,
      prodDate: manualInput.prodDate,
      expDate: manualInput.expDate,
      batch: manualInput.batch,
      qty: parseFloat(manualInput.qty) || 0,
      id: Date.now()
    }])
    setManualInput({
      container: '',
      sku: '',
      prodDate: '',
      expDate: '',
      batch: '',
      qty: ''
    })
  }

  const clearManualRows = () => {
    setManualRows([])
    setShowPreview(false)
    setPreviewData([])
  }

  const removeManualRow = (id) => {
    setManualRows(manualRows.filter(r => r.id !== id))
  }

  const generateASN = async () => {
    setGenerateStatus('Обработка...')
    
    try {
      let dataSource = null
      if (mode === 'file') {
        if (!sourceFile) {
          setGenerateStatus('❌ Загрузите файл 1 (данные склада)')
          return
        }
        dataSource = previewData
      } else {
        if (!manualRows.length) {
          setGenerateStatus('❌ Добавьте хотя бы одну строку вручную')
          return
        }
        dataSource = manualRows
      }

      const asnRef1 = reference1.trim()
      let outputName = asnRef1 || 'asn_output'
      if (!outputName.toLowerCase().endsWith('.xlsx')) {
        outputName += '.xlsx'
      }

      // Создаем простой Excel файл
      const workbook = new ExcelJS.Workbook()
      const worksheet = workbook.addWorksheet('ASN')

      const headers = ['Контейнер', 'Артикул', 'Дата пр-ва', 'Срок годн.', 'Партия', 'Количество']
      worksheet.addRow(headers)

      if (mode === 'file') {
        dataSource.forEach(row => {
          worksheet.addRow(row)
        })
      } else {
        dataSource.forEach(row => {
          worksheet.addRow([row.container, row.art, row.prodDate, row.expDate, row.batch, row.qty])
        })
      }

      const buffer = await workbook.xlsx.writeBuffer()
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
      const url = window.URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = outputName
      a.click()
      window.URL.revokeObjectURL(url)
      
      setGenerateStatus('Файл успешно скачан!')
      setTimeout(() => setGenerateStatus(''), 3000)
    } catch (error) {
      setGenerateStatus('Ошибка: ' + error.message)
    }
  }

  const clearASN = () => {
    setSourceFile(null)
    setTemplateFile(null)
    setSourceStatus('')
    setTemplateStatus('')
    setManualRows([])
    setReference1('')
    setGenerateStatus('')
    setShowPreview(false)
    setPreviewData([])
  }

  return (
    <div className="card">
      <div className="card-header">
        <div>
          <h2 className="card-title">Создание приемки</h2>
          <p className="card-subtitle">Загрузите два файла и получите заполненный шаблон</p>
        </div>
      </div>

      <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: '10px', padding: '16px', marginBottom: '16px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '14px' }}>
          <div style={{ width: '4px', height: '18px', background: 'var(--primary)', borderRadius: '2px' }}></div>
          <h3 style={{ fontSize: '14px', fontWeight: 700, margin: 0, color: 'var(--text-primary)' }}>Шаг 1: Режим ввода данных</h3>
        </div>
        <div style={{ display: 'flex', gap: '12px', marginBottom: '16px' }}>
          <button 
            onClick={() => switchMode('file')} 
            className={`btn ${mode === 'file' ? 'btn-primary' : 'btn-secondary'}`}
            style={{ padding: '8px 16px', fontSize: '13px' }}
          >
            Загрузка файлов
          </button>
          <button 
            onClick={() => switchMode('manual')} 
            className={`btn ${mode === 'manual' ? 'btn-primary' : 'btn-secondary'}`}
            style={{ padding: '8px 16px', fontSize: '13px' }}
          >
            Ручной ввод
          </button>
        </div>

        {mode === 'file' && (
          <div style={{ display: 'grid', gridTemplateColumns: '1fr', gap: '16px' }}>
            <div className="form-group" style={{ margin: 0 }}>
              <label className="form-label" style={{ fontSize: '12px' }}>📦 Файл — Остатки на складе</label>
              <input 
                type="file" 
                accept=".xlsx,.xls" 
                onChange={handleSourceFile}
                className="form-input" 
                style={{ padding: '8px', cursor: 'pointer', fontSize: '12px' }}
              />
              <div style={{ fontSize: '12px', color: 'var(--text-secondary)', marginTop: '4px' }}>{sourceStatus}</div>
            </div>
            <div className="form-group" style={{ margin: 0 }}>
              <label className="form-label" style={{ fontSize: '12px' }}>📋 Шаблон ASN (если не загрузился автоматически)</label>
              <input 
                type="file" 
                accept=".xlsx,.xls" 
                onChange={handleTemplateFile}
                className="form-input" 
                style={{ padding: '8px', cursor: 'pointer', fontSize: '12px' }}
              />
              <div style={{ fontSize: '12px', color: 'var(--text-secondary)', marginTop: '4px' }}>{templateStatus}</div>
            </div>
          </div>
        )}

        {mode === 'manual' && (
          <div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr 1fr 1fr 1fr', gap: '10px', marginBottom: '12px', alignItems: 'center' }}>
              <label className="form-label" style={{ fontSize: '12px', margin: 0 }}>Контейнер</label>
              <label className="form-label" style={{ fontSize: '12px', margin: 0 }}>Артикул</label>
              <label className="form-label" style={{ fontSize: '12px', margin: 0 }}>Дата пр-ва</label>
              <label className="form-label" style={{ fontSize: '12px', margin: 0 }}>Срок годн.</label>
              <label className="form-label" style={{ fontSize: '12px', margin: 0 }}>Партия</label>
              <label className="form-label" style={{ fontSize: '12px', margin: 0 }}>Кол-во</label>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr 1fr 1fr 1fr', gap: '10px', marginBottom: '10px' }}>
              <input 
                type="text" 
                value={manualInput.container}
                onChange={(e) => setManualInput({ ...manualInput, container: e.target.value })}
                className="form-input" 
                placeholder="L01..." 
                style={{ fontSize: '12px', padding: '8px' }}
              />
              <input 
                type="text" 
                value={manualInput.sku}
                onChange={(e) => setManualInput({ ...manualInput, sku: e.target.value })}
                className="form-input" 
                placeholder="арт" 
                style={{ fontSize: '12px', padding: '8px' }}
              />
              <input 
                type="text" 
                value={manualInput.prodDate}
                onChange={(e) => setManualInput({ ...manualInput, prodDate: e.target.value })}
                className="form-input" 
                placeholder="2026-01-01" 
                style={{ fontSize: '12px', padding: '8px' }}
              />
              <input 
                type="text" 
                value={manualInput.expDate}
                onChange={(e) => setManualInput({ ...manualInput, expDate: e.target.value })}
                className="form-input" 
                placeholder="2027-01-01" 
                style={{ fontSize: '12px', padding: '8px' }}
              />
              <input 
                type="text" 
                value={manualInput.batch}
                onChange={(e) => setManualInput({ ...manualInput, batch: e.target.value })}
                className="form-input" 
                placeholder="партия" 
                style={{ fontSize: '12px', padding: '8px' }}
              />
              <input 
                type="number" 
                value={manualInput.qty}
                onChange={(e) => setManualInput({ ...manualInput, qty: e.target.value })}
                className="form-input" 
                placeholder="100" 
                style={{ fontSize: '12px', padding: '8px' }}
              />
            </div>
            <div style={{ display: 'flex', gap: '8px' }}>
              <button onClick={addManualRow} className="btn btn-primary" style={{ padding: '8px 16px', fontSize: '12px' }}>
                + Добавить строку
              </button>
              <button onClick={clearManualRows} className="btn btn-secondary" style={{ padding: '8px 16px', fontSize: '12px' }}>
                Очистить все
              </button>
            </div>
            
            {manualRows.length > 0 && (
              <div id="manualRowsPreview" style={{ marginTop: '12px', maxHeight: '200px', overflowY: 'auto', border: '1px solid var(--border)', borderRadius: '8px', fontSize: '12px' }}>
                {manualRows.map((r, i) => (
                  <div key={r.id} style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr 1fr 1fr 1fr', gap: '8px', padding: '8px', borderBottom: '1px solid var(--border)', fontSize: '12px' }}>
                    <span>{r.container}</span>
                    <span>{r.art || '-'}</span>
                    <span>{r.prodDate || '-'}</span>
                    <span>{r.expDate || '-'}</span>
                    <span>{r.batch || '-'}</span>
                    <span>{r.qty}</span>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}
      </div>

      <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: '10px', padding: '16px', marginBottom: '16px', textAlign: 'center' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '14px', justifyContent: 'center' }}>
          <div style={{ width: '4px', height: '18px', background: 'var(--primary)', borderRadius: '2px' }}></div>
          <h3 style={{ fontSize: '14px', fontWeight: 700, margin: 0, color: 'var(--text-primary)' }}>Шаг 2: Параметры</h3>
        </div>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr', gap: '12px', maxWidth: '400px', margin: '0 auto' }}>
          <div className="form-group" style={{ margin: 0 }}>
            <label className="form-label" style={{ fontSize: '12px' }}>Номер заказа / Название файла</label>
            <input 
              type="text" 
              value={reference1}
              onChange={(e) => setReference1(e.target.value)}
              className="form-input" 
              placeholder="8422" 
              style={{ fontSize: '12px', padding: '8px' }}
            />
          </div>
        </div>
      </div>

      <div style={{ display: 'flex', gap: '10px', alignItems: 'center', flexWrap: 'wrap', justifyContent: 'center' }}>
        <button onClick={generateASN} className="btn btn-primary" style={{ padding: '11px 22px', fontSize: '12px' }}>
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ marginRight: '6px' }}>
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
            <polyline points="14 2 14 8 20 8"/>
            <path d="M8 13h8"/>
            <path d="M8 17h8"/>
          </svg>
          Генерировать и скачать
        </button>
        <button onClick={clearASN} className="btn btn-secondary" style={{ padding: '11px 18px', fontSize: '12px' }}>
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ marginRight: '5px' }}>
            <polyline points="1 4 1 10 7 10"/>
            <path d="M3.51 15a9 9 0 1 0 .49-3.5"/>
          </svg>
          Очистить
        </button>
        <div id="asnGenerateStatus" style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>{generateStatus}</div>
      </div>

      {showPreview && (
        <div id="asnPreview" style={{ display: 'block', marginTop: '16px', textAlign: 'center' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '10px', justifyContent: 'center' }}>
            <div style={{ width: '4px', height: '18px', background: '#22c55e', borderRadius: '2px' }}></div>
            <h3 style={{ fontSize: '12px', fontWeight: 700, margin: 0, color: 'var(--text-primary)' }}>Превью данных</h3>
            <span id="asnRowCount" style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>{previewData.length} строк</span>
          </div>
          <div style={{ overflowX: 'auto', maxHeight: '300px', overflowY: 'auto', border: '1px solid var(--border)', borderRadius: '8px', margin: '0 auto' }}>
            <table id="asnPreviewTable" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px' }}>
              <thead id="asnPreviewHead" style={{ background: 'var(--surface)', position: 'sticky', top: 0 }}>
                <tr>
                  {previewData.length > 0 && previewData[0].map((_, i) => (
                    <th key={i} style={{ padding: '8px', border: '1px solid var(--border)' }}>Кол {i + 1}</th>
                  ))}
                </tr>
              </thead>
              <tbody id="asnPreviewBody">
                {previewData.map((row, i) => (
                  <tr key={i}>
                    {row.map((cell, j) => (
                      <td key={j} style={{ padding: '8px', border: '1px solid var(--border)' }}>{cell}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  )
}

export default PalletOpt
