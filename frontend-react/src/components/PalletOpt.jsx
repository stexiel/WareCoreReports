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
  const [lotAtt03, setLotAtt03] = useState('')
  const [generateStatus, setGenerateStatus] = useState('')
  const [showPreview, setShowPreview] = useState(false)
  const [previewData, setPreviewData] = useState([])
  const [asnSourceData, setAsnSourceData] = useState([])
  const [asnTemplateWb, setAsnTemplateWb] = useState(null)
  const [asnTemplateLoaded, setAsnTemplateLoaded] = useState(false)

  const fmtDate = (val) => {
    if (!val) return ''
    if (val instanceof Date) {
      const y = val.getUTCFullYear()
      const m = String(val.getUTCMonth() + 1).padStart(2, '0')
      const d = String(val.getUTCDate()).padStart(2, '0')
      return `${y}-${m}-${d}`
    }
    const s = val.toString().trim()
    const m1 = s.match(/(\d{4})[.\-](\d{2})[.\-](\d{2})/)
    if (m1) return `${m1[1]}-${m1[2]}-${m1[3]}`
    const m2 = s.match(/(\d{2})\.(\d{2})\.(\d{4})/)
    if (m2) return `${m2[3]}-${m2[2]}-${m2[1]}`
    return s
  }

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

        const rows = []
        ws.eachRow({ includeEmpty: true }, (row) => {
          const values = row.values.slice(1).map(v => (v === null || v === undefined) ? '' : v)
          rows.push(values)
        })

        // Ищем строку с заголовками (где есть "Контейнер" и "Артикул")
        let headerRowIdx = -1
        const colIdx = {}
        for (let ri = 0; ri < rows.length; ri++) {
          const r = rows[ri].map(v => (v || '').toString().trim())
          if (r.some(v => v.includes('Контейнер')) && r.some(v => v.includes('Артикул'))) {
            headerRowIdx = ri
            r.forEach((h, ci) => {
              if (h.includes('Артикул')) colIdx.artIdx = ci
              if (h === 'Контейнер') colIdx.containerIdx = ci
              if (h.includes('Дата производства')) colIdx.prodDateIdx = ci
              if (h.includes('Срок годности')) colIdx.expDateIdx = ci
              if (h.includes('Партия номенклатуры') && !h.includes('Статус') && !h.includes('Срок') && !h.includes('Дата')) colIdx.batchIdx = ci
              if ((h.includes('Кол') || h.includes('КОЛ') || h.includes('qty') || h.includes('QTY') || h.includes('Количество')) && !h.includes('план') && !h.includes('ERP') && !h.includes('План') && colIdx.qtyIdx === undefined) colIdx.qtyIdx = ci
            })
            break
          }
        }

        if (headerRowIdx < 0) {
          setSourceStatus('❌ Не найдена строка заголовков (нужны Контейнер и Артикул)')
          return
        }

        const data = []
        for (let ri = headerRowIdx + 1; ri < rows.length; ri++) {
          const r = rows[ri]
          const container = colIdx.containerIdx !== undefined ? (r[colIdx.containerIdx] || '').toString().trim() : ''
          if (!container || container.toLowerCase().includes('итого') || container.length < 5) continue
          const art = colIdx.artIdx !== undefined ? (r[colIdx.artIdx] || '').toString().trim() : ''
          const prodDate = colIdx.prodDateIdx !== undefined ? fmtDate(r[colIdx.prodDateIdx]) : ''
          const expDate = colIdx.expDateIdx !== undefined ? fmtDate(r[colIdx.expDateIdx]) : ''
          const batch = colIdx.batchIdx !== undefined ? (r[colIdx.batchIdx] || '').toString().trim() : ''
          let qty = colIdx.qtyIdx !== undefined ? r[colIdx.qtyIdx] : 0
          if (qty === null || qty === undefined || qty === '') qty = 0
          else if (typeof qty === 'number') qty = Math.round(qty)
          else qty = Math.round(parseFloat(qty.toString().replace(',', '.')) || 0)
          if (!art && !prodDate) continue
          data.push({ container, art, prodDate, expDate, batch, qty })
        }

        setAsnSourceData(data)
        setPreviewData(data)
        setShowPreview(true)
        setSourceStatus(`✅ Загружено ${data.length} строк`)
      }
      reader.readAsArrayBuffer(file)
    } catch (error) {
      setSourceStatus('❌ Ошибка загрузки файла: ' + error.message)
    }
  }

  const handleTemplateFile = async (e) => {
    const file = e.target.files[0]
    if (file) {
      setTemplateFile(file)
      setTemplateStatus(`Выбран: ${file.name}`)
      
      try {
        const reader = new FileReader()
        reader.onload = async (e) => {
          const buffer = e.target.result
          const wb = new ExcelJS.Workbook()
          await wb.xlsx.load(buffer)
          setAsnTemplateWb({ wb, filename: file.name })
          setAsnTemplateLoaded(true)
        }
        reader.readAsArrayBuffer(file)
      } catch (error) {
        setTemplateStatus('Ошибка загрузки шаблона')
      }
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
        if (!asnSourceData.length) {
          setGenerateStatus('❌ Загрузите файл 1 (данные склада)')
          return
        }
        dataSource = asnSourceData
      } else {
        if (!manualRows.length) {
          setGenerateStatus('❌ Добавьте хотя бы одну строку вручную')
          return
        }
        dataSource = manualRows
      }

      // Загружаем шаблон если еще не загружен
      if (!asnTemplateLoaded && !asnTemplateWb) {
        setGenerateStatus('❌ Загрузите шаблон ASN')
        return
      }

      const asnRef1 = reference1.trim()
      let outputName = asnRef1 || 'asn_output'
      if (!outputName.toLowerCase().endsWith('.xlsx')) {
        outputName += '.xlsx'
      }

      // Используем шаблон
      const wb = asnTemplateWb.wb
      const ws = wb.worksheets[0]

      // Найти строку заголовков
      let headerRowNum = 2
      const headerRow = ws.getRow(headerRowNum)
      const fieldMap = {}
      headerRow.eachCell((cell, colNum) => {
        const v = (cell.value || '').toString().trim()
        if (v) fieldMap[v] = colNum
      })

      // Если не нашли ключевые поля — попробуем строку 1
      if (!fieldMap['containerId'] && !fieldMap['warehouseId']) {
        headerRowNum = 1
        const r1 = ws.getRow(1)
        r1.eachCell((cell, colNum) => {
          const v = (cell.value || '').toString().trim()
          if (v) fieldMap[v] = colNum
        })
      }

      // Удаляем все строки данных
      const lastRow = ws.lastRow ? ws.lastRow.number : headerRowNum + 1
      for (let r = lastRow; r > headerRowNum; r--) {
        ws.spliceRows(r, 1)
      }

      // Фиксированные значения
      const fixed = {
        warehouseId: 'WH01',
        asnType: 'ProductIn',
        customerId: 'SHIN-LINE',
        packUom: 'EA',
        lotAtt08: 'GQ',
        lotAtt05: 'GENERAL',
      }

      // Заполняем строки
      dataSource.forEach((item, idx) => {
        const rowNum = headerRowNum + 1 + idx
        const row = ws.getRow(rowNum)

        const setValue = (field, val) => {
          if (fieldMap[field]) {
            const cell = row.getCell(fieldMap[field])
            cell.value = val
            cell.alignment = { horizontal: 'center', vertical: 'middle' }
            if (rowNum > 2) {
              cell.font = { name: 'Calibri', size: 12 }
            }
            cell.border = {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' }
            }
          }
        }

        // Фиксированные
        Object.entries(fixed).forEach(([k, v]) => setValue(k, v))

        // Из параметров
        setValue('asnReference1', asnRef1)
        setValue('lotAtt03', lotAtt03.trim())

        // Из данных (одинаковая структура для файла и ручного ввода)
        setValue('sku', item.art)
        setValue('containerId', item.container)
        setValue('expectedQty', typeof item.qty === 'number' ? item.qty : Number(item.qty) || 0)
        setValue('lotAtt01', item.prodDate)
        setValue('lotAtt02', item.expDate)
        setValue('lotAtt04', item.batch)

        // Форматирование дат как текст
        if (fieldMap['lotAtt01']) {
          const c = row.getCell(fieldMap['lotAtt01'])
          c.numFmt = '@'
        }
        if (fieldMap['lotAtt02']) {
          const c = row.getCell(fieldMap['lotAtt02'])
          c.numFmt = '@'
        }

        row.commit()
      })

      const buffer = await wb.xlsx.writeBuffer()
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
    setLotAtt03('')
    setGenerateStatus('')
    setShowPreview(false)
    setPreviewData([])
    setAsnSourceData([])
    setAsnTemplateWb(null)
    setAsnTemplateLoaded(false)
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
              <label className="form-label" style={{ fontSize: '12px' }}>📋 Шаблон ASN</label>
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
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px', maxWidth: '600px', margin: '0 auto' }}>
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
          <div className="form-group" style={{ margin: 0 }}>
            <label className="form-label" style={{ fontSize: '12px' }}>Дата приемки</label>
            <input 
              type="text" 
              value={lotAtt03}
              onChange={(e) => setLotAtt03(e.target.value)}
              className="form-input" 
              placeholder="2026-04-26" 
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
                  {['Контейнер', 'Артикул', 'Дата производства', 'Срок годности', 'Партия', 'Кол-во'].map((c, i) => (
                    <th key={i} style={{ padding: '8px', border: '1px solid var(--border)', whiteSpace: 'nowrap' }}>{c}</th>
                  ))}
                </tr>
              </thead>
              <tbody id="asnPreviewBody">
                {previewData.map((row, i) => (
                  <tr key={i}>
                    <td style={{ padding: '8px', border: '1px solid var(--border)' }}>{row.container}</td>
                    <td style={{ padding: '8px', border: '1px solid var(--border)' }}>{row.art || '-'}</td>
                    <td style={{ padding: '8px', border: '1px solid var(--border)' }}>{row.prodDate || '-'}</td>
                    <td style={{ padding: '8px', border: '1px solid var(--border)' }}>{row.expDate || '-'}</td>
                    <td style={{ padding: '8px', border: '1px solid var(--border)' }}>{row.batch || '-'}</td>
                    <td style={{ padding: '8px', border: '1px solid var(--border)', textAlign: 'right' }}>{row.qty}</td>
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
