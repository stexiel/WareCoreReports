import { useState, useEffect } from 'react'
import ExcelJS from 'exceljs'
import { apiFetch, checkBackendHealth } from '../api'

const randomInRange = (min, max) => Math.floor(Math.random() * (max - min + 1)) + min

const QUICK_PRESETS = [
  { key: 'pallet_check', label: 'Проверка выставления паллет по ячейкам', action: 'Проверка на выставление паллеты по ячейкам', description: 'Проверка по сообщению кладовщика', min: 2, max: 4 },
]

function OperatorReport() {
  const [logs, setLogs] = useState([])
  const [loading, setLoading] = useState(true)
  const [backendAvailable, setBackendAvailable] = useState(null)
  const [shift, setShift] = useState('currentDay')
  const [shiftDate, setShiftDate] = useState('')
  const [dbFrom, setDbFrom] = useState('')
  const [dbTo, setDbTo] = useState('')
  const [showModal, setShowModal] = useState(false)
  const [editingLog, setEditingLog] = useState(null)
  const [filterType, setFilterType] = useState('all') // all, errors, shipments, receipts
  const [formData, setFormData] = useState({
    operatorName: '',
    action: '',
    description: '',
    timestamp: '',
    shift: ''
  })

  const updateDates = (shiftType) => {
    const now = new Date()
    let from, to

    switch (shiftType) {
      case 'all':
        from = new Date(2000, 0, 1)
        to = new Date(2030, 11, 31)
        break
      case 'currentDay':
        from = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 8, 0, 0)
        to = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 20, 0, 0)
        break
      case 'currentNight':
        from = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 20, 0, 0)
        to = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1, 8, 0, 0)
        break
      case 'lastDay':
        from = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1, 8, 0, 0)
        to = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1, 20, 0, 0)
        break
      case 'lastNight':
        from = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1, 20, 0, 0)
        to = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 8, 0, 0)
        break
    }

    setDbFrom(formatDateTimeLocal(from))
    setDbTo(formatDateTimeLocal(to))
    setShiftDate(shiftType === 'all' ? 'Все записи' : formatDate(from))
  }

  const formatDateTimeLocal = (date) => {
    const offset = date.getTimezoneOffset() * 60000
    const localISOTime = (new Date(date - offset)).toISOString().slice(0, 16)
    return localISOTime
  }

  const formatDate = (date) => {
    const year = date.getFullYear()
    const month = String(date.getMonth() + 1).padStart(2, '0')
    const day = String(date.getDate()).padStart(2, '0')
    return `${year}-${month}-${day}`
  }

  const handleShiftChange = (e) => {
    setShift(e.target.value)
    updateDates(e.target.value)
  }

  const handleDateChange = (e) => {
    setShiftDate(e.target.value)
    if (shift === 'currentDay' || shift === 'lastDay') {
      const date = new Date(e.target.value + 'T08:00')
      const to = new Date(e.target.value + 'T20:00')
      setDbFrom(formatDateTimeLocal(date))
      setDbTo(formatDateTimeLocal(to))
    } else {
      const date = new Date(e.target.value + 'T20:00')
      const nextDay = new Date(date)
      nextDay.setDate(nextDay.getDate() + 1)
      nextDay.setHours(8, 0, 0)
      setDbFrom(formatDateTimeLocal(date))
      setDbTo(formatDateTimeLocal(nextDay))
    }
  }

  const loadLogs = async () => {
    setLoading(true)
    try {
      const response = await apiFetch(`/api/operator-logs?from=${encodeURIComponent(dbFrom)}&to=${encodeURIComponent(dbTo)}`)
      if (!response.ok) throw new Error('Ошибка загрузки логов')
      const data = await response.json()
      setLogs(data.rows || [])
    } catch (err) {
      console.error('Error loading logs:', err)
    } finally {
      setLoading(false)
    }
  }

  const downloadExcel = async () => {
    try {
      const workbook = new ExcelJS.Workbook()
      const worksheet = workbook.addWorksheet('Отчет операторов')

      // Информация о смене, операторе и дате в верхних строках
      const shiftText = shift === 'currentDay' || shift === 'lastDay' ? 'День' : 'Ночь'
      worksheet.addRow(['Смена:', shiftText])
      worksheet.addRow(['Оператор:', 'Не указан'])
      worksheet.addRow(['Дата:', shiftDate])
      worksheet.addRow([]) // Пустая строка

      // Заголовки таблицы
      const headers = ['Дата', 'Время', 'Действие', 'Описание', 'Общее время (мин)']
      const headerRow = worksheet.addRow(headers)

      // Стилизация заголовков
      headerRow.eachCell((cell) => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: '4472C4' }
        }
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        }
      })

      // Цвета строк по типу действия
      const getRowFillColor = (action) => {
        if (action.includes('Ошибка')) return 'FFE2E2'
        if (action.includes('Отгрузка')) return 'E2F7E2'
        if (action.includes('Приемка')) return 'E2ECFB'
        return 'FFFFFF'
      }

      // Данные
      logs.forEach((log) => {
        const date = new Date(log.timestamp)
        const row = worksheet.addRow([
          date.toLocaleDateString('ru-RU'),
          date.toLocaleTimeString('ru-RU'),
          log.action,
          log.description,
          log.total_minutes || ''
        ])

        const fillColor = getRowFillColor(log.action)

        // Стилизация строк данных
        row.eachCell((cell) => {
          cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true }
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
          }
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: fillColor }
          }
        })
      })

      // Итоговая строка с общим временем
      const totalMinutes = logs.reduce((sum, log) => sum + (log.total_minutes || 0), 0)
      const summaryRow = worksheet.addRow(['', '', '', 'Общее время от начала до конца смены:', totalMinutes + ' мин'])
      summaryRow.eachCell((cell) => {
        cell.font = { bold: true }
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'C6F6D5' }
        }
      })

      // Автоширина колонок
      worksheet.columns.forEach((column) => {
        let maxLength = 0
        column.eachCell({ includeEmpty: true }, (cell) => {
          const cellValue = cell.value ? cell.value.toString() : ''
          maxLength = Math.max(maxLength, cellValue.length)
        })
        column.width = Math.min(maxLength + 4, 50) // Ограничиваем максимальную ширину
      })

      const buffer = await workbook.xlsx.writeBuffer()
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
      const url = window.URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = `operator_report_${shiftDate}.xlsx`
      a.click()
      window.URL.revokeObjectURL(url)
    } catch (err) {
      console.error('Error downloading Excel:', err)
    }
  }

  const openAddModal = () => {
    setEditingLog(null)
    setFormData({
      operatorName: '',
      action: '',
      description: '',
      timestamp: formatDateTimeLocal(new Date()),
      shift: shift === 'currentDay' || shift === 'lastDay' ? 'день' : 'ночь',
      total_minutes: ''
    })
    setShowModal(true)
  }

  const openEditModal = (log) => {
    setEditingLog(log)
    setFormData({
      operatorName: log.operator_name,
      action: log.action,
      description: log.description,
      timestamp: formatDateTimeLocal(new Date(log.timestamp)),
      shift: log.shift,
      total_minutes: log.total_minutes || ''
    })
    setShowModal(true)
  }

  const getFilteredLogs = () => {
    if (filterType === 'all') return logs
    if (filterType === 'errors') return logs.filter(l => l.action.includes('Ошибка'))
    if (filterType === 'shipments') return logs.filter(l => l.action.includes('Отгрузка'))
    if (filterType === 'receipts') return logs.filter(l => l.action.includes('Приемка'))
    return logs
  }

  const getRowColor = (action) => {
    if (action.includes('Ошибка')) return 'rgba(239, 68, 68, 0.08)'
    if (action.includes('Отгрузка')) return 'rgba(34, 197, 94, 0.08)'
    if (action.includes('Приемка')) return 'rgba(59, 130, 246, 0.08)'
    return 'transparent'
  }

  const applyPreset = (key) => {
    if (!key) return
    const preset = QUICK_PRESETS.find(p => p.key === key)
    if (!preset) return
    setFormData(prev => ({
      ...prev,
      action: preset.action,
      description: preset.description,
      total_minutes: randomInRange(preset.min, preset.max)
    }))
  }

  const saveLog = async () => {
    if (!formData.operatorName || !formData.operatorName.trim()) {
      alert('Укажите имя оператора')
      return
    }
    try {
      const payload = {
        operator_name: formData.operatorName,
        action: formData.action,
        description: formData.description,
        timestamp: formData.timestamp,
        shift: formData.shift,
        total_minutes: formData.total_minutes || null
      }

      let response
      if (editingLog) {
        response = await apiFetch(`/api/operator-logs/${editingLog.id}`, {
          method: 'PUT',
          body: JSON.stringify(payload)
        })
      } else {
        response = await apiFetch('/api/operator-logs', {
          method: 'POST',
          body: JSON.stringify(payload)
        })
      }

      if (!response.ok) throw new Error('Ошибка сохранения')
      
      setShowModal(false)
      loadLogs()
    } catch (err) {
      console.error('Error saving log:', err)
    }
  }

  const deleteLog = async (id) => {
    if (!confirm('Удалить запись?')) return
    try {
      const response = await apiFetch(`/api/operator-logs/${id}`, {
        method: 'DELETE'
      })
      if (!response.ok) throw new Error('Ошибка удаления')
      loadLogs()
    } catch (err) {
      console.error('Error deleting log:', err)
    }
  }

  const recheckBackend = async () => {
    setBackendAvailable(null)
    const ok = await checkBackendHealth()
    setBackendAvailable(ok)
    if (ok && dbFrom && dbTo) {
      loadLogs()
    }
  }

  useEffect(() => {
    updateDates('currentDay')
    recheckBackend()
  }, [])

  useEffect(() => {
    if (backendAvailable && dbFrom && dbTo) {
      loadLogs()
    }
  }, [dbFrom, dbTo])

  if (backendAvailable === false) {
    return (
      <div className="card">
        <div className="card-header">
          <div>
            <h2 className="card-title">Отчет операторов</h2>
          </div>
        </div>
        <div style={{ padding: '20px' }}>
          <div style={{ background: 'rgba(239, 68, 68, 0.1)', border: '1px solid rgba(239, 68, 68, 0.4)', borderRadius: '10px', padding: '20px', textAlign: 'center' }}>
            <div style={{ fontSize: '32px', marginBottom: '8px' }}>⚠️</div>
            <h3 style={{ fontSize: '16px', fontWeight: 700, color: '#ef4444', margin: '0 0 8px' }}>Backend не работает</h3>
            <p style={{ color: 'var(--text-secondary)', margin: '0 0 16px' }}>
              Не удалось подключиться к серверу. Эта страница требует запущенный backend с доступом к базе логов.
            </p>
            <button onClick={recheckBackend} className="btn btn-secondary">Повторить попытку</button>
          </div>
        </div>
      </div>
    )
  }

  if (backendAvailable === null) {
    return (
      <div className="card">
        <div style={{ padding: '40px', textAlign: 'center', color: 'var(--text-secondary)' }}>
          Проверка соединения с backend...
        </div>
      </div>
    )
  }

  return (
    <div className="card">
      <div className="card-header">
        <div>
          <h2 className="card-title">Отчет операторов</h2>
          <p className="card-subtitle">История действий операторов в системе</p>
        </div>
      </div>

      <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: '10px', padding: '16px', marginBottom: '16px' }}>
        <div style={{ display: 'flex', gap: '12px', marginBottom: '16px', flexWrap: 'wrap' }}>
          <div>
            <label className="form-label" style={{ fontSize: '12px' }}>Смена</label>
            <select
              value={shift}
              onChange={handleShiftChange}
              className="form-input"
              style={{ padding: '8px', fontSize: '13px' }}
            >
              <option value="all">Все записи</option>
              <option value="currentDay">Текущая дневная (8:00-20:00)</option>
              <option value="currentNight">Текущая ночная (20:00-8:00)</option>
              <option value="lastDay">Прошлая дневная (8:00-20:00)</option>
              <option value="lastNight">Прошлая ночная (20:00-8:00)</option>
            </select>
          </div>
          <div>
            <label className="form-label" style={{ fontSize: '12px' }}>Дата</label>
            <input 
              type="date" 
              value={shiftDate}
              onChange={handleDateChange}
              className="form-input"
              style={{ padding: '8px', fontSize: '13px' }}
            />
          </div>
          <div>
            <label className="form-label" style={{ fontSize: '12px' }}>Начало</label>
            <input 
              type="datetime-local" 
              value={dbFrom}
              onChange={(e) => setDbFrom(e.target.value)}
              className="form-input"
              style={{ padding: '8px', fontSize: '13px' }}
            />
          </div>
          <div>
            <label className="form-label" style={{ fontSize: '12px' }}>Окончание</label>
            <input 
              type="datetime-local" 
              value={dbTo}
              onChange={(e) => setDbTo(e.target.value)}
              className="form-input"
              style={{ padding: '8px', fontSize: '13px' }}
            />
          </div>
          <div>
            <label className="form-label" style={{ fontSize: '12px' }}>Тип</label>
            <select 
              value={filterType}
              onChange={(e) => setFilterType(e.target.value)}
              className="form-input"
              style={{ padding: '8px', fontSize: '13px' }}
            >
              <option value="all">Все</option>
              <option value="errors">Ошибки</option>
              <option value="shipments">Отгрузки</option>
              <option value="receipts">Приемки</option>
            </select>
          </div>
        </div>

        <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
          <button onClick={loadLogs} className="btn btn-primary" style={{ padding: '8px 16px', fontSize: '13px' }}>
            Загрузить
          </button>
          <button onClick={() => window.location.reload()} className="btn btn-secondary" style={{ padding: '8px 16px', fontSize: '13px' }}>
            Обновить страницу
          </button>
          <button onClick={downloadExcel} className="btn btn-secondary" style={{ padding: '8px 16px', fontSize: '13px' }}>
            Скачать Excel
          </button>
          <button onClick={openAddModal} className="btn btn-primary" style={{ padding: '8px 16px', fontSize: '13px' }}>
            + Добавить запись
          </button>
        </div>
      </div>

      {/* Информация о смене, операторе и дате */}
      <div style={{ display: 'flex', gap: '12px', marginBottom: '16px', flexWrap: 'wrap' }}>
        <div style={{ background: 'rgba(168, 85, 247, 0.15)', border: '1px solid rgba(168, 85, 247, 0.4)', borderRadius: '8px', padding: '10px 16px', fontSize: '14px' }}>
          <strong style={{ color: '#a855f7' }}>Смена:</strong> {shift === 'currentDay' || shift === 'lastDay' ? 'День' : 'Ночь'}
        </div>
        <div style={{ background: 'rgba(249, 115, 22, 0.15)', border: '1px solid rgba(249, 115, 22, 0.4)', borderRadius: '8px', padding: '10px 16px', fontSize: '14px' }}>
          <strong style={{ color: '#f97316' }}>Оператор:</strong> Не указан
        </div>
        <div style={{ background: 'rgba(20, 184, 166, 0.15)', border: '1px solid rgba(20, 184, 166, 0.4)', borderRadius: '8px', padding: '10px 16px', fontSize: '14px' }}>
          <strong style={{ color: '#14b8a6' }}>Дата:</strong> {shiftDate}
        </div>
      </div>

      {loading ? (
        <div style={{ textAlign: 'center', padding: '40px', color: 'var(--text-secondary)' }}>
          Загрузка...
        </div>
      ) : (
        <div style={{ overflowX: 'auto', border: '1px solid var(--border)', borderRadius: '8px' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
            <thead style={{ background: 'var(--surface)' }}>
              <tr>
                <th style={{ padding: '12px', border: '1px solid var(--border)', textAlign: 'left' }}>Дата</th>
                <th style={{ padding: '12px', border: '1px solid var(--border)', textAlign: 'left' }}>Время</th>
                <th style={{ padding: '12px', border: '1px solid var(--border)', textAlign: 'left' }}>Действие</th>
                <th style={{ padding: '12px', border: '1px solid var(--border)', textAlign: 'left' }}>Описание</th>
                <th style={{ padding: '12px', border: '1px solid var(--border)', textAlign: 'center' }}>Общее время (мин)</th>
                <th style={{ padding: '12px', border: '1px solid var(--border)', textAlign: 'center' }}>Действия</th>
              </tr>
            </thead>
            <tbody>
              {getFilteredLogs().map((log) => {
                const date = new Date(log.timestamp)
                return (
                  <tr key={log.id} style={{ background: getRowColor(log.action) }}>
                    <td style={{ padding: '10px', border: '1px solid var(--border)' }}>
                      {date.toLocaleDateString('ru-RU')}
                    </td>
                    <td style={{ padding: '10px', border: '1px solid var(--border)' }}>
                      {date.toLocaleTimeString('ru-RU')}
                    </td>
                    <td style={{ padding: '10px', border: '1px solid var(--border)' }}>
                      {log.action}
                    </td>
                    <td style={{ padding: '10px', border: '1px solid var(--border)' }}>
                      {log.description}
                    </td>
                    <td style={{ padding: '10px', border: '1px solid var(--border)', textAlign: 'center' }}>
                      {log.total_minutes || ''}
                    </td>
                    <td style={{ padding: '10px', border: '1px solid var(--border)', textAlign: 'center' }}>
                      <button 
                        onClick={() => openEditModal(log)}
                        style={{ padding: '4px 8px', fontSize: '11px', marginRight: '4px' }}
                        className="btn btn-secondary"
                      >
                        Редактировать
                      </button>
                      <button 
                        onClick={() => deleteLog(log.id)}
                        style={{ padding: '4px 8px', fontSize: '11px' }}
                        className="btn btn-secondary"
                      >
                        Удалить
                      </button>
                    </td>
                  </tr>
                )
              })}
              {getFilteredLogs().length === 0 && (
                <tr>
                  <td colSpan="6" style={{ padding: '20px', textAlign: 'center', color: 'var(--text-secondary)' }}>
                    Нет записей
                  </td>
                </tr>
              )}
              {getFilteredLogs().length > 0 && (
                <tr style={{ background: 'rgba(16, 185, 129, 0.15)', fontWeight: 'bold' }}>
                  <td colSpan="4" style={{ padding: '12px', border: '1px solid var(--border)', textAlign: 'right' }}>
                    Общее время от начала до конца смены:
                  </td>
                  <td style={{ padding: '12px', border: '1px solid var(--border)', textAlign: 'center' }}>
                    {getFilteredLogs().reduce((sum, log) => sum + (log.total_minutes || 0), 0)} мин
                  </td>
                  <td style={{ padding: '12px', border: '1px solid var(--border)' }}></td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {showModal && (
        <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000 }}>
          <div style={{ background: 'var(--surface)', padding: '24px', borderRadius: '12px', maxWidth: '500px', width: '100%', maxHeight: '90vh', overflowY: 'auto' }}>
            <h3 style={{ fontSize: '18px', fontWeight: 700, marginBottom: '16px', color: 'var(--text-primary)' }}>
              {editingLog ? 'Редактировать запись' : 'Добавить запись'}
            </h3>
            
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div>
                <label className="form-label">Быстрый шаблон</label>
                <select
                  className="form-input"
                  defaultValue=""
                  onChange={(e) => applyPreset(e.target.value)}
                >
                  <option value="">— Выбрать шаблон —</option>
                  {QUICK_PRESETS.map(p => (
                    <option key={p.key} value={p.key}>{p.label}</option>
                  ))}
                </select>
              </div>
              <div>
                <label className="form-label">Оператор <span style={{ color: '#ef4444' }}>*</span></label>
                <input 
                  type="text"
                  value={formData.operatorName}
                  onChange={(e) => setFormData({ ...formData, operatorName: e.target.value })}
                  className="form-input"
                  required
                  placeholder="Введите имя оператора"
                />
              </div>
              <div>
                <label className="form-label">Действие</label>
                <input 
                  type="text"
                  value={formData.action}
                  onChange={(e) => setFormData({ ...formData, action: e.target.value })}
                  className="form-input"
                  placeholder="Например: Исправление ошибки"
                />
              </div>
              <div>
                <label className="form-label">Описание</label>
                <textarea 
                  value={formData.description}
                  onChange={(e) => setFormData({ ...formData, description: e.target.value })}
                  className="form-input"
                  rows={3}
                  placeholder="Подробное описание действия..."
                />
              </div>
              <div>
                <label className="form-label">Дата и время</label>
                <input 
                  type="datetime-local"
                  value={formData.timestamp}
                  onChange={(e) => setFormData({ ...formData, timestamp: e.target.value })}
                  className="form-input"
                />
              </div>
              <div>
                <label className="form-label">Смена</label>
                <select
                  value={formData.shift}
                  onChange={(e) => setFormData({ ...formData, shift: e.target.value })}
                  className="form-input"
                >
                  <option value="день">Дневная</option>
                  <option value="ночь">Ночная</option>
                </select>
              </div>
              <div>
                <label className="form-label">Общее время (мин)</label>
                <input
                  type="number"
                  value={formData.total_minutes || ''}
                  onChange={(e) => setFormData({ ...formData, total_minutes: e.target.value ? parseInt(e.target.value) : '' })}
                  className="form-input"
                  placeholder="Например: 15"
                  min="0"
                />
              </div>
            </div>

            <div style={{ display: 'flex', gap: '10px', marginTop: '20px', justifyContent: 'flex-end' }}>
              <button 
                onClick={() => setShowModal(false)}
                className="btn btn-secondary"
                style={{ padding: '8px 16px' }}
              >
                Отмена
              </button>
              <button 
                onClick={saveLog}
                className="btn btn-primary"
                style={{ padding: '8px 16px' }}
              >
                Сохранить
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

export default OperatorReport
