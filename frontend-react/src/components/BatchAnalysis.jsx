import { useState } from 'react'

function BatchAnalysis() {
  const [input, setInput] = useState('')
  const [results, setResults] = useState(null)

  const parseTimeHMSS = (timeStr) => {
    const parts = timeStr.trim().split(':')
    if (parts.length !== 3) return null
    
    const hours = parseInt(parts[0])
    const minutes = parseInt(parts[1])
    const seconds = parseInt(parts[2])
    
    if (isNaN(hours) || isNaN(minutes) || isNaN(seconds)) return null
    
    return hours * 3600 + minutes * 60 + seconds
  }

  const formatTime = (seconds) => {
    const hours = Math.floor(seconds / 3600)
    const mins = Math.floor((seconds % 3600) / 60)
    const secs = seconds % 60
    return `${hours.toString().padStart(2, '0')}:${mins.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}`
  }

  const calculate = () => {
    if (!input.trim()) {
      alert('Пожалуйста, вставьте список временных меток')
      return
    }

    const lines = input.trim().split('\n').filter(line => line.trim())
    
    if (lines.length < 3) {
      alert('Недостаточно данных. Нужно минимум 3 временные метки')
      return
    }

    const times = []
    for (const line of lines) {
      const parsed = parseTimeHMSS(line.trim())
      if (parsed !== null) {
        times.push(parsed)
      }
    }

    if (times.length < 3) {
      alert('Не удалось распознать временные метки. Используйте формат HH:MM:SS')
      return
    }

    const groups = []
    for (let i = 0; i < times.length; i += 3) {
      if (i + 2 < times.length) {
        groups.push([times[i], times[i + 1], times[i + 2]])
      }
    }

    if (groups.length === 0) {
      alert('Недостаточно данных для формирования групп')
      return
    }

    const diff12List = []
    const diff13List = []

    for (const group of groups) {
      const [t1, t2, t3] = group
      
      const diff12Seconds = Math.abs(t2 - t1)
      const diff12Minutes = Math.round(diff12Seconds / 60)
      diff12List.push(diff12Minutes)
      
      const diff13Seconds = Math.abs(t3 - t1)
      const diff13Minutes = Math.round(diff13Seconds / 60)
      diff13List.push(diff13Minutes)
    }

    const sum12 = diff12List.reduce((a, b) => a + b, 0)
    const sum13 = diff13List.reduce((a, b) => a + b, 0)

    const f06Count = lines.filter(line => line.trim().includes('F06')).length

    const groupResults = groups.map((group, index) => ({
      index: index + 1,
      t1: formatTime(group[0]),
      t2: formatTime(group[1]),
      t3: formatTime(group[2]),
      diff12: diff12List[index],
      diff13: diff13List[index]
    }))

    // Подробный расчёт для первых 3 групп
    const maxDetail = Math.min(3, groups.length)
    const detailedResults = []
    for (let i = 0; i < maxDetail; i++) {
      const group = groups[i]
      const t1 = formatTime(group[0])
      const t2 = formatTime(group[1])
      const t3 = formatTime(group[2])
      
      const diff12Seconds = Math.abs(group[1] - group[0])
      const diff12Minutes = Math.round(diff12Seconds / 60)
      const diff13Seconds = Math.abs(group[2] - group[0])
      const diff13Minutes = Math.round(diff13Seconds / 60)
      
      detailedResults.push({
        index: i + 1,
        t1,
        t2,
        t3,
        t1Sec: group[0],
        t2Sec: group[1],
        t3Sec: group[2],
        diff12Sec,
        diff12Min: diff12Minutes,
        diff13Sec,
        diff13Min: diff13Minutes
      })
    }

    setResults({
      f06Count,
      sum12,
      sum13,
      groupResults,
      detailedResults
    })
  }

  return (
    <div className="card">
      <div className="card-header">
        <div>
          <h2 className="card-title">Пакетный анализ штабелеров</h2>
          <p className="card-subtitle">Вставьте таблицу статусов штабелеров из Excel</p>
        </div>
        <button onClick={calculate} className="btn btn-primary">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/>
            <polyline points="17 6 23 6 23 12"/>
          </svg>
          Рассчитать
        </button>
      </div>

      <div className="form-group">
        <label className="form-label">Список временных меток (HH:MM:SS)</label>
        <textarea
          value={input}
          onChange={(e) => setInput(e.target.value)}
          className="form-input"
          placeholder="Вставьте список временных меток&#10;&#10;16:56:30&#10;19:30:15&#10;20:45:00&#10;...&#10;&#10;Время будет разбито на группы по 3 элемента и рассчитаны разницы"
        />
      </div>

      {results && (
        <div className="mt-4">
          <div className="stats-grid">
            <div className="stat-card stat-danger">
              <div className="stat-label">Количество F06</div>
              <div className="stat-value">{results.f06Count}</div>
            </div>
            <div className="stat-card stat-warning">
              <div className="stat-label">Время ожидания технической службы</div>
              <div className="stat-value">{results.sum12}</div>
              <div className="stat-unit">МИНУТ</div>
            </div>
            <div className="stat-card stat-success">
              <div className="stat-label">Общее время ошибок</div>
              <div className="stat-value">{results.sum13}</div>
              <div className="stat-unit">МИНУТ</div>
            </div>
          </div>

          <div className="card" style={{ marginTop: '20px' }}>
            <div className="card-header">
              <h2 className="card-title">Результаты по группам</h2>
            </div>

            <div style={{ display: 'grid', gap: '8px', marginBottom: '20px' }}>
              {results.groupResults.map((group) => (
                <div key={group.index} style={{ padding: '12px', background: 'var(--background)', borderRadius: '8px', border: '1px solid var(--border)' }}>
                  <div style={{ fontWeight: 500, marginBottom: '4px' }}>
                    Группа {group.index}: {group.t1} → {group.t2} → {group.t3}
                  </div>
                  <div style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>
                    (1→2): {group.diff12} мин | (1→3): {group.diff13} мин
                  </div>
                </div>
              ))}
            </div>

            <div className="card-header">
              <h2 className="card-title">Подробный расчёт (первые 3 группы)</h2>
            </div>

            <div>
              {results.detailedResults && results.detailedResults.map((detail, index) => (
                <div key={index} style={{ padding: '16px', background: 'var(--background)', borderRadius: '8px', border: '1px solid var(--border)', marginBottom: '12px' }}>
                  <div style={{ fontWeight: 600, marginBottom: '12px' }}>Группа {detail.index}</div>
                  <div style={{ marginBottom: '8px' }}>
                    <strong>Времена:</strong> {detail.t1} (1), {detail.t2} (2), {detail.t3} (3)
                  </div>
                  <div style={{ marginBottom: '8px' }}>
                    <strong>(1 → 2):</strong> |{detail.t2Sec} - {detail.t1Sec}| = {detail.diff12Sec} сек = {detail.diff12Min} мин
                  </div>
                  <div>
                    <strong>(1 → 3):</strong> |{detail.t3Sec} - {detail.t1Sec}| = {detail.diff13Sec} сек = {detail.diff13Min} мин
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

export default BatchAnalysis
