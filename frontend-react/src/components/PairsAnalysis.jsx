import { useState } from 'react'

function PairsAnalysis() {
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
    
    if (lines.length < 2) {
      alert('Недостаточно данных. Нужно минимум 2 временные метки')
      return
    }

    const times = []
    for (const line of lines) {
      const parsed = parseTimeHMSS(line.trim())
      if (parsed !== null) {
        times.push(parsed)
      }
    }

    if (times.length < 2) {
      alert('Не удалось распознать временные метки. Используйте формат HH:MM:SS')
      return
    }

    const pairs = []
    for (let i = 0; i < times.length; i += 2) {
      if (i + 1 < times.length) {
        pairs.push([times[i], times[i + 1]])
      }
    }

    if (pairs.length === 0) {
      alert('Недостаточно данных для формирования пар')
      return
    }

    const diffList = []

    for (const pair of pairs) {
      const [t1, t2] = pair
      
      const diffSeconds = Math.abs(t2 - t1)
      const diffMinutes = Math.round(diffSeconds / 60)
      diffList.push(diffMinutes)
    }

    const sum = diffList.reduce((a, b) => a + b, 0)

    const f06Count = lines.filter(line => line.trim().includes('F06')).length

    const pairResults = pairs.map((pair, index) => ({
      index: index + 1,
      t1: formatTime(pair[0]),
      t2: formatTime(pair[1]),
      diff: diffList[index]
    }))

    // Подробный расчёт для первых 3 пар
    const maxDetail = Math.min(3, pairs.length)
    const detailedResults = []
    for (let i = 0; i < maxDetail; i++) {
      const pair = pairs[i]
      const t1 = formatTime(pair[0])
      const t2 = formatTime(pair[1])
      
      const earlier = Math.min(pair[0], pair[1])
      const later = Math.max(pair[0], pair[1])
      const diffSeconds = later - earlier
      const diffMinutes = diffSeconds / 60
      const rounded = Math.round(diffMinutes)
      
      detailedResults.push({
        index: i + 1,
        t1,
        t2,
        earlierTime: earlier,
        laterTime: later,
        earlierMin: (earlier / 60).toFixed(2),
        laterMin: (later / 60).toFixed(2),
        diffSeconds,
        diffMinutes: diffMinutes.toFixed(2),
        rounded
      })
    }

    setResults({
      f06Count,
      sum,
      pairResults,
      detailedResults
    })
  }

  return (
    <div className="card">
      <div className="card-header">
        <div>
          <h2 className="card-title">Пакетный анализ 2.0</h2>
          <p className="card-subtitle">Разбивка временных меток на пары и расчёт разниц</p>
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
          placeholder="Вставьте список временных меток&#10;&#10;16:56:30&#10;19:30:15&#10;20:45:00&#10;21:10:30&#10;...&#10;&#10;Время будет разбито на пары (1-2, 3-4 и т.д.) и рассчитаны разницы"
        />
      </div>

      {results && (
        <div className="mt-4">
          <div className="stats-grid">
            <div className="stat-card stat-danger">
              <div className="stat-label">Количество F06</div>
              <div className="stat-value">{results.f06Count}</div>
            </div>
            <div className="stat-card stat-primary">
              <div className="stat-label">Итоговая сумма</div>
              <div className="stat-value">{results.sum}</div>
              <div className="stat-unit">МИНУТ</div>
            </div>
          </div>

          <div className="card" style={{ marginTop: '20px' }}>
            <div className="card-header">
              <h2 className="card-title">Результаты по парам</h2>
            </div>

            <div style={{ display: 'grid', gap: '8px', marginBottom: '20px' }}>
              {results.pairResults.map((pair) => (
                <div key={pair.index} style={{ padding: '12px', background: 'var(--background)', borderRadius: '8px', border: '1px solid var(--border)' }}>
                  <div style={{ fontWeight: 500, marginBottom: '4px' }}>
                    Пара {pair.index}: {pair.t1} → {pair.t2}
                  </div>
                  <div style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>
                    Разница: {pair.diff} мин
                  </div>
                </div>
              ))}
            </div>

            <div className="card-header">
              <h2 className="card-title">Подробный расчёт (первые 3 пары)</h2>
            </div>

            <div>
              {results.detailedResults && results.detailedResults.map((detail, index) => (
                <div key={index} style={{ padding: '16px', background: 'var(--background)', borderRadius: '8px', border: '1px solid var(--border)', marginBottom: '12px' }}>
                  <div style={{ fontWeight: 600, marginBottom: '12px' }}>Пара {detail.index}</div>
                  <div style={{ marginBottom: '8px' }}>
                    <strong>Времена:</strong> {detail.t1} (1), {detail.t2} (2)
                  </div>
                  <div style={{ marginBottom: '8px' }}>
                    <strong>Раннее:</strong> {detail.earlierTime} сек ({detail.earlierMin} мин)
                  </div>
                  <div style={{ marginBottom: '8px' }}>
                    <strong>Позднее:</strong> {detail.laterTime} сек ({detail.laterMin} мин)
                  </div>
                  <div>
                    <strong>Разница:</strong> {detail.laterTime} - {detail.earlierTime} = {detail.diffSeconds} сек = {detail.diffMinutes} мин (округлено: {detail.rounded} мин)
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

export default PairsAnalysis
