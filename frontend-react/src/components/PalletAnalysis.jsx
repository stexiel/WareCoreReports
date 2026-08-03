import { useState } from 'react'

function PalletAnalysis() {
  const [input, setInput] = useState('')
  const [results, setResults] = useState(null)

  const analyze = () => {
    if (!input.trim()) {
      alert('Пожалуйста, вставьте список паллет')
      return
    }

    const lines = input.trim().split('\n').filter(line => line.trim())
    
    if (lines.length === 0) {
      alert('Не удалось распознать данные')
      return
    }

    const pallets = {}
    
    for (const line of lines) {
      const trimmed = line.trim()
      if (trimmed.length < 3) continue
      
      const type = trimmed.slice(-2)
      const id = trimmed.slice(0, -2)
      
      if (!pallets[id]) {
        pallets[id] = { d1: false, d2: false }
      }
      
      if (type === 'D1') pallets[id].d1 = true
      if (type === 'D2') pallets[id].d2 = true
    }

    const paired = []
    const unpaired = []

    for (const id in pallets) {
      const pallet = pallets[id]
      if (pallet.d1 && pallet.d2) {
        paired.push({ id, d1: true, d2: true })
      } else if (pallet.d1 && !pallet.d2) {
        unpaired.push({ id, missing: 'D2' })
      } else if (!pallet.d1 && pallet.d2) {
        unpaired.push({ id, missing: 'D1' })
      }
    }

    const pairedRows = paired.length * 2
    const unpairedRows = unpaired.length
    const totalRows = pairedRows + unpairedRows

    let output = ''
    
    if (paired.length > 0) {
      output += 'Парные паллеты:\n'
      for (const p of paired) {
        output += `${p.id} — D1, D2\n`
      }
      output += '\n'
    }

    if (unpaired.length > 0) {
      output += 'Непарные паллеты:\n'
      for (const p of unpaired) {
        output += `${p.id} — отсутствует ${p.missing}\n`
      }
      output += '\n'
    }

    output += `Парные\t${pairedRows}\n`
    output += `Не парные\t${unpairedRows}\n`
    output += `Общий кол палет\t${totalRows}`

    setResults({
      pairedCount: paired.length,
      unpairedCount: unpaired.length,
      totalCount: totalRows,
      output,
      paired,
      unpaired
    })
  }

  return (
    <div className="card">
      <div className="card-header">
        <div>
          <h2 className="card-title">Анализ паллет</h2>
          <p className="card-subtitle">Группировка паллет по идентификатору и определение пар</p>
        </div>
        <button onClick={analyze} className="btn btn-primary">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/>
            <polyline points="17 6 23 6 23 12"/>
          </svg>
          Анализировать
        </button>
      </div>

      <div className="form-group">
        <label className="form-label">Список паллет (ID\tD1 или ID\tD2)</label>
        <textarea
          value={input}
          onChange={(e) => setInput(e.target.value)}
          className="form-input"
          placeholder="L01X01Y73Z01D1&#10;L01X01Y73Z01D2&#10;L02X02Y74Z01D1&#10;..."
        />
      </div>

      {results && (
        <div className="mt-4">
          <div className="stats-grid">
            <div className="stat-card stat-success">
              <div className="stat-label">Парные паллеты</div>
              <div className="stat-value">{results.pairedCount}</div>
            </div>
            <div className="stat-card stat-danger">
              <div className="stat-label">Непарные паллеты</div>
              <div className="stat-value">{results.unpairedCount}</div>
            </div>
            <div className="stat-card stat-primary">
              <div className="stat-label">Всего паллет</div>
              <div className="stat-value">{results.totalCount}</div>
            </div>
          </div>

          <div className="card" style={{ marginTop: '20px' }}>
            <div className="card-header">
              <h2 className="card-title">Результаты анализа</h2>
            </div>
            <pre style={{ whiteSpace: 'pre-wrap', fontFamily: 'monospace', color: 'var(--text-primary)' }}>
              {results.output}
            </pre>
          </div>
        </div>
      )}
    </div>
  )
}

export default PalletAnalysis
