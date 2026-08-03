import { useState } from 'react'

function QuickCalc() {
  const [input, setInput] = useState('')
  const [results, setResults] = useState([])
  const [unit, setUnit] = useState('minutes')

  const parseTimeHM = (timeStr) => {
    const parts = timeStr.trim().split(':')
    if (parts.length !== 2) return null
    
    const hours = parseInt(parts[0])
    const minutes = parseInt(parts[1])
    
    if (isNaN(hours) || isNaN(minutes)) return null
    
    return hours * 60 + minutes
  }

  const calculate = () => {
    const lines = input.trim().split('\n').filter(l => l.trim())

    if (lines.length === 0) {
      setResults([])
      return
    }

    let totalDiff = 0
    const newResults = []
    let hasAny = false

    for (const line of lines) {
      const parts = line.trim().split(/[\t\s]+/).filter(p => p.trim())
      if (parts.length < 2) continue

      const start = parseTimeHM(parts[0])
      const end = parseTimeHM(parts[1])
      if (start === null || end === null) continue

      let diff = end - start
      if (diff < 0) diff += 24 * 60

      let displayDiff = diff
      if (unit === 'milliseconds') {
        displayDiff = diff * 60 * 1000
        totalDiff += displayDiff
      } else {
        totalDiff += diff
      }

      hasAny = true
      newResults.push({
        start: parts[0],
        end: parts[1],
        diff: displayDiff
      })
    }

    if (!hasAny) {
      setResults([])
      return
    }

    setResults(newResults)
  }

  const handleInputChange = (e) => {
    setInput(e.target.value)
    calculate()
  }

  const handleUnitChange = (e) => {
    setUnit(e.target.value)
    calculate()
  }

  const total = results.reduce((sum, r) => sum + r.diff, 0)

  return (
    <div className="card">
      <div className="card-header">
        <div>
          <h2 className="card-title">Быстрый расчёт разницы</h2>
          <p className="card-subtitle">Вычислите разницу между двумя временными отметками</p>
        </div>
      </div>

      <div className="form-group">
        <label className="form-label">Единица измерения</label>
        <select
          value={unit}
          onChange={handleUnitChange}
          className="form-input"
          style={{ maxWidth: '200px' }}
        >
          <option value="minutes">Минуты</option>
          <option value="milliseconds">Миллисекунды</option>
        </select>
      </div>

      <div className="form-group">
        <label className="form-label">Времена (каждая пара на новой строке через табуляцию или пробел)</label>
        <textarea
          value={input}
          onChange={handleInputChange}
          className="form-input"
          placeholder="20:21&#9;23:41&#10;21:30&#9;00:15&#10;22:42&#9;00:45"
          rows="5"
          style={{ minHeight: '100px' }}
        />
      </div>

      {results.length > 0 && (
        <>
          <div className="table-container" style={{ marginTop: '12px' }}>
            <table>
              <thead>
                <tr>
                  <th style={{ padding: '8px', border: '1px solid var(--border)', textAlign: 'center' }}>Начало</th>
                  <th style={{ padding: '8px', border: '1px solid var(--border)', textAlign: 'center' }}>Конец</th>
                  <th style={{ padding: '8px', border: '1px solid var(--border)', textAlign: 'center' }}>
                    Разница ({unit === 'minutes' ? 'мин' : 'мс'})
                  </th>
                </tr>
              </thead>
              <tbody>
                {results.map((result, index) => (
                  <tr key={index}>
                    <td style={{ padding: '8px', border: '1px solid var(--border)', textAlign: 'center' }}>{result.start}</td>
                    <td style={{ padding: '8px', border: '1px solid var(--border)', textAlign: 'center' }}>{result.end}</td>
                    <td style={{ padding: '8px', border: '1px solid var(--border)', textAlign: 'center', fontWeight: 600 }}>{result.diff.toLocaleString()}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            <div style={{ marginTop: '10px', fontSize: '14px', fontWeight: 600 }}>
              Итого: {total.toLocaleString()} {unit === 'minutes' ? 'мин' : 'мс'}
            </div>
          </div>
        </>
      )}
    </div>
  )
}

export default QuickCalc
