import { useState } from 'react'

function Selections() {
  const [selections, setSelections] = useState([])
  const [newSelection, setNewSelection] = useState({ name: '', value: '' })

  const addSelection = () => {
    if (!newSelection.name || !newSelection.value) {
      alert('Заполните название и значение')
      return
    }
    setSelections([...selections, { ...newSelection, id: Date.now() }])
    setNewSelection({ name: '', value: '' })
  }

  const removeSelection = (id) => {
    setSelections(selections.filter(s => s.id !== id))
  }

  return (
    <div className="card">
      <div className="card-header">
        <div>
          <h2 className="card-title">Управление выборами</h2>
          <p className="card-subtitle">Добавляйте и управляйте часто используемыми значениями для общих данных</p>
        </div>
        <button onClick={addSelection} className="btn btn-primary">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <line x1="12" y1="5" x2="12" y2="19"/>
            <line x1="5" y1="12" x2="19" y2="12"/>
          </svg>
          Добавить выбор
        </button>
      </div>

      <div className="form-group">
        <label className="form-label">Название</label>
        <input
          type="text"
          value={newSelection.name}
          onChange={(e) => setNewSelection({ ...newSelection, name: e.target.value })}
          className="form-input"
          placeholder="Название выбора"
        />
      </div>

      <div className="form-group">
        <label className="form-label">Значение</label>
        <input
          type="text"
          value={newSelection.value}
          onChange={(e) => setNewSelection({ ...newSelection, value: e.target.value })}
          className="form-input"
          placeholder="Значение"
        />
      </div>

      <div style={{ display: 'flex', flexDirection: 'column', gap: '12px', marginTop: '20px' }}>
        {selections.length === 0 ? (
          <p style={{ color: 'var(--text-secondary)', fontSize: '14px' }}>Нет добавленных выборов</p>
        ) : (
          selections.map((selection) => (
            <div
              key={selection.id}
              style={{
                background: 'var(--background)',
                border: '1px solid var(--border)',
                borderRadius: '8px',
                padding: '16px',
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center'
              }}
            >
              <div>
                <div style={{ fontWeight: 600, color: 'var(--text-primary)', marginBottom: '4px' }}>
                  {selection.name}
                </div>
                <div style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>
                  {selection.value}
                </div>
              </div>
              <button
                onClick={() => removeSelection(selection.id)}
                className="btn btn-danger"
                style={{ padding: '6px 12px', fontSize: '12px' }}
              >
                Удалить
              </button>
            </div>
          ))
        )}
      </div>
    </div>
  )
}

export default Selections
