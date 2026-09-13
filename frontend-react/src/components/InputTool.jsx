import { useState } from 'react'

function InputTool() {
  const [inputValue, setInputValue] = useState('')

  return (
    <div className="card">
      <div className="card-header">
        <h2 className="card-title">Поле ввода</h2>
      </div>

      <div style={{ padding: '20px' }}>
        <div style={{ background: 'var(--surface)', border: '1px solid var(--border)', borderRadius: '10px', padding: '16px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '14px' }}>
            <div style={{ width: '4px', height: '18px', background: 'var(--primary)', borderRadius: '2px' }}></div>
            <h3 style={{ fontSize: '14px', fontWeight: 700, margin: 0, color: 'var(--text-primary)' }}>Ввод данных</h3>
          </div>
          <div className="form-group">
            <label className="form-label">Введите текст</label>
            <textarea
              value={inputValue}
              onChange={(e) => setInputValue(e.target.value)}
              className="form-input"
              placeholder="Введите текст..."
              rows={16}
              style={{ width: '100%', resize: 'vertical', fontFamily: 'inherit', fontSize: '14px', lineHeight: 1.5, minHeight: '350px' }}
            />
          </div>
          {inputValue && (
            <div style={{ marginTop: '12px', padding: '12px', background: 'var(--background)', borderRadius: '8px' }}>
              <p style={{ fontSize: '13px', color: 'var(--text-secondary)', margin: 0 }}>
                Символов: <strong style={{ color: 'var(--text-primary)' }}>{inputValue.length}</strong>
              </p>
            </div>
          )}
        </div>
      </div>
    </div>
  )
}

export default InputTool
