import { useState, useEffect } from 'react'

function Settings() {
  const [theme, setTheme] = useState('gray')
  const [fontSize, setFontSize] = useState('medium')

  const themes = [
    { id: 'light', name: 'Светлая', gradient: 'linear-gradient(135deg, #ffffff 0%, #f8fafc 100%)' },
    { id: 'dark', name: 'Тёмная', gradient: 'linear-gradient(135deg, #1f2937 0%, #111827 100%)' },
    { id: 'gray', name: 'Серая', gradient: 'linear-gradient(135deg, #ffffff 0%, #f3f4f6 100%)' },
    { id: 'icon', name: 'Иконка', gradient: 'linear-gradient(135deg, #0ea5e9 0%, #0284c7 100%)' },
    { id: 'lime', name: 'Лайм', gradient: 'linear-gradient(135deg, #84cc16 0%, #65a30d 100%)' },
    { id: 'purple', name: 'Пурпур', gradient: 'linear-gradient(135deg, #8b5cf6 0%, #7c3aed 100%)' },
    { id: 'blue', name: 'Синяя', gradient: 'linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%)' },
    { id: 'red', name: 'Красная', gradient: 'linear-gradient(135deg, #dc2626 0%, #b91c1c 100%)' },
    { id: 'orange', name: 'Оранж', gradient: 'linear-gradient(135deg, #ea580c 0%, #c2410c 100%)' },
    { id: 'amber', name: 'Янтарь', gradient: 'linear-gradient(135deg, #d97706 0%, #b45309 100%)' },
    { id: 'rose', name: 'Розовая', gradient: 'linear-gradient(135deg, #e11d48 0%, #be123c 100%)' },
    { id: 'pink', name: 'Розовый', gradient: 'linear-gradient(135deg, #db2777 0%, #be185d 100%)' },
    { id: 'teal', name: 'Бирюза', gradient: 'linear-gradient(135deg, #0d9488 0%, #0f766e 100%)' },
    { id: 'cyan', name: 'Голубая', gradient: 'linear-gradient(135deg, #0891b2 0%, #0e7490 100%)' },
    { id: 'indigo', name: 'Индиго', gradient: 'linear-gradient(135deg, #4f46e5 0%, #4338ca 100%)' },
    { id: 'emerald', name: 'Изумруд', gradient: 'linear-gradient(135deg, #059669 0%, #047857 100%)' },
    { id: 'yellow', name: 'Жёлтая', gradient: 'linear-gradient(135deg, #eab308 0%, #a16207 100%)' },
    { id: 'gothic', name: 'Готика', gradient: 'linear-gradient(135deg, #18181b 0%, #09090b 100%)', border: '1px solid #a855f7' },
    { id: 'neon', name: 'Неон', gradient: 'linear-gradient(135deg, #0a0a0f 0%, #12121a 100%)', border: '1px solid #00ff88', boxShadow: '0 0 8px #00ff8866' },
  ]

  useEffect(() => {
    const savedTheme = localStorage.getItem('theme') || 'gray'
    const savedFontSize = localStorage.getItem('fontSize') || 'medium'
    
    setTheme(savedTheme)
    setFontSize(savedFontSize)
    
    document.documentElement.setAttribute('data-theme', savedTheme)
    document.documentElement.setAttribute('data-font-size', savedFontSize)
  }, [])

  const handleThemeChange = (newTheme) => {
    setTheme(newTheme)
    localStorage.setItem('theme', newTheme)
    document.documentElement.setAttribute('data-theme', newTheme)
  }

  const handleFontSizeChange = (newFontSize) => {
    setFontSize(newFontSize)
    localStorage.setItem('fontSize', newFontSize)
    document.documentElement.setAttribute('data-font-size', newFontSize)
  }

  return (
    <div className="card">
      <div className="card-header">
        <h2 className="card-title">Настройки</h2>
        <p className="card-subtitle">Персонализация интерфейса</p>
      </div>

      <div style={{ padding: '20px' }}>
        <div className="form-group">
          <label className="form-label">Тема оформления</label>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(100px, 1fr))', gap: '12px' }}>
            {themes.map((t) => (
              <button
                key={t.id}
                onClick={() => handleThemeChange(t.id)}
                style={{
                  padding: '12px',
                  border: `2px solid ${theme === t.id ? 'var(--primary)' : 'var(--border)'}`,
                  borderRadius: '8px',
                  background: 'var(--surface)',
                  cursor: 'pointer',
                  display: 'flex',
                  flexDirection: 'column',
                  alignItems: 'center',
                  gap: '8px',
                  transition: 'all 0.2s',
                }}
              >
                <div
                  style={{
                    width: '32px',
                    height: '32px',
                    borderRadius: '50%',
                    background: t.gradient,
                    border: t.border || 'none',
                    boxShadow: t.boxShadow || 'none',
                  }}
                />
                <span style={{ fontSize: '12px', fontWeight: '500', color: 'var(--text-primary)' }}>
                  {t.name}
                </span>
              </button>
            ))}
          </div>
        </div>

        <div className="form-group">
          <label className="form-label">Размер шрифта</label>
          <select
            value={fontSize}
            onChange={(e) => handleFontSizeChange(e.target.value)}
            className="form-input"
          >
            <option value="small">Мелкий</option>
            <option value="medium">Средний</option>
            <option value="large">Крупный</option>
          </select>
        </div>
      </div>
    </div>
  )
}

export default Settings
