import { useState, useEffect } from 'react'

function Settings({ user }) {
  const [theme, setTheme] = useState('gray')
  const [fontSize, setFontSize] = useState('medium')
  const [autoCalc, setAutoCalc] = useState(false)
  const [showHints, setShowHints] = useState(true)
  const [saveHistory, setSaveHistory] = useState(true)
  const [selections, setSelections] = useState([])
  const [newSelection, setNewSelection] = useState({ name: '', value: '' })

  const userPrefix = user?.username || 'default'

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
    const savedTheme = localStorage.getItem(`theme_${userPrefix}`) || 'gray'
    const savedFontSize = localStorage.getItem(`fontSize_${userPrefix}`) || 'medium'
    const savedAutoCalc = localStorage.getItem(`autoCalc_${userPrefix}`) === 'true'
    const savedShowHints = localStorage.getItem(`showHints_${userPrefix}`) !== 'false'
    const savedSaveHistory = localStorage.getItem(`saveHistory_${userPrefix}`) !== 'false'
    const savedSelections = localStorage.getItem(`selections_${userPrefix}`)
    
    setTheme(savedTheme)
    setFontSize(savedFontSize)
    setAutoCalc(savedAutoCalc)
    setShowHints(savedShowHints)
    setSaveHistory(savedSaveHistory)
    if (savedSelections) {
      setSelections(JSON.parse(savedSelections))
    } else {
      // Добавить значения по умолчанию для операторов и техслужбы
      const defaultSelections = [
        { id: 1, name: 'Оператор', value: 'Aser' },
        { id: 2, name: 'Оператор', value: 'Лёша' },
        { id: 3, name: 'Оператор', value: 'Владимир' },
        { id: 4, name: 'Тех служба', value: 'Лёша, Владимир' },
      ]
      setSelections(defaultSelections)
      localStorage.setItem(`selections_${userPrefix}`, JSON.stringify(defaultSelections))
    }
    
    document.documentElement.setAttribute('data-theme', savedTheme)
    document.documentElement.style.fontSize = savedFontSize === 'small' ? '14px' : savedFontSize === 'large' ? '18px' : '16px'
  }, [userPrefix])

  const handleThemeChange = (themeId) => {
    setTheme(themeId)
    document.documentElement.setAttribute('data-theme', themeId)
    localStorage.setItem(`theme_${userPrefix}`, themeId)
  }

  const handleFontSizeChange = (size) => {
    setFontSize(size)
    document.documentElement.style.fontSize = size === 'small' ? '14px' : size === 'large' ? '18px' : '16px'
    localStorage.setItem(`fontSize_${userPrefix}`, size)
  }

  const handleAutoCalcChange = (checked) => {
    setAutoCalc(checked)
    localStorage.setItem(`autoCalc_${userPrefix}`, checked)
  }

  const handleShowHintsChange = (checked) => {
    setShowHints(checked)
    localStorage.setItem(`showHints_${userPrefix}`, checked)
  }

  const handleSaveHistoryChange = (checked) => {
    setSaveHistory(checked)
    localStorage.setItem(`saveHistory_${userPrefix}`, checked)
  }

  const addSelection = () => {
    if (!newSelection.name || !newSelection.value) {
      alert('Заполните название и значение')
      return
    }
    const newSelections = [...selections, { ...newSelection, id: Date.now() }]
    setSelections(newSelections)
    setNewSelection({ name: '', value: '' })
    localStorage.setItem(`selections_${userPrefix}`, JSON.stringify(newSelections))
  }

  const removeSelection = (id) => {
    const newSelections = selections.filter(s => s.id !== id)
    setSelections(newSelections)
    localStorage.setItem(`selections_${userPrefix}`, JSON.stringify(newSelections))
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

        <div className="form-group">
          <label className="form-label">Автоматический расчёт</label>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <input
              type="checkbox"
              checked={autoCalc}
              onChange={(e) => handleAutoCalcChange(e.target.checked)}
              style={{ width: '18px', height: '18px', cursor: 'pointer' }}
            />
            <span style={{ fontSize: '14px', color: 'var(--text-primary)' }}>
              Автоматически рассчитывать при вставке данных
            </span>
          </div>
        </div>

        <div className="form-group">
          <label className="form-label">Показывать подсказки</label>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <input
              type="checkbox"
              checked={showHints}
              onChange={(e) => handleShowHintsChange(e.target.checked)}
              style={{ width: '18px', height: '18px', cursor: 'pointer' }}
            />
            <span style={{ fontSize: '14px', color: 'var(--text-primary)' }}>
              Показывать подсказки и примеры
            </span>
          </div>
        </div>

        <div className="form-group">
          <label className="form-label">Сохранять историю расчётов</label>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <input
              type="checkbox"
              checked={saveHistory}
              onChange={(e) => handleSaveHistoryChange(e.target.checked)}
              style={{ width: '18px', height: '18px', cursor: 'pointer' }}
            />
            <span style={{ fontSize: '14px', color: 'var(--text-primary)' }}>
              Сохранять историю расчётов
            </span>
          </div>
        </div>

        <div style={{ marginTop: '32px', paddingTop: '20px', borderTop: '1px solid var(--border)' }}>
          <h3 style={{ fontSize: '16px', fontWeight: 600, marginBottom: '16px', color: 'var(--text-primary)' }}>
            Управление выборами
          </h3>
          <p style={{ fontSize: '13px', color: 'var(--text-secondary)', marginBottom: '16px' }}>
            Добавляйте и управляйте часто используемыми значениями (операторы, техслужба и т.д.)
          </p>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px', marginBottom: '12px' }}>
            <div className="form-group" style={{ margin: 0 }}>
              <label className="form-label">Название</label>
              <input
                type="text"
                value={newSelection.name}
                onChange={(e) => setNewSelection({ ...newSelection, name: e.target.value })}
                className="form-input"
                placeholder="Оператор"
              />
            </div>
            <div className="form-group" style={{ margin: 0 }}>
              <label className="form-label">Значение</label>
              <input
                type="text"
                value={newSelection.value}
                onChange={(e) => setNewSelection({ ...newSelection, value: e.target.value })}
                className="form-input"
                placeholder="Aser"
              />
            </div>
          </div>

          <button onClick={addSelection} className="btn btn-primary" style={{ marginBottom: '16px' }}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <line x1="12" y1="5" x2="12" y2="19"/>
              <line x1="5" y1="12" x2="19" y2="12"/>
            </svg>
            Добавить выбор
          </button>

          <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
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
                  <div style={{ display: 'flex', gap: '8px' }}>
                    <button
                      onClick={() => {
                        setNewSelection({ name: selection.name, value: selection.value })
                      }}
                      className="btn btn-secondary"
                      style={{ padding: '6px 12px', fontSize: '12px' }}
                    >
                      Изменить
                    </button>
                    <button
                      onClick={() => removeSelection(selection.id)}
                      className="btn btn-danger"
                      style={{ padding: '6px 12px', fontSize: '12px' }}
                    >
                      Удалить
                    </button>
                  </div>
                </div>
              ))
            )}
          </div>
        </div>
      </div>
    </div>
  )
}

export default Settings
