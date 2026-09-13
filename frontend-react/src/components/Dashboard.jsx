import { useState } from 'react'
import QuickCalc from './QuickCalc'
import BatchAnalysis from './BatchAnalysis'
import PalletAnalysis from './PalletAnalysis'
import ExcelGenerator from './ExcelGenerator'
import Settings from './Settings'
import Translator from './Translator'
import PairsAnalysis from './PairsAnalysis'
import AutoDB from './AutoDB'
import PalletOpt from './PalletOpt'
import OperatorReport from './OperatorReport'
import InputTool from './InputTool'

function Dashboard() {
  const [activeTab, setActiveTab] = useState(() => localStorage.getItem('activeTab') || 'palletOpt')
  const [sidebarOpen, setSidebarOpen] = useState(false)

  // Локальные инструменты не требуют backend/БД, работают всегда.
  // Отчёты требуют backend + подключение к SQL Server (WMS/WCS).
  const navItems = [
    { id: 'translator', label: 'Переводчик', icon: 'translate', section: 'local' },
    { id: 'inputTool', label: 'Поле ввода', icon: 'edit', section: 'local' },
    { id: 'palletOpt', label: 'Создание приемки', icon: 'check', section: 'local' },
    { id: 'excel', label: 'Генератор Excel', icon: 'file', section: 'local' },
    { id: 'pallet', label: 'Анализ паллет', icon: 'grid', section: 'local' },
    { id: 'quick', label: 'Быстрый расчёт', icon: 'clock', section: 'local' },
    { id: 'batch', label: 'Пакетный анализ', icon: 'calendar', section: 'local' },
    { id: 'pairs', label: 'Пакетный анализ 2.0', icon: 'users', section: 'local' },
    { id: 'settings', label: 'Настройки', icon: 'settings', section: 'local' },
    { id: 'autodb', label: 'Автозагрузка из БД', icon: 'database', section: 'backend' },
    { id: 'operatorReport', label: 'Отчет операторов', icon: 'report', section: 'backend' },
  ]

  const sectionTitles = {
    local: 'Локальные инструменты',
    backend: 'Отчёты (нужен backend)',
  }

  const icons = {
    clock: (
      <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
        <circle cx="12" cy="12" r="10"/>
        <polyline points="12 6 12 12 16 14"/>
      </svg>
    ),
    calendar: (
      <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
        <rect x="3" y="4" width="18" height="18" rx="2"/>
        <line x1="16" y1="2" x2="16" y2="6"/>
        <line x1="8" y1="2" x2="8" y2="6"/>
        <line x1="3" y1="10" x2="21" y2="10"/>
      </svg>
    ),
    grid: (
      <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
        <rect x="3" y="3" width="18" height="18" rx="2"/>
        <path d="M9 3v18"/>
        <path d="M15 3v18"/>
      </svg>
    ),
    file: (
      <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
        <polyline points="14 2 14 8 20 8"/>
        <path d="M8 13h8"/>
        <path d="M8 17h8"/>
        <path d="M8 9h2"/>
      </svg>
    ),
    users: (
      <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
        <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/>
        <circle cx="9" cy="7" r="4"/>
        <path d="M23 21v-2a4 4 0 0 0-3-3.87"/>
        <path d="M16 3.13a4 4 0 0 1 0 7.75"/>
      </svg>
    ),
    settings: (
      <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
        <circle cx="12" cy="12" r="3"/>
        <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"/>
      </svg>
    ),
    edit: (
      <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
        <path d="M12 20h9"/>
        <path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"/>
      </svg>
    ),
    translate: (
      <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
        <path d="M5 8l6 6"/>
        <path d="M4 14l6-6 2-3"/>
        <path d="M2 5h12"/>
        <path d="M7 2h1"/>
        <path d="M22 22l-5-10-5 10"/>
        <path d="M14 18h6"/>
      </svg>
    ),
    check: (
      <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
        <path d="M9 11l3 3L22 4"/>
        <path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/>
      </svg>
    ),
    database: (
      <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
        <ellipse cx="12" cy="5" rx="9" ry="3"/>
        <path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"/>
        <path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"/>
      </svg>
    ),
    report: (
      <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
        <polyline points="14 2 14 8 20 8"/>
        <line x1="16" y1="13" x2="8" y2="13"/>
        <line x1="16" y1="17" x2="8" y2="17"/>
        <polyline points="10 9 9 9 8 9"/>
      </svg>
    ),
  }

  const groupedNavItems = navItems.reduce((acc, item) => {
    if (!acc[item.section]) {
      acc[item.section] = []
    }
    acc[item.section].push(item)
    return acc
  }, {})

  const renderContent = () => {
    switch (activeTab) {
      case 'quick':
        return <QuickCalc />
      case 'batch':
        return <BatchAnalysis />
      case 'pallet':
        return <PalletAnalysis />
      case 'excel':
        return <ExcelGenerator />
      case 'settings':
        return <Settings />
      case 'translator':
        return <Translator />
      case 'inputTool':
        return <InputTool />
      case 'pairs':
        return <PairsAnalysis />
      case 'autodb':
        return <AutoDB />
      case 'palletOpt':
        return <PalletOpt />
      case 'operatorReport':
        return <OperatorReport />
      default:
        return (
          <div className="card text-center">
            <p className="text-muted mb-4">Раздел в разработке</p>
            <p className="text-muted">Выбранный раздел: {activeTab}</p>
          </div>
        )
    }
  }

  return (
    <div className="app-layout">
      <button
        className="mobile-menu-btn"
        onClick={() => setSidebarOpen(!sidebarOpen)}
      >
        <svg width="20" height="20" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          {sidebarOpen ? (
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
          ) : (
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 6h16M4 12h16M4 18h16" />
          )}
        </svg>
      </button>

      <aside className={`sidebar ${sidebarOpen ? 'open' : ''}`}>
        <div className="logo">
          <div className="logo-content">
            <div>
              <div className="logo-text">WareCore</div>
              <div className="logo-subtitle">Reports System</div>
            </div>
          </div>
        </div>

        <nav className="nav">
          {Object.entries(groupedNavItems).map(([section, items]) => (
            <div key={section} className="nav-section">
              <div className="nav-section-title">{sectionTitles[section]}</div>
              {items.map((item) => (
                <button
                  key={item.id}
                  onClick={() => {
                    setActiveTab(item.id)
                    localStorage.setItem('activeTab', item.id)
                    setSidebarOpen(false)
                  }}
                  className={`nav-item ${activeTab === item.id ? 'active' : ''}`}
                >
                  {icons[item.icon]}
                  <span>{item.label}</span>
                </button>
              ))}
            </div>
          ))}
        </nav>

      </aside>

      {sidebarOpen && (
        <div className="sidebar-overlay open" onClick={() => setSidebarOpen(false)} />
      )}

      <main className="main-content">
        <div className="header">
          <div>
            <h1 className="header-title">
              {navItems.find(item => item.id === activeTab)?.label || 'Dashboard'}
            </h1>
          </div>
        </div>

        {renderContent()}
      </main>
    </div>
  )
}

export default Dashboard
