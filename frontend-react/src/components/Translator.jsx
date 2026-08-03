import { useState, useEffect } from 'react'

function Translator() {
  const [sourceText, setSourceText] = useState('')
  const [translatedText, setTranslatedText] = useState('')
  const [sourceLang, setSourceLang] = useState('zh')
  const [targetLang, setTargetLang] = useState('ru')
  const [isTranslating, setIsTranslating] = useState(false)
  const [history, setHistory] = useState([])

  const langNames = { zh: '🇨🇳 Китайский', en: '🇬🇧 Английский', ru: '🇷🇺 Русский' }

  useEffect(() => {
    const savedHistory = localStorage.getItem('translationHistory')
    if (savedHistory) {
      setHistory(JSON.parse(savedHistory))
    }
  }, [])

  const splitTextIntoChunks = (text, maxLength) => {
    const chunks = []
    const sentences = text.split(/[.!?。！？]/)
    let currentChunk = ''
    
    for (const sentence of sentences) {
      if ((currentChunk + sentence).length > maxLength && currentChunk) {
        chunks.push(currentChunk.trim())
        currentChunk = sentence
      } else {
        currentChunk += sentence + '.'
      }
    }
    
    if (currentChunk.trim()) {
      chunks.push(currentChunk.trim())
    }
    
    return chunks
  }

  const translateText = async () => {
    if (!sourceText.trim()) {
      setTranslatedText('Введите текст для перевода')
      return
    }

    if (sourceLang === targetLang) {
      setTranslatedText(sourceText)
      return
    }

    setIsTranslating(true)
    setTranslatedText('Перевод...')

    try {
      const langPair = `${sourceLang}|${targetLang}`
      const chunks = splitTextIntoChunks(sourceText, 400)
      let translatedParts = []
      
      for (let i = 0; i < chunks.length; i++) {
        const chunk = chunks[i]
        setTranslatedText(`Перевод... (${i + 1}/${chunks.length})`)
        
        const url = `https://api.mymemory.translated.net/get?q=${encodeURIComponent(chunk)}&langpair=${langPair}`
        const response = await fetch(url)
        const data = await response.json()
        
        if (data.responseStatus === 200) {
          translatedParts.push(data.responseData.translatedText)
        } else {
          translatedParts.push(`[Ошибка перевода части ${i + 1}]`)
        }
        
        if (i < chunks.length - 1) await new Promise(r => setTimeout(r, 300))
      }
      
      const translated = translatedParts.join('\n')
      setTranslatedText(translated)
      addToTranslationHistory(sourceText, translated, sourceLang, targetLang)
    } catch (error) {
      setTranslatedText('Ошибка соединения. Проверьте интернет.')
    } finally {
      setIsTranslating(false)
    }
  }

  const addToTranslationHistory = (source, translated, fromLang, toLang) => {
    const newHistory = [...history]
    newHistory.unshift({
      source,
      translated,
      fromLang,
      toLang,
      date: new Date().toISOString()
    })
    if (newHistory.length > 50) newHistory.pop()
    setHistory(newHistory)
    localStorage.setItem('translationHistory', JSON.stringify(newHistory))
  }

  const copyTranslation = (index) => {
    const item = history[index]
    if (!item) return

    navigator.clipboard.writeText(item.translated).then(() => {
      alert('Перевод скопирован в буфер обмена')
    })
  }

  const clearHistory = () => {
    setHistory([])
    localStorage.removeItem('translationHistory')
  }

  return (
    <div className="card">
      <div className="card-header">
        <h2 className="card-title">Переводчик</h2>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px', marginBottom: '16px' }}>
        <div className="form-group">
          <label className="form-label">Исходный язык</label>
          <select
            value={sourceLang}
            onChange={(e) => setSourceLang(e.target.value)}
            className="form-input"
          >
            <option value="zh">🇨🇳 Китайский</option>
            <option value="en">🇬🇧 Английский</option>
            <option value="ru">🇷🇺 Русский</option>
          </select>
        </div>
        <div className="form-group">
          <label className="form-label">Целевой язык</label>
          <select
            value={targetLang}
            onChange={(e) => setTargetLang(e.target.value)}
            className="form-input"
          >
            <option value="ru">🇷🇺 Русский</option>
            <option value="en">🇬🇧 Английский</option>
            <option value="zh">🇨🇳 Китайский</option>
          </select>
        </div>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px', marginBottom: '16px' }}>
        <div className="form-group">
          <label className="form-label">Исходный текст</label>
          <textarea
            value={sourceText}
            onChange={(e) => setSourceText(e.target.value)}
            className="form-input"
            rows="6"
            placeholder="Введите текст для перевода..."
            style={{ resize: 'vertical', fontSize: '18px', fontWeight: 700 }}
          />
        </div>
        <div className="form-group">
          <label className="form-label">Перевод</label>
          <textarea
            value={translatedText}
            readOnly
            className="form-input"
            rows="6"
            placeholder="Перевод появится здесь..."
            style={{ resize: 'vertical', background: 'var(--background)', fontSize: '18px', fontWeight: 700 }}
          />
        </div>
      </div>

      <button
        onClick={translateText}
        className="btn btn-primary"
        disabled={isTranslating}
        style={{ width: '100%', padding: '12px' }}
      >
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ marginRight: '8px' }}>
          <path d="M5 8l6 6"/>
          <path d="M4 14l6-6 2-3"/>
          <path d="M2 5h12"/>
          <path d="M7 2h1"/>
          <path d="M22 22l-5-10-5 10"/>
          <path d="M14 18h6"/>
        </svg>
        Перевести
      </button>

      <div style={{ marginTop: '24px' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px' }}>
          <h3 style={{ fontSize: '16px', fontWeight: 600, margin: 0, color: 'var(--text-primary)' }}>
            История переводов
          </h3>
          {history.length > 0 && (
            <button onClick={clearHistory} className="btn btn-secondary" style={{ padding: '6px 12px', fontSize: '12px' }}>
              Очистить
            </button>
          )}
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
          {history.length === 0 ? (
            <p style={{ color: 'var(--text-secondary)', fontSize: '14px' }}>История пуста</p>
          ) : (
            history.map((item, index) => (
              <div
                key={index}
                style={{
                  background: 'var(--surface)',
                  border: '1px solid var(--border)',
                  borderRadius: '8px',
                  padding: '12px',
                  cursor: 'pointer'
                }}
                onClick={() => copyTranslation(index)}
              >
                <div style={{ fontSize: '12px', color: 'var(--text-secondary)', marginBottom: '8px' }}>
                  {langNames[item.fromLang]} → {langNames[item.toLang]} • {new Date(item.date).toLocaleString('ru-RU')}
                </div>
                <div style={{ fontWeight: 600, color: 'var(--text-primary)', marginBottom: '4px' }}>
                  {item.source.substring(0, 100)}{item.source.length > 100 ? '...' : ''}
                </div>
                <div style={{ color: 'var(--text-secondary)' }}>
                  {item.translated.substring(0, 100)}{item.translated.length > 100 ? '...' : ''}
                </div>
              </div>
            ))
          )}
        </div>
      </div>
    </div>
  )
}

export default Translator
