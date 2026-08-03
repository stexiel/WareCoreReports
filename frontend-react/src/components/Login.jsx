import { useState } from 'react'

function Login({ onLogin }) {
  const [username, setUsername] = useState('')
  const [password, setPassword] = useState('')
  const [error, setError] = useState('')
  const [loading, setLoading] = useState(false)

  const getCSRFToken = async () => {
    try {
      const response = await fetch('http://localhost:3000/api/csrf-token')
      const data = await response.json()
      return data.csrfToken
    } catch (err) {
      console.error('Ошибка получения CSRF токена:', err)
      return null
    }
  }

  const handleSubmit = async (e) => {
    e.preventDefault()
    setError('')
    setLoading(true)

    try {
      const csrfToken = await getCSRFToken()
      
      const headers = {
        'Content-Type': 'application/json',
      }
      
      if (csrfToken) {
        headers['X-CSRF-Token'] = csrfToken
      }

      const response = await fetch('http://localhost:3000/api/login', {
        method: 'POST',
        headers,
        body: JSON.stringify({ username, password }),
      })

      const data = await response.json()

      if (!response.ok) {
        throw new Error(data.error || 'Ошибка входа')
      }

      localStorage.setItem('authToken', data.token)
      localStorage.setItem('username', data.username)
      localStorage.setItem('userRole', data.role)
      localStorage.setItem('userId', data.id)

      onLogin({ username: data.username, role: data.role, id: data.id })
    } catch (err) {
      setError(err.message)
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="login-page">
      <div className="login-card">
        <div className="login-header">
          <img src="/icon.png" className="login-logo" alt="WareCore Logo" />
          <h1 className="login-title">WareCore Reports</h1>
          <p className="login-subtitle">Система отчетности склада</p>
        </div>

        {error && (
          <div className="error-banner">{error}</div>
        )}

        <form onSubmit={handleSubmit}>
          <div className="form-group">
            <label className="form-label">Логин</label>
            <input
              type="text"
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              className="form-input"
              placeholder="Введите логин"
              required
              autoComplete="username"
            />
          </div>

          <div className="form-group">
            <label className="form-label">Пароль</label>
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              className="form-input"
              placeholder="Введите пароль"
              required
              autoComplete="current-password"
            />
          </div>

          <button
            type="submit"
            disabled={loading}
            className="btn btn-primary btn-block"
          >
            {loading ? 'Вход...' : 'Войти'}
          </button>
        </form>
      </div>
    </div>
  )
}

export default Login
