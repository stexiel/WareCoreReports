import { useState } from 'react'

function Profile({ user }) {
  const [currentPassword, setCurrentPassword] = useState('')
  const [newPassword, setNewPassword] = useState('')
  const [confirmPassword, setConfirmPassword] = useState('')
  const [message, setMessage] = useState('')
  const [messageType, setMessageType] = useState('')

  const getApiBase = () => 'http://localhost:3000'
  const getAuthToken = () => localStorage.getItem('authToken')

  const getCSRFToken = async () => {
    try {
      const response = await fetch(`${getApiBase()}/api/csrf-token`)
      const data = await response.json()
      return data.csrfToken
    } catch (err) {
      console.error('Ошибка получения CSRF токена:', err)
      return null
    }
  }

  const showMessage = (text, type) => {
    setMessage(text)
    setMessageType(type)
    setTimeout(() => {
      setMessage('')
      setMessageType('')
    }, 5000)
  }

  const handlePasswordChange = async (e) => {
    e.preventDefault()
    setMessage('')

    if (newPassword !== confirmPassword) {
      showMessage('Пароли не совпадают', 'error')
      return
    }

    if (newPassword.length < 4) {
      showMessage('Пароль должен быть минимум 4 символа', 'error')
      return
    }

    try {
      const token = getAuthToken()
      const csrfToken = await getCSRFToken()
      
      const headers = {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${token}`
      }
      
      if (csrfToken) {
        headers['X-CSRF-Token'] = csrfToken
      }

      const response = await fetch(`${getApiBase()}/api/users/${user.id}`, {
        method: 'PUT',
        headers,
        body: JSON.stringify({ 
          username: user.username,
          password: newPassword,
          role: user.role
        })
      })

      const data = await response.json()

      if (!response.ok) {
        throw new Error(data.error || 'Ошибка изменения пароля')
      }

      showMessage('Пароль успешно изменен', 'success')
      setCurrentPassword('')
      setNewPassword('')
      setConfirmPassword('')
    } catch (err) {
      showMessage(err.message, 'error')
    }
  }

  const getRoleLabel = (role) => {
    switch(role) {
      case 'admin': return 'Администратор'
      case 'operator': return 'Оператор'
      case 'warehouseman': return 'Кладовщик'
      default: return role
    }
  }

  const getRoleClass = (role) => {
    switch(role) {
      case 'admin': return 'badge badge-red'
      case 'operator': return 'badge badge-blue'
      case 'warehouseman': return 'badge badge-green'
      default: return 'badge'
    }
  }

  return (
    <div className="card">
      <div className="card-header">
        <h2 className="card-title">Профиль пользователя</h2>
      </div>

      <div className="flex items-center gap-4 mb-4">
        <div className="avatar" style={{ width: '72px', height: '72px', fontSize: '28px' }}>
          {user.username.charAt(0).toUpperCase()}
        </div>
        <div>
          <h3 style={{ fontSize: '20px', fontWeight: 700 }}>{user.username}</h3>
          <p className="text-muted">Пользователь системы WareCore Reports</p>
          <span className={getRoleClass(user.role)} style={{ marginTop: '8px', display: 'inline-flex' }}>
            {getRoleLabel(user.role)}
          </span>
        </div>
      </div>

      <div className="stats-grid">
        <div className="stat-card">
          <div className="stat-label">Логин</div>
          <div style={{ fontSize: '18px', fontWeight: 600 }}>{user.username}</div>
        </div>
        <div className="stat-card stat-success">
          <div className="stat-label">Роль</div>
          <div style={{ fontSize: '18px', fontWeight: 600 }}>{getRoleLabel(user.role)}</div>
        </div>
        <div className="stat-card stat-warning">
          <div className="stat-label">ID пользователя</div>
          <div style={{ fontSize: '18px', fontWeight: 600 }}>{user.id}</div>
        </div>
      </div>

      <div style={{ borderTop: '1px solid var(--border)', paddingTop: '20px' }}>
        <h3 className="card-title mb-4" style={{ fontSize: '16px' }}>Изменить пароль</h3>

        {message && (
          <div className={messageType === 'success' ? 'badge badge-green' : 'error-banner'} style={messageType === 'success' ? { display: 'block', padding: '12px 16px', marginBottom: '16px' } : {}}>
            {message}
          </div>
        )}

        <form onSubmit={handlePasswordChange} style={{ maxWidth: '400px' }}>
          <div className="form-group">
            <label className="form-label">Новый пароль</label>
            <input
              type="password"
              value={newPassword}
              onChange={(e) => setNewPassword(e.target.value)}
              className="form-input"
              required
              minLength={4}
            />
          </div>

          <div className="form-group">
            <label className="form-label">Подтвердите новый пароль</label>
            <input
              type="password"
              value={confirmPassword}
              onChange={(e) => setConfirmPassword(e.target.value)}
              className="form-input"
              required
              minLength={4}
            />
          </div>

          <button type="submit" className="btn btn-primary">
            Изменить пароль
          </button>
        </form>
      </div>
    </div>
  )
}

export default Profile
