import { useState, useEffect } from 'react'

function Admin({ user }) {
  const [users, setUsers] = useState([])
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')
  const [showModal, setShowModal] = useState(false)
  const [editingUser, setEditingUser] = useState(null)
  const [formData, setFormData] = useState({ username: '', password: '', role: 'operator' })

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

  const loadUsers = async () => {
    setLoading(true)
    setError('')

    try {
      const response = await fetch(`${getApiBase()}/api/users`, {
        headers: {
          'Authorization': `Bearer ${getAuthToken()}`
        }
      })

      if (!response.ok) {
        throw new Error('Ошибка загрузки пользователей')
      }

      const data = await response.json()
      setUsers(data)
    } catch (err) {
      setError(err.message)
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => {
    loadUsers()
  }, [])

  const openCreateModal = () => {
    setEditingUser(null)
    setFormData({ username: '', password: '', role: 'operator' })
    setShowModal(true)
  }

  const openEditModal = (user) => {
    setEditingUser(user)
    setFormData({ username: user.username, password: '', role: user.role })
    setShowModal(true)
  }

  const closeModal = () => {
    setShowModal(false)
    setEditingUser(null)
    setFormData({ username: '', password: '', role: 'operator' })
  }

  const handleSubmit = async (e) => {
    e.preventDefault()
    setError('')

    try {
      const token = getAuthToken()
      const csrfToken = await getCSRFToken()
      
      let url = `${getApiBase()}/api/users`
      let method = 'POST'

      if (editingUser) {
        url += `/${editingUser.id}`
        method = 'PUT'
      }

      const headers = {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${token}`
      }
      
      if (csrfToken) {
        headers['X-CSRF-Token'] = csrfToken
      }

      const response = await fetch(url, {
        method,
        headers,
        body: JSON.stringify(formData)
      })

      const data = await response.json()

      if (!response.ok) {
        throw new Error(data.error || 'Ошибка сохранения')
      }

      closeModal()
      loadUsers()
    } catch (err) {
      setError(err.message)
    }
  }

  const deleteUser = async (userId) => {
    if (!confirm('Вы уверены, что хотите удалить этого пользователя?')) {
      return
    }

    try {
      const csrfToken = await getCSRFToken()
      
      const headers = {
        'Authorization': `Bearer ${getAuthToken()}`
      }
      
      if (csrfToken) {
        headers['X-CSRF-Token'] = csrfToken
      }

      const response = await fetch(`${getApiBase()}/api/users/${userId}`, {
        method: 'DELETE',
        headers
      })

      if (!response.ok) {
        const data = await response.json()
        throw new Error(data.error || 'Ошибка удаления')
      }

      loadUsers()
    } catch (err) {
      setError(err.message)
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
    <>
      <div className="card">
        <div className="card-header">
          <h2 className="card-title">Управление пользователями</h2>
          <button onClick={openCreateModal} className="btn btn-primary">
            Добавить пользователя
          </button>
        </div>

        {error && <div className="error-banner">{error}</div>}

        {loading ? (
          <div className="text-center text-muted" style={{ padding: '32px 0' }}>Загрузка...</div>
        ) : (
          <div className="table-container">
            <table>
              <thead>
                <tr>
                  <th>ID</th>
                  <th>Логин</th>
                  <th>Роль</th>
                  <th>Действия</th>
                </tr>
              </thead>
              <tbody>
                {users.map((u) => (
                  <tr key={u.id}>
                    <td>{u.id}</td>
                    <td>{u.username}</td>
                    <td>
                      <span className={getRoleClass(u.role)}>{getRoleLabel(u.role)}</span>
                    </td>
                    <td>
                      <div className="flex gap-2">
                        <button
                          onClick={() => openEditModal(u)}
                          className="btn btn-secondary"
                          style={{ padding: '6px 12px' }}
                          disabled={u.username === 'stexiel'}
                        >
                          Редактировать
                        </button>
                        <button
                          onClick={() => deleteUser(u.id)}
                          className="btn btn-danger"
                          style={{ padding: '6px 12px' }}
                          disabled={u.username === 'stexiel'}
                        >
                          Удалить
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>

            {users.length === 0 && (
              <div className="text-center text-muted" style={{ padding: '32px 0' }}>
                Нет пользователей
              </div>
            )}
          </div>
        )}
      </div>

      {showModal && (
        <div className="modal">
          <div className="modal-content">
            <div className="modal-header">
              <h3>{editingUser ? 'Редактировать пользователя' : 'Добавить пользователя'}</h3>
            </div>

            <form onSubmit={handleSubmit}>
              <div className="form-group">
                <label className="form-label">Логин</label>
                <input
                  type="text"
                  value={formData.username}
                  onChange={(e) => setFormData({ ...formData, username: e.target.value })}
                  className="form-input"
                  required
                  disabled={editingUser?.username === 'stexiel'}
                />
              </div>

              <div className="form-group">
                <label className="form-label">
                  Пароль {editingUser ? '(оставьте пустым для сохранения текущего)' : ''}
                </label>
                <input
                  type="password"
                  value={formData.password}
                  onChange={(e) => setFormData({ ...formData, password: e.target.value })}
                  className="form-input"
                  required={!editingUser}
                />
              </div>

              <div className="form-group">
                <label className="form-label">Роль</label>
                <select
                  value={formData.role}
                  onChange={(e) => setFormData({ ...formData, role: e.target.value })}
                  className="form-select"
                  disabled={editingUser?.username === 'stexiel'}
                >
                  <option value="operator">Оператор</option>
                  <option value="warehouseman">Кладовщик</option>
                  <option value="admin">Администратор</option>
                </select>
              </div>

              <div className="flex gap-3 mt-4">
                <button type="button" onClick={closeModal} className="btn btn-secondary" style={{ flex: 1, justifyContent: 'center' }}>
                  Отмена
                </button>
                <button type="submit" className="btn btn-primary" style={{ flex: 1, justifyContent: 'center' }}>
                  {editingUser ? 'Сохранить' : 'Создать'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </>
  )
}

export default Admin
