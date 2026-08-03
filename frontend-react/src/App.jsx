import { useState, useEffect } from 'react'
import Login from './components/Login'
import Dashboard from './components/Dashboard'
import './index.css'

function App() {
  const [isAuthenticated, setIsAuthenticated] = useState(false)
  const [user, setUser] = useState(null)

  useEffect(() => {
    const token = localStorage.getItem('authToken')
    const username = localStorage.getItem('username')
    const role = localStorage.getItem('userRole')
    const userId = localStorage.getItem('userId')
    
    if (token && username) {
      setIsAuthenticated(true)
      setUser({ username, role, id: userId })
    }
  }, [])

  const handleLogin = (userData) => {
    setIsAuthenticated(true)
    setUser(userData)
  }

  const handleLogout = () => {
    localStorage.removeItem('authToken')
    localStorage.removeItem('username')
    localStorage.removeItem('userRole')
    localStorage.removeItem('userId')
    setIsAuthenticated(false)
    setUser(null)
  }

  if (!isAuthenticated) {
    return <Login onLogin={handleLogin} />
  }

  return <Dashboard user={user} onLogout={handleLogout} />
}

export default App
