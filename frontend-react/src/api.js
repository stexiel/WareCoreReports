const API_BASE = 'http://172.30.1.249:3000'

export const apiFetch = async (url, options = {}) => {
  const headers = {
    'Content-Type': 'application/json',
    ...options.headers
  }

  return fetch(`${API_BASE}${url}`, {
    ...options,
    headers
  })
}

export const getApiBase = () => API_BASE

// Проверка доступности backend. Используется страницами, которым нужен backend/БД
// (AutoDB, OperatorReport), чтобы показать понятную ошибку, если backend не запущен.
export const checkBackendHealth = async (timeoutMs = 3000) => {
  try {
    const controller = new AbortController()
    const timer = setTimeout(() => controller.abort(), timeoutMs)
    const response = await fetch(`${API_BASE}/api/health`, { signal: controller.signal })
    clearTimeout(timer)
    return response.ok
  } catch (err) {
    return false
  }
}
