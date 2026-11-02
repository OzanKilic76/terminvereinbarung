import outputs from '../amplify_outputs.json'

const apiUrl = import.meta.env.VITE_BOOKING_API_URL || outputs.custom?.bookingApiUrl || ''
export const bookingApiConfigured = Boolean(apiUrl)

async function request(path, options) {
  if (!apiUrl) throw new Error('Die Buchungs-API ist noch nicht konfiguriert.')
  const response = await fetch(`${apiUrl}${path}`, {
    ...options,
    headers: { 'Content-Type': 'application/json', ...(options?.headers || {}) },
  })
  const body = await response.json()
  if (!response.ok) throw new Error(body.message || 'Die Buchung konnte nicht verarbeitet werden.')
  return body
}

export function getAvailability(month) {
  return request(`?month=${encodeURIComponent(month)}`)
}

export function createBooking(data) {
  return request('', { method: 'POST', body: JSON.stringify(data) })
}
