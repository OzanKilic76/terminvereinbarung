import { PublicClientApplication } from '@azure/msal-browser'

const clientId = import.meta.env.VITE_MICROSOFT_CLIENT_ID
const tenantId = import.meta.env.VITE_MICROSOFT_TENANT_ID
const redirectUri = window.location.origin
const scopes = ['Calendars.ReadWrite']

export const microsoftConfigured = Boolean(clientId && tenantId)

const msal = microsoftConfigured ? new PublicClientApplication({
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri,
  },
  cache: {
    cacheLocation: 'localStorage',
  },
}) : null

let initialized = false

async function getClient() {
  if (!msal) throw new Error('Microsoft Outlook ist noch nicht konfiguriert.')
  if (!initialized) {
    await msal.initialize()
    initialized = true
  }
  return msal
}

export async function signInMicrosoft() {
  const client = await getClient()
  const result = await client.loginPopup({ scopes })
  return result.account
}

export async function getMicrosoftAccount() {
  if (!msal) return null
  const client = await getClient()
  return client.getAllAccounts()[0] || null
}

async function getToken() {
  const client = await getClient()
  const account = await getMicrosoftAccount()
  if (!account) throw new Error('Bitte zuerst mit Microsoft Outlook verbinden.')
  try {
    const result = await client.acquireTokenSilent({ account, scopes })
    return result.accessToken
  } catch {
    const result = await client.acquireTokenPopup({ account, scopes })
    return result.accessToken
  }
}

async function graphRequest(path, options = {}) {
  const token = await getToken()
  const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...(options.headers || {}),
    },
  })
  if (!response.ok) {
    const details = await response.text()
    throw new Error(`Microsoft Graph: ${response.status} ${details}`)
  }
  return response.status === 204 ? null : response.json()
}

export async function getCalendarEvents(start, end) {
  const params = new URLSearchParams({
    startDateTime: start.toISOString(),
    endDateTime: end.toISOString(),
    '$select': 'start,end,showAs',
    '$top': '500',
  })
  const data = await graphRequest(`/me/calendarView?${params}`)
  return data.value || []
}

export async function createTeamsAppointment({ date, time, name, email, phone, comment }) {
  const [hours, minutes] = time.split(':').map(Number)
  const start = new Date(date)
  start.setHours(hours, minutes, 0, 0)
  const end = new Date(start.getTime() + 60 * 60 * 1000)
  const body = {
    subject: `Beratungsgespräch – ${name}`,
    body: {
      contentType: 'HTML',
      content: `<p>Beratungsgespräch zur Baufinanzierung.</p><p>Telefon: ${phone}</p><p>${comment || ''}</p>`,
    },
    start: { dateTime: toGraphDateTime(start), timeZone: 'W. Europe Standard Time' },
    end: { dateTime: toGraphDateTime(end), timeZone: 'W. Europe Standard Time' },
    isOnlineMeeting: true,
    onlineMeetingProvider: 'teamsForBusiness',
    attendees: [{
      emailAddress: { address: email, name },
      type: 'required',
    }],
  }
  return graphRequest('/me/events', { method: 'POST', body: JSON.stringify(body) })
}

function toGraphDateTime(date) {
  const pad = (value) => String(value).padStart(2, '0')
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}T${pad(date.getHours())}:${pad(date.getMinutes())}:00`
}
