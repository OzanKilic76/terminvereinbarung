/// <reference types="node" />

import { randomBytes, createHmac } from 'node:crypto'
import { SecretsManagerClient, GetSecretValueCommand, PutSecretValueCommand } from '@aws-sdk/client-secrets-manager'

const graphBase = 'https://graph.microsoft.com/v1.0'
const secretClient = new SecretsManagerClient({})
type SecretValues = {
  tenantId: string
  clientId: string
  clientSecret: string
  accessToken?: string
  refreshToken?: string
  tokenExpiresAt?: number
}

let cachedSecret: SecretValues | null = null

type Slot = { date: string; time: string }

export const handler = async (event: any) => {
  try {
    if (event.httpMethod === 'GET') {
      const query = event.queryStringParameters || {}
      if (query.action === 'connect') return redirect(await getAuthorizationUrl())
      if (query.action === 'callback') return await handleOAuthCallback(query)
      return response(200, await getAvailability(query.month))
    }
    if (event.httpMethod === 'POST') return response(200, await createBooking(JSON.parse(event.body || '{}')))
    return response(405, { message: 'Method not allowed' })
  } catch (error) {
    console.error(error)
    return response(500, { message: error instanceof Error ? error.message : 'Booking request failed' })
  }
}

async function getAvailability(month?: string) {
  const targetMonth = /^\d{4}-\d{2}$/.test(month || '') ? month! : currentMonth()
  const events = await getCalendarView(targetMonth)
  const slots = buildSlots(targetMonth).filter((slot) => !slotOverlaps(slot, events))
  return { month: targetMonth, slots }
}

async function getAuthorizationUrl() {
  const secret = await getSecret()
  const state = createState(secret.clientSecret)
  const params = new URLSearchParams({
    client_id: secret.clientId,
    response_type: 'code',
    redirect_uri: process.env.OAUTH_REDIRECT_URI!,
    response_mode: 'query',
    scope: 'offline_access User.Read Calendars.ReadWrite',
    state,
  })
  return {
    authorizationUrl: `https://login.microsoftonline.com/${process.env.OAUTH_AUTHORITY}/oauth2/v2.0/authorize?${params}`,
    message: 'Öffne authorizationUrl einmalig mit deinem Microsoft-Konto.',
  }
}

async function handleOAuthCallback(query: Record<string, string>) {
  if (!query.code || !query.state) return response(400, { message: 'OAuth-Code oder Status fehlt.' })
  const secret = await getSecret()
  if (!verifyState(query.state, secret.clientSecret)) return response(400, { message: 'Ungültiger oder abgelaufener OAuth-Status.' })

  const tokenResponse = await fetch(`https://login.microsoftonline.com/${process.env.OAUTH_AUTHORITY}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: secret.clientId,
      client_secret: secret.clientSecret,
      code: query.code,
      redirect_uri: process.env.OAUTH_REDIRECT_URI!,
      grant_type: 'authorization_code',
      scope: 'offline_access User.Read Calendars.ReadWrite',
    }),
  })
  if (!tokenResponse.ok) return response(502, { message: `Microsoft OAuth fehlgeschlagen: ${await tokenResponse.text()}` })
  const tokens = await tokenResponse.json()
  await saveSecret({ ...secret, accessToken: tokens.access_token, refreshToken: tokens.refresh_token, tokenExpiresAt: Date.now() + (tokens.expires_in * 1000) })
  return response(200, { message: 'Outlook erfolgreich verbunden. Kunden benötigen keinen Microsoft-Login.' })
}

async function createBooking(input: any) {
  const { name, email, phone, comment, date, time } = input
  if (!name?.trim() || !email?.trim() || !phone?.trim() || !/^\d{4}-\d{2}-\d{2}$/.test(date || '') || !isAllowedTime(time)) {
    throw new Error('Ungültige Buchungsdaten.')
  }
  if (!isMondayOrFriday(date)) throw new Error('Termine sind nur montags und freitags verfügbar.')

  const events = await getCalendarView(date.slice(0, 7))
  if (slotOverlaps({ date, time }, events)) throw new Error('Dieser Termin ist inzwischen nicht mehr frei.')

  const start = `${date}T${time}:00`
  const end = `${date}T${addHour(time)}:00`
  const created = await graphRequest('/me/events', {
    method: 'POST',
    body: JSON.stringify({
      subject: `Beratungsgespräch – ${name.trim()}`,
      body: {
        contentType: 'HTML',
        content: `<p>Beratungsgespräch zur Baufinanzierung.</p><p>Telefon: ${escapeHtml(phone)}</p><p>${escapeHtml(comment || '')}</p>`,
      },
      start: { dateTime: start, timeZone: 'W. Europe Standard Time' },
      end: { dateTime: end, timeZone: 'W. Europe Standard Time' },
      isOnlineMeeting: true,
      onlineMeetingProvider: 'teamsForBusiness',
      attendees: [{ emailAddress: { address: email.trim(), name: name.trim() }, type: 'required' }],
    }),
  })

  return { ok: true, joinUrl: created.onlineMeeting?.joinUrl || '' }
}

async function getCalendarView(month: string) {
  const [year, monthNumber] = month.split('-').map(Number)
  const start = `${month}-01T00:00:00`
  const lastDay = new Date(Date.UTC(year, monthNumber, 0)).getUTCDate()
  const end = `${month}-${String(lastDay).padStart(2, '0')}T23:59:59`
  const params = new URLSearchParams({ startDateTime: start, endDateTime: end, '$select': 'start,end,showAs', '$top': '500' })
  const data = await graphRequest(`/me/calendarView?${params}`, {
    headers: { Prefer: 'outlook.timezone="W. Europe Standard Time"' },
  })
  return data.value || []
}

async function graphRequest(path: string, options: RequestInit = {}) {
  const token = await getGraphToken()
  const result = await fetch(`${graphBase}${path}`, {
    ...options,
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', ...(options.headers || {}) },
  })
  if (!result.ok) throw new Error(`Microsoft Graph returned ${result.status}: ${await result.text()}`)
  return result.json()
}

async function getGraphToken() {
  const secret = await getSecret()
  if (secret.refreshToken) {
    if (secret.accessToken && secret.tokenExpiresAt && secret.tokenExpiresAt > Date.now() + 60_000) return secret.accessToken
    const result = await fetch(`https://login.microsoftonline.com/${process.env.OAUTH_AUTHORITY}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: secret.clientId,
        client_secret: secret.clientSecret,
        refresh_token: secret.refreshToken,
        scope: 'offline_access User.Read Calendars.ReadWrite',
        grant_type: 'refresh_token',
      }),
    })
    if (!result.ok) throw new Error(`Microsoft token refresh failed: ${result.status}`)
    const tokens = await result.json()
    await saveSecret({ ...secret, accessToken: tokens.access_token, refreshToken: tokens.refresh_token || secret.refreshToken, tokenExpiresAt: Date.now() + (tokens.expires_in * 1000) })
    return tokens.access_token
  }
  throw new Error('Outlook ist noch nicht verbunden. Öffne zuerst die Connect-URL für den Administrator.')
}

async function saveSecret(value: SecretValues) {
  await secretClient.send(new PutSecretValueCommand({
    SecretId: process.env.AZURE_SECRET_NAME,
    SecretString: JSON.stringify(value),
  }))
  cachedSecret = value
}

function redirect(data: { authorizationUrl: string }) {
  return {
    statusCode: 302,
    headers: { Location: data.authorizationUrl, 'Cache-Control': 'no-store' },
    body: '',
  }
}

function createState(clientSecret: string) {
  const payload = `${Date.now()}.${randomBytes(24).toString('hex')}`
  const signature = createHmac('sha256', clientSecret).update(payload).digest('hex')
  return Buffer.from(`${payload}.${signature}`).toString('base64url')
}

function verifyState(state: string, clientSecret: string) {
  try {
    const decoded = Buffer.from(state, 'base64url').toString()
    const parts = decoded.split('.')
    const signature = parts.pop()
    const payload = parts.join('.')
    if (!signature || !payload || Date.now() - Number(payload.split('.')[0]) > 10 * 60 * 1000) return false
    const expected = createHmac('sha256', clientSecret).update(payload).digest('hex')
    return signature === expected
  } catch {
    return false
  }
}

async function getSecret() {
  if (cachedSecret) return cachedSecret
  const result = await secretClient.send(new GetSecretValueCommand({ SecretId: process.env.AZURE_SECRET_NAME }))
  cachedSecret = JSON.parse(result.SecretString || '{}')
  if (!cachedSecret?.tenantId || !cachedSecret?.clientId || !cachedSecret?.clientSecret) throw new Error('AWS Secret enthält nicht die erwarteten Azure-Werte.')
  return cachedSecret
}

function buildSlots(month: string): Slot[] {
  const [year, monthNumber] = month.split('-').map(Number)
  const lastDay = new Date(Date.UTC(year, monthNumber, 0)).getUTCDate()
  const result: Slot[] = []
  for (let day = 1; day <= lastDay; day += 1) {
    const date = `${month}-${String(day).padStart(2, '0')}`
    if (isMondayOrFriday(date)) for (const hour of [19, 20, 21]) result.push({ date, time: `${hour}:00` })
  }
  return result
}

function slotOverlaps(slot: Slot, events: any[]) {
  const start = `${slot.date}T${slot.time}:00`
  const end = `${slot.date}T${addHour(slot.time)}:00`
  return events.some((event) => {
    if (event.showAs === 'free' || !event.start?.dateTime || !event.end?.dateTime) return false
    const eventStart = normalizeDateTime(event.start.dateTime)
    const eventEnd = normalizeDateTime(event.end.dateTime)
    return eventStart < end && eventEnd > start
  })
}

function normalizeDateTime(value: string) {
  return value.replace(/\.(\d{3})\d+/, '.$1').replace(/\.\d{3}$/, '')
}

function addHour(time: string) {
  return `${String(Number(time.slice(0, 2)) + 1).padStart(2, '0')}:00`
}

function isAllowedTime(time: string) {
  return time === '19:00' || time === '20:00' || time === '21:00'
}

function isMondayOrFriday(date: string) {
  const day = new Date(`${date}T12:00:00`).getDay()
  return day === 1 || day === 5
}

function currentMonth() {
  return new Date().toISOString().slice(0, 7)
}

function escapeHtml(value: string) {
  return value.replace(/[&<>"']/g, (character) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' })[character]!)
}

function response(statusCode: number, body: unknown) {
  return {
    statusCode,
    headers: { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': 'Content-Type', 'Access-Control-Allow-Methods': 'GET,POST,OPTIONS' },
    body: JSON.stringify(body),
  }
}
