/// <reference types="node" />

import { SecretsManagerClient, GetSecretValueCommand } from '@aws-sdk/client-secrets-manager'

const graphBase = 'https://graph.microsoft.com/v1.0'
const secretClient = new SecretsManagerClient({})
let cachedSecret: { tenantId: string; clientId: string; clientSecret: string } | null = null

type Slot = { date: string; time: string }

export const handler = async (event: any) => {
  try {
    if (event.httpMethod === 'GET') return response(200, await getAvailability(event.queryStringParameters?.month))
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
  const created = await graphRequest(`/users/${encodeURIComponent(process.env.CALENDAR_OWNER!)}/events`, {
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
  const data = await graphRequest(`/users/${encodeURIComponent(process.env.CALENDAR_OWNER!)}/calendarView?${params}`, {
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
  const body = new URLSearchParams({
    client_id: secret.clientId,
    client_secret: secret.clientSecret,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  })
  const result = await fetch(`https://login.microsoftonline.com/${secret.tenantId}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body,
  })
  if (!result.ok) throw new Error(`Microsoft token request failed: ${result.status}`)
  return (await result.json()).access_token
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
