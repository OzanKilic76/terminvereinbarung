import { useEffect, useMemo, useState } from 'react'
import './App.css'
import { bookingApiConfigured, createBooking, getAvailability } from './bookingApi'

const weekdays = ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So']
const slots = ['09:00', '10:30', '12:00', '14:00', '15:30', '17:00']

function formatMonth(date) {
  return new Intl.DateTimeFormat('de-DE', { month: 'long', year: 'numeric' }).format(date)
}

function dateKey(date) {
  return date.toISOString().slice(0, 10)
}

function App() {
  const today = useMemo(() => {
    const date = new Date()
    date.setHours(0, 0, 0, 0)
    return date
  }, [])
  const [month, setMonth] = useState(new Date(today.getFullYear(), today.getMonth(), 1))
  const [selectedDate, setSelectedDate] = useState(null)
  const [selectedTime, setSelectedTime] = useState(null)
  const [submitted, setSubmitted] = useState(false)
  const [form, setForm] = useState({ name: '', email: '', phone: '', comment: '' })
  const [availableSlots, setAvailableSlots] = useState([])
  const [isBooking, setIsBooking] = useState(false)
  const [error, setError] = useState('')
  const [meetingUrl, setMeetingUrl] = useState('')

  useEffect(() => {
    if (!bookingApiConfigured) return
    setError('')
    getAvailability(`${month.getFullYear()}-${String(month.getMonth() + 1).padStart(2, '0')}`)
      .then((data) => setAvailableSlots(data.slots || []))
      .catch((requestError) => setError(requestError.message))
  }, [month])

  const calendarDays = useMemo(() => {
    const first = new Date(month.getFullYear(), month.getMonth(), 1)
    const offset = (first.getDay() + 6) % 7
    const daysInMonth = new Date(month.getFullYear(), month.getMonth() + 1, 0).getDate()
    return Array.from({ length: Math.ceil((offset + daysInMonth) / 7) * 7 }, (_, index) => {
      const day = index - offset + 1
      return day > 0 && day <= daysInMonth ? new Date(month.getFullYear(), month.getMonth(), day) : null
    })
  }, [month])

  const isSlotAvailable = (date, time) => {
    if (!date) return false
    return availableSlots.some((slot) => slot.date === dateKey(date) && slot.time === time)
  }
  const isAvailable = (date) => date && date >= today && slots.some((time) => isSlotAvailable(date, time))
  const selectDate = (date) => {
    if (!isAvailable(date)) return
    setSelectedDate(date)
    setSelectedTime(null)
    setSubmitted(false)
  }
  const appointmentReady = selectedDate && selectedTime && form.name.trim() && form.email.trim() && form.phone.trim()

  const submitAppointment = async (event) => {
    event.preventDefault()
    if (!appointmentReady) return
    setIsBooking(true)
    setError('')
    try {
      const createdEvent = await createBooking({ date: dateKey(selectedDate), time: selectedTime, ...form })
      setMeetingUrl(createdEvent.joinUrl || '')
      setSubmitted(true)
    } catch (requestError) {
      setError(requestError.message)
    } finally {
      setIsBooking(false)
    }
  }

  return (
    <main className="booking-page">
      <section className="intro-panel">
        <div className="intro-content">
          <div className="profile-photo">
            <img src="/foto.jpg" alt="Ozan Kilic" />
          </div>
          <p className="eyebrow">Kilic Baufinanzierung</p>
          <h1>Ozan Kilic</h1>
          <h2>Finanzberater<br />Beratungsgespräch</h2>
          <div className="appointment-meta">
            <span><strong>◷</strong> 1 h</span>
            <span><strong>⌁</strong> Einzelheiten zur Webkonferenz nach Bestätigung.</span>
          </div>
          <div className="copy">
            <p>Merhaba,</p>
            <p>🏠 Almanyada gayrimenkul satın alma ve konut kredisi hakkında sorularınız var mı?</p>
            <p>🎁 Konut kredisinde en avantajlı kredi şartları ve sunulan hibe ve devlet desteklerini biliyor musunuz?</p>
            <p>Bu ve benzeri tüm sorularınız için</p>
            <p>her Pazartesi ve Cuma günü sabit danışmanlık saatlerim var (isteğe bağlı olarak Türkçe).</p>
            <p>✅ Ücretsiz danışmanlık</p>
            <p>💻 Online randevu</p>
            <p>⏳ Sınırlı kontenjan – hemen yerinizi ayırın!</p>
            <p className="cta-copy">👉 Randevunuzu buradan alın</p>
          </div>
        </div>
        <footer>© {new Date().getFullYear()} Ozan Kilic</footer>
      </section>

      <section className="calendar-panel">
        <div className="calendar-inner">
          <div className="calendar-heading">
            <div>
              <p className="eyebrow">Termin vereinbaren</p>
              <h3>Wähle einen passenden Termin</h3>
              <p className="muted">Beratungsgespräche finden montags und freitags statt.</p>
            </div>
            <span className={bookingApiConfigured ? 'secure-badge' : 'connect-button'}>{bookingApiConfigured ? '● Kalender live' : 'Kalender wird eingerichtet'}</span>
          </div>
          {!bookingApiConfigured && <p className="integration-hint">Die Kalenderverbindung wird noch eingerichtet.</p>}
          {error && <p className="integration-error">{error}</p>}

          {submitted ? (
            <div className="confirmation">
              <div className="confirmation-icon">✓</div>
              <p className="eyebrow">Termin angefragt</p>
              <h3>Vielen Dank, {form.name}!</h3>
              <p>Dein Termin für den <strong>{selectedDate.toLocaleDateString('de-DE')} um {selectedTime} Uhr</strong> wurde in Outlook eingetragen. Die Bestätigung mit dem Teams-Link erhältst du per E-Mail.</p>
              {meetingUrl && <a className="meeting-link" href={meetingUrl} target="_blank" rel="noreferrer">Teams-Link öffnen →</a>}
              <button className="outline-button" onClick={() => setSubmitted(false)}>Anderen Termin wählen</button>
            </div>
          ) : (
            <form onSubmit={submitAppointment}>
              {!selectedDate ? <>
                <div className="month-nav">
                  <button type="button" aria-label="Vorheriger Monat" onClick={() => setMonth(new Date(month.getFullYear(), month.getMonth() - 1, 1))}>‹</button>
                  <strong>{formatMonth(month)}</strong>
                  <button type="button" aria-label="Nächster Monat" onClick={() => setMonth(new Date(month.getFullYear(), month.getMonth() + 1, 1))}>›</button>
                </div>
                <div className="calendar-grid weekdays">{weekdays.map((day) => <span key={day}>{day}</span>)}</div>
                <div className="calendar-grid days">{calendarDays.map((date, index) => <button type="button" key={date ? dateKey(date) : `empty-${index}`} className={`${date && isAvailable(date) ? 'available' : 'disabled'}`} onClick={() => selectDate(date)} disabled={!isAvailable(date)}>{date?.getDate()}</button>)}</div>
              </> : <div className="chosen-date"><span>Ausgewählter Tag</span><strong>{selectedDate.toLocaleDateString('de-DE', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}</strong><button type="button" onClick={() => { setSelectedDate(null); setSelectedTime(null) }}>← Datum ändern</button></div>}

              {selectedDate && <>
                <div className="selection-section">
                  <div className="section-label"><span>Uhrzeit auswählen</span><small>{selectedDate.toLocaleDateString('de-DE', { weekday: 'long', day: 'numeric', month: 'long' })}</small></div>
                  <div className="time-grid">{slots.map((time) => <button type="button" key={time} className={selectedTime === time ? 'selected' : ''} onClick={() => setSelectedTime(time)} disabled={!isSlotAvailable(selectedDate, time)}>{time}</button>)}</div>
                </div>

                <div className="details-section">
                  <div className="section-label"><span>Deine Kontaktdaten</span><small>für die Terminbestätigung</small></div>
                  <div className="form-grid">
                    <label>Name<input required value={form.name} onChange={(event) => setForm({ ...form, name: event.target.value })} placeholder="Vor- und Nachname" /></label>
                    <label>E-Mail-Adresse *<input required type="email" value={form.email} onChange={(event) => setForm({ ...form, email: event.target.value })} placeholder="name@beispiel.de" /></label>
                    <label>Telefonnummer *<input required type="tel" value={form.phone} onChange={(event) => setForm({ ...form, phone: event.target.value })} placeholder="+49 170 1234567" /></label>
                    <label className="comment-field">Kommentar <textarea value={form.comment} onChange={(event) => setForm({ ...form, comment: event.target.value })} placeholder="Worum geht es in deinem Beratungsgespräch?" rows="3" /></label>
                  </div>
                </div>

                <button className="submit-button" type="submit" disabled={!appointmentReady || isBooking}>{isBooking ? 'Termin wird erstellt...' : <>Termin bestätigen <span>→</span></>}</button>
                <p className="privacy-note">Mit der Anfrage stimmst du der Kontaktaufnahme zur Terminbestätigung zu.</p>
              </>}
            </form>
          )}
        </div>
      </section>
    </main>
  )
}

export default App
