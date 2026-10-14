import { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import './App.css'

const EVENT_DATES = ['2026-06-03', '2026-06-04', '2026-06-05', '2026-06-06', '2026-06-07']

const ROOM_TYPES = [
  { type: 'Cottage Dormitory', units: [{ name: 'Purandar', capacity: 11 }, { name: 'Sinhagad', capacity: 11 }, { name: 'Rajgad', capacity: 11 }, { name: 'Raigad', capacity: 11 }] },
  { type: 'Dormitory Hall', units: [{ name: 'Torna', capacity: 28 }] },
  { type: 'Bamboo House', units: [{ name: 'Bamboo House', capacity: 14 }] },
  { type: 'Bring your own tent', units: [{ name: 'Bring your own tent', capacity: 1, wholeUnitBooking: true }] },
  { type: '8 person tent', units: [{ name: '8 Person Tent - 1', capacity: 8 }, { name: '8 Person Tent - 2', capacity: 8 }] },
  { type: '4 person tent', units: [{ name: '4 person tent', capacity: 4, wholeUnitBooking: true }] },
  {
    type: '3 person tent',
    units: [
      { name: '3 person tent - 1', capacity: 3, wholeUnitBooking: true },
      { name: '3 person tent - 2', capacity: 3, wholeUnitBooking: true },
      { name: '3 person tent - 3', capacity: 3, wholeUnitBooking: true },
    ],
  },
]

const ROOM_UNITS = ROOM_TYPES.flatMap((roomType) =>
  roomType.units.map((unit, idx) => ({
    id: `${roomType.type}-${idx + 1}`,
    label: unit.name,
    type: roomType.type,
    capacity: unit.capacity,
    wholeUnitBooking: Boolean(unit.wholeUnitBooking),
  })),
)

function normalizeHeader(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '')
}

function normalizePhone(value) {
  return String(value || '').replace(/\D/g, '').replace(/^0+/, '')
}

function parseDate(value) {
  if (!value) return null
  if (typeof value === 'number' && Number.isFinite(value)) {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30))
    excelEpoch.setUTCDate(excelEpoch.getUTCDate() + Math.floor(value))
    return new Date(excelEpoch.getUTCFullYear(), excelEpoch.getUTCMonth(), excelEpoch.getUTCDate())
  }
  const raw = String(value).trim()
  const iso = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/)
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]))
  const slash = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/)
  if (slash) return new Date(Number(slash[3]), Number(slash[1]) - 1, Number(slash[2]))
  const shortSlash = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})$/)
  if (shortSlash) return new Date(2000 + Number(shortSlash[3]), Number(shortSlash[1]) - 1, Number(shortSlash[2]))
  const dayMonthOnly = raw.match(/^(\d{1,2})[-\s]([A-Za-z]+)$/)
  if (dayMonthOnly) {
    const day = Number(dayMonthOnly[1])
    const monthName = dayMonthOnly[2].toLowerCase()
    const monthMap = {
      jan: 0,
      january: 0,
      feb: 1,
      february: 1,
      mar: 2,
      march: 2,
      apr: 3,
      april: 3,
      may: 4,
      jun: 5,
      june: 5,
      jul: 6,
      july: 6,
      aug: 7,
      august: 7,
      sep: 8,
      sept: 8,
      september: 8,
      oct: 9,
      october: 9,
      nov: 10,
      november: 10,
      dec: 11,
      december: 11,
    }
    const monthIndex = monthMap[monthName]
    if (monthIndex !== undefined && day >= 1 && day <= 31) {
      return new Date(2026, monthIndex, day)
    }
  }
  const parsed = new Date(raw)
  if (Number.isNaN(parsed.getTime())) return null
  return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate())
}

function validateBookingDate(value) {
  if (typeof value === 'number' && Number.isFinite(value)) {
    const date = parseDate(value)
    return date ? { ok: true, date, reason: '' } : { ok: false, reason: 'Invalid spreadsheet date.' }
  }
  const raw = String(value || '').trim()
  if (!raw) {
    return { ok: false, reason: 'Date is blank. Use MM/DD/YYYY or M/D/YYYY.' }
  }
  const dayMonthOnly = raw.match(/^(\d{1,2})[-\s]([A-Za-z]+)$/)
  if (dayMonthOnly) {
    const parsed = parseDate(raw)
    return parsed ? { ok: true, date: parsed, reason: '' } : { ok: false, reason: 'Invalid date.' }
  }
  if (raw.includes('-')) {
    return { ok: false, reason: 'Wrong separator. Use MM/DD/YYYY or M/D/YYYY with "/" separators.' }
  }

  const parsedMatch = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/)
  if (!parsedMatch) {
    const shortYearMatch = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})$/)
    if (!shortYearMatch) {
      return { ok: false, reason: 'Wrong format. Expected MM/DD/YYYY or M/D/YYYY (example: 06/03/2026 or 6/3/2026).' }
    }
    const parsed = parseDate(raw)
    return parsed ? { ok: true, date: parsed, reason: '' } : { ok: false, reason: 'Invalid date.' }
  }

  const month = Number(parsedMatch[1])
  const day = Number(parsedMatch[2])
  const year = Number(parsedMatch[3])

  if (month < 1 || month > 12) {
    return { ok: false, reason: `Invalid month "${parsedMatch[1]}". Month must be 1 to 12.` }
  }
  if (year < 1900 || year > 2100) {
    return { ok: false, reason: `Invalid year "${parsedMatch[3]}".` }
  }

  const maxDay = new Date(year, month, 0).getDate()
  if (day < 1 || day > maxDay) {
    return { ok: false, reason: `Invalid day "${parsedMatch[2]}" for month ${parsedMatch[1]}/${parsedMatch[3]}.` }
  }

  return { ok: true, date: new Date(year, month - 1, day), reason: '' }
}

function toYmd(date) {
  return `${date.getFullYear()}-${`${date.getMonth() + 1}`.padStart(2, '0')}-${`${date.getDate()}`.padStart(2, '0')}`
}

function formatDate(value) {
  const d = parseDate(value)
  if (!d) return 'Invalid date'
  return d.toLocaleDateString('en-IN', { weekday: 'short', day: '2-digit', month: 'short', year: 'numeric' })
}

function formatDateNoDay(value) {
  const d = parseDate(value)
  if (!d) return 'Invalid date'
  return d.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' })
}

function getStayDates(checkIn, nights) {
  const parsed = validateBookingDate(checkIn)
  if (!parsed.ok) return []
  const d = parsed.date
  return Array.from({ length: Math.max(1, nights) }, (_, idx) => {
    const c = new Date(d)
    c.setDate(c.getDate() + idx)
    return toYmd(c)
  })
}

function getEventDateFlag(checkInDate) {
  const parsed = validateBookingDate(checkInDate)
  if (!parsed.ok) return `Invalid date: ${parsed.reason}`
  const d = parsed.date
  const start = parseDate(EVENT_DATES[0])
  const end = parseDate(EVENT_DATES[EVENT_DATES.length - 1])
  if (d < start) return 'Earlier than event dates'
  if (d > end) return 'Later than event dates'
  return ''
}

function csvToObjects(text) {
  const workbook = XLSX.read(text, { type: 'string', raw: true })
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]]
  return XLSX.utils.sheet_to_json(firstSheet, { defval: '', raw: true })
}

function mapAccommodation(rawValue) {
  const value = String(rawValue || '').toLowerCase().replace(/[–-]/g, ' ').replace(/\s+/g, ' ').trim()
  if (value.includes('dorm cottage')) return 'Cottage Dormitory'
  if (value.includes('dorm hall') || value.includes('dormitory hall')) return 'Dormitory Hall'
  if (value.includes('bamboo house')) return 'Bamboo House'
  if (value.includes('bring your own tent') || value.includes('bring your won tent')) return 'Bring your own tent'
  if (value.includes('large tent') || value.includes('8') || value.includes('big tent')) return '8 person tent'
  if (value.includes('4 person')) return '4 person tent'
  if (value.includes('3 person')) return '3 person tent'
  return rawValue || '-'
}

function normalizeGuestName(value) {
  const compacted = String(value || '')
    .replace(/\s+/g, ' ')
    .replace(/\bDr\.\s*/gi, 'Dr. ')
    .trim()

  return compacted.replace(/[A-Za-z]+(?:\.[A-Za-z]+)?/g, (part) => {
    if (part.toLowerCase() === 'dr') return 'Dr'
    return part.charAt(0).toUpperCase() + part.slice(1).toLowerCase()
  })
}

function mapAccommodationValue(rawValue) {
  const value = String(rawValue || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').replace(/\s+/g, ' ').trim()
  if (value.includes('dorm cottage')) return 'Cottage Dormitory'
  if (value.includes('dorm hall') || value.includes('dormitory hall')) return 'Dormitory Hall'
  if (value.includes('bamboo house')) return 'Bamboo House'
  if (value.includes('bring your own tent') || value.includes('bring your won tent')) return 'Bring your own tent'
  if (value.includes('large tent') || value.includes('8 person') || value.includes('big tent')) return '8 person tent'
  if (value.includes('4 person')) return '4 person tent'
  if (value.includes('3 person')) return '3 person tent'
  return mapAccommodation(rawValue)
}

function getNormalizedValue(row, ...keys) {
  for (const key of keys) {
    if (row[key]) return row[key]
    const prefixMatch = Object.entries(row).find(([rowKey]) => rowKey.startsWith(key))
    if (prefixMatch?.[1]) return prefixMatch[1]
  }
  return ''
}

function normalizeBookings(rows) {
  return rows
    .map((row, idx) => {
      const n = Object.entries(row).reduce((acc, [k, v]) => {
        acc[normalizeHeader(k)] = v
        return acc
      }, {})
      const nightsValue = getNormalizedValue(n, 'noofnightsofstay', 'nights')
      const nights = Number.parseInt(nightsValue || '1', 10)
      return {
        id: `booking-${idx + 1}`,
        fullName: normalizeGuestName(n.fullname || n.name || n.guestname || ''),
        phoneNumber: String(n.phonenumber || n.phone || '').trim(),
        gender: String(n.gender || '-').trim(),
        preferredAccommodation: mapAccommodationValue(
          getNormalizedValue(n, 'accommodationpernightcharges', 'accommodationtype', 'roomtype'),
        ),
        checkInDate: getNormalizedValue(n, 'dateofcheckinmmddyyyy', 'dateofcheckin', 'checkindate'),
        nights: Number.isNaN(nights) || nights < 1 ? 1 : nights,
      }
    })
    .filter((b) => b.fullName)
}

function renumberBookings(bookings) {
  return bookings.map((booking, idx) => ({ ...booking, id: `booking-${idx + 1}` }))
}

function bookingFingerprint(booking) {
  const parsed = validateBookingDate(booking.checkInDate)
  const date = parsed.ok ? parsed.date : null
  return [
    booking.fullName.toLowerCase(),
    normalizePhone(booking.phoneNumber),
    date ? toYmd(date) : 'invalid',
    booking.nights,
    String(booking.preferredAccommodation || '').toLowerCase(),
  ].join('|')
}

function bookingIdentityFingerprint(booking) {
  return [
    booking.fullName.toLowerCase(),
    normalizePhone(booking.phoneNumber),
    String(booking.preferredAccommodation || '').toLowerCase(),
  ].join('|')
}

function mergeBookingsWithoutDuplicates(existing, incoming) {
  const merged = [...existing]
  const idxByKey = new Map(merged.map((b, idx) => [bookingFingerprint(b), idx]))
  const idxByIdentity = new Map()
  merged.forEach((booking, idx) => {
    const identity = bookingIdentityFingerprint(booking)
    if (!idxByIdentity.has(identity)) idxByIdentity.set(identity, idx)
  })
  let added = 0
  let updated = 0
  incoming.forEach((item) => {
    const key = bookingFingerprint(item)
    const incomingDate = validateBookingDate(item.checkInDate)
    const identityKey = bookingIdentityFingerprint(item)
    const matchIdx = idxByKey.has(key)
      ? idxByKey.get(key)
      : !incomingDate.ok && idxByIdentity.has(identityKey)
        ? idxByIdentity.get(identityKey)
        : null

    if (matchIdx !== null) {
      const existingItem = merged[matchIdx]
      const nextItem = { ...existingItem, ...item, id: existingItem.id }
      if (!incomingDate.ok && validateBookingDate(existingItem.checkInDate).ok) {
        nextItem.checkInDate = existingItem.checkInDate
        nextItem.nights = existingItem.nights
      }
      merged[matchIdx] = nextItem
      updated += 1
    } else {
      merged.push({ ...item })
      idxByKey.set(key, merged.length - 1)
      idxByIdentity.set(identityKey, merged.length - 1)
      added += 1
    }
  })
  return { merged: renumberBookings(merged), added, updated }
}

function normalizeAssignments(rawAssignments) {
  const out = {}
  Object.entries(rawAssignments || {}).forEach(([bookingId, value]) => {
    if (!value || typeof value !== 'object' || !value.roomUnitId) return
    out[bookingId] = {
      roomUnitId: String(value.roomUnitId),
      seatNumber: String(value.seatNumber || ''),
      comment: String(value.comment || ''),
      assignedAt: String(value.assignedAt || ''),
    }
  })
  return out
}

function buildEmptyOccupancy() {
  const map = {}
  ROOM_UNITS.forEach((unit) => {
    map[unit.id] = {}
    EVENT_DATES.forEach((date) => {
      map[unit.id][date] = []
    })
  })
  return map
}

function buildOccupancy(bookings, assignments) {
  const map = buildEmptyOccupancy()
  bookings.forEach((booking) => {
    const assignment = assignments[booking.id]
    if (!assignment?.roomUnitId || !map[assignment.roomUnitId]) return
    getStayDates(booking.checkInDate, booking.nights)
      .filter((date) => EVENT_DATES.includes(date))
      .forEach((date) => {
        map[assignment.roomUnitId][date].push(booking.id)
      })
  })
  return map
}

function autoAssignBookings(bookings) {
  const sorted = [...bookings].sort((a, b) => {
    const da = parseDate(a.checkInDate)?.getTime() ?? 0
    const db = parseDate(b.checkInDate)?.getTime() ?? 0
    if (da !== db) return da - db
    return a.fullName.localeCompare(b.fullName)
  })
  const occupancy = buildEmptyOccupancy()
  const assignments = {}

  sorted.forEach((booking) => {
    const unitOptions = ROOM_UNITS.filter((u) => u.type === booking.preferredAccommodation)
    if (!unitOptions.length) return
    const stayDates = getStayDates(booking.checkInDate, booking.nights).filter((d) => EVENT_DATES.includes(d))
    if (!stayDates.length) return

    for (const unit of unitOptions) {
      // Purandar is reserved for female guests only.
      if (unit.label === 'Purandar' && !String(booking.gender || '').toLowerCase().startsWith('f')) {
        continue
      }

      const limit = unit.wholeUnitBooking ? 1 : unit.capacity
      const canFit = stayDates.every((date) => occupancy[unit.id][date].length < limit)
      if (!canFit) continue

      let seatNumber = 'FULL'
      if (!unit.wholeUnitBooking) {
        seatNumber = ''
        for (let seat = 1; seat <= unit.capacity; seat += 1) {
          const seatStr = String(seat)
          const blocked = stayDates.some((date) =>
            occupancy[unit.id][date].some((id) => assignments[id]?.seatNumber === seatStr),
          )
          if (!blocked) {
            seatNumber = seatStr
            break
          }
        }
      }
      if (!seatNumber) continue

      assignments[booking.id] = {
        roomUnitId: unit.id,
        seatNumber,
        comment: 'Auto-assigned',
        assignedAt: new Date().toISOString(),
      }
      stayDates.forEach((date) => occupancy[unit.id][date].push(booking.id))
      break
    }
  })

  return assignments
}

function buildAssignmentSummary(bookings, assignments) {
  const assigned = bookings
    .filter((booking) => assignments[booking.id])
    .map((booking) => booking.fullName)
  const pending = bookings
    .filter((booking) => !assignments[booking.id])
    .map((booking) => booking.fullName)
  return { assigned, pending }
}

function App() {
  const [bookings, setBookings] = useState([])
  const [assignments, setAssignments] = useState({})
  const [activeTab, setActiveTab] = useState('assign')
  const [selectedBookingId, setSelectedBookingId] = useState('')
  const [selectedDate, setSelectedDate] = useState(EVENT_DATES[0])
  const [assignmentComment, setAssignmentComment] = useState('')
  const [chartDateFilter, setChartDateFilter] = useState(EVENT_DATES[0])
  const [chartRoomFilter, setChartRoomFilter] = useState(ROOM_UNITS[0].id)
  const [chartDate, setChartDate] = useState(EVENT_DATES[0])
  const [chartRoomUnit, setChartRoomUnit] = useState(ROOM_UNITS[0].id)
  const [showAddGuestModal, setShowAddGuestModal] = useState(false)
  const [seatGuestPopup, setSeatGuestPopup] = useState(null)
  const [assignmentStatusPopup, setAssignmentStatusPopup] = useState(null)
  const [newGuestForm, setNewGuestForm] = useState({
    fullName: '',
    phoneNumber: '',
    checkInDate: '06/03/2026',
    nights: '1',
    accommodationType: ROOM_TYPES[0].type,
  })
  const [feedback, setFeedback] = useState('Loading data...')
  const [error, setError] = useState('')

  const occupancy = useMemo(() => buildOccupancy(bookings, assignments), [bookings, assignments])
  const pendingGuests = useMemo(() => bookings.filter((b) => !assignments[b.id]), [bookings, assignments])
  const selectedBooking = useMemo(
    () => pendingGuests.find((b) => b.id === selectedBookingId) || pendingGuests[0] || null,
    [pendingGuests, selectedBookingId],
  )
  const chartUnit = ROOM_UNITS.find((u) => u.id === chartRoomUnit) || ROOM_UNITS[0]
  const chartGuests = (occupancy[chartUnit.id]?.[chartDate] || [])
    .map((id) => {
      const booking = bookings.find((b) => b.id === id)
      if (!booking) return null
      return { booking, assignment: assignments[id] }
    })
    .filter(Boolean)

  useEffect(() => {
    void (async () => {
      try {
        const response = await fetch('/api/state')
        if (!response.ok) throw new Error(`HTTP ${response.status}`)
        const state = await response.json()
        setBookings(Array.isArray(state.bookings) ? state.bookings : [])
        setAssignments(normalizeAssignments(state.assignments || {}))
        setFeedback('Data loaded.')
        setError('')
      } catch (loadError) {
        setError(`Could not load API state: ${loadError.message}`)
      }
    })()
  }, [])

  async function saveBookingsAndAssignments(nextBookings, nextAssignments) {
    const bookingsResponse = await fetch('/api/bookings', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ bookings: nextBookings }),
    })
    if (!bookingsResponse.ok) throw new Error(`Bookings save failed: HTTP ${bookingsResponse.status}`)
    const assignResponse = await fetch('/api/assignments', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ assignments: nextAssignments }),
    })
    if (!assignResponse.ok) throw new Error(`Assignments save failed: HTTP ${assignResponse.status}`)
    setBookings(nextBookings)
    setAssignments(normalizeAssignments(nextAssignments))
  }

  async function importCsvFile(file) {
    if (!file) return
    setError('')
    try {
      const text = await file.text()
      const incoming = normalizeBookings(csvToObjects(text))
      if (!incoming.length) throw new Error('No valid rows found in CSV.')
      const { merged, added, updated } = mergeBookingsWithoutDuplicates(bookings, incoming)
      const autoAssignments = autoAssignBookings(merged)
      await saveBookingsAndAssignments(merged, autoAssignments)
      setAssignmentStatusPopup(buildAssignmentSummary(merged, autoAssignments))
      setFeedback(`CSV merged: ${added} new, ${updated} updated, no duplicates. Auto-assigned where possible.`)
    } catch (loadError) {
      setError(`CSV import failed: ${loadError.message}`)
    }
  }

  async function removePendingGuest(bookingId) {
    try {
      const nextBookings = bookings.filter((b) => b.id !== bookingId)
      const nextAssignments = { ...assignments }
      delete nextAssignments[bookingId]
      await saveBookingsAndAssignments(nextBookings, nextAssignments)
      setFeedback('Pending guest removed.')
    } catch (removeError) {
      setError(removeError.message)
    }
  }

  async function saveNewGuest() {
    const fullName = newGuestForm.fullName.trim()
    const nights = Number.parseInt(newGuestForm.nights, 10)
    if (!fullName) return setError('Guest name is required.')
    if (!parseDate(newGuestForm.checkInDate)) return setError('Valid check-in date required.')
    const dateValidation = validateBookingDate(newGuestForm.checkInDate)
    if (!dateValidation.ok) return setError(`Invalid date: ${dateValidation.reason}`)
    if (Number.isNaN(nights) || nights < 1) return setError('Nights must be 1 or more.')
    try {
      const nextBookings = renumberBookings([
        ...bookings,
        {
          id: 'booking-x',
          fullName,
          phoneNumber: newGuestForm.phoneNumber.trim(),
          gender: '-',
          preferredAccommodation: newGuestForm.accommodationType,
          checkInDate: newGuestForm.checkInDate,
          nights,
        },
      ])
      const autoAssignments = autoAssignBookings(nextBookings)
      await saveBookingsAndAssignments(nextBookings, autoAssignments)
      setAssignmentStatusPopup(buildAssignmentSummary(nextBookings, autoAssignments))
      setShowAddGuestModal(false)
      setFeedback(`${fullName} added and auto-assigned if possible.`)
      setError('')
    } catch (saveError) {
      setError(saveError.message)
    }
  }

  function getPendingReason(booking) {
    const dateValidation = validateBookingDate(booking.checkInDate)
    if (!dateValidation.ok) return `Flagged: ${dateValidation.reason}`
    const stay = getStayDates(booking.checkInDate, booking.nights).filter((d) => EVENT_DATES.includes(d))
    if (!stay.length) return 'Flagged: check-in date outside event dates'
    const units = ROOM_UNITS.filter((u) => u.type === booking.preferredAccommodation)
    if (!units.length) return 'Flagged: accommodation type not available'
    return 'Flagged: Accomodation FULL'
  }

  function handleSeatClick(date, unit, seat) {
    const bookingIds = occupancy[unit.id][date] || []
    const guests = bookingIds
      .map((id) => ({ booking: bookings.find((b) => b.id === id), assignment: assignments[id] }))
      .filter((x) => x.booking)
      .filter((x) => unit.wholeUnitBooking || x.assignment?.seatNumber === seat)
    setSeatGuestPopup({ date, unitLabel: unit.label, seat, guests })
  }

  function applyChartFilters() {
    setChartDate(chartDateFilter)
    setChartRoomUnit(chartRoomFilter)
  }

  function exportAssignmentsExcel() {
    const assignedRows = bookings
      .filter((b) => assignments[b.id])
      .map((booking) => {
        const assignment = assignments[booking.id]
        const room = ROOM_UNITS.find((x) => x.id === assignment.roomUnitId)
        return {
          Name: booking.fullName,
          'Phone Number': booking.phoneNumber || '',
          Gender: booking.gender,
          'Check-In Date': formatDateNoDay(booking.checkInDate),
          Nights: booking.nights,
          'Accommodation Type': booking.preferredAccommodation,
          'Assigned Room': room?.label || assignment.roomUnitId,
          Seat: assignment.seatNumber || '',
          Comment: assignment.comment || '',
        }
      })

    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(assignedRows.length ? assignedRows : [{ Info: 'No assignments yet' }]),
      'All Assignments',
    )
    EVENT_DATES.forEach((date) => {
      const rowsForDate = bookings
        .filter((b) => assignments[b.id] && getStayDates(b.checkInDate, b.nights).includes(date))
        .map((booking) => {
          const assignment = assignments[booking.id]
          const room = ROOM_UNITS.find((x) => x.id === assignment.roomUnitId)
          return {
            Name: booking.fullName,
            'Phone Number': booking.phoneNumber || '',
            Gender: booking.gender,
            'Check-In Date': formatDateNoDay(booking.checkInDate),
            Nights: booking.nights,
            'Accommodation Type': booking.preferredAccommodation,
            'Assigned Room': room?.label || assignment.roomUnitId,
            Seat: assignment.seatNumber || '',
            Comment: assignment.comment || '',
          }
        })
      XLSX.utils.book_append_sheet(
        workbook,
        XLSX.utils.json_to_sheet(rowsForDate.length ? rowsForDate : [{ Info: 'No assignments for this date' }]),
        date,
      )
    })
    XLSX.writeFile(workbook, `mawla-assignments-${Date.now()}.xlsx`)
    setFeedback('Assignments Excel exported.')
  }

  return (
    <main className="app-shell">
      <header className="topbar">
        <h1>Mawla Booking Console</h1>
      </header>

      <section className="panel data-loader">
        <div className="loader-head">
          <h2>Data Import</h2>
        </div>
        <div className="loader-row">
          <input type="file" accept=".csv,text/csv" onChange={(event) => importCsvFile(event.target.files?.[0])} />
          <button type="button" onClick={exportAssignmentsExcel}>Export Assignments Excel</button>
        </div>
        {error ? <p className="error">{error}</p> : null}
        {feedback ? <p className="ok">{feedback}</p> : null}
      </section>

      <section className="tabs panel">
        <button type="button" className={activeTab === 'assign' ? 'active' : ''} onClick={() => setActiveTab('assign')}>Assign Guest</button>
        <button type="button" className={activeTab === 'chart' ? 'active' : ''} onClick={() => setActiveTab('chart')}>Room-Date Charts</button>
      </section>

      {activeTab === 'assign' ? (
        <section className="assign-layout">
          <section className="panel">
            <div className="row-head">
              <h2>Guest Picker</h2>
              <button type="button" onClick={() => setShowAddGuestModal(true)}>Add New Guest</button>
            </div>
            <select value={selectedBooking?.id || ''} onChange={(event) => setSelectedBookingId(event.target.value)}>
              <option value="">Select Full Name</option>
              {pendingGuests.map((booking) => (
                <option key={booking.id} value={booking.id}>{booking.fullName}</option>
              ))}
            </select>
            {selectedBooking ? (
              <div className="guest-card">
                <h3>{selectedBooking.fullName}</h3>
                <dl>
                  <div><dt>Phone</dt><dd>{selectedBooking.phoneNumber || '-'}</dd></div>
                  <div><dt>Accommodation</dt><dd>{selectedBooking.preferredAccommodation}</dd></div>
                  <div><dt>Check In</dt><dd>{formatDate(selectedBooking.checkInDate)}</dd></div>
                  <div><dt>No of Nights</dt><dd>{selectedBooking.nights}</dd></div>
                  {getEventDateFlag(selectedBooking.checkInDate) ? (
                    <div><dt>Date Status</dt><dd className="date-flag">{getEventDateFlag(selectedBooking.checkInDate)}</dd></div>
                  ) : null}
                </dl>
              </div>
            ) : null}

            <label className="comment-box">
              Comment
              <textarea rows={3} value={assignmentComment} onChange={(event) => setAssignmentComment(event.target.value)} />
            </label>

            <h3>Pending Guests</h3>
            <ul className="compact-list pending-list">
              {pendingGuests.length ? (
                pendingGuests.map((booking) => (
                  <li key={`pending-${booking.id}`}>
                    <div>
                      <strong>{booking.fullName}</strong>
                      <p>{getPendingReason(booking)}</p>
                    </div>
                    <button type="button" onClick={() => removePendingGuest(booking.id)}>Remove</button>
                  </li>
                ))
              ) : (
                <li>No pending guests.</li>
              )}
            </ul>
          </section>

          <section className="panel">
            <h2>Live Availability</h2>
            <div className="date-tabs">
              {EVENT_DATES.map((date) => (
                <button key={date} type="button" className={selectedDate === date ? 'active' : ''} onClick={() => setSelectedDate(date)}>
                  {formatDate(date)}
                </button>
              ))}
            </div>
            <div className="room-grid">
              {ROOM_UNITS.map((unit) => {
                const bookingIds = occupancy[unit.id][selectedDate]
                const occupiedCount = bookingIds.length
                const occupiedSeats = bookingIds.map((id) => assignments[id]?.seatNumber).filter(Boolean)
                return (
                  <article key={unit.id} className="room-card">
                    <header>
                      <h3>{unit.label}</h3>
                      <p>{occupiedCount}/{unit.wholeUnitBooking ? 1 : unit.capacity} occupied</p>
                    </header>
                    <div className="seats">
                      {Array.from({ length: unit.wholeUnitBooking ? 1 : unit.capacity }, (_, idx) => {
                        const seat = String(idx + 1)
                        const occupied = unit.wholeUnitBooking ? occupiedCount > 0 : occupiedSeats.includes(seat)
                        return (
                          <button
                            key={`${unit.id}-${seat}`}
                            type="button"
                            className={`seat-btn ${occupied ? 'occupied' : 'free'}`}
                            onClick={() => handleSeatClick(selectedDate, unit, seat)}
                            title={unit.wholeUnitBooking ? unit.label : `Seat ${seat}`}
                          />
                        )
                      })}
                    </div>
                  </article>
                )
              })}
            </div>
          </section>
        </section>
      ) : null}

      {activeTab === 'chart' ? (
        <section className="panel chart-panel">
          <h2>Room-Date Charts</h2>
          <div className="chart-filters-top">
            <label>
              Date
              <select value={chartDateFilter} onChange={(event) => setChartDateFilter(event.target.value)}>
                {EVENT_DATES.map((date) => (
                  <option key={date} value={date}>{formatDate(date)}</option>
                ))}
              </select>
            </label>
            <label>
              Room
              <select value={chartRoomFilter} onChange={(event) => setChartRoomFilter(event.target.value)}>
                {ROOM_UNITS.map((unit) => (
                  <option key={unit.id} value={unit.id}>{unit.label}</option>
                ))}
              </select>
            </label>
            <button type="button" className="filter-btn" onClick={applyChartFilters}>Apply Filter</button>
          </div>

          <section className="detail-card">
            <h3>{chartUnit.label} on {formatDate(chartDate)}</h3>
            <ul className="compact-list">
              {chartGuests.length ? (
                chartGuests.map(({ booking, assignment }) => (
                  <li key={`${chartDate}-${booking.id}`}>
                    <span>{booking.fullName}</span>
                    <span>{assignment.seatNumber || '-'}</span>
                    <span>{assignment.comment || '-'}</span>
                  </li>
                ))
              ) : (
                <li>No guests assigned for this room/date selection.</li>
              )}
            </ul>
          </section>
        </section>
      ) : null}

      {showAddGuestModal ? (
        <section className="modal-backdrop" role="dialog" aria-modal="true">
          <div className="modal-card">
            <h3>Add New Guest</h3>
            <label>
              Full Name
              <input
                type="text"
                value={newGuestForm.fullName}
                onChange={(event) => setNewGuestForm((prev) => ({ ...prev, fullName: event.target.value }))}
              />
            </label>
            <label>
              Phone Number
              <input
                type="text"
                value={newGuestForm.phoneNumber}
                onChange={(event) => setNewGuestForm((prev) => ({ ...prev, phoneNumber: event.target.value }))}
              />
            </label>
            <label>
              Check In Date
              <input
                type="text"
                placeholder="MM/DD/YYYY"
                value={newGuestForm.checkInDate}
                onChange={(event) => setNewGuestForm((prev) => ({ ...prev, checkInDate: event.target.value }))}
              />
            </label>
            <label>
              No of Nights
              <input
                type="number"
                min="1"
                value={newGuestForm.nights}
                onChange={(event) => setNewGuestForm((prev) => ({ ...prev, nights: event.target.value }))}
              />
            </label>
            <label>
              Accommodation Type
              <select
                value={newGuestForm.accommodationType}
                onChange={(event) => setNewGuestForm((prev) => ({ ...prev, accommodationType: event.target.value }))}
              >
                {ROOM_TYPES.map((room) => (
                  <option key={room.type} value={room.type}>{room.type}</option>
                ))}
              </select>
            </label>
            <div className="modal-actions">
              <button type="button" onClick={() => setShowAddGuestModal(false)}>Cancel</button>
              <button type="button" className="assign-btn" onClick={saveNewGuest}>Save Guest</button>
            </div>
          </div>
        </section>
      ) : null}

      {seatGuestPopup ? (
        <section className="modal-backdrop" role="dialog" aria-modal="true">
          <div className="modal-card">
            <h3>{seatGuestPopup.unitLabel} | {formatDate(seatGuestPopup.date)} | Seat {seatGuestPopup.seat}</h3>
            {seatGuestPopup.guests.length ? (
              <ul className="compact-list">
                {seatGuestPopup.guests.map(({ booking, assignment }) => (
                  <li key={`seat-popup-${booking.id}`}>
                    <span>{booking.fullName}</span>
                    <span>{booking.phoneNumber || '-'}</span>
                    <span>{booking.preferredAccommodation}</span>
                    <span>{formatDate(booking.checkInDate)} | {booking.nights} night(s)</span>
                    <span>{assignment?.comment || '-'}</span>
                  </li>
                ))}
              </ul>
            ) : (
              <p>No guest assigned.</p>
            )}
            <button type="button" onClick={() => setSeatGuestPopup(null)}>Close</button>
          </div>
        </section>
      ) : null}

      {assignmentStatusPopup ? (
        <section className="modal-backdrop" role="dialog" aria-modal="true">
          <div className="modal-card">
            <button type="button" className="modal-close-x" onClick={() => setAssignmentStatusPopup(null)}>
              X
            </button>
            <h3>Assignment Summary</h3>
            <p>Assigned: {assignmentStatusPopup.assigned.length} | Pending: {assignmentStatusPopup.pending.length}</p>
            <div className="grid-two">
              <div>
                <h4>Assigned</h4>
                <ul className="compact-list">
                  {assignmentStatusPopup.assigned.length ? (
                    assignmentStatusPopup.assigned.map((name) => (
                      <li key={`assigned-${name}`}>
                        <span>{name}</span>
                      </li>
                    ))
                  ) : (
                    <li>No guests assigned.</li>
                  )}
                </ul>
              </div>
              <div>
                <h4>Pending</h4>
                <ul className="compact-list">
                  {assignmentStatusPopup.pending.length ? (
                    assignmentStatusPopup.pending.map((name) => (
                      <li key={`pending-popup-${name}`}>
                        <span>{name}</span>
                      </li>
                    ))
                  ) : (
                    <li>No pending guests.</li>
                  )}
                </ul>
              </div>
            </div>
            <button type="button" onClick={() => setAssignmentStatusPopup(null)}>Close</button>
          </div>
        </section>
      ) : null}
    </main>
  )
}

export default App
