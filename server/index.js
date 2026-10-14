import express from 'express'
import cors from 'cors'
import fs from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)
const rootDir = path.resolve(__dirname, '..')
const dataDir = path.join(rootDir, 'data')
const stateFile = path.join(dataDir, 'state.json')
const port = process.env.PORT || 8787

const app = express()
app.use(cors())
app.use(express.json({ limit: '2mb' }))

async function ensureStateFile() {
  await fs.mkdir(dataDir, { recursive: true })
  try {
    await fs.access(stateFile)
  } catch {
    const initial = { bookings: [], assignments: {} }
    await fs.writeFile(stateFile, JSON.stringify(initial, null, 2), 'utf-8')
  }
}

async function readState() {
  await ensureStateFile()
  const raw = await fs.readFile(stateFile, 'utf-8')
  try {
    const parsed = JSON.parse(raw)
    return {
      bookings: Array.isArray(parsed.bookings) ? parsed.bookings : [],
      assignments:
        parsed.assignments && typeof parsed.assignments === 'object' ? parsed.assignments : {},
    }
  } catch {
    return { bookings: [], assignments: {} }
  }
}

async function writeState(nextState) {
  await ensureStateFile()
  await fs.writeFile(stateFile, JSON.stringify(nextState, null, 2), 'utf-8')
}

app.get('/api/health', (_req, res) => {
  res.json({ ok: true })
})

app.get('/api/state', async (_req, res) => {
  const state = await readState()
  res.json(state)
})

app.put('/api/bookings', async (req, res) => {
  const incoming = req.body?.bookings
  if (!Array.isArray(incoming)) {
    return res.status(400).json({ error: 'bookings must be an array' })
  }

  const current = await readState()
  const validIds = new Set(incoming.map((item) => item.id))
  const prunedAssignments = Object.fromEntries(
    Object.entries(current.assignments).filter(([bookingId]) => validIds.has(bookingId)),
  )

  const nextState = {
    bookings: incoming,
    assignments: prunedAssignments,
  }
  await writeState(nextState)
  return res.json(nextState)
})

app.put('/api/assignments', async (req, res) => {
  const incoming = req.body?.assignments
  if (!incoming || typeof incoming !== 'object' || Array.isArray(incoming)) {
    return res.status(400).json({ error: 'assignments must be an object' })
  }

  const current = await readState()
  const validIds = new Set(current.bookings.map((item) => item.id))
  const sanitizedAssignments = Object.fromEntries(
    Object.entries(incoming).filter(([bookingId]) => validIds.has(bookingId)),
  )

  const nextState = {
    bookings: current.bookings,
    assignments: sanitizedAssignments,
  }
  await writeState(nextState)
  return res.json(nextState)
})

app.listen(port, () => {
  console.log(`Mawla local API running at http://localhost:${port}`)
})
