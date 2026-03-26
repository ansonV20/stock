import cors from 'cors'
import dotenv from 'dotenv'
import express from 'express'
import { google } from 'googleapis'

dotenv.config()
dotenv.config({ path: '.env.local', override: true })

const app = express()
app.use(express.json())

const allowedOrigin = process.env.APPEND_API_ALLOW_ORIGIN || '*'
app.use(cors({ origin: allowedOrigin }))

const PORT = Number(process.env.APPEND_API_PORT || 8787)
const spreadsheetId = process.env.SPREADSHEET_ID || process.env.VITE_SPREADSHEET_ID

const TABLE_FIELD_ORDER = {
  Stocks: ['stock', 'currency', 'price', 'action', 'time', 'quantity', 'handlingFees'],
  Dividend: ['stock', 'currency', 'div', 'time'],
  'Money Move': ['name', 'currency', 'price', 'time', 'do'],
}

const RANGE_BY_TABLE = {
  Stocks: process.env.SHEET_RANGE_STOCKS || 'Stocks!A:G',
  Dividend: process.env.SHEET_RANGE_DIVIDEND || 'Dividend!A:D',
  'Money Move': process.env.SHEET_RANGE_MONEY_MOVE || 'Money Move!A:E',
}

function getGoogleAuth() {
  const keyFile = process.env.GOOGLE_APPLICATION_CREDENTIALS
  const jsonRaw = process.env.GOOGLE_SERVICE_ACCOUNT_KEY_JSON

  if (jsonRaw) {
    return new google.auth.GoogleAuth({
      credentials: JSON.parse(jsonRaw),
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    })
  }

  if (keyFile) {
    return new google.auth.GoogleAuth({
      keyFile,
      scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    })
  }

  throw new Error(
    'Missing Google credentials. Set GOOGLE_APPLICATION_CREDENTIALS or GOOGLE_SERVICE_ACCOUNT_KEY_JSON.',
  )
}

app.get('/health', (_req, res) => {
  res.json({ ok: true })
})

app.post('/api/sheets/append', async (req, res) => {
  try {
    if (!spreadsheetId) {
      return res.status(500).json({ error: 'Missing SPREADSHEET_ID (or VITE_SPREADSHEET_ID).' })
    }

    const { table, row } = req.body ?? {}
    if (!table || !TABLE_FIELD_ORDER[table]) {
      return res.status(400).json({ error: 'Invalid table. Use Stocks, Dividend, or Money Move.' })
    }

    if (!row || typeof row !== 'object') {
      return res.status(400).json({ error: 'Invalid row payload.' })
    }

    const orderedRow = TABLE_FIELD_ORDER[table].map((fieldKey) => String(row[fieldKey] ?? ''))

    const auth = getGoogleAuth()
    const sheets = google.sheets({ version: 'v4', auth })

    const response = await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: RANGE_BY_TABLE[table],
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [orderedRow],
      },
    })

    return res.json({
      ok: true,
      table,
      updatedRows: response.data.updates?.updatedRows ?? 0,
      updatedRange: response.data.updates?.updatedRange ?? null,
    })
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unknown server error'
    return res.status(500).json({ error: message })
  }
})

app.listen(PORT, () => {
  console.log(`Append API listening on http://localhost:${PORT}`)
})
