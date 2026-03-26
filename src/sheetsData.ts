export type SheetData = {
  headers: string[]
  rows: string[][]
}

export type AllSheetsData = {
  stocks: SheetData
  dividend: SheetData
  moneyMove: SheetData
}

type Lot = {
  quantity: number
  unitCost: number
}

type GoogleVizResponse = {
  table?: {
    cols?: Array<{ label?: string; id?: string }>
    rows?: Array<{ c?: Array<{ v?: string | number | boolean | null }> }>
  }
}

const STOCKS_COLUMN_COUNT = 12
const OTHER_TABLE_COLUMN_COUNT = 10
const PROFIT_LOSS_HEADER = 'Profit/Loss'
const PROFIT_PERCENT_HEADER = '%'

export function parseNumber(value: string): number {
  const trimmed = value.trim()
  if (!trimmed) {
    return Number.NaN
  }

  const isNegativeByParentheses = trimmed.startsWith('(') && trimmed.endsWith(')')
  const normalized = trimmed
    .replace(/[,$%\s]/g, '')
    .replace(/[()]/g, '')
    .replace(/,/g, '')

  const parsed = Number(normalized)
  if (!Number.isFinite(parsed)) {
    return Number.NaN
  }

  return isNegativeByParentheses ? -parsed : parsed
}

function normalizeHeader(value: string): string {
  return value.toLowerCase().replace(/[^a-z0-9]/g, '')
}

function findColumnIndex(headers: string[], keys: string[]): number {
  const normalizedKeys = keys.map((key) => normalizeHeader(key))

  return headers.findIndex((header) => {
    const normalizedHeader = normalizeHeader(header)
    return normalizedKeys.some((key) => normalizedHeader.includes(key))
  })
}

export function resolveAction(actionValue: string, quantity: number): 'buy' | 'sell' | null {
  const normalized = actionValue.trim().toLowerCase()

  if (normalized.includes('sell')) {
    return 'sell'
  }

  if (normalized.includes('buy')) {
    return 'buy'
  }

  if (quantity < 0) {
    return 'sell'
  }

  if (quantity > 0) {
    return 'buy'
  }

  return null
}

function formatSignedNumber(value: number): string {
  const sign = value > 0 ? '+' : ''
  return `${sign}${value.toFixed(2)}`
}

function parseGoogleVizResponse(payload: string, visibleColumnCount: number): SheetData {
  const start = payload.indexOf('{')
  const end = payload.lastIndexOf('}')

  if (start === -1 || end === -1 || end <= start) {
    throw new Error('Unable to parse Google Sheets response.')
  }

  const json = JSON.parse(payload.slice(start, end + 1)) as GoogleVizResponse
  const columns = json.table?.cols ?? []
  const rows = json.table?.rows ?? []

  const headers = columns.slice(0, visibleColumnCount).map((col, index) => {
    const fallback = `Column ${index + 1}`
    return (col.label?.trim() || col.id?.trim() || fallback).toString()
  })

  const parsedRows = rows.map((row) =>
    (row.c ?? []).slice(0, visibleColumnCount).map((cell) => {
      if (cell?.v === null || cell?.v === undefined) {
        return ''
      }
      return String(cell.v)
    }),
  )

  return { headers, rows: parsedRows }
}

async function fetchSheetByGid(
  spreadsheetId: string,
  gid: string,
  visibleColumnCount: number,
): Promise<SheetData> {
  const response = await fetch(
    `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?tqx=out:json&gid=${gid}`,
  )

  if (!response.ok) {
    throw new Error(`Google Sheets request failed with status ${response.status}`)
  }

  const payload = await response.text()
  return parseGoogleVizResponse(payload, visibleColumnCount)
}

export async function fetchAllSheetsData(
  spreadsheetId: string,
  stocksGid: string,
  dividendGid: string,
  moneyMoveGid: string,
): Promise<AllSheetsData> {
  const [stocks, dividend, moneyMove] = await Promise.all([
    fetchSheetByGid(spreadsheetId, stocksGid, STOCKS_COLUMN_COUNT),
    fetchSheetByGid(spreadsheetId, dividendGid, OTHER_TABLE_COLUMN_COUNT),
    fetchSheetByGid(spreadsheetId, moneyMoveGid, OTHER_TABLE_COLUMN_COUNT),
  ])

  console.log('Fetched Stocks Sheet:', stocks)
  console.log('Fetched Dividend Sheet:', dividend)
  console.log('Fetched Money Move Sheet:', moneyMove)

  return { stocks, dividend, moneyMove }
}

export function enrichStocksWithProfitLoss(data: SheetData): SheetData {
  const actionIndex = findColumnIndex(data.headers, [
    'action',
    'type',
    'side',
    'buysell',
    'operation',
  ])
  const symbolIndex = findColumnIndex(data.headers, [
    'symbol',
    'ticker',
    'stock',
    'code',
    'company',
  ])
  const quantityIndex = findColumnIndex(data.headers, [
    'qty',
    'quantity',
    'share',
    'units',
    'vol',
    'volume',
  ])
  const priceIndex = findColumnIndex(data.headers, [
    'price',
    'avgprice',
    'unitprice',
    'cost',
    'rate',
  ])
  const totalIndex = findColumnIndex(data.headers, [
    'total',
    'amount',
    'value',
    'proceeds',
    'market value',
    'marketvalue',
  ])

  const nextHeaders = [...data.headers, PROFIT_LOSS_HEADER, PROFIT_PERCENT_HEADER]

  if (quantityIndex === -1 || symbolIndex === -1) {
    return {
      headers: nextHeaders,
      rows: data.rows.map((row) => [...row, '', '']),
    }
  }

  const inventory = new Map<string, Lot[]>()

  const nextRows = data.rows.map((row) => {
    const symbol = row[symbolIndex]?.trim().toUpperCase() || ''
    const quantityRaw = row[quantityIndex] ?? ''
    const quantitySigned = parseNumber(quantityRaw)
    const quantity = Math.abs(quantitySigned)
    const actionRaw = actionIndex >= 0 ? row[actionIndex] ?? '' : ''
    const action = resolveAction(actionRaw, quantitySigned)
    const price = priceIndex >= 0 ? parseNumber(row[priceIndex] ?? '') : Number.NaN
    const total = totalIndex >= 0 ? parseNumber(row[totalIndex] ?? '') : Number.NaN

    if (!symbol || !Number.isFinite(quantity) || quantity <= 0 || !action) {
      return [...row, '', '']
    }

    if (action === 'buy') {
      const unitCost = Number.isFinite(price)
        ? Math.abs(price)
        : Number.isFinite(total)
          ? Math.abs(total) / quantity
          : Number.NaN

      if (Number.isFinite(unitCost) && unitCost > 0) {
        const lots = inventory.get(symbol) ?? []
        lots.push({ quantity, unitCost })
        inventory.set(symbol, lots)
      }

      return [...row, '', '']
    }

    const lots = inventory.get(symbol) ?? []
    let remaining = quantity
    let matchedQuantity = 0
    let cost = 0

    while (remaining > 0 && lots.length > 0) {
      const currentLot = lots[0]
      const matched = Math.min(remaining, currentLot.quantity)

      cost += matched * currentLot.unitCost
      matchedQuantity += matched
      currentLot.quantity -= matched
      remaining -= matched

      if (currentLot.quantity <= 0) {
        lots.shift()
      }
    }

    const proceeds = Number.isFinite(total)
      ? Math.abs(total)
      : Number.isFinite(price)
        ? quantity * Math.abs(price)
        : Number.NaN

    if (!Number.isFinite(proceeds) || matchedQuantity <= 0 || cost <= 0) {
      return [...row, '', '']
    }

    const profitLoss = proceeds - cost
    const percentage = (profitLoss / cost) * 100

    return [...row, formatSignedNumber(profitLoss), `${formatSignedNumber(percentage)}%`]
  })

  return { headers: nextHeaders, rows: nextRows }
}
