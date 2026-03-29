import type { SheetData } from './sheetsData'
import { supabase } from './supabaseClient'

export type AppUser = {
  id: string
  full_name: string
  email: string | null
}

type StockRow = {
  stock: string
  currency: string
  price: number
  action: string
  time: string
  quantity: number
  handling_fees: number
}

type DividendRow = {
  stock: string
  currency: string
  div: number
  time: string
}

type MoneyMoveRow = {
  name: string
  currency: string
  price: number
  time: string
  action: string
}

const STOCK_HEADERS = ['Stock', 'Currency', 'Price', 'Action', 'Time', 'Quantity', 'Handling Fees']
const DIVIDEND_HEADERS = ['Stock', 'Currency', 'Div', 'Time']
const MONEY_MOVE_HEADERS = ['Name', 'Currency', 'Price', 'Time', 'Action']

function toSheetData(headers: string[], rows: string[][]): SheetData {
  return { headers, rows }
}

function normalizeNumberInput(value: string): number {
  const parsed = Number(value)
  if (!Number.isFinite(parsed)) {
    throw new Error(`Invalid number: ${value}`)
  }
  return parsed
}

export async function ensureUserProfile(userId: string, fullName: string, email: string | null): Promise<void> {
  const { error } = await supabase.from('user').upsert(
    {
      id: userId,
      full_name: fullName,
      email,
      updated_at: new Date().toISOString(),
    },
    { onConflict: 'id' },
  )

  if (error) {
    throw new Error(error.message)
  }
}

export async function loadUserSheetData(userId: string): Promise<{
  stocks: SheetData
  dividend: SheetData
  moneyMove: SheetData
}> {
  const [stocksResult, dividendResult, moneyMoveResult] = await Promise.all([
    supabase
      .from('stocks')
      .select('stock, currency, price, action, time, quantity, handling_fees')
      .eq('user_id', userId)
      .order('time', { ascending: true }),
    supabase
      .from('dividend')
      .select('stock, currency, div, time')
      .eq('user_id', userId)
      .order('time', { ascending: true }),
    supabase
      .from('money_move')
      .select('name, currency, price, time, action')
      .eq('user_id', userId)
      .order('time', { ascending: true }),
  ])

  if (stocksResult.error) {
    throw new Error(stocksResult.error.message)
  }
  if (dividendResult.error) {
    throw new Error(dividendResult.error.message)
  }
  if (moneyMoveResult.error) {
    throw new Error(moneyMoveResult.error.message)
  }

  const stocksRows = (stocksResult.data ?? []).map((row: StockRow) => [
    row.stock ?? '',
    row.currency ?? '',
    String(row.price ?? ''),
    row.action ?? '',
    row.time ?? '',
    String(row.quantity ?? ''),
    String(row.handling_fees ?? ''),
  ])

  const dividendRows = (dividendResult.data ?? []).map((row: DividendRow) => [
    row.stock ?? '',
    row.currency ?? '',
    String(row.div ?? ''),
    row.time ?? '',
  ])

  const moneyMoveRows = (moneyMoveResult.data ?? []).map((row: MoneyMoveRow) => [
    row.name ?? '',
    row.currency ?? '',
    String(row.price ?? ''),
    row.time ?? '',
    row.action ?? '',
  ])

  return {
    stocks: toSheetData(STOCK_HEADERS, stocksRows),
    dividend: toSheetData(DIVIDEND_HEADERS, dividendRows),
    moneyMove: toSheetData(MONEY_MOVE_HEADERS, moneyMoveRows),
  }
}

export async function insertUserRow(
  userId: string,
  table: 'Stocks' | 'Dividend' | 'Money Move',
  row: Record<string, string>,
): Promise<void> {
  if (table === 'Stocks') {
    const payload = {
      user_id: userId,
      stock: row.stock,
      currency: row.currency,
      price: normalizeNumberInput(row.price),
      action: row.action,
      time: row.time,
      quantity: normalizeNumberInput(row.quantity),
      handling_fees: normalizeNumberInput(row.handlingFees ?? '0'),
    }

    const { error } = await supabase.from('stocks').insert(payload)
    if (error) {
      throw new Error(error.message)
    }
    return
  }

  if (table === 'Dividend') {
    const payload = {
      user_id: userId,
      stock: row.stock,
      currency: row.currency,
      div: normalizeNumberInput(row.div),
      time: row.time,
    }

    const { error } = await supabase.from('dividend').insert(payload)
    if (error) {
      throw new Error(error.message)
    }
    return
  }

  const payload = {
    user_id: userId,
    name: row.name,
    currency: row.currency,
    price: normalizeNumberInput(row.price),
    time: row.time,
    action: row.do,
  }

  const { error } = await supabase.from('money_move').insert(payload)
  if (error) {
    throw new Error(error.message)
  }
}
