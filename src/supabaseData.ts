import { supabase } from './supabaseClient'
import { normalizeDateTimeInput } from './dateUtils'

export type AppUser = {
  id: string
  full_name: string
  email: string | null
}

type StockRow = {
  id?: string
  stock: string
  currency: string
  price: number
  action: string
  time: string
  quantity: number
  handling_fees: number
}

type DividendRow = {
  id?: string
  stock: string
  currency: string
  div: number
  time: string
}

type MoneyMoveRow = {
  id?: string
  name: string
  currency: string
  price: number
  time: string
  action: string
}

export type SheetDataWithIds = {
  headers: string[]
  rows: string[][]
  ids: string[]
}

export type UserTableName = 'Stocks' | 'Dividend' | 'Money Move'
export type RowMutationAction = 'update' | 'delete'

export type GoalScheduleType = 'frequency' | 'deadline'

export type GoalRow = {
  id: string
  user_id: string
  metric: string
  target_value: number | string
  target_currency: string
  schedule_type: GoalScheduleType
  frequency: string | null
  frequency_number: number | string | null
  deadline: string | null
  created_at: string
}

export type GoalInsertPayload = {
  id: string
  metric: string
  targetValue: number
  targetCurrency: string
  scheduleType: GoalScheduleType
  frequency: string | null
  frequencyNumber: number
  deadline: string | null
}

const STOCK_HEADERS = ['Stock', 'Currency', 'Price', 'Action', 'Time', 'Quantity', 'Handling Fees']
const DIVIDEND_HEADERS = ['Stock', 'Currency', 'Div', 'Time']
const MONEY_MOVE_HEADERS = ['Name', 'Currency', 'Price', 'Time', 'Action']

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
  stocks: SheetDataWithIds
  dividend: SheetDataWithIds
  moneyMove: SheetDataWithIds
}> {
  const [stocksResult, dividendResult, moneyMoveResult] = await Promise.all([
    supabase
      .from('stocks')
      .select('id, stock, currency, price, action, time, quantity, handling_fees')
      .eq('user_id', userId)
      .order('time', { ascending: true }),
    supabase
      .from('dividend')
      .select('id, stock, currency, div, time')
      .eq('user_id', userId)
      .order('time', { ascending: true }),
    supabase
      .from('money_move')
      .select('id, name, currency, price, time, action')
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
  const stocksIds = (stocksResult.data ?? []).map((row: StockRow) => row.id ?? '')

  const dividendRows = (dividendResult.data ?? []).map((row: DividendRow) => [
    row.stock ?? '',
    row.currency ?? '',
    String(row.div ?? ''),
    row.time ?? '',
  ])
  const dividendIds = (dividendResult.data ?? []).map((row: DividendRow) => row.id ?? '')

  const moneyMoveRows = (moneyMoveResult.data ?? []).map((row: MoneyMoveRow) => [
    row.name ?? '',
    row.currency ?? '',
    String(row.price ?? ''),
    row.time ?? '',
    row.action ?? '',
  ])
  const moneyMoveIds = (moneyMoveResult.data ?? []).map((row: MoneyMoveRow) => row.id ?? '')

  return {
    stocks: { headers: STOCK_HEADERS, rows: stocksRows, ids: stocksIds },
    dividend: { headers: DIVIDEND_HEADERS, rows: dividendRows, ids: dividendIds },
    moneyMove: { headers: MONEY_MOVE_HEADERS, rows: moneyMoveRows, ids: moneyMoveIds },
  }
}

export async function loadUserGoals(userId: string): Promise<GoalRow[]> {
  const { data, error } = await supabase
    .from('goals')
    .select('id, user_id, metric, target_value, target_currency, schedule_type, frequency, frequency_number, deadline, created_at')
    .eq('user_id', userId)
    .order('created_at', { ascending: false })

  if (error) {
    throw new Error(error.message)
  }

  return (data ?? []) as GoalRow[]
}

export async function insertUserGoal(userId: string, goal: GoalInsertPayload): Promise<void> {
  const payload = {
    id: goal.id,
    user_id: userId,
    metric: goal.metric,
    target_value: goal.targetValue,
    target_currency: goal.targetCurrency,
    schedule_type: goal.scheduleType,
    frequency: goal.frequency,
    frequency_number: goal.frequencyNumber,
    deadline: goal.deadline,
  }

  const { error } = await supabase.from('goals').insert(payload)

  if (error) {
    throw new Error(error.message)
  }
}

export async function deleteUserGoal(goalId: string): Promise<void> {
  const { error } = await supabase.from('goals').delete().eq('id', goalId).select()

  if (error) {
    throw new Error(error.message)
  }
}

export async function insertUserRow(
  userId: string,
  table: UserTableName,
  row: Record<string, string>,
): Promise<void> {
  if (table === 'Stocks') {
    const payload = {
      user_id: userId,
      stock: row.stock,
      currency: row.currency,
      price: normalizeNumberInput(row.price),
      action: row.action,
      time: normalizeDateTimeInput(row.time),
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
      time: normalizeDateTimeInput(row.time),
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
    time: normalizeDateTimeInput(row.time),
    action: row.do,
  }

  const { error } = await supabase.from('money_move').insert(payload)
  if (error) {
    throw new Error(error.message)
  }
}

export async function updateUserRow(
  table: UserTableName,
  recordId: string,
  row: Record<string, string>,
): Promise<void> {
  // console.log('Updating row', { table, recordId, row }) 
  if (table === 'Stocks') {
    const payload = {
      stock: row.stock,
      currency: row.currency,
      price: normalizeNumberInput(row.price),
      action: row.action,
      time: normalizeDateTimeInput(row.time),
      quantity: normalizeNumberInput(row.quantity),
      handling_fees: normalizeNumberInput(row.handlingFees ?? '0'),
    }

    const { error } = await supabase
      .from('stocks')
      .update(payload)
      .eq('id', recordId)
      .select()
    // console.log('Update result', { table, recordId, payload, error })
    if (error) {
      console.error('[Supabase][stocks][update] failed', {
        table,
        recordId,
        payload,
        error,
      })
      throw new Error(error.message)
    }
    return
  }

  if (table === 'Dividend') {
    const payload = {
      stock: row.stock,
      currency: row.currency,
      div: normalizeNumberInput(row.div),
      time: normalizeDateTimeInput(row.time),
    }

    const { error } = await supabase
      .from('dividend')
      .update(payload)
      .eq('id', recordId)
      .select()
    if (error) {
      console.error('[Supabase][dividend][update] failed', {
        table,
        recordId,
        payload,
        error,
      })
      throw new Error(error.message)
    }
    return
  }

  const payload = {
    name: row.name,
    currency: row.currency,
    price: normalizeNumberInput(row.price),
    time: normalizeDateTimeInput(row.time),
    action: row.do,
  }

  const { error } = await supabase
    .from('money_move')
    .update(payload)
    .eq('id', recordId)
    .select()
  if (error) {
    console.error('[Supabase][money_move][update] failed', {
      table,
      recordId,
      payload,
      error,
    })
    throw new Error(error.message)
  }
}

export async function deleteUserRow(
  table: UserTableName,
  recordId: string,
): Promise<void> {
  const { error } = await supabase
    .from(table === 'Stocks' ? 'stocks' : table === 'Dividend' ? 'dividend' : 'money_move')
    .delete()
    .eq('id', recordId)
    .select()

  if (error) {
    console.error('[Supabase][delete] failed', {
      table,
      recordId,
      error,
    })
    throw new Error(error.message)
  }
}

export async function mutateUserRowAndReload(
  userId: string,
  table: UserTableName,
  recordId: string,
  action: RowMutationAction,
  row: Record<string, string>,
): Promise<{
  stocks: SheetDataWithIds
  dividend: SheetDataWithIds
  moneyMove: SheetDataWithIds
}> {
  if (action === 'delete') {
    await deleteUserRow(table, recordId)
  } else {
    await updateUserRow(table, recordId, row)
  }

  return loadUserSheetData(userId)
}
