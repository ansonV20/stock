import { useEffect, useMemo, useRef, useState } from 'react'
import Alert from '@mui/material/Alert'
import CircularProgress from '@mui/material/CircularProgress'
import MenuItem from '@mui/material/MenuItem'
import Menu from '@mui/material/Menu'
import Paper from '@mui/material/Paper'
import Switch from '@mui/material/Switch'
import Table from '@mui/material/Table'
import TableBody from '@mui/material/TableBody'
import TableCell from '@mui/material/TableCell'
import TableContainer from '@mui/material/TableContainer'
import TableHead from '@mui/material/TableHead'
import TableRow from '@mui/material/TableRow'
import ToggleButton from '@mui/material/ToggleButton'
import ToggleButtonGroup from '@mui/material/ToggleButtonGroup'
import Autocomplete from '@mui/material/Autocomplete'
import Divider from '@mui/material/Divider'
import TextField from '@mui/material/TextField'
import type { Session } from '@supabase/supabase-js'
import {
  enrichStocksWithProfitLoss,
  parseNumber,
  resolveAction,
  type SheetData,
} from './sheetsData'
import { Box, Button, Modal } from '@mui/material'
import {
  MdDataUsage,
  MdTableRows,
  MdOutlineAdd,
  MdSettings,
  MdShowChart,
  MdBrightnessAuto,
  MdDarkMode,
  MdLightMode,
  MdCancel,
  MdEdit,
  MdQueryStats
} from 'react-icons/md'
import { LineChart } from '@mui/x-charts/LineChart'
import { PieChart } from '@mui/x-charts/PieChart'
import { BarChart } from '@mui/x-charts/BarChart'
import { Gauge } from '@mui/x-charts/Gauge'
import { supabase } from './supabaseClient'
import {
  deleteUserGoal,
  ensureUserProfile,
  insertUserGoal,
  insertUserRow,
  loadUserGoals,
  loadUserSheetData,
  mutateUserRowAndReload,
  type GoalRow,
  type GoalScheduleType,
  type SheetDataWithIds,
} from './supabaseData'
import { downloadExcelTemplate, parseExcelFile, excelRowToObject, type ExcelDataRow } from './excelData'
import './App.css'

const CURRENCY_OPTIONS = ['USD', 'HKD'] as const
type CurrencyCode = (typeof CURRENCY_OPTIONS)[number]

type CurrencyRatesCache = {
  timestamp: number
  rates: Record<CurrencyCode, number>
}

type HoldingSnapshot = {
  symbol: string
  displayName: string
  quantity: number
  avgBuyPriceUsd: number
}

type FinnhubQuoteResponse = {
  c?: number
  d?: number
  dp?: number
  h?: number
  l?: number
  o?: number
  pc?: number
}

type FinnhubMarketStatusResponse = {
  isOpen?: boolean
}

type LiveSymbolsCookiePayload = {
  user: string
  followList: string[]
}

const DEFAULT_RATES: Record<CurrencyCode, number> = {
  USD: 1,
  HKD: 7.8,
}

const CURRENCY_COOKIE_KEY = 'currency_rates_cache_v1'
const COOKIE_TTL_MS = 2 * 24 * 60 * 60 * 1000
const CURRENCY_API_URL =
  'https://api.currencyapi.com/v3/latest?apikey=' + import.meta.env.VITE_CURRENCYAPI_KEY?.trim() + '&currencies=USD%2CHKD'
const FINNHUB_TOKEN = import.meta.env.VITE_FINNHUB_TOKEN?.trim()
const FINNHUB_QUOTE_URL = 'https://finnhub.io/api/v1/quote'
const FINNHUB_MARKET_STATUS_URL = 'https://finnhub.io/api/v1/stock/market-status'
const LIVE_QUOTES_NORMAL_INTERVAL_MS = 5 * 60 * 1000
const LIVE_QUOTES_FAST_INTERVAL_MS = 60 * 1000
const MARKET_OPEN_CHECK_INTERVAL_MS = 15 * 1000
const US_MARKET_TIME_ZONE = 'America/New_York'
const US_MARKET_OPEN_MINUTES = 9 * 60 + 30
const US_MARKET_CLOSE_MINUTES = 16 * 60
const AUTH_REDIRECT_URL = import.meta.env.VITE_AUTH_REDIRECT_URL?.trim()
const THEME_STORAGE_KEY = 'stock_theme_mode_v1'
const NOTIFICATIONS_STORAGE_KEY = 'stock_notifications_enabled_v1'
const LIVE_SYMBOLS_STORAGE_KEY = 'stock_live_symbols_v1'
const LIVE_SYMBOLS_COOKIE_KEY = 'stock_live_symbols_cookie_v1'
const PRICE_ALERTS_STORAGE_KEY = 'stock_price_alerts_v1'
const MARKET_OPEN_NOTICE_DAY_KEY = 'stock_market_open_notice_day_v1'
const NOTIFICATION_SW_PATH = '/notification-sw.js'

type AddTableName = 'Stocks' | 'Dividend' | 'Money Move'
type PageName = 'table' | 'dataShow' | 'goal' | 'add' | 'live'
type ThemeMode = 'system' | 'light' | 'dark'
type PriceAlertCondition = 'above' | 'below'

type PriceAlert = {
  id: string
  symbol: string
  targetPrice: number
  condition: PriceAlertCondition
  enabled: boolean
  triggered: boolean
}

type AddFieldConfig = {
  key: string
  label: string
  inputType: 'text' | 'number' | 'datetime' | 'select'
  required?: boolean
  options?: string[]
  step?: string
  defaultValue?: string
}

type GoalMetric = 'total-money' | 'earn-money' | 'earn-percent' | 'avg-earn-percent' | 'trade-count'
type GoalFrequency = 'day' | 'week' | 'month' | 'year'

type GoalFieldConfig = {
  key: string
  label: string
  inputType: 'text' | 'number' | 'select' | 'date'
  required?: boolean
  options?: string[]
  step?: string
  defaultValue?: string
}

type GoalRecord = {
  id: string
  metric: GoalMetric
  targetValue: number
  targetCurrency: CurrencyCode
  scheduleType: GoalScheduleType
  frequency: GoalFrequency | null
  frequencyNumber: number
  deadline: string | null
  createdAt: string
}

type GoalMetricOption = {
  value: GoalMetric
  label: string
  unit: 'currency' | 'percent' | 'count'
}

type GoalMetricSnapshot = {
  totalMoney: number
  earnMoney: number
  earnPercent: number
  avgEarnPercent: number
  tradeCount: number
}

type GoalPeriodMetrics = {
  earnMoney: number
  earnPercent: number
  avgEarnPercent: number
  tradeCount: number
}

const ADD_TABLE_OPTIONS: AddTableName[] = ['Stocks', 'Dividend', 'Money Move']
const EXCEL_PREVIEW_HEADERS: Record<AddTableName, string[]> = {
  Stocks: ['Stock', 'Currency', 'Price', 'Action', 'Time', 'Quantity', 'Handling Fees'],
  Dividend: ['Stock', 'Currency', 'Div', 'Time'],
  'Money Move': ['Name', 'Currency', 'Price', 'Time', 'Action'],
}

const ADD_FORM_CONFIG: Record<AddTableName, AddFieldConfig[]> = {
  Stocks: [
    { key: 'stock', label: 'Stock', inputType: 'text', required: true },
    {
      key: 'currency',
      label: 'Currency',
      inputType: 'select',
      options: [...CURRENCY_OPTIONS],
      required: true,
      defaultValue: 'USD',
    },
    { key: 'price', label: 'Price', inputType: 'number', step: '0.01', required: true },
    {
      key: 'action',
      label: 'Action',
      inputType: 'select',
      options: ['Buy', 'Sell'],
      required: true,
      defaultValue: 'Buy',
    },
    { key: 'time', label: 'Time', inputType: 'datetime', required: true },
    { key: 'quantity', label: 'Quantity', inputType: 'number', step: '1', required: true },
    {
      key: 'handlingFees',
      label: 'Handiling Fees',
      inputType: 'number',
      step: '0.01',
      required: true,
      defaultValue: '0',
    },
  ],
  Dividend: [
    { key: 'stock', label: 'Stock', inputType: 'text', required: true },
    {
      key: 'currency',
      label: 'Currency',
      inputType: 'select',
      options: [...CURRENCY_OPTIONS],
      required: true,
      defaultValue: 'USD',
    },
    { key: 'div', label: 'Div', inputType: 'number', step: '0.01', required: true },
    { key: 'time', label: 'Time', inputType: 'datetime', required: true },
  ],
  'Money Move': [
    { key: 'name', label: 'Name', inputType: 'text', required: true },
    {
      key: 'currency',
      label: 'Currency',
      inputType: 'select',
      options: [...CURRENCY_OPTIONS],
      required: true,
      defaultValue: 'USD',
    },
    { key: 'price', label: 'Price', inputType: 'number', step: '0.01', required: true },
    { key: 'time', label: 'Time', inputType: 'datetime', required: true },
    {
      key: 'do',
      label: 'Do',
      inputType: 'select',
      options: ['In', 'Out', 'Bor', 'Back'],
      required: true,
      defaultValue: 'In',
    },
  ],
}

const GOAL_METRIC_OPTIONS: GoalMetricOption[] = [
  { value: 'total-money', label: 'Total money', unit: 'currency' },
  { value: 'earn-money', label: 'Earn money', unit: 'currency' },
  { value: 'earn-percent', label: 'Earn %', unit: 'percent' },
  { value: 'avg-earn-percent', label: 'Avg earn %', unit: 'percent' },
  { value: 'trade-count', label: 'No. of trade', unit: 'count' },
]

const GOAL_FREQUENCY_OPTIONS: Array<{ value: GoalFrequency; label: string }> = [
  { value: 'day', label: 'Day' },
  { value: 'week', label: 'Week' },
  { value: 'month', label: 'Month' },
  { value: 'year', label: 'Year' },
]

const GOAL_FORM_CONFIG: GoalFieldConfig[] = [
  {
    key: 'metric',
    label: 'Goal type',
    inputType: 'select',
    required: true,
    options: GOAL_METRIC_OPTIONS.map((option) => option.value),
    defaultValue: 'total-money',
  },
  { key: 'targetValue', label: 'Target value', inputType: 'number', step: '0.01', required: true },
  {
    key: 'targetCurrency',
    label: 'Target currency',
    inputType: 'select',
    required: true,
    options: [...CURRENCY_OPTIONS],
    defaultValue: 'USD',
  },
  {
    key: 'scheduleType',
    label: 'Schedule type',
    inputType: 'select',
    required: true,
    options: ['frequency', 'deadline'],
    defaultValue: 'frequency',
  },
  {
    key: 'frequency',
    label: 'Frequency',
    inputType: 'select',
    required: true,
    options: GOAL_FREQUENCY_OPTIONS.map((option) => option.value),
    defaultValue: 'month',
  },
  {
    key: 'frequencyNumber',
    label: 'Frequency number',
    inputType: 'number',
    required: true,
    step: '1',
    defaultValue: '1',
  },
  { key: 'deadline', label: 'Deadline', inputType: 'date' },
]

function getDefaultFormValues(table: AddTableName): Record<string, string> {
  return ADD_FORM_CONFIG[table].reduce<Record<string, string>>((acc, field) => {
    acc[field.key] = field.defaultValue ?? ''
    return acc
  }, {})
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

function formatCurrency(value: number, currency: CurrencyCode): string {
  return new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency,
    maximumFractionDigits: 2,
  }).format(value)
}

function formatPercent(value: number): string {
  return `${value.toFixed(2)}%`
}

function resolveMoneyMoveType(value: string): 'in' | 'bor' | 'back' | 'out' | null {
  const normalized = value.trim().toLowerCase()

  if (normalized === 'in' || normalized.includes('deposit')) {
    return 'in'
  }

  if (normalized === 'bor' || normalized.includes('borrow')) {
    return 'bor'
  }

  if (normalized === 'back' || normalized.includes('repay')) {
    return 'back'
  }

  if (normalized === 'out' || normalized.includes('withdraw')) {
    return 'out'
  }

  return null
}

function getCookie(name: string): string | null {
  const cookieEntries = document.cookie ? document.cookie.split('; ') : []
  const keyPrefix = `${name}=`
  const matched = cookieEntries.find((entry) => entry.startsWith(keyPrefix))

  if (!matched) {
    return null
  }

  return decodeURIComponent(matched.slice(keyPrefix.length))
}

function setCookie(name: string, value: string, maxAgeSeconds: number): void {
  document.cookie = `${name}=${encodeURIComponent(value)}; max-age=${maxAgeSeconds}; path=/; samesite=lax`
}

function convertAmount(
  amount: number,
  fromCurrency: CurrencyCode,
  toCurrency: CurrencyCode,
  rates: Record<CurrencyCode, number>,
): number {
  if (fromCurrency === toCurrency) {
    return amount
  }

  const fromRate = rates[fromCurrency]
  const toRate = rates[toCurrency]

  if (!Number.isFinite(fromRate) || !Number.isFinite(toRate) || fromRate <= 0 || toRate <= 0) {
    return amount
  }

  const amountInUsd = amount / fromRate
  return amountInUsd * toRate
}

function resolveCurrencyCode(value: string | undefined): CurrencyCode {
  const normalized = (value ?? '').trim().toUpperCase()
  return normalized === 'HKD' ? 'HKD' : 'USD'
}

function parseDateValue(raw: string): Date | null {
  const trimmed = raw.trim()
  if (!trimmed) {
    return null
  }

  const googleDateMatch = trimmed.match(
    /^Date\((\d+),\s*(\d+),\s*(\d+)(?:,\s*(\d+),\s*(\d+),\s*(\d+))?\)$/i,
  )
  if (googleDateMatch) {
    const [, yearRaw, monthRaw, dayRaw, hourRaw, minuteRaw, secondRaw] = googleDateMatch
    const year = Number(yearRaw)
    // Google Charts Date(...) month is already zero-based.
    const month = Number(monthRaw)
    const day = Number(dayRaw)
    const hour = hourRaw ? Number(hourRaw) : 0
    const minute = minuteRaw ? Number(minuteRaw) : 0
    const second = secondRaw ? Number(secondRaw) : 0

    const parsedGoogleDate = new Date(year, month, day, hour, minute, second)
    return Number.isNaN(parsedGoogleDate.getTime()) ? null : parsedGoogleDate
  }

  const normalized = trimmed.replace(/\./g, '-').replace(/\//g, '-')
  const parsed = new Date(normalized)
  return Number.isNaN(parsed.getTime()) ? null : parsed
}

function toMonthKey(date: Date): string {
  const year = date.getFullYear()
  const month = String(date.getMonth() + 1).padStart(2, '0')
  return `${year}-${month}`
}

function formatMonthLabel(monthKey: string): string {
  const matched = monthKey.match(/^(\d{4})-(\d{2})$/)
  if (!matched) {
    return monthKey
  }

  const [, year, month] = matched
  return `${month}/${year.slice(-2)}`
}

function formatDateDDMMYYYY(date: Date): string {
  const day = String(date.getDate()).padStart(2, '0')
  const month = String(date.getMonth() + 1).padStart(2, '0')
  const year = date.getFullYear()
  return `${day}/${month}/${year}`
}

function formatTimelineLabel(date: Date, index: number): string {
  return `${formatDateDDMMYYYY(date)} #${String(index + 1).padStart(2, '0')}`
}

function formatCompactAxisDate(label: string): string {
  const datePart = label.split(' #')[0]
  const match = datePart.match(/^(\d{2})\/(\d{2})\/(\d{4})$/)

  if (!match) {
    return datePart
  }

  const [, day, month, year] = match
  return `${day}/${month}/${year.slice(-2)}`
}

function buildMonthlyTotalMoneySeries(
  moneyMoveData: SheetDataWithIds,
  stocksData: SheetDataWithIds,
  dividendData: SheetDataWithIds,
  selectedCurrency: CurrencyCode,
  currencyRates: Record<CurrencyCode, number>,
): Map<string, number> {
  type TimelineEvent = {
    date: Date
    sourceOrder: number
    cashDelta: number
    stockAction?: { symbol: string; action: 'buy' | 'sell'; quantity: number; costPerShare: number }
  }

  const events: TimelineEvent[] = []

  const moneyMoveAmountIndex = findColumnIndex(moneyMoveData.headers, [
    'amount',
    'total',
    'value',
    'money',
    'price',
  ])
  const moneyMoveActionIndex = findColumnIndex(moneyMoveData.headers, [
    'action',
    'type',
    'note',
    'description',
    'do',
  ])
  const moneyMoveDateIndex = findColumnIndex(moneyMoveData.headers, ['date', 'time'])
  const moneyMoveCurrencyIndex = findColumnIndex(moneyMoveData.headers, ['currency'])

  moneyMoveData.rows.forEach((row, rowIndex) => {
    if (moneyMoveAmountIndex === -1 || moneyMoveActionIndex === -1 || moneyMoveDateIndex === -1) {
      return
    }

    const amount = parseNumber(row[moneyMoveAmountIndex] ?? '')
    const action = resolveMoneyMoveType(row[moneyMoveActionIndex] ?? '')
    const date = parseDateValue(row[moneyMoveDateIndex] ?? '')

    if (!Number.isFinite(amount) || !action || !date) {
      return
    }

    const rowCurrency = resolveCurrencyCode(row[moneyMoveCurrencyIndex] ?? '')
    const normalizedAmount = convertAmount(
      Math.abs(amount),
      rowCurrency,
      selectedCurrency,
      currencyRates,
    )

    const delta = action === 'in' || action === 'bor' ? normalizedAmount : -normalizedAmount

    events.push({
      date,
      sourceOrder: rowIndex,
      cashDelta: delta,
    })
  })

  const stocksDateIndex = findColumnIndex(stocksData.headers, ['date', 'time'])
  const stocksActionIndex = findColumnIndex(stocksData.headers, [
    'action',
    'type',
    'side',
    'buysell',
    'operation',
  ])
  const stocksSymbolIndex = findColumnIndex(stocksData.headers, [
    'symbol',
    'ticker',
    'stock',
    'code',
    'company',
  ])
  const stocksQuantityIndex = findColumnIndex(stocksData.headers, [
    'qty',
    'quantity',
    'share',
    'units',
    'vol',
    'volume',
  ])
  const stocksTotalIndex = findColumnIndex(stocksData.headers, [
    'total',
    'amount',
    'value',
    'proceeds',
    'market value',
    'marketvalue',
  ])
  const stocksPriceIndex = findColumnIndex(stocksData.headers, [
    'price',
    'avgprice',
    'unitprice',
    'cost',
    'rate',
  ])
  const stocksCurrencyIndex = findColumnIndex(stocksData.headers, ['currency'])
  const stocksFeeIndex = findColumnIndex(stocksData.headers, [
    'handlingfee',
    'handling',
    'fee',
    'commission',
    'charge',
    'charges',
  ])

  stocksData.rows.forEach((row, rowIndex) => {
    if (stocksDateIndex === -1 || stocksQuantityIndex === -1) {
      return
    }

    const date = parseDateValue(row[stocksDateIndex] ?? '')
    const quantitySigned = parseNumber(row[stocksQuantityIndex] ?? '')
    const actionRaw = stocksActionIndex >= 0 ? row[stocksActionIndex] ?? '' : ''
    const action = resolveAction(actionRaw, quantitySigned) ?? (quantitySigned > 0 ? 'buy' : 'sell')

    let total = stocksTotalIndex >= 0 ? parseNumber(row[stocksTotalIndex] ?? '') : Number.NaN
    if (!Number.isFinite(total) && stocksPriceIndex >= 0) {
      const price = parseNumber(row[stocksPriceIndex] ?? '')
      if (Number.isFinite(price)) {
        total = Math.abs(price * Math.abs(quantitySigned))
      }
    }

    if (!date || !Number.isFinite(total)) {
      return
    }

    const rowCurrency = resolveCurrencyCode(row[stocksCurrencyIndex] ?? '')
    const tradeAmount = convertAmount(
      Math.abs(total),
      rowCurrency,
      selectedCurrency,
      currencyRates,
    )

    const feeRaw = stocksFeeIndex >= 0 ? parseNumber(row[stocksFeeIndex] ?? '') : 0
    const fee = Number.isFinite(feeRaw)
      ? convertAmount(Math.abs(feeRaw), rowCurrency, selectedCurrency, currencyRates)
      : 0

    const quantity = Math.abs(quantitySigned)
    const costPerShare = quantity > 0 ? tradeAmount / quantity : 0
    const symbol = stocksSymbolIndex >= 0 ? (row[stocksSymbolIndex] ?? '').trim().toUpperCase() : 'UNKNOWN'
    const cashDelta = action === 'buy' ? -(tradeAmount + fee) : tradeAmount - fee

    events.push({
      date,
      sourceOrder: 10_000 + rowIndex,
      cashDelta,
      stockAction:
        quantity > 0
          ? { symbol, action, quantity, costPerShare }
          : undefined,
    })
  })

  const dividendAmountIndex = findColumnIndex(dividendData.headers, ['dividend', 'div', 'amount', 'value'])
  const dividendDateIndex = findColumnIndex(dividendData.headers, ['date', 'time'])
  const dividendCurrencyIndex = findColumnIndex(dividendData.headers, ['currency'])

  dividendData.rows.forEach((row, rowIndex) => {
    if (dividendAmountIndex === -1 || dividendDateIndex === -1) {
      return
    }

    const amount = parseNumber(row[dividendAmountIndex] ?? '')
    const date = parseDateValue(row[dividendDateIndex] ?? '')

    if (!Number.isFinite(amount) || !date) {
      return
    }

    const rowCurrency = resolveCurrencyCode(row[dividendCurrencyIndex] ?? '')
    const normalizedAmount = convertAmount(amount, rowCurrency, selectedCurrency, currencyRates)

    events.push({
      date,
      sourceOrder: 20_000 + rowIndex,
      cashDelta: normalizedAmount,
    })
  })

  events.sort((a, b) => {
    const timeDiff = a.date.getTime() - b.date.getTime()
    if (timeDiff !== 0) {
      return timeDiff
    }
    return a.sourceOrder - b.sourceOrder
  })

  let runningCash = 0
  const holdingsBySymbol = new Map<string, { quantity: number; costBasis: number }>()
  const monthLatestTotalMoney = new Map<string, number>()

  events.forEach((event) => {
    runningCash += event.cashDelta

    if (event.stockAction) {
      const { symbol, action, quantity, costPerShare } = event.stockAction
      const costForThisPurchase = quantity * costPerShare
      const holding = holdingsBySymbol.get(symbol) ?? { quantity: 0, costBasis: 0 }

      if (action === 'buy') {
        holding.quantity += quantity
        holding.costBasis += costForThisPurchase
      } else if (action === 'sell') {
        const quantityToSell = Math.min(quantity, holding.quantity)
        const avgCostPerShare = holding.quantity > 0 ? holding.costBasis / holding.quantity : 0
        const soldCostBasis = quantityToSell * avgCostPerShare
        holding.quantity -= quantityToSell
        holding.costBasis = Math.max(0, holding.costBasis - soldCostBasis)

        if (holding.quantity < 0.0001) {
          holding.quantity = 0
          holding.costBasis = 0
        }
      }

      holdingsBySymbol.set(symbol, holding)
    }

    const totalHoldingsCostBasis = Array.from(holdingsBySymbol.values()).reduce(
      (sum, holding) => sum + holding.costBasis,
      0,
    )
    const totalMoney = runningCash + totalHoldingsCostBasis
    const monthKey = toMonthKey(event.date)
    monthLatestTotalMoney.set(monthKey, Number(totalMoney.toFixed(2)))
  })

  return monthLatestTotalMoney
}

function resolveOAuthRedirectUrl(): string {
  if (typeof window === 'undefined') {
    return AUTH_REDIRECT_URL ?? ''
  }

  const isLocalhost =
    window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1'
  if (isLocalhost) {
    return window.location.origin
  }

  if (!AUTH_REDIRECT_URL) {
    return window.location.origin
  }

  try {
    return new URL(AUTH_REDIRECT_URL).toString()
  } catch {
    return window.location.origin
  }
}

function normalizeSymbolInput(input: string): string | null {
  const normalized = input.trim().toUpperCase()

  if (!/^[A-Z.]{1,10}$/.test(normalized)) {
    return null
  }

  return normalized
}

function sanitizeSymbols(input: unknown): string[] {
  if (!Array.isArray(input)) {
    return []
  }

  const normalized = input
    .map((symbol) => String(symbol).trim().toUpperCase())
    .filter((symbol) => /^[A-Z.]{1,10}$/.test(symbol))

  return Array.from(new Set(normalized))
}

function resolveLiveQuotePrice(quote: FinnhubQuoteResponse | undefined): number {
  if (!quote) {
    return Number.NaN
  }

  if (typeof quote.c === 'number' && Number.isFinite(quote.c) && quote.c > 0) {
    return quote.c
  }

  if (typeof quote.pc === 'number' && Number.isFinite(quote.pc) && quote.pc > 0) {
    return quote.pc
  }

  return Number.NaN
}

function isGoalMetric(value: string): value is GoalMetric {
  return (
    value === 'total-money' ||
    value === 'earn-money' ||
    value === 'earn-percent' ||
    value === 'avg-earn-percent' ||
    value === 'trade-count'
  )
}

function isGoalScheduleType(value: string): value is GoalScheduleType {
  return value === 'frequency' || value === 'deadline'
}

function mapGoalRowToRecord(row: GoalRow): GoalRecord | null {
  if (!isGoalMetric(row.metric) || !isGoalScheduleType(row.schedule_type)) {
    return null
  }

  const targetValue = Number(row.target_value)
  if (!Number.isFinite(targetValue)) {
    return null
  }

  const targetCurrency = resolveCurrencyCode(row.target_currency)
  const frequency =
    row.frequency === 'day' || row.frequency === 'week' || row.frequency === 'month' || row.frequency === 'year'
      ? row.frequency
      : null
  const parsedFrequencyNumber = Number(row.frequency_number)
  const frequencyNumber =
    Number.isInteger(parsedFrequencyNumber) && parsedFrequencyNumber > 0 ? parsedFrequencyNumber : 1

  return {
    id: row.id,
    metric: row.metric,
    targetValue,
    targetCurrency,
    scheduleType: row.schedule_type,
    frequency,
    frequencyNumber,
    deadline: typeof row.deadline === 'string' ? row.deadline : null,
    createdAt: row.created_at,
  }
}

function getGoalMetricOption(metric: GoalMetric): GoalMetricOption {
  return GOAL_METRIC_OPTIONS.find((option) => option.value === metric) ?? GOAL_METRIC_OPTIONS[0]
}

function getGoalMetricUnit(metric: GoalMetric): GoalMetricOption['unit'] {
  return getGoalMetricOption(metric).unit
}

function getGoalMetricLabel(metric: GoalMetric): string {
  return getGoalMetricOption(metric).label
}

function formatGoalValue(value: number, metric: GoalMetric, currency: CurrencyCode): string {
  const unit = getGoalMetricUnit(metric)

  if (unit === 'currency') {
    return formatCurrency(value, currency)
  }

  if (unit === 'percent') {
    return formatPercent(value)
  }

  return Number.isFinite(value) ? `${Math.max(0, value).toFixed(0)} trades` : '-'
}

function clampPercent(value: number): number {
  if (!Number.isFinite(value)) {
    return 0
  }

  return Math.min(100, Math.max(0, value))
}

function parseGoalDeadlineLabel(deadline: string): string {
  if (!deadline.trim()) {
    return 'No deadline'
  }

  const parsed = new Date(deadline)
  return Number.isNaN(parsed.getTime()) ? deadline : formatDateDDMMYYYY(parsed)
}

function getGoalScheduleLabel(goal: GoalRecord): string {
  if (goal.scheduleType === 'deadline') {
    return goal.deadline ? parseGoalDeadlineLabel(goal.deadline) : 'No deadline'
  }

  if (!goal.frequency) {
    return '1 Month'
  }

  const frequencyLabel =
    GOAL_FREQUENCY_OPTIONS.find((option) => option.value === goal.frequency)?.label ?? goal.frequency
  const frequencyNumber = Math.max(1, Math.trunc(goal.frequencyNumber || 1))

  if (frequencyNumber === 1) {
    return `1 ${frequencyLabel}`
  }

  return `${frequencyNumber} ${frequencyLabel}s`
}

function getGoalDisplayTitle(goal: GoalRecord): string {
  return `${getGoalMetricLabel(goal.metric)}`
}

function getPeriodStartEnd(frequency: GoalFrequency, frequencyNumber: number): { start: Date; end: Date } {
  const normalizedFrequencyNumber = Math.max(1, Math.trunc(frequencyNumber || 1))
  const now = new Date()
  const start = new Date(now)
  start.setHours(0, 0, 0, 0)

  if (frequency === 'day') {
    start.setDate(start.getDate() - (normalizedFrequencyNumber - 1))
    const end = new Date(start)
    end.setDate(end.getDate() + normalizedFrequencyNumber)
    return { start, end }
  }

  if (frequency === 'week') {
    const dayOfWeek = start.getDay()
    const daysSinceMonday = (dayOfWeek + 6) % 7
    start.setDate(start.getDate() - daysSinceMonday - (normalizedFrequencyNumber - 1) * 7)
    const end = new Date(start)
    end.setDate(end.getDate() + normalizedFrequencyNumber * 7)
    return { start, end }
  }

  if (frequency === 'month') {
    start.setDate(1)
    start.setMonth(start.getMonth() - (normalizedFrequencyNumber - 1))
    const end = new Date(start)
    end.setMonth(end.getMonth() + normalizedFrequencyNumber)
    return { start, end }
  }

  if (frequency === 'year') {
    start.setMonth(0, 1)
    start.setFullYear(start.getFullYear() - (normalizedFrequencyNumber - 1))
    const end = new Date(start)
    end.setFullYear(end.getFullYear() + normalizedFrequencyNumber)
    return { start, end }
  }

  return {
    start,
    end: now,
  }
}

function isWithinGoalPeriod(date: Date, frequency: GoalFrequency, frequencyNumber: number): boolean {
  const range = getPeriodStartEnd(frequency, frequencyNumber)

  const timestamp = date.getTime()
  return timestamp >= range.start.getTime() && timestamp < range.end.getTime()
}

function getGoalDefaultFormValues(defaultCurrency: CurrencyCode = 'USD'): Record<string, string> {
  return GOAL_FORM_CONFIG.reduce<Record<string, string>>((acc, field) => {
    if (field.key === 'targetCurrency') {
      acc[field.key] = defaultCurrency
    } else {
      acc[field.key] = field.defaultValue ?? ''
    }
    return acc
  }, {})
}

function getGoalValueInTargetCurrency(
  goal: GoalRecord,
  currentValue: number,
  selectedCurrency: CurrencyCode,
  currencyRates: Record<CurrencyCode, number>,
): number {
  if (getGoalMetricUnit(goal.metric) !== 'currency') {
    return currentValue
  }

  return convertAmount(currentValue, selectedCurrency, goal.targetCurrency, currencyRates)
}

function getGoalMetricValue(goal: GoalRecord, snapshot: GoalMetricSnapshot, periodSnapshot: GoalPeriodMetrics): number {
  if (goal.metric === 'total-money') {
    return snapshot.totalMoney
  }

  if (goal.metric === 'earn-money') {
    return goal.scheduleType === 'deadline' ? snapshot.earnMoney : periodSnapshot.earnMoney
  }

  if (goal.metric === 'earn-percent') {
    return goal.scheduleType === 'deadline' ? snapshot.earnPercent : periodSnapshot.earnPercent
  }

  if (goal.metric === 'avg-earn-percent') {
    return goal.scheduleType === 'deadline' ? snapshot.avgEarnPercent : periodSnapshot.avgEarnPercent
  }

  return goal.scheduleType === 'deadline' ? snapshot.tradeCount : periodSnapshot.tradeCount
}

function getGoalCompletion(
  goal: GoalRecord,
  snapshot: GoalMetricSnapshot,
  periodSnapshot: GoalPeriodMetrics,
  selectedCurrency: CurrencyCode,
  currencyRates: Record<CurrencyCode, number>,
): number {
  const currentValue = getGoalValueInTargetCurrency(
    goal,
    getGoalMetricValue(goal, snapshot, periodSnapshot),
    selectedCurrency,
    currencyRates,
  )

  if (!Number.isFinite(goal.targetValue) || goal.targetValue <= 0) {
    return 0
  }

  return clampPercent((currentValue / goal.targetValue) * 100)
}

function getGoalValueSummary(goal: GoalRecord, currentValue: number, targetValue: number, currency: CurrencyCode): string {
  const unit = getGoalMetricUnit(goal.metric)

  if (unit === 'currency') {
    return `${formatGoalValue(currentValue, goal.metric, currency)} / ${formatGoalValue(targetValue, goal.metric, currency)}`
    // return `${formatGoalValue(targetValue, goal.metric, currency)}`
  }

  if (unit === 'percent') {
    return `${formatGoalValue(currentValue, goal.metric, currency)} / ${formatGoalValue(targetValue, goal.metric, currency)}`
    // return `${formatGoalValue(targetValue, goal.metric, currency)}`
  }

  return `${Math.max(0, currentValue).toFixed(0)} / ${Math.max(0, targetValue).toFixed(0)} trades`
  // return `${Math.max(0, targetValue).toFixed(0)} trades`
}

function parseCurrencyToNumber(input: string): number {
  return parseFloat(input.replace(/[^\d.-]/g, '')) || 0;
}

function formatNumberAbbrev(num: number): string {
  if (num === 0) return '0';
  const absNum = Math.abs(num);
  const sign = num < 0 ? '-' : '';
  const suffixes = ['', 'K', 'M', 'B'];
  let i = 0;
  let scaled = absNum;
  while (scaled >= 1000 && i < 3) {
    scaled /= 1000;
    i++;
  }
  return sign + scaled.toFixed(1).replace(/\.0$/, '') + suffixes[i];
}

function formatCurrencyAbbrev(input: string): string {
  const num = parseCurrencyToNumber(input);
  const symbol = input.includes('HK$') ? 'HK$' : '$';
  return symbol + formatNumberAbbrev(num);
}


function getGoalValue(goal: GoalRecord, _currentValue: number, targetValue: number, currency: CurrencyCode): string {
  const unit = getGoalMetricUnit(goal.metric)

  if (unit === 'currency') {
    // return `${formatGoalValue(currentValue, goal.metric, currency)} / ${formatGoalValue(targetValue, goal.metric, currency)}`
    return `${formatCurrencyAbbrev(formatGoalValue(targetValue, goal.metric, currency))}`
  }

  if (unit === 'percent') {
    // return `${formatGoalValue(currentValue, goal.metric, currency)} / ${formatGoalValue(targetValue, goal.metric, currency)}`
    return `${formatGoalValue(targetValue, goal.metric, currency)}`
  }

  // return `${Math.max(0, currentValue).toFixed(0)} / ${Math.max(0, targetValue).toFixed(0)} trades`
  return `${Math.max(0, targetValue).toFixed(0)} trades`
}

function getCurrentMarketDayKey(): string {
  const marketDateParts = new Intl.DateTimeFormat('en-CA', {
    timeZone: US_MARKET_TIME_ZONE,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).formatToParts(new Date())

  const year = marketDateParts.find((part) => part.type === 'year')?.value ?? '0000'
  const month = marketDateParts.find((part) => part.type === 'month')?.value ?? '00'
  const day = marketDateParts.find((part) => part.type === 'day')?.value ?? '00'

  return `${year}-${month}-${day}`
}

function getCurrentUsMarketClock(): { weekday: string; minutesSinceMidnight: number } {
  const marketTimeParts = new Intl.DateTimeFormat('en-US', {
    timeZone: US_MARKET_TIME_ZONE,
    weekday: 'short',
    hour: '2-digit',
    minute: '2-digit',
    hour12: false,
  }).formatToParts(new Date())

  const weekday = marketTimeParts.find((part) => part.type === 'weekday')?.value ?? ''
  const hour = Number(marketTimeParts.find((part) => part.type === 'hour')?.value ?? '0')
  const minute = Number(marketTimeParts.find((part) => part.type === 'minute')?.value ?? '0')

  return {
    weekday,
    minutesSinceMidnight: hour * 60 + minute,
  }
}

function isUsTradingDay(weekday: string): boolean {
  return weekday !== 'Sat' && weekday !== 'Sun'
}

function safeReadStorage<T>(key: string, fallback: T): T {
  if (typeof window === 'undefined') {
    return fallback
  }

  try {
    const raw = window.localStorage.getItem(key)
    if (!raw) {
      return fallback
    }

    return JSON.parse(raw) as T
  } catch {
    return fallback
  }
}

function safeWriteStorage(key: string, value: unknown): void {
  if (typeof window === 'undefined') {
    return
  }

  try {
    window.localStorage.setItem(key, JSON.stringify(value))
  } catch {
    // Ignore storage write failures.
  }
}

function applyThemeMode(mode: ThemeMode): void {
  if (typeof document === 'undefined') {
    return
  }

  if (mode === 'system') {
    document.documentElement.removeAttribute('data-theme')
    return
  }

  document.documentElement.setAttribute('data-theme', mode)
}

function App() {
  const [stocksData, setStocksData] = useState<SheetDataWithIds>({ headers: [], rows: [], ids: [] })
  const [dividendData, setDividendData] = useState<SheetDataWithIds>({ headers: [], rows: [], ids: [] })
  const [moneyMoveData, setMoneyMoveData] = useState<SheetDataWithIds>({ headers: [], rows: [], ids: [] })
  const [isLoading, setIsLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)
  const [page, setPage] = useState<PageName>('dataShow')
  const [selectedCurrency, setSelectedCurrency] = useState<CurrencyCode>('USD')
  const [themeMode, setThemeMode] = useState<ThemeMode>('system')
  const [settingsAnchorEl, setSettingsAnchorEl] = useState<HTMLElement | null>(null)
  const [notificationsEnabled, setNotificationsEnabled] = useState(false)
  const [notificationPermission, setNotificationPermission] = useState<NotificationPermission>(
    typeof window !== 'undefined' && 'Notification' in window ? Notification.permission : 'default',
  )
  const [currencyRates, setCurrencyRates] = useState<Record<CurrencyCode, number>>(DEFAULT_RATES)
  const [isUsMarketOpen, setIsUsMarketOpen] = useState<boolean | null>(null)
  // const [isQuoteLoading, setIsQuoteLoading] = useState(false)
  // const [quoteUpdatedAt, setQuoteUpdatedAt] = useState<number | null>(null)
  const [selectedAddTable, setSelectedAddTable] = useState<AddTableName>('Stocks')
  const [addFormValues, setAddFormValues] = useState<Record<string, string>>(
    getDefaultFormValues('Stocks'),
  )
  const [isConfirmingAdd, setIsConfirmingAdd] = useState(false)
  const [isSubmittingAdd, setIsSubmittingAdd] = useState(false)
  const [addFormMessage, setAddFormMessage] = useState<{ type: 'success' | 'error'; text: string } | null>(null)
  const [isAuthReady, setIsAuthReady] = useState(false)
  const [currentUserId, setCurrentUserId] = useState<string | null>(null)
  const [currentUserName, setCurrentUserName] = useState<string>('')
  const [currentUserEmail, setCurrentUserEmail] = useState<string | null>(null) 
  const [isUploadingExcel, setIsUploadingExcel] = useState(false)
  const [excelUploadMessage, setExcelUploadMessage] = useState<{ type: 'success' | 'error'; text: string } | null>(null)
  const [selectedExcelFile, setSelectedExcelFile] = useState<File | null>(null)
  const [excelPreviewData, setExcelPreviewData] = useState<ExcelDataRow | null>(null)
  const [selectedExcelPreviewTable, setSelectedExcelPreviewTable] = useState<AddTableName>('Stocks')
  const [isParsingExcelPreview, setIsParsingExcelPreview] = useState(false)
  const [isConfirmingExcelAdd, setIsConfirmingExcelAdd] = useState(false)
  const [followSymbolInput, setFollowSymbolInput] = useState('')
  const [userFollowSymbols, setUserFollowSymbols] = useState<string[]>(['AAPL', 'TSLA'])
  const [liveQuotes, setLiveQuotes] = useState<Record<string, FinnhubQuoteResponse>>({})
  const [isLiveQuotesLoading, setIsLiveQuotesLoading] = useState(false)
  const [liveQuotesUpdatedAt, setLiveQuotesUpdatedAt] = useState<number | null>(null)
  const [isFastMode, setIsFastMode] = useState(false)
  const [alertSymbol, setAlertSymbol] = useState('AAPL')
  const [alertPriceInput, setAlertPriceInput] = useState('')
  const [alertCondition, setAlertCondition] = useState<PriceAlertCondition>('above')
  const [priceAlerts, setPriceAlerts] = useState<PriceAlert[]>([])
  const [editingTable, setEditingTable] = useState<AddTableName | null>(null)
  const [editingRowIndex, setEditingRowIndex] = useState<number | null>(null)
  const [editFormValues, setEditFormValues] = useState<Record<string, string>>({})
  const [isConfirmingEdit, setIsConfirmingEdit] = useState(false)
  const [isSubmittingEdit, setIsSubmittingEdit] = useState(false)
  const [editFormMessage, setEditFormMessage] = useState<{ type: 'success' | 'error'; text: string } | null>(null)
  const [isConfirmingDelete, setIsConfirmingDelete] = useState(false)
  const [goalRecords, setGoalRecords] = useState<GoalRecord[]>([])
  const [selectedGoalId, setSelectedGoalId] = useState<string | null>(null)
  const [goalFormValues, setGoalFormValues] = useState<Record<string, string>>(getGoalDefaultFormValues())
  const [isConfirmingGoal, setIsConfirmingGoal] = useState(false)
  const [isSubmittingGoal, setIsSubmittingGoal] = useState(false)
  const [goalMessage, setGoalMessage] = useState<{ type: 'success' | 'error'; text: string } | null>(null)
  const [isGoalModalOpen, setIsGoalModalOpen] = useState(false)
  const [viewingGoalIdMain, setViewingGoalIdMain] = useState<string | null>(null)
  const [showGoalTimelineAsPercent, setShowGoalTimelineAsPercent] = useState(true)
  const [isToolsModalOpen, setIsToolsModalOpen] = useState(false)
  const [calcBudgetInput, setCalcBudgetInput] = useState('')
  const [calcStockPriceInput, setCalcStockPriceInput] = useState('')
  const [calcQuantityInput, setCalcQuantityInput] = useState('')
  const [calcSelectedSymbol, setCalcSelectedSymbol] = useState<string | null>(null)
  const userSymbolsReadyForIdRef = useRef<string | null>(null)
  const notificationRegistrationRef = useRef<ServiceWorkerRegistration | null>(null)
  const excelFileInputRef = useRef<HTMLInputElement | null>(null)
  const [calculatorType, setCalculatorType] = useState<'buy' | 'sell'>('buy')

  const isConfigValid = Boolean(
    (import.meta.env.VITE_PUBLIC_SUPABASE_URL?.trim() || import.meta.env.VITE_SUPABASE_URL?.trim()) &&
      (import.meta.env.VITE_PUBLIC_SUPABASE_ANON_KEY?.trim() || import.meta.env.VITE_SUPABASE_ANON_KEY?.trim()),
  )

  const displayStocksData = useMemo(() => enrichStocksWithProfitLoss(stocksData), [stocksData])
  const selectedAddFields = useMemo(() => ADD_FORM_CONFIG[selectedAddTable], [selectedAddTable])
  const excelActionButtonSx = {
    height: 40,
    minWidth: 140,
    borderRadius: '12px',
    backgroundColor: 'var(--bg)',
    border: '0px solid',
    color: 'var(--text)',
    textTransform: 'none',
    '&.MuiButton-contained': {
      backgroundColor: 'var(--special)',
      color: 'var(--text)',
      '&:hover': {
        borderColor: 'var(--accent)',
        opacity: 0.92,
      },
    },
    '&.MuiButton-outlined:hover': {
      borderColor: 'var(--accent)',
      opacity: 0.92,
    },
  }

  const previewRows = useMemo(() => {
    if (!excelPreviewData) {
      return [] as string[][]
    }

    if (selectedExcelPreviewTable === 'Stocks') {
      return excelPreviewData.stocks
    }

    if (selectedExcelPreviewTable === 'Dividend') {
      return excelPreviewData.dividend
    }

    return excelPreviewData.moneyMove
  }, [excelPreviewData, selectedExcelPreviewTable])

  const formatExcelPreviewCellValue = (value: string, headerName: string): string => {
    const headerLower = headerName.toLowerCase()
    const isTimeColumn = headerLower.includes('time')

    if (!isTimeColumn || !value.trim()) {
      return value
    }

    const asNumber = Number(value.trim())
    if (!Number.isNaN(asNumber) && asNumber > 0) {
      try {
        const excelBaseDate = new Date(1899, 11, 30)
        const jsDate = new Date(excelBaseDate.getTime() + asNumber * 24 * 60 * 60 * 1000)
        return jsDate.toLocaleString()
      } catch (err) {
        console.warn('Failed to format Excel serial:', value, err)
        return value
      }
    }

    try {
      const parsed = new Date(value.trim())
      if (!Number.isNaN(parsed.getTime())) {
        return parsed.toLocaleString()
      }
    } catch (err) {
      console.warn('Failed to format date string:', value, err)
    }

    return value
  }

  useEffect(() => {
    setAddFormValues(getDefaultFormValues(selectedAddTable))
    setIsConfirmingAdd(false)
    setAddFormMessage(null)
  }, [selectedAddTable])

  useEffect(() => {
    if (!currentUserId) {
      setGoalRecords([])
      setSelectedGoalId(null)
      setGoalMessage(null)
      setGoalFormValues(getGoalDefaultFormValues(selectedCurrency))
      setIsConfirmingGoal(false)
      return
    }

    let isActive = true

    void (async () => {
      try {
        const storedGoals = (await loadUserGoals(currentUserId))
          .map(mapGoalRowToRecord)
          .filter((goal): goal is GoalRecord => goal !== null)

        if (!isActive) {
          return
        }

        setGoalRecords(storedGoals)
        setSelectedGoalId((currentSelectedId) =>
          storedGoals.some((goal) => goal.id === currentSelectedId) ? currentSelectedId : storedGoals[0]?.id ?? null,
        )
        setGoalMessage(null)
        setGoalFormValues(getGoalDefaultFormValues(selectedCurrency))
        setIsConfirmingGoal(false)
      } catch (error) {
        if (!isActive) {
          return
        }

        const message = error instanceof Error ? error.message : 'Unknown error'
        setGoalRecords([])
        setSelectedGoalId(null)
        setGoalMessage({ type: 'error', text: `Unable to load goals. ${message}` })
      }
    })()

    return () => {
      isActive = false
    }
  }, [currentUserId])

  useEffect(() => {
    const storedTheme = safeReadStorage<ThemeMode | null>(THEME_STORAGE_KEY, null)
    const preferredTheme = storedTheme ?? 'system'
    setThemeMode(preferredTheme)
    applyThemeMode(preferredTheme)

    const storedNotificationsEnabled = safeReadStorage<boolean>(NOTIFICATIONS_STORAGE_KEY, false)
    setNotificationsEnabled(Boolean(storedNotificationsEnabled))
    setUserFollowSymbols(['AAPL', 'TSLA'])
    setAlertSymbol('AAPL')

    const storedAlerts = safeReadStorage<PriceAlert[]>(PRICE_ALERTS_STORAGE_KEY, [])
    if (Array.isArray(storedAlerts)) {
      setPriceAlerts(
        storedAlerts.filter(
          (alert) =>
            typeof alert.id === 'string' &&
            typeof alert.symbol === 'string' &&
            typeof alert.targetPrice === 'number' &&
            Number.isFinite(alert.targetPrice) &&
            (alert.condition === 'above' || alert.condition === 'below'),
        ),
      )
    }
  }, [])

  useEffect(() => {
    if (!currentUserId) {
      userSymbolsReadyForIdRef.current = null
      setUserFollowSymbols(['AAPL', 'TSLA'])
      return
    }

    const cookieRaw = getCookie(LIVE_SYMBOLS_COOKIE_KEY)
    if (!cookieRaw) {
      const legacySymbols = sanitizeSymbols(safeReadStorage<string[]>(LIVE_SYMBOLS_STORAGE_KEY, []))
      setUserFollowSymbols(legacySymbols.length > 0 ? legacySymbols : ['AAPL', 'TSLA'])
      userSymbolsReadyForIdRef.current = currentUserId
      return
    }

    try {
      const payload = JSON.parse(cookieRaw) as LiveSymbolsCookiePayload

      if (payload.user !== currentUserId) {
        setUserFollowSymbols(['AAPL', 'TSLA'])
        userSymbolsReadyForIdRef.current = currentUserId
        return
      }

      const nextSymbols = sanitizeSymbols(payload.followList)
      setUserFollowSymbols(nextSymbols.length > 0 ? nextSymbols : ['AAPL', 'TSLA'])
      userSymbolsReadyForIdRef.current = currentUserId
    } catch {
      setUserFollowSymbols(['AAPL', 'TSLA'])
      userSymbolsReadyForIdRef.current = currentUserId
    }
  }, [currentUserId])

  useEffect(() => {
    if (!currentUserId || userSymbolsReadyForIdRef.current !== currentUserId) {
      return
    }

    const payload: LiveSymbolsCookiePayload = {
      user: currentUserId,
      followList: sanitizeSymbols(userFollowSymbols),
    }

    setCookie(LIVE_SYMBOLS_COOKIE_KEY, JSON.stringify(payload), COOKIE_TTL_MS / 1000)
    safeWriteStorage(LIVE_SYMBOLS_STORAGE_KEY, payload.followList)
  }, [currentUserId, userFollowSymbols])

  const isSettingsOpen = Boolean(settingsAnchorEl)

  const getNotificationRegistration = async (): Promise<ServiceWorkerRegistration | null> => {
    if (typeof window === 'undefined' || !('serviceWorker' in navigator)) {
      return null
    }

    if (notificationRegistrationRef.current) {
      return notificationRegistrationRef.current
    }

    try {
      const existing = await navigator.serviceWorker.getRegistration(NOTIFICATION_SW_PATH)
      if (existing) {
        notificationRegistrationRef.current = existing
        return existing
      }

      const next = await navigator.serviceWorker.register(NOTIFICATION_SW_PATH)
      notificationRegistrationRef.current = next
      return next
    } catch {
      return null
    }
  }

  const handleOpenSettings = (event: React.MouseEvent<HTMLButtonElement>) => {
    setSettingsAnchorEl(event.currentTarget)
  }

  const handleCloseSettings = () => {
    setSettingsAnchorEl(null)
  }

  const handleToggleThemeMode = (
    _event: React.MouseEvent<HTMLElement>,
    nextMode: ThemeMode | null,
  ) => {
    if (!nextMode) {
      return
    }

    setThemeMode(nextMode)
    safeWriteStorage(THEME_STORAGE_KEY, nextMode)
    applyThemeMode(nextMode)
  }

  const sendBrowserNotification = (title: string, body: string) => {
    if (
      !notificationsEnabled ||
      notificationPermission !== 'granted' ||
      typeof window === 'undefined' ||
      !('Notification' in window)
    ) {
      return
    }

    void (async () => {
      const registration = await getNotificationRegistration()

      if (!registration) {
        new Notification(title, { body })
        return
      }

      await registration.showNotification(title, {
        body,
        tag: 'stock-alert',
        icon: '/vite.svg',
        badge: '/vite.svg',
        requireInteraction: true,
      })
    })()
  }

  const handleToggleNotifications = async (_event: React.ChangeEvent<HTMLInputElement>, checked: boolean) => {
    if (typeof window === 'undefined' || !('Notification' in window)) {
      setNotificationsEnabled(false)
      safeWriteStorage(NOTIFICATIONS_STORAGE_KEY, false)
      return
    }

    if (!checked) {
      setNotificationsEnabled(false)
      safeWriteStorage(NOTIFICATIONS_STORAGE_KEY, false)
      return
    }

    let permission = Notification.permission
    if (permission === 'default') {
      permission = await Notification.requestPermission()
    }

    setNotificationPermission(permission)
    const enabled = permission === 'granted'

    if (enabled) {
      await getNotificationRegistration()
    }

    setNotificationsEnabled(enabled)
    safeWriteStorage(NOTIFICATIONS_STORAGE_KEY, enabled)
  }

  const handleAddFollowSymbol = () => {
    const symbol = normalizeSymbolInput(followSymbolInput)

    if (!symbol) {
      return
    }

    setUserFollowSymbols((prev) => {
      if (prev.includes(symbol)) {
        return prev
      }

      return [symbol, ...prev]
    })

    setFollowSymbolInput('')
  }

  const handleUnfollowSymbol = (symbol: string) => {
    if (requiredHoldingSymbolSet.has(symbol)) {
      return
    }

    setUserFollowSymbols((prev) => {
      return prev.filter((item) => item !== symbol)
    })
  }

  const handleAddPriceAlert = () => {
    const symbol = alertSymbol.trim().toUpperCase()
    const targetPrice = Number(alertPriceInput)

    if (!symbol || !Number.isFinite(targetPrice) || targetPrice <= 0) {
      return
    }

    const nextAlert: PriceAlert = {
      id: `${symbol}-${Date.now()}`,
      symbol,
      targetPrice,
      condition: alertCondition,
      enabled: true,
      triggered: false,
    }

    setPriceAlerts((prev) => {
      const next = [nextAlert, ...prev]
      safeWriteStorage(PRICE_ALERTS_STORAGE_KEY, next)
      return next
    })
    setAlertPriceInput('')
  }

  const handleTogglePriceAlert = (alertId: string) => {
    setPriceAlerts((prev) => {
      const next = prev.map((alert) =>
        alert.id === alertId
          ? { ...alert, enabled: !alert.enabled, triggered: false }
          : alert,
      )
      safeWriteStorage(PRICE_ALERTS_STORAGE_KEY, next)
      return next
    })
  }

  const handleDeletePriceAlert = (alertId: string) => {
    setPriceAlerts((prev) => {
      const next = prev.filter((alert) => alert.id !== alertId)
      safeWriteStorage(PRICE_ALERTS_STORAGE_KEY, next)
      return next
    })
  }

  const handleGoalFieldChange = (fieldKey: string, value: string) => {
    setIsConfirmingGoal(false)
    setGoalMessage(null)
    setGoalFormValues((prev) => {
      const next = { ...prev, [fieldKey]: value }

      if (fieldKey === 'metric') {
        const nextMetric = value as GoalMetric
        if (getGoalMetricUnit(nextMetric) === 'currency' && !next.targetCurrency) {
          next.targetCurrency = selectedCurrency
        }
      }

      if (fieldKey === 'scheduleType') {
        if (value === 'frequency') {
          next.deadline = ''
          if (!next.frequency) {
            next.frequency = 'month'
          }
          if (!next.frequencyNumber) {
            next.frequencyNumber = '1'
          }
        } else {
          next.frequency = ''
          next.frequencyNumber = '1'
          next.deadline = next.deadline || ''
        }
      }

      return next
    })
  }

  const handleSelectGoal = (goalId: string) => {
    setSelectedGoalId(goalId)
  }

  const handleDeleteGoal = async (goalId: string) => {
    if (!currentUserId) {
      setGoalMessage({ type: 'error', text: 'Please sign in with Google before deleting goals.' })
      return
    }

    try {
      await deleteUserGoal(goalId)
      setGoalRecords((prev) => {
        const next = prev.filter((goal) => goal.id !== goalId)
        if (selectedGoalId === goalId) {
          setSelectedGoalId(next[0]?.id ?? null)
        }
        return next
      })
      setGoalMessage({ type: 'success', text: 'Goal removed.' })
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error'
      setGoalMessage({ type: 'error', text: `Failed to delete goal. ${message}` })
    }
  }

  const handleSubmitGoal = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    setGoalMessage(null)

    if (!isConfirmingGoal) {
      setIsConfirmingGoal(true)
      return
    }

    if (!currentUserId) {
      setIsConfirmingGoal(false)
      setGoalMessage({ type: 'error', text: 'Please sign in with Google before saving goals.' })
      return
    }

    const targetValue = Number(goalFormValues.targetValue)
    const metric = goalFormValues.metric as GoalMetric
    const targetCurrency = resolveCurrencyCode(goalFormValues.targetCurrency)
    const scheduleType = goalFormValues.scheduleType as GoalScheduleType
    const frequency = goalFormValues.frequency as GoalFrequency
    const frequencyNumber = Math.max(1, Math.trunc(Number(goalFormValues.frequencyNumber)))
    const deadline = goalFormValues.deadline.trim()

    if (!Number.isFinite(targetValue) || targetValue <= 0) {
      setIsConfirmingGoal(false)
      setGoalMessage({ type: 'error', text: 'Target value must be greater than 0.' })
      return
    }

    if (getGoalMetricUnit(metric) === 'currency' && !goalFormValues.targetCurrency.trim()) {
      setIsConfirmingGoal(false)
      setGoalMessage({ type: 'error', text: 'Target currency is required for price-based goals.' })
      return
    }

    if (scheduleType === 'frequency' && !frequency) {
      setIsConfirmingGoal(false)
      setGoalMessage({ type: 'error', text: 'Choose a frequency or switch to deadline.' })
      return
    }

    if (scheduleType === 'frequency' && (!Number.isInteger(frequencyNumber) || frequencyNumber <= 0)) {
      setIsConfirmingGoal(false)
      setGoalMessage({ type: 'error', text: 'Frequency number must be a positive whole number.' })
      return
    }

    if (scheduleType === 'deadline' && !deadline) {
      setIsConfirmingGoal(false)
      setGoalMessage({ type: 'error', text: 'Deadline is required when deadline mode is selected.' })
      return
    }

    setIsSubmittingGoal(true)

    try {
      const goalId = crypto.randomUUID()
      const newGoal: GoalRecord = {
        id: goalId,
        metric,
        targetValue,
        targetCurrency,
        scheduleType,
        frequency: scheduleType === 'frequency' ? frequency : null,
        frequencyNumber: scheduleType === 'frequency' ? frequencyNumber : 1,
        deadline: scheduleType === 'deadline' ? deadline : null,
        createdAt: new Date().toISOString(),
      }

      await insertUserGoal(currentUserId, {
        id: goalId,
        metric,
        targetValue,
        targetCurrency,
        scheduleType,
        frequency: newGoal.frequency,
        frequencyNumber: newGoal.frequencyNumber,
        deadline: newGoal.deadline,
      })

      setGoalRecords((prev) => {
        const next = [newGoal, ...prev]
        return next
      })
      setSelectedGoalId(newGoal.id)
      setGoalFormValues(getGoalDefaultFormValues(selectedCurrency))
      setIsConfirmingGoal(false)
      setGoalMessage({ type: 'success', text: 'Goal saved.' })
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unknown error'
      setGoalMessage({ type: 'error', text: `Unable to save goal. ${message}` })
    } finally {
      setIsSubmittingGoal(false)
      setIsConfirmingGoal(false)
    }
  }

  const handleOpenGoalModalMain = () => {
    setGoalFormValues(getGoalDefaultFormValues(selectedCurrency))
    setIsConfirmingGoal(false)
    setGoalMessage(null)
    setIsGoalModalOpen(true)
  }

  const handleCloseGoalModalMain = () => {
    setIsGoalModalOpen(false)
    setViewingGoalIdMain(null)
  }

  const handleOpenEditModal = (table: AddTableName, rowIndex: number) => {
    setEditingTable(table)
    setEditingRowIndex(rowIndex)
    setEditFormMessage(null)

    let data: SheetDataWithIds
    if (table === 'Stocks') {
      data = stocksData
    } else if (table === 'Dividend') {
      data = dividendData
    } else {
      data = moneyMoveData
    }

    const row = data.rows[rowIndex]
    if (!row) {
      return
    }

    const formValues: Record<string, string> = {}
    data.headers.forEach((header, colIndex) => {
      const value = row[colIndex] ?? ''
      const key = header
        .toLowerCase()
        .replace(/\s+/g, ' ')
        .split(' ')
        .map((word, i) => (i === 0 ? word : word.charAt(0).toUpperCase() + word.slice(1)))
        .join('')

      if (key === 'time' || key === 'date') {
        const parsed = parseDateValue(value)
        if (parsed) {
          formValues['time'] = parsed.toISOString().slice(0, 16)
        }
      } else if (key === 'action' && table === 'Money Move') {
        formValues['do'] = value
      } else {
        formValues[key] = value
      }
    })

    setEditFormValues(formValues)
    setIsConfirmingEdit(false)
    setIsConfirmingDelete(false)
  }

  const handleCloseEditModal = () => {
    setEditingTable(null)
    setEditingRowIndex(null)
    setEditFormValues({})
    setIsConfirmingEdit(false)
    setIsConfirmingDelete(false)
    setEditFormMessage(null)
  }

  const handleEditFieldChange = (fieldKey: string, value: string) => {
    setIsConfirmingEdit(false)
    setEditFormValues((prev) => ({ ...prev, [fieldKey]: value }))
  }

  const handleSubmitEditDelete = async (action: 'update' | 'delete') => {
    if (action === 'delete') {
      if (!isConfirmingDelete) {
        setIsConfirmingEdit(false)
        setIsConfirmingDelete(true)
        return
      }

      if (!editingTable || editingRowIndex === null || !currentUserId) {
        setEditFormMessage({ type: 'error', text: 'Invalid edit state' })
        return
      }

      let data: SheetDataWithIds
      if (editingTable === 'Stocks') {
        data = stocksData
      } else if (editingTable === 'Dividend') {
        data = dividendData
      } else {
        data = moneyMoveData
      }

      const recordId = data.ids[editingRowIndex]
      if (!recordId) {
        setEditFormMessage({ type: 'error', text: 'Record ID not found' })
        return
      }

      setIsSubmittingEdit(true)

      try {
        const { stocks, dividend, moneyMove } = await mutateUserRowAndReload(
          currentUserId,
          editingTable,
          recordId,
          'delete',
          editFormValues,
        )
        setStocksData(stocks)
        setDividendData(dividend)
        setMoneyMoveData(moneyMove)

        handleCloseEditModal()
        setEditFormMessage(null)
      } catch (err) {
        console.error('[Edit/Delete][Delete] failed', {
          table: editingTable,
          rowIndex: editingRowIndex,
          recordId,
          formValues: editFormValues,
          error: err,
        })
        const message = err instanceof Error ? err.message : 'Unknown error'
        setEditFormMessage({ type: 'error', text: `Failed to delete row. ${message}` })
      } finally {
        setIsSubmittingEdit(false)
        setIsConfirmingDelete(false)
      }
    } else {
      // Update action
      if (!isConfirmingEdit) {
        setIsConfirmingDelete(false)
        setIsConfirmingEdit(true)
        return
      }

      if (!editingTable || editingRowIndex === null || !currentUserId) {
        setEditFormMessage({ type: 'error', text: 'Invalid edit state' })
        return
      }

      let data: SheetDataWithIds
      if (editingTable === 'Stocks') {
        data = stocksData
      } else if (editingTable === 'Dividend') {
        data = dividendData
      } else {
        data = moneyMoveData
      }

      const recordId = data.ids[editingRowIndex]
      if (!recordId) {
        setEditFormMessage({ type: 'error', text: 'Record ID not found' })
        return
      }

      setIsSubmittingEdit(true)

      try {
        const { stocks, dividend, moneyMove } = await mutateUserRowAndReload(
          currentUserId,
          editingTable,
          recordId,
          'update',
          editFormValues,
        )
        setStocksData(stocks)
        setDividendData(dividend)
        setMoneyMoveData(moneyMove)

        handleCloseEditModal()
        setEditFormMessage(null)
      } catch (err) {
        console.error('[Edit/Delete][Update] failed', {
          table: editingTable,
          rowIndex: editingRowIndex,
          recordId,
          formValues: editFormValues,
          error: err,
        })
        const message = err instanceof Error ? err.message : 'Unknown error'
        setEditFormMessage({ type: 'error', text: `Failed to update row. ${message}` })
      } finally {
        setIsSubmittingEdit(false)
        setIsConfirmingEdit(false)
      }
    }
  }

  useEffect(() => {
    let isMounted = true

    const applySession = async (session: Session | null) => {
      if (!isMounted) {
        return
      }

      const user = session?.user ?? null
      if (!user) {
        setCurrentUserId(null)
        setCurrentUserName('')
        setCurrentUserEmail(null)
        setIsAuthReady(true)
        return
      }

      const fullName =
        (typeof user.user_metadata?.full_name === 'string' && user.user_metadata.full_name) ||
        (typeof user.user_metadata?.name === 'string' && user.user_metadata.name) ||
        user.email ||
        'User'

      setCurrentUserId(user.id)
      setCurrentUserName(fullName)
      setCurrentUserEmail(typeof user.email === 'string' ? user.email : null)
      setIsAuthReady(true)

      try {
        await ensureUserProfile(user.id, fullName, user.email ?? null)
      } catch (profileError) {
        const message =
          profileError instanceof Error ? profileError.message : 'Unable to sync user profile'
        setError(message)
      }
    }

    void supabase.auth.getSession().then(({ data }) => {
      void applySession(data.session)
    })

    const {
      data: { subscription },
    } = supabase.auth.onAuthStateChange((_event, session) => {
      void applySession(session)
    })

    return () => {
      isMounted = false
      subscription.unsubscribe()
    }
  }, [])

  useEffect(() => {
    if (typeof window === 'undefined' || !('Notification' in window)) {
      setNotificationPermission('denied')
      return
    }

    setNotificationPermission(Notification.permission)

    if ('serviceWorker' in navigator) {
      void getNotificationRegistration()
    }
  }, [])

  const handleGoogleSignIn = async () => {
    setError(null)
    const { error: authError } = await supabase.auth.signInWithOAuth({
      provider: 'google',
      options: {
        redirectTo: resolveOAuthRedirectUrl(),
      },
    })

    if (authError) {
      setError(`Google sign-in failed: ${authError.message}`)
    }
  }

  const handleSignOut = async () => {
    const { error: signOutError } = await supabase.auth.signOut()
    if (signOutError) {
      setError(`Sign out failed: ${signOutError.message}`)
      return
    }

    setCurrentUserId(null)
    setCurrentUserName('')
    setCurrentUserEmail(null)
    setStocksData({ headers: [], rows: [], ids: [] })
    setDividendData({ headers: [], rows: [], ids: [] })
    setMoneyMoveData({ headers: [], rows: [], ids: [] })
  }

  const handleAddFieldChange = (fieldKey: string, value: string) => {
    setIsConfirmingAdd(false)
    setAddFormValues((prev) => ({ ...prev, [fieldKey]: value }))
  }

  const handleCycleAddTable = () => {
    setSelectedAddTable((prev) => {
      const currentIndex = ADD_TABLE_OPTIONS.indexOf(prev)
      const nextIndex = (currentIndex + 1) % ADD_TABLE_OPTIONS.length
      return ADD_TABLE_OPTIONS[nextIndex]
    })
  }

  const renderAddField = (field: AddFieldConfig, className?: string) => {
    if (field.inputType === 'select') {
      return (
        <TextField
          key={field.key}
          className={className}
          select
          variant="standard"
          label={field.label}
          required={Boolean(field.required)}
          value={addFormValues[field.key] ?? ''}
          onChange={(event) => handleAddFieldChange(field.key, event.target.value)}
        >
          {(field.options ?? []).map((option) => (
            <MenuItem key={`${field.key}-${option}`} value={option}>
              {option}
            </MenuItem>
          ))}
        </TextField>
      )
    }

    if (field.inputType === 'datetime') {
      return (
        <TextField
          key={field.key}
          className={className}
          label={field.label}
          type="datetime-local"
          variant="standard"
          required={Boolean(field.required)}
          value={addFormValues[field.key] ?? ''}
          onChange={(event) => handleAddFieldChange(field.key, event.target.value)}
          InputLabelProps={{ shrink: true }}
        />
      )
    }

    return (
      <TextField
        key={field.key}
        className={className}
        label={field.label}
        type={field.inputType === 'number' ? 'number' : 'text'}
        variant="standard"
        required={Boolean(field.required)}
        value={addFormValues[field.key] ?? ''}
        onChange={(event) => handleAddFieldChange(field.key, event.target.value)}
        inputProps={field.inputType === 'number' && field.step ? { step: field.step } : undefined}
      />
    )
  }

  const renderEditField = (field: AddFieldConfig, className?: string) => {
    if (field.inputType === 'select') {
      return (
        <TextField
          key={field.key}
          className={className}
          select
          variant="standard"
          label={field.label}
          required={Boolean(field.required)}
          value={editFormValues[field.key] ?? ''}
          onChange={(event) => handleEditFieldChange(field.key, event.target.value)}
        >
          {(field.options ?? []).map((option) => (
            <MenuItem key={`${field.key}-${option}`} value={option}>
              {option}
            </MenuItem>
          ))}
        </TextField>
      )
    }

    if (field.inputType === 'datetime') {
      return (
        <TextField
          key={field.key}
          className={className}
          label={field.label}
          type="datetime-local"
          variant="standard"
          required={Boolean(field.required)}
          value={editFormValues[field.key] ?? ''}
          onChange={(event) => handleEditFieldChange(field.key, event.target.value)}
          InputLabelProps={{ shrink: true }}
        />
      )
    }

    return (
      <TextField
        key={field.key}
        className={className}
        label={field.label}
        type={field.inputType === 'number' ? 'number' : 'text'}
        variant="standard"
        required={Boolean(field.required)}
        value={editFormValues[field.key] ?? ''}
        onChange={(event) => handleEditFieldChange(field.key, event.target.value)}
        inputProps={field.inputType === 'number' && field.step ? { step: field.step } : undefined}
      />
    )
  }

  const renderGoalField = (field: GoalFieldConfig, className?: string) => {
    if (field.inputType === 'select') {
      return (
        <TextField
          key={field.key}
          className={className}
          select
          variant="standard"
          label={field.label}
          required={Boolean(field.required)}
          value={goalFormValues[field.key] ?? ''}
          onChange={(event) => handleGoalFieldChange(field.key, event.target.value)}
        >
          {(field.options ?? []).map((option) => {
            let label = option

            if (field.key === 'metric') {
              label = getGoalMetricLabel(option as GoalMetric)
            } else if (field.key === 'scheduleType') {
              label = option === 'frequency' ? 'Frequency' : 'Deadline'
            } else if (field.key === 'frequency') {
              label = GOAL_FREQUENCY_OPTIONS.find((item) => item.value === option)?.label ?? option
            }

            return (
              <MenuItem key={`${field.key}-${option}`} value={option}>
                {label}
              </MenuItem>
            )
          })}
        </TextField>
      )
    }

    if (field.inputType === 'date') {
      return (
        <TextField
          key={field.key}
          className={className}
          label={field.label}
          type="date"
          variant="standard"
          required={Boolean(field.required)}
          value={goalFormValues[field.key] ?? ''}
          onChange={(event) => handleGoalFieldChange(field.key, event.target.value)}
          InputLabelProps={{ shrink: true }}
        />
      )
    }

    return (
      <TextField
        key={field.key}
        className={className}
        label={field.label}
        type={field.inputType === 'number' ? 'number' : 'text'}
        variant="standard"
        required={Boolean(field.required)}
        value={goalFormValues[field.key] ?? ''}
        onChange={(event) => handleGoalFieldChange(field.key, event.target.value)}
        inputProps={field.inputType === 'number' && field.step ? { step: field.step } : undefined}
      />
    )
  }

  const renderGoalFields = () => {
    return GOAL_FORM_CONFIG.filter((field) => shouldShowGoalField(field)).map((field) => {
      if (field.key === 'frequency') {
        const frequencyNumberField = GOAL_FORM_CONFIG.find((item) => item.key === 'frequencyNumber')

        if (frequencyNumberField && shouldShowGoalField(frequencyNumberField)) {
          return (
            <Box className="goal-form-frequency-row" key="goal-frequency-row">
              {renderGoalField(field, 'goal-field-frequency')}
              {renderGoalField(frequencyNumberField, 'goal-field-frequencyNumber')}
            </Box>
          )
        }
      }

      if (field.key === 'frequencyNumber') {
        return null
      }

      return renderGoalField(field, `goal-field-${field.key}`)
    })
  }

  const shouldShowGoalField = (field: GoalFieldConfig): boolean => {
    const selectedMetric = goalFormValues.metric as GoalMetric
    const selectedScheduleType = goalFormValues.scheduleType as GoalScheduleType

    if (field.key === 'targetCurrency') {
      return getGoalMetricUnit(selectedMetric) === 'currency'
    }

    if (field.key === 'frequency') {
      return selectedScheduleType === 'frequency'
    }

    if (field.key === 'frequencyNumber') {
      return selectedScheduleType === 'frequency'
    }

    if (field.key === 'deadline') {
      return selectedScheduleType === 'deadline'
    }

    return true
  }

  const handleSubmitAdd = async (event: React.FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    setAddFormMessage(null)

    if (!isConfirmingAdd) {
      setIsConfirmingAdd(true)
      return
    }

    const missingRequiredLabel = selectedAddFields.find(
      (field) => field.required && !(addFormValues[field.key] ?? '').trim(),
    )?.label

    if (missingRequiredLabel) {
      setIsConfirmingAdd(false)
      setAddFormMessage({ type: 'error', text: `${missingRequiredLabel} is required.` })
      return
    }

    if (!currentUserId) {
      setIsConfirmingAdd(false)
      setAddFormMessage({
        type: 'error',
        text: 'Please sign in with Google before adding data.',
      })
      return
    }

    setIsSubmittingAdd(true)

    try {
      await insertUserRow(currentUserId, selectedAddTable, addFormValues)
      const { stocks, dividend, moneyMove } = await loadUserSheetData(currentUserId)
      setStocksData(stocks)
      setDividendData(dividend)
      setMoneyMoveData(moneyMove)

      setAddFormValues(getDefaultFormValues(selectedAddTable))
      setIsConfirmingAdd(false)
      setAddFormMessage({ type: 'success', text: `Added new row to ${selectedAddTable} in Supabase.` })
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Unknown error'
      setAddFormMessage({ type: 'error', text: `Unable to insert row. ${message}` })
    } finally {
      setIsConfirmingAdd(false)
      setIsSubmittingAdd(false)
    }
  }

  const handleDownloadExcel = () => {
    downloadExcelTemplate()
  }

  const handleExcelFileChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0] ?? null
    setSelectedExcelFile(file)
    setIsConfirmingExcelAdd(false)
    setExcelUploadMessage(null)

    if (!file) {
      setExcelPreviewData(null)
      return
    }

    setIsParsingExcelPreview(true)

    try {
      const parsedData = await parseExcelFile(file)
      setExcelPreviewData(parsedData)
      setSelectedExcelPreviewTable('Stocks')
    } catch (error) {
      setExcelPreviewData(null)
      const message = error instanceof Error ? error.message : 'Failed to parse Excel file'
      setExcelUploadMessage({ type: 'error', text: message })
    } finally {
      setIsParsingExcelPreview(false)
    }
  }

  const handleSubmitExcelUpload = async () => {
    if (!currentUserId) {
      setExcelUploadMessage({ type: 'error', text: 'Please sign in with Google before uploading data.' })
      return
    }

    if (!selectedExcelFile) {
      setExcelUploadMessage({ type: 'error', text: 'Please choose an Excel file first.' })
      return
    }

    if (!isConfirmingExcelAdd) {
      setIsConfirmingExcelAdd(true)
      return
    }

    setIsUploadingExcel(true)
    setExcelUploadMessage(null)

    try {
      const excelData = excelPreviewData ?? (await parseExcelFile(selectedExcelFile))
      let successCount = 0
      let errorCount = 0

      // Process Stocks
      for (const row of excelData.stocks) {
        try {
          const rowObj = excelRowToObject(row, [
            'Stock',
            'Currency',
            'Price',
            'Action',
            'Time',
            'Quantity',
            'Handling Fees',
          ])
          if (rowObj.stock?.trim()) {
            await insertUserRow(currentUserId, 'Stocks', rowObj)
            successCount++
          }
        } catch (err) {
          errorCount++
          console.error('Stock row error:', err)
        }
      }

      // Process Dividend
      for (const row of excelData.dividend) {
        try {
          const rowObj = excelRowToObject(row, ['Stock', 'Currency', 'Div', 'Time'])
          if (rowObj.stock?.trim()) {
            await insertUserRow(currentUserId, 'Dividend', rowObj)
            successCount++
          }
        } catch (err) {
          errorCount++
          console.error('Dividend row error:', err)
        }
      }

      // Process Money Move
      for (const row of excelData.moneyMove) {
        try {
          const rowObj = excelRowToObject(row, ['Name', 'Currency', 'Price', 'Time', 'Action'])
          if (rowObj.name?.trim()) {
            await insertUserRow(currentUserId, 'Money Move', rowObj)
            successCount++
          }
        } catch (err) {
          errorCount++
          console.error('Money Move row error:', err)
        }
      }

      // Reload data
      try {
        const newData = await loadUserSheetData(currentUserId)
        setStocksData(newData.stocks)
        setDividendData(newData.dividend)
        setMoneyMoveData(newData.moneyMove)
      } catch (reloadErr) {
        console.error('Reload error:', reloadErr)
      }

      const message = `Added ${successCount} rows successfully${errorCount > 0 ? ` (${errorCount} errors)` : ''}`
      setExcelUploadMessage({
        type: errorCount === 0 ? 'success' : 'error',
        text: message,
      })
      setSelectedExcelFile(null)
      setExcelPreviewData(null)
      setSelectedExcelPreviewTable('Stocks')
      if (excelFileInputRef.current) {
        excelFileInputRef.current.value = ''
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Failed to parse Excel file'
      setExcelUploadMessage({ type: 'error', text: message })
    } finally {
      setIsConfirmingExcelAdd(false)
      setIsUploadingExcel(false)
    }
  }

  const holdings = useMemo<HoldingSnapshot[]>(() => {
    const symbolIndex = findColumnIndex(stocksData.headers, [
      'symbol',
      'ticker',
      'code',
      'stock',
      'company',
    ])
    const nameIndex = findColumnIndex(stocksData.headers, ['name', 'company'])
    const actionIndex = findColumnIndex(stocksData.headers, [
      'action',
      'type',
      'side',
      'buysell',
      'operation',
    ])
    const quantityIndex = findColumnIndex(stocksData.headers, [
      'qty',
      'quantity',
      'share',
      'units',
      'vol',
      'volume',
    ])
    const priceIndex = findColumnIndex(stocksData.headers, [
      'price',
      'avgprice',
      'unitprice',
      'cost',
      'rate',
    ])
    const totalIndex = findColumnIndex(stocksData.headers, [
      'total',
      'amount',
      'value',
      'proceeds',
      'market value',
      'marketvalue',
    ])

    if (symbolIndex === -1 && nameIndex === -1) {
      // Need at least a symbol or name to identify the holding
      return []
    }

    if (quantityIndex === -1) {
      return []
    }

    const positionBySymbol = new Map<string, { name: string; quantity: number; totalCost: number }>()

    stocksData.rows.forEach((row) => {
      let symbol = symbolIndex >= 0 ? (row[symbolIndex] ?? '').trim().toUpperCase() : ''
      const name = nameIndex >= 0 ? (row[nameIndex] ?? '').trim() : ''

      if (!symbol && name) {
        // Fallback: use name as symbol if symbol is missing logic-wise
        symbol = name.toUpperCase().substring(0, 10).replace(/\s/g, '')
      }

      if (!symbol) {
        return
      }

      const quantitySigned = parseNumber(row[quantityIndex] ?? '')
      const quantity = Math.abs(quantitySigned)
      if (!Number.isFinite(quantity) || quantity <= 0) {
        return
      }

      const actionRaw = actionIndex >= 0 ? row[actionIndex] ?? '' : ''
      const action = resolveAction(actionRaw, quantitySigned)
      if (!action) {
        return
      }

      const price = priceIndex >= 0 ? parseNumber(row[priceIndex] ?? '') : Number.NaN
      const total = totalIndex >= 0 ? parseNumber(row[totalIndex] ?? '') : Number.NaN

      const cost = Number.isFinite(total)
        ? Math.abs(total)
        : Number.isFinite(price)
          ? Math.abs(price) * quantity
          : Number.NaN

      const existing = positionBySymbol.get(symbol) ?? { name: '', quantity: 0, totalCost: 0 }
      if (!existing.name && nameIndex >= 0) {
        existing.name = (row[nameIndex] ?? '').trim()
      }

      if (action === 'buy') {
        if (!Number.isFinite(cost)) {
          return
        }

        existing.quantity += quantity
        existing.totalCost += cost
        positionBySymbol.set(symbol, existing)
        return
      }

      if (existing.quantity <= 0) {
        positionBySymbol.set(symbol, existing)
        return
      }

      const sellQty = Math.min(quantity, existing.quantity)
      const avgCostBeforeSell = existing.quantity > 0 ? existing.totalCost / existing.quantity : 0
      existing.quantity -= sellQty
      existing.totalCost -= sellQty * avgCostBeforeSell

      if (existing.quantity < 0.000001) {
        existing.quantity = 0
        existing.totalCost = 0
      }

      positionBySymbol.set(symbol, existing)
    })

    return Array.from(positionBySymbol.entries())
      .filter(([, position]) => position.quantity > 0)
      .map(([symbol, position]) => ({
        symbol,
        displayName: position.name ? `${position.name} (${symbol})` : symbol,
        quantity: position.quantity,
        avgBuyPriceUsd: position.totalCost / position.quantity,
      }))
      .sort((a, b) => a.symbol.localeCompare(b.symbol))
  }, [stocksData])

  const requiredHoldingSymbols = useMemo(() => holdings.map((holding) => holding.symbol), [holdings])

  const requiredHoldingSymbolSet = useMemo(
    () => new Set(requiredHoldingSymbols),
    [requiredHoldingSymbols],
  )

  const liveSymbols = useMemo(() => {
    const merged = [...requiredHoldingSymbols, ...userFollowSymbols]
    return Array.from(new Set(merged))
  }, [requiredHoldingSymbols, userFollowSymbols])

  const followSymbolOptions = useMemo(() => {
    const symbols = [...requiredHoldingSymbols, ...userFollowSymbols, ...Object.keys(liveQuotes)]
    return Array.from(new Set(symbols)).sort((a, b) => a.localeCompare(b))
  }, [requiredHoldingSymbols, userFollowSymbols, liveQuotes])

  useEffect(() => {
    const loadCurrencyRates = async () => {
      const cachedRaw = getCookie(CURRENCY_COOKIE_KEY)

      if (cachedRaw) {
        try {
          const cached = JSON.parse(cachedRaw) as CurrencyRatesCache
          if (
            Number.isFinite(cached.timestamp) &&
            Date.now() - cached.timestamp < COOKIE_TTL_MS &&
            Number.isFinite(cached.rates.USD) &&
            Number.isFinite(cached.rates.HKD)
          ) {
            setCurrencyRates({
              USD: cached.rates.USD,
              HKD: cached.rates.HKD,
            })
            return
          }
        } catch {
          // Ignore invalid cache and continue with API request.
        }
      }

      try {
        const response = await fetch(CURRENCY_API_URL)
        if (!response.ok) {
          return
        }

        const payload = (await response.json()) as {
          data?: Record<string, { value?: number }>
        }

        const usdValue = payload.data?.USD?.value
        const hkdValue = payload.data?.HKD?.value

        if (
          typeof usdValue !== 'number' ||
          typeof hkdValue !== 'number' ||
          !Number.isFinite(usdValue) ||
          !Number.isFinite(hkdValue) ||
          usdValue <= 0 ||
          hkdValue <= 0
        ) {
          return
        }

        const nextRates: Record<CurrencyCode, number> = {
          USD: usdValue,
          HKD: hkdValue,
        }

        setCurrencyRates(nextRates)

        const cachePayload: CurrencyRatesCache = {
          timestamp: Date.now(),
          rates: nextRates,
        }

        setCookie(CURRENCY_COOKIE_KEY, JSON.stringify(cachePayload), COOKIE_TTL_MS / 1000)
      } catch {
        // Keep default rates when API fails.
      }
    }

    void loadCurrencyRates()
  }, [])

  useEffect(() => {
    if (liveSymbols.length === 0) {
      setAlertSymbol('')
      return
    }

    if (!liveSymbols.includes(alertSymbol)) {
      setAlertSymbol(liveSymbols[0])
    }
  }, [liveSymbols, alertSymbol])

  useEffect(() => {
    if (!FINNHUB_TOKEN || liveSymbols.length === 0) {
      setLiveQuotes({})
      setIsLiveQuotesLoading(false)
      return
    }

    let isCancelled = false

    const fetchLiveQuotes = async () => {
      setIsLiveQuotesLoading(true)

      try {
        const marketStatusResponse = await fetch(
          `${FINNHUB_MARKET_STATUS_URL}?exchange=US&token=${FINNHUB_TOKEN}`,
        )

        if (marketStatusResponse.ok) {
          const marketStatus = (await marketStatusResponse.json()) as FinnhubMarketStatusResponse
          if (!isCancelled) {
            setIsUsMarketOpen(marketStatus.isOpen === true)
          }
        } else if (!isCancelled) {
          setIsUsMarketOpen(null)
        }

        const quoteResults = await Promise.all(
          liveSymbols.map(async (symbol) => {
            const quoteResponse = await fetch(
              `${FINNHUB_QUOTE_URL}?symbol=${encodeURIComponent(symbol)}&token=${FINNHUB_TOKEN}`,
            )

            if (!quoteResponse.ok) {
              return [symbol, null] as const
            }

            const quote = (await quoteResponse.json()) as FinnhubQuoteResponse
            return [symbol, quote] as const
          }),
        )

        if (isCancelled) {
          return
        }

        const nextQuotes: Record<string, FinnhubQuoteResponse> = {}
        quoteResults.forEach(([symbol, quote]) => {
          if (quote) {
            nextQuotes[symbol] = quote
          }
        })

        setLiveQuotes((prev) => {
          const merged = { ...prev, ...nextQuotes }
          liveSymbols.forEach((symbol) => {
            if (!(symbol in merged)) {
              merged[symbol] = {}
            }
          })
          return merged
        })
        setLiveQuotesUpdatedAt(Date.now())
      } catch {
        if (!isCancelled) {
          setIsUsMarketOpen(null)
        }
      } finally {
        if (!isCancelled) {
          setIsLiveQuotesLoading(false)
        }
      }
    }

    const intervalMs =
      isUsMarketOpen === true
        ? isFastMode
          ? LIVE_QUOTES_FAST_INTERVAL_MS
          : LIVE_QUOTES_NORMAL_INTERVAL_MS
        : LIVE_QUOTES_NORMAL_INTERVAL_MS

    void fetchLiveQuotes()
    const timer = window.setInterval(fetchLiveQuotes, intervalMs)

    return () => {
      isCancelled = true
      window.clearInterval(timer)
    }
  }, [liveSymbols, isFastMode, isUsMarketOpen])

  useEffect(() => {
    if (!notificationsEnabled || !FINNHUB_TOKEN) {
      return
    }

    let isCancelled = false

    const checkMarketOpenNotification = () => {
      if (isCancelled) {
        return
      }

      const marketDayKey = getCurrentMarketDayKey()
      const notifiedDay = safeReadStorage<string | null>(MARKET_OPEN_NOTICE_DAY_KEY, null)
      const marketClock = getCurrentUsMarketClock()
      const isMarketSessionOpen =
        isUsTradingDay(marketClock.weekday) &&
        marketClock.minutesSinceMidnight >= US_MARKET_OPEN_MINUTES &&
        marketClock.minutesSinceMidnight < US_MARKET_CLOSE_MINUTES

      if (isMarketSessionOpen && notifiedDay !== marketDayKey) {
        sendBrowserNotification('US Market Open', 'US stock market is now open.')
        safeWriteStorage(MARKET_OPEN_NOTICE_DAY_KEY, marketDayKey)
      }
    }

    checkMarketOpenNotification()
    const timer = window.setInterval(checkMarketOpenNotification, MARKET_OPEN_CHECK_INTERVAL_MS)

    return () => {
      isCancelled = true
      window.clearInterval(timer)
    }
  }, [notificationsEnabled, notificationPermission])

  useEffect(() => {
    if (!notificationsEnabled || notificationPermission !== 'granted') {
      return
    }

    setPriceAlerts((prev) => {
      let didChange = false

      const next = prev.map((alert) => {
        if (!alert.enabled) {
          return alert
        }

        const currentPrice = liveQuotes[alert.symbol]?.c
        if (!Number.isFinite(currentPrice)) {
          return alert
        }

        const matchCondition =
          alert.condition === 'above'
            ? Number(currentPrice) >= alert.targetPrice
            : Number(currentPrice) <= alert.targetPrice

        if (matchCondition && !alert.triggered) {
          sendBrowserNotification(
            'Price Alert',
            `${alert.symbol} is ${alert.condition} ${alert.targetPrice.toFixed(2)}. Current: ${Number(currentPrice).toFixed(2)}`,
          )
          didChange = true
          return { ...alert, triggered: true }
        }

        if (!matchCondition && alert.triggered) {
          didChange = true
          return { ...alert, triggered: false }
        }

        return alert
      })

      if (didChange) {
        safeWriteStorage(PRICE_ALERTS_STORAGE_KEY, next)
      }

      return next
    })
  }, [liveQuotes, notificationsEnabled, notificationPermission])

  const summary = useMemo(() => {
    const moneyMoveAmountIndex = findColumnIndex(moneyMoveData.headers, [
      'amount',
      'total',
      'value',
      'money',
      'price',
    ])
    const moneyMoveActionIndex = findColumnIndex(moneyMoveData.headers, [
      'action',
      'type',
      'note',
      'description',
      'do',
    ])
    const moneyMoveCurrencyIndex = findColumnIndex(moneyMoveData.headers, ['currency'])

    let moneyMoveNet = 0
    let borrowed = 0

    moneyMoveData.rows.forEach((row) => {
      if (moneyMoveAmountIndex === -1 || moneyMoveActionIndex === -1) {
        return
      }

      const amount = parseNumber(row[moneyMoveAmountIndex] ?? '')
      if (!Number.isFinite(amount)) {
        return
      }

      const rowCurrency = resolveCurrencyCode(row[moneyMoveCurrencyIndex] ?? '')
      const amountInSelectedCurrency = convertAmount(
        Math.abs(amount),
        rowCurrency,
        selectedCurrency,
        currencyRates,
      )

      const action = resolveMoneyMoveType(row[moneyMoveActionIndex] ?? '')
      if (!action) {
        return
      }

      if (action === 'in' || action === 'bor') {
        moneyMoveNet += amountInSelectedCurrency
      } else if (action === 'out' || action === 'back') {
        moneyMoveNet -= amountInSelectedCurrency
      }

      if (action === 'bor') {
        borrowed += amountInSelectedCurrency
      } else if (action === 'back') {
        borrowed -= amountInSelectedCurrency
      }
    })

    const dividendAmountIndex = findColumnIndex(dividendData.headers, [
      'dividend',
      'div',
      'amount',
      'total',
      'value',
    ])
    const dividendCurrencyIndex = findColumnIndex(dividendData.headers, ['currency'])

    let dividendTotal = 0
    dividendData.rows.forEach((row) => {
      if (dividendAmountIndex === -1) {
        return
      }

      const amount = parseNumber(row[dividendAmountIndex] ?? '')
      if (Number.isFinite(amount)) {
        const rowCurrency = resolveCurrencyCode(row[dividendCurrencyIndex] ?? '')
        dividendTotal += convertAmount(amount, rowCurrency, selectedCurrency, currencyRates)
      }
    })

    const stocksCurrencyIndex = findColumnIndex(stocksData.headers, ['currency'])
    const stocksFeeIndex = findColumnIndex(stocksData.headers, [
      'handlingfee',
      'handling',
      'fee',
      'commission',
      'charge',
      'charges',
    ])
    const profitLossIndex = findColumnIndex(displayStocksData.headers, ['profitloss'])
    const percentIndex = findColumnIndex(displayStocksData.headers, ['%'])

    let earnAll = 0
    let earnTradeCount = 0
    let percentSum = 0
    let percentCount = 0
    let handlingFeeTotal = 0

    displayStocksData.rows.forEach((row, rowIndex) => {
      if (profitLossIndex !== -1) {
        const profitLoss = parseNumber(row[profitLossIndex] ?? '')
        if (Number.isFinite(profitLoss)) {
          const rowCurrency = resolveCurrencyCode(
            stocksData.rows[rowIndex]?.[stocksCurrencyIndex] ?? '',
          )
          earnAll += convertAmount(profitLoss, rowCurrency, selectedCurrency, currencyRates)
          earnTradeCount += 1
        }
      }

      if (stocksFeeIndex !== -1) {
        const feeValue = parseNumber(stocksData.rows[rowIndex]?.[stocksFeeIndex] ?? '')
        if (Number.isFinite(feeValue)) {
          const rowCurrency = resolveCurrencyCode(
            stocksData.rows[rowIndex]?.[stocksCurrencyIndex] ?? '',
          )
          handlingFeeTotal += convertAmount(
            Math.abs(feeValue),
            rowCurrency,
            selectedCurrency,
            currencyRates,
          )
        }
      }

      if (percentIndex !== -1) {
        const percentage = parseNumber(row[percentIndex] ?? '')
        if (Number.isFinite(percentage)) {
          percentSum += percentage
          percentCount += 1
        }
      }

    })

    const earnPerTrade = earnTradeCount > 0 ? earnAll / earnTradeCount : 0
    const percentAverage = percentCount > 0 ? percentSum / percentCount : 0
    const totalMoney = moneyMoveNet + earnAll + dividendTotal - handlingFeeTotal

    return {
      totalMoney,
      borrowed,
      dividendTotal,
      earnAll,
      handlingFeeTotal,
      earnPerTrade,
      percentAverage,
    }
  }, [moneyMoveData, dividendData, displayStocksData, stocksData, selectedCurrency, currencyRates])

  const goalSnapshot = useMemo(() => {
    const totalMoney = summary.totalMoney
    const earnMoney = summary.earnAll
    const earnPercent = totalMoney > 0 ? (earnMoney / totalMoney) * 100 : 0
    const avgEarnPercent = summary.percentAverage
    const tradeCount = stocksData.rows.length

    const fullPeriodMetrics: GoalPeriodMetrics = {
      earnMoney,
      earnPercent,
      avgEarnPercent,
      tradeCount,
    }

    const buildPeriodMetrics = (frequency: GoalFrequency, frequencyNumber: number): GoalPeriodMetrics => {
      const stocksDateIndex = findColumnIndex(stocksData.headers, ['date', 'time'])
      const stocksCurrencyIndex = findColumnIndex(stocksData.headers, ['currency'])
      const profitLossIndex = findColumnIndex(displayStocksData.headers, ['profitloss'])
      const percentIndex = findColumnIndex(displayStocksData.headers, ['%'])
      const dividendAmountIndex = findColumnIndex(dividendData.headers, ['dividend', 'div', 'amount', 'total', 'value'])
      const dividendDateIndex = findColumnIndex(dividendData.headers, ['date', 'time'])
      const dividendCurrencyIndex = findColumnIndex(dividendData.headers, ['currency'])

      let periodEarnMoney = 0
      let periodPercentSum = 0
      let periodPercentCount = 0
      let periodTradeCount = 0

      if (stocksDateIndex !== -1 && profitLossIndex !== -1) {
        displayStocksData.rows.forEach((row, rowIndex) => {
          const date = parseDateValue(stocksData.rows[rowIndex]?.[stocksDateIndex] ?? '')
          if (!date || !isWithinGoalPeriod(date, frequency, frequencyNumber)) {
            return
          }

          const profitLoss = parseNumber(row[profitLossIndex] ?? '')
          if (Number.isFinite(profitLoss)) {
            const rowCurrency = resolveCurrencyCode(stocksData.rows[rowIndex]?.[stocksCurrencyIndex] ?? '')
            periodEarnMoney += convertAmount(profitLoss, rowCurrency, selectedCurrency, currencyRates)
            periodTradeCount += 1
          }

          if (percentIndex !== -1) {
            const percentage = parseNumber(row[percentIndex] ?? '')
            if (Number.isFinite(percentage)) {
              periodPercentSum += percentage
              periodPercentCount += 1
            }
          }
        })
      }

      if (dividendAmountIndex !== -1 && dividendDateIndex !== -1) {
        dividendData.rows.forEach((row) => {
          const date = parseDateValue(row[dividendDateIndex] ?? '')
          if (!date || !isWithinGoalPeriod(date, frequency, frequencyNumber)) {
            return
          }

          const amount = parseNumber(row[dividendAmountIndex] ?? '')
          if (!Number.isFinite(amount)) {
            return
          }

          const rowCurrency = resolveCurrencyCode(row[dividendCurrencyIndex] ?? '')
          periodEarnMoney += convertAmount(amount, rowCurrency, selectedCurrency, currencyRates)
        })
      }

      return {
        earnMoney: periodEarnMoney,
        earnPercent: totalMoney > 0 ? (periodEarnMoney / totalMoney) * 100 : 0,
        avgEarnPercent: periodPercentCount > 0 ? periodPercentSum / periodPercentCount : 0,
        tradeCount: periodTradeCount,
      }
    }

    return {
      snapshot: {
        totalMoney,
        earnMoney,
        earnPercent,
        avgEarnPercent,
        tradeCount,
      },
      fullPeriodMetrics,
      buildPeriodMetrics,
    }
  }, [summary, stocksData, dividendData, displayStocksData, selectedCurrency, currencyRates])

  const selectedGoal = useMemo(() => {
    if (goalRecords.length === 0) {
      return null
    }

    return goalRecords.find((goal) => goal.id === selectedGoalId) ?? goalRecords[0]
  }, [goalRecords, selectedGoalId])

  const goalProgressRows = useMemo(() => {
    return goalRecords.map((goal) => {
      const periodSnapshot =
        goal.scheduleType === 'deadline' || !goal.frequency
          ? goalSnapshot.fullPeriodMetrics
          : goalSnapshot.buildPeriodMetrics(goal.frequency, goal.frequencyNumber)
      const currentValue = getGoalMetricValue(goal, goalSnapshot.snapshot, periodSnapshot)
      const displayCurrentValue = getGoalValueInTargetCurrency(goal, currentValue, selectedCurrency, currencyRates)
      const completion = getGoalCompletion(
        goal,
        goalSnapshot.snapshot,
        periodSnapshot,
        selectedCurrency,
        currencyRates,
      )

      return {
        ...goal,
        currentValue,
        displayCurrentValue,
        completion,
      }
    })
  }, [goalRecords, goalSnapshot, selectedCurrency, currencyRates])

  const selectedGoalProgress = useMemo(() => {
    if (!selectedGoal) {
      return null
    }

    const periodSnapshot =
      selectedGoal.scheduleType === 'deadline' || !selectedGoal.frequency
        ? goalSnapshot.fullPeriodMetrics
        : goalSnapshot.buildPeriodMetrics(selectedGoal.frequency, selectedGoal.frequencyNumber)
    const currentValue = getGoalMetricValue(selectedGoal, goalSnapshot.snapshot, periodSnapshot)
    const displayCurrentValue = getGoalValueInTargetCurrency(
      selectedGoal,
      currentValue,
      selectedCurrency,
      currencyRates,
    )
    const completion = getGoalCompletion(
      selectedGoal,
      goalSnapshot.snapshot,
      periodSnapshot,
      selectedCurrency,
      currencyRates,
    )

    return {
      goal: selectedGoal,
      currentValue,
      displayCurrentValue,
      completion,
      periodSnapshot,
    }
  }, [selectedGoal, goalSnapshot, selectedCurrency, currencyRates])

  const selectedMainGoalProgress = useMemo(() => {
    if (!viewingGoalIdMain) {
      return null
    }

    return goalProgressRows.find((goal) => goal.id === viewingGoalIdMain) ?? null
  }, [viewingGoalIdMain, goalProgressRows])

  const selectedMainGoalMonthlyCompletionHistory = useMemo(() => {
    if (!selectedMainGoalProgress) {
      return {
        labels: [] as string[],
        periodCompletionValues: [] as number[],
        timelineCompletionValues: [] as number[],
        timelineRawValues: [] as number[],
        metricUnit: '',
      }
    }

    const stocksDateIndex = findColumnIndex(stocksData.headers, ['date', 'time'])
    const stocksCurrencyIndex = findColumnIndex(stocksData.headers, ['currency'])
    const profitLossIndex = findColumnIndex(displayStocksData.headers, ['profitloss'])
    const percentIndex = findColumnIndex(displayStocksData.headers, ['%'])
    const dividendAmountIndex = findColumnIndex(dividendData.headers, ['dividend', 'div', 'amount', 'value'])
    const dividendDateIndex = findColumnIndex(dividendData.headers, ['date', 'time'])
    const dividendCurrencyIndex = findColumnIndex(dividendData.headers, ['currency'])

    const monthMap = new Map<string, { earnMoney: number; percentSum: number; percentCount: number; tradeCount: number }>()
    const deadlineWindowMonthKeys = (() => {
      if (selectedMainGoalProgress.scheduleType !== 'deadline' || !selectedMainGoalProgress.deadline) {
        return null
      }

      const deadlineDate = parseDateValue(selectedMainGoalProgress.deadline)
      if (!deadlineDate) {
        return null
      }

      const deadlineMonth = new Date(deadlineDate.getFullYear(), deadlineDate.getMonth(), 1)
      const startMonth = new Date(deadlineMonth)
      startMonth.setMonth(startMonth.getMonth() - 11)

      const allowedMonths = new Set<string>()
      const cursor = new Date(startMonth)

      while (cursor <= deadlineMonth) {
        allowedMonths.add(toMonthKey(cursor))
        cursor.setMonth(cursor.getMonth() + 1)
      }

      return allowedMonths
    })()

    const shouldUseDeadlineTotalMoneyTimeline =
      selectedMainGoalProgress.scheduleType === 'deadline' && selectedMainGoalProgress.metric === 'total-money'

    const deadlineTotalMoneyByMonth = shouldUseDeadlineTotalMoneyTimeline
      ? buildMonthlyTotalMoneySeries(
          moneyMoveData,
          stocksData,
          dividendData,
          selectedCurrency,
          currencyRates,
        )
      : null

    const deadlineMonths = deadlineWindowMonthKeys ? Array.from(deadlineWindowMonthKeys).sort((a, b) => a.localeCompare(b)) : []

    if (shouldUseDeadlineTotalMoneyTimeline && deadlineMonths.length > 0) {
      const allMonthlyKeysSorted = Array.from(deadlineTotalMoneyByMonth?.keys() ?? []).sort((a, b) => a.localeCompare(b))
      const firstWindowMonth = deadlineMonths[0]

      let lastKnownTotalMoney = 0
      allMonthlyKeysSorted.forEach((monthKey) => {
        if (monthKey < firstWindowMonth) {
          const totalMoney = deadlineTotalMoneyByMonth?.get(monthKey)
          if (typeof totalMoney === 'number') {
            lastKnownTotalMoney = totalMoney
          }
        }
      })

      const deadlineTimelineValues = deadlineMonths.map((monthKey) => {
        const monthTotalMoney = deadlineTotalMoneyByMonth?.get(monthKey)
        if (typeof monthTotalMoney === 'number') {
          lastKnownTotalMoney = monthTotalMoney
        }

        const normalizedTotalMoney = getGoalValueInTargetCurrency(
          selectedMainGoalProgress,
          lastKnownTotalMoney,
          selectedCurrency,
          currencyRates,
        )

        return clampPercent((normalizedTotalMoney / selectedMainGoalProgress.targetValue) * 100)
      })

      if (import.meta.env.DEV) {
        const latestTimelinePercent = deadlineTimelineValues[deadlineTimelineValues.length - 1] ?? 0
        const latestMonth = deadlineMonths[deadlineMonths.length - 1] ?? 'n/a'
        const latestMonthTotalMoney = deadlineTotalMoneyByMonth?.get(latestMonth) ?? lastKnownTotalMoney

        console.log('[GoalTimeline][DeadlineTotalMoney][Debug]', {
          goalId: selectedMainGoalProgress.id,
          metric: selectedMainGoalProgress.metric,
          targetValue: selectedMainGoalProgress.targetValue,
          targetCurrency: selectedMainGoalProgress.targetCurrency,
          selectedCurrency,
          gaugeCompletion: selectedMainGoalProgress.completion,
          gaugeCurrentValue: selectedMainGoalProgress.displayCurrentValue,
          summaryTotalMoney: summary.totalMoney,
          deadlineMonths,
          latestMonth,
          latestMonthTotalMoney,
          latestTimelinePercent,
          monthEndTotalsInWindow: deadlineMonths.map((monthKey, index) => ({
            monthKey,
            totalMoney: deadlineTotalMoneyByMonth?.get(monthKey) ?? null,
            timelinePercent: Number((deadlineTimelineValues[index] ?? 0).toFixed(2)),
          })),
        })
      }

      // Compute raw values for deadline total-money case
      let lastKnownTotalMoneyForRaw = 0
      // Initialize with the last known value before the window starts
      allMonthlyKeysSorted.forEach((monthKey) => {
        if (monthKey < firstWindowMonth) {
          const totalMoney = deadlineTotalMoneyByMonth?.get(monthKey)
          if (typeof totalMoney === 'number') {
            lastKnownTotalMoneyForRaw = totalMoney
          }
        }
      })

      const deadlineTimelineRawValues = deadlineMonths.map((monthKey) => {
        const monthTotalMoney = deadlineTotalMoneyByMonth?.get(monthKey)
        if (typeof monthTotalMoney === 'number') {
          lastKnownTotalMoneyForRaw = monthTotalMoney
        }
        return Number(lastKnownTotalMoneyForRaw.toFixed(2))
      })

      return {
        labels: deadlineMonths.map((month) => formatMonthLabel(month)),
        periodCompletionValues: deadlineTimelineValues.map((value) => Number(value.toFixed(2))),
        timelineCompletionValues: deadlineTimelineValues.map((value) => Number(value.toFixed(2))),
        timelineRawValues: deadlineTimelineRawValues,
        metricUnit: 'money',
      }
    }

    if (stocksDateIndex !== -1) {
      displayStocksData.rows.forEach((row, rowIndex) => {
        const dateValue = parseDateValue(stocksData.rows[rowIndex]?.[stocksDateIndex] ?? '')
        if (!dateValue) {
          return
        }

        const monthKey = toMonthKey(dateValue)
        if (deadlineWindowMonthKeys && !deadlineWindowMonthKeys.has(monthKey)) {
          return
        }

        const current = monthMap.get(monthKey) ?? { earnMoney: 0, percentSum: 0, percentCount: 0, tradeCount: 0 }

        if (profitLossIndex !== -1) {
          const profitLoss = parseNumber(row[profitLossIndex] ?? '')
          if (Number.isFinite(profitLoss)) {
            const rowCurrency = resolveCurrencyCode(stocksData.rows[rowIndex]?.[stocksCurrencyIndex] ?? '')
            current.earnMoney += convertAmount(profitLoss, rowCurrency, selectedCurrency, currencyRates)
            current.tradeCount += 1
          }
        }

        if (percentIndex !== -1) {
          const percentValue = parseNumber(row[percentIndex] ?? '')
          if (Number.isFinite(percentValue)) {
            current.percentSum += percentValue
            current.percentCount += 1
          }
        }

        monthMap.set(monthKey, current)
      })
    }

    if (dividendAmountIndex !== -1 && dividendDateIndex !== -1) {
      dividendData.rows.forEach((row) => {
        const dateValue = parseDateValue(row[dividendDateIndex] ?? '')
        if (!dateValue) {
          return
        }

        const monthKey = toMonthKey(dateValue)
        if (deadlineWindowMonthKeys && !deadlineWindowMonthKeys.has(monthKey)) {
          return
        }

        const current = monthMap.get(monthKey) ?? { earnMoney: 0, percentSum: 0, percentCount: 0, tradeCount: 0 }
        const amount = parseNumber(row[dividendAmountIndex] ?? '')
        if (!Number.isFinite(amount)) {
          return
        }

        const rowCurrency = resolveCurrencyCode(row[dividendCurrencyIndex] ?? '')
        current.earnMoney += convertAmount(amount, rowCurrency, selectedCurrency, currencyRates)
        monthMap.set(monthKey, current)
      })
    }

    const recentMonths = deadlineWindowMonthKeys
      ? Array.from(deadlineWindowMonthKeys)
      : Array.from(monthMap.keys()).sort((a, b) => a.localeCompare(b)).slice(-12)

    // Helper function to find the last available month with data on or before a given month
    const findLastValueAtOrBefore = (targetMonth: string) => {
      const row = monthMap.get(targetMonth)
      if (row) {
        return row
      }

      // Look backwards to find the last month with data
      const allMonthsSorted = Array.from(monthMap.keys()).sort((a, b) => a.localeCompare(b))
      for (let i = allMonthsSorted.length - 1; i >= 0; i--) {
        if (allMonthsSorted[i] <= targetMonth) {
          const foundRow = monthMap.get(allMonthsSorted[i])
          if (foundRow) {
            return foundRow
          }
        }
      }

      return null
    }

    const monthMetricValues = recentMonths.map((monthKey) => {
      const row = findLastValueAtOrBefore(monthKey)
      if (!row) {
        return 0
      }

      let currentValue = 0
      if (selectedMainGoalProgress.metric === 'earn-money') {
        currentValue = getGoalValueInTargetCurrency(
          selectedMainGoalProgress,
          row.earnMoney,
          selectedCurrency,
          currencyRates,
        )
      } else if (selectedMainGoalProgress.metric === 'earn-percent') {
        currentValue = summary.totalMoney > 0 ? (row.earnMoney / summary.totalMoney) * 100 : 0
      } else if (selectedMainGoalProgress.metric === 'avg-earn-percent') {
        currentValue = row.percentCount > 0 ? row.percentSum / row.percentCount : 0
      } else if (selectedMainGoalProgress.metric === 'trade-count') {
        currentValue = row.tradeCount
      } else {
        currentValue = getGoalValueInTargetCurrency(
          selectedMainGoalProgress,
          row.earnMoney,
          selectedCurrency,
          currencyRates,
        )
      }

      return currentValue
    })

    const hasValidTarget = Number.isFinite(selectedMainGoalProgress.targetValue) && selectedMainGoalProgress.targetValue > 0
    const periodCompletionValues = hasValidTarget
      ? monthMetricValues.map((value) => clampPercent((value / selectedMainGoalProgress.targetValue) * 100))
      : monthMetricValues.map(() => 0)

    let runningTotal = 0
    const timelineCompletionValues = hasValidTarget
      ? monthMetricValues.map((value) => {
          const unit = getGoalMetricUnit(selectedMainGoalProgress.metric)

          if (unit === 'percent') {
            runningTotal = value
          } else {
            runningTotal += value
          }

          return clampPercent((runningTotal / selectedMainGoalProgress.targetValue) * 100)
        })
      : monthMetricValues.map(() => 0)

    if (import.meta.env.DEV && selectedMainGoalProgress.scheduleType === 'deadline') {
      console.log('[GoalTimeline][DeadlineFallback][Debug]', {
        goalId: selectedMainGoalProgress.id,
        metric: selectedMainGoalProgress.metric,
        targetValue: selectedMainGoalProgress.targetValue,
        targetCurrency: selectedMainGoalProgress.targetCurrency,
        selectedCurrency,
        gaugeCompletion: selectedMainGoalProgress.completion,
        gaugeCurrentValue: selectedMainGoalProgress.displayCurrentValue,
        summaryTotalMoney: summary.totalMoney,
        recentMonths,
        monthMetricValues,
        timelineCompletionValues,
      })
    }

    // Compute raw timeline values (running total in actual units, not percentages)
    let runningTotalRaw = 0
    const timelineRawValues = monthMetricValues.map((value) => {
      const unit = getGoalMetricUnit(selectedMainGoalProgress.metric)

      if (unit === 'percent') {
        runningTotalRaw = value
      } else {
        runningTotalRaw += value
      }

      return runningTotalRaw
    })

    const metricUnit = getGoalMetricUnit(selectedMainGoalProgress.metric)

    return {
      labels: recentMonths.map((month) => formatMonthLabel(month)),
      periodCompletionValues: periodCompletionValues.map((value) => Number(value.toFixed(2))),
      timelineCompletionValues: timelineCompletionValues.map((value) => Number(value.toFixed(2))),
      timelineRawValues: timelineRawValues.map((value) => Number(value.toFixed(2))),
      metricUnit,
    }
  }, [selectedMainGoalProgress, stocksData, displayStocksData, dividendData, selectedCurrency, currencyRates, summary.totalMoney])

  const goalBarChartData = useMemo(() => {
    return {
      labels: goalProgressRows.map((goal) => getGoalDisplayTitle(goal)),
      values: goalProgressRows.map((goal) => goal.completion),
    }
  }, [goalProgressRows])

  const moneyTimelineChart = useMemo(() => {
    type TimelineEvent = {
      date: Date
      sourceOrder: number
      cashDelta: number
      stockAction?: { symbol: string; action: 'buy' | 'sell'; quantity: number; costPerShare: number }
    }

    const events: TimelineEvent[] = []

    const moneyMoveAmountIndex = findColumnIndex(moneyMoveData.headers, [
      'amount',
      'total',
      'value',
      'money',
      'price',
    ])
    const moneyMoveActionIndex = findColumnIndex(moneyMoveData.headers, [
      'action',
      'type',
      'note',
      'description',
      'do',
    ])
    const moneyMoveDateIndex = findColumnIndex(moneyMoveData.headers, ['date', 'time'])
    const moneyMoveCurrencyIndex = findColumnIndex(moneyMoveData.headers, ['currency'])

    moneyMoveData.rows.forEach((row, rowIndex) => {
      if (moneyMoveAmountIndex === -1 || moneyMoveActionIndex === -1 || moneyMoveDateIndex === -1) {
        return
      }

      const amount = parseNumber(row[moneyMoveAmountIndex] ?? '')
      const action = resolveMoneyMoveType(row[moneyMoveActionIndex] ?? '')
      const date = parseDateValue(row[moneyMoveDateIndex] ?? '')

      if (!Number.isFinite(amount) || !action || !date) {
        console.warn('[Timeline] Skipped money move row:', { amount, action, date, row })
        return
      }

      const rowCurrency = resolveCurrencyCode(row[moneyMoveCurrencyIndex] ?? '')
      const normalizedAmount = convertAmount(
        Math.abs(amount),
        rowCurrency,
        selectedCurrency,
        currencyRates,
      )

      const delta = action === 'in' || action === 'bor' ? normalizedAmount : -normalizedAmount

      events.push({
        date,
        sourceOrder: rowIndex,
        cashDelta: delta,
      })
    })

    const stocksDateIndex = findColumnIndex(stocksData.headers, ['date', 'time'])
    const stocksActionIndex = findColumnIndex(stocksData.headers, [
      'action',
      'type',
      'side',
      'buysell',
      'operation',
    ])
    const stocksSymbolIndex = findColumnIndex(stocksData.headers, [
      'symbol',
      'ticker',
      'stock',
      'code',
      'company',
    ])
    const stocksQuantityIndex = findColumnIndex(stocksData.headers, [
      'qty',
      'quantity',
      'share',
      'units',
      'vol',
      'volume',
    ])
    const stocksTotalIndex = findColumnIndex(stocksData.headers, [
      'total',
      'amount',
      'value',
      'proceeds',
      'market value',
      'marketvalue',
    ])
    const stocksPriceIndex = findColumnIndex(stocksData.headers, [
      'price',
      'avgprice',
      'unitprice',
      'cost',
      'rate',
    ])
    const stocksCurrencyIndex = findColumnIndex(stocksData.headers, ['currency'])
    const stocksFeeIndex = findColumnIndex(stocksData.headers, [
      'handlingfee',
      'handling',
      'fee',
      'commission',
      'charge',
      'charges',
    ])

    stocksData.rows.forEach((row, rowIndex) => {
      // We absolutely need Date and Quantity to process a stock row for the timeline
      if (stocksDateIndex === -1 || stocksQuantityIndex === -1) {
        return
      }

      const date = parseDateValue(row[stocksDateIndex] ?? '')
      const quantitySigned = parseNumber(row[stocksQuantityIndex] ?? '')
      
      // Resolve action: Explicit column > inferred from sign > default to buy if +qty
      const actionRaw = stocksActionIndex >= 0 ? row[stocksActionIndex] ?? '' : ''
      const action = resolveAction(actionRaw, quantitySigned) ?? (quantitySigned > 0 ? 'buy' : 'sell')
      
      // Resolve Total: Explicit total > Price * Qty
      let total = stocksTotalIndex >= 0 ? parseNumber(row[stocksTotalIndex] ?? '') : Number.NaN
      
      // Fallback if Total is missing but we have Price
      if (!Number.isFinite(total) && stocksPriceIndex >= 0) {
        const price = parseNumber(row[stocksPriceIndex] ?? '')
        if (Number.isFinite(price)) {
          total = Math.abs(price * Math.abs(quantitySigned))
        }
      }

      if (!date || !Number.isFinite(total)) {
        // Only warn if we really can't calculate a financial impact
        if (rowIndex < 5) {
             console.warn('[Timeline] Skipped stock row due to missing data:', { date, total, rowIndex })
        }
        return
      }

      const rowCurrency = resolveCurrencyCode(row[stocksCurrencyIndex] ?? '')
      const tradeAmount = convertAmount(
        Math.abs(total),
        rowCurrency,
        selectedCurrency,
        currencyRates,
      )

      const feeRaw = stocksFeeIndex >= 0 ? parseNumber(row[stocksFeeIndex] ?? '') : 0
      const fee = Number.isFinite(feeRaw)
        ? convertAmount(Math.abs(feeRaw), rowCurrency, selectedCurrency, currencyRates)
        : 0

      const quantity = Math.abs(quantitySigned)
      // Cost per share used for Cost Basis tracking
      // If we bought, cost is total / qty.
      // If we sold, we use the historical cost basis (FIFO/Avg) logic later, 
      // but here we just need to know the money flow.
      const costPerShare = quantity > 0 ? tradeAmount / quantity : 0
      
      const symbol = stocksSymbolIndex >= 0 ? (row[stocksSymbolIndex] ?? '').trim().toUpperCase() : 'UNKNOWN'

      const cashDelta = action === 'buy' ? -(tradeAmount + fee) : tradeAmount - fee

      events.push({
        date,
        sourceOrder: 10_000 + rowIndex,
        cashDelta,
        stockAction:
          quantity > 0
            ? { symbol, action, quantity, costPerShare }
            : undefined,
      })
    })

    const dividendAmountIndex = findColumnIndex(dividendData.headers, ['dividend', 'div', 'amount', 'value'])
    const dividendDateIndex = findColumnIndex(dividendData.headers, ['date', 'time'])
    const dividendCurrencyIndex = findColumnIndex(dividendData.headers, ['currency'])

    dividendData.rows.forEach((row, rowIndex) => {
      if (dividendAmountIndex === -1 || dividendDateIndex === -1) {
        return
      }

      const amount = parseNumber(row[dividendAmountIndex] ?? '')
      const date = parseDateValue(row[dividendDateIndex] ?? '')

      if (!Number.isFinite(amount) || !date) {
        console.warn('[Timeline] Skipped dividend row:', { amount, date, row })
        return
      }

      const rowCurrency = resolveCurrencyCode(row[dividendCurrencyIndex] ?? '')
      const normalizedAmount = convertAmount(amount, rowCurrency, selectedCurrency, currencyRates)

      events.push({
        date,
        sourceOrder: 20_000 + rowIndex,
        cashDelta: normalizedAmount,
      })
    })

    console.log(`[Timeline] Collected ${events.length} total events`)

    events.sort((a, b) => {
      const timeDiff = a.date.getTime() - b.date.getTime()
      if (timeDiff !== 0) {
        return timeDiff
      }
      return a.sourceOrder - b.sourceOrder
    })

    const labels: string[] = []
    const moneyIHaveValues: number[] = []
    const totalMoneyValues: number[] = []

    let runningCash = 0
    const holdingsBySymbol = new Map<string, { quantity: number; costBasis: number }>()

    events.forEach((event, index) => {
      runningCash += event.cashDelta

      if (event.stockAction) {
        const { symbol, action, quantity, costPerShare } = event.stockAction
        const costForThisPurchase = quantity * costPerShare
        const holding = holdingsBySymbol.get(symbol) ?? { quantity: 0, costBasis: 0 }

        if (action === 'buy') {
          holding.quantity += quantity
          holding.costBasis += costForThisPurchase
        } else if (action === 'sell') {
          const quantityToSell = Math.min(quantity, holding.quantity)
          const avgCostPerShare = holding.quantity > 0 ? holding.costBasis / holding.quantity : 0
          const soldCostBasis = quantityToSell * avgCostPerShare
          holding.quantity -= quantityToSell
          holding.costBasis = Math.max(0, holding.costBasis - soldCostBasis)

          if (holding.quantity < 0.0001) {
            holding.quantity = 0
            holding.costBasis = 0
          }
        }

        holdingsBySymbol.set(symbol, holding)
      }

      const totalHoldingsCostBasis = Array.from(holdingsBySymbol.values()).reduce(
        (sum, holding) => sum + holding.costBasis,
        0,
      )
      const totalMoney = runningCash + totalHoldingsCostBasis

      labels.push(formatTimelineLabel(event.date, index))
      moneyIHaveValues.push(Number(runningCash.toFixed(2)))
      totalMoneyValues.push(Number(totalMoney.toFixed(2)))
    })

    console.log(
      `[Timeline] Generated ${labels.length} data points | Cash range: ${moneyIHaveValues[0]} to ${moneyIHaveValues[moneyIHaveValues.length - 1]} | Total range: ${totalMoneyValues[0]} to ${totalMoneyValues[totalMoneyValues.length - 1]}`,
    )

    return { labels, moneyIHaveValues, totalMoneyValues }
  }, [
    moneyMoveData,
    stocksData,
    dividendData,
    selectedCurrency,
    currencyRates,
  ])

  const holdingsDonutChart = useMemo(() => {
    const data = holdings
      .map((holding, index) => {
        const nowPriceUsd = resolveLiveQuotePrice(liveQuotes[holding.symbol])
        const priceUsd = Number.isFinite(nowPriceUsd) ? nowPriceUsd : holding.avgBuyPriceUsd
        const priceInSelectedCurrency = convertAmount(
          priceUsd,
          'USD',
          selectedCurrency,
          currencyRates,
        )

        return {
          id: index,
          label: holding.symbol,
          value: Number((priceInSelectedCurrency * holding.quantity).toFixed(2)),
        }
      })
      .filter((item) => Number.isFinite(item.value) && item.value > 0)

    return data
  }, [holdings, liveQuotes, selectedCurrency, currencyRates])

  const monthlyEarnChart = useMemo(() => {
    const monthEarnMap = new Map<string, number>()

    const stocksDateIndex = findColumnIndex(stocksData.headers, ['date', 'time'])
    const stocksCurrencyIndex = findColumnIndex(stocksData.headers, ['currency'])
    const profitLossIndex = findColumnIndex(displayStocksData.headers, ['profitloss'])

    if (stocksDateIndex !== -1 && profitLossIndex !== -1) {
      displayStocksData.rows.forEach((row, rowIndex) => {
        const profitLoss = parseNumber(row[profitLossIndex] ?? '')
        const dateValue = parseDateValue(stocksData.rows[rowIndex]?.[stocksDateIndex] ?? '')

        if (!Number.isFinite(profitLoss) || !dateValue) {
          return
        }

        const rowCurrency = resolveCurrencyCode(stocksData.rows[rowIndex]?.[stocksCurrencyIndex] ?? '')
        const normalizedProfitLoss = convertAmount(
          profitLoss,
          rowCurrency,
          selectedCurrency,
          currencyRates,
        )

        const monthKey = toMonthKey(dateValue)
        monthEarnMap.set(monthKey, (monthEarnMap.get(monthKey) ?? 0) + normalizedProfitLoss)
      })
    }

    const dividendAmountIndex = findColumnIndex(dividendData.headers, ['dividend', 'div', 'amount', 'value'])
    const dividendDateIndex = findColumnIndex(dividendData.headers, ['date', 'time'])
    const dividendCurrencyIndex = findColumnIndex(dividendData.headers, ['currency'])

    if (dividendAmountIndex !== -1 && dividendDateIndex !== -1) {
      dividendData.rows.forEach((row) => {
        const amount = parseNumber(row[dividendAmountIndex] ?? '')
        const dateValue = parseDateValue(row[dividendDateIndex] ?? '')

        if (!Number.isFinite(amount) || !dateValue) {
          return
        }

        const rowCurrency = resolveCurrencyCode(row[dividendCurrencyIndex] ?? '')
        const normalizedAmount = convertAmount(amount, rowCurrency, selectedCurrency, currencyRates)

        const monthKey = toMonthKey(dateValue)
        monthEarnMap.set(monthKey, (monthEarnMap.get(monthKey) ?? 0) + normalizedAmount)
      })
    }

    const labels = Array.from(monthEarnMap.keys()).sort((a, b) => a.localeCompare(b))
    const values = labels.map((label) => Number((monthEarnMap.get(label) ?? 0).toFixed(2)))

    return { labels, values }
  }, [stocksData, displayStocksData, dividendData, selectedCurrency, currencyRates])

  useEffect(() => {
    if (!isConfigValid) {
      setError(
        'Missing required env vars in .env.local: VITE_PUBLIC_SUPABASE_URL + VITE_PUBLIC_SUPABASE_ANON_KEY (or VITE_SUPABASE_URL + VITE_SUPABASE_ANON_KEY)',
      )
      setIsLoading(false)
      return
    }

    if (!isAuthReady) {
      setIsLoading(true)
      return
    }

    if (!currentUserId) {
      setStocksData({ headers: [], rows: [], ids: [] })
      setDividendData({ headers: [], rows: [], ids: [] })
      setMoneyMoveData({ headers: [], rows: [], ids: [] })
      setError(null)
      setIsLoading(false)
      return
    }

    const loadData = async () => {
      setIsLoading(true)
      setError(null)

      try {
        const { stocks, dividend, moneyMove } = await loadUserSheetData(currentUserId)

        setStocksData(stocks)
        setDividendData(dividend)
        setMoneyMoveData(moneyMove)
      } catch (err) {
        const message = err instanceof Error ? err.message : 'Unknown error'
        setError(`Failed to load your Supabase data. Details: ${message}`)
      } finally {
        setIsLoading(false)
      }
    }

    void loadData()
  }, [currentUserId, isAuthReady, isConfigValid])

  const renderSheetTable = (title: string, data: SheetData) => {
    const shouldIncludeColumn = (header: string): boolean => {
      const normalized = normalizeHeader(header)

      if (title === 'Stocks') {
        if (header.trim() === '%') {
          return true
        }

        return [
          'stock',
          'symbol',
          'ticker',
          'currency',
          'price',
          'action',
          'time',
          'date',
          'quantity',
          'qty',
          'handlingfee',
          'handling',
          'fee',
          'profitloss',
          'percent',
        ].some((key) => normalized.includes(normalizeHeader(key)))
      }

      if (title === 'Dividend') {
        return ['stock', 'symbol', 'ticker', 'currency', 'div', 'dividend', 'time', 'date'].some(
          (key) => normalized.includes(normalizeHeader(key)),
        )
      }

      if (title === 'Money Move') {
        return ['name', 'currency', 'price', 'amount', 'time', 'date', 'do', 'action', 'type'].some(
          (key) => normalized.includes(normalizeHeader(key)),
        )
      }

      return true
    }

    const matchedColumnIndexes = data.headers
      .map((header, index) => ({ header, index }))
      .filter(({ header }) => shouldIncludeColumn(header))
      .map(({ index }) => index)

    const fallbackColumnIndexes = (() => {
      if (title === 'Stocks') {
        const stockDefaults = [0, 1, 2, 3, 4, 5, 6]
        const derivedDefaults =
          data.headers.length >= 2 ? [data.headers.length - 2, data.headers.length - 1] : []

        return [...stockDefaults, ...derivedDefaults].filter(
          (index, position, source) =>
            index >= 0 && index < data.headers.length && source.indexOf(index) === position,
        )
      }

      if (title === 'Dividend') {
        return [0, 1, 2, 3].filter((index) => index < data.headers.length)
      }

      if (title === 'Money Move') {
        return [0, 1, 2, 3, 4].filter((index) => index < data.headers.length)
      }

      return data.headers.map((_, index) => index)
    })()

    const visibleColumnIndexes =
      matchedColumnIndexes.length > 0 ? matchedColumnIndexes : fallbackColumnIndexes

    const formatCellValue = (value: string, header: string): string => {
      const headerLower = header.toLowerCase()
      const isDateColumn = headerLower.includes('date') || headerLower.includes('time')
      
      if (isDateColumn && value.trim()) {
        const parsed = parseDateValue(value)
        if (parsed) {
          return formatDateDDMMYYYY(parsed)
        }
      }
      
      return value
    }

    return (
      <Box className="sheet-section">
        <h1 className="sheet-title text-black">{title}</h1>
        <TableContainer component={Paper} elevation={0} className="table-shell">
          <Table size="small" aria-label={`${title.toLowerCase()} table`}>
            <TableHead>
              <TableRow>
                {visibleColumnIndexes.map((headerIndex) => (
                  <TableCell key={`${title}-header-${headerIndex}`} sx={{ fontWeight: 700 }}>
                    {data.headers[headerIndex]}
                  </TableCell>
                ))}
                <TableCell sx={{ fontWeight: 700 }}>Action</TableCell>
              </TableRow>
            </TableHead>
            <TableBody>
              {data.rows.map((row, rowIndex) => (
                <TableRow key={`${title}-${rowIndex}-${row.join('-')}`}>
                  {visibleColumnIndexes.map((colIndex) => (
                    <TableCell key={`${title}-${rowIndex}-${colIndex}`}>
                      {formatCellValue(row[colIndex] ?? '', data.headers[colIndex] ?? '')}
                    </TableCell>
                  ))}
                  <TableCell>
                    <Button
                      variant="text"
                      className="live-edit-button"
                      onClick={() => handleOpenEditModal(title as AddTableName, rowIndex)}
                      aria-label={`Edit row ${rowIndex}`}
                    >
                      <MdEdit size={20} />
                    </Button>
                  </TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
        </TableContainer>
      </Box>
    )
  }

  const renderEditModal = () => {
    if (!editingTable || editingRowIndex === null) {
      return null
    }

    const editFields = ADD_FORM_CONFIG[editingTable]

    return (
      <Modal
        open={editingTable !== null && editingRowIndex !== null}
        onClose={handleCloseEditModal}
        sx={{
          display: 'flex',
          justifyContent: 'center',
          alignItems: 'center',
        }}
      >
        <Paper
          elevation={8}
          className="add-form-card"
          sx={{
            width: '90%',
            maxWidth: '600px',
            height: 'auto',
            maxHeight: '90vh',
            overflow: 'auto',
            backgroundColor: 'var(--bg)',
            color: 'var(--text)',
            padding: '2rem',
          }}
        >
          <Box component="form" className="add-form-grid edit-form-grid">
            <h2 style={{ textAlign: 'center', marginBottom: '1.5rem', color: 'var(--text)' }}>Edit {editingTable}</h2>

            {editFields.map((field) => {
              const pairedFieldKey = editingTable === 'Dividend' ? 'div' : 'price'

              if (editingTable === 'Stocks' && field.key === 'stock') {
                const actionField = editFields.find((candidate) => candidate.key === 'action')
                const quantityField = editFields.find((candidate) => candidate.key === 'quantity')

                return (
                  <Box className="add-form-triple-row" key="stocks-main-row">
                    {renderEditField(field, 'add-field-stock')}
                    {actionField ? renderEditField(actionField, 'add-field-action') : null}
                    {quantityField ? renderEditField(quantityField, 'add-field-quantity') : null}
                  </Box>
                )
              }

              if (editingTable === 'Stocks' && (field.key === 'action' || field.key === 'quantity')) {
                return null
              }

              if (editingTable === 'Money Move' && field.key === 'name') {
                const doField = editFields.find((candidate) => candidate.key === 'do')

                return (
                  <Box className="add-form-name-do-row" key="money-move-name-do-row">
                    {renderEditField(field, 'add-field-name')}
                    {doField ? renderEditField(doField, 'add-field-do') : null}
                  </Box>
                )
              }

              if (editingTable === 'Money Move' && field.key === 'do') {
                return null
              }

              if (field.key === pairedFieldKey) {
                return null
              }

              if (field.key === 'currency') {
                const pairedField = editFields.find((candidate) => candidate.key === pairedFieldKey)

                return (
                  <Box className="add-form-pair-row" key={`pair-${editingTable}-${pairedFieldKey}`}>
                    {renderEditField(field, 'add-field-currency')}
                    {pairedField ? renderEditField(pairedField, 'add-field-value') : null}
                  </Box>
                )
              }

              return renderEditField(field)
            })}

            {editFormMessage ? (
              <Alert severity={editFormMessage.type}>{editFormMessage.text}</Alert>
            ) : null}

            <Box
              className="edit-form-button"
              // sx={{
              //   display: 'flex',
              //   justifyContent: 'space-between',
              //   gap: '1rem',
              //   marginTop: '1.5rem',
              // }}
            >
              <Button
                type="button"
                variant="contained"
                disabled={isSubmittingEdit || !currentUserId}
                className={isConfirmingDelete ? 'is-confirming' : 'delete'}
                onClick={() => handleSubmitEditDelete('delete')}
                // sx={{
                //   backgroundColor: 'var(--special)',
                //   color: 'var(--text)',
                //   '&:hover': {
                //     backgroundColor: 'var(--accent)',
                //     opacity: 0.92,
                //   },
                //   '&:disabled': {
                //     backgroundColor: 'var(--special)',
                //     opacity: 0.5,
                //   },
                // }}
              >
                {isSubmittingEdit
                  ? 'Deleting...'
                  : isConfirmingDelete
                    ? 'Confirm'
                    : 'Delete'}
              </Button>

              <Button
                type="button"
                variant="contained"
                disabled={isSubmittingEdit || !currentUserId}
                className={isConfirmingEdit ? 'is-confirming' : ''}
                onClick={() => handleSubmitEditDelete('update')}
                // sx={{
                //   backgroundColor: '#000000',
                //   borderRadius: '99px',
                //   color: 'var(--bg)',
                //   '&:hover': {
                //     backgroundColor: '#333333',
                //     opacity: 0.92,
                //   },
                //   '&:disabled': {
                //     backgroundColor: '#000000',
                //     opacity: 0.5,
                //   },
                // }}
              >
                {isSubmittingEdit
                  ? 'Saving...'
                  : isConfirmingEdit
                    ? 'Confirm'
                    : 'Change'}
              </Button>
            </Box>
          </Box>
        </Paper>
      </Modal>
    )
  }

  const renderGoalModal = () => {
    return (
      <Modal
        open={isGoalModalOpen}
        onClose={handleCloseGoalModalMain}
        sx={{
          display: 'flex',
          justifyContent: 'center',
          alignItems: 'center',
        }}
      >
        <Paper
          elevation={8}
          className="add-form-card"
          sx={{
            width: '90%',
            maxWidth: '600px',
            height: 'auto',
            maxHeight: '90vh',
            overflow: 'auto',
            backgroundColor: 'var(--white)',
            color: 'var(--black)',
            padding: '2rem',
          }}
        >
          <Box component="form" className="add-form-grid goal-form-grid" onSubmit={handleSubmitGoal}>
            <h1 className='sheet-title-box text-black' >Create New Goal</h1>

            {renderGoalFields()}

            <Box className="add-form-actions">
              <Button
                type="submit"
                variant="contained"
                disabled={isSubmittingGoal || !currentUserId}
                className={isConfirmingGoal ? 'is-confirming' : ''}
              >
                {isSubmittingGoal
                  ? 'Saving...'
                  : isConfirmingGoal
                    ? 'Confirm'
                    : currentUserId
                      ? 'Save Goal'
                      : 'Sign in required'}
              </Button>
              {/* <Button
                type="button"
                variant="outlined"
                onClick={handleCloseGoalModalMain}
              >
                Cancel
              </Button> */}
            </Box>

            {goalMessage ? <Alert severity={goalMessage.type}>{goalMessage.text}</Alert> : null}
          </Box>
        </Paper>
      </Modal>
    )
  }

  const formatLiveQuoteValue = (value: number | undefined, type: 'price' | 'change' | 'percent') => {
    if (typeof value !== 'number' || !Number.isFinite(value)) {
      return '-'
    }

    if (type === 'percent') {
      return `${value.toFixed(2)}%`
    }

    return value.toFixed(2)
  }

  const latestMoneyIHave =
    moneyTimelineChart.moneyIHaveValues.length > 0
      ? moneyTimelineChart.moneyIHaveValues[moneyTimelineChart.moneyIHaveValues.length - 1]
      : 0

  const selectedLivePriceUsd =
    calcSelectedSymbol && liveQuotes[calcSelectedSymbol]
      ? resolveLiveQuotePrice(liveQuotes[calcSelectedSymbol])
      : Number.NaN
  const selectedLivePrice = Number.isFinite(selectedLivePriceUsd)
    ? convertAmount(selectedLivePriceUsd, 'USD', selectedCurrency, currencyRates)
    : Number.NaN

  const handleCalculatorSymbolChange = (_event: unknown, value: string | null) => {
    setCalcSelectedSymbol(value)

    if (!value) {
      return
    }

    const quote = liveQuotes[value]
    const quotePriceUsd = resolveLiveQuotePrice(quote)

    if (Number.isFinite(quotePriceUsd)) {
      const convertedPrice = convertAmount(quotePriceUsd, 'USD', selectedCurrency, currencyRates)

      if (Number.isFinite(convertedPrice)) {
        setCalcStockPriceInput(convertedPrice.toFixed(2))
      }
    }

    if (calculatorType === 'sell') {
      const holding = holdings.find((item) => item.symbol === value)
      if (holding) {
        setCalcQuantityInput(holding.quantity.toFixed(2))
      }
    }
  }

  const handleOpenToolsModal = () => {
    if (!calcBudgetInput && Number.isFinite(latestMoneyIHave)) {
      setCalcBudgetInput(latestMoneyIHave.toFixed(2))
    }

    setIsToolsModalOpen(true)
  }

  const calculatorSymbolOptions = calculatorType === 'sell' ? requiredHoldingSymbols : liveSymbols

  const selectedCalcHolding = useMemo(() => {
    if (!calcSelectedSymbol) {
      return null
    }

    return holdings.find((holding) => holding.symbol === calcSelectedSymbol) ?? null
  }, [holdings, calcSelectedSymbol])

  useEffect(() => {
    if (calculatorType !== 'sell') {
      return
    }

    if (!calcSelectedSymbol) {
      return
    }

    const holding = holdings.find((item) => item.symbol === calcSelectedSymbol)
    if (!holding) {
      setCalcSelectedSymbol(null)
      setCalcQuantityInput('')
      return
    }

    if (!calcQuantityInput) {
      setCalcQuantityInput(holding.quantity.toFixed(2))
    }
  }, [calculatorType, calcSelectedSymbol, calcQuantityInput, holdings])

  const calcBudget = Number(calcBudgetInput)
  const calcStockPrice = Number(calcStockPriceInput)
  const calcQuantity = Number(calcQuantityInput)
  const selectedCalcHoldingPrice = selectedCalcHolding
    ? convertAmount(selectedCalcHolding.avgBuyPriceUsd, 'USD', selectedCurrency, currencyRates)
    : Number.NaN

  const canCalculateBuyCount =
    calculatorType === 'buy' &&
    Number.isFinite(calcBudget) &&
    Number.isFinite(calcStockPrice) &&
    calcBudget > 0 &&
    calcStockPrice > 0

  const maxBuyShares = canCalculateBuyCount ? Math.floor(calcBudget / calcStockPrice) : 0
  const remainingMoney = canCalculateBuyCount
    ? Number((calcBudget - maxBuyShares * calcStockPrice).toFixed(2))
    : 0

  const calcSellHoldingQuantity = selectedCalcHolding?.quantity ?? 0
  const isSellQuantityValid =
    calculatorType === 'sell' &&
    Number.isFinite(calcStockPrice) &&
    Number.isFinite(calcQuantity) &&
    calcStockPrice > 0 &&
    calcQuantity > 0 &&
    Boolean(selectedCalcHolding) &&
    calcQuantity <= calcSellHoldingQuantity

  const sellGainLoss = isSellQuantityValid
    ? Number(((calcStockPrice - selectedCalcHoldingPrice) * calcQuantity).toFixed(2))
    : 0
  const sellGainLossPercent = isSellQuantityValid && selectedCalcHoldingPrice > 0
    ? Number((((calcStockPrice - selectedCalcHoldingPrice) / selectedCalcHoldingPrice) * 100).toFixed(2))
    : 0
  const sellGainLossAbs = Math.abs(sellGainLoss)

  const sellStatusMessage = (() => {
    if (calculatorType !== 'sell') {
      return null
    }

    if (!calcSelectedSymbol) {
      return 'Select a stock you currently hold to calculate realized gain or loss.'
    }

    if (!selectedCalcHolding) {
      return `No holding record found for ${calcSelectedSymbol}.`
    }

    if (!Number.isFinite(calcStockPrice) || calcStockPrice <= 0) {
      return 'Enter a valid sell price.'
    }

    if (!Number.isFinite(calcQuantity) || calcQuantity <= 0) {
      return 'Enter a valid quantity to sell.'
    }

    if (calcQuantity > calcSellHoldingQuantity) {
      return `Quantity exceeds your current holding of ${calcSellHoldingQuantity.toFixed(2)} shares.`
    }

    return null
  })()

  return (
    <main className="page">
      <header className="page-header">
        <Box className="header-all">
        <div className="page-nav">
          <Button
            variant={page === 'dataShow' ? 'contained' : 'outlined'}
            sx={{
              width: 40,
              minWidth: 40,
              height: 40,
              padding: 0,
              borderRadius: '50%',
              backgroundColor: 'var(--bg)', 
              color: 'var(--wbg)',
              borderColor: 'var(--wbg)',
              '&.MuiButton-contained': {
              backgroundColor: 'var(--special)',
              color: 'var(--text)',
              '&:hover': {
                borderColor: 'var(--accent)',
                opacity: 0.92,
                },
              }
            }}
            onClick={() => setPage('dataShow')}
          >
            <MdDataUsage size={20} />
          </Button>
          <Button
            variant={page === 'live' ? 'contained' : 'outlined'}
            sx={{
              width: 40,
              minWidth: 40,
              height: 40,
              padding: 0,
              borderRadius: '50%',
              backgroundColor: 'var(--bg)',
              color: 'var(--wbg)',
              borderColor: 'var(--wbg)',
              '&.MuiButton-contained': {
                backgroundColor: 'var(--special)',
                color: 'var(--text)',
                '&:hover': {
                  borderColor: 'var(--accent)',
                  opacity: 0.92,
                },
              },
            }}
            onClick={() => setPage('live')}
          >
            <MdShowChart size={20} />
          </Button>
          <Button 
            variant={page === 'table' ? 'contained' : 'outlined'} 
            sx={{
                width: 40,
                minWidth: 40,
                height: 40,
                padding: 0,
                borderRadius: '50%',
                backgroundColor: 'var(--bg)', 
                color: 'var(--wbg)',
                borderColor: 'var(--wbg)',
                '&.MuiButton-contained': {
                backgroundColor: 'var(--special)',
                color: 'var(--text)',
                '&:hover': {
                  borderColor: 'var(--accent)',
                  opacity: 0.92,
                  },
                }
              }}
            onClick={() => setPage('table')}
          >
            <MdTableRows size={20} />
          </Button>
          <Button
            variant={page === 'add' ? 'contained' : 'outlined'}
            sx={{
              width: 40,
              minWidth: 40,
              height: 40,
              padding: 0,
              borderRadius: '50%',
              backgroundColor: 'var(--bg)',
              color: 'var(--wbg)',
              borderColor: 'var(--wbg)',
              '&.MuiButton-contained': {
                backgroundColor: 'var(--special)',
                color: 'var(--text)',
                '&:hover': {
                  borderColor: 'var(--accent)',
                  opacity: 0.92,
                },
              },
            }}
            onClick={() => setPage('add')}
          >
            <MdOutlineAdd size={20} />
          </Button>
        </div>

        <div className="header-controls page-nav">
          <Button
            aria-haspopup="menu"
            aria-expanded={isSettingsOpen ? 'true' : undefined}
            aria-controls={isSettingsOpen ? 'header-settings-menu' : undefined}
            variant="outlined"
            sx={{
              width: 40,
              minWidth: 40,
              height: 40,
              padding: 0,
              borderRadius: '50%',
              backgroundColor: 'var(--bg)',
              color: 'var(--wbg)',
              borderColor: 'var(--wbg)',
              '&:hover': {
                borderColor: 'var(--wbg)',
                opacity: 0.85,
              },
            }}
            onClick={handleOpenSettings}
          >
            <MdSettings size={20} />
          </Button>

          <Menu
            id="header-settings-menu"
            anchorEl={settingsAnchorEl}
            open={isSettingsOpen}
            onClose={handleCloseSettings}
            anchorOrigin={{ vertical: 'bottom', horizontal: 'right' }}
            transformOrigin={{ vertical: 'top', horizontal: 'right' }}
            slotProps={{
              paper: {
                className: 'settings-menu-paper',
              },
            }}
          >
            <Box className="settings-menu-content">
              {currentUserName && currentUserEmail && currentUserId &&
                <Box>
                  <Box className="settings-menu-header">
                    <h3 className="settings-menu-title">{currentUserName}</h3>
                    <Button
                    variant="outlined"
                    onClick={handleSignOut}
                    className="settings-action-button"
                    >
                      Logout
                    </Button>
                  </Box>
                  <p className="settings-menu-text">{currentUserEmail}</p>
                  <p className="settings-menu-text">ID: {currentUserId}</p>
                </Box>
              }
              <div className="settings-menu-row">
                {!currentUserId && (
                  <Button
                    variant="contained"
                    onClick={handleGoogleSignIn}
                    className="settings-action-button settings-action-button-primary"
                  >
                    Login
                  </Button>
                )}
              </div>

              {currentUserId ? (
                <Autocomplete
                  className="currency-picker settings-currency-picker"
                  sx={{
                    width: '100%',
                    '& .MuiOutlinedInput-root': {
                      height: '40px',
                      borderRadius: '12px',
                      color: 'var(--text)',
                      '& fieldset': { borderColor: 'var(--wbg)' },
                      '&:hover fieldset': { borderColor: 'var(--wbg)', opacity: 0.92 },
                      '&.Mui-focused fieldset': { borderColor: 'var(--special)', borderWidth: '1px' },
                    },
                    '& .MuiInputLabel-root': { color: 'var(--wbg)', fontSize: '0.85rem' },
                    '& .MuiInputLabel-root.Mui-focused': { color: 'var(--special)' },
                    '& .MuiSvgIcon-root': { color: 'var(--text)' },
                  }}
                  size="small"
                  disableClearable
                  options={[...CURRENCY_OPTIONS]}
                  value={selectedCurrency}
                  onChange={(_, value) => setSelectedCurrency(value)}
                  renderInput={(params) => <TextField {...params} label="Currency" />}
                />
              ) : null}

              <Box className="settings-item-row">
                <p className="settings-item-label">Mode</p>
                <ToggleButtonGroup
                  value={themeMode}
                  exclusive
                  size="small"
                  onChange={handleToggleThemeMode}
                  className="settings-mode-toggle"
                  aria-label="Theme mode"
                >
                  <ToggleButton value="system" aria-label="System mode" title="System">
                    <MdBrightnessAuto size={16} />
                  </ToggleButton>
                  <ToggleButton value="dark" aria-label="Dark mode" title="Dark">
                    <MdDarkMode size={16} />
                  </ToggleButton>
                  <ToggleButton value="light" aria-label="Light mode" title="Light">
                    <MdLightMode size={16} />
                  </ToggleButton>
                </ToggleButtonGroup>
              </Box>

              <Box className="settings-item-row">
                <p className="settings-item-label">Notifications</p>
                <Switch checked={notificationsEnabled} onChange={handleToggleNotifications} />
              </Box>

              {notificationPermission === 'denied' ? (
                <p className="settings-note">Browser notifications are blocked. Enable them in browser settings.</p>
              ) : null}
            </Box>
          </Menu>
        </div>
        </Box>
      </header>

      <Box className="fixed-button">
        <Button
          variant="contained"
          sx={{
                width: 50,
                minWidth: 50,
                height: 50,
                padding: 0,
                borderRadius: '50%',
                backgroundColor: 'var(--bg)', 
                color: 'var(--wbg)',
                borderColor: 'var(--wbg)',
                '&.MuiButton-contained': {
                backgroundColor: 'var(--special)',
                color: 'var(--text)',
                '&:hover': {
                  borderColor: 'var(--accent)',
                  opacity: 0.92,
                  },
                }
              }}
          onClick={handleOpenToolsModal}
        >
          <MdQueryStats size={20} />
        </Button>
      </Box>

      <Modal
        open={isToolsModalOpen}
        onClose={() => setIsToolsModalOpen(false)}
        sx={{
          display: 'flex',
          justifyContent: 'center',
          alignItems: 'center',
        }}
      >
        <Box className="tools-modal">
            {/* <h1 className="sheet-title-box text-black">Extra Functions</h1> */}
            {/* <Button variant="text" className="live-unfollow-button" onClick={() => setIsToolsModalOpen(false)}>
              <MdCancel size={20} />
            </Button> */}

          <Box className="tools-section">
            <Box
              className="tools-section-header"
              sx={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 1 }}
            >
              <h3 className="tools-section-title">
                {calculatorType === 'buy' ? 'Stock Buy Calculator' : 'Stock Sell Calculator'}
              </h3>
              <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
                {/* <p className="tools-section-note" style={{ margin: 0 }}>Buy</p> */}
                <Switch
                  checked={calculatorType === 'buy'}
                  onChange={(event) => setCalculatorType(event.target.checked ? 'buy' : 'sell')}
                />
                {/* <p className="tools-section-note" style={{ margin: 0 }}>Sell</p> */}
              </Box>
            </Box>

            <Box className="tools-calc-grid">
              {calculatorType === 'buy' ? (
                <TextField
                  type="number"
                  label={`Money Available (Money I Have, ${selectedCurrency})`}
                  value={calcBudgetInput}
                  onChange={(event) => setCalcBudgetInput(event.target.value)}
                  inputProps={{ min: '0', step: '0.01' }}
                  size="small"
                  variant="standard"
                />
              ) : (
                <Autocomplete
                  options={calculatorSymbolOptions}
                  value={calcSelectedSymbol}
                  onChange={handleCalculatorSymbolChange}
                  renderInput={(params) => (
                    <TextField
                      {...params}
                      label="Stock to sell"
                      size="small"
                      variant="standard"
                    />
                  )}
                />
              )}

              {calculatorType === 'buy' ? (
                <Autocomplete
                  options={calculatorSymbolOptions}
                  value={calcSelectedSymbol}
                  onChange={handleCalculatorSymbolChange}
                  renderInput={(params) => (
                    <TextField
                      {...params}
                      label="Live Symbol (optional)"
                      size="small"
                      variant="standard"
                    />
                  )}
                />
              ) : (
                <TextField
                  type="number"
                  label={`Sell Price (${selectedCurrency})`}
                  value={calcStockPriceInput}
                  onChange={(event) => setCalcStockPriceInput(event.target.value)}
                  inputProps={{ min: '0', step: '0.01' }}
                  size="small"
                  variant="standard"
                />
              )}

              {calculatorType === 'buy' ? (
                <TextField
                  type="number"
                  label={`Stock Price (manual or live, ${selectedCurrency})`}
                  value={calcStockPriceInput}
                  onChange={(event) => setCalcStockPriceInput(event.target.value)}
                  inputProps={{ min: '0', step: '0.01' }}
                  size="small"
                  variant="standard"
                />
              ) : (
                <TextField
                  type="number"
                  label="Quantity to sell"
                  value={calcQuantityInput}
                  onChange={(event) => setCalcQuantityInput(event.target.value)}
                  inputProps={{ min: '0', step: '0.01' }}
                  size="small"
                  variant="standard"
                  sx={{ gridColumn: '1 / -1' }}
                />
              )}
            </Box>

            {calculatorType === 'buy' ? (
              calcSelectedSymbol ? (
                <p className="tools-section-note">
                  {Number.isFinite(selectedLivePrice)
                    ? `Live price for ${calcSelectedSymbol}: ${formatCurrency(selectedLivePrice, selectedCurrency)} (auto-filled)`
                    : `No live quote available now for ${calcSelectedSymbol}. You can still type price manually.`}
                </p>
              ) : null
            ) : (
              <>
                {calcSelectedSymbol ? (
                  <p className="tools-section-note">
                    {Number.isFinite(selectedLivePrice)
                      ? `Live price for ${calcSelectedSymbol}: ${formatCurrency(selectedLivePrice, selectedCurrency)} (auto-filled)`
                      : `No live quote available now for ${calcSelectedSymbol}. You can still type price manually.`}
                  </p>
                ) : null}
                {selectedCalcHolding ? (
                  <p className="tools-section-note">
                    Holding average buy price: {formatCurrency(selectedCalcHoldingPrice, selectedCurrency)} | Available quantity: {selectedCalcHolding.quantity.toFixed(2)}
                  </p>
                ) : null}
              </>
            )}

            {calculatorType === 'buy' ? (
              canCalculateBuyCount ? (
                <p className="tools-calc-result">
                  You can buy <span style={{ fontWeight: 700, color: 'var(--special)' }}>{maxBuyShares}</span> share{maxBuyShares === 1 ? '' : 's'} and keep <span style={{ fontWeight: 700, color: 'var(--special)' }}>{formatCurrency(remainingMoney, selectedCurrency)}</span> remaining.
                </p>
              ) : (
                <p className="tools-calc-result">Enter valid values above to calculate your buy count.</p>
              )
            ) : sellStatusMessage ? (
              <p className="tools-calc-result">{sellStatusMessage}</p>
            ) : isSellQuantityValid ? (
              <p className="tools-calc-result">
                You would {sellGainLoss >= 0 ? 'gain' : 'lose'}{' '}
                <span style={{ fontWeight: 700, color: 'var(--special)' }}>
                  {formatCurrency(sellGainLossAbs, selectedCurrency)}
                </span>{' '}
                on this sale{sellGainLoss !== 0 ? ` (${sellGainLossPercent >= 0 ? '+' : ''}${formatPercent(sellGainLossPercent)})` : ''}.
              </p>
            ) : (
              <p className="tools-calc-result">Select a holding, enter a sell price, and enter a quantity to calculate gain or loss.</p>
            )}
          </Box>

          <Divider className="live-mode-divider" />

          <Box className="tools-section">
            <h3 className="tools-section-title">Useful Links</h3>
            <Box className="tools-links">
              <a
                href="https://finance.yahoo.com/news/"
                target="_blank"
                rel="noreferrer"
                className="tools-link"
              >
                Yahoo Finance News
              </a>
              <a
                href="https://www.tradingview.com/heatmap/stock/"
                target="_blank"
                rel="noreferrer"
                className="tools-link"
              >
                S&P 500 Heatmap
              </a>
            </Box>
          </Box>
        </Box>
      </Modal>

      {renderEditModal()}
      {renderGoalModal()}

      {isLoading ? (
        <div className="state-block" role="status" aria-live="polite">
          <CircularProgress size={28} />
          <span>Loading...</span>
        </div>
        ) : error ? (
          <Alert severity="error">{error}</Alert>
        ) : !currentUserId ? (
          <Alert severity="info">Sign in with Google to view and save your personal data.</Alert>
        ) : (
          page === 'table' ? (
          <>
            {renderSheetTable('Stocks', displayStocksData)}
            {renderSheetTable('Dividend', dividendData)}
            {renderSheetTable('Money Move', moneyMoveData)}
          </>
          ) : page === 'dataShow' ? (
            <Box>
              <Box className="user-show">
                <h1 className="user-title text-black">{currentUserName}</h1>
                <p className="user-title">{currentUserEmail}</p>
              </Box>

              <Box>
                <Box className="sheet-title-container" style={{ display: 'flex', alignItems: 'end', justifyContent: 'start', gap: '1rem', position: 'relative' }}>
                  <h1 className="sheet-title text-black">Goal</h1>
                  <Button
                    variant="outlined"
                    onClick={handleOpenGoalModalMain}
                    sx={{
                      width: 40,
                      minWidth: 40,
                      height: 40,
                      padding: 0,
                      marginBottom: '4px',
                      borderRadius: '50%',
                      backgroundColor: 'var(--bg)',
                      color: 'var(--wbg)',
                      borderColor: 'var(--wbg)',
                      '&:hover': {
                        borderColor: 'var(--accent)',
                        opacity: 0.92,
                      },
                    }}
                  >
                    <MdOutlineAdd size={20} />
                  </Button>
                </Box>

                {goalProgressRows.length > 0 && (
                  <Box className="goal-main-scroll-wrap">
                    <Box className="goal-main-scroll">
                      {goalProgressRows.map((goal) => (
                        <Box
                          key={goal.id}
                          className={`goal-main-card ${viewingGoalIdMain === goal.id ? ' goal-main-card-selected' : ''}`}
                          onClick={() =>
                            setViewingGoalIdMain((currentGoalId) =>
                              currentGoalId === goal.id ? null : goal.id,
                            )
                          }
                        >
                          <Box className="goal-main-card-head">
                            <p className="summary-label">{getGoalDisplayTitle(goal)} | {getGoalValue(goal, goal.displayCurrentValue, goal.targetValue, goal.targetCurrency)}</p>
                            <Button
                              variant="text"
                              className="live-unfollow-button"
                              onClick={(event) => {
                                event.stopPropagation()
                                if (viewingGoalIdMain === goal.id) {
                                  setViewingGoalIdMain(null)
                                }
                                handleDeleteGoal(goal.id)
                              }}
                              aria-label={`Delete goal ${getGoalDisplayTitle(goal)}`}
                              sx={{ minWidth: 32, width: 32, height: 32, borderRadius: '50%' }}
                            >
                              <MdCancel size={18} />
                            </Button>
                          </Box>
                          <p className="summary-value" style={{ color: goal.completion === 100 ? '#16a34a' : '' }}>
                            {/* {getGoalValueSummary(goal, goal.displayCurrentValue, goal.targetValue, goal.targetCurrency)} */}
                            {goal.completion.toFixed(0)}%
                          </p>
                          {/* <p className="goal-main-progress">{goal.completion.toFixed(1)}% complete</p> */}
                        </Box>
                      ))}
                    </Box>
                  </Box>
                )}

                {goalProgressRows.length === 0 && (
                  <p style={{ color: 'var(--text-secondary)', marginBottom: '2rem' }}>No goals yet. Click + to create one.</p>
                )}

                {selectedMainGoalProgress ? (
                  <Box className="dashboard-layout goal-main-detail-grid">
                    <Box className="chart-card goal-main-gauge-card">
                      <h2 className="chart-title">Selected Goal Progress</h2>
                      <div className="chart-canvas chart-canvas-center">
                        <Gauge
                          width={300}
                          height={230}
                          value={selectedMainGoalProgress.completion}
                          valueMin={0}
                          valueMax={100}
                          sx={{ '& .MuiGauge-valueArc': { fill: 'var(--special)' } }}
                        />
                      </div>
                      <p className="goal-gauge-value">{selectedMainGoalProgress.completion.toFixed(2)}%</p>
                      <p className="goal-gauge-label">{getGoalDisplayTitle(selectedMainGoalProgress)}</p>
                      <p className="goal-gauge-detail">
                        {getGoalValueSummary(
                          selectedMainGoalProgress,
                          selectedMainGoalProgress.displayCurrentValue,
                          selectedMainGoalProgress.targetValue,
                          selectedMainGoalProgress.targetCurrency,
                        )}
                      </p>
                    </Box>

                    {selectedMainGoalProgress.scheduleType === 'frequency' ? (
                      <Box className="chart-card goal-main-history-card">
                        <h2 className="chart-title">Goal Completion History (Monthly)</h2>
                        <div className="chart-canvas">
                          {selectedMainGoalMonthlyCompletionHistory.labels.length === 0 ? (
                            <p className="chart-empty">No monthly data available for this goal yet.</p>
                          ) : (
                            <BarChart
                              height={280}
                              xAxis={[{ scaleType: 'band', data: selectedMainGoalMonthlyCompletionHistory.labels }]}
                              yAxis={[{ label: 'Completion %', max: 100 }]}
                              series={[
                                {
                                  label: `Completion % (target ${formatGoalValue(selectedMainGoalProgress.targetValue, selectedMainGoalProgress.metric, selectedMainGoalProgress.targetCurrency)})`,
                                  color: '#ff8447',
                                  data: selectedMainGoalMonthlyCompletionHistory.periodCompletionValues,
                                },
                              ]}
                            />
                          )}
                        </div>
                      </Box>
                    ) : (
                      <Box className="chart-card goal-main-history-card">
                        <Box className="goal-timeline-header">
                          <h2 className="chart-title">Goal Timeline</h2>
                          <Button
                            variant="outlined"
                            size="small"
                            onClick={() => setShowGoalTimelineAsPercent(!showGoalTimelineAsPercent)}
                            sx={{
                              textTransform: 'none',
                              borderRadius: '8px',
                              fontSize: '0.875rem',
                              height: '32px',
                              color: 'var(--special)',
                              borderColor: 'var(--special)',
                            }}
                          >
                            {showGoalTimelineAsPercent ? 'Value' : '%'}
                          </Button>
                        </Box>
                        <div className="chart-canvas">
                          {selectedMainGoalMonthlyCompletionHistory.labels.length === 0 ||
                          (showGoalTimelineAsPercent &&
                            (!selectedMainGoalMonthlyCompletionHistory.timelineCompletionValues ||
                              selectedMainGoalMonthlyCompletionHistory.timelineCompletionValues.length === 0)) ||
                          (!showGoalTimelineAsPercent &&
                            (!selectedMainGoalMonthlyCompletionHistory.timelineRawValues ||
                              selectedMainGoalMonthlyCompletionHistory.timelineRawValues.length === 0)) ? (
                            <p className="chart-empty">No timeline data available for this goal yet.</p>
                          ) : (
                            <LineChart
                              height={280}
                              series={[
                                {
                                  data: showGoalTimelineAsPercent
                                    ? selectedMainGoalMonthlyCompletionHistory.timelineCompletionValues
                                    : selectedMainGoalMonthlyCompletionHistory.timelineRawValues,
                                  label: showGoalTimelineAsPercent
                                    ? `Completion % (target ${formatGoalValue(selectedMainGoalProgress.targetValue, selectedMainGoalProgress.metric, selectedMainGoalProgress.targetCurrency)})`
                                    : `Value (target ${formatGoalValue(selectedMainGoalProgress.targetValue, selectedMainGoalProgress.metric, selectedMainGoalProgress.targetCurrency)})`,
                                  showMark: false,
                                },
                              ]}
                              xAxis={[
                                {
                                  scaleType: 'point',
                                  data: selectedMainGoalMonthlyCompletionHistory.labels,
                                },
                              ]}
                              yAxis={[{ label: showGoalTimelineAsPercent ? 'Completion %' : `Value (${selectedMainGoalMonthlyCompletionHistory.metricUnit})`, min: 0, max: showGoalTimelineAsPercent ? 100 : undefined }]}
                            />
                          )}
                        </div>
                      </Box>
                    )}

                    <Box className="chart-card goal-main-meta-card">
                      <h2 className="chart-title">Goal Details</h2>
                      <div className="chart-canvas">
                        <p className="chart-empty">
                          {`Metric: ${getGoalMetricLabel(selectedMainGoalProgress.metric)}`}
                        </p>
                        <p className="chart-empty">
                          {`Target: ${formatGoalValue(selectedMainGoalProgress.targetValue, selectedMainGoalProgress.metric, selectedMainGoalProgress.targetCurrency)}`}
                        </p>
                        <p className="chart-empty">
                          {selectedMainGoalProgress.scheduleType === 'deadline'
                            ? `Schedule: Deadline ${getGoalScheduleLabel(selectedMainGoalProgress)}`
                            : `Schedule: Frequency ${getGoalScheduleLabel(selectedMainGoalProgress)}`}
                        </p>
                      </div>
                    </Box>
                  </Box>
                ) : null}
              </Box>

              <h1 className="sheet-title text-black">Data</h1>
              <Box className="dashboard-layout">
                <section className="summary-grid" aria-label="Portfolio summary">
                  <Box className="summary-card">
                    <p className="summary-label">Total</p>
                    <p className="summary-value">{formatCurrency(summary.totalMoney, selectedCurrency)}</p>
                  </Box>
                  <Box className="summary-card">
                    <p className="summary-label">Borrowed</p>
                    <p className="summary-value">{formatCurrency(summary.borrowed, selectedCurrency)}</p>
                  </Box>
                  <Box className="summary-card">
                    <p className="summary-label">Dividend</p>
                    <p className="summary-value">{formatCurrency(summary.dividendTotal, selectedCurrency)}</p>
                  </Box>
                  <Box className="summary-card">
                    <p className="summary-label">% (Average)</p>
                    <p className="summary-value">{formatPercent(summary.percentAverage)}</p>
                  </Box>
                  <Box className="summary-card summary-card-large">
                    <p className="summary-label">Earn (All / Per Trade)</p>
                    <p className="summary-value">
                      {formatCurrency(summary.earnAll, selectedCurrency)} / {formatCurrency(summary.earnPerTrade, selectedCurrency)}
                    </p>
                  </Box>
                </section>

                <Box className="chart-card chart-card-timeline">
                  <h2 className="chart-title">Money Timeline (Total vs Working Capital)</h2>
                  <div className="chart-canvas">
                    {moneyTimelineChart.labels.length === 0 ? (
                      <p className="chart-empty">No timeline data available to display.</p>
                    ) : (
                      <LineChart
                        height={300}
                        series={[
                          {
                            data: moneyTimelineChart.totalMoneyValues,
                            label: 'Total Money',
                            showMark: false,
                          },
                          {
                            data: moneyTimelineChart.moneyIHaveValues,
                            label: 'Money I Have',
                            showMark: false,
                          },
                        ]}
                        xAxis={[
                          {
                            scaleType: 'point',
                            data: moneyTimelineChart.labels,
                            valueFormatter: (value, context) =>
                              context.location === 'tick'
                                ? formatCompactAxisDate(String(value))
                                : String(value),
                          }
                        ]}
                      />
                    )}
                  </div>
                </Box>

                <Box className="chart-card chart-card-allocation">
                  <h2 className="chart-title">Current Holdings Allocation</h2>
                  <div className="chart-canvas chart-canvas-center">
                    {holdingsDonutChart.length === 0 ? (
                      <p className="chart-empty">No holdings allocation data available.</p>
                    ) : (
                      <PieChart
                        height={230}
                        width={300}
                        series={[
                          {
                            innerRadius: 65,
                            outerRadius: 110,
                            paddingAngle: 2,
                            cornerRadius: 4,
                            data: holdingsDonutChart,
                          },
                        ]}
                      />
                    )}
                  </div>
                </Box>

                <Box className="chart-card chart-card-monthly">
                  <h2 className="chart-title">Monthly Earn / Loss</h2>
                  <div className="chart-canvas">
                    {monthlyEarnChart.labels.length === 0 ? (
                      <p className="chart-empty">No monthly earn/loss data available.</p>
                    ) : (
                      <BarChart
                        height={280}
                        xAxis={[
                          {
                            scaleType: 'band',
                            data: monthlyEarnChart.labels.map((label) => formatMonthLabel(label)),
                          },
                        ]}
                        yAxis={[
                          {
                            label: `Earn / Loss (${selectedCurrency})`,
                          },
                        ]}
                        series={[
                          {
                            label: 'Earn / Loss',
                            color: '#16a34a',
                            data: monthlyEarnChart.values,
                          },
                        ]}
                      />
                    )}
                  </div>
                </Box>

                <Box className="sheet-section dashboard-holdings">
                  {/* <div className="holdings-header"> */}
                    <h1 className="sheet-title text-black">Current Holdings</h1>
                    {/* <p className="holdings-status">
                      {isUsMarketOpen === true
                        ? `Market Open${isQuoteLoading ? ' - Updating...' : ''}`
                        : isUsMarketOpen === false
                          ? 'Market Closed'
                          : 'Market Status Unknown'}
                      {quoteUpdatedAt ? ` | Last update: ${new Date(quoteUpdatedAt).toLocaleTimeString()}` : ''}
                    </p> */}
                  {/* </div> */}

                  <TableContainer component={Paper} elevation={0} className="table-shell">
                    <Table size="small" aria-label="current holdings table">
                      <TableHead>
                        <TableRow>
                          <TableCell sx={{ fontWeight: 700 }}>Stock Name/Number</TableCell>
                          <TableCell sx={{ fontWeight: 700 }}>Quantity</TableCell>
                          <TableCell sx={{ fontWeight: 700 }}>Now Price</TableCell>
                          <TableCell sx={{ fontWeight: 700 }}>Buy In Price</TableCell>
                          <TableCell sx={{ fontWeight: 700 }}>$ Earn/Loss</TableCell>
                          <TableCell sx={{ fontWeight: 700 }}>% Change</TableCell>
                        </TableRow>
                      </TableHead>
                      <TableBody>
                        {holdings.length === 0 ? (
                          <TableRow>
                            <TableCell colSpan={6}>No current holdings found.</TableCell>
                          </TableRow>
                        ) : (
                          holdings.map((holding) => {
                            const nowPriceUsd = resolveLiveQuotePrice(liveQuotes[holding.symbol])
                            const nowPrice = Number.isFinite(nowPriceUsd)
                              ? convertAmount(nowPriceUsd, 'USD', selectedCurrency, currencyRates)
                              : Number.NaN
                            const buyInPrice = convertAmount(
                              holding.avgBuyPriceUsd,
                              'USD',
                              selectedCurrency,
                              currencyRates,
                            )

                            const earnLoss = Number.isFinite(nowPrice)
                              ? (nowPrice - buyInPrice) * holding.quantity
                              : Number.NaN
                            const changePercent = buyInPrice > 0 && Number.isFinite(nowPrice)
                              ? ((nowPrice - buyInPrice) / buyInPrice) * 100
                              : Number.NaN

                            return (
                              <TableRow key={holding.symbol}>
                                <TableCell>{holding.displayName}</TableCell>
                                <TableCell>{holding.quantity.toFixed(2)}</TableCell>
                                <TableCell>
                                  {Number.isFinite(nowPrice)
                                    ? formatCurrency(nowPrice, selectedCurrency)
                                    : '-'}
                                </TableCell>
                                <TableCell>{formatCurrency(buyInPrice, selectedCurrency)}</TableCell>
                                <TableCell
                                  className={
                                    Number.isFinite(earnLoss)
                                      ? earnLoss >= 0
                                        ? 'value-positive'
                                        : 'value-negative'
                                      : undefined
                                  }
                                >
                                  {Number.isFinite(earnLoss)
                                    ? formatCurrency(earnLoss, selectedCurrency)
                                    : '-'}
                                </TableCell>
                                <TableCell
                                  className={
                                    Number.isFinite(changePercent)
                                      ? changePercent >= 0
                                        ? 'value-positive'
                                        : 'value-negative'
                                      : undefined
                                  }
                                >
                                  {Number.isFinite(changePercent) ? formatPercent(changePercent) : '-'}
                                </TableCell>
                              </TableRow>
                            )
                          })
                        )}
                      </TableBody>
                    </Table>
                  </TableContainer>
                </Box>
              </Box>
            </Box>
          ) : page === 'goal' ? (
            <Box className="goal-layout">
              <Box className="goal-header">
                <Box>
                  <h1 className="sheet-title text-black">Goal</h1>
                  <p className="goal-subtitle">
                    Set multiple targets and track completion with a live gauge and bar view.
                  </p>
                </Box>
                {selectedGoalProgress ? (
                  <Box className="goal-highlight">
                    <p className="goal-highlight-label">{getGoalDisplayTitle(selectedGoalProgress.goal)}</p>
                    <p className="goal-highlight-value">
                      {getGoalValueSummary(
                        selectedGoalProgress.goal,
                        selectedGoalProgress.displayCurrentValue,
                        selectedGoalProgress.goal.targetValue,
                        selectedGoalProgress.goal.targetCurrency,
                      )}
                    </p>
                  </Box>
                ) : null}
              </Box>

              <Box className="goal-grid">
                <Paper elevation={0} className="goal-card goal-form-card">
                  <h2 className="chart-title">Create Goal</h2>
                  <Box component="form" className="goal-form-grid" onSubmit={handleSubmitGoal}>
                    {renderGoalFields()}
                    <Box className="add-form-actions">
                      <Button
                        type="submit"
                        variant="contained"
                        disabled={isSubmittingGoal || !currentUserId}
                        className={isConfirmingGoal ? 'is-confirming' : ''}
                      >
                        {isSubmittingGoal
                          ? 'Saving...'
                          : isConfirmingGoal
                            ? 'Confirm'
                            : currentUserId
                              ? 'Save Goal'
                              : 'Sign in required'}
                      </Button>
                    </Box>
                    {goalMessage ? <Alert severity={goalMessage.type}>{goalMessage.text}</Alert> : null}
                  </Box>
                </Paper>

                <Paper elevation={0} className="goal-card goal-gauge-card">
                  <h2 className="chart-title">Selected Goal Progress</h2>
                  {selectedGoalProgress ? (
                    <>
                      <Box className="goal-gauge-wrap">
                        <Gauge
                          width={240}
                          height={240}
                          value={selectedGoalProgress.completion}
                          valueMin={0}
                          valueMax={100}
                          sx={{ '& .MuiGauge-valueArc': { fill: 'var(--special)' } }}
                        />
                      </Box>
                      <p className="goal-gauge-value">{selectedGoalProgress.completion.toFixed(1)}%</p>
                      <p className="goal-gauge-label">{getGoalDisplayTitle(selectedGoalProgress.goal)}</p>
                      <p className="goal-gauge-detail">
                        {getGoalValueSummary(
                          selectedGoalProgress.goal,
                          selectedGoalProgress.displayCurrentValue,
                          selectedGoalProgress.goal.targetValue,
                          selectedGoalProgress.goal.targetCurrency,
                        )}
                      </p>
                      <p className="goal-gauge-meta">
                        {getGoalMetricLabel(selectedGoalProgress.goal.metric)} |{' '}
                        {selectedGoalProgress.goal.scheduleType === 'deadline'
                          ? `Deadline ${getGoalScheduleLabel(selectedGoalProgress.goal)}`
                          : `Frequency ${getGoalScheduleLabel(selectedGoalProgress.goal)}`}
                      </p>
                    </>
                  ) : (
                    <p className="chart-empty">Add a goal to see progress here.</p>
                  )}
                </Paper>

                <Paper elevation={0} className="goal-card goal-bar-card">
                  <h2 className="chart-title">Goal Completion Bars</h2>
                  {goalBarChartData.labels.length === 0 ? (
                    <p className="chart-empty">No goals yet.</p>
                  ) : (
                    <BarChart
                      height={260}
                      xAxis={[{ scaleType: 'band', data: goalBarChartData.labels }]}
                      yAxis={[{ label: 'Completion %', max: 100 }]}
                      series={[{ label: 'Completion %', data: goalBarChartData.values, color: '#ff8447' }]}
                    />
                  )}
                </Paper>

                <Paper elevation={0} className="goal-card goal-list-card">
                  <h2 className="chart-title">All Goals</h2>
                  {goalProgressRows.length === 0 ? (
                    <p className="chart-empty">No goals saved yet.</p>
                  ) : (
                    <Box className="goal-list">
                      {goalProgressRows.map((goal) => (
                        <Box
                          key={goal.id}
                          className={`goal-item${goal.id === selectedGoalId ? ' goal-item-selected' : ''}`}
                          onClick={() => handleSelectGoal(goal.id)}
                        >
                          <Box className="goal-item-main">
                            <p className="goal-item-name">{getGoalDisplayTitle(goal)}</p>
                            <p className="goal-item-meta">
                              {getGoalMetricLabel(goal.metric)} |{' '}
                              {goal.scheduleType === 'deadline'
                                ? `Deadline ${getGoalScheduleLabel(goal)}`
                                : `Frequency ${getGoalScheduleLabel(goal)}`}
                            </p>
                            <p className="goal-item-value">
                              {getGoalValueSummary(goal, goal.displayCurrentValue, goal.targetValue, goal.targetCurrency)}
                            </p>
                            <p className="goal-item-progress">{goal.completion.toFixed(1)}% complete</p>
                          </Box>
                          <Button
                            variant="text"
                            className="live-unfollow-button"
                            onClick={(event) => {
                              event.stopPropagation()
                              handleDeleteGoal(goal.id)
                            }}
                            aria-label={`Delete goal ${getGoalDisplayTitle(goal)}`}
                            sx={{ minWidth: 36, width: 36, height: 36, borderRadius: '50%' }}
                          >
                            <MdCancel size={20} />
                          </Button>
                        </Box>
                      ))}
                    </Box>
                  )}
                </Paper>
              </Box>
            </Box>

          ) : page === 'live' ? (
            <Box>
              <h1 className="sheet-title text-black">Live US Stocks</h1>
              <Box className="live-layout">
                <Paper elevation={0} className="add-form-card live-controls-card">
                  <Box className="live-controls-row">
                    <Autocomplete
                      freeSolo
                      options={followSymbolOptions}
                      fullWidth
                      inputValue={followSymbolInput}
                      onInputChange={(_, value) => setFollowSymbolInput(value.toUpperCase())}
                      renderInput={(params) => (
                        <TextField
                          {...params}
                          label="Search symbol to follow"
                          variant="standard"
                          helperText="No comma or space. Example: NVDA"
                        />
                      )}
                    />
                    <Button
                      variant="contained"
                      onClick={handleAddFollowSymbol}
                      className="live-action-button"
                    >
                      Add
                    </Button>
                  </Box>

                  <Box className="live-mode-row">
                    <p className="holdings-status">
                      {`Market: ${
                        isUsMarketOpen === true
                          ? 'Open'
                          : isUsMarketOpen === false
                            ? 'Closed'
                            : 'Unknown'
                      }`}
                    </p>

                    <Divider orientation="vertical" flexItem className="live-mode-divider" />

                    <p className="holdings-status">
                      {liveQuotesUpdatedAt
                        ? `Last update: ${new Date(liveQuotesUpdatedAt).toLocaleTimeString()}`
                        : ''}
                    </p>

                    <Divider orientation="vertical" flexItem className="live-mode-divider" />

                    <p className="holdings-status">
                      {isLiveQuotesLoading ? 'Updating quotes...' : 'Quotes ready'}
                    </p>

                    <Divider orientation="vertical" flexItem className="live-mode-divider" />

                    <p className="holdings-status">
                      {`Interval: ${
                        isUsMarketOpen === true
                          ? isFastMode
                            ? '1 min (fast mode)'
                            : '5 min (normal mode)'
                          : '5 min'
                      }`}
                    </p>

                    <Box className="live-fast-mode-toggle">
                      <p className="holdings-status">Fast mode (1 min)</p>
                      <Switch
                        checked={isFastMode}
                        onChange={(_, checked) => setIsFastMode(checked)}
                        size="small"
                      />
                    </Box>
                  </Box>
                </Paper>

                <TableContainer component={Paper} elevation={0} className="table-shell">
                  <Table size="small" aria-label="live us stocks table">
                    <TableHead>
                      <TableRow>
                        <TableCell sx={{ fontWeight: 700 }}>Stocks</TableCell>
                        <TableCell sx={{ fontWeight: 700 }}>Current</TableCell>
                        <TableCell sx={{ fontWeight: 700 }}>Change</TableCell>
                        <TableCell sx={{ fontWeight: 700 }}>% change</TableCell>
                        <TableCell sx={{ fontWeight: 700 }}>Highest</TableCell>
                        <TableCell sx={{ fontWeight: 700 }}>Lowest</TableCell>
                        <TableCell sx={{ fontWeight: 700 }}>Open</TableCell>
                        <TableCell sx={{ fontWeight: 700 }}>Prev Close</TableCell>
                        <TableCell sx={{ fontWeight: 700 }}>Action</TableCell>
                      </TableRow>
                    </TableHead>
                    <TableBody>
                      {liveSymbols.length === 0 ? (
                        <TableRow>
                          <TableCell colSpan={9}>No symbols selected.</TableCell>
                        </TableRow>
                      ) : (
                        liveSymbols.map((symbol) => {
                          const quote = liveQuotes[symbol] ?? {}
                          const changeValue = quote.d
                          const percentValue = quote.dp
                          const isRequiredHolding = requiredHoldingSymbolSet.has(symbol)

                          return (
                            <TableRow key={symbol}>
                              <TableCell>
                                <a
                                  href={`https://finance.yahoo.com/quote/${encodeURIComponent(symbol)}/`}
                                  target="_blank"
                                  rel="noreferrer"
                                  className="live-symbol-link"
                                >
                                  {symbol}
                                </a>
                              </TableCell>
                              <TableCell>{formatLiveQuoteValue(quote.c, 'price')}</TableCell>
                              <TableCell
                                className={
                                  typeof changeValue === 'number' && Number.isFinite(changeValue)
                                    ? changeValue >= 0
                                      ? 'value-positive'
                                      : 'value-negative'
                                    : undefined
                                }
                              >
                                {formatLiveQuoteValue(changeValue, 'change')}
                              </TableCell>
                              <TableCell
                                className={
                                  typeof percentValue === 'number' && Number.isFinite(percentValue)
                                    ? percentValue >= 0
                                      ? 'value-positive'
                                      : 'value-negative'
                                    : undefined
                                }
                              >
                                {formatLiveQuoteValue(percentValue, 'percent')}
                              </TableCell>
                              <TableCell>{formatLiveQuoteValue(quote.h, 'price')}</TableCell>
                              <TableCell>{formatLiveQuoteValue(quote.l, 'price')}</TableCell>
                              <TableCell>{formatLiveQuoteValue(quote.o, 'price')}</TableCell>
                              <TableCell>{formatLiveQuoteValue(quote.pc, 'price')}</TableCell>
                              <TableCell>
                                {isRequiredHolding ? (
                                  <span className="live-required-symbol">Holding</span>
                                ) : (
                                  <Button
                                    variant="text"
                                    className="live-unfollow-button"
                                    onClick={() => handleUnfollowSymbol(symbol)}
                                    aria-label={`Unfollow ${symbol}`}
                                  >
                                    <MdCancel size={20} />
                                  </Button>
                                )}
                              </TableCell>
                            </TableRow>
                          )
                        })
                      )}
                    </TableBody>
                  </Table>
                </TableContainer>
              </Box>

              <h1 className="sheet-title text-black">Price Alerts</h1>
              <Box className="live-layout">
                <Paper elevation={0} className="add-form-card live-alert-card">
                  <Box className="live-alert-form">
                    <TextField
                      select
                      variant="standard"
                      label="Stock"
                      value={alertSymbol}
                      onChange={(event) => setAlertSymbol(event.target.value.toUpperCase())}
                      className="live-alert-field"
                    >
                      {liveSymbols.map((symbol) => (
                        <MenuItem key={`alert-symbol-${symbol}`} value={symbol}>
                          {symbol}
                        </MenuItem>
                      ))}
                    </TextField>

                    <TextField
                      variant="standard"
                      label="Target price"
                      type="number"
                      value={alertPriceInput}
                      onChange={(event) => setAlertPriceInput(event.target.value)}
                      className="live-alert-field"
                      inputProps={{ step: '0.01', min: '0' }}
                    />

                    <TextField
                      select
                      variant="standard"
                      label="Condition"
                      value={alertCondition}
                      onChange={(event) => setAlertCondition(event.target.value as PriceAlertCondition)}
                      className="live-alert-field"
                    >
                      <MenuItem value="above">Above</MenuItem>
                      <MenuItem value="below">Below</MenuItem>
                    </TextField>

                    <Button variant="contained" onClick={handleAddPriceAlert} className="live-action-button">
                      Add Alert
                    </Button>
                  </Box>

                  <TableContainer component={Paper} elevation={0} className="table-shell live-alert-table">
                    <Table size="small" aria-label="price alerts table">
                      <TableHead>
                        <TableRow>
                          <TableCell sx={{ fontWeight: 700 }}>Stock</TableCell>
                          <TableCell sx={{ fontWeight: 700 }}>Condition</TableCell>
                          <TableCell sx={{ fontWeight: 700 }}>Target</TableCell>
                          <TableCell sx={{ fontWeight: 700 }}>Enabled</TableCell>
                          <TableCell sx={{ fontWeight: 700 }}>Action</TableCell>
                        </TableRow>
                      </TableHead>
                      <TableBody>
                        {priceAlerts.length === 0 ? (
                          <TableRow>
                            <TableCell colSpan={5}>No alerts set.</TableCell>
                          </TableRow>
                        ) : (
                          priceAlerts.map((alert) => (
                            <TableRow key={alert.id}>
                              <TableCell>{alert.symbol}</TableCell>
                              <TableCell>{alert.condition === 'above' ? 'Above' : 'Below'}</TableCell>
                              <TableCell>{alert.targetPrice.toFixed(2)}</TableCell>
                              <TableCell>
                                <Switch
                                  checked={alert.enabled}
                                  onChange={() => handleTogglePriceAlert(alert.id)}
                                  size="small"
                                />
                              </TableCell>
                              {/* ewL */}
                              <TableCell>
                                <Button
                                  variant="text"
                                  className="live-unfollow-button"
                                  onClick={() => handleDeletePriceAlert(alert.id)}
                                  sx={{ minWidth: 36, width: 36, height: 36, borderRadius: '50%' }}
                                >
                                  <MdCancel size={20} />
                                </Button>
                              </TableCell>
                              {/* <TableCell>
                                {isRequiredHolding ? (
                                  <span className="live-required-symbol">Holding</span>
                                ) : (
                                  <Button
                                    variant="text"
                                    className="live-unfollow-button"
                                    onClick={() => handleUnfollowSymbol(symbol)}
                                    aria-label={`Unfollow ${symbol}`}
                                  >
                                    <MdCancel size={20} />
                                  </Button>
                                )}
                              </TableCell> */}
                            </TableRow>
                          ))
                        )}
                      </TableBody>
                    </Table>
                  </TableContainer>
                </Paper>
              </Box>
            </Box>
          ) : page === 'add' ? (
            <Box className="add-layout">
              <h1 className='sheet-title text-black'>Add data</h1>
              <Paper elevation={0} className="add-form-card">
                <Box component="form" className="add-form-grid" onSubmit={handleSubmitAdd}>
                  <Box className="add-table-switcher" role="group" aria-label="Choose table">
                    {/* <p className="add-table-label">Table</p> */}
                    <Button
                      type="button"
                      className="add-table-button"
                      variant="text"
                      onClick={handleCycleAddTable}
                    >
                      {selectedAddTable}
                    </Button>
                    <p className='add-table-label'>* click the table name to switch tables</p>
                  </Box>

                  {selectedAddFields.map((field) => {
                    const pairedFieldKey = selectedAddTable === 'Dividend' ? 'div' : 'price'

                    if (selectedAddTable === 'Stocks' && field.key === 'stock') {
                      const actionField = selectedAddFields.find((candidate) => candidate.key === 'action')
                      const quantityField = selectedAddFields.find((candidate) => candidate.key === 'quantity')

                      return (
                        <Box className="add-form-triple-row" key="stocks-main-row">
                          {renderAddField(field, 'add-field-stock')}
                          {actionField ? renderAddField(actionField, 'add-field-action') : null}
                          {quantityField ? renderAddField(quantityField, 'add-field-quantity') : null}
                        </Box>
                      )
                    }

                    if (selectedAddTable === 'Stocks' && (field.key === 'action' || field.key === 'quantity')) {
                      return null
                    }

                    if (selectedAddTable === 'Money Move' && field.key === 'name') {
                      const doField = selectedAddFields.find((candidate) => candidate.key === 'do')

                      return (
                        <Box className="add-form-name-do-row" key="money-move-name-do-row">
                          {renderAddField(field, 'add-field-name')}
                          {doField ? renderAddField(doField, 'add-field-do') : null}
                        </Box>
                      )
                    }

                    if (selectedAddTable === 'Money Move' && field.key === 'do') {
                      return null
                    }

                    if (field.key === pairedFieldKey) {
                      return null
                    }

                    if (field.key === 'currency') {
                      const pairedField = selectedAddFields.find((candidate) => candidate.key === pairedFieldKey)

                      return (
                        <Box className="add-form-pair-row" key={`pair-${selectedAddTable}-${pairedFieldKey}`}>
                          {renderAddField(field, 'add-field-currency')}
                          {pairedField ? renderAddField(pairedField, 'add-field-value') : null}
                        </Box>
                      )
                    }

                    return renderAddField(field)
                  })}

                  <Box className="add-form-actions">
                    <Button
                      type="submit"
                      variant="contained"
                      disabled={isSubmittingAdd || !currentUserId}
                      className={isConfirmingAdd ? 'is-confirming' : ''}
                    >
                      {isSubmittingAdd ? 'Adding...' : isConfirmingAdd ? 'Conform' : currentUserId ? 'Add Row' : 'Sign in required'}
                    </Button>
                  </Box>

                  {addFormMessage ? (
                    <Alert severity={addFormMessage.type}>{addFormMessage.text}</Alert>
                  ) : null}
                </Box>
              </Paper>

              <h1 className='sheet-title text-black'>Upload via Excel</h1>
              <Paper elevation={0} className="add-form-card">
                <Box className="excel-section">
                  <Box className="excel-actions" style={{ display: 'flex', gap: '1rem', marginBottom: '1rem' }}>
                    <Button
                      component="label"
                      variant="outlined"
                      sx={{ ...excelActionButtonSx, maxWidth: 280, justifyContent: 'flex-start' }}
                      disabled={isUploadingExcel || !currentUserId}
                      startIcon={<MdTableRows />}
                    >
                      {selectedExcelFile ? selectedExcelFile.name : 'Choose Excel File'}
                      <input
                        ref={excelFileInputRef}
                        hidden
                        accept=".xlsx,.xls"
                        type="file"
                        onChange={handleExcelFileChange}
                        disabled={isUploadingExcel || !currentUserId}
                      />
                    </Button>
                    <Button
                      variant="outlined"
                      sx={{ ... excelActionButtonSx, backgroundColor: 'var(--special)', color: 'var(--text)', '&:hover': { backgroundColor: 'var(--accent)' } }}
                      onClick={handleDownloadExcel}
                      disabled={!currentUserId}
                      startIcon={<MdDataUsage />}
                    >
                      Download Template
                    </Button>
                  </Box>
                  {excelUploadMessage ? (
                    <Alert severity={excelUploadMessage.type}>{excelUploadMessage.text}</Alert>
                  ) : null}

                  {!selectedExcelFile ? (
                    <div style={{ fontSize: '0.875rem', color: '#666', marginTop: '1rem' }}>
                      <p>
                        <strong>Instructions:</strong>
                      </p>
                      <ul style={{ marginTop: '0.5rem', marginBottom: '0.5rem', textAlign: 'left' }}>
                        <li>Download the template to see the required format</li>
                        <li>Fill in your data in the Stocks, Dividend, and Money Move sheets</li>
                        <li>Upload the file to import all rows to the database at once</li>
                      </ul>
                    </div>
                  ) : (
                    <Box className="excel-preview-panel">
                      {isParsingExcelPreview ? (
                        <Box className="state-block">
                          <CircularProgress size={20} />
                          <span>Reading Excel file preview...</span>
                        </Box>
                      ) : (
                        <>
                          <ToggleButtonGroup
                            value={selectedExcelPreviewTable}
                            exclusive
                            onChange={(_, value: AddTableName | null) => {
                              if (value) {
                                setSelectedExcelPreviewTable(value)
                              }
                            }}
                            size="small"
                            className="excel-toggle-group"
                          >
                            <ToggleButton value="Stocks">Stocks ({excelPreviewData?.stocks.length ?? 0})</ToggleButton>
                            <ToggleButton value="Dividend">Dividend ({excelPreviewData?.dividend.length ?? 0})</ToggleButton>
                            <ToggleButton value="Money Move">Money Move ({excelPreviewData?.moneyMove.length ?? 0})</ToggleButton>
                          </ToggleButtonGroup>

                          <TableContainer component={Paper} elevation={0} className="table-shell excel-preview-table">
                            <Table size="small" stickyHeader>
                              <TableHead>
                                <TableRow>
                                  {EXCEL_PREVIEW_HEADERS[selectedExcelPreviewTable].map((header) => (
                                    <TableCell key={`${selectedExcelPreviewTable}-${header}`}>{header}</TableCell>
                                  ))}
                                </TableRow>
                              </TableHead>
                              <TableBody>
                                {previewRows.length === 0 ? (
                                  <TableRow>
                                    <TableCell colSpan={EXCEL_PREVIEW_HEADERS[selectedExcelPreviewTable].length}>
                                      No rows found in this sheet.
                                    </TableCell>
                                  </TableRow>
                                ) : (
                                  previewRows.map((row, rowIndex) => (
                                    <TableRow key={`${selectedExcelPreviewTable}-${rowIndex}`}>
                                      {EXCEL_PREVIEW_HEADERS[selectedExcelPreviewTable].map((header, columnIndex) => (
                                        <TableCell key={`${selectedExcelPreviewTable}-${rowIndex}-${columnIndex}`}>
                                          {formatExcelPreviewCellValue(row[columnIndex] ?? '', header)}
                                        </TableCell>
                                      ))}
                                    </TableRow>
                                  ))
                                )}
                              </TableBody>
                            </Table>
                          </TableContainer>
                        </>
                      )}
                    </Box>
                  )}

                  <Box className="add-form-actions">
                    <Button
                      type="button"
                      variant="contained"
                      disabled={isUploadingExcel || !currentUserId || !selectedExcelFile || isParsingExcelPreview}
                      className={isConfirmingExcelAdd ? 'is-confirming' : ''}
                      onClick={handleSubmitExcelUpload}
                    >
                      {isUploadingExcel
                        ? 'Adding...'
                        : isConfirmingExcelAdd
                          ? 'Conform'
                          : currentUserId
                            ? 'Add Data'
                            : 'Sign in required'}
                    </Button>
                  </Box>
                </Box>
              </Paper>
            </Box>
          ) : null
        ) 
      }

      {!isLoading &&
        !error &&
        currentUserId &&
        displayStocksData.rows.length === 0 &&
        dividendData.rows.length === 0 &&
        moneyMoveData.rows.length === 0 &&
        <Alert severity="info">No rows found in Supabase yet.</Alert>}
    </main>
  )
}

export default App
