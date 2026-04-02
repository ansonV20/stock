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
  MdDeleteOutline,
  MdBrightnessAuto,
  MdDarkMode,
  MdLightMode,
  MdCancel,
  MdEdit
} from 'react-icons/md'
import { LineChart } from '@mui/x-charts/LineChart'
import { PieChart } from '@mui/x-charts/PieChart'
import { BarChart } from '@mui/x-charts/BarChart'
import { supabase } from './supabaseClient'
import {
  ensureUserProfile,
  insertUserRow,
  loadUserSheetData,
  mutateUserRowAndReload,
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
  'https://api.currencyapi.com/v3/latest?apikey=cur_live_PJHFzvzysrami5rnf9VB8aIqKnhYcb3Fp3JP7Lvk&currencies=USD%2CHKD'
const FINNHUB_TOKEN = import.meta.env.VITE_FINNHUB_TOKEN ?? 'd6pcdu9r01qo88aim4i0d6pcdu9r01qo88aim4ig'
const FINNHUB_QUOTE_URL = 'https://finnhub.io/api/v1/quote'
const FINNHUB_MARKET_STATUS_URL = 'https://finnhub.io/api/v1/stock/market-status'
const LIVE_QUOTES_NORMAL_INTERVAL_MS = 5 * 60 * 1000
const LIVE_QUOTES_FAST_INTERVAL_MS = 60 * 1000
const MARKET_STATUS_UPDATE_INTERVAL_MS = 60 * 1000
const AUTH_REDIRECT_URL = import.meta.env.VITE_AUTH_REDIRECT_URL?.trim()
const THEME_STORAGE_KEY = 'stock_theme_mode_v1'
const NOTIFICATIONS_STORAGE_KEY = 'stock_notifications_enabled_v1'
const LIVE_SYMBOLS_STORAGE_KEY = 'stock_live_symbols_v1'
const LIVE_SYMBOLS_COOKIE_KEY = 'stock_live_symbols_cookie_v1'
const PRICE_ALERTS_STORAGE_KEY = 'stock_price_alerts_v1'
const MARKET_OPEN_NOTICE_DAY_KEY = 'stock_market_open_notice_day_v1'
const NOTIFICATION_SW_PATH = '/notification-sw.js'

type AddTableName = 'Stocks' | 'Dividend' | 'Money Move'
type PageName = 'table' | 'dataShow' | 'add' | 'live'
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

function getCurrentMarketDayKey(): string {
  const date = new Date()
  const year = date.getFullYear()
  const month = String(date.getMonth() + 1).padStart(2, '0')
  const day = String(date.getDate()).padStart(2, '0')
  return `${year}-${month}-${day}`
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
  const lastKnownMarketOpenRef = useRef<boolean | null>(null)
  const userSymbolsReadyForIdRef = useRef<string | null>(null)
  const notificationRegistrationRef = useRef<ServiceWorkerRegistration | null>(null)
  const excelFileInputRef = useRef<HTMLInputElement | null>(null)

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

    const pollMarketStatus = async () => {
      try {
        const response = await fetch(
          `${FINNHUB_MARKET_STATUS_URL}?exchange=US&token=${FINNHUB_TOKEN}`,
        )

        if (!response.ok || isCancelled) {
          return
        }

        const marketStatus = (await response.json()) as FinnhubMarketStatusResponse
        const isOpen = marketStatus.isOpen === true
        const previousOpen = lastKnownMarketOpenRef.current
        const marketDayKey = getCurrentMarketDayKey()
        const notifiedDay = safeReadStorage<string | null>(MARKET_OPEN_NOTICE_DAY_KEY, null)

        if (isOpen && (previousOpen === false || notifiedDay !== marketDayKey)) {
          sendBrowserNotification('US Market Open', 'US stock market is now open.')
          safeWriteStorage(MARKET_OPEN_NOTICE_DAY_KEY, marketDayKey)
        }

        lastKnownMarketOpenRef.current = isOpen
      } catch {
        // Ignore market status polling errors.
      }
    }

    void pollMarketStatus()
    const timer = window.setInterval(pollMarketStatus, MARKET_STATUS_UPDATE_INTERVAL_MS)

    return () => {
      isCancelled = true
      window.clearInterval(timer)
    }
  }, [notificationsEnabled])

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
    const profitLossIndex = findColumnIndex(displayStocksData.headers, ['profitloss'])
    const percentIndex = findColumnIndex(displayStocksData.headers, ['%'])

    let earnAll = 0
    let earnTradeCount = 0
    let percentSum = 0
    let percentCount = 0

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
    const totalMoney = moneyMoveNet + earnAll + dividendTotal

    return {
      totalMoney,
      borrowed,
      dividendTotal,
      earnAll,
      earnPerTrade,
      percentAverage,
    }
  }, [moneyMoveData, dividendData, displayStocksData, stocksData, selectedCurrency, currencyRates])

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
              sx={{
                display: 'flex',
                justifyContent: 'space-between',
                gap: '1rem',
                marginTop: '1.5rem',
              }}
            >
              <Button
                type="button"
                variant="contained"
                disabled={isSubmittingEdit || !currentUserId}
                className={isConfirmingDelete ? 'is-confirming' : ''}
                onClick={() => handleSubmitEditDelete('delete')}
                sx={{
                  backgroundColor: 'var(--special)',
                  color: 'var(--text)',
                  '&:hover': {
                    backgroundColor: 'var(--accent)',
                    opacity: 0.92,
                  },
                  '&:disabled': {
                    backgroundColor: 'var(--special)',
                    opacity: 0.5,
                  },
                }}
              >
                {isSubmittingEdit
                  ? 'Deleting...'
                  : isConfirmingDelete
                    ? 'Confirm Delete'
                    : 'Delete'}
              </Button>

              <Button
                type="button"
                variant="contained"
                disabled={isSubmittingEdit || !currentUserId}
                className={isConfirmingEdit ? 'is-confirming' : ''}
                onClick={() => handleSubmitEditDelete('update')}
                sx={{
                  backgroundColor: '#000000',
                  color: 'var(--text)',
                  '&:hover': {
                    backgroundColor: '#333333',
                    opacity: 0.92,
                  },
                  '&:disabled': {
                    backgroundColor: '#000000',
                    opacity: 0.5,
                  },
                }}
              >
                {isSubmittingEdit
                  ? 'Saving...'
                  : isConfirmingEdit
                    ? 'Confirm Change'
                    : 'Change'}
              </Button>
            </Box>
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

      {renderEditModal()}

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
                <Box className="summary-card summary-card-earn">
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
                              <TableCell>
                                <Button
                                  variant="outlined"
                                  onClick={() => handleDeletePriceAlert(alert.id)}
                                  sx={{ minWidth: 36, width: 36, height: 36, borderRadius: '50%' }}
                                >
                                  <MdDeleteOutline size={18} />
                                </Button>
                              </TableCell>
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
