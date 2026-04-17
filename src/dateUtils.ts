function pad(value: number): string {
  return String(Math.trunc(value)).padStart(2, '0')
}

function padMilliseconds(value: number): string {
  return String(Math.trunc(value)).padStart(3, '0')
}

function buildLocalDate(
  year: number,
  month: number,
  day: number,
  hour = 0,
  minute = 0,
  second = 0,
  millisecond = 0,
): Date | null {
  const date = new Date(year, month - 1, day, hour, minute, second, millisecond)
  return Number.isNaN(date.getTime()) ? null : date
}

export function parseDateValue(raw: string): Date | null {
  const trimmed = raw.trim()
  if (!trimmed) {
    return null
  }

  const normalizedWhitespace = trimmed.replace(/[\u00A0\u202F\u2009]/g, ' ').replace(/\s+/g, ' ').trim()

  const googleDateMatch = normalizedWhitespace.match(
    /^Date\((\d+),\s*(\d+),\s*(\d+)(?:,\s*(\d+),\s*(\d+),\s*(\d+))?\)$/i,
  )
  if (googleDateMatch) {
    const [, yearRaw, monthRaw, dayRaw, hourRaw, minuteRaw, secondRaw] = googleDateMatch
    const year = Number(yearRaw)
    const month = Number(monthRaw) + 1
    const day = Number(dayRaw)
    const hour = hourRaw ? Number(hourRaw) : 0
    const minute = minuteRaw ? Number(minuteRaw) : 0
    const second = secondRaw ? Number(secondRaw) : 0

    return buildLocalDate(year, month, day, hour, minute, second)
  }

  const meridiemMatch = normalizedWhitespace.match(/(上午|下午|\bAM\b|\bPM\b)/i)
  const meridiem = meridiemMatch?.[1]?.toUpperCase() ?? ''
  const stripped = normalizedWhitespace
    .replace(/(上午|下午|\bAM\b|\bPM\b)/gi, '')
    .replace(/\s+/g, ' ')
    .trim()

  const dateTimeMatch = stripped.match(
    /^(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})(?:[T\s]+(\d{1,2})(?::(\d{2}))?(?::(\d{2}))?(?:\.(\d{1,3}))?)?(?:\s*(?:Z|[+-]\d{2}:?\d{2}))?$/i,
  )
  if (dateTimeMatch) {
    const [, yearRaw, monthRaw, dayRaw, hourRaw, minuteRaw, secondRaw, millisecondRaw] = dateTimeMatch
    let hour = hourRaw ? Number(hourRaw) : 0
    const minute = minuteRaw ? Number(minuteRaw) : 0
    const second = secondRaw ? Number(secondRaw) : 0
    const millisecond = millisecondRaw ? Number(millisecondRaw.padEnd(3, '0')) : 0

    if (meridiem === 'PM' || meridiem === '下午') {
      if (hour < 12) {
        hour += 12
      }
    } else if (meridiem === 'AM' || meridiem === '上午') {
      if (hour === 12) {
        hour = 0
      }
    }

    // Check if the input has a UTC indicator (Z or timezone offset)
    const hasUTCIndicator = /Z|[+-]\d{2}:?\d{2}$/.test(stripped)

    if (hasUTCIndicator) {
      // Parse as UTC time
      const timestamp = Date.UTC(
        Number(yearRaw),
        Number(monthRaw) - 1,
        Number(dayRaw),
        hour,
        minute,
        second,
        millisecond,
      )
      return Number.isFinite(timestamp) ? new Date(timestamp) : null
    }

    // Parse as local time (no UTC indicator)
    return buildLocalDate(Number(yearRaw), Number(monthRaw), Number(dayRaw), hour, minute, second, millisecond)
  }

  // Try parsing as standard ISO string (handles Z suffix and offsets)
  const parsed = new Date(normalizedWhitespace)
  return Number.isNaN(parsed.getTime()) ? null : parsed
}

export function formatDateTimeLocalValue(date: Date): string {
  const year = date.getFullYear()
  const month = pad(date.getMonth() + 1)
  const day = pad(date.getDate())
  const hour = pad(date.getHours())
  const minute = pad(date.getMinutes())
  return `${year}-${month}-${day}T${hour}:${minute}`
}

export function formatDateDDMMYYYY(date: Date): string {
  const day = pad(date.getDate())
  const month = pad(date.getMonth() + 1)
  const year = date.getFullYear()
  return `${day}/${month}/${year}`
}

export function formatDateTimeWithOffset(date: Date): string {
  const year = date.getFullYear()
  const month = pad(date.getMonth() + 1)
  const day = pad(date.getDate())
  const hour = pad(date.getHours())
  const minute = pad(date.getMinutes())
  const second = pad(date.getSeconds())
  const millisecond = padMilliseconds(date.getMilliseconds())
  const offsetMinutes = -date.getTimezoneOffset()
  const offsetSign = offsetMinutes >= 0 ? '+' : '-'
  const absoluteOffsetMinutes = Math.abs(offsetMinutes)
  const offsetHours = pad(Math.floor(absoluteOffsetMinutes / 60))
  const offsetRemainderMinutes = pad(absoluteOffsetMinutes % 60)

  return `${year}-${month}-${day}T${hour}:${minute}:${second}.${millisecond}${offsetSign}${offsetHours}:${offsetRemainderMinutes}`
}

export function normalizeDateTimeInput(raw: string): string {
  const parsed = parseDateValue(raw)
  // Send as pure UTC ISO string to DB; DB will convert timestamptz automatically
  return parsed ? parsed.toISOString() : raw.trim()
}