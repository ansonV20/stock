import * as XLSX from 'xlsx'

export interface ExcelDataRow {
  stocks: string[][]
  dividend: string[][]
  moneyMove: string[][]
}

const STOCK_HEADERS = ['Stock', 'Currency', 'Price', 'Action', 'Time', 'Quantity', 'Handling Fees']
const DIVIDEND_HEADERS = ['Stock', 'Currency', 'Div', 'Time']
const MONEY_MOVE_HEADERS = ['Name', 'Currency', 'Price', 'Time', 'Action']

/**
 * Create and download an Excel template file
 */
export function downloadExcelTemplate(): void {
  const wb = XLSX.utils.book_new()

  // Add Stocks sheet
  const stocksWs = XLSX.utils.aoa_to_sheet([STOCK_HEADERS])
  XLSX.utils.book_append_sheet(wb, stocksWs, 'stocks')

  // Add Dividend sheet
  const dividendWs = XLSX.utils.aoa_to_sheet([DIVIDEND_HEADERS])
  XLSX.utils.book_append_sheet(wb, dividendWs, 'dividend')

  // Add Money Move sheet
  const moneyMoveWs = XLSX.utils.aoa_to_sheet([MONEY_MOVE_HEADERS])
  XLSX.utils.book_append_sheet(wb, moneyMoveWs, 'money_move')

  XLSX.writeFile(wb, 'stocks_template.xlsx')
}

/**
 * Parse an Excel file and extract data from the three sheets
 */
export async function parseExcelFile(file: File): Promise<ExcelDataRow> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()

    reader.onload = (e) => {
      try {
        const data = e.target?.result
        if (!data) {
          throw new Error('Failed to read file')
        }

        const wb = XLSX.read(data, { type: 'array' })

        // Parse Stocks sheet
        const stocksWs = wb.Sheets['stocks']
        if (!stocksWs) {
          throw new Error('Missing "stocks" sheet in Excel file')
        }
        const stocksData = XLSX.utils.sheet_to_json(stocksWs, {
          header: 1,
        }) as unknown[][]
        const stocksRows = (stocksData.slice(1) as (string | number | boolean)[][]).map((row) =>
          row.map((cell) => String(cell ?? '')),
        )

        // Parse Dividend sheet
        const dividendWs = wb.Sheets['dividend']
        if (!dividendWs) {
          throw new Error('Missing "dividend" sheet in Excel file')
        }
        const dividendData = XLSX.utils.sheet_to_json(dividendWs, {
          header: 1,
        }) as unknown[][]
        const dividendRows = (dividendData.slice(1) as (string | number | boolean)[][]).map((row) =>
          row.map((cell) => String(cell ?? '')),
        )

        // Parse Money Move sheet
        const moneyMoveWs = wb.Sheets['money_move']
        if (!moneyMoveWs) {
          throw new Error('Missing "money_move" sheet in Excel file')
        }
        const moneyMoveData = XLSX.utils.sheet_to_json(moneyMoveWs, {
          header: 1,
        }) as unknown[][]
        const moneyMoveRows = (moneyMoveData.slice(1) as (string | number | boolean)[][]).map((row) =>
          row.map((cell) => String(cell ?? '')),
        )

        resolve({
          stocks: stocksRows.filter((row) => row.some((cell) => cell?.trim() !== '')),
          dividend: dividendRows.filter((row) => row.some((cell) => cell?.trim() !== '')),
          moneyMove: moneyMoveRows.filter((row) => row.some((cell) => cell?.trim() !== '')),
        })
      } catch (error) {
        reject(error)
      }
    }

    reader.onerror = () => {
      reject(new Error('Failed to read file'))
    }

    reader.readAsArrayBuffer(file)
  })
}

/**
 * Convert Excel serial date number to ISO string
 * Excel stores dates as number of days since Dec 30, 1899
 */
function excelSerialToISOString(excelSerial: number): string {
  // Excel base date: December 30, 1899
  const excelBaseDate = new Date(1899, 11, 30)
  // Convert Excel serial (days) to milliseconds and add to base date
  const jsDate = new Date(excelBaseDate.getTime() + excelSerial * 24 * 60 * 60 * 1000)
  return jsDate.toISOString()
}

/**
 * Parse various time input formats and return ISO string
 * Handles: Excel serials, date strings, formatted dates with time
 */
function parseTimeToISO(timeValue: string): string {
  if (!timeValue || !timeValue.trim()) {
    return new Date().toISOString()
  }

  const trimmed = timeValue.trim()

  // Check if it's an Excel serial number
  const asNumber = Number(trimmed)
  if (!Number.isNaN(asNumber) && asNumber > 0) {
    try {
      return excelSerialToISOString(asNumber)
    } catch (err) {
      console.warn('Failed to parse Excel serial:', trimmed, err)
    }
  }

  // Try to parse as standard date string
  try {
    const parsed = new Date(trimmed)
    if (!Number.isNaN(parsed.getTime())) {
      return parsed.toISOString()
    }
  } catch (err) {
    console.warn('Failed to parse date string:', trimmed, err)
  }

  // Fallback: return current time
  return new Date().toISOString()
}

/**
 * Convert Excel row array to object based on headers
 */
export function excelRowToObject(
  row: string[],
  headers: string[],
): Record<string, string> {
  const obj: Record<string, string> = {}
  headers.forEach((header, index) => {
    const value = row[index] ?? ''
    // Convert header to camelCase for consistency with form data
    const key = header
      .toLowerCase()
      .replace(/\s+/g, ' ')
      .split(' ')
      .map((word, i) => (i === 0 ? word : word.charAt(0).toUpperCase() + word.slice(1)))
      .join('')
    // Handle special cases
    if (key === 'time') {
      obj[key] = parseTimeToISO(String(value))
    } else if (key === 'handlingFees' && !obj[key]) {
      obj[key] = String(value ?? '0')
    } else if (key === 'div') {
      obj[key] = String(value) // Keep 'Div' as 'div' for Dividend table
    } else if (key === 'action') {
      obj['do'] = String(value) // Map 'Action' to 'do' for Money Move table
    } else {
      obj[key] = String(value)
    }
  })
  return obj
}
