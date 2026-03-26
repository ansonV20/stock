/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_SPREADSHEET_ID: string
  readonly VITE_STOCKS: string
  readonly VITE_DIVIDEND: string
  readonly VITE_MONEYMOVE: string
  readonly VITE_SHEETS_APPEND_ENDPOINT?: string
}

interface ImportMeta {
  readonly env: ImportMetaEnv
}
