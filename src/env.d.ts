/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_SPREADSHEET_ID: string
  readonly VITE_STOCKS: string
  readonly VITE_DIVIDEND: string
  readonly VITE_MONEYMOVE: string
}

interface ImportMeta {
  readonly env: ImportMetaEnv
}
