# React + TypeScript + Vite

This template provides a minimal setup to get React working in Vite with HMR and some ESLint rules.

Currently, two official plugins are available:

- [@vitejs/plugin-react](https://github.com/vitejs/vite-plugin-react/blob/main/packages/plugin-react) uses [Oxc](https://oxc.rs)
- [@vitejs/plugin-react-swc](https://github.com/vitejs/vite-plugin-react/blob/main/packages/plugin-react-swc) uses [SWC](https://swc.rs/)

## React Compiler

The React Compiler is enabled on this template. See [this documentation](https://react.dev/learn/react-compiler) for more information.

Note: This will impact Vite dev & build performances.

## Expanding the ESLint configuration

If you are developing a production application, we recommend updating the configuration to enable type-aware lint rules:

```js
export default defineConfig([
  globalIgnores(['dist']),
  {
    files: ['**/*.{ts,tsx}'],
    extends: [
      // Other configs...

      // Remove tseslint.configs.recommended and replace with this
      tseslint.configs.recommendedTypeChecked,
      // Alternatively, use this for stricter rules
      tseslint.configs.strictTypeChecked,
      // Optionally, add this for stylistic rules
      tseslint.configs.stylisticTypeChecked,

      // Other configs...
    ],
    languageOptions: {
      parserOptions: {
        project: ['./tsconfig.node.json', './tsconfig.app.json'],
        tsconfigRootDir: import.meta.dirname,
      },
      // other options...
    },
  },
])
```

You can also install [eslint-plugin-react-x](https://github.com/Rel1cx/eslint-react/tree/main/packages/plugins/eslint-plugin-react-x) and [eslint-plugin-react-dom](https://github.com/Rel1cx/eslint-react/tree/main/packages/plugins/eslint-plugin-react-dom) for React-specific lint rules:

```js
// eslint.config.js
import reactX from 'eslint-plugin-react-x'
import reactDom from 'eslint-plugin-react-dom'

export default defineConfig([
  globalIgnores(['dist']),
  {
    files: ['**/*.{ts,tsx}'],
    extends: [
      // Other configs...
      // Enable lint rules for React
      reactX.configs['recommended-typescript'],
      // Enable lint rules for React DOM
      reactDom.configs.recommended,
    ],
    languageOptions: {
      parserOptions: {
        project: ['./tsconfig.node.json', './tsconfig.app.json'],
        tsconfigRootDir: import.meta.dirname,
      },
      // other options...
    },
  },
])
```
# stock

## Enable Add Row (Google Sheets Append API)

The Add page posts to `VITE_SHEETS_APPEND_ENDPOINT`. This project now includes a local API server using `spreadsheets.values.append`.

### 1. Configure env

Copy `.env.example` to `.env.local` and fill values:

- `VITE_SHEETS_APPEND_ENDPOINT=http://localhost:8787/api/sheets/append`
- `SPREADSHEET_ID` (or `VITE_SPREADSHEET_ID`)
- Google credentials:
  - `GOOGLE_APPLICATION_CREDENTIALS=./credentials.json` (recommended), or
  - `GOOGLE_SERVICE_ACCOUNT_KEY_JSON=...`

Default ranges:

- `SHEET_RANGE_STOCKS=Stocks!A:G`
- `SHEET_RANGE_DIVIDEND=Dividend!A:D`
- `SHEET_RANGE_MONEY_MOVE=Money Move!A:E`

### 2. Run frontend and API

In two terminals:

```bash
npm run dev
```

```bash
npm run server
```

Append endpoint health check:

```bash
curl http://localhost:8787/health
```

### 3. Notes

- Service account must have access to the target spreadsheet.
- `valueInputOption` is `USER_ENTERED`.
- Rows are appended via `POST /api/sheets/append` with body:

```json
{
  "table": "Stocks | Dividend | Money Move",
  "row": { "...": "..." }
}
```
