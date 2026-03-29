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

## Supabase Migration + Google Login

This app now stores data in Supabase and scopes all records by logged-in user.

### 1. Install and configure env

Create `.env.local`:

```bash
VITE_SUPABASE_URL=https://YOUR_PROJECT_REF.supabase.co
VITE_SUPABASE_ANON_KEY=YOUR_PUBLIC_ANON_KEY
VITE_FINNHUB_TOKEN=YOUR_FINNHUB_TOKEN
```

Then run:

```bash
npm install
npm run dev
```

### 2. SQL schema (run in Supabase SQL Editor)

```sql
-- Needed for gen_random_uuid() if your project does not already have it.
create extension if not exists pgcrypto;

-- User profile table (requested name: "user").
create table if not exists public."user" (
  id uuid primary key references auth.users(id) on delete cascade,
  full_name text not null,
  email text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.stocks (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  stock text not null,
  currency text not null,
  price numeric(18, 6) not null,
  action text not null,
  time timestamptz not null,
  quantity numeric(18, 6) not null,
  handling_fees numeric(18, 6) not null default 0,
  created_at timestamptz not null default now()
);

create table if not exists public.dividend (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  stock text not null,
  currency text not null,
  div numeric(18, 6) not null,
  time timestamptz not null,
  created_at timestamptz not null default now()
);

create table if not exists public.money_move (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  name text not null,
  currency text not null,
  price numeric(18, 6) not null,
  time timestamptz not null,
  action text not null,
  created_at timestamptz not null default now()
);

create index if not exists idx_stocks_user_time on public.stocks(user_id, time);
create index if not exists idx_dividend_user_time on public.dividend(user_id, time);
create index if not exists idx_money_move_user_time on public.money_move(user_id, time);

alter table public."user" enable row level security;
alter table public.stocks enable row level security;
alter table public.dividend enable row level security;
alter table public.money_move enable row level security;

drop policy if exists "user_select_own" on public."user";
create policy "user_select_own"
on public."user"
for select
to authenticated
using (auth.uid() = id);

drop policy if exists "user_insert_own" on public."user";
create policy "user_insert_own"
on public."user"
for insert
to authenticated
with check (auth.uid() = id);

drop policy if exists "user_update_own" on public."user";
create policy "user_update_own"
on public."user"
for update
to authenticated
using (auth.uid() = id)
with check (auth.uid() = id);

drop policy if exists "stocks_select_own" on public.stocks;
create policy "stocks_select_own"
on public.stocks
for select
to authenticated
using (auth.uid() = user_id);

drop policy if exists "stocks_insert_own" on public.stocks;
create policy "stocks_insert_own"
on public.stocks
for insert
to authenticated
with check (auth.uid() = user_id);

drop policy if exists "dividend_select_own" on public.dividend;
create policy "dividend_select_own"
on public.dividend
for select
to authenticated
using (auth.uid() = user_id);

drop policy if exists "dividend_insert_own" on public.dividend;
create policy "dividend_insert_own"
on public.dividend
for insert
to authenticated
with check (auth.uid() = user_id);

drop policy if exists "money_move_select_own" on public.money_move;
create policy "money_move_select_own"
on public.money_move
for select
to authenticated
using (auth.uid() = user_id);

drop policy if exists "money_move_insert_own" on public.money_move;
create policy "money_move_insert_own"
on public.money_move
for insert
to authenticated
with check (auth.uid() = user_id);
```

### 3. Supabase Auth setup (Google)

1. In Supabase Dashboard, go to Authentication -> Providers -> Google and enable it.
2. Create OAuth credentials in Google Cloud Console:
   - Application type: Web application
   - Authorized redirect URI:
     - `https://YOUR_PROJECT_REF.supabase.co/auth/v1/callback`
3. Copy Client ID and Client Secret into Supabase Google provider settings.
4. In Supabase Dashboard, open Authentication -> URL Configuration and set:
   - Site URL: `http://localhost:5173`
   - Additional Redirect URLs: `http://localhost:5173`
5. Save settings and test sign-in from the app using the "Sign in with Google" button.

### 4. Behavior after migration

- Each user signs in with Google.
- App upserts profile into table `public."user"`.
- All `stocks`, `dividend`, `money_move` rows include `user_id`.
- RLS ensures each user can only read/write their own rows.
