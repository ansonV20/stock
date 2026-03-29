import { createClient } from '@supabase/supabase-js'

const supabaseUrl =
  import.meta.env.VITE_PUBLIC_SUPABASE_URL?.trim() || import.meta.env.VITE_SUPABASE_URL?.trim()
const supabaseAnonKey =
  import.meta.env.VITE_PUBLIC_SUPABASE_ANON_KEY?.trim() || import.meta.env.VITE_SUPABASE_ANON_KEY?.trim()

if (!supabaseUrl || !supabaseAnonKey) {
  throw new Error(
    'Missing Supabase env vars: set VITE_PUBLIC_SUPABASE_URL/VITE_PUBLIC_SUPABASE_ANON_KEY (or VITE_SUPABASE_URL/VITE_SUPABASE_ANON_KEY)',
  )
}

export const supabase = createClient(supabaseUrl, supabaseAnonKey)
