import { createClient } from '@supabase/supabase-js'

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL
const supabaseKey = import.meta.env.VITE_SUPABASE_ANON_KEY

console.log('URL:', supabaseUrl)
console.log('KEY:', supabaseKey?.slice(0, 20) + '...')

export const supabase = createClient(supabaseUrl, supabaseKey)