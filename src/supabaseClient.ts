import { createClient, type SupabaseClient } from '@supabase/supabase-js'

/**
 * Cloud sync is optional. With no credentials configured the app runs exactly
 * as it did before — localStorage only — so a fork or a local checkout without
 * a Supabase project still works.
 *
 * Set these in `.env.local` (and in the Vercel project settings for deploys):
 *   VITE_SUPABASE_URL=https://<project-ref>.supabase.co
 *   VITE_SUPABASE_ANON_KEY=<anon public key>
 *
 * The anon key is a publishable key and is safe in the browser bundle: every
 * table is guarded by row-level security (see supabase/schema.sql).
 */

const url = import.meta.env.VITE_SUPABASE_URL as string | undefined
const anonKey = import.meta.env.VITE_SUPABASE_ANON_KEY as string | undefined

export const isCloudConfigured = Boolean(url && anonKey)

export const supabase: SupabaseClient | null = isCloudConfigured
  ? createClient(url!, anonKey!, {
      auth: {
        persistSession: true,
        autoRefreshToken: true,
        // The magic-link callback lands back on the app with tokens in the URL
        // hash; the client consumes them and cleans up after itself.
        detectSessionInUrl: true,
      },
    })
  : null
