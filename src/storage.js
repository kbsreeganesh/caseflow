/**
 * CaseFlow — Supabase Storage Adapter
 * ------------------------------------
 * Replaces window.storage with Supabase so all users share the same data.
 *
 * Table schema (created via SQL in Supabase dashboard):
 *
 *   create table caseflow_store (
 *     key   text primary key,
 *     value text  not null,
 *     updated_at timestamptz default now()
 *   );
 *
 *   -- Allow anonymous read/write (fine for internal trial)
 *   alter table caseflow_store enable row level security;
 *   create policy "anon all" on caseflow_store for all using (true) with check (true);
 */

import { createClient } from '@supabase/supabase-js'

const SUPABASE_URL  = import.meta.env.VITE_SUPABASE_URL
const SUPABASE_ANON = import.meta.env.VITE_SUPABASE_ANON_KEY
const TABLE         = 'caseflow_store'

if (!SUPABASE_URL || !SUPABASE_ANON) {
  console.error('[CaseFlow] Missing VITE_SUPABASE_URL or VITE_SUPABASE_ANON_KEY in .env')
}

const sb = createClient(SUPABASE_URL, SUPABASE_ANON)

window.storage = {
  async get(key) {
    const { data, error } = await sb
      .from(TABLE)
      .select('value')
      .eq('key', key)
      .single()
    if (error || !data) throw new Error(`Key not found: ${key}`)
    return { key, value: data.value }
  },

  async set(key, value) {
    const { error } = await sb
      .from(TABLE)
      .upsert({ key, value, updated_at: new Date().toISOString() }, { onConflict: 'key' })
    if (error) { console.error('[CaseFlow] set error:', error); return null }
    return { key, value }
  },

  async delete(key) {
    await sb.from(TABLE).delete().eq('key', key)
    return { key, deleted: true }
  },

  async list(prefix = '') {
    const { data } = await sb.from(TABLE).select('key')
    const keys = (data || []).map(r => r.key).filter(k => !prefix || k.startsWith(prefix))
    return { keys, prefix }
  },
}

console.log('[CaseFlow] Supabase storage adapter installed →', SUPABASE_URL)
