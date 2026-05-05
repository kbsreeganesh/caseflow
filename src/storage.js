/**
 * CaseFlow — Supabase Storage Adapter
 * Env vars read at build time by Vite:
 *   VITE_SUPABASE_URL
 *   VITE_SUPABASE_ANON_KEY
 */

const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL  || '';
const SUPABASE_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY || '';
const TABLE        = 'caseflow_store';

async function sbFetch(path, opts = {}) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    ...opts,
    headers: {
      'apikey':        SUPABASE_KEY,
      'Authorization': `Bearer ${SUPABASE_KEY}`,
      'Content-Type':  'application/json',
      'Prefer':        'return=representation',
      ...(opts.headers || {}),
    },
  });
  if (!res.ok) {
    const txt = await res.text().catch(() => res.statusText);
    throw new Error(`Supabase ${opts.method||'GET'} -> ${res.status}: ${txt}`);
  }
  const text = await res.text();
  return text ? JSON.parse(text) : null;
}

/* localStorage fallback (used locally when no env vars) */
const ls = {
  async get(key)        { const v=localStorage.getItem(key); if(v===null) throw new Error('Key not found: '+key); return {key,value:v}; },
  async set(key,value)  { localStorage.setItem(key,value); return {key,value}; },
  async delete(key)     { localStorage.removeItem(key); return {key,deleted:true}; },
  async list(prefix='') { const keys=Object.keys(localStorage).filter(k=>!prefix||k.startsWith(prefix)); return {keys}; },
};

/* Supabase adapter */
const sb = {
  async get(key) {
    const rows = await sbFetch(`${TABLE}?key=eq.${encodeURIComponent(key)}&select=key,value`);
    if (!rows || !rows.length) throw new Error('Key not found: ' + key);
    return { key, value: rows[0].value };
  },
  async set(key, value) {
    await sbFetch(TABLE, {
      method:  'POST',
      headers: { 'Prefer': 'resolution=merge-duplicates,return=representation' },
      body:    JSON.stringify({ key, value, updated_at: new Date().toISOString() }),
    });
    return { key, value };
  },
  async delete(key) {
    await sbFetch(`${TABLE}?key=eq.${encodeURIComponent(key)}`, { method: 'DELETE' });
    return { key, deleted: true };
  },
  async list(prefix = '') {
    const path = prefix
      ? `${TABLE}?key=like.${encodeURIComponent(prefix + '%')}&select=key`
      : `${TABLE}?select=key`;
    const rows = await sbFetch(path);
    return { keys: (rows||[]).map(r => r.key) };
  },
};

if (SUPABASE_URL && SUPABASE_KEY) {
  window.storage = sb;
  console.log('[CaseFlow] Supabase connected:', SUPABASE_URL);
} else {
  window.storage = ls;
  console.warn('[CaseFlow] Using localStorage — set VITE_SUPABASE_URL + VITE_SUPABASE_ANON_KEY for shared mode.');
}
