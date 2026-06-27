// supabaseClient.js
import { createClient } from "@supabase/supabase-js";

const url = import.meta.env.VITE_SUPABASE_URL;
const key = import.meta.env.VITE_SUPABASE_ANON_KEY;

// db.schema => todas las queries .from() usan el schema 'recorridas'
// (Storage no depende del schema).
export const supabase = createClient(url, key, {
  db: { schema: "recorridas" },
  auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: false },
});

export const BUCKET = "recorrida-fotos";
