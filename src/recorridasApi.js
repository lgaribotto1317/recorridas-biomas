// recorridasApi.js
// Capa de datos entre la UI (que trabaja con NOMBRES) y Supabase (que usa IDs/FK).
import { supabase, BUCKET } from "./supabaseClient";

const PLANTA = "Biomas";

/* ---------- Mapas nombre<->id de las maestras (se cargan una vez) ---------- */
let maps = null;
export async function loadMaps() {
  if (maps) return maps;
  const [plantas, sectores, areas, personas] = await Promise.all([
    supabase.from("plantas").select("id,nombre"),
    supabase.from("sectores").select("id,nombre"),
    supabase.from("areas_responsables").select("id,nombre"),
    supabase.from("personas").select("id,nombre"),
  ]);
  const idByName = (rows) => Object.fromEntries((rows.data || []).map((r) => [r.nombre, r.id]));
  const nameById = (rows) => Object.fromEntries((rows.data || []).map((r) => [r.id, r.nombre]));
  maps = {
    plantaId: idByName(plantas)[PLANTA],
    sector: { id: idByName(sectores), name: nameById(sectores) },
    area: { id: idByName(areas), name: nameById(areas) },
    persona: { id: idByName(personas), name: nameById(personas) },
  };
  return maps;
}

/* ---------- URL pública de una foto ---------- */
export const publicUrl = (path) => supabase.storage.from(BUCKET).getPublicUrl(path).data.publicUrl;

/* ---------- DB -> UI ---------- */
function rowToUi(h, m, fotosByH) {
  const fs = fotosByH[h.id] || [];
  const last = (tipo) => {
    const arr = fs.filter((f) => f.tipo === tipo).sort((a, b) => a.created_at.localeCompare(b.created_at));
    return arr.length ? publicUrl(arr[arr.length - 1].storage_path) : null;
  };
  return {
    id: h.id,
    numero: h.numero,
    planta: PLANTA,
    sector: m.sector.name[h.sector_id] || "",
    sectorResp: m.area.name[h.area_resp_id] || "",
    responsable: m.persona.name[h.responsable_id] || "",
    relevadoPor: m.persona.name[h.relevado_por_id] || h.relevado_por_texto || "",
    criticidad: h.criticidad || "",
    descripcion: h.descripcion || "",
    estado: h.estado,
    fechaRegistro: h.fecha_registro,
    fechaCierre: h.fecha_cierre,
    comentarios: h.comentarios || "",
    fotoAntes: last("antes"),
    fotoDespues: last("despues"),
    createdAt: h.created_at,
  };
}

/* ---------- UI -> DB ---------- */
function uiToRow(ui, m) {
  const relId = m.persona.id[ui.relevadoPor] || null;
  return {
    planta_id: m.plantaId,
    sector_id: m.sector.id[ui.sector] || null,
    area_resp_id: m.area.id[ui.sectorResp] || null,
    responsable_id: m.persona.id[ui.responsable] || null,
    relevado_por_id: relId,
    relevado_por_texto: relId ? null : ((ui.relevadoPor || "").trim() || null),
    criticidad: ui.criticidad || null,
    descripcion: (ui.descripcion || "").trim() || null,
    estado: ui.estado || "No comenzado",
    fecha_cierre: ui.fechaCierre || null,
    comentarios: (ui.comentarios || "").trim() || null,
  };
}

/* ---------- Cargar todos los hallazgos (con fotos) ---------- */
export async function loadHallazgos() {
  const m = await loadMaps();
  const [{ data: hs }, { data: fs }] = await Promise.all([
    supabase.from("hallazgos").select("*").order("fecha_registro", { ascending: false }).order("numero", { ascending: false }),
    supabase.from("fotos").select("*"),
  ]);
  const fotosByH = {};
  (fs || []).forEach((f) => { (fotosByH[f.hallazgo_id] ||= []).push(f); });
  return (hs || []).map((h) => rowToUi(h, m, fotosByH));
}

/* ---------- Comprimir imagen (dataURL -> Blob jpeg) para no llenar storage ---------- */
export function compressImage(dataUrl, maxW = 1280, quality = 0.7) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      const scale = Math.min(1, maxW / img.width);
      const c = document.createElement("canvas");
      c.width = Math.round(img.width * scale);
      c.height = Math.round(img.height * scale);
      c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
      c.toBlob((b) => resolve(b), "image/jpeg", quality);
    };
    img.src = dataUrl;
  });
}

async function subirFoto(hallazgoId, tipo, dataUrl) {
  const blob = await compressImage(dataUrl);
  const path = `${hallazgoId}/${tipo}_${Date.now()}.jpg`;
  const up = await supabase.storage.from(BUCKET).upload(path, blob, { contentType: "image/jpeg", upsert: true });
  if (up.error) throw up.error;
  const ins = await supabase.from("fotos").insert({ hallazgo_id: hallazgoId, tipo, storage_path: path });
  if (ins.error) throw ins.error;
}

const esDataUrl = (s) => typeof s === "string" && s.startsWith("data:");

/* ---------- Crear hallazgo (+ foto antes) ---------- */
export async function createHallazgo(ui) {
  const m = await loadMaps();
  const { data, error } = await supabase.from("hallazgos").insert(uiToRow(ui, m)).select().single();
  if (error) throw error;
  if (esDataUrl(ui.fotoAntes)) await subirFoto(data.id, "antes", ui.fotoAntes);
  return data.id;
}

/* ---------- Actualizar hallazgo (+ fotos nuevas si las hay) ---------- */
export async function updateHallazgo(ui) {
  const m = await loadMaps();
  const { error } = await supabase.from("hallazgos").update(uiToRow(ui, m)).eq("id", ui.id);
  if (error) throw error;
  if (esDataUrl(ui.fotoAntes)) await subirFoto(ui.id, "antes", ui.fotoAntes);
  if (esDataUrl(ui.fotoDespues)) await subirFoto(ui.id, "despues", ui.fotoDespues);
}

/* ---------- Realtime: avisa cuando cambian hallazgos/fotos ---------- */
export function subscribe(onChange) {
  const ch = supabase
    .channel("recorridas-cambios")
    .on("postgres_changes", { event: "*", schema: "recorridas", table: "hallazgos" }, onChange)
    .on("postgres_changes", { event: "*", schema: "recorridas", table: "fotos" }, onChange)
    .subscribe();
  return () => supabase.removeChannel(ch);
}

/* ===================== AUTH (email + password) ===================== */

// Iniciar sesión
export async function signIn(email, password) {
  const { data, error } = await supabase.auth.signInWithPassword({ email, password });
  if (error) throw error;
  return data;
}

// Cerrar sesión
export async function signOut() {
  await supabase.auth.signOut();
}

// Sesión actual (o null)
export async function getSession() {
  const { data } = await supabase.auth.getSession();
  return data.session || null;
}

// Escuchar cambios de sesión (login/logout/refresh). Devuelve función para desuscribir.
export function onAuthChange(cb) {
  const { data } = supabase.auth.onAuthStateChange((_event, session) => cb(session || null));
  return () => data.subscription.unsubscribe();
}

// Persona vinculada al usuario logueado (por user_id). Devuelve { nombre, email } o null.
export async function currentPersona() {
  const { data: u } = await supabase.auth.getUser();
  const user = u?.user;
  if (!user) return null;
  const { data } = await supabase.from("personas").select("nombre").eq("user_id", user.id).maybeSingle();
  return { nombre: data?.nombre || null, email: user.email, userId: user.id };
}

// Cambiar la contraseña del usuario logueado (no usa mail)
export async function changePassword(newPassword) {
  const { error } = await supabase.auth.updateUser({ password: newPassword });
  if (error) throw error;
}

/* ===================== TRAZABILIDAD (audit log) ===================== */

// Trae los registros de auditoría (más recientes primero)
export async function loadAuditoria(limit = 500) {
  const { data, error } = await supabase
    .from("auditoria")
    .select("*")
    .order("creado_en", { ascending: false })
    .limit(limit);
  if (error) { console.error("loadAuditoria", error); return []; }
  return data || [];
}

// Realtime de la auditoría
export function subscribeAuditoria(onChange) {
  const ch = supabase
    .channel("recorridas-auditoria")
    .on("postgres_changes", { event: "*", schema: "recorridas", table: "auditoria" }, onChange)
    .subscribe();
  return () => supabase.removeChannel(ch);
}

// ---------- Vistas de usuario (aviso de hallazgos nuevos) ----------
// Lee la marca de "última vez que el usuario vio la sección Hallazgos".
// Devuelve un ISO timestamp, o null si el usuario nunca la vio (primera vez).
export async function getUltimaVista() {
  const { data: u } = await supabase.auth.getUser();
  const user = u?.user;
  if (!user) return null;
  const { data } = await supabase
    .from("vistas_usuario")
    .select("ultima_vista_hallazgos")
    .eq("user_id", user.id)
    .maybeSingle();
  return data?.ultima_vista_hallazgos || null;
}

// Marca "ahora" como última vista de Hallazgos para el usuario logueado.
// Upsert: crea la fila la primera vez, actualiza después.
export async function marcarVistoHallazgos() {
  const { data: u } = await supabase.auth.getUser();
  const user = u?.user;
  if (!user) return;
  const { error } = await supabase
    .from("vistas_usuario")
    .upsert({ user_id: user.id, ultima_vista_hallazgos: new Date().toISOString() }, { onConflict: "user_id" });
  if (error) console.error("marcarVistoHallazgos", error);
}
