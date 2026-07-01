import React, { useState, useRef } from "react";
import {
  Camera, Plus, X, ChevronLeft, Download, ClipboardList, BarChart3, Table2,
  MapPin, User, Search, SlidersHorizontal, Check, Image as ImageIcon, Ban, Save, AlertTriangle, LogOut, KeyRound, ChevronDown, History
} from "lucide-react";
import * as XLSX from "xlsx";
import { BarChart, Bar, XAxis, YAxis, ResponsiveContainer, Cell, Tooltip } from "recharts";
import { useEffect } from "react";
import * as api from "./recorridasApi";

/* ───────── Branding ───────── */
const LOGO = "data:image/png;base64,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";
const C = {
  blue: "#0088CE", blueD: "#0072AD", orange: "#E37222", ink: "#1E1E1E",
  page: "#F5F7FA", card: "#FFFFFF", border: "#E2E8F0", muted: "#64748B",
  green: "#16A34A", amber: "#F59E0B", red: "#DC2626", grey: "#94A3B8",
};
// Tipografía corporativa Ferring (Helvetica para web / Arial primaria)
const FONT = '"Helvetica Neue", Helvetica, Arial, sans-serif';

/* ───────── Listas maestras (Planta Biomas — del Excel del proyecto) ───────── */
const PLANTA = "Biomas";
const SECTORES = ["RO", "FO", "FA", "F1", "FB", "F2", "PTEL", "Bioterio", "QC", "QA", "General", "Planta de Agua", "Área técnica"];
const SECTOR_RESP = ["Mantenimiento", "Facilities", "Ingeniería", "Proyectos"];
const PERSONAS = [
  "Sergio Parcheiczuk", "Leonardo Garibotto", "Fernando Alarcon", "Mauricio Mangieri",
  "Leandro Grovas", "Sergio Gallego", "Juan Alasia", "Luciano Fioretti",
  "Gustavo Pare", "Carlos Rosic", "Ignacio Nieto",
]; // relevadores y responsables salen de la misma lista por ahora
const CRITICIDADES = ["Baja", "Media", "Alta"];
const ESTADOS = ["No comenzado", "En curso", "Finalizado", "No aplica"];
const ESTADOS_KANBAN = ["No comenzado", "En curso", "Finalizado"]; // "No aplica" va a su propia pestaña
const REQ = ["sector", "responsable", "criticidad", "descripcion"]; // requeridos: carga completa (tarjetas/Dashboard/Excel/filtro "Por completar")
const REQ_FIN = ["sector", "responsable", "criticidad", "comentarios"]; // requeridos para FINALIZAR (caso B): obligatorios + comentario final, SIN descripción

// Sólo estas personas (jefes/gerentes/directores) ven el gráfico "Por responsable" del Dashboard.
const VE_RESPONSABLE = ["Carlos Rosic", "Sergio Parcheiczuk", "Leonardo Garibotto", "Fernando Alarcon"];
const normNombre = (s) => (s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toLowerCase();
const puedeVerResponsable = (nombre) => !!nombre && VE_RESPONSABLE.some((n) => normNombre(n) === normNombre(nombre));

const critColor = { Baja: C.green, Media: C.amber, Alta: C.red, "": C.grey };
const critText  = { Baja: "#15803D", Media: "#B45309", Alta: "#B91C1C", "": C.muted };
const estadoPill = {
  "No comenzado": { bg: "#FEF2F2", fg: "#B91C1C" },
  "En curso":     { bg: "#FFF7ED", fg: "#B45309" },
  "Finalizado":   { bg: "#F0FDF4", fg: "#15803D" },
  "No aplica":    { bg: "#F1F5F9", fg: C.grey },
};
const estadoColor = { "No comenzado": C.red, "En curso": C.amber, "Finalizado": C.green, "No aplica": C.grey };

const incompleto = (h) => REQ.some((k) => !h[k]);
const hoy = () => new Date().toISOString().slice(0, 10);
const fmtNro = (n) => (n == null ? "" : "#" + String(n).padStart(4, "0"));
const ph = (t) => `data:image/svg+xml;utf8,` + encodeURIComponent(
  `<svg xmlns='http://www.w3.org/2000/svg' width='400' height='300'><rect width='400' height='300' fill='#E2E8F0'/>
   <text x='50%' y='50%' fill='#94A3B8' font-family='sans-serif' font-size='20' text-anchor='middle' dominant-baseline='middle'>${t}</text></svg>`);

const SEED = [
  { id: 1, planta: PLANTA, sector: "FO", sectorResp: "Mantenimiento", responsable: "Fernando Alarcon", relevadoPor: "Sergio Gallego",
    criticidad: "Media", descripcion: "Óxido en baranda e instalaciones perimetrales.", estado: "Finalizado",
    fechaRegistro: "2026-06-10", fechaCierre: "2026-06-18", comentarios: "Lijado y pintado con epoxi. OK.",
    fotoAntes: ph("ANTES"), fotoDespues: ph("DESPUÉS") },
  { id: 2, planta: PLANTA, sector: "Planta de Agua", sectorResp: "Ingeniería", responsable: "Juan Alasia", relevadoPor: "Esteban (jerárquico)",
    criticidad: "Alta", descripcion: "Instalación eléctrica a acomodar, riesgo de contacto.", estado: "En curso",
    fechaRegistro: "2026-06-15", fechaCierre: null, comentarios: "", fotoAntes: ph("ANTES"), fotoDespues: null },
  { id: 3, planta: PLANTA, sector: "FA", sectorResp: "", responsable: "", relevadoPor: "Santiago Ramos Mejia",
    criticidad: "", descripcion: "", estado: "No comenzado", fechaRegistro: "2026-06-20", fechaCierre: null,
    comentarios: "", fotoAntes: ph("ANTES"), fotoDespues: null }, // ejemplo: cargado solo con foto
];

/* ───────── UI helpers ───────── */
const Label = ({ children }) => <span style={{ display: "block", fontSize: 10, fontWeight: 600, letterSpacing: ".10em", textTransform: "uppercase", color: C.muted, whiteSpace: "nowrap" }}>{children}</span>;
const fieldStyle = { marginTop: 4, width: "100%", boxSizing: "border-box", borderRadius: 8, border: `1px solid ${C.border}`, background: C.card, padding: "10px 12px", fontSize: 14, color: C.ink, outline: "none" };

function Select({ value, onChange, placeholder, options, icon: Icon }) {
  return (
    <div style={{ position: "relative" }}>
      {Icon && <Icon size={15} style={{ position: "absolute", left: 10, top: 12, color: C.muted, pointerEvents: "none" }} />}
      <ChevronDown size={16} style={{ position: "absolute", right: 10, top: 12, color: C.muted, pointerEvents: "none" }} />
      <select value={value} onChange={(e) => onChange(e.target.value)}
        style={{ ...fieldStyle, paddingLeft: Icon ? 32 : 12, paddingRight: 30, appearance: "none", color: value ? C.ink : C.muted }}>
        <option value="">{placeholder}</option>
        {options.map((o) => <option key={o} value={o} style={{ color: C.ink }}>{o}</option>)}
      </select>
    </div>
  );
}

// Combo: sugiere las personas del padrón pero permite escribir un nombre libre.
// Mismo look que Select (incluye el chevron) para que se vean iguales.
function PersonaCombo({ value, onChange, placeholder, options, icon: Icon }) {
  return (
    <div style={{ position: "relative" }}>
      {Icon && <Icon size={15} style={{ position: "absolute", left: 10, top: 12, color: C.muted, pointerEvents: "none" }} />}
      <ChevronDown size={16} style={{ position: "absolute", right: 10, top: 12, color: C.muted, pointerEvents: "none" }} />
      <input list="dl-personas" value={value} onChange={(e) => onChange(e.target.value)} placeholder={placeholder}
        style={{ ...fieldStyle, paddingLeft: Icon ? 32 : 12, paddingRight: 30, color: value ? C.ink : C.muted }} />
      <datalist id="dl-personas">{options.map((o) => <option key={o} value={o} />)}</datalist>
    </div>
  );
}

function PhotoTile({ src, onPick, label }) {
  const ref = useRef(null);
  return (
    <div style={{ flex: 1 }}>
      <Label>{label}</Label>
      <button type="button" onClick={() => ref.current?.click()}
        style={{ marginTop: 4, position: "relative", width: "100%", aspectRatio: "4/3", overflow: "hidden", borderRadius: 8, border: `1px solid ${C.border}`, background: C.page, cursor: "pointer" }}>
        {src ? <img src={src} alt={label} style={{ height: "100%", width: "100%", objectFit: "cover" }} />
          : <span style={{ position: "absolute", inset: 0, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 4, color: C.muted }}>
              <Camera size={26} /><span style={{ fontSize: 12 }}>Tocar para capturar</span></span>}
      </button>
      <input ref={ref} type="file" accept="image/*" capture="environment" style={{ display: "none" }}
        onChange={(e) => { const f = e.target.files?.[0]; if (!f) return; const r = new FileReader(); r.onload = () => onPick(r.result); r.readAsDataURL(f); }} />
    </div>
  );
}

function CompactPhoto({ src, onPick, onClear, label = "Tomar o adjuntar foto" }) {
  const ref = useRef(null);
  return (
    <div style={{ marginTop: 6 }}>
      <button type="button" onClick={() => ref.current?.click()}
        style={{ display: "inline-flex", alignItems: "center", gap: 8, borderRadius: 8, border: `1px solid ${C.blue}`, background: "#fff", color: C.blue, padding: "9px 14px", fontSize: 14, fontWeight: 600, cursor: "pointer" }}>
        <Camera size={17} /> {src ? "Cambiar foto" : label}
      </button>
      <input ref={ref} type="file" accept="image/*" style={{ display: "none" }}
        onChange={(e) => { const f = e.target.files?.[0]; if (!f) return; const r = new FileReader(); r.onload = () => onPick(r.result); r.readAsDataURL(f); }} />
      {src && (
        <div style={{ marginTop: 10, display: "flex", alignItems: "flex-start", gap: 10 }}>
          <img src={src} alt={label} style={{ width: 120, height: 90, objectFit: "cover", borderRadius: 8, border: `1px solid ${C.border}` }} />
          <button type="button" onClick={onClear} style={{ display: "inline-flex", alignItems: "center", gap: 4, fontSize: 12, color: C.muted, background: "none", border: "none", cursor: "pointer" }}><X size={14} /> Quitar</button>
        </div>
      )}
    </div>
  );
}

/* ───────── Nuevo hallazgo (carga mínima: solo foto obligatoria) ───────── */
function NuevoHallazgo({ onClose, onSave, defaultRelevadoPor = "" }) {
  const [f, setF] = useState({ fotoAntes: null, fotoDespues: null, sector: "", sectorResp: "", responsable: "", relevadoPor: defaultRelevadoPor, criticidad: "Media", descripcion: "" });
  const ok = !!f.fotoAntes || !!f.descripcion.trim();
  // Si además de la foto después están todos los campos requeridos, el hallazgo nace cerrado.
  const completo = !!f.sector && !!f.responsable && !!f.criticidad && !!f.descripcion.trim();
  const cierraAhora = !!f.fotoDespues && completo;
  return (
    <div style={{ position: "fixed", inset: 0, zIndex: 30, display: "flex", flexDirection: "column", background: C.page }}>
      <header style={{ display: "flex", alignItems: "center", justifyContent: "space-between", borderBottom: `1px solid ${C.border}`, background: C.card, padding: "12px 16px" }}>
        <button onClick={onClose} style={{ display: "flex", alignItems: "center", gap: 4, color: C.muted, background: "none", border: "none", cursor: "pointer" }}><X size={20} /><span style={{ fontSize: 14 }}>Cancelar</span></button>
        <h2 style={{ fontSize: 14, fontWeight: 600, color: C.ink }}>Nuevo hallazgo</h2><span style={{ width: 64 }} />
      </header>
      <div style={{ flex: 1, overflowY: "auto", padding: 16, display: "flex", flexDirection: "column", gap: 14 }}>
        <div style={{ borderBottom: `1px solid ${C.border}`, paddingBottom: 14, display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
          <div style={{ display: "flex", gap: 20, flexWrap: "wrap" }}>
            <div>
              <Label>Foto — Antes</Label>
              <CompactPhoto src={f.fotoAntes} onPick={(v) => setF({ ...f, fotoAntes: v })} onClear={() => setF({ ...f, fotoAntes: null })} label="Foto antes" />
            </div>
            <div>
              <Label>Foto — Después (opcional)</Label>
              <CompactPhoto src={f.fotoDespues} onPick={(v) => setF({ ...f, fotoDespues: v })} onClear={() => setF({ ...f, fotoDespues: null })} label="Foto después" />
            </div>
          </div>
          <div style={{ minWidth: 160 }}>
            <Label>Criticidad</Label>
            <div style={{ marginTop: 6 }}>
              <Select value={f.criticidad} onChange={(v) => setF({ ...f, criticidad: v })} placeholder="Criticidad" options={CRITICIDADES} />
            </div>
          </div>
        </div>
        <div className="rec-nh-row" style={{ display: "flex", gap: 14, flexWrap: "wrap", alignItems: "stretch" }}>
          <div className="rec-nh-grid" style={{ flex: "0 0 60%", minWidth: 280, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, alignContent: "start" }}>
            <div><Label>Relevado por</Label><PersonaCombo value={f.relevadoPor} onChange={(v) => setF({ ...f, relevadoPor: v })} placeholder="Quién releva" options={PERSONAS} icon={User} /></div>
            <div><Label>Sector</Label><Select value={f.sector} onChange={(v) => setF({ ...f, sector: v })} placeholder="Sector" options={SECTORES} icon={MapPin} /></div>
            <div><Label>Responsable</Label><Select value={f.responsable} onChange={(v) => setF({ ...f, responsable: v })} placeholder="Responsable" options={PERSONAS} icon={User} /></div>
            <div><Label>Sector responsable</Label><Select value={f.sectorResp} onChange={(v) => setF({ ...f, sectorResp: v })} placeholder="Área responsable" options={SECTOR_RESP} /></div>
          </div>
          <div className="rec-nh-desc" style={{ flex: 1, minWidth: 220, display: "flex", flexDirection: "column" }}>
            <Label>Descripción</Label>
            <textarea value={f.descripcion} onChange={(e) => setF({ ...f, descripcion: e.target.value })}
              placeholder="Qué se observó y qué corresponde corregir… (opcional)" style={{ ...fieldStyle, resize: "none", flex: 1, minHeight: 120 }} />
          </div>
        </div>
        {cierraAhora && (
          <div style={{ display: "flex", alignItems: "center", gap: 8, background: "#F0FDF4", border: `1px solid #BBF7D0`, color: "#15803D", borderRadius: 8, padding: "8px 12px", fontSize: 12.5 }}>
            <Check size={15} /> Tiene foto después y todos los datos: se va a registrar como «Finalizado».
          </div>
        )}
      </div>
      <footer style={{ borderTop: `1px solid ${C.border}`, background: C.card, padding: "12px 16px" }}>
        <button disabled={!ok} onClick={() => onSave({ ...f, id: Date.now(), planta: PLANTA, estado: cierraAhora ? "Finalizado" : "No comenzado", fechaRegistro: hoy(), fechaCierre: cierraAhora ? hoy() : null, comentarios: "" })}
          style={{ width: "100%", borderRadius: 8, padding: "12px 0", fontSize: 14, fontWeight: 600, border: "none", cursor: ok ? "pointer" : "not-allowed", background: ok ? (cierraAhora ? C.green : C.blue) : C.border, color: ok ? "#fff" : C.muted }}>
          {cierraAhora ? "Registrar y finalizar" : "Registrar hallazgo"}
        </button>
      </footer>
    </div>
  );
}

/* ───────── Detalle ───────── */
function Detalle({ h, onClose, onUpdate }) {
  const refDespues = useRef(null);
  const [coment, setComent] = useState(h.comentarios || "");
  const [desc, setDesc] = useState(h.descripcion || "");
  const [zoomSrc, setZoomSrc] = useState(null);
  // Layout 70/30 sólo en PC; en celular se apila (mobile queda como estaba).
  const [esMobile, setEsMobile] = useState(() => typeof window !== "undefined" && window.matchMedia("(max-width:640px)").matches);
  useEffect(() => {
    const mq = window.matchMedia("(max-width:640px)");
    const fn = (e) => setEsMobile(e.matches);
    mq.addEventListener("change", fn);
    return () => mq.removeEventListener("change", fn);
  }, []);
  const dirty = coment !== (h.comentarios || "") || desc !== (h.descripcion || "");
  const cerrado = h.estado === "Finalizado" || h.estado === "No aplica";
  // Gate de FINALIZAR (caso B): sector + responsable + criticidad + comentario final. La descripción NO bloquea.
  const faltan = REQ_FIN.filter((k) => k === "comentarios" ? !coment.trim() : !h[k]);

  const patch = (extra) => onUpdate({ ...h, descripcion: desc, comentarios: coment, ...extra });
  const attachDespues = (dataUrl) => patch({ fotoDespues: dataUrl });
  const finalizar = () => { if (!faltan.length) patch({ estado: "Finalizado", fechaCierre: hoy() }); };

  // Auto-guardar desc/comentarios al volver si hay cambios pendientes
  const handleClose = () => { if (dirty) patch({}); onClose(); };

  // ── Bloques reutilizables (mismo contenido; en mobile/PC cambia sólo la disposición) ──
  const bFotos = (
    <div style={{ display: "flex", gap: 12 }}>
      <figure style={{ flex: 1, margin: 0 }}>
        <Label>Antes</Label>
        <div style={{ marginTop: 4, aspectRatio: "4/3", overflow: "hidden", borderRadius: 8, border: `1px solid ${C.border}`, cursor: h.fotoAntes ? "zoom-in" : "default" }}
          onClick={() => h.fotoAntes && setZoomSrc(h.fotoAntes)}>
          {h.fotoAntes ? <img src={h.fotoAntes} alt="Antes" style={{ height: "100%", width: "100%", objectFit: "cover" }} /> : <div style={{ height: "100%", display: "flex", alignItems: "center", justifyContent: "center", color: C.grey }}><ImageIcon size={22} /></div>}
        </div>
      </figure>
      <figure style={{ flex: 1, margin: 0 }}>
        <Label>Después</Label>
        <div style={{ marginTop: 4, aspectRatio: "4/3", overflow: "hidden", borderRadius: 8, border: `1px solid ${C.border}`, background: C.card, cursor: h.fotoDespues ? "zoom-in" : "default" }}
          onClick={() => h.fotoDespues ? setZoomSrc(h.fotoDespues) : refDespues.current?.click()}>
          {h.fotoDespues
            ? <img src={h.fotoDespues} alt="Después" style={{ height: "100%", width: "100%", objectFit: "cover" }} />
            : <div style={{ height: "100%", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 4, color: C.muted }}>
                <Camera size={22} /><span style={{ fontSize: 11 }}>Capturar o galería</span></div>}
        </div>
        <input ref={refDespues} type="file" accept="image/*" style={{ display: "none" }}
          onChange={(e) => { const f = e.target.files?.[0]; if (!f) return; const r = new FileReader(); r.onload = () => attachDespues(r.result); r.readAsDataURL(f); }} />
      </figure>
    </div>
  );

  const bPills = (
    <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
      <span style={{ borderRadius: 6, padding: "4px 8px", fontSize: 12, fontWeight: 500, background: estadoPill[h.estado].bg, color: estadoPill[h.estado].fg }}>{h.estado}</span>
      <span style={{ borderRadius: 6, padding: "4px 8px", fontSize: 12, fontWeight: 500, background: "#fff", border: `1px solid ${C.border}`, color: critText[h.criticidad] }}>Criticidad {h.criticidad || "Sin clasificar"}</span>
      {incompleto({ ...h, descripcion: desc }) && <span style={{ display: "inline-flex", alignItems: "center", gap: 4, borderRadius: 6, padding: "4px 8px", fontSize: 12, fontWeight: 500, background: "#FFF7ED", color: "#B45309" }}><AlertTriangle size={13} /> Por completar</span>}
    </div>
  );

  const bDesc = (<div><Label>Descripción</Label><textarea rows={3} value={desc} onChange={(e) => setDesc(e.target.value)} placeholder="Qué corresponde corregir…" style={{ ...fieldStyle, resize: "none" }} /></div>);
  const bComent = (<div><Label>Comentarios (los completa el responsable)</Label><textarea rows={3} value={coment} onChange={(e) => setComent(e.target.value)} placeholder="Avance, cotizaciones, motivo si no aplica…" style={{ ...fieldStyle, resize: "none" }} /></div>);

  // Campos mobile: 4 selectores en rec-2col + criticidad en 3 botones (igual que antes).
  const bCamposMobile = (
    <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
      <div className="rec-2col">
        <div><Label>Relevado por</Label><PersonaCombo value={h.relevadoPor} onChange={(v) => onUpdate({ ...h, descripcion: desc, comentarios: coment, relevadoPor: v })} placeholder="Elegir o escribir…" options={PERSONAS} /></div>
        <div><Label>Sector</Label><Select value={h.sector} onChange={(v) => onUpdate({ ...h, descripcion: desc, comentarios: coment, sector: v })} placeholder="Elegir…" options={SECTORES} /></div>
        <div><Label>Responsable</Label><Select value={h.responsable} onChange={(v) => onUpdate({ ...h, descripcion: desc, comentarios: coment, responsable: v })} placeholder="Elegir…" options={PERSONAS} /></div>
        <div><Label>Sector responsable</Label><Select value={h.sectorResp} onChange={(v) => onUpdate({ ...h, descripcion: desc, comentarios: coment, sectorResp: v })} placeholder="Elegir…" options={SECTOR_RESP} /></div>
      </div>
      <div>
        <Label>Criticidad</Label>
        <div style={{ marginTop: 6, display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 8 }}>
          {CRITICIDADES.map((c) => { const active = h.criticidad === c;
            return <button key={c} onClick={() => onUpdate({ ...h, descripcion: desc, comentarios: coment, criticidad: active ? "" : c })}
              style={{ borderRadius: 8, padding: "8px 0", fontSize: 13, fontWeight: 500, cursor: "pointer", border: active ? "1px solid transparent" : `1px solid ${C.border}`, background: active ? critColor[c] : C.card, color: active ? "#fff" : C.ink }}>{c}</button>; })}
        </div>
      </div>
    </div>
  );

  // Campos PC: 5 selectores apilados en la columna del 30%; criticidad como dropdown.
  const bCamposPC = (
    <>
      <div><Label>Relevado por</Label><PersonaCombo value={h.relevadoPor} onChange={(v) => onUpdate({ ...h, descripcion: desc, comentarios: coment, relevadoPor: v })} placeholder="Elegir o escribir…" options={PERSONAS} /></div>
      <div><Label>Sector</Label><Select value={h.sector} onChange={(v) => onUpdate({ ...h, descripcion: desc, comentarios: coment, sector: v })} placeholder="Elegir…" options={SECTORES} /></div>
      <div><Label>Sector responsable</Label><Select value={h.sectorResp} onChange={(v) => onUpdate({ ...h, descripcion: desc, comentarios: coment, sectorResp: v })} placeholder="Elegir…" options={SECTOR_RESP} /></div>
      <div><Label>Responsable</Label><Select value={h.responsable} onChange={(v) => onUpdate({ ...h, descripcion: desc, comentarios: coment, responsable: v })} placeholder="Elegir…" options={PERSONAS} /></div>
      <div><Label>Criticidad</Label><Select value={h.criticidad} onChange={(v) => onUpdate({ ...h, descripcion: desc, comentarios: coment, criticidad: v })} placeholder="Elegir…" options={CRITICIDADES} /></div>
    </>
  );

  const bAcciones = (
    <>
      {!cerrado && h.estado === "No comenzado" && (
        <button onClick={() => patch({ estado: "En curso" })} style={{ width: "100%", borderRadius: 8, border: `1px solid ${C.amber}`, background: "#FFF7ED", color: "#B45309", padding: "10px 0", fontSize: 14, fontWeight: 500, cursor: "pointer" }}>Marcar «En curso»</button>
      )}
      {!cerrado && (
        <button onClick={() => patch({ estado: "No aplica", fechaCierre: hoy() })} style={{ width: "100%", display: "flex", alignItems: "center", justifyContent: "center", gap: 8, borderRadius: 8, border: `1px solid ${C.border}`, background: C.card, color: C.muted, padding: "10px 0", fontSize: 14, fontWeight: 500, cursor: "pointer" }}><Ban size={16} /> No aplica (descartar)</button>
      )}
      {h.estado === "Finalizado" && <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 8, background: "#F0FDF4", color: "#15803D", borderRadius: 8, padding: "12px 0", fontSize: 14 }}><Check size={17} /> Hallazgo resuelto</div>}
      {h.estado === "No aplica" && <button onClick={() => patch({ estado: "No comenzado", fechaCierre: null })} style={{ width: "100%", borderRadius: 8, border: `1px solid ${C.border}`, background: C.card, color: C.muted, padding: "10px 0", fontSize: 14, cursor: "pointer" }}>Reactivar hallazgo</button>}
    </>
  );

  return (
    <div style={{ position: "fixed", inset: 0, zIndex: 30, display: "flex", flexDirection: "column", background: C.page }}>
      {zoomSrc && <Lightbox src={zoomSrc} onClose={() => setZoomSrc(null)} />}

      {/* Header siempre visible con "Volver" */}
      <header style={{ display: "flex", alignItems: "center", gap: 8, borderBottom: `1px solid ${C.border}`, background: C.card, padding: "12px 16px", flexShrink: 0 }}>
        <button onClick={handleClose} style={{ display: "flex", alignItems: "center", gap: 4, color: C.muted, background: "none", border: "none", cursor: "pointer" }}><ChevronLeft size={22} /><span style={{ fontSize: 14 }}>Volver</span></button>
        <h2 style={{ fontSize: 14, fontWeight: 600, color: C.ink, marginLeft: 4 }}>{h.numero != null && <span style={{ color: C.muted, fontVariantNumeric: "tabular-nums" }}>{fmtNro(h.numero)} · </span>}{h.sector || "Sin sector"} · {h.planta}</h2>
      </header>

      {/* Scroll area — padding-bottom para que No aplica no quede tapado por el footer */}
      <div style={{ flex: 1, overflowY: "auto", WebkitOverflowScrolling: "touch", padding: 16, paddingBottom: 100, display: "flex", flexDirection: "column", gap: 16 }}>
        {esMobile ? (
          /* MOBILE: orden y disposición como estaba (fotos → pills → campos → desc → comentarios) */
          <>
            {bFotos}
            {bPills}
            {bCamposMobile}
            {bDesc}
            {bComent}
          </>
        ) : (
          /* PC: split 70% fotos / 30% campos; descripción y comentarios debajo a todo el ancho */
          <>
            {bPills}
            <div style={{ display: "flex", gap: 16, alignItems: "flex-start" }}>
              <div style={{ flex: "0 0 70%", maxWidth: "70%", minWidth: 0 }}>{bFotos}</div>
              <div style={{ flex: 1, minWidth: 0, display: "flex", flexDirection: "column", gap: 12 }}>{bCamposPC}</div>
            </div>
            {bDesc}
            {bComent}
          </>
        )}

        {/* Marcar en curso + No aplica + estado (dentro del scroll, por encima del footer) */}
        {bAcciones}
      </div>

      {/* Footer sticky: aviso + foto después + guardar/finalizar — siempre visible */}
      {!cerrado ? (
        <footer style={{ borderTop: `1px solid ${C.border}`, background: C.card, flexShrink: 0 }}>
          {faltan.length > 0 && (
            <div style={{ background: "#FFF7ED", borderBottom: `1px solid #FDE68A`, padding: "6px 16px", fontSize: 12, color: "#B45309", display: "flex", gap: 6, alignItems: "center" }}>
              <AlertTriangle size={13} /> Falta: {faltan.map((k) => ({ sector: "sector", responsable: "responsable", criticidad: "criticidad", comentarios: "comentario final" }[k])).join(", ")}
            </div>
          )}
          <div style={{ padding: "10px 16px", display: "flex", gap: 8 }}>
            <button onClick={() => refDespues.current?.click()}
              style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 6, borderRadius: 8, border: `1px solid ${C.blue}`, background: "#fff", color: C.blue, padding: "11px 0", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
              <Camera size={16} /> {h.fotoDespues ? "Cambiar foto" : "Foto después"}
            </button>
            <button onClick={finalizar} disabled={faltan.length > 0}
              style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 6, borderRadius: 8, border: "none", padding: "11px 0", fontSize: 13, fontWeight: 600, cursor: faltan.length > 0 ? "default" : "pointer", background: faltan.length > 0 ? C.amber : C.green, color: "#fff" }}>
              {faltan.length > 0 ? <><Save size={16} /> Guardar</> : <><Check size={16} /> Finalizar</>}
            </button>
          </div>
        </footer>
      ) : (
        /* Cerrado (Finalizado / No aplica): se puede seguir agregando o cambiando la foto «después» sin reabrir el hallazgo */
        <footer style={{ borderTop: `1px solid ${C.border}`, background: C.card, flexShrink: 0 }}>
          <div style={{ padding: "10px 16px" }}>
            <button onClick={() => refDespues.current?.click()}
              style={{ width: "100%", display: "flex", alignItems: "center", justifyContent: "center", gap: 6, borderRadius: 8, border: `1px solid ${C.blue}`, background: "#fff", color: C.blue, padding: "11px 0", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
              <Camera size={16} /> {h.fotoDespues ? "Cambiar foto después" : "Agregar foto después"}
            </button>
          </div>
        </footer>
      )}
    </div>
  );
}

/* ───────── Lightbox + miniaturas ampliables ───────── */
function Lightbox({ src, onClose }) {
  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, zIndex: 50, background: "rgba(15,23,42,.85)", display: "flex", alignItems: "center", justifyContent: "center", padding: 24, cursor: "zoom-out" }}>
      <img src={src} alt="" style={{ maxHeight: "90%", maxWidth: "90%", borderRadius: 8, boxShadow: "0 10px 40px rgba(0,0,0,.5)" }} />
      <button onClick={onClose} style={{ position: "absolute", top: 16, right: 16, height: 40, width: 40, borderRadius: 999, border: "none", background: "rgba(255,255,255,.15)", color: "#fff", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}><X size={22} /></button>
    </div>
  );
}

function ZoomImg({ src, w = 38, h = 38 }) {
  const [zoom, setZoom] = useState(false);
  if (!src) return null;
  return (
    <>
      <button type="button" onClick={() => setZoom(true)} title="Ampliar"
        style={{ height: h, width: w, borderRadius: 4, border: `1px solid ${C.border}`, padding: 0, overflow: "hidden", cursor: "zoom-in" }}>
        <img src={src} alt="" style={{ height: "100%", width: "100%", objectFit: "cover" }} />
      </button>
      {zoom && <Lightbox src={src} onClose={() => setZoom(false)} />}
    </>
  );
}

/* ───────── Foto compacta para la planilla (Después) ───────── */
function GridPhoto({ src, onPick }) {
  const ref = useRef(null);
  const [zoom, setZoom] = useState(false);
  return (
    <>
      <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
        {src && (
          <button type="button" onClick={() => setZoom(true)} title="Ampliar"
            style={{ height: 38, width: 50, borderRadius: 4, border: `1px solid ${C.border}`, padding: 0, overflow: "hidden", cursor: "zoom-in" }}>
            <img src={src} alt="Después" style={{ height: "100%", width: "100%", objectFit: "cover" }} />
          </button>
        )}
        <button type="button" onClick={() => ref.current?.click()} title={src ? "Cambiar" : "Cargar foto"}
          style={{ height: 38, width: src ? 32 : 50, borderRadius: 4, border: `1px dashed ${C.border}`, background: C.page, padding: 0, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", color: C.muted }}>
          <Camera size={15} />
        </button>
      </div>
      <input ref={ref} type="file" accept="image/*" capture="environment" style={{ display: "none" }}
        onChange={(e) => { const f = e.target.files?.[0]; if (!f) return; const r = new FileReader(); r.onload = () => onPick(r.result); r.readAsDataURL(f); }} />
      {zoom && src && <Lightbox src={src} onClose={() => setZoom(false)} />}
    </>
  );
}

/* ───────── Planilla (edición tipo Excel) ───────── */
function Planilla({ items, onUpdate }) {
  const [solo, setSolo] = useState(false);
  const PF0 = { q: "", relevadoPor: "", estado: "", sector: "", sectorResp: "", responsable: "", criticidad: "" };
  const [pf, setPf] = useState(PF0);
  const activos = Object.values(pf).some(Boolean) || solo;
  const rows = items.filter((h) =>
    (!solo || incompleto(h))
    && (!pf.relevadoPor || h.relevadoPor === pf.relevadoPor)
    && (!pf.estado || h.estado === pf.estado)
    && (!pf.sector || h.sector === pf.sector)
    && (!pf.sectorResp || h.sectorResp === pf.sectorResp)
    && (!pf.responsable || h.responsable === pf.responsable)
    && (!pf.criticidad || (h.criticidad || "Sin clasificar") === pf.criticidad)
    && (!pf.q || [h.descripcion, h.comentarios, h.sector, h.responsable, h.relevadoPor, h.sectorResp].some((x) => (x || "").toLowerCase().includes(pf.q.toLowerCase())))
  );
  const searchStyle = { borderRadius: 8, border: `1px solid ${C.border}`, background: C.card, padding: "6px 8px 6px 28px", fontSize: 12.5, color: C.ink, outline: "none", width: 180 };
  const th = { textAlign: "left", padding: "8px 10px", fontSize: 10, fontWeight: 700, letterSpacing: ".08em", textTransform: "uppercase", color: C.muted, whiteSpace: "nowrap", borderBottom: `1px solid ${C.border}`, position: "sticky", top: 0, background: C.page };
  const HF = ({ k, label, opts }) => {
    const on = !!pf[k];
    return (
      <th style={{ ...th, padding: "6px 8px" }}>
        <select value={pf[k]} onChange={(e) => setPf({ ...pf, [k]: e.target.value })}
          style={{ borderRadius: 6, border: `1px solid ${on ? C.blue : C.border}`, background: on ? C.blue : C.card, color: on ? "#fff" : C.muted, padding: "4px 6px", fontSize: 10, fontWeight: 700, letterSpacing: ".06em", textTransform: "uppercase", cursor: "pointer", outline: "none", maxWidth: 150 }}>
          <option value="" style={{ color: C.ink, background: "#fff", textTransform: "none", fontWeight: 400 }}>{label}</option>
          {opts.map((o) => <option key={o} value={o} style={{ color: C.ink, background: "#fff", textTransform: "none", fontWeight: 400 }}>{o}</option>)}
        </select>
      </th>
    );
  };
  const td = { padding: "6px 8px", borderBottom: `1px solid ${C.border}`, verticalAlign: "middle" };
  const cell = (miss) => ({ width: "100%", borderRadius: 6, border: `1px solid ${miss ? C.amber : C.border}`, background: miss ? "#FFFBEB" : C.card, padding: "6px 8px", fontSize: 13, color: C.ink, outline: "none", appearance: "none" });
  const Opt = ({ v, set, ph2, opts, miss }) => (
    <select value={v} onChange={(e) => set(e.target.value)} style={{ ...cell(miss), minWidth: 130, color: v ? C.ink : C.muted }}>
      <option value="">{ph2}</option>{opts.map((o) => <option key={o} value={o} style={{ color: C.ink }}>{o}</option>)}
    </select>
  );
  const OptCombo = ({ v, set, ph2, opts, miss }) => (
    <>
      <input list="dl-personas" value={v} onChange={(e) => set(e.target.value)} placeholder={ph2} style={{ ...cell(miss), minWidth: 130, color: v ? C.ink : C.muted }} />
      <datalist id="dl-personas">{opts.map((o) => <option key={o} value={o} />)}</datalist>
    </>
  );
  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%" }}>
      <div className="rec-planilla-toolbar">
        <div className="rec-toolbar-title">
          <p style={{ fontSize: 13, fontWeight: 600, color: C.ink }}>Completar en planilla</p>
          <p style={{ fontSize: 11.5, color: C.muted }}>Editá las celdas; las amarillas faltan.</p>
        </div>
        <div style={{ position: "relative", flex: 1, minWidth: 120 }}>
          <Search size={14} style={{ position: "absolute", left: 8, top: 9, color: C.muted, pointerEvents: "none" }} />
          <input value={pf.q} onChange={(e) => setPf({ ...pf, q: e.target.value })} placeholder="Buscar texto…" style={{ ...searchStyle, width: "100%" }} />
        </div>
        <span style={{ fontSize: 12, color: C.muted, whiteSpace: "nowrap" }}>{rows.length} hallazgo(s)</span>
        <button onClick={() => setSolo((v) => !v)} style={{ borderRadius: 999, padding: "6px 12px", fontSize: 12, fontWeight: 500, cursor: "pointer", border: `1px solid ${solo ? C.blue : C.border}`, background: solo ? C.blue : C.card, color: solo ? "#fff" : C.ink, whiteSpace: "nowrap" }}>Solo por completar</button>
        {activos && <button onClick={() => { setPf(PF0); setSolo(false); }} style={{ fontSize: 12, color: C.blue, background: "none", border: "none", cursor: "pointer", whiteSpace: "nowrap" }}>Limpiar</button>}
      </div>
      <div style={{ flex: 1, overflow: "auto" }}>
        <table style={{ borderCollapse: "collapse", width: "max-content", minWidth: "100%" }}>
          <thead><tr>
            <th style={th}>N°</th>
            <HF k="relevadoPor" label="Relevado por" opts={PERSONAS} />
            <th style={th}>Registro</th>
            <HF k="criticidad" label="Criticidad" opts={[...CRITICIDADES, "Sin clasificar"]} />
            <th style={th}>Descripción</th>
            <th style={th}>Antes</th>
            <HF k="sector" label="Sector" opts={SECTORES} />
            <HF k="sectorResp" label="Sector resp." opts={SECTOR_RESP} />
            <HF k="responsable" label="Responsable" opts={PERSONAS} />
            <HF k="estado" label="Estado" opts={ESTADOS} />
            <th style={th}>Comentarios</th><th style={th}>Después</th><th style={th}>Realización</th>
          </tr></thead>
          <tbody>
            {rows.map((h) => (
              <tr key={h.id} style={{ background: C.card }}>
                <td style={{ ...td, fontSize: 12, fontWeight: 600, color: C.muted, whiteSpace: "nowrap", fontVariantNumeric: "tabular-nums" }}>{fmtNro(h.numero)}</td>
                <td style={td}><OptCombo v={h.relevadoPor} set={(v) => onUpdate({ ...h, relevadoPor: v })} ph2="Relevado por" opts={PERSONAS} miss={false} /></td>
                <td style={{ ...td, fontSize: 12, color: C.muted, whiteSpace: "nowrap" }}>{h.fechaRegistro}</td>
                <td style={td}><Opt v={h.criticidad} set={(v) => onUpdate({ ...h, criticidad: v })} ph2="Sin clasificar" opts={CRITICIDADES} miss={!h.criticidad} /></td>
                <td style={td}><input value={h.descripcion} onChange={(e) => onUpdate({ ...h, descripcion: e.target.value })} placeholder="Descripción" style={{ ...cell(!h.descripcion), minWidth: 220 }} /></td>
                <td style={td}><ZoomImg src={h.fotoAntes} w={38} h={38} /></td>
                <td style={td}><Opt v={h.sector} set={(v) => onUpdate({ ...h, sector: v })} ph2="Sector" opts={SECTORES} miss={!h.sector} /></td>
                <td style={td}><Opt v={h.sectorResp} set={(v) => onUpdate({ ...h, sectorResp: v })} ph2="Área" opts={SECTOR_RESP} miss={false} /></td>
                <td style={td}><Opt v={h.responsable} set={(v) => onUpdate({ ...h, responsable: v })} ph2="Responsable" opts={PERSONAS} miss={!h.responsable} /></td>
                <td style={td}><select value={h.estado} onChange={(e) => onUpdate({ ...h, estado: e.target.value })} style={{ ...cell(false), minWidth: 130, fontWeight: 600, color: estadoPill[h.estado].fg, background: estadoPill[h.estado].bg, border: `1px solid ${estadoColor[h.estado]}` }}>{ESTADOS.map((s) => <option key={s} style={{ color: C.ink, background: "#fff" }}>{s}</option>)}</select></td>
                <td style={td}><input value={h.comentarios || ""} onChange={(e) => onUpdate({ ...h, comentarios: e.target.value })} placeholder="Comentarios de resolución" style={{ ...cell(false), minWidth: 200 }} /></td>
                <td style={td}><GridPhoto src={h.fotoDespues} onPick={(v) => onUpdate({ ...h, fotoDespues: v })} /></td>
                <td style={{ ...td, fontSize: 12, color: C.muted, whiteSpace: "nowrap" }}>{h.fechaCierre || "—"}</td>
              </tr>
            ))}
          </tbody>
        </table>
        {rows.length === 0 && <p style={{ textAlign: "center", color: C.muted, fontSize: 14, marginTop: 40 }}>No hay hallazgos por completar.</p>}
      </div>
    </div>
  );
}

/* ───────── Tablero ───────── */
const tipStyle = { background: "#fff", border: `1px solid ${C.border}`, borderRadius: 8, fontSize: 12 };
function BarCard({ title, data, colorFor, color, height, nota }) {
  const max = Math.max(1, ...data.map((d) => d.v));
  const ticks = Array.from({ length: max + 1 }, (_, i) => i);
  return (
    <div>
      <p style={{ fontSize: 13, fontWeight: 700, letterSpacing: ".06em", textTransform: "uppercase", color: C.ink }}>{title}</p>
      {nota && <p style={{ fontSize: 11, color: C.muted, marginTop: 2 }}>{nota}</p>}
      <div style={{ height: height || 230, marginTop: 10 }}>
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={data} margin={{ top: 4, right: 8, left: 0, bottom: 4 }}>
            <XAxis dataKey="name" interval={0} tick={{ fill: C.muted, fontSize: 11 }} axisLine={false} tickLine={false} />
            <YAxis allowDecimals={false} domain={[0, max]} ticks={ticks} tick={{ fill: C.muted, fontSize: 12 }} axisLine={false} tickLine={false} width={28} />
            <Tooltip cursor={{ fill: C.page }} contentStyle={tipStyle} />
            <Bar dataKey="v" radius={[3, 3, 0, 0]} fill={color || C.blue}>{colorFor && data.map((d) => <Cell key={d.name} fill={colorFor(d.name)} />)}</Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

function Tablero({ items, me }) {
  const [desde, setDesde] = useState(""); const [hasta, setHasta] = useState("");
  const f = items.filter((h) => (!desde || h.fechaRegistro >= desde) && (!hasta || h.fechaRegistro <= hasta));
  const total = f.length;
  const noAplica = f.filter((h) => h.estado === "No aplica").length;
  const relev = total - noAplica;
  const fin = f.filter((h) => h.estado === "Finalizado").length;
  const pct = relev ? Math.round((fin / relev) * 100) : 0;
  const incom = f.filter(incompleto).length;
  const altaAbiertas = f.filter((h) => h.criticidad === "Alta" && h.estado !== "Finalizado" && h.estado !== "No aplica").length;
  const finc = f.filter((h) => h.estado === "Finalizado" && h.fechaCierre);
  const avgDias = finc.length ? Math.round(finc.reduce((a, h) => a + (new Date(h.fechaCierre) - new Date(h.fechaRegistro)) / 86400000, 0) / finc.length) : null;

  const cc = { "Baja": C.green, "Media": C.amber, "Alta": C.red, "Sin clasif.": C.grey };
  const porEstado = ESTADOS.map((e) => ({ name: e, v: f.filter((h) => h.estado === e).length }));
  const porCrit = [...CRITICIDADES, ""].map((c) => ({ name: c || "Sin clasif.", v: f.filter((h) => h.criticidad === c).length }));
  // Top 5 por cantidad, mayor a menor (evita que se enciman los nombres)
  const porSector = SECTORES.map((s) => ({ name: s, v: f.filter((h) => h.sector === s).length })).filter((d) => d.v > 0).sort((a, b) => b.v - a.v).slice(0, 5);
  const porResp = PERSONAS.map((p) => ({ name: p, v: f.filter((h) => h.responsable === p).length })).filter((d) => d.v > 0).sort((a, b) => b.v - a.v).slice(0, 5);
  const porArea = SECTOR_RESP.map((a) => ({ name: a, v: f.filter((h) => h.sectorResp === a).length }));

  const inp = { borderRadius: 8, border: `1px solid ${C.border}`, background: C.card, padding: "6px 8px", fontSize: 14, color: C.ink, outline: "none" };
  const card = { borderRadius: 10, border: `1px solid ${C.border}`, background: C.card, padding: 12 };
  const kpi = (label, value, color) => (
    <div key={label} style={{ ...card, textAlign: "center" }}><Label>{label}</Label><p style={{ marginTop: 4, fontSize: 24, fontWeight: 600, color: color || C.ink, fontVariantNumeric: "tabular-nums" }}>{value}</p></div>
  );
  const hBar = (d) => Math.max(150, d.length * 30 + 30);

  return (
    <div style={{ padding: 16, display: "flex", flexDirection: "column", gap: 18 }}>
      <div style={{ display: "flex", alignItems: "flex-end", gap: 8 }}>
        <div><Label>Desde</Label><input type="date" value={desde} onChange={(e) => setDesde(e.target.value)} style={{ ...inp, marginTop: 4, display: "block" }} /></div>
        <div><Label>Hasta</Label><input type="date" value={hasta} onChange={(e) => setHasta(e.target.value)} style={{ ...inp, marginTop: 4, display: "block" }} /></div>
        {(desde || hasta) && <button onClick={() => { setDesde(""); setHasta(""); }} style={{ paddingBottom: 6, fontSize: 12, color: C.blue, background: "none", border: "none", cursor: "pointer" }}>limpiar</button>}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(120px,1fr))", gap: 8 }}>
        {kpi("Total", total)}
        {kpi("Abiertos", relev - fin, C.amber)}
        {kpi("Finalizados", fin, C.green)}
        {kpi("% resuelto", pct + "%")}
        {kpi("Alta abiertas", altaAbiertas, altaAbiertas ? C.red : C.ink)}
        {kpi("Por completar", incom, incom ? "#B45309" : C.ink)}
        {kpi("T. prom. cierre", avgDias === null ? "—" : avgDias + " d")}
      </div>
      {noAplica > 0 && <p style={{ marginTop: -8, fontSize: 12, color: C.muted }}>{noAplica} marcado(s) «No aplica», excluidos del % resuelto.</p>}

      <div className="rec-charts">
        <BarCard title="Por estado" data={porEstado} colorFor={(n) => estadoColor[n]} />
        <BarCard title="Por criticidad" data={porCrit} colorFor={(n) => cc[n]} />
        <BarCard title="Por sector" data={porSector} nota="Top 5 sectores con más hallazgos" />
        {puedeVerResponsable(me?.nombre) && <BarCard title="Por responsable" data={porResp} nota="Top 5 responsables con más hallazgos" />}
        <BarCard title="Por área responsable" data={porArea} />
      </div>
    </div>
  );
}

/* ───────── Card reutilizable (lista + kanban) ───────── */
function HallazgoCard({ h, onClick, showEstado = true }) {
  const inc = incompleto(h);
  return (
    <button onClick={onClick} style={{ width: "100%", flexShrink: 0, display: "flex", alignItems: "stretch", overflow: "hidden", borderRadius: 10, border: `1px solid ${C.border}`, background: C.card, textAlign: "left", cursor: "pointer", padding: 0 }}>
      <span style={{ width: 6, flexShrink: 0, background: critColor[h.criticidad] }} />
      <span style={{ margin: 8, height: 56, width: 56, flexShrink: 0, overflow: "hidden", borderRadius: 6, background: C.page }}>{h.fotoAntes && <img src={h.fotoAntes} alt="" style={{ height: "100%", width: "100%", objectFit: "cover" }} />}</span>
      <span style={{ minWidth: 0, flex: 1, padding: "8px 12px 8px 0" }}>
        <span style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
          <span style={{ fontSize: 10, fontWeight: 700, letterSpacing: ".06em", textTransform: "uppercase", color: C.muted }}>{h.numero != null && <span style={{ fontVariantNumeric: "tabular-nums" }}>{fmtNro(h.numero)} · </span>}{h.sector || "Sin sector"}</span>
          {showEstado && <span style={{ flexShrink: 0, borderRadius: 6, padding: "2px 7px", fontSize: 10, fontWeight: 600, background: estadoPill[h.estado].bg, color: estadoPill[h.estado].fg }}>{h.estado}</span>}
        </span>
        <span style={{ marginTop: 2, display: "-webkit-box", WebkitLineClamp: 2, WebkitBoxOrient: "vertical", overflow: "hidden", fontSize: 13.5, color: h.descripcion ? C.ink : C.grey }}>{h.descripcion || "Sin descripción"}</span>
        <span style={{ marginTop: 4, display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap", fontSize: 11, color: C.muted }}>
          <span>{h.responsable || "Sin responsable"} · {h.fechaRegistro}</span>
          {inc && <span style={{ display: "inline-flex", alignItems: "center", gap: 3, color: "#B45309" }}><AlertTriangle size={12} /> Por completar</span>}
        </span>
      </span>
    </button>
  );
}

/* ───────── Vista Kanban (columnas por estado, scroll independiente) ───────── */
function Kanban({ items, onOpen }) {
  const scrollRef = useRef(null);
  const [colIdx, setColIdx] = useState(0);

  const onScroll = () => {
    const el = scrollRef.current;
    if (!el) return;
    const colW = el.scrollWidth / ESTADOS_KANBAN.length;
    setColIdx(Math.round(el.scrollLeft / colW));
  };

  return (
    <div className="rec-kanban-wrap">
      {colIdx > 0 && <div className="rec-kanban-arrow rec-kanban-arrow-l">‹</div>}
      {colIdx < ESTADOS_KANBAN.length - 1 && <div className="rec-kanban-arrow rec-kanban-arrow-r">›</div>}
      <div ref={scrollRef} className="rec-kanban" style={{ flex: 1, minHeight: 0 }} onScroll={onScroll}>
        {ESTADOS_KANBAN.map((e) => {
          const col = items.filter((h) => h.estado === e);
          const color = estadoColor[e];
          return (
            <div key={e} className="rec-kanban-col">
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "10px 12px", borderTop: `3px solid ${color}`, borderTopLeftRadius: 10, borderTopRightRadius: 10 }}>
                <span style={{ fontSize: 12.5, fontWeight: 700, color: C.ink }}>{e}</span>
                <span style={{ fontSize: 11, fontWeight: 700, color: "#fff", background: color, borderRadius: 999, padding: "1px 8px", minWidth: 20, textAlign: "center" }}>{col.length}</span>
              </div>
              <div style={{ flex: 1, minHeight: 0, overflowY: "auto", padding: "0 8px 8px", display: "flex", flexDirection: "column", gap: 8 }}>
                {col.length === 0 ? <p style={{ fontSize: 12, color: C.muted, textAlign: "center", padding: "16px 0" }}>—</p>
                  : col.map((h) => <HallazgoCard key={h.id} h={h} onClick={() => onOpen(h)} showEstado={false} />)}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ───────── Vista No aplica (hallazgos descartados) ───────── */
function NoAplica({ items, onOpen }) {
  const col = items.filter((h) => h.estado === "No aplica").sort((a, b) => (b.fechaCierre || "").localeCompare(a.fechaCierre || ""));
  return (
    <div style={{ height: "100%", overflowY: "auto", padding: 12 }}>
      {col.length === 0
        ? <p style={{ textAlign: "center", color: C.muted, fontSize: 13, padding: "32px 0" }}>No hay hallazgos marcados como «No aplica».</p>
        : <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))", gap: 8, maxWidth: 1100, margin: "0 auto" }}>
            {col.map((h) => <HallazgoCard key={h.id} h={h} onClick={() => onOpen(h)} showEstado={false} />)}
          </div>}
    </div>
  );
}

/* ───────── App ───────── */
const F0 = { q: "", estado: "Todos", sector: "Todos", sectorResp: "Todos", responsable: "Todos", criticidad: "Todos", soloIncompletos: false, soloNuevos: false, desde: "", hasta: "" };

/* ───────── Trazabilidad (audit log) ───────── */
function Trazabilidad({ items }) {
  const [q, setQ] = useState("");
  const accionTxt = { crear: "Creó", modificar: "Modificó", borrar: "Borró" };
  const fmt = (iso) => { try { return new Date(iso).toLocaleString("es-AR", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" }); } catch { return iso; } };
  const rows = items.filter((r) => {
    if (!q.trim()) return true;
    const s = q.toLowerCase();
    return [r.usuario_nombre, r.usuario_email, r.campo, r.valor_anterior, r.valor_nuevo, r.accion].some((x) => (x || "").toLowerCase().includes(s));
  });
  const th = { textAlign: "left", padding: "8px 10px", fontSize: 10, fontWeight: 700, letterSpacing: ".08em", textTransform: "uppercase", color: C.muted, whiteSpace: "nowrap", borderBottom: `1px solid ${C.border}`, position: "sticky", top: 0, background: C.page };
  const td = { padding: "8px 10px", borderBottom: `1px solid ${C.border}`, fontSize: 13, color: C.ink, verticalAlign: "top" };
  const accColor = { crear: C.green, modificar: C.blue, borrar: C.red };
  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%" }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap", padding: "12px 16px", borderBottom: `1px solid ${C.border}`, background: C.card }}>
        <div><p style={{ fontSize: 13, fontWeight: 600, color: C.ink }}>Trazabilidad</p><p style={{ fontSize: 11.5, color: C.muted }}>Quién cambió qué y cuándo.</p></div>
        <div style={{ position: "relative" }}>
          <Search size={14} style={{ position: "absolute", left: 8, top: 9, color: C.muted, pointerEvents: "none" }} />
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Buscar usuario, campo, valor…" style={{ borderRadius: 8, border: `1px solid ${C.border}`, background: C.card, padding: "7px 10px 7px 28px", fontSize: 13, color: C.ink, outline: "none", minWidth: 240 }} />
        </div>
        <span style={{ fontSize: 12, color: C.muted, whiteSpace: "nowrap" }}>{rows.length} registro(s)</span>
      </div>
      <div style={{ flex: 1, overflow: "auto" }}>
        <table style={{ borderCollapse: "collapse", width: "max-content", minWidth: "100%" }}>
          <thead><tr>
            <th style={th}>Fecha y hora</th><th style={th}>Usuario</th><th style={th}>Acción</th>
            <th style={th}>Campo</th><th style={th}>Antes</th><th style={th}>Después</th>
          </tr></thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.id} style={{ background: C.card }}>
                <td style={{ ...td, whiteSpace: "nowrap", color: C.muted }}>{fmt(r.creado_en)}</td>
                <td style={{ ...td, fontWeight: 600, whiteSpace: "nowrap" }}>{r.usuario_nombre || r.usuario_email || "—"}</td>
                <td style={td}><span style={{ borderRadius: 6, padding: "2px 8px", fontSize: 11, fontWeight: 600, color: "#fff", background: accColor[r.accion] || C.muted }}>{accionTxt[r.accion] || r.accion}</span></td>
                <td style={{ ...td, whiteSpace: "nowrap" }}>{r.campo || "—"}</td>
                <td style={{ ...td, color: C.muted }}>{r.valor_anterior || "—"}</td>
                <td style={td}>{r.valor_nuevo || "—"}</td>
              </tr>
            ))}
          </tbody>
        </table>
        {rows.length === 0 && <p style={{ textAlign: "center", color: C.muted, fontSize: 14, marginTop: 40 }}>Sin registros todavía.</p>}
      </div>
    </div>
  );
}

/* ───────── Login (email + password) ───────── */
function Login() {
  const [email, setEmail] = useState("");
  const [pass, setPass] = useState("");
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);
  const entrar = async () => {
    setErr(""); setBusy(true);
    try { await api.signIn(email.trim(), pass); }
    catch (e) { setErr(e?.message === "Invalid login credentials" ? "Mail o contraseña incorrectos." : (e?.message || "No se pudo iniciar sesión.")); }
    finally { setBusy(false); }
  };
  const onKey = (e) => { if (e.key === "Enter" && email && pass) entrar(); };
  return (
    <div style={{ minHeight: "100dvh", display: "flex", alignItems: "center", justifyContent: "center", background: C.page, padding: 16, fontFamily: FONT }}>
      <div style={{ width: "100%", maxWidth: 360, background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: 24, boxShadow: "0 8px 30px rgba(0,0,0,.06)" }}>
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 6, marginBottom: 18 }}>
          <img src={LOGO} alt="Biomás" style={{ height: 40 }} />
          <p style={{ fontSize: 10, fontWeight: 700, letterSpacing: ".16em", textTransform: "uppercase", color: C.orange }}>Recorrida · {PLANTA}</p>
          <h1 style={{ fontSize: 18, fontWeight: 700, color: C.ink }}>Iniciar sesión</h1>
        </div>
        <Label>Email</Label>
        <input value={email} onChange={(e) => setEmail(e.target.value)} onKeyDown={onKey} type="email" autoComplete="username" placeholder="tu mail" style={{ ...fieldStyle, marginBottom: 12 }} />
        <Label>Contraseña</Label>
        <input value={pass} onChange={(e) => setPass(e.target.value)} onKeyDown={onKey} type="password" autoComplete="current-password" placeholder="••••••••" style={{ ...fieldStyle }} />
        {err && <p style={{ marginTop: 10, fontSize: 12.5, color: C.red }}>{err}</p>}
        <button disabled={busy || !email || !pass} onClick={entrar}
          style={{ width: "100%", marginTop: 16, borderRadius: 8, border: "none", padding: "12px 0", fontSize: 14, fontWeight: 600, cursor: busy || !email || !pass ? "not-allowed" : "pointer", background: busy || !email || !pass ? C.border : C.blue, color: busy || !email || !pass ? C.muted : "#fff" }}>
          {busy ? "Ingresando…" : "Ingresar"}
        </button>
        <p style={{ marginTop: 14, fontSize: 11, color: C.muted, textAlign: "center" }}>¿Olvidaste tu contraseña? Pedile a un administrador que te la reinicie.</p>
      </div>
    </div>
  );
}

/* ───────── Cambiar contraseña ───────── */
function ChangePassword({ onClose }) {
  const [p1, setP1] = useState("");
  const [p2, setP2] = useState("");
  const [msg, setMsg] = useState("");
  const [busy, setBusy] = useState(false);
  const guardar = async () => {
    setMsg("");
    if (p1.length < 6) { setMsg("La contraseña debe tener al menos 6 caracteres."); return; }
    if (p1 !== p2) { setMsg("Las contraseñas no coinciden."); return; }
    setBusy(true);
    try { await api.changePassword(p1); setMsg("ok"); }
    catch (e) { setMsg(e?.message || "No se pudo cambiar la contraseña."); }
    finally { setBusy(false); }
  };
  return (
    <div style={{ position: "fixed", inset: 0, zIndex: 40, display: "flex", alignItems: "center", justifyContent: "center", background: "rgba(0,0,0,.4)", padding: 16 }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} style={{ width: "100%", maxWidth: 360, background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: 22 }}>
        <h2 style={{ fontSize: 16, fontWeight: 700, color: C.ink, marginBottom: 14 }}>Cambiar contraseña</h2>
        {msg === "ok" ? (
          <>
            <p style={{ fontSize: 13.5, color: C.green, marginBottom: 16 }}>Contraseña actualizada correctamente.</p>
            <button onClick={onClose} style={{ width: "100%", borderRadius: 8, border: "none", padding: "11px 0", fontSize: 14, fontWeight: 600, cursor: "pointer", background: C.blue, color: "#fff" }}>Listo</button>
          </>
        ) : (
          <>
            <Label>Nueva contraseña</Label>
            <input value={p1} onChange={(e) => setP1(e.target.value)} type="password" autoComplete="new-password" style={{ ...fieldStyle, marginBottom: 12 }} />
            <Label>Repetir contraseña</Label>
            <input value={p2} onChange={(e) => setP2(e.target.value)} type="password" autoComplete="new-password" style={{ ...fieldStyle }} />
            {msg && <p style={{ marginTop: 10, fontSize: 12.5, color: C.red }}>{msg}</p>}
            <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
              <button onClick={onClose} style={{ flex: 1, borderRadius: 8, border: `1px solid ${C.border}`, background: C.card, padding: "11px 0", fontSize: 14, color: C.ink, cursor: "pointer" }}>Cancelar</button>
              <button disabled={busy} onClick={guardar} style={{ flex: 1, borderRadius: 8, border: "none", padding: "11px 0", fontSize: 14, fontWeight: 600, cursor: busy ? "not-allowed" : "pointer", background: C.blue, color: "#fff" }}>{busy ? "Guardando…" : "Guardar"}</button>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

/* ───────── Botón flotante directivo (idle configurable) ───────── */
const IDLE_SEGUNDOS = 7; // ← configurable

function BotonDirectivo({ onSave, defaultRelevadoPor }) {
  const [esMobile, setEsMobile] = useState(() => typeof window !== "undefined" && window.matchMedia("(max-width:640px)").matches);
  const [visible, setVisible] = useState(true);
  const [fase, setFase] = useState(null); // null | "confirmar" | "guardando"
  const [fotoDataUrl, setFotoDataUrl] = useState(null);
  const refInput = useRef(null);
  const timerRef = useRef(null);
  const visibleRef = useRef(true);
  const faseRef = useRef(null);
  visibleRef.current = visible;
  faseRef.current = fase;

  // Solo en celular: escuchar cambios de ancho (rotación, resize)
  useEffect(() => {
    const mq = window.matchMedia("(max-width:640px)");
    const handler = (e) => setEsMobile(e.matches);
    mq.addEventListener("change", handler);
    return () => mq.removeEventListener("change", handler);
  }, []);

  // Programa la reaparición tras la última actividad.
  // Solo cuenta cuando el overlay está oculto y no hay confirmación abierta.
  const programarAparicion = () => {
    if (visibleRef.current || faseRef.current !== null) return;
    clearTimeout(timerRef.current);
    timerRef.current = setTimeout(() => setVisible(true), IDLE_SEGUNDOS * 1000);
  };

  useEffect(() => {
    const events = ["touchstart", "pointerdown", "keydown", "scroll"];
    events.forEach((e) => window.addEventListener(e, programarAparicion, { passive: true }));
    return () => {
      events.forEach((e) => window.removeEventListener(e, programarAparicion));
      clearTimeout(timerRef.current);
    };
  }, []);

  // Cada vez que el overlay se oculta, arranca la cuenta para reaparecer.
  useEffect(() => {
    if (!visible && fase === null) {
      clearTimeout(timerRef.current);
      timerRef.current = setTimeout(() => setVisible(true), IDLE_SEGUNDOS * 1000);
    }
  }, [visible, fase]);

  const abrirCamara = () => { setVisible(false); setFase(null); setFotoDataUrl(null); refInput.current?.click(); };

  const onFoto = (e) => {
    const f = e.target.files?.[0];
    e.target.value = "";
    if (!f) return;
    const r = new FileReader();
    r.onload = () => { setFotoDataUrl(r.result); setFase("confirmar"); };
    r.readAsDataURL(f);
  };

  const guardar = async (otraFoto) => {
    setFase("guardando");
    await onSave({ fotoAntes: fotoDataUrl });
    setFotoDataUrl(null);
    if (otraFoto) { setFase(null); refInput.current?.click(); }
    else { setFase(null); }
  };

  if (!esMobile) return null;

  return (
    <>
      {/* Input oculto — capture=environment abre la cámara trasera directamente */}
      <input ref={refInput} type="file" accept="image/*" capture="environment" style={{ display: "none" }} onChange={onFoto} />

      {/* Overlay de confirmación */}
      {fase === "confirmar" && (
        <div style={{ position: "fixed", inset: 0, zIndex: 60, background: "rgba(15,23,42,.80)", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: 24, gap: 16 }}>
          <img src={fotoDataUrl} alt="preview" style={{ maxHeight: "55vh", maxWidth: "100%", borderRadius: 12, objectFit: "contain", boxShadow: "0 8px 32px rgba(0,0,0,.5)" }} />
          <p style={{ color: "#fff", fontSize: 15, fontWeight: 600, textAlign: "center" }}>¿Registrar este hallazgo?</p>
          <div style={{ display: "flex", flexDirection: "column", gap: 10, width: "100%", maxWidth: 320 }}>
            <button onClick={() => guardar(false)} style={{ borderRadius: 10, border: "none", padding: "13px 0", fontSize: 15, fontWeight: 700, background: C.blue, color: "#fff", cursor: "pointer" }}>
              ✓ Registrar
            </button>
            <button onClick={() => guardar(true)} style={{ borderRadius: 10, border: `1px solid rgba(255,255,255,.35)`, padding: "13px 0", fontSize: 15, fontWeight: 600, background: "rgba(255,255,255,.12)", color: "#fff", cursor: "pointer" }}>
              + Registrar y sacar otra foto
            </button>
            <button onClick={() => { setFase(null); setFotoDataUrl(null); }} style={{ borderRadius: 10, border: "none", padding: "11px 0", fontSize: 14, fontWeight: 500, background: "none", color: "rgba(255,255,255,.6)", cursor: "pointer" }}>
              Cancelar
            </button>
          </div>
        </div>
      )}

      {fase === "guardando" && (
        <div style={{ position: "fixed", inset: 0, zIndex: 60, background: "rgba(15,23,42,.70)", display: "flex", alignItems: "center", justifyContent: "center" }}>
          <p style={{ color: "#fff", fontSize: 15, fontWeight: 600 }}>Guardando…</p>
        </div>
      )}

      {/* Botón flotante — aparece tras idle */}
      {visible && fase === null && (
        <div onClick={() => setVisible(false)}
          style={{ position: "fixed", inset: 0, zIndex: 50, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", background: "rgba(15,23,42,.18)", backdropFilter: "blur(1px)" }}>
          <div onClick={(e) => e.stopPropagation()} style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 18 }}>
            <button onClick={abrirCamara}
              style={{ width: 180, height: 180, borderRadius: 999, border: "none", background: C.blue, color: "#fff", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 10, cursor: "pointer", boxShadow: "0 8px 36px rgba(0,136,206,.55)", fontSize: 16, fontWeight: 700, lineHeight: 1.2, textAlign: "center", padding: "0 18px" }}>
              <Camera size={40} />
              Registrar<br />hallazgo
            </button>
            <p style={{ fontSize: 12, color: "#fff", opacity: .9, margin: 0, textShadow: "0 1px 3px rgba(0,0,0,.4)" }}>Tocá fuera para cerrar</p>
          </div>
        </div>
      )}
    </>
  );
}

/* ───────── App ───────── */
export default function App() {
  const [sess, setSess] = useState(undefined); // undefined=verificando, null=sin sesión, obj=logueado
  const [me, setMe] = useState(null);          // { nombre, email } de la persona logueada
  const [showPass, setShowPass] = useState(false);
  const [items, setItems] = useState([]);
  const [cargando, setCargando] = useState(true);
  const recargar = () => api.loadHallazgos().then((d) => { setItems(d); setCargando(false); }).catch((e) => { console.error(e); setCargando(false); });

  // Sesión: estado inicial + escucha de cambios (login/logout/refresh)
  useEffect(() => {
    api.getSession().then(setSess);
    const off = api.onAuthChange(setSess);
    return off;
  }, []);

  // Al haber sesión: traer la persona logueada, cargar datos y suscribir realtime
  useEffect(() => {
    if (!sess) { setMe(null); return; }
    api.currentPersona().then(setMe).catch(() => setMe(null));
    recargar();
    const off = api.subscribe(recargar);
    return off;
  }, [sess]);

  // Trazabilidad (audit log)
  const [audit, setAudit] = useState([]);
  useEffect(() => {
    if (!sess) return;
    api.loadAuditoria().then(setAudit);
    const off = api.subscribeAuditoria(() => api.loadAuditoria().then(setAudit));
    return off;
  }, [sess]);
  const [tab, setTab] = useState("recorrida");
  // Aviso de hallazgos nuevos (criterio B): set congelado al entrar a Hallazgos
  const [nuevosIds, setNuevosIds] = useState(() => new Set());
  const vistaTomada = useRef(false);

  // Al entrar a Hallazgos con datos cargados: tomar UNA foto de los "nuevos"
  // (createdAt > última vista) y marcar visto en segundo plano.
  useEffect(() => {
    if (tab !== "recorrida" || cargando || vistaTomada.current) return;
    vistaTomada.current = true;
    api.getUltimaVista().then((ultima) => {
      if (ultima) {
        const ids = items.filter((h) => h.createdAt && h.createdAt > ultima).map((h) => h.id);
        setNuevosIds(new Set(ids));
      } // primera vez (ultima=null): todo ya visto -> set vacío
      api.marcarVistoHallazgos();
    }).catch((e) => console.error("vista", e));
  }, [tab, cargando, items]);

  const [flt, setFlt] = useState(F0);
  const [openFilters, setOpenFilters] = useState(false);
  const [nuevo, setNuevo] = useState(false);
  const [sel, setSel] = useState(null);
  const set = (k, v) => setFlt((p) => ({ ...p, [k]: v }));
  const activos = ["sector", "sectorResp", "responsable", "criticidad"].filter((k) => flt[k] !== "Todos").length + (flt.desde ? 1 : 0) + (flt.hasta ? 1 : 0) + (flt.soloIncompletos ? 1 : 0) + (flt.soloNuevos ? 1 : 0);

  const lista = items.filter((h) => {
    const q = flt.q.trim().toLowerCase();
    const enTexto = !q || [h.descripcion, h.sector, h.sectorResp, h.responsable, h.relevadoPor, h.comentarios].some((x) => (x || "").toLowerCase().includes(q));
    return enTexto
      && (flt.estado === "Todos" || h.estado === flt.estado)
      && (flt.sector === "Todos" || h.sector === flt.sector)
      && (flt.sectorResp === "Todos" || h.sectorResp === flt.sectorResp)
      && (flt.responsable === "Todos" || h.responsable === flt.responsable)
      && (flt.criticidad === "Todos" || (h.criticidad || "Sin clasificar") === flt.criticidad)
      && (!flt.soloIncompletos || incompleto(h))
      && (!flt.soloNuevos || nuevosIds.has(h.id))
      && (!flt.desde || h.fechaRegistro >= flt.desde) && (!flt.hasta || h.fechaRegistro <= flt.hasta);
  }).sort((a, b) => b.fechaRegistro.localeCompare(a.fechaRegistro));

  const upd = (u) => {
    setItems((arr) => arr.map((i) => (i.id === u.id ? u : i)));
    setSel((s) => (s && s.id === u.id ? u : s));
    api.updateHallazgo(u).then(recargar).catch((e) => console.error("update", e));
  };

  const exportar = () => {
    const rows = items.map((h) => ({
      "N°": fmtNro(h.numero),
      Planta: h.planta, Sector: h.sector, "Sector responsable": h.sectorResp, "Relevado por": h.relevadoPor, Responsable: h.responsable,
      Criticidad: h.criticidad || "Sin clasificar", Estado: h.estado, "Fecha registro": h.fechaRegistro, "Fecha cierre": h.fechaCierre || "",
      Descripción: h.descripcion, Comentarios: h.comentarios || "", "Foto antes": h.fotoAntes ? "Sí" : "No", "Foto después": h.fotoDespues ? "Sí" : "No",
      Completo: incompleto(h) ? "No" : "Sí",
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [{ wch: 8 }, { wch: 9 }, { wch: 14 }, { wch: 16 }, { wch: 18 }, { wch: 18 }, { wch: 12 }, { wch: 13 }, { wch: 13 }, { wch: 13 }, { wch: 40 }, { wch: 30 }, { wch: 10 }, { wch: 11 }, { wch: 9 }];
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Recorrida"); XLSX.writeFile(wb, "recorrida_biomas.xlsx");
  };

  const chev = "url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%2364748B' stroke-width='2'%3E%3Cpath d='M6 9l6 6 6-6'/%3E%3C/svg%3E\")";
  const fSelStyle = { height: 34, boxSizing: "border-box", borderRadius: 8, border: `1px solid ${C.border}`, background: C.card, backgroundImage: chev, backgroundRepeat: "no-repeat", backgroundPosition: "right 8px center", padding: "0 26px 0 9px", fontSize: 13, outline: "none", appearance: "none", cursor: "pointer" };
  const FSel = ({ k, label, opts, w }) => (
    <select value={flt[k]} onChange={(e) => set(k, e.target.value)} style={{ ...fSelStyle, minWidth: w || 108, color: flt[k] === "Todos" ? C.muted : C.ink }}>
      <option value="Todos">{label}</option>{opts.map((o) => <option key={o} value={o} style={{ color: C.ink }}>{o}</option>)}
    </select>
  );

  if (sess === undefined) return (
    <div style={{ minHeight: "100dvh", display: "flex", alignItems: "center", justifyContent: "center", background: C.page, color: C.muted, fontFamily: FONT, fontSize: 14 }}>Cargando…</div>
  );
  if (sess === null) return <Login />;

  return (
    <div style={{ position: "relative", margin: "0 auto", width: "100%", maxWidth: "100%", height: "100dvh", display: "flex", flexDirection: "column", background: C.page, color: C.ink, fontFamily: FONT }}>
      <style>{`
        html,body,#root{ margin:0; padding:0; width:100%; max-width:100%; background:${C.page}; font-family:${FONT}; }
        body{ display:block; overflow-x:hidden; }
        #root{ max-width:none; padding:0; text-align:left; min-height:100dvh; }
        *{ box-sizing:border-box; }
        .rec-fgrid{ display:grid; grid-template-columns:1fr 1fr; gap:8px; }
        @media(min-width:768px){ .rec-fgrid{ grid-template-columns:repeat(3,1fr); } }
        @media(min-width:1100px){ .rec-fgrid{ grid-template-columns:repeat(6,1fr); } }
        .rec-newbtn{ align-self:stretch; }
        .rec-4{ display:grid; grid-template-columns:1fr; gap:10px; }
        @media(min-width:768px){ .rec-4{ grid-template-columns:repeat(4,1fr); } }
        .rec-2col{ display:grid; grid-template-columns:1fr 1fr; gap:10px; }
        /* ── Nuevo hallazgo: en mobile, grid de campos ocupa todo el ancho (50%/50%) y la descripción baja debajo ── */
        .rec-nh-grid{ flex:0 0 60% !important; }
        @media(max-width:640px){
          .rec-nh-row{ display:block !important; }
          .rec-nh-grid{ width:100% !important; min-width:0 !important; margin-bottom:14px; }
          .rec-nh-grid > div{ min-width:0 !important; }
          .rec-nh-desc{ width:100% !important; min-width:0 !important; }
        }
        .rec-list{ display:grid; grid-template-columns:1fr; gap:10px; align-items:start; }
        @media(min-width:680px){ .rec-list{ grid-template-columns:repeat(2,1fr); } }
        @media(min-width:1024px){ .rec-list{ grid-template-columns:repeat(3,1fr); } }
        @media(min-width:1500px){ .rec-list{ grid-template-columns:repeat(4,1fr); } }
        .rec-charts{ display:grid; grid-template-columns:1fr; gap:18px; }
        @media(min-width:768px){ .rec-charts{ grid-template-columns:1fr 1fr; } }
        /* ── Kanban mobile: una columna visible con snap ── */
        .rec-kanban{
          display:flex; gap:10px; overflow-x:auto; padding:8px 12px 12px;
          scroll-snap-type:x mandatory; -webkit-overflow-scrolling:touch;
          align-items:stretch; scrollbar-width:none;
        }
        .rec-kanban::-webkit-scrollbar{ display:none; }
        .rec-kanban-col{
          flex:0 0 calc(100vw - 28px); min-width:0; max-width:calc(100vw - 28px);
          scroll-snap-align:center; display:flex; flex-direction:column;
          background:#EEF2F6; border-radius:10px;
        }
        @media(min-width:640px){
          .rec-kanban-col{ flex:1 1 0; max-width:none; }
        }
        /* flechas laterales kanban (solo mobile) */
        .rec-kanban-wrap{ position:relative; flex:1; min-height:0; display:flex; flex-direction:column; }
        .rec-kanban-arrow{
          position:absolute; top:50%; transform:translateY(-50%);
          width:22px; height:44px; display:flex; align-items:center; justify-content:center;
          background:rgba(255,255,255,0.72); border-radius:999px;
          color:${C.muted}; font-size:14px; pointer-events:none; z-index:5;
          box-shadow:0 1px 4px rgba(0,0,0,.10);
        }
        .rec-kanban-arrow-l{ left:2px; }
        .rec-kanban-arrow-r{ right:2px; }
        @media(min-width:640px){ .rec-kanban-arrow{ display:none; } }
        /* ── Planilla toolbar inline en mobile ── */
        .rec-planilla-toolbar{
          display:flex; align-items:center; gap:8px; flex-wrap:wrap;
          padding:10px 12px; border-bottom:1px solid ${C.border}; background:${C.card};
        }
        .rec-planilla-toolbar .rec-toolbar-title{ display:flex; flex-direction:column; }
        @media(max-width:639px){
          .rec-planilla-toolbar .rec-toolbar-title{ display:none; }
        }
        /* ── Header mobile: ocultar nombre usuario ── */
        .rec-username{ display:none; }
        @media(min-width:640px){ .rec-username{ display:flex; } }
        /* ── Nav: padding-bottom safe area iOS ── */
        .rec-nav{
          display:grid; grid-template-columns:repeat(5,1fr);
          border-top:1px solid ${C.border}; background:${C.card};
          padding-bottom:env(safe-area-inset-bottom,0px);
        }
      `}</style>
      <header style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, borderBottom: `1px solid ${C.border}`, background: C.card, padding: "8px 12px", flexShrink: 0 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, minWidth: 0, overflow: "hidden" }}>
          <img src={LOGO} alt="Biomás" style={{ height: 28, width: "auto", flexShrink: 0 }} />
          <p style={{ fontSize: 12, fontWeight: 700, letterSpacing: ".12em", textTransform: "uppercase", color: C.blue, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>Recorrida · {PLANTA}</p>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 6, flexShrink: 0 }}>
          <button onClick={exportar} title="Exportar a Excel" style={{ display: "flex", alignItems: "center", gap: 5, borderRadius: 8, border: `1px solid ${C.blue}`, background: C.card, padding: "6px 10px", fontSize: 12, color: C.blue, cursor: "pointer" }}><Download size={15} /> Excel</button>
          <span style={{ width: 1, height: 20, background: C.border }} />
          <button onClick={() => setShowPass(true)} title="Cambiar contraseña" style={{ display: "flex", alignItems: "center", borderRadius: 8, border: `1px solid ${C.border}`, background: C.card, padding: "6px 7px", fontSize: 12, color: C.muted, cursor: "pointer" }}><KeyRound size={15} /></button>
          <button onClick={() => api.signOut()} title="Cerrar sesión" style={{ display: "flex", alignItems: "center", borderRadius: 8, border: `1px solid ${C.border}`, background: C.card, padding: "6px 7px", fontSize: 12, color: C.muted, cursor: "pointer" }}><LogOut size={15} /></button>
        </div>
      </header>

      {showPass && <ChangePassword onClose={() => setShowPass(false)} />}

      {tab === "recorrida" && <>
        <div style={{ display: "flex", flexWrap: "wrap", alignItems: "center", gap: 8, padding: "12px 12px 0" }}>
          <div style={{ position: "relative", flex: "1 1 150px", minWidth: 140 }}>
            <Search size={15} style={{ position: "absolute", left: 10, top: 9, color: C.muted, pointerEvents: "none" }} />
            <input value={flt.q} onChange={(e) => set("q", e.target.value)} placeholder="Buscar…"
              style={{ width: "100%", boxSizing: "border-box", height: 34, borderRadius: 8, border: `1px solid ${C.border}`, background: C.card, padding: "0 8px 0 30px", fontSize: 13, color: C.ink, outline: "none" }} />
          </div>
          <FSel k="sector" label="Sector" opts={SECTORES} w={96} />
          <FSel k="sectorResp" label="Sector resp." opts={SECTOR_RESP} w={124} />
          <FSel k="responsable" label="Responsable" opts={PERSONAS} w={118} />
          <FSel k="criticidad" label="Criticidad" opts={[...CRITICIDADES, "Sin clasificar"]} w={104} />
          <input type="date" value={flt.desde} onChange={(e) => set("desde", e.target.value)} title="Desde"
            style={{ height: 34, boxSizing: "border-box", borderRadius: 8, border: `1px solid ${C.border}`, background: C.card, padding: "0 8px", fontSize: 12.5, color: flt.desde ? C.ink : C.muted, outline: "none" }} />
          <input type="date" value={flt.hasta} onChange={(e) => set("hasta", e.target.value)} title="Hasta"
            style={{ height: 34, boxSizing: "border-box", borderRadius: 8, border: `1px solid ${C.border}`, background: C.card, padding: "0 8px", fontSize: 12.5, color: flt.hasta ? C.ink : C.muted, outline: "none" }} />
          <button onClick={() => nuevosIds.size && set("soloNuevos", !flt.soloNuevos)} disabled={!nuevosIds.size}
            style={{ height: 34, display: "inline-flex", alignItems: "center", gap: 6, borderRadius: 999, padding: "0 12px", fontSize: 12.5, fontWeight: nuevosIds.size ? 500 : 400, whiteSpace: "nowrap", cursor: nuevosIds.size ? "pointer" : "default", border: `1px solid ${flt.soloNuevos ? C.blue : (nuevosIds.size ? C.orange : C.border)}`, background: flt.soloNuevos ? "#E6F1FB" : (nuevosIds.size ? "#FDECE0" : C.card), color: flt.soloNuevos ? C.blueD : (nuevosIds.size ? C.orange : C.ink), opacity: nuevosIds.size ? 1 : .5 }}>
            Solo nuevos{nuevosIds.size ? ` (${nuevosIds.size})` : ""}
          </button>
          {(activos > 0 || flt.q) && <button onClick={() => setFlt(F0)} style={{ fontSize: 12, color: C.blue, background: "none", border: "none", cursor: "pointer", whiteSpace: "nowrap" }}>Limpiar</button>}
        </div>

        <button onClick={() => setNuevo(true)} className="rec-newbtn" style={{ margin: "8px 12px 0", display: "flex", alignItems: "center", justifyContent: "center", gap: 8, borderRadius: 8, border: "none", background: C.blue, color: "#fff", padding: "11px 0", fontSize: 14, fontWeight: 600, cursor: "pointer" }}><Plus size={18} /> Nuevo hallazgo</button>

        <Kanban items={lista} onOpen={setSel} />
      </>}

      {tab === "planilla" && <main style={{ flex: 1, overflow: "hidden" }}><Planilla items={items} onUpdate={upd} /></main>}
      {tab === "tablero" && <main style={{ flex: 1, overflowY: "auto" }}><Tablero items={items} me={me} /></main>}
      {tab === "noaplica" && <main style={{ flex: 1, overflow: "hidden" }}><NoAplica items={items} onOpen={setSel} /></main>}
      {tab === "trazabilidad" && <main style={{ flex: 1, overflow: "hidden" }}><Trazabilidad items={audit} /></main>}

      <nav className="rec-nav">
        {[["recorrida", "Hallazgos", ClipboardList], ["planilla", "Planilla", Table2], ["tablero", "Dashboard", BarChart3], ["noaplica", "No aplica", Ban], ["trazabilidad", "Historial", History]].map(([k, label, Icon]) => (
          <button key={k} onClick={() => setTab(k)} style={{ position: "relative", display: "flex", flexDirection: "column", alignItems: "center", gap: 2, padding: "9px 0", fontSize: 11, fontWeight: 500, border: "none", background: "none", cursor: "pointer", color: tab === k ? C.blue : C.muted }}>
            <span style={{ position: "relative", display: "inline-flex" }}>
              <Icon size={20} />
              {k === "recorrida" && nuevosIds.size > 0 && <span style={{ position: "absolute", top: -5, right: -9, minWidth: 15, height: 15, padding: "0 3px", display: "flex", alignItems: "center", justifyContent: "center", borderRadius: 999, background: C.orange, fontSize: 9, fontWeight: 700, color: "#fff", lineHeight: 1 }}>{nuevosIds.size}</span>}
            </span>
            {label}</button>
        ))}
      </nav>

      {nuevo && <NuevoHallazgo onClose={() => setNuevo(false)} defaultRelevadoPor={me?.nombre || ""} onSave={(h) => { setNuevo(false); api.createHallazgo(h).then(recargar).catch((e) => console.error("create", e)); }} />}
      {sel && <Detalle h={items.find((i) => i.id === sel.id) || sel} onClose={() => setSel(null)} onUpdate={upd} />}

      {/* Botón flotante directivo — solo cuando no hay modal abierto */}
      {!nuevo && !sel && !showPass && (
        <BotonDirectivo
          defaultRelevadoPor={me?.nombre || ""}
          onSave={(campos) => {
            const h = { ...campos, id: Date.now(), planta: PLANTA, estado: "No comenzado", fechaRegistro: hoy(), fechaCierre: null, fotoDespues: null, comentarios: "", sector: "", sectorResp: "", responsable: "", relevadoPor: me?.nombre || "", criticidad: "Media", descripcion: "" };
            return api.createHallazgo(h).then(recargar).catch((e) => console.error("directivo create", e));
          }}
        />
      )}
    </div>
  );
}
