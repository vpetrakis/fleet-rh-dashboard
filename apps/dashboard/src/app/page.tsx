"use client";

import React, {
  useState, useEffect, useCallback, useRef, useMemo,
} from "react";
import {
  UploadCloud, Ship, Activity, Server, Cpu, Database,
  ShieldCheck, AlertTriangle, RefreshCw, Clock,
  TrendingUp, ChevronRight, CheckCircle2, XCircle,
  Zap, Gauge, BarChart3, Layers, Bell, ChevronDown, ChevronUp,
} from "lucide-react";

// ─── Types ────────────────────────────────────────────────────────────────────

interface MatrixRow {
  component: string;
  limit: number;
  hours: string[];
}
interface AuxRow {
  component: string;
  limit: number;
  dg_hours: string[][];
}
interface Matrices {
  main_engine: MatrixRow[];
  dg_general:  MatrixRow[];
  aux_engine:  AuxRow[];
}
interface VesselConfig {
  me_cyls:  number;
  dg_count: number;
  aux_cyls: number;
}
interface DiagAlert {
  severity:  "critical" | "warning";
  section:   string;
  component: string;
  slot:      number;
  hours:     number;
  limit:     number;
  pct:       number;
  message:   string;
}
interface VesselData {
  status:       string;
  vessel:       string;
  is_fallback?: boolean;
  config:       VesselConfig;
  integrity:    number;
  diagnostics:  DiagAlert[];
  matrices:     Matrices;
  _upload?:     { original_filename: string; bytes: number; uploaded_at: string };
  _saved_at?:   string;
}
type ActiveTab = "ME" | "DG_GEN" | "DG_CYL";

// ─── Design tokens ────────────────────────────────────────────────────────────
const T = {
  bg:          "#06080e",
  surface:     "#0b0f1a",
  surfaceHigh: "#0f1524",
  border:      "rgba(255,255,255,0.06)",
  gold:        "#c9a84c",
  goldDim:     "#7a6128",
  goldFaint:   "rgba(201,168,76,0.07)",
  amber:       "#f5a623",
  rose:        "#f04060",
  emerald:     "#34d399",
  cyanDim:     "#1a7a72",
  text:        "#dde3ef",
  textMuted:   "#4e5d78",
  textDim:     "#2a3349",
} as const;
const MONO = `"JetBrains Mono", ui-monospace, monospace`;

// ─── Safe helpers — NEVER crash on undefined ──────────────────────────────────

function safeHours(row: MatrixRow | undefined): string[] {
  return row?.hours ?? [];
}

function parseH(raw: string | undefined): number | null {
  if (!raw || raw === "-") return null;
  const n = parseInt((raw ?? "").replace(/[,.\s]/g, ""), 10);
  return isNaN(n) ? null : n;
}

function getHealth(raw: string | undefined, limit: number) {
  const h = parseH(raw);
  if (h === null || !limit) return null;
  const pct = Math.min((h / limit) * 100, 100);
  return { h, pct, overdue: h > limit, critical: !!(h <= limit && pct >= 85) };
}

// CRASH-FIX: countOverdue — guards every level
function countOverdue(rows: MatrixRow[] | undefined): number {
  if (!Array.isArray(rows)) return 0;
  return rows.reduce((acc, r) => {
    if (!r || !Array.isArray(r.hours)) return acc;
    return acc + r.hours.filter((h) => getHealth(h, r.limit)?.overdue).length;
  }, 0);
}

function fmtBytes(b: number) {
  return b < 1_048_576 ? `${(b / 1024).toFixed(1)} KB` : `${(b / 1_048_576).toFixed(1)} MB`;
}

// ─── Micro-components ─────────────────────────────────────────────────────────

function Dot({ color, pulse = false }: { color: string; pulse?: boolean }) {
  return (
    <span style={{ position: "relative", display: "inline-flex", width: 8, height: 8 }}>
      <span style={{ width: 8, height: 8, borderRadius: "50%", background: color, display: "block" }} />
      {pulse && <span className="animate-ping" style={{ position: "absolute", inset: 0, borderRadius: "50%", background: color, opacity: 0.4 }} />}
    </span>
  );
}

function Lbl({ children }: { children: React.ReactNode }) {
  return <span style={{ fontSize: 10, fontWeight: 700, letterSpacing: "0.18em", textTransform: "uppercase", color: T.textMuted }}>{children}</span>;
}

type ChipV = "default" | "gold" | "rose" | "amber" | "emerald";
const CHIP: Record<ChipV, React.CSSProperties> = {
  default: { background: "rgba(255,255,255,0.03)", border: `0.5px solid ${T.border}`,            color: T.textMuted },
  gold:    { background: T.goldFaint,               border: `0.5px solid ${T.goldDim}`,           color: T.gold      },
  rose:    { background: "rgba(240,64,96,0.07)",    border: "0.5px solid rgba(240,64,96,0.25)",   color: T.rose      },
  amber:   { background: "rgba(245,166,35,0.07)",   border: "0.5px solid rgba(245,166,35,0.25)",  color: T.amber     },
  emerald: { background: "rgba(52,211,153,0.07)",   border: "0.5px solid rgba(52,211,153,0.25)",  color: T.emerald   },
};

function Chip({ children, v = "default" }: { children: React.ReactNode; v?: ChipV }) {
  return (
    <span style={{ display: "inline-flex", alignItems: "center", gap: 5, padding: "4px 10px", borderRadius: 6, fontSize: 10, fontWeight: 700, letterSpacing: "0.14em", textTransform: "uppercase", ...CHIP[v] }}>
      {children}
    </span>
  );
}

// ─── Integrity ring ───────────────────────────────────────────────────────────

function IntegrityRing({ value }: { value: number }) {
  const R = 30, circ = 2 * Math.PI * R;
  const dash = ((value ?? 0) / 100) * circ;
  const clr = value === 100 ? T.emerald : value > 70 ? T.amber : T.rose;
  return (
    <div style={{ position: "relative", width: 80, height: 80, flexShrink: 0 }}>
      <svg width="80" height="80" style={{ transform: "rotate(-90deg)" }}>
        <circle cx="40" cy="40" r={R} fill="none" stroke={T.textDim} strokeWidth="2.5" />
        <circle cx="40" cy="40" r={R} fill="none" stroke={clr} strokeWidth="2.5"
          strokeDasharray={`${dash} ${circ}`} strokeLinecap="round"
          style={{ transition: "stroke-dasharray 1.2s ease" }} />
      </svg>
      <div style={{ position: "absolute", inset: 0, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center" }}>
        <span style={{ color: clr, fontSize: 18, fontWeight: 900, fontFamily: MONO, letterSpacing: "-0.04em", lineHeight: 1 }}>{value ?? 0}</span>
        <span style={{ color: T.textDim, fontSize: 8, fontWeight: 700, letterSpacing: "0.18em", textTransform: "uppercase", marginTop: 2 }}>INT%</span>
      </div>
    </div>
  );
}

// ─── Tab button ───────────────────────────────────────────────────────────────

function TabBtn({ active, onClick, icon: Icon, label, badge }: {
  active: boolean; onClick: () => void; icon: React.ElementType; label: string; badge?: number;
}) {
  return (
    <button onClick={onClick} style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 20px", borderRadius: 10, fontSize: 11, fontWeight: 700, letterSpacing: "0.16em", textTransform: "uppercase", color: active ? T.gold : T.textMuted, background: active ? T.goldFaint : "transparent", border: active ? `0.5px solid ${T.goldDim}` : "0.5px solid transparent", cursor: "pointer", transition: "all 0.15s" }}>
      <Icon size={13} />
      {label}
      {(badge ?? 0) > 0 && (
        <span style={{ display: "flex", alignItems: "center", justifyContent: "center", width: 16, height: 16, borderRadius: "50%", background: "rgba(240,64,96,0.15)", color: T.rose, fontSize: 9, fontWeight: 900 }}>{badge}</span>
      )}
    </button>
  );
}

// ─── Health cell ──────────────────────────────────────────────────────────────

function HealthCell({ raw, limit }: { raw?: string; limit: number }) {
  const m = getHealth(raw, limit);
  if (!m) {
    return (
      <td style={{ padding: "18px 16px", textAlign: "center", borderLeft: `0.5px solid ${T.border}` }}>
        <span style={{ color: T.textDim, fontFamily: MONO, fontSize: 13 }}>—</span>
      </td>
    );
  }
  const { h, pct, overdue, critical } = m;
  const numColor = overdue ? T.rose : critical ? T.amber : T.emerald;
  const barColor = overdue ? T.rose : critical ? T.amber : T.cyanDim;
  return (
    <td style={{ padding: "18px 16px", textAlign: "center", borderLeft: `0.5px solid ${T.border}`, background: overdue ? "rgba(240,64,96,0.03)" : undefined }}>
      <div style={{ fontSize: 15, fontWeight: 700, fontFamily: MONO, letterSpacing: "-0.02em", color: numColor, display: "flex", alignItems: "center", justifyContent: "center", gap: 4 }}>
        {h.toLocaleString()}
        {overdue && <span style={{ fontSize: 8, fontWeight: 900, color: T.rose, letterSpacing: "0.14em", textTransform: "uppercase", animation: "pulse 1.5s infinite" }}>OVR</span>}
      </div>
      <div style={{ marginTop: 8, width: "75%", height: 2, background: T.textDim, borderRadius: 99, overflow: "hidden", margin: "8px auto 0" }}>
        <div style={{ height: "100%", width: `${pct}%`, background: barColor, borderRadius: 99, transition: "width 0.8s ease" }} />
      </div>
      <div style={{ marginTop: 4, fontSize: 9, fontFamily: MONO, color: T.textDim }}>{pct.toFixed(0)}%</div>
    </td>
  );
}

// ─── Table row — CRASH-PROOF ──────────────────────────────────────────────────

function TableRow({ row }: { row: MatrixRow | undefined }) {
  // CRASH-FIX: every access guarded
  const comp      = row?.component ?? "";
  const limit     = row?.limit     ?? 0;
  const hours     = safeHours(row);
  const overdueCt = hours.filter((h) => getHealth(h, limit)?.overdue).length;

  return (
    <tr style={{ borderBottom: `0.5px solid ${T.border}` }}>
      <td style={{ paddingLeft: 24, paddingRight: 16, paddingTop: 18, paddingBottom: 18 }}>
        <div style={{ display: "flex", alignItems: "flex-start", gap: 10 }}>
          <ChevronRight size={12} style={{ color: T.textDim, marginTop: 2, flexShrink: 0 }} />
          <div>
            <div style={{ fontSize: 13, fontWeight: 600, color: T.text, letterSpacing: "-0.01em" }}>{comp}</div>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 5 }}>
              <Gauge size={9} style={{ color: T.goldDim }} />
              <span style={{ fontSize: 10, fontFamily: MONO, color: T.textMuted, letterSpacing: "0.08em" }}>
                LIMIT {limit.toLocaleString()} H
              </span>
              {overdueCt > 0 && <Chip v="rose">{overdueCt} overdue</Chip>}
            </div>
          </div>
        </div>
      </td>
      {/* CRASH-FIX: map over safeHours, not row.hours */}
      {hours.map((val, j) => <HealthCell key={j} raw={val} limit={limit} />)}
    </tr>
  );
}

// ─── Diagnostics panel ────────────────────────────────────────────────────────

function DiagnosticsPanel({ alerts }: { alerts: DiagAlert[] | undefined }) {
  const [open, setOpen] = useState(true);
  const safe = Array.isArray(alerts) ? alerts : [];
  const criticals = safe.filter((a) => a.severity === "critical");
  const warnings  = safe.filter((a) => a.severity === "warning");
  if (safe.length === 0) return null;

  return (
    <div style={{ borderRadius: 20, overflow: "hidden", border: `0.5px solid ${T.border}`, background: T.surface }}>
      {/* Header */}
      <button
        onClick={() => setOpen((o) => !o)}
        style={{ width: "100%", display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px 24px", background: T.surfaceHigh, border: "none", cursor: "pointer" }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <Bell size={14} style={{ color: T.gold }} />
          <span style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.18em", textTransform: "uppercase", color: T.text }}>
            Maintenance Diagnostics
          </span>
          {criticals.length > 0 && (
            <Chip v="rose">{criticals.length} critical</Chip>
          )}
          {warnings.length > 0 && (
            <Chip v="amber">{warnings.length} warning</Chip>
          )}
        </div>
        {open ? <ChevronUp size={14} style={{ color: T.textMuted }} /> : <ChevronDown size={14} style={{ color: T.textMuted }} />}
      </button>

      {open && (
        <div style={{ padding: "12px 0" }}>
          {safe.map((a, i) => {
            const isCrit = a.severity === "critical";
            return (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: 14, padding: "10px 24px", borderBottom: i < safe.length - 1 ? `0.5px solid ${T.border}` : undefined }}>
                <div style={{ width: 4, height: 32, borderRadius: 4, background: isCrit ? T.rose : T.amber, flexShrink: 0 }} />
                <div style={{ minWidth: 80 }}>
                  <Chip v={isCrit ? "rose" : "amber"}>{isCrit ? "Critical" : "Warning"}</Chip>
                </div>
                <div style={{ flex: 1 }}>
                  <span style={{ fontSize: 12, fontWeight: 600, color: T.text }}>{a.section} — {a.component}</span>
                  <span style={{ fontSize: 11, color: T.textMuted, marginLeft: 8 }}>
                    {a.section.startsWith("Gen") ? `Cyl ${a.slot}` : `#${a.slot}`}
                  </span>
                </div>
                <div style={{ textAlign: "right" }}>
                  <div style={{ fontSize: 11, color: isCrit ? T.rose : T.amber, fontFamily: MONO, fontWeight: 700 }}>{a.message}</div>
                  <div style={{ fontSize: 10, color: T.textDim, fontFamily: MONO, marginTop: 2 }}>
                    {a.hours.toLocaleString()} / {a.limit.toLocaleString()} h
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ─── Main component ───────────────────────────────────────────────────────────

export default function FleetCommandNexus() {
  const [data,        setData       ] = useState<VesselData | null>(null);
  const [loading,     setLoading    ] = useState(true);
  const [error,       setError      ] = useState<string | null>(null);
  const [activeTab,   setActiveTab  ] = useState<ActiveTab>("ME");
  const [activeDg,    setActiveDg   ] = useState(0);
  const [isDragging,  setIsDragging ] = useState(false);
  const [uploadState, setUploadState] = useState<"idle" | "busy" | "ok">("idle");
  const fileRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    fetch("/api/init")
      .then((r) => r.json())
      .then((j: VesselData) => { if (j?.status === "success") setData(j); })
      .catch(() => {})
      .finally(() => setLoading(false));
  }, []);

  const handleUpload = useCallback(async (file: File) => {
    setLoading(true);
    setError(null);
    setUploadState("busy");
    const fd = new FormData();
    fd.append("file", file);
    try {
      const res    = await fetch("/api/upload", { method: "POST", body: fd });
      const result = await res.json() as VesselData & { detail?: string };
      if (result?.status === "success") {
        setData(result);
        setActiveTab("ME");
        setActiveDg(0);
        setUploadState("ok");
        setTimeout(() => setUploadState("idle"), 4000);
      } else {
        setError(result?.detail ?? "The document could not be parsed. Verify the file format.");
        setUploadState("idle");
      }
    } catch {
      setError("Cannot reach the Python backend on port 8000. Run: cd services/brain && python3 main.py");
      setUploadState("idle");
    } finally {
      setLoading(false);
    }
  }, []);

  const onDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const f = e.dataTransfer.files?.[0];
    if (f) handleUpload(f);
  }, [handleUpload]);

  // CRASH-FIX: all array accesses safe
  const meOverdue  = useMemo(() => countOverdue(data?.matrices?.main_engine), [data]);
  const dgOverdue  = useMemo(() => countOverdue(data?.matrices?.dg_general),  [data]);
  const savedLabel = data?._saved_at
    ? new Date(data._saved_at).toLocaleString(undefined, { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" })
    : null;

  // Loading splash
  if (loading && !data) {
    return (
      <div style={{ minHeight: "100vh", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 16, background: T.bg }}>
        <div style={{ width: 48, height: 48, borderRadius: 14, display: "flex", alignItems: "center", justifyContent: "center", background: T.goldFaint, border: `0.5px solid ${T.goldDim}` }}>
          <Activity size={22} className="animate-spin" style={{ color: T.gold }} />
        </div>
        <Lbl>Initializing Fleet Nexus</Lbl>
      </div>
    );
  }

  return (
    <div
      style={{ minHeight: "100vh", background: T.bg, color: T.text }}
      onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
      onDragLeave={(e) => { if (!e.currentTarget.contains(e.relatedTarget as Node)) setIsDragging(false); }}
      onDrop={onDrop}
    >
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;700;900&display=swap');
        @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:.3} }
        *{box-sizing:border-box;margin:0;padding:0}
        button{font-family:inherit}
        tr:hover td{background:rgba(201,168,76,0.018)!important}
      `}</style>

      {/* Drop overlay */}
      {isDragging && (
        <div style={{ position: "fixed", inset: 0, zIndex: 50, display: "flex", alignItems: "center", justifyContent: "center", background: "rgba(6,8,14,0.94)", border: `2px dashed ${T.goldDim}` }}>
          <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 20 }}>
            <div style={{ width: 80, height: 80, borderRadius: 24, display: "flex", alignItems: "center", justifyContent: "center", background: T.goldFaint, border: `1px solid ${T.goldDim}` }}>
              <UploadCloud size={36} className="animate-bounce" style={{ color: T.gold }} />
            </div>
            <div style={{ textAlign: "center" }}>
              <p style={{ fontSize: 20, fontWeight: 900, color: T.text }}>Release to upload</p>
              <p style={{ color: T.textMuted, fontSize: 13, marginTop: 6, fontFamily: MONO }}>.doc · .docx engine report</p>
            </div>
          </div>
        </div>
      )}

      {/* Header */}
      <header style={{ position: "sticky", top: 0, zIndex: 30, height: 58, background: `${T.bg}f2`, backdropFilter: "blur(20px)", borderBottom: `0.5px solid ${T.border}`, display: "flex", alignItems: "center" }}>
        <div style={{ maxWidth: 1600, width: "100%", margin: "0 auto", padding: "0 32px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <div style={{ width: 30, height: 30, borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center", background: T.goldFaint, border: `0.5px solid ${T.goldDim}` }}>
              <Ship size={14} style={{ color: T.gold }} />
            </div>
            <span style={{ fontSize: 14, fontWeight: 900, color: T.text, letterSpacing: "-0.02em" }}>
              Fleet Command <span style={{ color: T.gold }}>Nexus</span>
            </span>
          </div>

          <div style={{ display: "flex", alignItems: "center", gap: 20 }}>
            {data && <div style={{ display: "flex", alignItems: "center", gap: 6 }}><Dot color={T.emerald} pulse /><span style={{ color: T.textMuted, fontSize: 10, fontFamily: MONO, fontWeight: 700, letterSpacing: "0.18em", textTransform: "uppercase" }}>Live</span></div>}
            {uploadState === "ok" && (
              <div style={{ display: "flex", alignItems: "center", gap: 6, color: T.emerald, fontSize: 11, fontWeight: 700, letterSpacing: "0.14em", textTransform: "uppercase" }}>
                <CheckCircle2 size={13} />Parse OK
              </div>
            )}
            <div style={{ position: "relative" }}>
              <input ref={fileRef} type="file" accept=".doc,.docx" style={{ position: "absolute", inset: 0, opacity: 0, cursor: "pointer", zIndex: 10, width: "100%", height: "100%" }} onChange={(e) => e.target.files?.[0] && handleUpload(e.target.files[0])} />
              <button style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 18px", borderRadius: 10, fontSize: 11, fontWeight: 700, letterSpacing: "0.16em", textTransform: "uppercase", background: T.goldFaint, border: `0.5px solid ${T.goldDim}`, color: T.gold, cursor: "pointer" }}>
                {loading && uploadState === "busy" ? <RefreshCw size={13} className="animate-spin" /> : <UploadCloud size={13} />}
                Upload Report
              </button>
            </div>
          </div>
        </div>
      </header>

      {/* Error banner */}
      {error && (
        <div style={{ maxWidth: 1600, margin: "0 auto", padding: "20px 32px 0" }}>
          <div style={{ display: "flex", alignItems: "flex-start", gap: 12, padding: "14px 20px", borderRadius: 12, background: "rgba(240,64,96,0.06)", border: "0.5px solid rgba(240,64,96,0.2)" }}>
            <XCircle size={16} style={{ color: T.rose, marginTop: 1, flexShrink: 0 }} />
            <div style={{ flex: 1 }}>
              <p style={{ fontSize: 10, fontWeight: 700, letterSpacing: "0.18em", textTransform: "uppercase", color: T.rose }}>Error</p>
              <p style={{ fontSize: 13, color: "rgba(240,64,96,0.7)", marginTop: 3 }}>{error}</p>
            </div>
            <button onClick={() => setError(null)} style={{ fontSize: 10, fontWeight: 700, letterSpacing: "0.14em", textTransform: "uppercase", color: T.textMuted, background: "none", border: "none", cursor: "pointer" }}>Dismiss</button>
          </div>
        </div>
      )}

      {/* No-data splash */}
      {!data && !loading && (
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", minHeight: "80vh", gap: 32, padding: "0 32px" }}>
          <div style={{ textAlign: "center", maxWidth: 380 }}>
            <div style={{ width: 56, height: 56, borderRadius: 16, display: "flex", alignItems: "center", justifyContent: "center", background: T.surface, border: `0.5px solid ${T.border}`, margin: "0 auto 24px" }}>
              <Ship size={24} style={{ color: T.textDim }} />
            </div>
            <p style={{ fontSize: 20, fontWeight: 900, color: T.text, letterSpacing: "-0.02em", marginBottom: 10 }}>No vessel data loaded</p>
            <p style={{ fontSize: 14, color: T.textMuted, lineHeight: 1.7 }}>
              Drop a <code style={{ color: T.gold, fontFamily: MONO }}>.doc</code> engine running-hours report anywhere, or click <strong style={{ color: T.gold }}>Upload Report</strong>.
            </p>
          </div>
          <div
            onClick={() => fileRef.current?.click()}
            style={{ width: 280, height: 150, borderRadius: 20, cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 12, border: `1px dashed ${T.border}`, transition: "border-color 0.2s" }}
            onMouseEnter={(e) => (e.currentTarget.style.borderColor = T.goldDim)}
            onMouseLeave={(e) => (e.currentTarget.style.borderColor = T.border)}
          >
            <UploadCloud size={22} style={{ color: T.textDim }} />
            <Lbl>Drop .doc / .docx here</Lbl>
          </div>
        </div>
      )}

      {/* Dashboard */}
      {data && (
        <div style={{ maxWidth: 1600, margin: "0 auto", padding: "28px 32px", display: "flex", flexDirection: "column", gap: 20 }}>

          {/* Vessel header */}
          <div style={{ borderRadius: 20, padding: "24px 28px", background: T.surface, border: `0.5px solid ${T.border}`, display: "flex", flexWrap: "wrap", alignItems: "center", justifyContent: "space-between", gap: 24 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 20 }}>
              <div style={{ position: "relative", flexShrink: 0 }}>
                <div style={{ width: 52, height: 52, borderRadius: 16, display: "flex", alignItems: "center", justifyContent: "center", background: T.goldFaint, border: `0.5px solid ${T.goldDim}` }}>
                  <Ship size={22} style={{ color: T.gold }} />
                </div>
                <span style={{ position: "absolute", bottom: -3, right: -3, width: 12, height: 12, borderRadius: "50%", background: T.emerald, border: `2px solid ${T.surface}` }} />
              </div>
              <div>
                <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                  <h1 style={{ fontSize: 22, fontWeight: 900, color: T.text, letterSpacing: "-0.03em" }}>{data.vessel}</h1>
                  {data.is_fallback && <Chip v="amber">Fallback match</Chip>}
                  {(meOverdue + dgOverdue) > 0
                    ? <Chip v="rose"><AlertTriangle size={9} />{meOverdue + dgOverdue} overdue</Chip>
                    : <Chip v="emerald"><ShieldCheck size={9} />All clear</Chip>}
                </div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginTop: 12 }}>
                  <Chip><Cpu size={10} />{data.config?.me_cyls ?? 0} M/E cyls</Chip>
                  <Chip><Database size={10} />{data.config?.dg_count ?? 0} generators</Chip>
                  <Chip><Layers size={10} />{data.config?.aux_cyls ?? 0} aux cyls/gen</Chip>
                  {savedLabel && <Chip><Clock size={10} />Synced {savedLabel}</Chip>}
                </div>
              </div>
            </div>

            <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
              <IntegrityRing value={data.integrity ?? 0} />
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                {[
                  { icon: BarChart3, label: "M/E overdue",  val: meOverdue },
                  { icon: Zap,       label: "D/G overdue",  val: dgOverdue },
                ].map(({ icon: I, label, val }) => (
                  <div key={label} style={{ padding: "12px 16px", borderRadius: 12, background: T.surfaceHigh, border: `0.5px solid ${T.border}`, display: "flex", flexDirection: "column", gap: 6 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 6 }}><I size={11} style={{ color: T.goldDim }} /><Lbl>{label}</Lbl></div>
                    <span style={{ fontSize: 22, fontWeight: 900, fontFamily: MONO, letterSpacing: "-0.04em", color: val > 0 ? T.rose : T.emerald }}>{val}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>

          {/* Diagnostics */}
          <DiagnosticsPanel alerts={data.diagnostics} />

          {/* Tab bar */}
          <div style={{ display: "flex", gap: 4, padding: 4, background: T.surface, border: `0.5px solid ${T.border}`, borderRadius: 14, width: "fit-content" }}>
            <TabBtn active={activeTab === "ME"}     onClick={() => setActiveTab("ME")}     icon={Cpu}      label="Main Engine"   badge={meOverdue} />
            <TabBtn active={activeTab === "DG_GEN"} onClick={() => setActiveTab("DG_GEN")} icon={Database} label="Generators"    badge={dgOverdue} />
            <TabBtn active={activeTab === "DG_CYL"} onClick={() => setActiveTab("DG_CYL")} icon={Server}   label="Aux Cylinders" />
          </div>

          {/* Matrix table */}
          <div style={{ borderRadius: 20, overflow: "hidden", background: T.surface, border: `0.5px solid ${T.border}` }}>
            {activeTab === "DG_CYL" && (
              <div style={{ display: "flex", gap: 8, padding: "12px 20px", borderBottom: `0.5px solid ${T.border}`, background: T.surfaceHigh }}>
                {Array.from({ length: data.config?.dg_count ?? 0 }).map((_, i) => (
                  <button key={i} onClick={() => setActiveDg(i)} style={{ display: "flex", alignItems: "center", gap: 6, padding: "6px 16px", borderRadius: 8, fontSize: 10, fontWeight: 700, letterSpacing: "0.16em", textTransform: "uppercase", cursor: "pointer", color: activeDg === i ? T.gold : T.textMuted, background: activeDg === i ? T.goldFaint : "transparent", border: activeDg === i ? `0.5px solid ${T.goldDim}` : "0.5px solid transparent", transition: "all 0.15s" }}>
                    <Zap size={10} />Gen {i + 1}
                  </button>
                ))}
              </div>
            )}

            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <colgroup>
                  <col style={{ width: 260 }} />
                  {activeTab === "ME"     && Array.from({ length: data.config?.me_cyls  ?? 0 }).map((_, i) => <col key={i} />)}
                  {activeTab === "DG_GEN" && Array.from({ length: data.config?.dg_count ?? 0 }).map((_, i) => <col key={i} />)}
                  {activeTab === "DG_CYL" && Array.from({ length: data.config?.aux_cyls ?? 0 }).map((_, i) => <col key={i} />)}
                </colgroup>
                <thead>
                  <tr style={{ borderBottom: `0.5px solid ${T.border}`, background: T.surfaceHigh }}>
                    <th style={{ paddingLeft: 24, paddingRight: 16, paddingTop: 14, paddingBottom: 14, textAlign: "left" }}><Lbl>Component</Lbl></th>
                    {activeTab === "ME"     && Array.from({ length: data.config?.me_cyls  ?? 0 }).map((_, i) => <th key={i} style={{ padding: "14px 16px", textAlign: "center", borderLeft: `0.5px solid ${T.border}` }}><Lbl>Cyl {i + 1}</Lbl></th>)}
                    {activeTab === "DG_GEN" && Array.from({ length: data.config?.dg_count ?? 0 }).map((_, i) => <th key={i} style={{ padding: "14px 16px", textAlign: "center", borderLeft: `0.5px solid ${T.border}` }}><Lbl>D/G {i + 1}</Lbl></th>)}
                    {activeTab === "DG_CYL" && Array.from({ length: data.config?.aux_cyls ?? 0 }).map((_, i) => <th key={i} style={{ padding: "14px 16px", textAlign: "center", borderLeft: `0.5px solid ${T.border}` }}><Lbl>Cyl {i + 1}</Lbl></th>)}
                  </tr>
                </thead>
                <tbody>
                  {/* CRASH-FIX: every matrix array access uses ?. and ?? [] */}
                  {activeTab === "ME" && (data.matrices?.main_engine ?? []).map((row, i) => (
                    <TableRow key={`me-${i}`} row={row} />
                  ))}
                  {activeTab === "DG_GEN" && (data.matrices?.dg_general ?? []).map((row, i) => (
                    <TableRow key={`dg-${i}`} row={row} />
                  ))}
                  {activeTab === "DG_CYL" && (data.matrices?.aux_engine ?? []).map((row, i) => {
                    // CRASH-FIX: safe access to dg_hours per generator
                    const chunk: string[] =
                      Array.isArray(row?.dg_hours)
                        ? (row.dg_hours[activeDg] ?? Array.from({ length: data.config?.aux_cyls ?? 0 }, () => "-"))
                        : Array.from({ length: data.config?.aux_cyls ?? 0 }, () => "-");
                    const synRow: MatrixRow = { component: row?.component ?? "", limit: row?.limit ?? 0, hours: chunk };
                    return <TableRow key={`aux-${i}`} row={synRow} />;
                  })}
                </tbody>
              </table>
            </div>

            {data._upload && (
              <div style={{ padding: "10px 24px", borderTop: `0.5px solid ${T.border}`, background: T.surfaceHigh, display: "flex", alignItems: "center", gap: 12 }}>
                <TrendingUp size={11} style={{ color: T.textDim }} />
                {[data._upload.original_filename, fmtBytes(data._upload.bytes), new Date(data._upload.uploaded_at).toLocaleString()].map((s, i) => (
                  <React.Fragment key={i}>
                    {i > 0 && <span style={{ color: T.textDim, fontSize: 11 }}>·</span>}
                    <span style={{ color: T.textDim, fontSize: 11, fontFamily: MONO }}>{s}</span>
                  </React.Fragment>
                ))}
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}