"use client"

import { useState, useMemo, useEffect, useCallback, useRef } from "react"

// ── LOCALSTORAGE HELPERS ───────────────────────────────────────
const LS_PREFIX = "poin_cabang_"

function lsKey(month: number, year: number) {
  return `${LS_PREFIX}${year}_${String(month + 1).padStart(2, "0")}`
}

function loadFromStorage(month: number, year: number): DateData | null {
  try {
    const raw = localStorage.getItem(lsKey(month, year))
    return raw ? JSON.parse(raw) : null
  } catch { return null }
}

function saveToStorage(month: number, year: number, data: DateData) {
  try {
    localStorage.setItem(lsKey(month, year), JSON.stringify(data))
  } catch { /* kuota penuh */ }
}

function listSavedPeriods(): { key: string; label: string }[] {
  try {
    return Object.keys(localStorage)
      .filter((k) => k.startsWith(LS_PREFIX))
      .map((k) => {
        const [yr, mo] = k.replace(LS_PREFIX, "").split("_")
        return { key: k, label: `${MONTHS[parseInt(mo) - 1]} ${yr}` }
      })
      .sort((a, b) => a.key.localeCompare(b.key))
  } catch { return [] }
}

// ── DATA KATEGORI ──────────────────────────────────────────────
const categories = [
  { id:  1, no:  "1", label: "MTB > 25 New CIF - EDC",                   sheet: "1.MTB 25- EDC",                 poin:  8 },
  { id:  2, no:  "2", label: "GIRO > 25 New CIF",                         sheet: "2.GIRO 50",                     poin:  8 },
  { id:  3, no:  "3", label: "KOPRA / TABREG > 10 / TRM",                 sheet: "3.KOPRA TABREG>10- TRM",        poin:  4 },
  { id:  4, no:  "4", label: "AXA CC Retail",                             sheet: "4.AXA RETAIL- CC APPROVE",      poin:  6 },
  { id:  5, no:  "5", label: "HVC > 100 jt",                              sheet: "5.AXA HVC",                     poin: 10 },
  { id:  6, no:  "6", label: "KSM < 100 / LVM Usaha",                     sheet: "6.KSM LVM USAHA",               poin:  4 },
  { id:  7, no:  "7", label: "New CIF < 25 jt (Gir-Tabis-Tabreg) / GMM", sheet: "7.CC-Incomming- LVM USAHA ORG", poin:  2 },
  { id:  8, no:  "8", label: "KPR DTBO / KSM > 100 jt / CC Approve",     sheet: "8.KPR DTBO KSM CC",             poin:  8 },
  { id:  9, no:  "9", label: "KKB / PKS Mitra ID",                        sheet: "9. KKB OR PKS MITRAID",         poin:  4 },
  { id: 10, no: "10", label: "E-Commerce / NTP",                          sheet: "10.ECOMMERCE-NTP",              poin:  8 },
  { id: 11, no: "11", label: "Livin USAK / Payroll PMP",                  sheet: "11.LIVIN USAK OR PAYROLL PMP",  poin:  1 },
]

const MAX_POIN  = categories.reduce((s, c) => s + c.poin, 0) // 63
const DATES     = Array.from({ length: 31 }, (_, i) => i + 1)
const MONTHS    = ["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"]
const employees = [
  { id:  1, name: "10900 - KC BATAM IMAM BONJOL" },
  { id:  2, name: "10901 - KCP BATAM LUBUK BAJA" },
  { id:  3, name: "10902 - KCP BATAM RAJA ALI HAJI" },
  { id:  4, name: "10903 - KCP BATAM SEKUPANG" },
  { id:  5, name: "10904 - KCP BATAM INDUSTRIAL PARK" },
  { id:  6, name: "10905 - KC TANJUNGPINANG" },
  { id:  7, name: "10906 - KCP TANJUNG UBAN" },
  { id:  8, name: "10907 - KCP BATAM BANDARA HANG NADIM" },
  { id:  9, name: "10908 - KCP BATAM CENTER" },
  { id: 10, name: "10909 - KCP BATAM SP PLAZA" },
  { id: 11, name: "10910 - KCP BATAM KAWASAN INDUSTRI TUNAS" },
  { id: 12, name: "10911 - KCP BATAM TIBAN" },
  { id: 13, name: "10912 - KCP BATAM PANBIL" },
  { id: 14, name: "10913 - KCP TANJUNG BALAI KARIMUN" },
  { id: 15, name: "10914 - KCP KIJANG" },
  { id: 16, name: "10915 - KCP NATUNA" },
  { id: 17, name: "10916 - KCP BATAM KAWASAN INDUSTRI KABIL" },
  { id: 18, name: "10917 - KCP BINTAN CENTER" },
  { id: 19, name: "10918 - KCP BATAM FANINDO" },
  { id: 20, name: "10919 - KCP BATAM KEPRI MALL" },
  { id: 21, name: "10920 - KCP BATAM PALM SPRING" },
  { id: 22, name: "10922 - KCP BATAM BOTANIA" },
  { id: 23, name: "10924 - KCP BATAM GRAND NIAGA MAS" },
  { id: 24, name: "10925 - KCP BATAM BATU AMPAR" },
  { id: 25, name: "10926 - KCP BINTAN ALUMINA INDONESIA" },
  { id: 26, name: "10977 - KCP TANJUNG BATU" },
  { id: 27, name: "10980 - KCP BATAM TANJUNG PIAYU" },
]

// ── TYPES ──────────────────────────────────────────────────────
// dateData[empId][date][catId] = string nilai
type DateData = Record<number, Record<number, Record<number, string>>>
type TabType  = "tanggal" | "simulasi" | "formula"
type FormulaVariant = "boolean" | "iferror" | "sign" | "choose"

const formulaVariants: { key: FormulaVariant; label: string; desc: string; badge: string }[] = [
  { key: "boolean", label: "Boolean Arithmetic", desc: "Ekspresi logika ringkas: positif dikali poin, nihil menghasilkan -1.",          badge: "Direkomendasikan" },
  { key: "iferror", label: "IF + IFERROR",        desc: "Aman dari error jika sheet tidak ditemukan. Positif dikali poin, nihil -1.",   badge: "Paling Aman" },
  { key: "sign",    label: "IF + SIGN",            desc: "Deteksi positif via SIGN(). Positif dikali poin, nihil -1.",                   badge: "Mudah Dibaca" },
  { key: "choose",  label: "IF Tunggal",            desc: "Pola IF klasik: positif dikali poin, nihil -1. Hasil angka murni.",            badge: "Klasik" },
]

// ── HELPERS ────────────────────────────────────────────────────
function buildFormula(variant: FormulaVariant, cell: string): string {
  const perCat: Record<FormulaVariant, (sheet: string, poin: number) => string> = {
    boolean: (s, p) => `(('${s}'!${cell}>0)*'${s}'!${cell}*${p})+(('${s}'!${cell}<1)*-1)`,
    iferror: (s, p) => `IFERROR(IF('${s}'!${cell}>0,'${s}'!${cell}*${p},-1),0)`,
    sign:    (s, p) => `IF(SIGN('${s}'!${cell})>0,'${s}'!${cell}*${p},-1)`,
    choose:  (s, p) => `IF('${s}'!${cell}>0,'${s}'!${cell}*${p},-1)`,
  }
  return `=${categories.map((c) => perCat[variant](c.sheet, c.poin)).join("\n+")}`
}

function calcScore(raw: string, poin: number) {
  const num = parseFloat(raw)
  if (raw === "" || isNaN(num)) return { score: 0, status: "empty" as const }
  if (num > 0)                  return { score: num * poin, status: "positif" as const }
  return                               { score: -1, status: "nihil" as const }
}

function calcDateScore(empId: number, date: number, dateData: DateData) {
  const vals = dateData[empId]?.[date] ?? {}
  let total = 0, pos = 0, neg = 0
  for (const cat of categories) {
    const { score, status } = calcScore(vals[cat.id] ?? "", cat.poin)
    total += score
    if (status === "positif") pos += score
    if (status === "nihil")   neg += score
  }
  return { total, pos, neg }
}

function calcEmployeeMonthTotal(empId: number, dateData: DateData) {
  let total = 0
  for (const d of DATES) {
    total += calcDateScore(empId, d, dateData).total
  }
  return total
}

function ScorePill({ v, size = "sm" }: { v: number; size?: "sm" | "xs" }) {
  const base = size === "sm"
    ? "tabular-nums text-xs font-bold rounded-md px-2 py-0.5"
    : "tabular-nums text-[10px] font-bold rounded px-1.5 py-0.5"
  if (v > 0)  return <span className={`${base} text-emerald-700 bg-emerald-500/15`}>+{v.toFixed(0)}</span>
  if (v < 0)  return <span className={`${base} text-red-600   bg-red-500/15`}>{v.toFixed(0)}</span>
  return              <span className={`${base} text-muted-foreground bg-secondary`}>0</span>
}

function CopyButton({ text }: { text: string }) {
  const [copied, setCopied] = useState(false)
  return (
    <button
      onClick={async () => {
        try { await navigator.clipboard.writeText(text); setCopied(true); setTimeout(() => setCopied(false), 2000) } catch { /* */ }
      }}
      className="inline-flex items-center gap-1.5 rounded-md border border-border bg-background px-3 py-1.5 text-xs font-medium text-muted-foreground hover:text-foreground hover:bg-secondary transition-colors"
    >
      {copied
        ? <><svg className="w-3.5 h-3.5 text-emerald-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}><path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7"/></svg>Tersalin</>
        : <><svg className="w-3.5 h-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1"/></svg>Salin</>
      }
    </button>
  )
}

// ── INIT STATE ─────────────────────────────────────────────────
function initDateData(): DateData {
  const data: DateData = {}
  for (const emp of employees) {
    data[emp.id] = {}
    for (const d of DATES) {
      data[emp.id][d] = Object.fromEntries(categories.map((c) => [c.id, ""]))
    }
  }
  return data
}

// ══════════════ KOMPONEN UTAMA ══════════════════════════════════
export default function Page() {
  const today = new Date()
  const [selectedEmp,    setSelectedEmp]    = useState(1)
  const [activeTab,      setActiveTab]      = useState<TabType>("tanggal")
  const [activeVariant,  setActiveVariant]  = useState<FormulaVariant>("boolean")
  const [cellRef,        setCellRef]        = useState("F6")
  const [selectedMonth,  setSelectedMonth]  = useState(today.getMonth())
  const [selectedYear,   setSelectedYear]   = useState(today.getFullYear())
  const [selectedDate,   setSelectedDate]   = useState<number>(today.getDate())
  const [dateData,       setDateData]       = useState<DateData>(initDateData)
  const [savedAt,        setSavedAt]        = useState<string | null>(null)
  const [showImport,     setShowImport]     = useState(false)
  const [importText,     setImportText]     = useState("")
  const [importError,    setImportError]    = useState("")
  const [savedPeriods,   setSavedPeriods]   = useState<{ key: string; label: string }[]>([])
  const saveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null)

  // Load data saat bulan/tahun berubah
  useEffect(() => {
    const stored = loadFromStorage(selectedMonth, selectedYear)
    setDateData(stored ?? initDateData())
    setSavedAt(null)
  }, [selectedMonth, selectedYear])

  // Auto-save dengan debounce 800ms setiap kali dateData berubah
  useEffect(() => {
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current)
    saveTimerRef.current = setTimeout(() => {
      saveToStorage(selectedMonth, selectedYear, dateData)
      setSavedAt(new Date().toLocaleTimeString("id-ID", { hour: "2-digit", minute: "2-digit", second: "2-digit" }))
      setSavedPeriods(listSavedPeriods())
    }, 800)
    return () => { if (saveTimerRef.current) clearTimeout(saveTimerRef.current) }
  }, [dateData, selectedMonth, selectedYear])

  // Sync daftar periode tersimpan saat pertama load
  useEffect(() => { setSavedPeriods(listSavedPeriods()) }, [])

  // Ekspor JSON bulan aktif
  const handleExport = useCallback(() => {
    const blob = new Blob([JSON.stringify({ month: selectedMonth, year: selectedYear, data: dateData }, null, 2)], { type: "application/json" })
    const url  = URL.createObjectURL(blob)
    const a    = document.createElement("a")
    a.href     = url
    a.download = `poin_cabang_${selectedYear}_${String(selectedMonth + 1).padStart(2, "0")}.json`
    a.click()
    URL.revokeObjectURL(url)
  }, [dateData, selectedMonth, selectedYear])

  // Impor JSON
  const handleImport = useCallback(() => {
    try {
      const parsed = JSON.parse(importText)
      if (!parsed.data || typeof parsed.month !== "number" || typeof parsed.year !== "number") throw new Error("Format tidak valid")
      setSelectedMonth(parsed.month)
      setSelectedYear(parsed.year)
      setDateData(parsed.data)
      saveToStorage(parsed.month, parsed.year, parsed.data)
      setShowImport(false)
      setImportText("")
      setImportError("")
    } catch (e: unknown) {
      setImportError(e instanceof Error ? e.message : "Format JSON tidak valid")
    }
  }, [importText])

  // Hapus periode tersimpan
  const handleDeletePeriod = useCallback((key: string) => {
    try { localStorage.removeItem(key) } catch { /* */ }
    setSavedPeriods(listSavedPeriods())
  }, [])

  const formula         = useMemo(() => buildFormula(activeVariant, cellRef), [activeVariant, cellRef])
  const selectedEmpData = employees.find((e) => e.id === selectedEmp)!

  // Skor per tanggal untuk karyawan terpilih (semua 31 hari)
  const dateScores = useMemo(() =>
    DATES.map((d) => ({ date: d, ...calcDateScore(selectedEmp, d, dateData) })),
    [selectedEmp, dateData]
  )

  // Total bulan
  const monthTotal = useMemo(() =>
    dateScores.reduce((s, d) => s + d.total, 0),
    [dateScores]
  )
  const monthPos = useMemo(() => dateScores.reduce((s, d) => s + d.pos, 0), [dateScores])
  const monthNeg = useMemo(() => dateScores.reduce((s, d) => s + d.neg, 0), [dateScores])

  // Detail kategori untuk tanggal terpilih
  const dateDetail = useMemo(() => {
    const vals = dateData[selectedEmp]?.[selectedDate] ?? {}
    return categories.map((c) => {
      const raw = vals[c.id] ?? ""
      const { score, status } = calcScore(raw, c.poin)
      return { ...c, raw, score, status }
    })
  }, [dateData, selectedEmp, selectedDate])

  const detailPos = useMemo(() => dateDetail.filter((r) => r.status === "positif").reduce((s, r) => s + r.score, 0), [dateDetail])
  const detailNeg = useMemo(() => dateDetail.filter((r) => r.status === "nihil").reduce((s, r) => s + r.score, 0), [dateDetail])
  const detailTotal = detailPos + detailNeg

  // Summary sidebar semua karyawan
  const empMonthTotals = useMemo(() =>
    Object.fromEntries(employees.map((e) => [e.id, calcEmployeeMonthTotal(e.id, dateData)])),
    [dateData]
  )

  function setVal(date: number, catId: number, val: string) {
    setDateData((prev) => ({
      ...prev,
      [selectedEmp]: {
        ...prev[selectedEmp],
        [date]: { ...prev[selectedEmp][date], [catId]: val },
      },
    }))
  }

  function resetDate(date: number) {
    setDateData((prev) => ({
      ...prev,
      [selectedEmp]: {
        ...prev[selectedEmp],
        [date]: Object.fromEntries(categories.map((c) => [c.id, ""])),
      },
    }))
  }

  const daysInMonth = new Date(selectedYear, selectedMonth + 1, 0).getDate()

  return (
    <div className="flex h-screen bg-background font-sans overflow-hidden">

      {/* ═══════ SIDEBAR KIRI ═══════ */}
      <aside className="w-60 flex-shrink-0 border-r border-border flex flex-col bg-card">
        <div className="px-4 py-4 border-b border-border">
          <p className="text-[10px] font-semibold uppercase tracking-widest text-muted-foreground mb-0.5">Wilayah Kepri</p>
          <h2 className="text-sm font-bold text-foreground">Daftar Cabang</h2>
          <div className="flex items-center gap-2 mt-2">
            <select
              value={selectedMonth}
              onChange={(e) => setSelectedMonth(+e.target.value)}
              className="flex-1 rounded-md border border-border bg-background px-2 py-1 text-xs text-foreground focus:outline-none focus:ring-2 focus:ring-primary/50"
            >
              {MONTHS.map((m, i) => <option key={i} value={i}>{m}</option>)}
            </select>
            <input
              type="number"
              value={selectedYear}
              onChange={(e) => setSelectedYear(+e.target.value)}
              className="w-16 rounded-md border border-border bg-background px-2 py-1 text-xs font-mono text-center text-foreground focus:outline-none focus:ring-2 focus:ring-primary/50"
            />
          </div>
        </div>

        <nav className="flex-1 overflow-y-auto py-1" aria-label="Daftar karyawan">
          {employees.map((emp) => {
            const total    = empMonthTotals[emp.id]
            const isActive = emp.id === selectedEmp
            const hasData  = DATES.some((d) =>
              Object.values(dateData[emp.id]?.[d] ?? {}).some((v) => v !== "")
            )
            return (
              <button
                key={emp.id}
                onClick={() => setSelectedEmp(emp.id)}
                className={`w-full flex items-center justify-between px-3 py-2 text-left transition-colors group ${
                  isActive ? "bg-primary/10 border-r-2 border-primary" : "hover:bg-accent border-r-2 border-transparent"
                }`}
              >
                <div className="flex items-center gap-2 min-w-0">
                  <div className={`w-7 h-7 rounded-md flex-shrink-0 flex flex-col items-center justify-center leading-none ${
                    isActive ? "bg-primary text-primary-foreground" : "bg-secondary text-muted-foreground group-hover:bg-primary/15"
                  }`}>
                    <span className="text-[8px] font-bold tabular-nums">{emp.name.slice(0, 5)}</span>
                  </div>
                  <div className="min-w-0">
                    <p className={`text-[11px] font-semibold leading-tight truncate ${isActive ? "text-primary" : "text-foreground"}`}>
                      {emp.name.slice(8)}
                    </p>
                  </div>
                </div>
                {hasData && <ScorePill v={total} />}
              </button>
            )
          })}
        </nav>

        {/* Footer rata-rata tim */}
        <div className="border-t border-border px-4 py-3 bg-secondary/30">
          <p className="text-[10px] text-muted-foreground uppercase tracking-wider font-semibold mb-1">
            Rata-rata Cabang &mdash; {MONTHS[selectedMonth]} {selectedYear}
          </p>
          {(() => {
            const filled = employees.filter((e) =>
              DATES.some((d) => Object.values(dateData[e.id]?.[d] ?? {}).some((v) => v !== ""))
            )
            const avg = filled.length > 0
              ? filled.reduce((s, e) => s + empMonthTotals[e.id], 0) / filled.length
              : 0
            return (
              <div className="flex items-baseline gap-1.5">
                <span className={`text-lg font-bold tabular-nums ${avg > 0 ? "text-emerald-600" : avg < 0 ? "text-destructive" : "text-foreground"}`}>
                  {avg > 0 ? "+" : ""}{avg.toFixed(1)}
                </span>
                <span className="text-xs text-muted-foreground">{filled.length} diisi</span>
              </div>
            )
          })()}
        </div>
      </aside>

      {/* ═══════ PANEL KANAN ═══════ */}
      <div className="flex-1 flex flex-col overflow-hidden">

        {/* Header */}
        <header className="border-b border-border bg-card px-6 py-3 flex items-center justify-between gap-4 flex-shrink-0">
          <div>
            <p className="text-[10px] font-semibold uppercase tracking-widest text-muted-foreground">Kalkulator Poin Cabang</p>
            <h1 className="text-base font-bold text-foreground leading-tight flex items-center gap-2">
              {selectedEmpData.name}
              <span className="text-xs font-medium bg-primary/10 text-primary rounded-full px-2 py-0.5">
                {MONTHS[selectedMonth]} {selectedYear}
              </span>
            </h1>
          </div>
          <div className="flex items-center gap-3 flex-wrap justify-end">
            {/* Auto-save indicator */}
            <div className="flex items-center gap-1.5">
              {savedAt ? (
                <>
                  <svg className="w-3.5 h-3.5 text-emerald-500 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7"/>
                  </svg>
                  <span className="text-[10px] text-emerald-600 font-medium">Tersimpan {savedAt}</span>
                </>
              ) : (
                <>
                  <svg className="w-3.5 h-3.5 text-muted-foreground animate-spin flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"/>
                  </svg>
                  <span className="text-[10px] text-muted-foreground">Menyimpan...</span>
                </>
              )}
            </div>

            {/* Riwayat periode */}
            {savedPeriods.length > 0 && (
              <div className="relative group">
                <button className="inline-flex items-center gap-1.5 rounded-md border border-border bg-background px-2.5 py-1.5 text-xs text-muted-foreground hover:text-foreground hover:bg-secondary transition-colors">
                  <svg className="w-3.5 h-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z"/>
                  </svg>
                  Riwayat ({savedPeriods.length})
                </button>
                <div className="absolute right-0 top-full mt-1 w-52 rounded-lg border border-border bg-card shadow-lg z-50 hidden group-hover:block py-1">
                  <p className="text-[10px] text-muted-foreground uppercase tracking-wider px-3 py-1 font-semibold">Periode Tersimpan</p>
                  {savedPeriods.map((p) => (
                    <div key={p.key} className="flex items-center justify-between px-3 py-1.5 hover:bg-secondary/50 group/item">
                      <span className="text-xs text-foreground">{p.label}</span>
                      <button
                        onClick={() => handleDeletePeriod(p.key)}
                        className="text-[10px] text-muted-foreground hover:text-destructive transition-colors hidden group-hover/item:block"
                        title="Hapus periode ini"
                      >Hapus</button>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Ekspor */}
            <button
              onClick={handleExport}
              className="inline-flex items-center gap-1.5 rounded-md border border-border bg-background px-2.5 py-1.5 text-xs font-medium text-muted-foreground hover:text-foreground hover:bg-secondary transition-colors"
            >
              <svg className="w-3.5 h-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"/>
              </svg>
              Ekspor
            </button>

            {/* Impor */}
            <button
              onClick={() => setShowImport(true)}
              className="inline-flex items-center gap-1.5 rounded-md border border-border bg-background px-2.5 py-1.5 text-xs font-medium text-muted-foreground hover:text-foreground hover:bg-secondary transition-colors"
            >
              <svg className="w-3.5 h-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l4-4m0 0l4 4m-4-4v12"/>
              </svg>
              Impor
            </button>

            {/* Maks poin */}
            <div className="text-right border-l border-border pl-3">
              <p className="text-[10px] text-muted-foreground">Maks/hari</p>
              <span className="text-sm font-bold text-foreground tabular-nums">{MAX_POIN} poin</span>
            </div>
            {/* Total bulan */}
            <div className="text-right">
              <p className="text-[10px] text-muted-foreground">Total Bulan</p>
              <ScorePill v={monthTotal} />
            </div>
          </div>
        </header>

        {/* Tabs */}
        <div className="flex border-b border-border bg-card px-6 flex-shrink-0">
          {([ ["tanggal","Poin per Tanggal"], ["simulasi","Input Nilai"], ["formula","Formula Excel"] ] as const).map(([key, label]) => (
            <button
              key={key}
              onClick={() => setActiveTab(key)}
              className={`px-4 py-3 text-xs font-semibold border-b-2 transition-colors ${
                activeTab === key
                  ? "border-primary text-primary"
                  : "border-transparent text-muted-foreground hover:text-foreground"
              }`}
            >
              {label}
            </button>
          ))}
        </div>

        {/* ═══ TAB: POIN PER TANGGAL ═══════════════════════════════════ */}
        {activeTab === "tanggal" && (
          <div className="flex-1 overflow-hidden flex flex-col">

            {/* Ringkasan bulan */}
            <div className="grid grid-cols-3 divide-x divide-border border-b border-border flex-shrink-0">
              <div className="px-6 py-3">
                <p className="text-[10px] text-muted-foreground uppercase tracking-wider">Total Poin Positif</p>
                <p className="text-xl font-bold text-emerald-600 tabular-nums mt-0.5">+{monthPos.toFixed(0)}</p>
              </div>
              <div className="px-6 py-3">
                <p className="text-[10px] text-muted-foreground uppercase tracking-wider">Total Penalti</p>
                <p className="text-xl font-bold text-red-600 tabular-nums mt-0.5">{monthNeg.toFixed(0)}</p>
              </div>
              <div className="px-6 py-3">
                <p className="text-[10px] text-muted-foreground uppercase tracking-wider">Saldo Bersih</p>
                <p className={`text-xl font-bold tabular-nums mt-0.5 ${monthTotal > 0 ? "text-emerald-600" : monthTotal < 0 ? "text-red-600" : "text-foreground"}`}>
                  {monthTotal > 0 ? "+" : ""}{monthTotal.toFixed(0)}
                </p>
              </div>
            </div>

            {/* Tabel tanggal 1–31 */}
            <div className="flex-1 overflow-auto">
              <table className="w-full text-xs border-collapse">
                <thead className="sticky top-0 z-10">
                  <tr className="bg-secondary border-b border-border">
                    <th className="py-2.5 px-3 text-left font-semibold uppercase tracking-wider text-muted-foreground w-16 sticky left-0 bg-secondary z-20">Tgl</th>
                    <th className="py-2.5 px-3 text-left font-semibold uppercase tracking-wider text-muted-foreground w-12">Hari</th>
                    {categories.map((c) => (
                      <th key={c.id} className="py-2 px-2 text-center font-semibold text-muted-foreground min-w-[4.5rem] max-w-[5rem]">
                        <div className="leading-tight" title={c.label}>
                          <div className="text-[9px] truncate max-w-[4rem]">{c.label.split(" ").slice(0,3).join(" ")}</div>
                          <div className="text-[10px] text-primary font-bold">{c.poin}p</div>
                        </div>
                      </th>
                    ))}
                    <th className="py-2.5 px-3 text-center font-semibold uppercase tracking-wider text-emerald-700 w-20">Positif</th>
                    <th className="py-2.5 px-3 text-center font-semibold uppercase tracking-wider text-red-600 w-20">Penalti</th>
                    <th className="py-2.5 px-3 text-center font-semibold uppercase tracking-wider text-foreground w-20">Total</th>
                    <th className="py-2.5 px-3 text-center font-semibold uppercase tracking-wider text-muted-foreground w-16">Aksi</th>
                  </tr>
                </thead>
                <tbody>
                  {DATES.map((d) => {
                    const isValid     = d <= daysInMonth
                    const dateObj     = isValid ? new Date(selectedYear, selectedMonth, d) : null
                    const dayName     = dateObj ? ["Min","Sen","Sel","Rab","Kam","Jum","Sab"][dateObj.getDay()] : "-"
                    const isSunday    = dateObj?.getDay() === 0
                    const isSaturday  = dateObj?.getDay() === 6
                    const isToday     = d === today.getDate() && selectedMonth === today.getMonth() && selectedYear === today.getFullYear()
                    const isSelected  = d === selectedDate
                    const { total: dTotal, pos: dPos, neg: dNeg } = calcDateScore(selectedEmp, d, dateData)
                    const hasAnyData  = Object.values(dateData[selectedEmp]?.[d] ?? {}).some((v) => v !== "")

                    return (
                      <tr
                        key={d}
                        onClick={() => { if (isValid) { setSelectedDate(d); setActiveTab("simulasi") } }}
                        className={`border-b border-border transition-colors cursor-pointer
                          ${!isValid ? "opacity-30 cursor-not-allowed" : ""}
                          ${isSelected ? "bg-primary/10" : isSunday ? "bg-red-50/30 dark:bg-red-950/10" : isSaturday ? "bg-amber-50/30 dark:bg-amber-950/10" : "hover:bg-accent"}
                        `}
                      >
                        {/* Tanggal */}
                        <td className={`py-2 px-3 sticky left-0 z-10 font-bold ${
                          isToday
                            ? "bg-primary text-primary-foreground"
                            : isSelected
                              ? "bg-primary/10 text-primary"
                              : isSunday
                                ? "bg-red-50/50 dark:bg-red-950/20 text-red-600"
                                : "bg-card text-foreground"
                        }`}>
                          {String(d).padStart(2,"0")}
                          {isToday && <span className="ml-1 text-[9px] font-normal opacity-75">hari ini</span>}
                        </td>
                        {/* Nama hari */}
                        <td className={`py-2 px-3 font-medium ${isSunday ? "text-red-500" : isSaturday ? "text-amber-600" : "text-muted-foreground"}`}>
                          {dayName}
                        </td>
                        {/* Nilai per kategori */}
                        {categories.map((c) => {
                          const raw = dateData[selectedEmp]?.[d]?.[c.id] ?? ""
                          const { score, status } = calcScore(raw, c.poin)
                          return (
                            <td key={c.id} className="py-1.5 px-1 text-center">
                              {status === "empty"
                                ? <span className="text-muted-foreground/30">—</span>
                                : status === "positif"
                                  ? <span className="tabular-nums font-bold text-emerald-700 bg-emerald-500/10 rounded px-1.5 py-0.5">+{score.toFixed(0)}</span>
                                  : <span className="tabular-nums font-bold text-red-600 bg-red-500/10 rounded px-1.5 py-0.5">-1</span>
                              }
                            </td>
                          )
                        })}
                        {/* Kolom total */}
                        <td className="py-2 px-3 text-center">
                          {dPos > 0 ? <span className="tabular-nums font-bold text-emerald-700">+{dPos.toFixed(0)}</span> : <span className="text-muted-foreground/40">—</span>}
                        </td>
                        <td className="py-2 px-3 text-center">
                          {dNeg < 0 ? <span className="tabular-nums font-bold text-red-600">{dNeg.toFixed(0)}</span> : <span className="text-muted-foreground/40">—</span>}
                        </td>
                        <td className="py-2 px-3 text-center">
                          {hasAnyData
                            ? <ScorePill v={dTotal} size="xs" />
                            : <span className="text-muted-foreground/30">—</span>
                          }
                        </td>
                        <td className="py-2 px-3 text-center">
                          {isValid && (
                            <span className={`inline-block text-[10px] font-medium px-1.5 py-0.5 rounded-full border ${
                              isSelected
                                ? "bg-primary text-primary-foreground border-primary"
                                : "border-border text-muted-foreground hover:border-primary hover:text-primary"
                            }`}>
                              Input
                            </span>
                          )}
                        </td>
                      </tr>
                    )
                  })}
                </tbody>

                {/* Baris total bulan */}
                <tfoot className="sticky bottom-0">
                  <tr className="border-t-2 border-border bg-secondary/80">
                    <td colSpan={2} className="py-3 px-3 text-xs font-bold uppercase tracking-wider text-foreground sticky left-0 bg-secondary/80">
                      TOTAL
                    </td>
                    {categories.map((c) => {
                      const catTotal = DATES.reduce((s, d) => {
                        const { score, status } = calcScore(dateData[selectedEmp]?.[d]?.[c.id] ?? "", c.poin)
                        return status !== "empty" ? s + score : s
                      }, 0)
                      return (
                        <td key={c.id} className="py-3 px-1 text-center">
                          {catTotal !== 0
                            ? <ScorePill v={catTotal} size="xs" />
                            : <span className="text-muted-foreground/30">—</span>
                          }
                        </td>
                      )
                    })}
                    <td className="py-3 px-3 text-center font-bold text-emerald-700 tabular-nums">
                      {monthPos > 0 ? `+${monthPos.toFixed(0)}` : "—"}
                    </td>
                    <td className="py-3 px-3 text-center font-bold text-red-600 tabular-nums">
                      {monthNeg < 0 ? monthNeg.toFixed(0) : "—"}
                    </td>
                    <td className="py-3 px-3 text-center">
                      <ScorePill v={monthTotal} />
                    </td>
                    <td />
                  </tr>
                </tfoot>
              </table>
            </div>
          </div>
        )}

        {/* ═══ TAB: INPUT NILAI ════════════════════════════════════════ */}
        {activeTab === "simulasi" && (
          <div className="flex-1 overflow-y-auto px-6 py-5 space-y-5">

            {/* Pilih tanggal */}
            <div className="flex items-center gap-3 flex-wrap">
              <span className="text-xs text-muted-foreground font-medium">Pilih Tanggal:</span>
              <div className="flex flex-wrap gap-1.5">
                {DATES.map((d) => {
                  const isValid = d <= daysInMonth
                  const hasData = Object.values(dateData[selectedEmp]?.[d] ?? {}).some((v) => v !== "")
                  const { total } = calcDateScore(selectedEmp, d, dateData)
                  const isToday = d === today.getDate() && selectedMonth === today.getMonth() && selectedYear === today.getFullYear()
                  return (
                    <button
                      key={d}
                      disabled={!isValid}
                      onClick={() => setSelectedDate(d)}
                      title={`Tanggal ${d}${hasData ? ` — skor: ${total}` : ""}`}
                      className={`relative w-8 h-8 rounded-lg text-xs font-bold transition-all
                        ${!isValid ? "opacity-20 cursor-not-allowed bg-secondary text-muted-foreground" :
                          selectedDate === d
                            ? "bg-primary text-primary-foreground shadow-sm"
                            : isToday
                              ? "bg-primary/20 text-primary border-2 border-primary"
                              : "bg-secondary text-foreground hover:bg-primary/15"
                        }
                      `}
                    >
                      {d}
                      {hasData && isValid && selectedDate !== d && (
                        <span className={`absolute -top-0.5 -right-0.5 w-2 h-2 rounded-full border border-card ${total > 0 ? "bg-emerald-500" : total < 0 ? "bg-red-500" : "bg-secondary"}`} />
                      )}
                    </button>
                  )
                })}
              </div>
            </div>

            {/* Header tanggal terpilih */}
            <div className="rounded-xl border border-border bg-card overflow-hidden">
              <div className="flex items-center justify-between px-5 py-3 border-b border-border bg-secondary/40">
                <div>
                  <span className="text-sm font-bold text-foreground">
                    {String(selectedDate).padStart(2,"0")} {MONTHS[selectedMonth]} {selectedYear}
                  </span>
                  <span className="ml-2 text-xs text-muted-foreground">&mdash; {selectedEmpData.name}</span>
                </div>
                <button
                  onClick={() => resetDate(selectedDate)}
                  className="text-xs text-muted-foreground hover:text-destructive border border-border rounded-md px-3 py-1.5 hover:bg-background transition-colors"
                >
                  Reset Tanggal Ini
                </button>
              </div>

              {/* Grid input */}
              <div className="px-5 py-4">
                <div className="grid grid-cols-[2rem_1fr_3rem_5rem_5rem_5rem] gap-x-3 pb-2 mb-2 border-b border-border">
                  <span className="text-[10px] font-semibold uppercase tracking-wider text-muted-foreground">#</span>
                  <span className="text-[10px] font-semibold uppercase tracking-wider text-muted-foreground">Kategori</span>
                  <span className="text-[10px] font-semibold uppercase tracking-wider text-muted-foreground text-center">Poin</span>
                  <span className="text-[10px] font-semibold uppercase tracking-wider text-muted-foreground text-center">Nilai</span>
                  <span className="text-[10px] font-semibold uppercase tracking-wider text-muted-foreground text-center">Skor</span>
                  <span className="text-[10px] font-semibold uppercase tracking-wider text-muted-foreground text-center">Status</span>
                </div>
                <div className="space-y-2">
                  {dateDetail.map((row) => (
                    <div key={row.id} className="grid grid-cols-[2rem_1fr_3rem_5rem_5rem_5rem] gap-x-3 items-center">
                      <span className="text-xs font-mono text-muted-foreground text-right">{row.no}</span>
                      <span className="text-xs text-foreground truncate">{row.label}</span>
                      <span className="text-xs font-bold text-primary text-center tabular-nums">{row.poin}</span>
                      <input
                        type="number"
                        value={row.raw}
                        onChange={(e) => setVal(selectedDate, row.id, e.target.value)}
                        placeholder="0"
                        className="rounded-md border border-border bg-background px-2 py-1 text-xs font-mono text-center text-foreground focus:outline-none focus:ring-2 focus:ring-primary/50 w-full"
                      />
                      <div className="flex justify-center">
                        {row.status !== "empty"
                          ? <span className={`tabular-nums text-xs font-bold rounded-md px-2 py-0.5 ${row.status === "positif" ? "text-emerald-700 bg-emerald-500/15" : "text-red-600 bg-red-500/15"}`}>
                              {row.status === "positif" ? `+${row.score.toFixed(0)}` : "-1"}
                            </span>
                          : <span className="text-xs text-muted-foreground">—</span>
                        }
                      </div>
                      <div className="flex justify-center">
                        {row.status !== "empty"
                          ? <span className={`text-[10px] font-medium px-1.5 py-0.5 rounded-full ${row.status === "positif" ? "bg-emerald-500/10 text-emerald-700" : "bg-red-500/10 text-red-600"}`}>
                              {row.status === "positif" ? "Positif" : "Nihil"}
                            </span>
                          : <span className="text-[10px] text-muted-foreground">—</span>
                        }
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Total baris */}
              <div className="border-t-2 border-border">
                <div className="grid grid-cols-[2rem_1fr_3rem_5rem_5rem_5rem] gap-x-3 items-center px-5 py-2.5 bg-emerald-50/50 dark:bg-emerald-950/20 border-b border-border">
                  <span /><span className="text-xs font-semibold text-emerald-700 uppercase tracking-wide">Jumlah Positif</span>
                  <span className="text-center text-xs text-muted-foreground">{dateDetail.filter(r=>r.status==="positif").length} kat.</span>
                  <span />
                  <span className="flex justify-center"><span className="tabular-nums text-sm font-bold text-emerald-700 bg-emerald-500/15 rounded-md px-2 py-0.5">+{detailPos.toFixed(0)}</span></span>
                  <span />
                </div>
                <div className="grid grid-cols-[2rem_1fr_3rem_5rem_5rem_5rem] gap-x-3 items-center px-5 py-2.5 bg-red-50/50 dark:bg-red-950/20 border-b border-border">
                  <span /><span className="text-xs font-semibold text-red-700 uppercase tracking-wide">Jumlah Penalti</span>
                  <span className="text-center text-xs text-muted-foreground">{dateDetail.filter(r=>r.status==="nihil").length} kat.</span>
                  <span />
                  <span className="flex justify-center"><span className="tabular-nums text-sm font-bold text-red-600 bg-red-500/15 rounded-md px-2 py-0.5">{detailNeg.toFixed(0)}</span></span>
                  <span />
                </div>
                <div className="grid grid-cols-[2rem_1fr_3rem_5rem_5rem_5rem] gap-x-3 items-center px-5 py-3 bg-primary/5">
                  <span />
                  <div>
                    <span className="text-sm font-bold text-foreground uppercase tracking-wide">Total Skor Tanggal Ini</span>
                    <p className="text-[10px] text-muted-foreground mt-0.5">
                      +{detailPos.toFixed(0)} + ({detailNeg.toFixed(0)}) = {detailTotal > 0 ? "+" : ""}{detailTotal.toFixed(0)}
                    </p>
                  </div>
                  <span className="text-center text-xs text-muted-foreground">{dateDetail.filter(r=>r.status!=="empty").length}/{categories.length}</span>
                  <span />
                  <span className="flex justify-center">
                    <span className={`tabular-nums text-base font-extrabold rounded-lg px-3 py-1 border-2 ${
                      detailTotal > 0 ? "text-emerald-700 bg-emerald-500/15 border-emerald-400/40"
                      : detailTotal < 0 ? "text-red-600 bg-red-500/15 border-red-400/40"
                      : "text-foreground bg-secondary border-border"
                    }`}>
                      {detailTotal > 0 ? "+" : ""}{detailTotal.toFixed(0)}
                    </span>
                  </span>
                  <span className="text-center text-[10px] text-muted-foreground">/ {MAX_POIN}</span>
                </div>
              </div>
            </div>

            {/* Ringkasan bulan di bawah input */}
            <div className="rounded-xl border border-border bg-card p-4">
              <p className="text-xs font-semibold text-muted-foreground uppercase tracking-wider mb-3">
                Akumulasi Bulan &mdash; {MONTHS[selectedMonth]} {selectedYear}
              </p>
              <div className="grid grid-cols-3 gap-3">
                <div className="rounded-lg bg-emerald-500/10 p-3 text-center">
                  <p className="text-[10px] text-emerald-700 font-semibold uppercase">Total Positif</p>
                  <p className="text-lg font-bold text-emerald-700 tabular-nums">+{monthPos.toFixed(0)}</p>
                </div>
                <div className="rounded-lg bg-red-500/10 p-3 text-center">
                  <p className="text-[10px] text-red-600 font-semibold uppercase">Total Penalti</p>
                  <p className="text-lg font-bold text-red-600 tabular-nums">{monthNeg.toFixed(0)}</p>
                </div>
                <div className={`rounded-lg p-3 text-center ${monthTotal > 0 ? "bg-primary/10" : monthTotal < 0 ? "bg-destructive/10" : "bg-secondary"}`}>
                  <p className="text-[10px] text-muted-foreground font-semibold uppercase">Saldo Bersih</p>
                  <p className={`text-lg font-bold tabular-nums ${monthTotal > 0 ? "text-primary" : monthTotal < 0 ? "text-destructive" : "text-foreground"}`}>
                    {monthTotal > 0 ? "+" : ""}{monthTotal.toFixed(0)}
                  </p>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ═══ TAB: FORMULA EXCEL ════════════════════════════════════== */}
        {activeTab === "formula" && (
          <div className="flex-1 overflow-y-auto px-6 py-5 space-y-5">
            {/* Logika */}
            <section>
              <h2 className="text-xs font-semibold uppercase tracking-wider text-muted-foreground mb-3">Logika Penilaian</h2>
              <div className="grid grid-cols-2 gap-3">
                <div className="rounded-xl border border-emerald-200 bg-emerald-50/50 dark:border-emerald-900/40 dark:bg-emerald-950/20 p-4 flex items-start gap-3">
                  <div className="mt-0.5 flex-shrink-0 w-7 h-7 rounded-full bg-emerald-500/15 flex items-center justify-center">
                    <svg className="w-3.5 h-3.5 text-emerald-600" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}><path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7"/></svg>
                  </div>
                  <div>
                    <p className="text-xs font-semibold text-emerald-800 dark:text-emerald-300">Jika Positif (F &gt; 0)</p>
                    <p className="text-xs text-emerald-700/80 mt-0.5 leading-relaxed">
                      Nilai sel dikali <strong>bobot poin</strong> kategori.<br/>
                      Contoh: F6 = 3, poin = 8 &rarr; <strong>3 &times; 8 = 24</strong>
                    </p>
                  </div>
                </div>
                <div className="rounded-xl border border-red-200 bg-red-50/50 dark:border-red-900/40 dark:bg-red-950/20 p-4 flex items-start gap-3">
                  <div className="mt-0.5 flex-shrink-0 w-7 h-7 rounded-full bg-red-500/15 flex items-center justify-center">
                    <svg className="w-3.5 h-3.5 text-red-600" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}><path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12"/></svg>
                  </div>
                  <div>
                    <p className="text-xs font-semibold text-red-800 dark:text-red-300">Jika Nihil (F &lt; 1)</p>
                    <p className="text-xs text-red-700/80 mt-0.5 leading-relaxed">
                      Penalti <strong>-1</strong> untuk kategori tersebut.<br/>
                      Contoh: F6 = 0 &rarr; <strong>-1</strong>
                    </p>
                  </div>
                </div>
              </div>
            </section>

            {/* Tabel kategori */}
            <section>
              <h2 className="text-xs font-semibold uppercase tracking-wider text-muted-foreground mb-3">Tabel Kategori &amp; Bobot</h2>
              <div className="rounded-xl border border-border overflow-hidden shadow-sm">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="bg-secondary border-b border-border">
                      <th className="py-2.5 px-4 text-left text-xs font-semibold uppercase tracking-wider text-muted-foreground w-8">No</th>
                      <th className="py-2.5 px-4 text-left text-xs font-semibold uppercase tracking-wider text-muted-foreground">Kategori</th>
                      <th className="py-2.5 px-4 text-center text-xs font-semibold uppercase tracking-wider text-muted-foreground w-16">Poin</th>
                      <th className="py-2.5 px-4 text-center text-xs font-semibold uppercase tracking-wider text-muted-foreground w-28">Hasil Positif</th>
                      <th className="py-2.5 px-4 text-center text-xs font-semibold uppercase tracking-wider text-muted-foreground w-20">Nihil</th>
                    </tr>
                  </thead>
                  <tbody>
                    {categories.map((cat, idx) => (
                      <tr key={cat.id} className={`border-b border-border hover:bg-accent transition-colors ${idx % 2 === 0 ? "bg-card" : "bg-secondary/30"}`}>
                        <td className="py-2.5 px-4 font-mono text-xs text-muted-foreground">{cat.no}</td>
                        <td className="py-2.5 px-4 font-medium text-foreground text-xs">{cat.label}</td>
                        <td className="py-2.5 px-4 text-center">
                          <span className="inline-flex items-center justify-center rounded-md bg-primary/10 text-primary font-bold px-2 py-0.5 text-xs tabular-nums">{cat.poin}</span>
                        </td>
                        <td className="py-2.5 px-4 text-center">
                          <span className="inline-flex items-center justify-center rounded-md bg-emerald-500/10 text-emerald-700 font-mono text-xs px-2 py-0.5">F &times; {cat.poin}</span>
                        </td>
                        <td className="py-2.5 px-4 text-center">
                          <span className="inline-flex items-center justify-center rounded-md bg-destructive/10 text-destructive font-bold px-2 py-0.5 text-xs tabular-nums">-1</span>
                        </td>
                      </tr>
                    ))}
                    <tr className="bg-primary text-primary-foreground">
                      <td colSpan={2} className="py-2.5 px-4 text-xs font-semibold uppercase tracking-wider opacity-80">Total</td>
                      <td className="py-2.5 px-4 text-center font-bold tabular-nums">{MAX_POIN}</td>
                      <td className="py-2.5 px-4 text-center opacity-60 text-xs">nilai &times; poin</td>
                      <td className="py-2.5 px-4 text-center opacity-60 text-xs">-{categories.length}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </section>

            {/* Formula builder */}
            <section className="space-y-4">
              <div className="flex items-center justify-between gap-3">
                <h2 className="text-xs font-semibold uppercase tracking-wider text-muted-foreground">Pilih Varian Formula</h2>
                <div className="flex items-center gap-2">
                  <label htmlFor="cellref" className="text-xs text-muted-foreground">Referensi Sel:</label>
                  <input
                    id="cellref"
                    type="text"
                    value={cellRef}
                    onChange={(e) => setCellRef(e.target.value.toUpperCase())}
                    className="w-16 rounded-md border border-border bg-background px-2 py-1 text-xs font-mono text-center text-foreground focus:outline-none focus:ring-2 focus:ring-primary/50"
                    maxLength={5}
                    placeholder="F6"
                  />
                </div>
              </div>
              <div className="grid grid-cols-2 gap-3">
                {formulaVariants.map((v) => (
                  <button
                    key={v.key}
                    onClick={() => setActiveVariant(v.key)}
                    className={`text-left rounded-xl border p-3 transition-all ${
                      activeVariant === v.key ? "border-primary bg-primary/5 shadow-sm" : "border-border bg-card hover:bg-secondary/50"
                    }`}
                  >
                    <div className="flex items-center justify-between mb-1">
                      <span className="text-xs font-semibold text-foreground">{v.label}</span>
                      <span className={`text-[10px] font-medium px-1.5 py-0.5 rounded-full ${activeVariant === v.key ? "bg-primary text-primary-foreground" : "bg-secondary text-muted-foreground"}`}>
                        {v.badge}
                      </span>
                    </div>
                    <p className="text-[11px] text-muted-foreground leading-relaxed">{v.desc}</p>
                  </button>
                ))}
              </div>
              <div className="rounded-xl border border-border bg-card overflow-hidden">
                <div className="flex items-center justify-between px-4 py-3 border-b border-border bg-secondary/40">
                  <span className="text-xs font-semibold text-foreground">
                    {formulaVariants.find((v) => v.key === activeVariant)?.label}
                  </span>
                  <CopyButton text={formula} />
                </div>
                <div className="p-4 overflow-x-auto max-h-56">
                  <pre className="font-mono text-xs text-foreground leading-6 whitespace-pre">{formula}</pre>
                </div>
              </div>
            </section>
          </div>
        )}

      </div>

      {/* ═══ MODAL IMPOR ══════════════════════════════════════════ */}
      {showImport && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 backdrop-blur-sm" onClick={() => setShowImport(false)}>
          <div className="bg-card border border-border rounded-2xl shadow-2xl w-full max-w-lg mx-4 overflow-hidden" onClick={(e) => e.stopPropagation()}>
            <div className="flex items-center justify-between px-5 py-4 border-b border-border">
              <div>
                <h2 className="text-sm font-bold text-foreground">Impor Data</h2>
                <p className="text-xs text-muted-foreground mt-0.5">Tempel isi file JSON hasil ekspor di bawah ini</p>
              </div>
              <button onClick={() => setShowImport(false)} className="text-muted-foreground hover:text-foreground transition-colors">
                <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12"/>
                </svg>
              </button>
            </div>
            <div className="px-5 py-4 space-y-3">
              <textarea
                rows={8}
                value={importText}
                onChange={(e) => { setImportText(e.target.value); setImportError("") }}
                placeholder='{"month":0,"year":2025,"data":{...}}'
                className="w-full rounded-lg border border-border bg-background px-3 py-2.5 text-xs font-mono text-foreground focus:outline-none focus:ring-2 focus:ring-primary/50 resize-none"
              />
              {importError && (
                <p className="text-xs text-destructive flex items-center gap-1.5">
                  <svg className="w-3.5 h-3.5 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v2m0 4h.01M12 3a9 9 0 100 18A9 9 0 0012 3z"/>
                  </svg>
                  {importError}
                </p>
              )}
            </div>
            <div className="px-5 py-4 border-t border-border flex items-center justify-end gap-2">
              <button
                onClick={() => { setShowImport(false); setImportText(""); setImportError("") }}
                className="rounded-lg border border-border bg-background px-4 py-2 text-xs font-medium text-muted-foreground hover:text-foreground hover:bg-secondary transition-colors"
              >Batal</button>
              <button
                onClick={handleImport}
                disabled={!importText.trim()}
                className="rounded-lg bg-primary text-primary-foreground px-4 py-2 text-xs font-semibold hover:opacity-90 disabled:opacity-40 disabled:cursor-not-allowed transition-opacity"
              >Impor Data</button>
            </div>
          </div>
        </div>
      )}

    </div>
  )
}
