import { useState, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ── CONSTANTS ────────────────────────────────────────────────
const PASS_DEFAULT = 460;

const PRODI_LIST = [
  { nama: "Statistika",       short: "Statistika",   color: "#a78bfa", bg: "#2e1065" },
  { nama: "Matematika",       short: "Matematika",   color: "#22d3ee", bg: "#082f49" },
  { nama: "Ilmu Aktuaria",    short: "Aktuaria",     color: "#2dd4bf", bg: "#042f2e" },
  { nama: "Fisika",           short: "Fisika",       color: "#fbbf24", bg: "#292524" },
  { nama: "Sistem Informasi", short: "Sist. Info",   color: "#34d399", bg: "#052e16" },
  { nama: "Informatika",      short: "Informatika",  color: "#60a5fa", bg: "#0c1a3d" },
  { nama: "Bisnis Digital",   short: "Bisnis Dig.",  color: "#f472b6", bg: "#3b0764" },
  { nama: "Teknik Elektro",   short: "T. Elektro",   color: "#fb923c", bg: "#431407" },
];
const PRODI_MAP   = Object.fromEntries(PRODI_LIST.map(p => [p.nama, p]));
const PRODI_NAMES = PRODI_LIST.map(p => p.nama);

// ── COLUMN ALIASES ───────────────────────────────────────────
const COL_MAP = {
  "nama mahasiswa":"nama","nama":"nama","name":"nama","student name":"nama",
  "username":"username","nim":"username","nrp":"username","user":"username",
  "prodi":"prodi","program studi":"prodi","jurusan":"prodi","department":"prodi",
  "listening":"listening","list.":"listening",
  "grammar":"grammar","structure":"grammar","str.":"grammar",
  "reading":"reading","read.":"reading",
  "total":"total","total skor":"total","total score":"total","skor total":"total",
  "iaet":"iaet","result":"iaet","result (iaet)":"iaet","skor iaet":"iaet",
  "iaet score":"iaet","score":"iaet","nilai":"iaet","nilai iaet":"iaet",
  "email":"email",
};

// ── 8 KOLOM WAJIB — jika kosong → baris ditolak ─────────────
const REQUIRED = ["nama","username","prodi","listening","grammar","reading","total","iaet"];
const REQ_LBL  = {
  nama:"Nama Mahasiswa", username:"Username", prodi:"Prodi",
  listening:"Listening", grammar:"Grammar",   reading:"Reading",
  total:"Total",         iaet:"Result (IAET)",
};

// ── SAMPLE DATA (32 mahasiswa, 8 prodi) ──────────────────────
const SAMPLE = [
  {nama:"Anisa Rahmawati",   username:"a.rahmawati",  prodi:"Statistika",       listening:42,grammar:33,reading:38,total:113,iaet:534},
  {nama:"Bima Saputra",      username:"b.saputra",    prodi:"Matematika",       listening:55,grammar:44,reading:52,total:151,iaet:610},
  {nama:"Citra Kusuma",      username:"c.kusuma",     prodi:"Ilmu Aktuaria",    listening:38,grammar:22,reading:31,total:91, iaet:491},
  {nama:"Dafa Maulana",      username:"d.maulana",    prodi:"Fisika",           listening:61,grammar:46,reading:59,total:166,iaet:638},
  {nama:"Elisa Putri",       username:"e.putri",      prodi:"Sistem Informasi", listening:36,grammar:20,reading:28,total:84, iaet:477},
  {nama:"Fajar Nugroho",     username:"f.nugroho",    prodi:"Informatika",      listening:50,grammar:38,reading:47,total:135,iaet:578},
  {nama:"Gita Nurdiana",     username:"g.nurdiana",   prodi:"Bisnis Digital",   listening:29,grammar:18,reading:26,total:73, iaet:454},
  {nama:"Hendra Wijaya",     username:"h.wijaya",     prodi:"Teknik Elektro",   listening:57,grammar:43,reading:54,total:154,iaet:615},
  {nama:"Indah Permata",     username:"i.permata",    prodi:"Statistika",       listening:33,grammar:19,reading:29,total:81, iaet:469},
  {nama:"Jaka Santoso",      username:"j.santoso",    prodi:"Matematika",       listening:48,grammar:35,reading:44,total:127,iaet:562},
  {nama:"Kartika Dewi",      username:"k.dewi",       prodi:"Ilmu Aktuaria",    listening:62,grammar:47,reading:60,total:169,iaet:643},
  {nama:"Lukman Hakim",      username:"l.hakim",      prodi:"Fisika",           listening:31,grammar:17,reading:25,total:73, iaet:454},
  {nama:"Mira Sari",         username:"m.sari",       prodi:"Sistem Informasi", listening:53,grammar:40,reading:49,total:142,iaet:591},
  {nama:"Naufal Arif",       username:"n.arif",       prodi:"Informatika",      listening:28,grammar:16,reading:24,total:68, iaet:441},
  {nama:"Olivia Susanti",    username:"o.susanti",    prodi:"Bisnis Digital",   listening:45,grammar:32,reading:41,total:118,iaet:543},
  {nama:"Pandu Pratama",     username:"p.pratama",    prodi:"Teknik Elektro",   listening:39,grammar:25,reading:34,total:98, iaet:504},
  {nama:"Qisthi Amalia",     username:"q.amalia",     prodi:"Statistika",       listening:58,grammar:45,reading:55,total:158,iaet:622},
  {nama:"Ridho Fernanda",    username:"r.fernanda",   prodi:"Matematika",       listening:34,grammar:21,reading:30,total:85, iaet:479},
  {nama:"Salsabila Ahmad",   username:"s.ahmad",      prodi:"Ilmu Aktuaria",    listening:46,grammar:34,reading:43,total:123,iaet:554},
  {nama:"Taufiq Hidayat",    username:"t.hidayat",    prodi:"Fisika",           listening:64,grammar:49,reading:62,total:175,iaet:657},
  {nama:"Ulfiana Sari",      username:"u.sari",       prodi:"Sistem Informasi", listening:37,grammar:23,reading:33,total:93, iaet:494},
  {nama:"Valdi Ramadhan",    username:"v.ramadhan",   prodi:"Informatika",      listening:51,grammar:39,reading:48,total:138,iaet:584},
  {nama:"Winda Lestari",     username:"w.lestari",    prodi:"Bisnis Digital",   listening:26,grammar:15,reading:22,total:63, iaet:428},
  {nama:"Xena Puspita",      username:"x.puspita",    prodi:"Teknik Elektro",   listening:43,grammar:30,reading:39,total:112,iaet:532},
  {nama:"Yoga Pratama",      username:"y.pratama",    prodi:"Statistika",       listening:67,grammar:50,reading:65,total:182,iaet:671},
  {nama:"Zahra Aulia",       username:"z.aulia",      prodi:"Matematika",       listening:30,grammar:18,reading:27,total:75, iaet:458},
  {nama:"Andi Wirawan",      username:"a.wirawan",    prodi:"Ilmu Aktuaria",    listening:44,grammar:31,reading:40,total:115,iaet:538},
  {nama:"Bella Kinasih",     username:"b.kinasih",    prodi:"Fisika",           listening:52,grammar:41,reading:50,total:143,iaet:593},
  {nama:"Cahya Nugraha",     username:"c.nugraha",    prodi:"Sistem Informasi", listening:35,grammar:22,reading:31,total:88, iaet:483},
  {nama:"Dini Rahayu",       username:"d.rahayu",     prodi:"Informatika",      listening:59,grammar:44,reading:56,total:159,iaet:624},
  {nama:"Efan Maulana",      username:"e.maulana",    prodi:"Bisnis Digital",   listening:40,grammar:27,reading:36,total:103,iaet:514},
  {nama:"Fikri Habibie",     username:"f.habibie",    prodi:"Teknik Elektro",   listening:47,grammar:36,reading:44,total:127,iaet:562},
];

// ── HELPERS ──────────────────────────────────────────────────
const getCat = (iaet) => {
  if (iaet >= 500) return { label:"UNGGUL", color:"#22c55e", bg:"#052e16", border:"#166534" };
  if (iaet >= 460) return { label:"LULUS",  color:"#4ade80", bg:"#052e16", border:"#166534" };
  if (iaet >= 400) return { label:"SEDANG", color:"#fb923c", bg:"#1c0f00", border:"#9a3412" };
  if (iaet >= 350) return { label:"RENDAH", color:"#f87171", bg:"#1c0404", border:"#991b1b" };
  return               { label:"KRITIS", color:"#ef4444", bg:"#1c0000", border:"#b91c1c" };
};
const getPred = (iaet) => {
  if (iaet >= 500) return "Memuaskan";
  if (iaet >= 460) return "Baik";
  if (iaet >= 400) return "Cukup";
  if (iaet >= 350) return "Kurang";
  return "Sangat Kurang";
};

// enrichRow: TIDAK ada kalkulasi — Total & IAET langsung dari data
const enrich = (r, thr = PASS_DEFAULT) => ({
  nama:      String(r.nama ?? ""),
  username:  String(r.username ?? ""),
  prodi:     String(r.prodi ?? ""),
  listening: Number(r.listening),
  grammar:   Number(r.grammar),
  reading:   Number(r.reading),
  total:     Number(r.total),
  iaet:      Number(r.iaet),
  predicate: getPred(Number(r.iaet)),
  passed:    Number(r.iaet) >= thr,
  email:     r.email || `${r.username}@student.itk.ac.id`,
  cat:       getCat(Number(r.iaet)),
});

// validateRow: 8 kolom wajib — jika kosong → tolak
const validate = (raw) => {
  const m = {};
  Object.entries(raw).forEach(([k, v]) => {
    const key = COL_MAP[k.toLowerCase().trim()];
    if (key) m[key] = v;
  });
  const miss = REQUIRED.filter(c => m[c] == null || String(m[c]).trim() === "");
  if (miss.length) return { ok: false, reason: `Kolom wajib kosong: ${miss.map(c => REQ_LBL[c]).join(", ")}` };
  const noNum = ["listening","grammar","reading","total","iaet"].filter(c => isNaN(Number(m[c])));
  if (noNum.length) return { ok: false, reason: `Bukan angka: ${noNum.join(", ")}` };
  return { ok: true, data: m };
};

// ── MAILTO HELPERS ───────────────────────────────────────────
const EMAIL_TEMPLATES = {
  KRITIS: (r) => ({
    subject: `[GREAT ITK] PERINGATAN KRITIS — Nilai IAET Anda Sangat Rendah`,
    body:
`Yth. ${r.nama},

Kami ingin menyampaikan bahwa hasil tes IAET Anda saat ini berada pada kategori KRITIS.

Detail Hasil Tes:
- Nama        : ${r.nama}
- Program Studi: ${r.prodi}
- Listening   : ${r.listening}
- Grammar     : ${r.grammar}
- Reading     : ${r.reading}
- Total Skor  : ${r.total}
- Skor IAET   : ${r.iaet}
- Kategori    : KRITIS (IAET < 350)

Anda WAJIB mengikuti program remedial intensif segera. Silakan hubungi koordinator Program GREAT FSTI ITK untuk jadwal remedial.

Hormat kami,
Koordinator Program GREAT
Fakultas Sains dan Teknologi Informasi — ITK`,
  }),
  RENDAH: (r) => ({
    subject: `[GREAT ITK] Perhatian — Nilai IAET Anda Perlu Ditingkatkan`,
    body:
`Yth. ${r.nama},

Berdasarkan hasil tes IAET terbaru, nilai Anda berada pada kategori RENDAH dan belum memenuhi standar kelulusan.

Detail Hasil Tes:
- Nama        : ${r.nama}
- Program Studi: ${r.prodi}
- Listening   : ${r.listening}
- Grammar     : ${r.grammar}
- Reading     : ${r.reading}
- Total Skor  : ${r.total}
- Skor IAET   : ${r.iaet}
- Kategori    : RENDAH (IAET 350–399)
- Target      : IAET ≥ 460

Anda dianjurkan untuk mengikuti program remedial dan aktif menggunakan sumber belajar yang tersedia di Program GREAT.

Hormat kami,
Koordinator Program GREAT
Fakultas Sains dan Teknologi Informasi — ITK`,
  }),
  SEDANG: (r, thr = 460) => ({
    subject: `[GREAT ITK] Notifikasi EAP — Tingkatkan Skor IAET Anda`,
    body:
`Yth. ${r.nama},

Hasil tes IAET Anda saat ini berada pada kategori SEDANG. Anda sudah menunjukkan kemajuan, namun masih perlu peningkatan untuk mencapai standar kelulusan.

Detail Hasil Tes:
- Nama        : ${r.nama}
- Program Studi: ${r.prodi}
- Listening   : ${r.listening}
- Grammar     : ${r.grammar}
- Reading     : ${r.reading}
- Total Skor  : ${r.total}
- Skor IAET   : ${r.iaet}
- Kategori    : SEDANG (IAET 400–${thr - 1})
- Target      : IAET ≥ ${thr}

Kami menganjurkan Anda untuk mengikuti program English for Academic Purposes (EAP) dan memanfaatkan fasilitas English Community yang tersedia.

Hormat kami,
Koordinator Program GREAT
Fakultas Sains dan Teknologi Informasi — ITK`,
  }),
};

// Buka draft email di Gmail/Outlook via mailto:
const openMailto = (to, subject, body) => {
  const url = `mailto:${encodeURIComponent(to)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
  window.open(url, "_blank");
};

// Bulk: buka mailto dengan semua penerima di BCC, body berisi daftar nama
const openBulkMailto = (mhs, category, thr = 460) => {
  const bcc    = mhs.map(r => r.email).join(",");
  const tmplFn = EMAIL_TEMPLATES[category];
  // Gunakan data mahasiswa pertama sebagai template, list semua di body
  const daftar = mhs.map(r => `  - ${r.nama} (${r.prodi}) | IAET: ${r.iaet}`).join("\n");
  const subj   = `[GREAT ITK] Notifikasi Program GREAT — Kategori ${category}`;
  const body   =
`Yth. Mahasiswa FSTI ITK,

Berikut ini adalah notifikasi hasil tes IAET untuk kategori ${category}.

Daftar Penerima:
${daftar}

${category === "KRITIS"
  ? "Seluruh mahasiswa di atas WAJIB mengikuti program remedial intensif segera."
  : category === "RENDAH"
  ? "Seluruh mahasiswa di atas dianjurkan untuk mengikuti program remedial."
  : `Seluruh mahasiswa di atas dianjurkan mengikuti program EAP untuk mencapai target IAET ≥ ${thr}.`}

Untuk informasi lebih lanjut, silakan hubungi koordinator Program GREAT FSTI ITK.

Hormat kami,
Koordinator Program GREAT
Fakultas Sains dan Teknologi Informasi — ITK`;
  const url = `mailto:?bcc=${encodeURIComponent(bcc)}&subject=${encodeURIComponent(subj)}&body=${encodeURIComponent(body)}`;
  window.open(url, "_blank");
};

// ── DESIGN TOKENS ────────────────────────────────────────────
const C = {
  bg:"#0b1120", panel:"#111827", card:"#1e293b", border:"#1e2d3d",
  text:"#e2e8f0", muted:"#64748b", accent:"#fde962",
  blue:"#3b82f6", green:"#22c55e", red:"#ef4444", orange:"#f97316",
};

// ── CHART COMPONENTS ─────────────────────────────────────────
function HBar({ pct, color, h = 5 }) {
  return (
    <div style={{ height: h, borderRadius: h, background: "#0f172a", overflow: "hidden" }}>
      <div style={{ height: "100%", width: `${Math.min(pct, 100)}%`, background: color, borderRadius: h }} />
    </div>
  );
}

function Pill({ label, color, bg, border }) {
  return (
    <span style={{ display:"inline-block", padding:"2px 9px", borderRadius:20,
      background: bg, color, border: `1px solid ${border || bg}`, fontSize:14, fontWeight:700 }}>
      {label}
    </span>
  );
}

function DonutChart({ segments, size = 92 }) {
  const total = segments.reduce((a, b) => a + b.value, 0) || 1;
  let cum = 0;
  return (
    <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`}>
      {segments.map((seg, i) => {
        const pct = seg.value / total, s = cum * 360, e = (cum + pct) * 360; cum += pct;
        const toR = d => (d - 90) * Math.PI / 180, r = size * .38, cx = size / 2, cy = size / 2;
        const x1 = cx + r * Math.cos(toR(s)), y1 = cy + r * Math.sin(toR(s));
        const x2 = cx + r * Math.cos(toR(e)), y2 = cy + r * Math.sin(toR(e));
        const path = pct < .005 ? "" : `M${cx} ${cy}L${x1} ${y1}A${r} ${r} 0 ${pct > .5 ? 1 : 0} 1 ${x2} ${y2}Z`;
        return path ? <path key={i} d={path} fill={seg.color} stroke={C.panel} strokeWidth={1.5} /> : null;
      })}
      <circle cx={size / 2} cy={size / 2} r={size * .22} fill={C.card} />
    </svg>
  );
}

// ── MAIN APP ─────────────────────────────────────────────────
export default function App() {
  const [rawData, setRawData]       = useState(SAMPLE);
  const [thr, setThr]               = useState(PASS_DEFAULT);
  const [page, setPage]             = useState("dashboard");
  const [search, setSearch]         = useState("");
  const [fProdi, setFProdi]         = useState("Semua");
  const [fCat, setFCat]             = useState("Semua");
  const [ewsProdi, setEwsProdi]     = useState("Semua");
  const [sCol, setSCol]             = useState("iaet");
  const [sDir, setSDir]             = useState("desc");
  const [impLog, setImpLog]         = useState(null);
  const [editRow, setEditRow]       = useState(null);
  const [showAdd, setShowAdd]       = useState(false);
  const [newRow, setNewRow]         = useState({ nama:"", username:"", prodi:"Statistika", listening:"", grammar:"", reading:"", total:"", iaet:"", email:"" });
  const [toast, setToast]           = useState(null);
  const fileRef = useRef();

  const notify = (msg, type = "ok") => { setToast({ msg, type }); setTimeout(() => setToast(null), 3500); };

  const data     = useMemo(() => rawData.map(r => enrich(r, thr)), [rawData, thr]);
  const filtered = useMemo(() => {
    let d = [...data];
    if (fProdi !== "Semua") d = d.filter(r => r.prodi === fProdi);
    if (fCat   !== "Semua") d = d.filter(r => r.cat.label === fCat);
    if (search) { const q = search.toLowerCase(); d = d.filter(r => r.nama.toLowerCase().includes(q) || r.username.toLowerCase().includes(q) || r.prodi.toLowerCase().includes(q)); }
    return [...d].sort((a, b) => {
      let av = a[sCol], bv = b[sCol];
      if (typeof av === "string") { av = av.toLowerCase(); bv = bv.toLowerCase(); }
      return sDir === "asc" ? (av > bv ? 1 : -1) : (av < bv ? 1 : -1);
    });
  }, [data, fProdi, fCat, search, sCol, sDir]);

  // dashData: filtered data yang sama dipakai Dashboard & Tabel
  // Filter dari halaman Data (fProdi, fCat, search) ikut mempengaruhi Dashboard
  const dashData = useMemo(() => {
    let d = [...data];
    if (fProdi !== "Semua") d = d.filter(r => r.prodi === fProdi);
    if (fCat   !== "Semua") d = d.filter(r => r.cat.label === fCat);
    if (search) {
      const q = search.toLowerCase();
      d = d.filter(r => r.nama.toLowerCase().includes(q) || r.username.toLowerCase().includes(q) || r.prodi.toLowerCase().includes(q));
    }
    return d;
  }, [data, fProdi, fCat, search]);

  const isFiltered = fProdi !== "Semua" || fCat !== "Semua" || search !== "";

  // KPIs — semua berbasis dashData (ikut filter)
  const N      = dashData.length;
  const lulus  = dashData.filter(r => r.passed).length;
  const belum  = N - lulus;
  const rate   = N ? Math.round(lulus / N * 100) : 0;
  const avgIAET = N ? Math.round(dashData.reduce((a, b) => a + b.iaet, 0) / N) : 0;
  const alerts  = dashData.filter(r => r.iaet < 400).length;

  const catCounts = ["KRITIS","RENDAH","SEDANG","LULUS","UNGGUL"].map(lbl => ({
    label: lbl, value: dashData.filter(r => r.cat.label === lbl).length,
    color: getCat({ KRITIS:300, RENDAH:370, SEDANG:430, LULUS:470, UNGGUL:510 }[lbl]).color,
  }));
  const prodiStats = PRODI_LIST.map(p => {
    const pd = dashData.filter(r => r.prodi === p.nama);
    return { ...p, total: pd.length, lulus: pd.filter(r => r.passed).length,
      avg: pd.length ? Math.round(pd.reduce((a, b) => a + b.iaet, 0) / pd.length) : 0 };
  });
  const compAvg = {
    l: N ? (dashData.reduce((a, b) => a + b.listening, 0) / N).toFixed(1) : 0,
    g: N ? (dashData.reduce((a, b) => a + b.grammar,   0) / N).toFixed(1) : 0,
    r: N ? (dashData.reduce((a, b) => a + b.reading,   0) / N).toFixed(1) : 0,
  };

  // Import
  const handleImport = useCallback((e) => {
    const file = e.target.files?.[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb   = XLSX.read(ev.target.result, { type: "binary" });
        const ws   = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
        if (!json.length) { notify("File kosong", "err"); return; }
        const acc = [], rej = [];
        json.forEach((row, i) => {
          const res = validate(row);
          if (res.ok) acc.push(res.data);
          else rej.push({ row: i + 2, reason: res.reason });
        });
        setImpLog({ total: json.length, accepted: acc.length, rejected: rej, file: file.name });
        if (!acc.length) { notify(`0 baris diterima — ${rej.length} ditolak`, "err"); return; }
        setRawData(acc);
        notify(`✓ ${acc.length} baris diimport${rej.length ? ` · ${rej.length} ditolak` : ""}`);
      } catch { notify("Gagal membaca file", "err"); }
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  }, []);

  // Export
  const handleExport = (type = "xlsx") => {
    const rows = filtered.map(r => ({
      "Nama Mahasiswa": r.nama, "Username": r.username, "Prodi": r.prodi,
      "Listening": r.listening, "Grammar": r.grammar, "Reading": r.reading,
      "Total": r.total, "Result (IAET)": r.iaet, "Predicate": r.predicate,
      "Passed": r.passed ? "LULUS" : "BELUM LULUS", "Kategori EWS": r.cat.label, "Email": r.email,
      "Pesan EWS": r.iaet < 400 ? "WAJIB REMEDIAL" : r.iaet < thr ? "PERLU EAP" : "Lulus",
    }));
    const ws2 = XLSX.utils.json_to_sheet(rows);
    ws2["!cols"] = [22,16,20,10,10,10,8,12,14,14,12,30,18].map(w => ({ wch: w }));
    const wb2 = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb2, ws2, "Data EWS");
    XLSX.writeFile(wb2, `EWS_GREAT_FSTI_${new Date().toISOString().slice(0, 10)}.${type === "csv" ? "csv" : "xlsx"}`);
    notify("✓ File diekspor");
  };

  const handleExportTemplate = () => {
    const tmpl = [{ "Nama Mahasiswa":"Contoh Mahasiswa", "Username":"c.mahasiswa", "Prodi":"Informatika", "Listening":45, "Grammar":35, "Reading":40, "Total":120, "Result (IAET)":548, "Email":"" }];
    const ws2 = XLSX.utils.json_to_sheet(tmpl); const wb2 = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb2, ws2, "Template");
    XLSX.writeFile(wb2, "Template_Import_EWS_GREAT_FSTI.xlsx"); notify("✓ Template diunduh");
  };

  // Sort helper
  const doSort = (col) => { if (sCol === col) setSDir(d => d === "asc" ? "desc" : "asc"); else { setSCol(col); setSDir("desc"); } };
  const SI = ({ col }) => <span style={{ opacity: .35, fontSize: 11, marginLeft: 2 }}>{sCol === col ? (sDir === "asc" ? "↑" : "↓") : "↕"}</span>;

  // Styles
  const st = {
    card:    { background: C.card, borderRadius: 10, border: `1px solid ${C.border}`, padding: 14 },
    cardH:   { fontSize: 12, fontWeight: 700, color: C.accent, letterSpacing: .8, textTransform: "uppercase", marginBottom: 11 },
    th:      { background: "#1a2550", color: C.accent, padding: "10px 11px", textAlign: "left", fontSize: 12, fontWeight: 700, letterSpacing: .5, textTransform: "uppercase", cursor: "pointer", whiteSpace: "nowrap" },
    td:      { padding: "9px 11px", borderBottom: `1px solid ${C.border}`, verticalAlign: "middle" },
    inp:     { background: "#0f172a", border: `1px solid ${C.border}`, borderRadius: 7, color: C.text, padding: "7px 10px", fontSize: 12, outline: "none", width: "100%" },
    sel:     { background: "#0f172a", border: `1px solid ${C.border}`, borderRadius: 7, color: C.text, padding: "7px 10px", fontSize: 12, outline: "none", cursor: "pointer", width: "100%" },
    btnP:    { padding: "6px 14px", borderRadius: 7, border: "none", cursor: "pointer", fontSize: 12, fontWeight: 700, background: C.accent, color: "#0b1120" },
    btnG:    { padding: "6px 14px", borderRadius: 7, border: `1px solid ${C.border}`, cursor: "pointer", fontSize: 12, fontWeight: 700, background: "transparent", color: C.text },
    btnD:    { padding: "6px 14px", borderRadius: 7, border: "1px solid #991b1b", cursor: "pointer", fontSize: 12, fontWeight: 700, background: "transparent", color: "#f87171" },
  };

  const NAVS = [
    { id: "dashboard", lbl: "📊 Dashboard" },
    { id: "data",      lbl: "📋 Data" },
    { id: "ews",       lbl: "🔔 EWS Alert" },
    { id: "import",    lbl: "⬆ Import / Export" },
    { id: "settings",  lbl: "⚙ Settings" },
  ];

  return (
    <div style={{ position:"fixed", inset:0, display:"flex", flexDirection:"column",
      background: C.bg, color: C.text, fontFamily: "'IBM Plex Sans','Segoe UI',sans-serif", overflow: "hidden" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;600&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        ::-webkit-scrollbar{width:5px;height:5px}
        ::-webkit-scrollbar-track{background:#0b1120}
        ::-webkit-scrollbar-thumb{background:#2d3f5e;border-radius:3px}
        .rh:hover{background:#1a2744!important}
        .bh:hover{opacity:.8;cursor:pointer}
        @keyframes fi{from{opacity:0;transform:translateY(4px)}to{opacity:1;transform:none}}
        @keyframes nt{0%{opacity:0;transform:translateY(-8px)}10%,85%{opacity:1;transform:none}100%{opacity:0}}
        button{cursor:pointer;transition:opacity .15s}
        button:hover{opacity:.8}
      `}</style>

      {/* TOAST */}
      {toast && (
        <div style={{ position:"fixed", top:14, right:16, zIndex:9999, padding:"9px 16px", borderRadius:8,
          background: toast.type === "err" ? "#7f1d1d" : "#052e16",
          border: `1px solid ${toast.type === "err" ? "#991b1b" : "#166534"}`,
          color: toast.type === "err" ? "#fca5a5" : "#4ade80",
          fontSize:15, fontWeight:700, animation:"nt 3.5s forwards", boxShadow:"0 8px 24px rgba(0,0,0,.5)" }}>
          {toast.msg}
        </div>
      )}

      {/* TOPBAR */}
      <div style={{ height:52, flexShrink:0, background:C.panel, borderBottom:`1px solid ${C.border}`,
        display:"flex", alignItems:"center", justifyContent:"space-between", padding:"0 18px", gap:8 }}>
        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
          <div style={{ width:30, height:30, borderRadius:8, background:C.accent, display:"flex",
            alignItems:"center", justifyContent:"center", fontSize:15, flexShrink:0, color:"#0b1120", fontWeight:900 }}>EW</div>
          <div style={{ lineHeight:1.25 }}>
            <div style={{ fontSize:15, fontWeight:700 }}>EWS GREAT · FSTI ITK</div>
            <div style={{ fontSize:13, color:C.muted }}>Early Warning System — Latsar CPNS 2026</div>
          </div>
          <div style={{ width:1, height:22, background:C.border, margin:"0 4px" }} />
          {NAVS.map(n => (
            <button key={n.id} onClick={() => setPage(n.id)}
              style={{ padding:"5px 12px", borderRadius:7, border:`1px solid ${page===n.id ? C.accent : C.border}`,
                fontSize:14, fontWeight:700, background: page===n.id ? "#1a2550" : "transparent",
                color: page===n.id ? C.accent : C.text }}>
              {n.lbl}
            </button>
          ))}
        </div>
        <div style={{ display:"flex", gap:8, alignItems:"center" }}>
          {alerts > 0 && (
            <div style={{ background:"#1c0000", border:"1px solid #991b1b", borderRadius:20,
              padding:"3px 12px", fontSize:14, fontWeight:700, color:"#f87171" }}>
              🚨 {alerts} Butuh Intervensi
            </div>
          )}
          <button style={st.btnG} onClick={() => handleExport("xlsx")}>⬇ XLSX</button>
          <button style={st.btnG} onClick={() => handleExport("csv")}>⬇ CSV</button>
        </div>
      </div>

      {/* BODY */}
      <div style={{ flex:1, overflowY:"auto", padding:"15px 18px" }}>

        {/* ════ DASHBOARD ════ */}
        {page === "dashboard" && (
          <div style={{ animation:"fi .3s ease" }}>
            {/* Filter active badge */}
            {isFiltered && (
              <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:12,
                background:"#1a2550", border:`1px solid ${C.accent}`, borderRadius:9, padding:"9px 14px" }}>
                <span style={{ fontSize:14, color:C.accent, fontWeight:700 }}>📊 Dashboard menampilkan data terfilter</span>
                <div style={{ display:"flex", gap:7, flexWrap:"wrap" }}>
                  {fProdi !== "Semua" && <span style={{ background:PRODI_MAP[fProdi]?.bg||"#1e293b", color:PRODI_MAP[fProdi]?.color||C.accent, border:`1px solid ${PRODI_MAP[fProdi]?.color||C.border}`, borderRadius:20, padding:"2px 10px", fontSize:13, fontWeight:700 }}>{fProdi}</span>}
                  {fCat   !== "Semua" && <span style={{ background:"#1e293b", color:C.accent, border:`1px solid ${C.border}`, borderRadius:20, padding:"2px 10px", fontSize:13, fontWeight:700 }}>{fCat}</span>}
                  {search !== ""      && <span style={{ background:"#1e293b", color:C.accent, border:`1px solid ${C.border}`, borderRadius:20, padding:"2px 10px", fontSize:13, fontWeight:700 }}>"{search}"</span>}
                </div>
                <button onClick={() => { setFProdi("Semua"); setFCat("Semua"); setSearch(""); }}
                  style={{ marginLeft:"auto", padding:"4px 12px", borderRadius:6, border:`1px solid ${C.accent}`, background:"transparent", color:C.accent, fontSize:13, fontWeight:700, cursor:"pointer" }}>
                  ✕ Reset Filter
                </button>
              </div>
            )}
            {/* KPI */}
            <div style={{ display:"grid", gridTemplateColumns:"repeat(6,1fr)", gap:10, marginBottom:14 }}>
              {[
                { lbl:"Total",       val:N,          sub:"mahasiswa",       col:C.blue   },
                { lbl:"Lulus",       val:lulus,       sub:`IAET ≥ ${thr}`,  col:C.green  },
                { lbl:"Belum",       val:belum,       sub:`IAET < ${thr}`,  col:C.orange },
                { lbl:"Pass Rate",   val:`${rate}%`,  sub:"keseluruhan",    col:rate>=60?C.green:C.red },
                { lbl:"Avg IAET",    val:avgIAET,     sub:"dari data asli", col:C.accent },
                { lbl:"Intervensi",  val:alerts,      sub:"IAET < 400",     col:C.red    },
              ].map((k, i) => (
                <div key={i} style={{ background:C.card, borderRadius:10, border:`1px solid ${C.border}`, padding:"13px 15px" }}>
                  <div style={{ fontSize:30, fontWeight:700, color:k.col, fontFamily:"'IBM Plex Mono',monospace", lineHeight:1 }}>{k.val}</div>
                  <div style={{ fontSize:15, color:C.text, textTransform:"uppercase", letterSpacing:.4, marginTop:6, fontWeight:700 }}>{k.lbl}</div>
                  <div style={{ fontSize:13, color:C.muted, marginTop:2 }}>{k.sub}</div>
                </div>
              ))}
            </div>

            {/* Row 2 — donut besar | komponen | bar chart full */}
            <div style={{ display:"grid", gridTemplateColumns:"320px 1fr 2fr", gap:12, marginBottom:12 }}>

              {/* Donut — centered, besar */}
              <div style={{ ...st.card, display:"flex", flexDirection:"column" }}>
                <div style={st.cardH}>Distribusi Kategori</div>
                <div style={{ flex:1, display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", gap:16, paddingTop:6 }}>
                  <DonutChart segments={catCounts.filter(c => c.value > 0)} size={160} />
                  <div style={{ width:"100%" }}>
                    {catCounts.map(c => (
                      <div key={c.label} style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:8 }}>
                        <div style={{ display:"flex", alignItems:"center", gap:7 }}>
                          <div style={{ width:10, height:10, borderRadius:3, background:c.color, flexShrink:0 }} />
                          <span style={{ fontSize:14, color:C.muted, fontWeight:600 }}>{c.label}</span>
                        </div>
                        <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                          <div style={{ width:52, height:5, background:"#0f172a", borderRadius:3, overflow:"hidden" }}>
                            <div style={{ height:"100%", width:`${(c.value/(N||1))*100}%`, background:c.color }} />
                          </div>
                          <span style={{ fontSize:14, fontWeight:700, color:c.color, fontFamily:"'IBM Plex Mono',monospace", minWidth:20, textAlign:"right" }}>{c.value}</span>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              {/* Komponen rata-rata */}
              <div style={st.card}>
                <div style={st.cardH}>Rata-rata Komponen</div>
                {[["Listening",compAvg.l,68,"#60a5fa"],["Grammar",compAvg.g,50,"#a78bfa"],["Reading",compAvg.r,67,"#34d399"]].map(([lbl,val,mx,col]) => (
                  <div key={lbl} style={{ marginBottom:16 }}>
                    <div style={{ display:"flex", justifyContent:"space-between", marginBottom:6 }}>
                      <span style={{ fontSize:14, color:C.muted }}>{lbl}</span>
                      <span style={{ fontSize:15, fontWeight:700, color:col, fontFamily:"'IBM Plex Mono',monospace" }}>{val}<span style={{ fontSize:13, color:C.muted }}>/{mx}</span></span>
                    </div>
                    <HBar pct={(val / mx) * 100} color={col} h={10} />
                  </div>
                ))}
                <div style={{ padding:"10px 12px", background:"#0f172a", borderRadius:7, display:"flex", justifyContent:"space-between", alignItems:"center", marginTop:8 }}>
                  <span style={{ fontSize:14, color:C.muted }}>Avg IAET</span>
                  <span style={{ fontSize:26, fontWeight:700, color:avgIAET>=thr?C.green:C.red, fontFamily:"'IBM Plex Mono',monospace" }}>{avgIAET}</span>
                </div>
              </div>

              {/* Bar chart — pass rate per prodi, bar tebal, full space */}
              <div style={st.card}>
                <div style={st.cardH}>Pass Rate per Program Studi</div>
                <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
                  {prodiStats.map(p => {
                    const pr = p.total ? Math.round(p.lulus / p.total * 100) : 0;
                    return (
                      <div key={p.nama}>
                        <div style={{ display:"flex", justifyContent:"space-between", marginBottom:5 }}>
                          <span style={{ fontSize:14, color:p.color, fontWeight:700 }}>{p.short}</span>
                          <span style={{ fontSize:14, fontFamily:"'IBM Plex Mono',monospace", color:C.muted }}>
                            <span style={{ color:p.color, fontWeight:700 }}>{p.lulus}</span>/{p.total}
                            <span style={{ marginLeft:6, color: pr>=60?C.green:C.orange, fontWeight:700 }}>{pr}%</span>
                          </span>
                        </div>
                        <HBar pct={pr} color={p.color} h={12} />
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>

            {/* Row 3 */}
            <div style={{ display:"grid", gridTemplateColumns:"1.5fr 1fr 1fr", gap:12 }}>
              {/* Statistik prodi */}
              <div style={st.card}>
                <div style={st.cardH}>Statistik 8 Program Studi</div>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:14 }}>
                  <thead>
                    <tr>
                      {["Program Studi","N","Lulus","Avg IAET","Pass%"].map(h => (
                        <th key={h} style={{ ...st.th, padding:"8px 10px" }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {prodiStats.map((p, i) => (
                      <tr key={i} className="rh" style={{ borderBottom:`1px solid ${C.border}` }}>
                        <td style={st.td}>
                          <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                            <div style={{ width:8, height:8, borderRadius:2, background:p.color, flexShrink:0 }} />
                            <span style={{ fontWeight:700, fontSize:14 }}>{p.short}</span>
                          </div>
                        </td>
                        <td style={{ ...st.td, textAlign:"center", fontFamily:"'IBM Plex Mono',monospace" }}>{p.total}</td>
                        <td style={{ ...st.td, textAlign:"center", color:C.green, fontFamily:"'IBM Plex Mono',monospace", fontWeight:700 }}>{p.lulus}</td>
                        <td style={{ ...st.td, textAlign:"center", color:C.accent, fontFamily:"'IBM Plex Mono',monospace", fontWeight:700 }}>{p.avg}</td>
                        <td style={{ ...st.td, textAlign:"center", fontFamily:"'IBM Plex Mono',monospace", fontWeight:700,
                          color: p.total && Math.round(p.lulus/p.total*100) >= 60 ? C.green : C.orange }}>
                          {p.total ? `${Math.round(p.lulus/p.total*100)}%` : "0%"}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              {/* Top & Bottom */}
              {[
                ["🏆 Top 5 Tertinggi", [...dashData].sort((a,b) => b.iaet-a.iaet).slice(0,5)],
                ["⚠ 5 Terendah",       [...dashData].sort((a,b) => a.iaet-b.iaet).slice(0,5)],
              ].map(([title, rows]) => (
                <div key={title} style={st.card}>
                  <div style={st.cardH}>{title}</div>
                  {rows.map((r, i) => {
                    const pc = PRODI_MAP[r.prodi] || PRODI_LIST[0];
                    return (
                      <div key={i} style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"6px 0", borderBottom:`1px solid ${C.border}` }}>
                        <div>
                          <div style={{ fontSize:14, fontWeight:700 }}>{r.nama}</div>
                          <div style={{ fontSize:13, color:pc.color, fontWeight:600 }}>{pc.short}</div>
                        </div>
                        <div style={{ display:"flex", alignItems:"center", gap:7 }}>
                          <Pill {...r.cat} />
                          <span style={{ fontFamily:"'IBM Plex Mono',monospace", fontWeight:700, color:r.cat.color, fontSize:17 }}>{r.iaet}</span>
                        </div>
                      </div>
                    );
                  })}
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ════ DATA ════ */}
        {page === "data" && (
          <div style={{ animation:"fi .3s ease" }}>
            {/* Filter bar */}
            <div style={{ display:"flex", gap:9, marginBottom:11, flexWrap:"wrap", alignItems:"center",
              background:C.card, padding:"10px 13px", borderRadius:10, border:`1px solid ${C.border}` }}>
              <input style={{ ...st.inp, width:220 }} placeholder="🔍 Cari nama / username / prodi..."
                value={search} onChange={e => setSearch(e.target.value)} />
              <select style={{ ...st.sel, width:175 }} value={fProdi} onChange={e => setFProdi(e.target.value)}>
                <option>Semua</option>
                {PRODI_NAMES.map(p => <option key={p}>{p}</option>)}
              </select>
              <select style={{ ...st.sel, width:130 }} value={fCat} onChange={e => setFCat(e.target.value)}>
                {["Semua","UNGGUL","LULUS","SEDANG","RENDAH","KRITIS"].map(c => <option key={c}>{c}</option>)}
              </select>
              <span style={{ fontSize:14, color:C.muted }}>{filtered.length}/{N}</span>
              <div style={{ marginLeft:"auto", display:"flex", gap:8 }}>
                <button style={st.btnP} onClick={() => setShowAdd(v => !v)}>+ Tambah</button>
                <button style={st.btnG} onClick={() => handleExport("xlsx")}>⬇ Export</button>
              </div>
            </div>

            {/* Add form */}
            {showAdd && (
              <div style={{ ...st.card, border:`1px solid ${C.blue}`, marginBottom:11 }}>
                <div style={{ fontSize:16, fontWeight:700, color:C.accent, marginBottom:10 }}>+ Tambah Mahasiswa Baru</div>
                <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))", gap:9, marginBottom:10 }}>
                  {[{k:"nama",lbl:"Nama ★",t:"text"},{k:"username",lbl:"Username ★",t:"text"},
                    {k:"listening",lbl:"Listening ★",t:"number"},{k:"grammar",lbl:"Grammar ★",t:"number"},
                    {k:"reading",lbl:"Reading ★",t:"number"},{k:"total",lbl:"Total Skor ★",t:"number"},
                    {k:"iaet",lbl:"Result (IAET) ★",t:"number"},{k:"email",lbl:"Email",t:"text"}].map(({k,lbl,t}) => (
                    <div key={k}>
                      <div style={{ fontSize:13, color:C.muted, marginBottom:3 }}>{lbl}</div>
                      <input style={st.inp} type={t} value={newRow[k]} onChange={e => setNewRow(p => ({...p,[k]:e.target.value}))} />
                    </div>
                  ))}
                  <div>
                    <div style={{ fontSize:13, color:C.muted, marginBottom:3 }}>Prodi ★</div>
                    <select style={st.sel} value={newRow.prodi} onChange={e => setNewRow(p => ({...p,prodi:e.target.value}))}>
                      {PRODI_NAMES.map(p => <option key={p}>{p}</option>)}
                    </select>
                  </div>
                </div>
                <div style={{ fontSize:13, color:C.muted, marginBottom:9 }}>★ 8 kolom wajib diisi. Total & IAET harus dari SIAKAD — tidak ada kalkulasi otomatis.</div>
                <div style={{ display:"flex", gap:8 }}>
                  <button style={st.btnP} onClick={() => {
                    const miss = REQUIRED.filter(k => !newRow[k] && newRow[k] !== 0);
                    if (miss.length) { notify(`Wajib: ${miss.map(k=>REQ_LBL[k]).join(", ")}`, "err"); return; }
                    const noNum = ["listening","grammar","reading","total","iaet"].filter(c => isNaN(Number(newRow[c])));
                    if (noNum.length) { notify(`Harus angka: ${noNum.join(", ")}`, "err"); return; }
                    setRawData(p => [...p, newRow]);
                    setNewRow({nama:"",username:"",prodi:"Statistika",listening:"",grammar:"",reading:"",total:"",iaet:"",email:""});
                    setShowAdd(false); notify("✓ Ditambahkan");
                  }}>Simpan</button>
                  <button style={st.btnG} onClick={() => setShowAdd(false)}>Batal</button>
                </div>
              </div>
            )}

            {/* Table */}
            <div style={{ background:C.card, borderRadius:10, border:`1px solid ${C.border}`, overflow:"hidden" }}>
              <div style={{ overflowX:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse", fontSize:14 }}>
                  <thead>
                    <tr>
                      <th style={st.th}>#</th>
                      {[["nama","Nama Mahasiswa"],["username","Username"],["prodi","Prodi"],
                        ["listening","Listening"],["grammar","Grammar"],["reading","Reading"],
                        ["total","Total"],["iaet","Result (IAET)"],["predicate","Predicate"],
                        ["passed","Passed"],["email","Email"]].map(([col,lbl]) => (
                        <th key={col} style={st.th} onClick={() => doSort(col)}>{lbl}<SI col={col}/></th>
                      ))}
                      <th style={st.th}>Aksi</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filtered.map((r, i) => {
                      const rawIdx = rawData.findIndex(d => d.nama === r.nama && d.username === r.username);
                      const isEd   = editRow?.index === rawIdx;
                      const pc     = PRODI_MAP[r.prodi] || PRODI_LIST[0];
                      const rowBg  = i % 2 === 0 ? C.card : "#172032";
                      return (
                        <tr key={i} className="rh" style={{ background:rowBg }}>
                          <td style={{ ...st.td, color:C.muted, fontFamily:"'IBM Plex Mono',monospace" }}>{i+1}</td>
                          <td style={st.td}>
                            {isEd
                              ? <input style={{ ...st.inp, width:140 }} value={editRow.data.nama||""} onChange={e => setEditRow(p=>({...p,data:{...p.data,nama:e.target.value}}))}/>
                              : <span style={{ fontWeight:700, fontSize:14 }}>{r.nama}</span>}
                          </td>
                          <td style={st.td}><span style={{ fontSize:13, color:C.muted, fontFamily:"'IBM Plex Mono',monospace" }}>{r.username}</span></td>
                          <td style={st.td}>
                            {isEd
                              ? <select style={{ ...st.sel, width:140 }} value={editRow.data.prodi||r.prodi} onChange={e=>setEditRow(p=>({...p,data:{...p.data,prodi:e.target.value}}))}>
                                  {PRODI_NAMES.map(p=><option key={p}>{p}</option>)}
                                </select>
                              : <span style={{ fontSize:14, fontWeight:700, color:pc.color }}>{pc.short}</span>}
                          </td>
                          {[["listening",68],["grammar",50],["reading",67]].map(([k,mx]) => {
                            const pct = ((r[k]/mx)*100);
                            const fg  = pct>=70?C.green:pct>=50?C.accent:C.red;
                            return (
                              <td key={k} style={{ ...st.td, minWidth:62 }}>
                                {isEd
                                  ? <input style={{ ...st.inp, width:55 }} type="number" value={editRow.data[k]??r[k]} onChange={e=>setEditRow(p=>({...p,data:{...p.data,[k]:e.target.value}}))}/>
                                  : <div>
                                      <span style={{ fontFamily:"'IBM Plex Mono',monospace", fontWeight:700, fontSize:15, color:fg }}>{r[k]}</span>
                                      <HBar pct={pct} color={fg} h={3}/>
                                    </div>}
                              </td>
                            );
                          })}
                          <td style={{ ...st.td, minWidth:65 }}>
                            {isEd
                              ? <input style={{ ...st.inp, width:58 }} type="number" value={editRow.data.total??r.total} onChange={e=>setEditRow(p=>({...p,data:{...p.data,total:e.target.value}}))}/>
                              : <div>
                                  <span style={{ fontFamily:"'IBM Plex Mono',monospace", fontWeight:700, fontSize:15 }}>{r.total}</span>
                                  <div style={{ fontSize:13, color:C.muted }}>/185</div>
                                </div>}
                          </td>
                          <td style={{ ...st.td, minWidth:80 }}>
                            {isEd
                              ? <input style={{ ...st.inp, width:70 }} type="number" value={editRow.data.iaet??r.iaet} onChange={e=>setEditRow(p=>({...p,data:{...p.data,iaet:e.target.value}}))}/>
                              : <div>
                                  <span style={{ fontFamily:"'IBM Plex Mono',monospace", fontWeight:700, color:r.cat.color, fontSize:17 }}>{r.iaet}</span>
                                  <HBar pct={Math.min((r.iaet-310)/(677-310)*100,100)} color={r.cat.color} h={3}/>
                                </div>}
                          </td>
                          <td style={st.td}><span style={{ fontSize:14, color:C.muted }}>{r.predicate}</span></td>
                          <td style={st.td}><Pill label={r.passed?"✓ LULUS":"✗ BELUM"} color={r.passed?"#4ade80":"#f87171"} bg={r.passed?"#052e16":"#1c0000"} border={r.passed?"#166534":"#991b1b"}/></td>
                          <td style={st.td}><span style={{ fontSize:13, color:"#818cf8" }}>{r.email}</span></td>
                          <td style={st.td}>
                            {isEd
                              ? <div style={{ display:"flex", gap:4 }}>
                                  <button style={{ ...st.btnP, padding:"3px 9px", fontSize:13 }} onClick={() => {
                                    const miss = REQUIRED.filter(k => editRow.data[k]==null || String(editRow.data[k]).trim()==="");
                                    if (miss.length) { notify(`Wajib: ${miss.map(k=>REQ_LBL[k]).join(", ")}`, "err"); return; }
                                    setRawData(prev => prev.map((r,idx) => idx===rawIdx ? {...r,...editRow.data} : r));
                                    setEditRow(null); notify("✓ Diperbarui");
                                  }}>✓</button>
                                  <button style={{ ...st.btnG, padding:"3px 9px", fontSize:13 }} onClick={() => setEditRow(null)}>✕</button>
                                </div>
                              : <div style={{ display:"flex", gap:4 }}>
                                  <button style={{ ...st.btnG, padding:"3px 9px", fontSize:13 }} onClick={() => setEditRow({index:rawIdx,data:{...rawData[rawIdx]}})}>✏</button>
                                  <button style={{ ...st.btnD, padding:"3px 9px", fontSize:13 }} onClick={() => { if(confirm(`Hapus ${rawData[rawIdx].nama}?`)) { setRawData(p=>p.filter((_,idx)=>idx!==rawIdx)); notify("✓ Dihapus"); }}}>🗑</button>
                                </div>}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              {!filtered.length && <div style={{ textAlign:"center", padding:36, color:C.muted, fontSize:15 }}>Tidak ada data sesuai filter</div>}
            </div>
          </div>
        )}

        {/* ════ EWS ALERT ════ */}
        {page === "ews" && (
          <div style={{ animation:"fi .3s ease" }}>
            {/* Filter bar EWS */}
            <div style={{ display:"flex", gap:10, marginBottom:12, alignItems:"center", flexWrap:"wrap",
              background:C.card, padding:"10px 14px", borderRadius:10, border:`1px solid ${C.border}` }}>
              <div style={{ fontSize:14, fontWeight:700, color:C.accent, marginRight:4 }}>Filter:</div>
              <select style={{ ...st.sel, width:200 }} value={ewsProdi} onChange={e => setEwsProdi(e.target.value)}>
                <option value="Semua">Semua Program Studi</option>
                {PRODI_NAMES.map(p => <option key={p} value={p}>{p}</option>)}
              </select>
              {ewsProdi !== "Semua" && (
                <button style={{ ...st.btnG, padding:"5px 12px", fontSize:13 }} onClick={() => setEwsProdi("Semua")}>✕ Reset</button>
              )}
              <div style={{ marginLeft:"auto", fontSize:14, color:C.muted }}>
                Menampilkan: <span style={{ color:C.accent, fontWeight:700 }}>
                  {data.filter(r => r.iaet < thr && (ewsProdi === "Semua" || r.prodi === ewsProdi)).length}
                </span> mahasiswa belum lulus
                {ewsProdi !== "Semua" && <span style={{ color:PRODI_MAP[ewsProdi]?.color, fontWeight:700 }}> · {ewsProdi}</span>}
              </div>
            </div>

            {/* Prodi pills shortcut */}
            <div style={{ display:"flex", gap:7, marginBottom:12, flexWrap:"wrap" }}>
              {PRODI_LIST.map(p => {
                const cnt = data.filter(r => r.iaet < thr && r.prodi === p.nama).length;
                if (!cnt) return null;
                return (
                  <button key={p.nama} onClick={() => setEwsProdi(ewsProdi === p.nama ? "Semua" : p.nama)}
                    style={{ padding:"4px 12px", borderRadius:20, border:`1px solid ${ewsProdi===p.nama?p.color:C.border}`,
                      background: ewsProdi===p.nama ? p.bg : "transparent",
                      color: ewsProdi===p.nama ? p.color : C.muted,
                      fontSize:13, fontWeight:700, cursor:"pointer" }}>
                    {p.short} <span style={{ fontFamily:"monospace" }}>({cnt})</span>
                  </button>
                );
              })}
            </div>

            <div style={{ background:"#1c0404", border:"1px solid #991b1b", borderRadius:10, padding:"11px 15px", marginBottom:12 }}>
              <div style={{ fontSize:15, fontWeight:700, color:"#fca5a5" }}>🚨 {data.filter(r => r.iaet < 400 && (ewsProdi==="Semua"||r.prodi===ewsProdi)).length} Mahasiswa Perlu Intervensi Segera (IAET &lt; 400){ewsProdi!=="Semua"?` — ${ewsProdi}`:""}</div>
              <div style={{ fontSize:14, color:"#f87171", marginTop:3 }}>Skor IAET diambil langsung dari data — tidak ada kalkulasi. Threshold lulus: {thr}.</div>
            </div>
            {[
              { lbl:`🔴 KRITIS (IAET < 350)`,          f:r=>r.iaet<350,              a:"#b91c1c", bg:"#1c0000", cat:"KRITIS", act:"Kirim Warning Kritis"   },
              { lbl:`🔴 RENDAH (IAET 350–399)`,         f:r=>r.iaet>=350&&r.iaet<400, a:"#c62828", bg:"#1c0404", cat:"RENDAH", act:"Kirim Warning Remedial" },
              { lbl:`🟡 SEDANG (IAET 400–${thr-1})`,    f:r=>r.iaet>=400&&r.iaet<thr, a:"#ea580c", bg:"#1c1000", cat:"SEDANG", act:"Kirim EAP Notice"       },
            ].map(({lbl,f,a,bg,cat,act}) => {
              const mhs = data.filter(r => f(r) && (ewsProdi === "Semua" || r.prodi === ewsProdi)).sort((x,y) => x.iaet-y.iaet);
              if (!mhs.length) return null;
              return (
                <div key={lbl} style={{ background:C.card, borderRadius:10, border:`1px solid ${C.border}`, marginBottom:12, overflow:"hidden" }}>
                  <div style={{ background:a, padding:"9px 15px", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                    <span style={{ fontSize:15, fontWeight:700, color:"#fff" }}>{lbl} — {mhs.length} mahasiswa</span>
                    <button style={{ background:"rgba(255,255,255,.15)", color:"#fff", border:"none", padding:"5px 12px", borderRadius:6, fontSize:14, fontWeight:700, cursor:"pointer" }}
                      onClick={() => { openBulkMailto(mhs, cat, thr); notify(`📧 Draft dibuka untuk ${mhs.length} mahasiswa`); }}>
                      📧 {act}
                    </button>
                  </div>
                  {mhs.map((r, i) => {
                    const pc = PRODI_MAP[r.prodi] || PRODI_LIST[0];
                    return (
                      <div key={i} className="rh" style={{ display:"grid", gridTemplateColumns:"1fr 58px 58px 58px 68px 80px 82px auto",
                        gap:9, padding:"9px 15px", borderBottom:`1px solid ${C.border}`, background:i%2===0?C.card:"#172032", alignItems:"center" }}>
                        <div>
                          <div style={{ fontWeight:700, fontSize:14 }}>{r.nama}</div>
                          <div style={{ fontSize:13, color:pc.color, fontWeight:600 }}>{pc.short} · {r.email}</div>
                        </div>
                        {[["L",r.listening,"#60a5fa"],["G",r.grammar,"#a78bfa"],["R",r.reading,"#34d399"]].map(([k,v,col]) => (
                          <div key={k} style={{ textAlign:"center", background:"#0f172a", borderRadius:6, padding:"4px 6px" }}>
                            <div style={{ fontSize:13, color:C.muted }}>{k}</div>
                            <div style={{ fontSize:15, fontWeight:700, color:col, fontFamily:"'IBM Plex Mono',monospace" }}>{v}</div>
                          </div>
                        ))}
                        <div style={{ textAlign:"center", background:"#0f172a", borderRadius:6, padding:"4px 6px" }}>
                          <div style={{ fontSize:13, color:C.muted }}>Total</div>
                          <div style={{ fontSize:15, fontWeight:700, color:C.muted, fontFamily:"'IBM Plex Mono',monospace" }}>{r.total}</div>
                        </div>
                        <div style={{ textAlign:"center", background:bg, borderRadius:6, padding:"4px 8px" }}>
                          <div style={{ fontSize:13, color:a }}>IAET</div>
                          <div style={{ fontSize:18, fontWeight:700, color:a, fontFamily:"'IBM Plex Mono',monospace" }}>{r.iaet}</div>
                        </div>
                        <Pill {...r.cat} />
                        <button style={{ ...st.btnG, padding:"4px 9px", fontSize:14 }} onClick={() => {
                          const tmplFn = EMAIL_TEMPLATES[r.cat.label] || EMAIL_TEMPLATES["SEDANG"];
                          const tmpl   = typeof tmplFn === "function" ? tmplFn(r, thr) : tmplFn;
                          openMailto(r.email, tmpl.subject, tmpl.body);
                          notify(`📧 Draft email ke ${r.nama} dibuka`);
                        }}>📧 Email</button>
                      </div>
                    );
                  })}
                </div>
              );
            })}
            <div style={{ background:"#052e16", border:"1px solid #166534", borderRadius:10, padding:"12px 15px" }}>
              <div style={{ fontSize:15, fontWeight:700, color:C.green, marginBottom:9 }}>
                ✅ {data.filter(r => r.passed && (ewsProdi==="Semua"||r.prodi===ewsProdi)).length} Mahasiswa Lulus IAET ≥ {thr}
                {ewsProdi!=="Semua" && <span style={{ fontSize:13, color:PRODI_MAP[ewsProdi]?.color, marginLeft:8 }}>— {ewsProdi}</span>}
              </div>
              <div style={{ display:"flex", flexWrap:"wrap", gap:6 }}>
                {data.filter(r => r.passed && (ewsProdi==="Semua"||r.prodi===ewsProdi)).sort((a,b) => b.iaet-a.iaet).map((r,i) => (
                  <div key={i} style={{ background:"#0f172a", borderRadius:7, padding:"4px 10px", fontSize:14 }}>
                    <span style={{ color:C.text, fontWeight:500 }}>{r.nama}</span>
                    <span style={{ color:r.cat.color, fontWeight:700, marginLeft:6, fontFamily:"'IBM Plex Mono',monospace" }}>{r.iaet}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {/* ════ IMPORT / EXPORT ════ */}
        {page === "import" && (
          <div style={{ animation:"fi .3s ease", maxWidth:740 }}>
            <div style={st.card}>
              <div style={st.cardH}>⬆ Import Data dari SIAKAD / Excel</div>
              <div style={{ background:"#0f172a", border:`2px dashed ${C.border}`, borderRadius:10, padding:28, textAlign:"center", marginBottom:14, cursor:"pointer" }}
                onClick={() => fileRef.current?.click()}
                onDragOver={e => { e.preventDefault(); e.currentTarget.style.borderColor = C.blue; }}
                onDragLeave={e => { e.currentTarget.style.borderColor = C.border; }}
                onDrop={e => { e.preventDefault(); const f = e.dataTransfer.files[0]; if(f) handleImport({target:{files:[f],value:""}}); }}>
                <div style={{ fontSize:30, marginBottom:7 }}>📂</div>
                <div style={{ fontSize:15, fontWeight:700, marginBottom:3 }}>Klik atau drag & drop file</div>
                <div style={{ fontSize:14, color:C.muted }}>Format: .xlsx · .xls · .csv</div>
                <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display:"none" }} onChange={handleImport} />
              </div>
              {impLog && (
                <div style={{ marginBottom:14 }}>
                  <div style={{ background:"#052e16", border:"1px solid #166534", borderRadius:8, padding:"10px 13px", marginBottom:8, fontSize:14, color:C.green }}>
                    ✓ <strong>{impLog.accepted}</strong> dari {impLog.total} baris diterima · <em>{impLog.file}</em>
                  </div>
                  {impLog.rejected.length > 0 && (
                    <div style={{ background:"#1c0404", border:"1px solid #991b1b", borderRadius:8, padding:"10px 13px" }}>
                      <div style={{ fontSize:14, fontWeight:700, color:"#fca5a5", marginBottom:6 }}>✗ {impLog.rejected.length} baris ditolak:</div>
                      <div style={{ maxHeight:110, overflowY:"auto" }}>
                        {impLog.rejected.map((rej,i) => (
                          <div key={i} style={{ fontSize:13, color:"#f87171", padding:"2px 0" }}>Baris {rej.row}: {rej.reason}</div>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              )}
              <div style={{ fontSize:15, fontWeight:700, color:C.accent, marginBottom:9 }}>8 Kolom Wajib Ada di File Import:</div>
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:7, marginBottom:14 }}>
                {[
                  ["Nama Mahasiswa","nama, Nama, Student Name","SIAKAD"],
                  ["Username","username, nim, nrp","SIAKAD"],
                  ["Prodi","prodi, Program Studi, jurusan","SIAKAD"],
                  ["Listening","listening","Skor IAET"],
                  ["Grammar","grammar, Structure","Skor IAET"],
                  ["Reading","reading","Skor IAET"],
                  ["Total","total, Total Skor","⚠ Data asli"],
                  ["Result (IAET)","iaet, result, Score","⚠ Data asli"],
                ].map(([col,alias,src],i) => (
                  <div key={i} style={{ background:"#0f172a", borderRadius:8, padding:"9px 11px", display:"flex", justifyContent:"space-between", alignItems:"flex-start" }}>
                    <div>
                      <div style={{ fontSize:14, fontWeight:700, color:C.accent }}>{col}</div>
                      <div style={{ fontSize:13, color:C.muted, marginTop:2 }}>Alias: {alias}</div>
                    </div>
                    <span style={{ fontSize:13, fontWeight:700, color:src.includes("⚠")?C.orange:C.green, marginLeft:8, whiteSpace:"nowrap" }}>{src}</span>
                  </div>
                ))}
              </div>
              <div style={{ fontSize:14, color:"#f87171", background:"#1c0000", border:"1px solid #991b1b", borderRadius:7, padding:"9px 13px", marginBottom:14 }}>
                ⛔ Jika salah satu dari 8 kolom kosong, baris tersebut akan DITOLAK sepenuhnya. Tidak ada kalkulasi otomatis.
              </div>
              <div style={{ display:"flex", gap:10 }}>
                <button style={st.btnP} onClick={handleExportTemplate}>⬇ Download Template</button>
                <button style={st.btnG} onClick={() => fileRef.current?.click()}>⬆ Pilih File</button>
              </div>
            </div>
            <div style={{ ...st.card, marginTop:12 }}>
              <div style={st.cardH}>⬇ Export Data</div>
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12, marginBottom:10 }}>
                {[
                  { lbl:"Export XLSX", desc:"Semua kolom + status EWS", fn:()=>handleExport("xlsx") },
                  { lbl:"Export CSV",  desc:"Untuk mail merge / sistem lain", fn:()=>handleExport("csv") },
                ].map((opt,i) => (
                  <div key={i} className="bh" style={{ ...st.card, cursor:"pointer" }} onClick={opt.fn}>
                    <div style={{ fontSize:15, fontWeight:700, marginBottom:4 }}>{opt.lbl}</div>
                    <div style={{ fontSize:14, color:C.muted, marginBottom:10 }}>{opt.desc}</div>
                    <button style={st.btnP}>Download</button>
                  </div>
                ))}
              </div>
              <div style={{ fontSize:14, color:C.muted }}>* Menggunakan filter aktif — {filtered.length} baris.</div>
            </div>
          </div>
        )}

        {/* ════ SETTINGS ════ */}
        {page === "settings" && (
          <div style={{ animation:"fi .3s ease", maxWidth:500 }}>
            <div style={st.card}>
              <div style={st.cardH}>⚙ Konfigurasi EWS</div>
              <div style={{ display:"grid", gap:15 }}>
                <div>
                  <div style={{ fontSize:15, fontWeight:700, marginBottom:5 }}>Threshold Lulus (IAET ≥)</div>
                  <input style={{ ...st.inp, width:170 }} type="number" value={thr} onChange={e => setThr(Number(e.target.value)||460)} />
                  <div style={{ fontSize:14, color:C.muted, marginTop:4 }}>Mengubah nilai ini memperbarui Passed, kategori EWS, dan semua chart secara otomatis.</div>
                </div>
                <div style={{ padding:"11px 13px", background:"#0f172a", borderRadius:8 }}>
                  <div style={{ fontSize:14, fontWeight:700, color:C.accent, marginBottom:8 }}>Kategori Saat Ini</div>
                  {[["KRITIS","< 350","#ef4444"],["RENDAH","350–399","#f87171"],
                    ["SEDANG",`400–${thr-1}`,"#fb923c"],["LULUS",`${thr}–499`,"#4ade80"],["UNGGUL","≥ 500","#22c55e"]].map(([lbl,rng,col]) => (
                    <div key={lbl} style={{ display:"flex", justifyContent:"space-between", padding:"6px 0", borderBottom:`1px solid ${C.border}`, fontSize:15 }}>
                      <span style={{ color:col, fontWeight:700 }}>{lbl}</span>
                      <span style={{ color:C.muted, fontFamily:"'IBM Plex Mono',monospace", fontSize:14 }}>IAET {rng}</span>
                    </div>
                  ))}
                </div>
                <div style={{ padding:"11px 13px", background:"#0f172a", borderRadius:8 }}>
                  <div style={{ fontSize:14, fontWeight:700, color:C.accent, marginBottom:8 }}>Kebijakan Data</div>
                  <div style={{ fontSize:14, color:C.muted, lineHeight:1.9 }}>
                    ✅ Kolom <strong style={{ color:C.text }}>Total</strong> dan <strong style={{ color:C.text }}>Result (IAET)</strong> diambil langsung dari file import<br/>
                    ⛔ Tidak ada kalkulasi otomatis — jika kolom kosong, baris ditolak<br/>
                    ℹ Predicate &amp; Passed diderivasi dari nilai IAET yang diimport
                  </div>
                </div>
                <div>
                  <div style={{ fontSize:15, fontWeight:700, marginBottom:5 }}>Reset ke Data Sampel</div>
                  <button style={st.btnD} onClick={() => {
                    if (confirm("Reset ke 32 data sampel?")) { setRawData(SAMPLE); setImpLog(null); notify("✓ Data direset"); }
                  }}>🔄 Reset Data</button>
                </div>
              </div>
            </div>
          </div>
        )}

      </div>
    </div>
  );
}