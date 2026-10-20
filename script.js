// =================================================================
// Konverter Faktur Pajak (Excel DJP) -> XML Jurnal Accurate 5
// Satu halaman untuk 4 jenis faktur: Keluaran, Masukan B1, B2, B3.
// =================================================================
"use strict";

const BRANCH_CODE = "1472498169";
const STORAGE_PREFIX = "fpconv.";

const MONTHS = [
  "Januari", "Februari", "Maret", "April", "Mei", "Juni",
  "Juli", "Agustus", "September", "Oktober", "November", "Desember",
];

// Nama bulan yang dikenali saat membaca Masa Pajak (Indonesia & Inggris)
const MONTH_ALIASES = [
  ["januari", "january"],
  ["februari", "pebruari", "february"],
  ["maret", "march"],
  ["april"],
  ["mei", "may"],
  ["juni", "june"],
  ["juli", "july"],
  ["agustus", "agt", "ags", "august"],
  ["september"],
  ["oktober", "october"],
  ["november", "nopember"],
  ["desember", "december"],
];

// Kolom Masa Pajak (mis. "Januari 2026") menentukan YY.MM pada nomor jurnal.
const PERIOD_COLUMNS = ["Masa Pajak"];
const TAX_YEAR_COLUMNS = ["Tahun Pajak"];

// -----------------------------------------------------------------
// KONFIGURASI JENIS FAKTUR
// `columns` berisi nama-nama kolom Excel yang dicoba berurutan
// (kolom pertama yang ada isinya yang dipakai). Nama pertama adalah
// format lampiran SPT Masa PPN Coretax; sisanya format lama.
// -----------------------------------------------------------------
const MODES = [
  {
    id: "keluaran",
    category: "sales",
    icon: "i-sales",
    title: "Faktur Pajak Keluaran A2",
    tag: "Penjualan",
    desc: "Faktur yang Anda terbitkan saat menjual barang/jasa ke pembeli.",
    prefix: "PENJ",
    party: "Pembeli",
    parseNum: parseFloat,
    descSep: " - ",
    columns: {
      invoiceNo: [
        "Nomor Faktur/Dokumen",
        "Nomor Faktur Pajak",
        "Faktur Pajak/Dokumen Tertentu/Nota Retur/Nota Pembatalan - Nomor",
      ],
      date: [
        "Tanggal Faktur/Dokumen",
        "Tanggal Faktur Pajak",
        "Faktur Pajak/Dokumen Tertentu/Nota Retur/Nota Pembatalan - Tanggal",
      ],
      name: ["Nama Pembeli", "Nama Pembeli BKP/Penerima Manfaat BKP Tidak Berwujud/Penerima JKP"],
      npwp: ["NPWP/NIK Pembeli", "NPWP Pembeli / Identitas lainnya", "NPWP/NIK/Nomor Paspor"],
      dpp: ["Harga Jual / Penggantian / DPP", "Harga Jual/Penggantian/DPP (Rupiah)"],
      ppn: ["PPN", "PPN (Rupiah)"],
    },
    accounts: [
      {
        key: "receivable", label: "Akun Piutang Usaha", side: "debit", amount: "DPP + PPN",
        hint: "Tagihan kepada pembeli (nilai faktur termasuk PPN).",
        suggestions: [["110302", "Piutang Usaha"]],
      },
      {
        key: "sales", label: "Akun Penjualan", side: "credit", amount: "DPP",
        hint: "Pendapatan dari penjualan barang/jasa (sebelum PPN).",
        suggestions: [["4000.03", "Penjualan"]],
      },
      {
        key: "tax", label: "Akun PPN Keluaran", side: "credit", amount: "PPN",
        hint: "PPN yang Anda pungut dari pembeli.",
        suggestions: [["2100.01", "PPN Penjualan"]],
      },
    ],
    lines(r, a, total) {
      return `
            <ACCOUNTLINE operation="Add">
                <KeyID>0</KeyID>
                <GLACCOUNT>${a.receivable}</GLACCOUNT>
                <GLAMOUNT>${total}</GLAMOUNT>
                <CUSTOMERNO>1000</CUSTOMERNO>
                <DESCRIPTION>${r.name} - ${r.invoiceNo}</DESCRIPTION>
                <RATE>1</RATE>
                <PRIMEAMOUNT>${total}</PRIMEAMOUNT>
                <TXDATE/>
                <POSTED/>
                <CURRENCYNAME>IDR</CURRENCYNAME>
            </ACCOUNTLINE>
            <ACCOUNTLINE operation="Add">
                <KeyID>1</KeyID>
                <GLACCOUNT>${a.sales}</GLACCOUNT>
                <GLAMOUNT>${-r.dpp}</GLAMOUNT>
                <DESCRIPTION>${r.name} - ${r.invoiceNo}</DESCRIPTION>
                <RATE>1</RATE>
                <PRIMEAMOUNT>${-r.dpp}</PRIMEAMOUNT>
                <TXDATE/>
                <POSTED/>
                <CURRENCYNAME>IDR</CURRENCYNAME>
            </ACCOUNTLINE>
            <ACCOUNTLINE operation="Add">
                <KeyID>2</KeyID>
                <GLACCOUNT>${a.tax}</GLACCOUNT>
                <GLAMOUNT>${-r.ppn}</GLAMOUNT>
                <DESCRIPTION>PPN 12% - ${r.name}</DESCRIPTION>
                <RATE>1</RATE>
                <PRIMEAMOUNT>${-r.ppn}</PRIMEAMOUNT>
                <TXDATE/>
                <POSTED/>
                <CURRENCYNAME>IDR</CURRENCYNAME>
            </ACCOUNTLINE>`;
    },
  },
  {
    id: "b1",
    category: "purchase_b1",
    icon: "i-ship",
    title: "Faktur Pajak Masukan B1",
    tag: "Impor / PIB",
    desc: "Pembelian dari luar negeri (impor) dengan dokumen PIB.",
    prefix: "PIB",
    party: "Penjual",
    parseNum: parseFloat,
    descSep: "\n ",
    columns: {
      invoiceNo: ["Nomor Faktur/Dokumen", "Nomor Dokumen"],
      date: ["Tanggal Faktur/Dokumen", "Tanggal Dokumen"],
      name: ["Nama Penjual"],
      npwp: ["NPWP/NIK Penjual", "NPWP Penjual"],
      dpp: ["Harga Jual / Penggantian / DPP", "DPP / Penghasilan Kotor"],
      ppn: ["PPN"],
    },
    accounts: purchaseAccounts(true),
    lines: purchaseLines,
  },
  {
    id: "b2",
    category: "purchase_b2",
    icon: "i-buy",
    title: "Faktur Pajak Masukan B2",
    tag: "Dalam negeri",
    desc: "Pembelian dalam negeri yang PPN-nya dapat dikreditkan.",
    prefix: "PEMB",
    party: "Penjual",
    parseNum: parseFloat,
    descSep: " - ",
    columns: purchaseColumns(false),
    accounts: purchaseAccounts(true),
    lines: purchaseLines,
  },
  {
    id: "b3",
    category: "purchase_b3",
    icon: "i-lock",
    title: "Faktur Pajak Masukan B3",
    tag: "PPN tidak dikreditkan",
    desc: "Pembelian yang PPN-nya tidak dapat dikreditkan (ikut jadi biaya).",
    prefix: "PEMB.NON",
    party: "Penjual",
    parseNum: parseInt,
    descSep: "\n ",
    columns: purchaseColumns(true),
    // Kode transaksi yang PPN-nya tidak benar-benar dibayar ke penjual:
    // 07 = PPN tidak dipungut, 08 = PPN dibebaskan. Hutang/pembelian = harga jual saja.
    ppnIgnoredCodes: ["07", "08"],
    codeColumns: ["Kode Transaksi", "Kode Jenis Transaksi"],
    schemaNote:
      "Kode transaksi 07/08 (PPN tidak dipungut/dibebaskan): PPN diabaikan, pembelian & hutang = Harga Jual saja.",
    accounts: purchaseAccounts(false),
    lines(r, a, total) {
      return `
            <ACCOUNTLINE operation="Add">
                <KeyID>0</KeyID>
                <GLACCOUNT>${a.purchase}</GLACCOUNT>
                <GLAMOUNT>${total}</GLAMOUNT>
                <DESCRIPTION>${r.name} - ${r.invoiceNo}</DESCRIPTION>
                <RATE>1</RATE>
                <PRIMEAMOUNT>${total}</PRIMEAMOUNT>
                <TXDATE/>
                <POSTED/>
                <CURRENCYNAME>IDR</CURRENCYNAME>
            </ACCOUNTLINE>
            <ACCOUNTLINE operation="Add">
                <KeyID>2</KeyID>
                <GLACCOUNT>${a.payable}</GLACCOUNT>
                <GLAMOUNT>${-total}</GLAMOUNT>
                <vendorNO>1000</vendorNO>
                <DESCRIPTION>${r.name} - ${r.invoiceNo}</DESCRIPTION>
                <RATE>1</RATE>
                <PRIMEAMOUNT>${-total}</PRIMEAMOUNT>
                <TXDATE/>
                <POSTED/>
                <CURRENCYNAME>IDR</CURRENCYNAME>
            </ACCOUNTLINE>`;
    },
  },
];

function purchaseColumns(withImportDpp) {
  const dpp = ["Harga Jual / Penggantian / DPP", "Harga Jual/Penggantian/DPP (Rupiah)", "DPP (Rupiah)"];
  if (withImportDpp) dpp.push("Harga Jual/Penggantian/Nilai Impor/DPP (Rupiah)");
  return {
    invoiceNo: [
      "Nomor Faktur/Dokumen",
      "Nomor Faktur Pajak",
      "Faktur Pajak/Dokumen Tertentu/Nota Retur/Nota Pembatalan - Nomor",
      "Dokumen Tertentu - Nomor",
    ],
    date: [
      "Tanggal Faktur/Dokumen",
      "Tanggal Faktur Pajak",
      "Faktur Pajak/Dokumen Tertentu/Nota Retur/Nota Pembatalan - Tanggal",
      "Dokumen Tertentu - Tanggal",
    ],
    name: [
      "Nama Penjual",
      "Nama Penjual Barang Kena Pajak/Barang Kena Pajak Tidak Berwujud/Jasa Kena Pajak",
      "Nama Penjual BKP/BKP Tidak Berwujud/Pemberi JKP",
    ],
    npwp: ["NPWP/NIK Penjual", "Nomor Identitas WP", "NPWP Penjual", "NPWP"],
    dpp,
    ppn: ["PPN", "PPN (Rupiah)"],
  };
}

function purchaseAccounts(creditable) {
  return [
    {
      key: "purchase", label: "Akun Pembelian", side: "debit",
      amount: creditable ? "DPP" : "DPP + PPN",
      hint: creditable
        ? "Biaya/persediaan dari pembelian (sebelum PPN)."
        : "Biaya/persediaan dari pembelian. PPN ikut digabung ke akun ini.",
      suggestions: [
        ["5000.01", "Material"],
        ["5000.02", "Bahan Pembantu"],
        ["5000.03", "Jasa/Subcont"],
        ["5000.04", "Consumable"],
      ],
    },
    {
      key: "tax", label: "Akun PPN Masukan", side: "debit", amount: "PPN",
      hint: creditable
        ? "PPN yang Anda bayar ke penjual dan dapat dikreditkan."
        : "Tidak dipakai — PPN jenis B3 tidak dapat dikreditkan.",
      suggestions: [["1600.01", "PPN Pembelian"]],
      disabled: !creditable,
    },
    {
      key: "payable", label: "Akun Hutang Usaha", side: "credit", amount: "DPP + PPN",
      hint: "Kewajiban bayar ke penjual (nilai faktur termasuk PPN).",
      suggestions: [["210102", "Hutang Usaha"]],
    },
  ];
}

// Jurnal pembelian B1 & B2: Pembelian (D), PPN Masukan (D), Hutang (K)
function purchaseLines(r, a, total) {
  return `
            <ACCOUNTLINE operation="Add">
                <KeyID>0</KeyID>
                <GLACCOUNT>${a.purchase}</GLACCOUNT>
                <GLAMOUNT>${r.dpp}</GLAMOUNT>
                <DESCRIPTION>${r.name} - ${r.invoiceNo}</DESCRIPTION>
                <RATE>1</RATE>
                <PRIMEAMOUNT>${r.dpp}</PRIMEAMOUNT>
                <TXDATE/>
                <POSTED/>
                <CURRENCYNAME>IDR</CURRENCYNAME>
            </ACCOUNTLINE>
            <ACCOUNTLINE operation="Add">
                <KeyID>1</KeyID>
                <GLACCOUNT>${a.tax}</GLACCOUNT>
                <GLAMOUNT>${r.ppn}</GLAMOUNT>
                <DESCRIPTION>PPN 12% - ${r.name}</DESCRIPTION>
                <RATE>1</RATE>
                <PRIMEAMOUNT>${r.ppn}</PRIMEAMOUNT>
                <TXDATE/>
                <POSTED/>
                <CURRENCYNAME>IDR</CURRENCYNAME>
            </ACCOUNTLINE>
            <ACCOUNTLINE operation="Add">
                <KeyID>2</KeyID>
                <GLACCOUNT>${a.payable}</GLACCOUNT>
                <GLAMOUNT>${-total}</GLAMOUNT>
                <vendorNO>1000</vendorNO>
                <DESCRIPTION>${r.name} - ${r.invoiceNo}</DESCRIPTION>
                <RATE>1</RATE>
                <PRIMEAMOUNT>${-total}</PRIMEAMOUNT>
                <TXDATE/>
                <POSTED/>
                <CURRENCYNAME>IDR</CURRENCYNAME>
            </ACCOUNTLINE>`;
}

// =================================================================
// LOGIKA KONVERSI
// =================================================================
// Judul kolom dibandingkan tanpa spasi/tanda baca & huruf besar-kecil,
// jadi "Harga Jual / Penggantian / DPP" = "Harga Jual/Penggantian/DPP".
const normKey = (s) => String(s).toLowerCase().replace(/[^a-z0-9]/g, "");

function normalizeRow(row) {
  const out = {};
  for (const [k, v] of Object.entries(row)) {
    const nk = normKey(k);
    if (!(nk in out)) out[nk] = v;
  }
  Object.defineProperty(out, "__rowNum__", { value: row.__rowNum__ });
  return out;
}

// Baris judul kolom = baris pertama (dari 20 teratas) yang memuat minimal
// dua nama kolom yang dikenal, untuk berjaga-jaga jika ada judul laporan
// di atas tabel. File Coretax biasanya langsung di baris 1.
function readRows(worksheet) {
  if (!worksheet || !worksheet["!ref"]) return [];
  const known = new Set(
    [...MODES.flatMap((m) => Object.values(m.columns).flat()), ...PERIOD_COLUMNS].map(normKey)
  );
  const top = XLSX.utils.sheet_to_json(worksheet, { header: 1, blankrows: true }).slice(0, 20);
  const headerAt = top.findIndex((cells) => (cells || []).filter((c) => known.has(normKey(c))).length >= 2);
  const firstRow = XLSX.utils.decode_range(worksheet["!ref"]).s.r;
  return XLSX.utils.sheet_to_json(worksheet, { range: firstRow + Math.max(headerAt, 0) }).map(normalizeRow);
}

function pick(row, names) {
  for (const n of names) {
    const v = row[normKey(n)];
    if (v !== undefined && v !== null && v !== "") return v;
  }
  return undefined;
}

// Masa Pajak -> kunci periode "YYMM", mis. "Januari 2026" -> "2601".
// Juga dikenali: "01/2026", "2026-01", "012026", tanggal Excel, atau
// angka bulan 1–12 dengan tahun dari kolom Tahun Pajak.
function parseMasaPajak(value, taxYear) {
  if (value === undefined) return null;
  let month = null;
  let year = null;
  let m;
  if (typeof value === "number" && Number.isInteger(value) && value >= 1 && value <= 12) {
    month = value;
  } else {
    const s = String(value).trim().toLowerCase();
    const words = s.match(/[a-z]+/g) || [];
    const named = MONTH_ALIASES.findIndex((names) =>
      words.some((w) => w.length >= 3 && names.some((n) => n.startsWith(w)))
    );
    if (named >= 0) {
      month = named + 1;
      m = /(?:^|\D)(20\d{2})(?!\d)/.exec(s);
      year = m && Number(m[1]);
    } else if ((m = /^(\d{1,2})\s*[-/.\s]\s*(20\d{2})$/.exec(s))) {
      [month, year] = [Number(m[1]), Number(m[2])];
    } else if ((m = /^(20\d{2})\s*[-/.\s]\s*(\d{1,2})$/.exec(s))) {
      [year, month] = [Number(m[1]), Number(m[2])];
    } else if ((m = /^(\d{1,2})(20\d{2})$/.exec(s))) {
      [month, year] = [Number(m[1]), Number(m[2])];
    } else if (/^\d{1,2}$/.test(s)) {
      month = Number(s);
    } else if ((m = /^(\d{4})-(\d{2})-\d{2}$/.exec(formatDate(value)))) {
      [year, month] = [Number(m[1]), Number(m[2])];
    }
  }
  if (!year && taxYear !== undefined) {
    m = /(20\d{2})/.exec(String(taxYear));
    year = m && Number(m[1]);
  }
  if (!(month >= 1 && month <= 12 && year >= 2000 && year <= 2099)) return null;
  return String(year).slice(2) + String(month).padStart(2, "0");
}

function readPeriod(row) {
  return parseMasaPajak(pick(row, PERIOD_COLUMNS), pick(row, TAX_YEAR_COLUMNS));
}

// Kode transaksi faktur = 2 digit pertama nomor faktur
// (mis. 08002500012345678 atau 080.000-25.12345678 -> "08").
function transactionCode(mode, row, invoiceNo) {
  const explicit = pick(row, mode.codeColumns || []);
  const m = /\d{1,2}/.exec(String(explicit ?? ""));
  if (m) return m[0].padStart(2, "0");
  if (invoiceNo === undefined) return null;
  let digits = String(invoiceNo).replace(/\D/g, "");
  // Excel menyimpan nomor sebagai angka -> nol di depan hilang
  if (typeof invoiceNo === "number" && digits.length === 16) digits = "0" + digits;
  return digits.length >= 16 ? digits.slice(0, 2) : null;
}

function extractRecord(mode, row) {
  const c = mode.columns;
  const r = {
    rowNum: row.__rowNum__ + 1,
    invoiceNo: pick(row, c.invoiceNo),
    rawDate: pick(row, c.date),
    name: pick(row, c.name),
    npwp: pick(row, c.npwp),
    dpp: mode.parseNum(pick(row, c.dpp)),
    ppn: mode.parseNum(pick(row, c.ppn)),
    period: readPeriod(row),
  };
  if (mode.ppnIgnoredCodes) {
    r.code = transactionCode(mode, row, r.invoiceNo);
    if (mode.ppnIgnoredCodes.includes(r.code)) {
      r.ppnIgnored = isNaN(r.ppn) ? 0 : r.ppn;
      r.ppn = 0;
    }
  }
  r.date = formatDate(r.rawDate);
  const reasons = [];
  if (!r.invoiceNo) reasons.push("Nomor faktur kosong");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(r.date ?? ""))) reasons.push("Tanggal kosong/tidak dikenali");
  if (state.autoPeriod && !r.period) reasons.push("Masa Pajak kosong/tidak dikenali");
  if (!r.name) reasons.push(`Nama ${mode.party.toLowerCase()} kosong`);
  if (r.npwp === undefined) reasons.push("NPWP kosong");
  if (isNaN(r.dpp)) reasons.push("DPP kosong/bukan angka");
  if (isNaN(r.ppn)) reasons.push("PPN kosong/bukan angka");
  r.valid = reasons.length === 0;
  r.reasons = reasons;
  r.total = r.valid ? r.dpp + r.ppn : NaN;
  return r;
}

// `key` = periode "YYMM", mis. jvNumber(mode, "2601", 1) -> "PENJ.26.01.001"
function jvNumber(mode, key, index) {
  return `${mode.prefix}.${key.slice(0, 2)}.${key.slice(2)}.${String(index).padStart(3, "0")}`;
}

// `jvs` = nomor jurnal per baris, sejajar dengan `records`
function buildXml(mode, records, accounts, jvs) {
  let transactions = "";
  let requestId = 1;

  records.forEach((rec, i) => {
    if (!rec.valid) return;
    // Teks dari Excel di-escape agar nama seperti "CV. MAJU & JAYA" tidak merusak XML
    const r = { ...rec, name: esc(rec.name), npwp: esc(rec.npwp), invoiceNo: esc(rec.invoiceNo) };
    const total = r.dpp + r.ppn;
    transactions += `
        <JV operation="Add" REQUESTID="${requestId}">
            <TRANSACTIONID>148</TRANSACTIONID>
            ${mode.lines(r, accounts, total)}
            <JVNUMBER>${jvs[i]}</JVNUMBER>
            <TRANSDATE>${r.date}</TRANSDATE>
            <SOURCE>GL</SOURCE>
            <TRANSTYPE>journal voucher</TRANSTYPE>
            <TRANSDESCRIPTION>${r.name}${mode.descSep}${r.npwp}${mode.descSep}${r.invoiceNo}</TRANSDESCRIPTION>
            <JVAMOUNT>${total}</JVAMOUNT>
        </JV>`;
    requestId++;
  });

  return `<?xml version="1.0"?>\n<NMEXML EximID="1" BranchCode="${BRANCH_CODE}" ACCOUNTANTCOPYID="">\n<TRANSACTIONS OnError="CONTINUE">\n${transactions}\n</TRANSACTIONS>\n</NMEXML>`;
}

function formatDate(serial) {
  if (typeof serial === "string") {
    const s = serial.trim();
    const pad = (n) => n.padStart(2, "0");
    // Tanggal berupa teks: dd/mm/yyyy (format Indonesia) atau yyyy-mm-dd
    let m = /^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$/.exec(s);
    if (m) return `${m[3]}-${pad(m[2])}-${pad(m[1])}`;
    m = /^(\d{4})[/.-](\d{1,2})[/.-](\d{1,2})/.exec(s);
    if (m) return `${m[1]}-${pad(m[2])}-${pad(m[3])}`;
  }
  if (typeof serial === "number") {
    const utc_days = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);
    const year = date_info.getFullYear();
    const month = String(date_info.getMonth() + 1).padStart(2, "0");
    const day = String(date_info.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }
  return serial;
}

// =================================================================
// STATE & ELEMEN
// =================================================================
const $ = (id) => document.getElementById(id);
const el = {
  modeGrid: $("modeGrid"),
  month: $("journalMonth"),
  year: $("journalYear"),
  startIndex: $("startIndex"),
  jvPreview: $("jvPreview"),
  continueIndex: $("continueIndex"),
  periodManual: $("periodManual"),
  periodNote: $("periodNote"),
  periodAuto: $("periodAuto"),
  periodSource: $("periodSource"),
  periodList: $("periodList"),
  accountFields: $("accountFields"),
  journalSchema: $("journalSchema"),
  dropzone: $("dropzone"),
  fileInput: $("fileInput"),
  fileCard: $("fileCard"),
  fileName: $("fileName"),
  fileInfo: $("fileInfo"),
  changeFile: $("changeFile"),
  columnsHelpBody: $("columnsHelpBody"),
  previewArea: $("previewArea"),
  stats: $("stats"),
  formatWarning: $("formatWarning"),
  rowFilter: $("rowFilter"),
  partyHeader: $("partyHeader"),
  previewBody: $("previewBody"),
  sideMode: $("sideMode"),
  checklist: $("checklist"),
  convertButton: $("convertButton"),
  convertHint: $("convertHint"),
  resultCard: $("resultCard"),
  resultText: $("resultText"),
  downloadLink: $("downloadLink"),
  xmlOutput: $("xmlOutput"),
  historyList: $("historyList"),
  historyModeLabel: $("historyModeLabel"),
  serverStatus: $("serverStatus"),
  toasts: $("toasts"),
};

const state = {
  mode: MODES[0],
  file: null,
  rawRows: [],
  records: [],
  filter: "all",
  periodTouched: false,
  // true jika file punya kolom Masa Pajak: periode jurnal diambil per baris
  autoPeriod: false,
  // nomor urut awal per periode "YYMM" yang diisi user (mode Masa Pajak)
  starts: {},
  storageOk: null,
  downloadUrl: null,
};

// =================================================================
// UTILITAS
// =================================================================
const rupiah = new Intl.NumberFormat("id-ID", { maximumFractionDigits: 2 });
const fmtNum = (n) => (isNaN(n) ? "–" : rupiah.format(n));

function esc(v) {
  return String(v ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function store(key, value) {
  try {
    localStorage.setItem(STORAGE_PREFIX + key, JSON.stringify(value));
  } catch (_) {}
}
function load(key, fallback) {
  try {
    const v = localStorage.getItem(STORAGE_PREFIX + key);
    return v === null ? fallback : JSON.parse(v);
  } catch (_) {
    return fallback;
  }
}

function icon(name, cls = "icon") {
  return `<svg class="${cls}"><use href="#${name}" /></svg>`;
}

function toast(message, type = "info") {
  const t = document.createElement("div");
  t.className = `toast ${type}`;
  const ic = type === "success" ? "i-check" : type === "error" ? "i-alert" : "i-info";
  t.innerHTML = `${icon(ic)}<div>${esc(message)}</div>`;
  el.toasts.appendChild(t);
  setTimeout(() => t.remove(), type === "error" ? 6000 : 3500);
}

function confirmDialog(title, text, okLabel = "Hapus") {
  const dlg = $("confirmDialog");
  $("confirmTitle").textContent = title;
  $("confirmText").textContent = text;
  $("confirmOk").textContent = okLabel;
  dlg.showModal();
  return new Promise((resolve) => {
    const done = (v) => {
      dlg.close();
      $("confirmOk").removeEventListener("click", ok);
      dlg.removeEventListener("close", cancel);
      resolve(v);
    };
    const ok = () => done(true);
    const cancel = () => done(false);
    $("confirmOk").addEventListener("click", ok);
    dlg.addEventListener("close", cancel, { once: true });
  });
}

function displayDate(d) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(d ?? ""));
  return m ? `${m[3]}/${m[2]}/${m[1]}` : esc(d ?? "–");
}

// Periode "YYMM" dari pilihan bulan & tahun (dipakai jika tidak ada Masa Pajak)
function manualPeriodKey() {
  return el.year.value + el.month.value;
}

function periodLabel(key) {
  const mm = key.slice(2);
  return `${MONTHS[Number(mm) - 1] ?? mm} 20${key.slice(0, 2)}`;
}

function startIndex() {
  const n = parseInt(el.startIndex.value, 10);
  return Number.isFinite(n) && n > 0 ? n : 1;
}

// Nomor urut awal untuk satu periode Masa Pajak; jika belum diisi,
// ikut nomor urut di isian manual (bawaan 1).
function periodStart(key) {
  const n = parseInt(state.starts[key], 10);
  return Number.isFinite(n) && n > 0 ? n : startIndex();
}

// Faktur valid dikelompokkan per periode jurnal: [{ key: "2601", count, start }].
// Dengan kolom Masa Pajak, periode diambil dari tiap baris; tanpa itu
// semua faktur memakai bulan & tahun yang dipilih manual.
function periodGroups() {
  const valid = state.records.filter((r) => r.valid);
  if (!state.autoPeriod) {
    return [{ key: manualPeriodKey(), count: valid.length, start: startIndex() }];
  }
  const counts = {};
  for (const r of valid) counts[r.period] = (counts[r.period] || 0) + 1;
  return Object.keys(counts)
    .sort()
    .map((key) => ({ key, count: counts[key], start: periodStart(key) }));
}

// Nomor jurnal tiap baris (null untuk yang dilewati), berurutan per periode
function jvNumbers(groups) {
  const next = {};
  for (const g of groups) next[g.key] = g.start;
  return state.records.map((r) => {
    if (!r.valid) return null;
    const key = state.autoPeriod ? r.period : groups[0].key;
    return jvNumber(state.mode, key, next[key]++);
  });
}

function jvRange(g) {
  const first = jvNumber(state.mode, g.key, g.start);
  return g.count > 1 ? `${first} s/d ${jvNumber(state.mode, g.key, g.start + g.count - 1)}` : first;
}

function accountValues() {
  const values = {};
  for (const a of state.mode.accounts) {
    if (a.disabled) continue;
    const input = $(`acc-${a.key}`);
    values[a.key] = input ? input.value.trim() : "";
  }
  return values;
}

// =================================================================
// RENDER: JENIS FAKTUR
// =================================================================
function renderModes() {
  el.modeGrid.innerHTML = MODES.map(
    (m) => `
      <button type="button" class="mode-card" role="radio" data-mode="${m.id}">
        <span class="mode-icon">${icon(m.icon)}</span>
        <span>
          <span class="mode-title">${esc(m.title.replace("Faktur Pajak ", ""))}<span class="tag">${esc(m.tag)}</span></span>
          <span class="mode-desc">${esc(m.desc)}</span>
        </span>
      </button>`
  ).join("");
  el.modeGrid.addEventListener("click", (e) => {
    const card = e.target.closest("[data-mode]");
    if (card) location.hash = card.dataset.mode;
  });
}

function setMode(id) {
  const mode = MODES.find((m) => m.id === id) || MODES[0];
  const changed = mode !== state.mode;
  state.mode = mode;
  store("lastMode", mode.id);

  el.modeGrid.querySelectorAll("[data-mode]").forEach((c) => {
    const active = c.dataset.mode === mode.id;
    c.classList.toggle("active", active);
    c.setAttribute("aria-checked", String(active));
  });
  document.title = `${mode.title} — Konverter ke Accurate 5`;
  el.partyHeader.textContent = mode.party;
  el.sideMode.innerHTML = `${esc(mode.title)}<span class="muted">${esc(mode.tag)} · nomor jurnal ${esc(mode.prefix)}.…</span>`;
  el.historyModeLabel.textContent = mode.title.replace("Faktur Pajak ", "");

  renderAccounts();
  renderColumnsHelp();
  if (state.rawRows.length) processRows();
  if (changed) hideResult();
  updateAll();
  loadHistory();
}

// =================================================================
// RENDER: PERIODE
// =================================================================
function renderPeriod() {
  const now = new Date();
  el.month.innerHTML = MONTHS.map(
    (name, i) => `<option value="${String(i + 1).padStart(2, "0")}">${name}</option>`
  ).join("");
  const thisYear = now.getFullYear();
  const years = [];
  for (let y = 2020; y <= thisYear + 1; y++) years.push(y);
  el.year.innerHTML = years
    .reverse()
    .map((y) => `<option value="${String(y).slice(2)}">${y}</option>`)
    .join("");
  el.month.value = String(now.getMonth() + 1).padStart(2, "0");
  el.year.value = String(thisYear).slice(2);

  const onChange = () => {
    state.periodTouched = true;
    hideResult();
    updateAll();
  };
  el.month.addEventListener("change", onChange);
  el.year.addEventListener("change", onChange);
  el.startIndex.addEventListener("input", () => {
    hideResult();
    updateAll();
  });
  el.continueIndex.addEventListener("click", () => {
    el.startIndex.value = el.continueIndex.dataset.next;
    hideResult();
    updateAll();
  });

  el.periodList.addEventListener("input", (e) => {
    const input = e.target.closest("[data-start]");
    if (!input) return;
    state.starts[input.dataset.start] = input.value;
    hideResult();
    updateAll();
  });
  el.periodList.addEventListener("click", (e) => {
    const chip = e.target.closest("[data-continue]");
    if (!chip) return;
    const key = chip.dataset.continue;
    state.starts[key] = chip.dataset.next;
    $(`start-${key}`).value = chip.dataset.next;
    hideResult();
    updateAll();
  });
}

function lastIndexKey(key) {
  return `lastIndex.${state.mode.category}.${key}`;
}

// Tampilkan tombol "Lanjutkan dari …" jika nomor awal bentrok dengan
// jurnal periode yang sama yang sudah pernah dibuat.
function updateContinueChip(chip, key, start) {
  const last = load(lastIndexKey(key), null);
  const show = !!last && start <= last;
  chip.classList.toggle("hidden", !show);
  if (!show) return;
  const pad = (n) => String(n).padStart(3, "0");
  chip.dataset.next = last + 1;
  chip.textContent = `Lanjutkan dari ${pad(last + 1)} (terakhir dipakai ${pad(last)})`;
}

function renderPeriods(groups) {
  const auto = state.autoPeriod;
  el.periodManual.classList.toggle("hidden", auto);
  el.periodAuto.classList.toggle("hidden", !auto);

  if (!auto) {
    const key = manualPeriodKey();
    el.jvPreview.textContent = jvNumber(state.mode, key, startIndex());
    updateContinueChip(el.continueIndex, key, startIndex());
    const noColumn = !!state.file;
    el.periodNote.className = `notice ${noColumn ? "warn" : "info"}`;
    el.periodNote.innerHTML = `${icon(noColumn ? "i-alert" : "i-info")}<div class="notice-body">${
      noColumn
        ? "File ini tidak memiliki kolom <b>Masa Pajak</b>. Pastikan bulan &amp; tahun di atas sudah benar."
        : "Bulan &amp; tahun akan terisi otomatis dari kolom <b>Masa Pajak</b> setelah file Excel diunggah di langkah 4."
    }</div>`;
    return;
  }

  el.periodSource.innerHTML =
    groups.length > 1
      ? `Bulan &amp; tahun dibaca dari kolom <b>Masa Pajak</b>. File ini berisi ${groups.length} masa pajak, nomor urut dihitung terpisah untuk masing-masing.`
      : "Bulan &amp; tahun dibaca dari kolom <b>Masa Pajak</b> di file Excel.";

  // Susun ulang baris hanya jika daftar periodenya berubah, supaya kursor
  // tidak hilang saat mengetik nomor urut.
  const sig = groups.map((g) => g.key).join(",");
  if (el.periodList.dataset.sig !== sig) {
    el.periodList.dataset.sig = sig;
    el.periodList.innerHTML =
      groups
        .map(
          (g) => `
        <div class="period-row" data-period="${g.key}">
          <div class="period-name">
            <strong>${esc(periodLabel(g.key))}</strong>
            <span class="muted small period-count"></span>
          </div>
          <div class="period-start">
            <label for="start-${g.key}">Mulai dari nomor urut</label>
            <input type="number" id="start-${g.key}" data-start="${g.key}" min="1" step="1"
              inputmode="numeric" value="${g.start}" />
          </div>
          <div class="period-jv">
            <span class="period-range"></span>
            <button type="button" class="chip hidden" data-continue="${g.key}"></button>
          </div>
        </div>`
        )
        .join("") || `<p class="muted small">Belum ada faktur yang bisa diproses.</p>`;
  }
  for (const g of groups) {
    const row = el.periodList.querySelector(`[data-period="${g.key}"]`);
    row.querySelector(".period-count").textContent = `${g.count} faktur`;
    row.querySelector(".period-range").innerHTML = jvRange(g)
      .split(" s/d ")
      .map((jv) => `<code>${esc(jv)}</code>`)
      .join(`<span class="muted small">s/d</span>`);
    updateContinueChip(row.querySelector("[data-continue]"), g.key, g.start);
  }
}

// Dipanggil saat file baru dimuat agar daftar periode disusun ulang
function resetPeriods() {
  state.starts = {};
  delete el.periodList.dataset.sig;
}

// =================================================================
// RENDER: AKUN
// =================================================================
function renderAccounts() {
  const saved = load(`accounts.${state.mode.category}`, {});
  el.accountFields.innerHTML = state.mode.accounts
    .map((a) => {
      const sideLabel = a.side === "debit" ? "Debit" : "Kredit";
      const value = a.disabled ? "" : saved[a.key] ?? "";
      const chips = a.disabled
        ? ""
        : `<div class="chips">${a.suggestions
            .map(
              ([code, name]) =>
                `<button type="button" class="chip" data-fill="${a.key}" data-value="${esc(code)}"><code>${esc(code)}</code> · ${esc(name)}</button>`
            )
            .join("")}</div>`;
      return `
        <div class="account ${a.disabled ? "disabled" : ""}">
          <div>
            <label class="account-label" for="acc-${a.key}">
              ${esc(a.label)}
              <span class="side-badge ${a.side}">${sideLabel}</span>
            </label>
            <div class="account-hint">${esc(a.hint)}</div>
          </div>
          <div>
            <input type="text" id="acc-${a.key}" value="${esc(value)}"
              placeholder="${a.disabled ? "Tidak diperlukan" : "Contoh: " + esc(a.suggestions[0][0])}"
              ${a.disabled ? "disabled" : ""} autocomplete="off" spellcheck="false" />
            ${chips}
          </div>
        </div>`;
    })
    .join("");
}

function onAccountChange() {
  const values = accountValues();
  store(`accounts.${state.mode.category}`, values);
  hideResult();
  updateAll();
}

function updateChipsAndSchema() {
  const values = accountValues();
  el.accountFields.querySelectorAll(".chip").forEach((chip) => {
    chip.classList.toggle("selected", values[chip.dataset.fill] === chip.dataset.value);
  });

  const rows = state.mode.accounts
    .filter((a) => !a.disabled)
    .map((a) => {
      const acc = values[a.key];
      const d = a.side === "debit";
      const star = state.mode.schemaNote && a.amount.includes("PPN") ? "*" : "";
      return `<tr>
        <td>${d ? "" : "&emsp;&emsp;"}${acc ? `<strong>${esc(acc)}</strong>` : `<span class="muted">[belum diisi]</span>`} <span class="muted">${esc(a.label.replace("Akun ", ""))}</span></td>
        <td class="amt d">${d ? esc(a.amount) + star : ""}</td>
        <td class="amt c">${d ? "" : esc(a.amount) + star}</td>
      </tr>`;
    })
    .join("");
  el.journalSchema.innerHTML = `
    <div class="schema-title">Bentuk jurnal yang akan dibuat untuk setiap faktur</div>
    <table>
      <tr><td class="muted small">Akun</td><td class="amt muted small">Debit</td><td class="amt muted small">Kredit</td></tr>
      ${rows}
    </table>
    ${state.mode.schemaNote ? `<div class="schema-note">* ${esc(state.mode.schemaNote)}</div>` : ""}`;
}

// =================================================================
// FILE EXCEL
// =================================================================
function renderColumnsHelp() {
  const labels = {
    invoiceNo: "No. Faktur",
    date: "Tanggal",
    name: `Nama ${state.mode.party}`,
    npwp: "NPWP",
    dpp: "DPP",
    ppn: "PPN",
  };
  const codes = (names) => names.map((n) => `<code>${esc(n)}</code>`).join("");
  const rows = Object.entries(state.mode.columns)
    .map(([k, names]) => `<tr><td>${labels[k]}</td><td>${codes(names)}</td></tr>`)
    .join("");
  el.columnsHelpBody.innerHTML = `
    <p class="muted">Judul kolom di Excel harus memuat salah satu nama berikut
      (huruf besar/kecil dan spasi tidak berpengaruh):</p>
    <table>${rows}
      <tr><td>Masa Pajak</td><td>${codes(PERIOD_COLUMNS)}
        <div class="muted">Opsional — menentukan bulan &amp; tahun pada nomor jurnal, mis. Januari 2026 → .26.01.</div></td></tr>
    </table>`;
}

function handleFile(file) {
  if (!file) return;
  if (!/\.xlsx?$/i.test(file.name)) {
    toast("File harus berformat Excel (.xlsx atau .xls).", "error");
    return;
  }
  if (typeof XLSX === "undefined") {
    toast("Pembaca Excel gagal dimuat. Periksa koneksi internet lalu muat ulang halaman.", "error");
    return;
  }

  const reader = new FileReader();
  reader.onload = (event) => {
    try {
      const workbook = XLSX.read(event.target.result, { type: "array" });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      state.rawRows = readRows(worksheet);
      state.autoPeriod = state.rawRows.some((row) => readPeriod(row));
      resetPeriods();
      state.file = file;
      state.filter = "all";
      hideResult();
      processRows();
      if (state.autoPeriod) announcePeriods();
      else autoDetectPeriod();
      updateAll();
    } catch (error) {
      console.error("Error processing file:", error);
      toast("File tidak dapat dibaca: " + error.message, "error");
    }
  };
  reader.onerror = () => toast("Gagal membaca file.", "error");
  reader.readAsArrayBuffer(file);
}

function processRows() {
  state.records = state.rawRows.map((row) => extractRecord(state.mode, row));
}

function announcePeriods() {
  const labels = periodGroups().map((g) => periodLabel(g.key));
  if (labels.length) toast(`Periode jurnal diambil dari Masa Pajak: ${labels.join(", ")}.`, "info");
}

// File tanpa kolom Masa Pajak: jika user belum memilih periode sendiri,
// isi periode otomatis sesuai bulan terbanyak pada tanggal faktur.
function autoDetectPeriod() {
  if (state.periodTouched) return;
  const counts = {};
  for (const r of state.records) {
    const m = /^(\d{4})-(\d{2})-\d{2}$/.exec(String(r.date ?? ""));
    if (r.valid && m) counts[m[1] + m[2]] = (counts[m[1] + m[2]] || 0) + 1;
  }
  const top = Object.entries(counts).sort((a, b) => b[1] - a[1])[0];
  if (!top) return;
  const yy = top[0].slice(2, 4);
  const mm = top[0].slice(4, 6);
  if (![...el.year.options].some((o) => o.value === yy)) return;
  if (el.year.value === yy && el.month.value === mm) return;
  el.year.value = yy;
  el.month.value = mm;
  toast(`Periode disesuaikan ke ${periodLabel(yy + mm)} sesuai tanggal faktur.`, "info");
}

function renderFile(jvs = jvNumbers(periodGroups())) {
  const hasFile = !!state.file;
  el.dropzone.classList.toggle("hidden", hasFile);
  el.fileCard.classList.toggle("hidden", !hasFile);
  el.previewArea.classList.toggle("hidden", !hasFile);
  if (!hasFile) return;

  el.fileName.textContent = state.file.name;
  el.fileInfo.textContent = `${(state.file.size / 1024).toFixed(1)} KB · ${state.rawRows.length} baris data`;

  const valid = state.records.filter((r) => r.valid);
  const skipped = state.records.length - valid.length;
  const sum = (k) => valid.reduce((s, r) => s + r[k], 0);

  el.stats.innerHTML = `
    <div class="stat good"><div class="stat-label">Akan diproses</div><div class="stat-value">${valid.length}</div></div>
    <div class="stat ${skipped ? "bad" : ""}"><div class="stat-label">Dilewati</div><div class="stat-value">${skipped}</div></div>
    <div class="stat"><div class="stat-label">Total DPP</div><div class="stat-value" title="${fmtNum(sum("dpp"))}">${fmtNum(sum("dpp"))}</div></div>
    <div class="stat"><div class="stat-label">Total PPN</div><div class="stat-value" title="${fmtNum(sum("ppn"))}">${fmtNum(sum("ppn"))}</div></div>`;

  renderFormatWarning(valid.length);

  // Filter
  el.rowFilter.querySelectorAll("button").forEach((b) => {
    b.classList.toggle("active", b.dataset.filter === state.filter);
    const n = b.dataset.filter === "all" ? state.records.length : b.dataset.filter === "valid" ? valid.length : skipped;
    b.innerHTML = `${b.textContent.replace(/\s*\d+$/, "")}<span class="count">${n}</span>`;
  });

  // Tabel
  const rows = [];
  state.records.forEach((r, i) => {
    const jv = jvs[i];
    if (state.filter === "valid" && !r.valid) return;
    if (state.filter === "skip" && r.valid) return;
    if (rows.length >= 1000) return;
    rows.push(`
      <tr class="${r.valid ? "" : "skip"}">
        <td class="muted">${r.rowNum}</td>
        <td class="inv">${esc(r.invoiceNo ?? "–")}<div class="npwp">${
          r.code ? `<span class="code-tag">Kode ${esc(r.code)}</span> ` : ""
        }${displayDate(r.date)}</div></td>
        <td class="name">${esc(r.name ?? "–")}<div class="npwp">${esc(r.npwp ?? "")}</div></td>
        <td class="num hide-sm">${fmtNum(r.dpp)}</td>
        <td class="num hide-sm">${
          r.ppnIgnored !== undefined
            ? `0<div class="ignored" title="PPN tidak dihitung karena kode ${esc(r.code)}">${fmtNum(r.ppnIgnored)}</div>`
            : fmtNum(r.ppn)
        }</td>
        <td class="num"><strong>${fmtNum(r.total)}</strong></td>
        <td>${
          r.valid
            ? `<span class="badge ok">${icon("i-check")}${esc(jv)}</span>`
            : `<span class="badge skip">${icon("i-alert")}Dilewati</span><div class="reason">${esc(r.reasons.join(", "))}</div>`
        }</td>
      </tr>`);
  });
  el.previewBody.innerHTML =
    rows.join("") ||
    `<tr class="empty-row"><td colspan="7">${
      state.records.length ? "Tidak ada baris untuk filter ini." : "File tidak berisi data."
    }</td></tr>`;
}

function renderFormatWarning(validCount) {
  if (!state.records.length || validCount > 0) {
    el.formatWarning.classList.add("hidden");
    return;
  }
  const headers = new Set(state.rawRows.flatMap((r) => Object.keys(r)));
  const labels = { invoiceNo: "No. Faktur", date: "Tanggal", name: "Nama", npwp: "NPWP", dpp: "DPP", ppn: "PPN" };
  const missing = Object.entries(state.mode.columns)
    .filter(([, names]) => !names.some((n) => headers.has(normKey(n))))
    .map(([k]) => labels[k]);

  // Cari jenis faktur lain yang cocok dengan file ini. Lampiran B1/B2/B3
  // Coretax memakai kolom yang hampir sama, jadi jika jumlahnya seri
  // dipakai tebakan dari kolom pembedanya.
  const guess = guessModeId(headers);
  const better = MODES.filter((m) => m !== state.mode)
    .map((m) => ({ m, n: state.rawRows.filter((row) => extractRecord(m, row).valid).length }))
    .sort((a, b) => b.n - a.n || (b.m.id === guess) - (a.m.id === guess))[0];

  el.formatWarning.innerHTML = `
    ${icon("i-alert")}
    <div class="notice-body">
      <strong>Tidak ada faktur yang bisa diproses.</strong>
      ${
        missing.length
          ? `Kolom berikut tidak ditemukan di file: <strong>${esc(missing.join(", "))}</strong>.`
          : "Semua baris datanya tidak lengkap."
      }
      ${
        better && better.n > 0
          ? `<div style="margin-top:8px">File ini sepertinya untuk <strong>${esc(better.m.title)}</strong> (${better.n} faktur cocok).
             <button type="button" class="btn btn-ghost btn-sm" data-switch="${better.m.id}" style="margin-left:6px">Ganti ke ${esc(better.m.title.replace("Faktur Pajak ", ""))}</button></div>`
          : `<div style="margin-top:4px">Pastikan Anda memilih jenis faktur yang benar dan file berasal dari unduhan DJP/Coretax.</div>`
      }
    </div>`;
  el.formatWarning.classList.remove("hidden");
}

// Tebakan lampiran SPT dari judul kolom (A2: pembeli, B1: ada Keterangan,
// B2: ada Kode Transaksi, B3: tanpa keduanya). Hanya untuk saran.
function guessModeId(headers) {
  if (headers.has("namapembeli")) return "keluaran";
  if (headers.has("keterangan")) return "b1";
  if (headers.has("kodetransaksi")) return "b2";
  if (headers.has("namapenjual")) return "b3";
  return null;
}

function clearFile() {
  state.file = null;
  state.rawRows = [];
  state.records = [];
  state.autoPeriod = false;
  resetPeriods();
  el.fileInput.value = "";
  hideResult();
  updateAll();
}

// =================================================================
// RINGKASAN & TOMBOL KONVERSI
// =================================================================
function readiness(groups) {
  const accounts = accountValues();
  const accTotal = Object.keys(accounts).length;
  const accFilled = Object.values(accounts).filter(Boolean).length;
  const validCount = state.records.filter((r) => r.valid).length;
  const [first] = groups;

  return [
    { ok: true, label: "Jenis faktur", sub: state.mode.title, step: "step1" },
    {
      ok: true,
      label: "Periode",
      sub: !first
        ? "Dari Masa Pajak · menunggu faktur valid"
        : groups.length > 1
        ? `Dari Masa Pajak · ${groups.map((g) => periodLabel(g.key)).join(", ")}`
        : `${periodLabel(first.key)}${state.autoPeriod ? " (Masa Pajak)" : ""} · mulai ${jvNumber(
            state.mode, first.key, first.start
          )}`,
      step: "step2",
    },
    {
      ok: accFilled === accTotal,
      label: "Akun Accurate",
      sub: accFilled === accTotal ? "Semua akun terisi" : `${accFilled} dari ${accTotal} akun terisi`,
      step: "step3",
    },
    {
      ok: validCount > 0,
      label: "File Excel",
      sub: !state.file
        ? "Belum ada file"
        : validCount
        ? `${validCount} faktur siap diproses`
        : "Tidak ada faktur yang valid",
      step: "step4",
    },
  ];
}

function updateAll() {
  const groups = periodGroups();
  renderPeriods(groups);
  updateChipsAndSchema();
  renderFile(jvNumbers(groups));

  const items = readiness(groups);
  el.checklist.innerHTML = items
    .map(
      (it) => `<li class="${it.ok ? "ok" : ""}"><span class="tick">${icon("i-check")}</span>
        <span>${esc(it.label)}<span class="sub">${esc(it.sub)}</span></span></li>`
    )
    .join("");
  items.forEach((it) => $(it.step).classList.toggle("done", it.ok));

  const ready = items.every((it) => it.ok);
  el.convertButton.disabled = !ready;
  const firstMissing = items.find((it) => !it.ok);
  el.convertHint.textContent = ready
    ? "Semua siap. Klik untuk membuat file XML."
    : firstMissing.step === "step3"
    ? "Isi semua nomor akun di langkah 3."
    : "Unggah file Excel di langkah 4.";
  el.convertHint.classList.toggle("hidden", !el.resultCard.classList.contains("hidden"));
}

function convert() {
  const accounts = accountValues();
  const groups = periodGroups();
  const valid = state.records.filter((r) => r.valid);
  if (!valid.length) return;

  const xml = buildXml(state.mode, state.records, accounts, jvNumbers(groups));
  const keys = groups.map((g) => g.key);
  const span = keys.length > 1 ? `${keys[0]}-${keys[keys.length - 1]}` : keys[0];
  const filename = `${state.mode.category}_${span}_${Date.now()}.xml`;

  el.xmlOutput.textContent = xml;
  if (state.downloadUrl) URL.revokeObjectURL(state.downloadUrl);
  state.downloadUrl = URL.createObjectURL(new Blob([xml], { type: "application/xml" }));
  el.downloadLink.href = state.downloadUrl;
  el.downloadLink.download = filename;
  el.resultText.textContent = `${valid.length} jurnal · ${groups.map(jvRange).join("; ")}`;
  el.resultCard.classList.remove("hidden");
  el.convertHint.classList.add("hidden");

  for (const g of groups) {
    const last = g.start + g.count - 1;
    store(lastIndexKey(g.key), Math.max(last, load(lastIndexKey(g.key), 0)));
  }
  saveToHistory(filename, xml);
  toast(`${valid.length} jurnal berhasil dibuat.`, "success");
  el.resultCard.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

function hideResult() {
  el.resultCard.classList.add("hidden");
  el.convertHint.classList.remove("hidden");
}

// =================================================================
// RIWAYAT (disimpan di browser ini dengan IndexedDB)
// =================================================================
const historyDb = (() => {
  let opening = null;
  function open() {
    if (!opening) {
      opening = new Promise((resolve, reject) => {
        if (!window.indexedDB) return reject(new Error("IndexedDB tidak didukung"));
        const req = indexedDB.open("fpconv", 1);
        req.onupgradeneeded = () => {
          const s = req.result.createObjectStore("history", { keyPath: "id", autoIncrement: true });
          s.createIndex("category", "category");
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      });
    }
    return opening;
  }
  async function run(mode, fn) {
    const db = await open();
    return new Promise((resolve, reject) => {
      const tx = db.transaction("history", mode);
      const req = fn(tx.objectStore("history"));
      tx.oncomplete = () => resolve(req.result);
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error);
    });
  }
  return {
    add: (item) => run("readwrite", (s) => s.add(item)),
    list: (category) => run("readonly", (s) => s.index("category").getAll(category)),
    get: (id) => run("readonly", (s) => s.get(id)),
    remove: (id) => run("readwrite", (s) => s.delete(id)),
  };
})();

function setStorageStatus(ok) {
  state.storageOk = ok;
  el.serverStatus.classList.toggle("ok", ok);
  el.serverStatus.classList.toggle("off", !ok);
  el.serverStatus.querySelector(".label").textContent = ok ? "Riwayat di browser ini" : "Riwayat nonaktif";
  el.serverStatus.title = ok
    ? "Setiap file XML otomatis disimpan di browser ini saja. Data tidak dikirim ke mana pun."
    : "Browser tidak mengizinkan penyimpanan (mis. mode penyamaran). Konversi tetap bisa, tetapi riwayat tidak tersimpan.";
}

async function saveToHistory(filename, content) {
  if (!state.storageOk) return;
  try {
    await historyDb.add({
      category: state.mode.category,
      filename,
      xml_content: content,
      created_at: new Date().toISOString(),
    });
    loadHistory();
  } catch (err) {
    console.error("Gagal menyimpan riwayat:", err);
    toast("File XML dibuat, tetapi gagal disimpan ke riwayat.", "error");
  }
}

async function loadHistory() {
  const category = state.mode.category;
  try {
    const data = await historyDb.list(category);
    setStorageStatus(true);
    if (category !== state.mode.category) return;
    renderHistory(data.sort((a, b) => b.id - a.id));
  } catch (err) {
    console.error("Gagal memuat riwayat:", err);
    setStorageStatus(false);
    el.historyList.innerHTML = `<li class="empty muted small">
      ${icon("i-info")}<br />
      Riwayat tidak tersedia karena browser tidak mengizinkan penyimpanan
      (misalnya mode penyamaran/incognito).<br />Konversi tetap dapat digunakan.</li>`;
  }
}

function renderHistory(data) {
  if (!data.length) {
    el.historyList.innerHTML = `<li class="empty muted small">Belum ada file untuk jenis faktur ini.</li>`;
    return;
  }
  el.historyList.innerHTML = "";
  for (const item of data) {
    // Nama file: <kategori>_<YYMM>[-<YYMM>]_<waktu>.xml
    const m = /_(\d{4})(?:-(\d{4}))?_\d+\.xml$/.exec(item.filename);
    const title = m
      ? `Periode ${periodLabel(m[1])}${m[2] ? ` – ${periodLabel(m[2])}` : ""}`
      : item.filename;
    const when = new Date(item.created_at).toLocaleString("id-ID", {
      day: "numeric", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit",
    });

    const li = document.createElement("li");
    li.innerHTML = `
      <div class="h-meta">
        <div class="h-name">${esc(title)}</div>
        <div class="h-file" title="${esc(item.filename)}">Dibuat ${esc(when)}</div>
      </div>
      <div class="h-actions">
        <button class="icon-btn" type="button" data-act="download" title="Unduh">${icon("i-download")}</button>
        <button class="icon-btn danger" type="button" data-act="delete" title="Hapus">${icon("i-trash")}</button>
      </div>`;
    li.querySelector('[data-act="download"]').addEventListener("click", () => downloadFromHistory(item.id, item.filename));
    li.querySelector('[data-act="delete"]').addEventListener("click", () => deleteHistory(item.id, title, when));
    el.historyList.appendChild(li);
  }
}

async function downloadFromHistory(id, filename) {
  try {
    const item = await historyDb.get(id);
    if (!item) throw new Error("File tidak ditemukan");
    const url = URL.createObjectURL(new Blob([item.xml_content], { type: "application/xml" }));
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 100);
  } catch (err) {
    toast("Gagal mengunduh file: " + err.message, "error");
  }
}

async function deleteHistory(id, title, when) {
  const ok = await confirmDialog(
    "Hapus dari riwayat?",
    `${title} (dibuat ${when}) akan dihapus permanen dari riwayat. File yang sudah Anda unduh tidak terpengaruh.`
  );
  if (!ok) return;
  try {
    await historyDb.remove(id);
    toast("File dihapus dari riwayat.", "success");
    loadHistory();
  } catch (_) {
    toast("Gagal menghapus data.", "error");
  }
}

// =================================================================
// EVENT
// =================================================================
function bindEvents() {
  el.accountFields.addEventListener("input", onAccountChange);
  el.accountFields.addEventListener("click", (e) => {
    const chip = e.target.closest("[data-fill]");
    if (!chip) return;
    $(`acc-${chip.dataset.fill}`).value = chip.dataset.value;
    onAccountChange();
  });

  el.fileInput.addEventListener("change", () => handleFile(el.fileInput.files[0]));
  el.changeFile.addEventListener("click", () => {
    clearFile();
    el.fileInput.click();
  });
  ["dragenter", "dragover"].forEach((ev) =>
    el.dropzone.addEventListener(ev, (e) => {
      e.preventDefault();
      el.dropzone.classList.add("drag");
    })
  );
  ["dragleave", "drop"].forEach((ev) =>
    el.dropzone.addEventListener(ev, (e) => {
      e.preventDefault();
      el.dropzone.classList.remove("drag");
    })
  );
  el.dropzone.addEventListener("drop", (e) => handleFile(e.dataTransfer.files[0]));
  // Cegah browser membuka file jika dilepas di luar area
  window.addEventListener("dragover", (e) => e.preventDefault());
  window.addEventListener("drop", (e) => e.preventDefault());

  el.rowFilter.addEventListener("click", (e) => {
    const b = e.target.closest("[data-filter]");
    if (!b) return;
    state.filter = b.dataset.filter;
    renderFile();
  });
  el.formatWarning.addEventListener("click", (e) => {
    const b = e.target.closest("[data-switch]");
    if (b) location.hash = b.dataset.switch;
  });

  el.convertButton.addEventListener("click", convert);

  $("helpButton").addEventListener("click", () => $("helpDialog").showModal());
  document.querySelectorAll("dialog [data-close]").forEach((b) =>
    b.addEventListener("click", () => b.closest("dialog").close())
  );
  document.querySelectorAll("dialog").forEach((d) =>
    d.addEventListener("click", (e) => {
      if (e.target === d) d.close();
    })
  );

  window.addEventListener("hashchange", () => setMode(location.hash.slice(1)));
}

// =================================================================
// MULAI
// =================================================================
renderModes();
renderPeriod();
bindEvents();
const initial = location.hash.slice(1) || load("lastMode", "keluaran");
state.mode = null;
if (location.hash.slice(1) !== initial) history.replaceState(null, "", "#" + initial);
setMode(initial);
if (!load("seenHelp", false)) {
  store("seenHelp", true);
  $("helpDialog").showModal();
}
