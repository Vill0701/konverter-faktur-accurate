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

// -----------------------------------------------------------------
// KONFIGURASI JENIS FAKTUR
// `columns` berisi nama-nama kolom Excel yang dicoba berurutan
// (kolom pertama yang ada isinya yang dipakai).
// -----------------------------------------------------------------
const MODES = [
  {
    id: "keluaran",
    category: "sales",
    icon: "i-sales",
    title: "Faktur Pajak Keluaran",
    tag: "Penjualan",
    desc: "Faktur yang Anda terbitkan saat menjual barang/jasa ke pembeli.",
    prefix: "PENJ",
    party: "Pembeli",
    parseNum: parseFloat,
    descSep: " - ",
    columns: {
      invoiceNo: ["Nomor Faktur Pajak", "Faktur Pajak/Dokumen Tertentu/Nota Retur/Nota Pembatalan - Nomor"],
      date: ["Tanggal Faktur Pajak", "Faktur Pajak/Dokumen Tertentu/Nota Retur/Nota Pembatalan - Tanggal"],
      name: ["Nama Pembeli", "Nama Pembeli BKP/Penerima Manfaat BKP Tidak Berwujud/Penerima JKP"],
      npwp: ["NPWP Pembeli / Identitas lainnya", "NPWP/NIK/Nomor Paspor"],
      dpp: ["Harga Jual/Penggantian/DPP", "Harga Jual/Penggantian/DPP (Rupiah)"],
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
      invoiceNo: ["Nomor Dokumen"],
      date: ["Tanggal Dokumen"],
      name: ["Nama Penjual"],
      npwp: ["NPWP Penjual"],
      dpp: ["DPP / Penghasilan Kotor"],
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
  const dpp = ["Harga Jual/Penggantian/DPP", "Harga Jual/Penggantian/DPP (Rupiah)", "DPP (Rupiah)"];
  if (withImportDpp) dpp.push("Harga Jual/Penggantian/Nilai Impor/DPP (Rupiah)");
  return {
    invoiceNo: [
      "Nomor Faktur Pajak",
      "Faktur Pajak/Dokumen Tertentu/Nota Retur/Nota Pembatalan - Nomor",
      "Dokumen Tertentu - Nomor",
    ],
    date: [
      "Tanggal Faktur Pajak",
      "Faktur Pajak/Dokumen Tertentu/Nota Retur/Nota Pembatalan - Tanggal",
      "Dokumen Tertentu - Tanggal",
    ],
    name: [
      "Nama Penjual",
      "Nama Penjual Barang Kena Pajak/Barang Kena Pajak Tidak Berwujud/Jasa Kena Pajak",
      "Nama Penjual BKP/BKP Tidak Berwujud/Pemberi JKP",
    ],
    npwp: ["Nomor Identitas WP", "NPWP Penjual", "NPWP"],
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
function pick(row, names) {
  for (const n of names) {
    if (row[n] !== undefined && row[n] !== null) return row[n];
  }
  return undefined;
}

function extractRecord(mode, row) {
  const c = mode.columns;
  const r = {
    invoiceNo: pick(row, c.invoiceNo),
    rawDate: pick(row, c.date),
    name: pick(row, c.name),
    npwp: pick(row, c.npwp),
    dpp: mode.parseNum(pick(row, c.dpp)),
    ppn: mode.parseNum(pick(row, c.ppn)),
  };
  const reasons = [];
  if (!r.invoiceNo) reasons.push("Nomor faktur kosong");
  if (!r.name) reasons.push(`Nama ${mode.party.toLowerCase()} kosong`);
  if (!r.npwp) reasons.push("NPWP kosong");
  if (isNaN(r.dpp)) reasons.push("DPP kosong/bukan angka");
  if (isNaN(r.ppn)) reasons.push("PPN kosong/bukan angka");
  r.valid = reasons.length === 0;
  r.reasons = reasons;
  r.total = r.valid ? r.dpp + r.ppn : NaN;
  r.date = formatDate(r.rawDate);
  return r;
}

function jvNumber(mode, year, month, index) {
  return `${mode.prefix}.${year}.${month}.${String(index).padStart(3, "0")}`;
}

function buildXml(mode, records, accounts, year, month, startIndex) {
  let transactions = "";
  let indexNum = startIndex;
  let requestId = 1;

  for (const r of records) {
    if (!r.valid) continue;
    const total = r.dpp + r.ppn;
    transactions += `
        <JV operation="Add" REQUESTID="${requestId}">
            <TRANSACTIONID>148</TRANSACTIONID>
            ${mode.lines(r, accounts, total)}
            <JVNUMBER>${jvNumber(mode, year, month, indexNum)}</JVNUMBER>
            <TRANSDATE>${r.date}</TRANSDATE>
            <SOURCE>GL</SOURCE>
            <TRANSTYPE>journal voucher</TRANSTYPE>
            <TRANSDESCRIPTION>${r.name}${mode.descSep}${r.npwp}${mode.descSep}${r.invoiceNo}</TRANSDESCRIPTION>
            <JVAMOUNT>${total}</JVAMOUNT>
        </JV>`;
    indexNum++;
    requestId++;
  }

  return `<?xml version="1.0"?>\n<NMEXML EximID="1" BranchCode="${BRANCH_CODE}" ACCOUNTANTCOPYID="">\n<TRANSACTIONS OnError="CONTINUE">\n${transactions}\n</TRANSACTIONS>\n</NMEXML>`;
}

function formatDate(serial) {
  if (typeof serial === "string" && (serial.includes("-") || serial.includes("/"))) {
    const d = new Date(serial);
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
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

function period() {
  return { year: el.year.value, month: el.month.value };
}

function periodLabel(yy, mm) {
  return `${MONTHS[Number(mm) - 1] ?? mm} 20${yy}`;
}

function startIndex() {
  const n = parseInt(el.startIndex.value, 10);
  return Number.isFinite(n) && n > 0 ? n : 1;
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
}

function lastIndexKey() {
  const { year, month } = period();
  return `lastIndex.${state.mode.category}.${year}${month}`;
}

function updatePeriodPreview() {
  const { year, month } = period();
  el.jvPreview.textContent = jvNumber(state.mode, year, month, startIndex());

  const last = load(lastIndexKey(), null);
  if (last && startIndex() <= last) {
    const next = last + 1;
    el.continueIndex.dataset.next = next;
    el.continueIndex.textContent = `Lanjutkan dari ${String(next).padStart(3, "0")} (terakhir dipakai ${String(last).padStart(3, "0")})`;
    el.continueIndex.classList.remove("hidden");
  } else {
    el.continueIndex.classList.add("hidden");
  }
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
      return `<tr>
        <td>${d ? "" : "&emsp;&emsp;"}${acc ? `<strong>${esc(acc)}</strong>` : `<span class="muted">[belum diisi]</span>`} <span class="muted">${esc(a.label.replace("Akun ", ""))}</span></td>
        <td class="amt d">${d ? esc(a.amount) : ""}</td>
        <td class="amt c">${d ? "" : esc(a.amount)}</td>
      </tr>`;
    })
    .join("");
  el.journalSchema.innerHTML = `
    <div class="schema-title">Bentuk jurnal yang akan dibuat untuk setiap faktur</div>
    <table>
      <tr><td class="muted small">Akun</td><td class="amt muted small">Debit</td><td class="amt muted small">Kredit</td></tr>
      ${rows}
    </table>`;
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
  const rows = Object.entries(state.mode.columns)
    .map(([k, names]) => `<tr><td>${labels[k]}</td><td>${names.map((n) => `<code>${esc(n)}</code>`).join("")}</td></tr>`)
    .join("");
  el.columnsHelpBody.innerHTML = `
    <p class="muted">Baris pertama (judul kolom) di Excel harus memuat salah satu nama berikut:</p>
    <table>${rows}</table>`;
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
      state.rawRows = XLSX.utils.sheet_to_json(worksheet);
      state.file = file;
      state.filter = "all";
      hideResult();
      processRows();
      autoDetectPeriod();
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

// Jika semua faktur ada di satu bulan yang sama dan user belum memilih
// periode sendiri, isi periode otomatis sesuai tanggal faktur.
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
  toast(`Periode disesuaikan ke ${periodLabel(yy, mm)} sesuai tanggal faktur.`, "info");
}

function renderFile() {
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
  const { year, month } = period();
  let jvIndex = startIndex();
  const rows = [];
  state.records.forEach((r, i) => {
    const jv = r.valid ? jvNumber(state.mode, year, month, jvIndex++) : null;
    if (state.filter === "valid" && !r.valid) return;
    if (state.filter === "skip" && r.valid) return;
    if (rows.length >= 1000) return;
    rows.push(`
      <tr class="${r.valid ? "" : "skip"}">
        <td class="muted">${i + 2}</td>
        <td class="name">${esc(r.invoiceNo ?? "–")}<div class="npwp">${displayDate(r.date)}</div></td>
        <td class="name">${esc(r.name ?? "–")}<div class="npwp">${esc(r.npwp ?? "")}</div></td>
        <td class="num hide-sm">${fmtNum(r.dpp)}</td>
        <td class="num hide-sm">${fmtNum(r.ppn)}</td>
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
    .filter(([, names]) => !names.some((n) => headers.has(n)))
    .map(([k]) => labels[k]);

  // Cari jenis faktur lain yang cocok dengan file ini
  const better = MODES.filter((m) => m !== state.mode)
    .map((m) => ({ m, n: state.rawRows.filter((row) => extractRecord(m, row).valid).length }))
    .sort((a, b) => b.n - a.n)[0];

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

function clearFile() {
  state.file = null;
  state.rawRows = [];
  state.records = [];
  el.fileInput.value = "";
  hideResult();
  updateAll();
}

// =================================================================
// RINGKASAN & TOMBOL KONVERSI
// =================================================================
function readiness() {
  const { year, month } = period();
  const accounts = accountValues();
  const accTotal = Object.keys(accounts).length;
  const accFilled = Object.values(accounts).filter(Boolean).length;
  const validCount = state.records.filter((r) => r.valid).length;

  return [
    { ok: true, label: "Jenis faktur", sub: state.mode.title, step: "step1" },
    {
      ok: !!year && !!month,
      label: "Periode",
      sub: `${periodLabel(year, month)} · mulai ${jvNumber(state.mode, year, month, startIndex())}`,
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
  updatePeriodPreview();
  updateChipsAndSchema();
  renderFile();

  const items = readiness();
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
  const { year, month } = period();
  const accounts = accountValues();
  const start = startIndex();
  const valid = state.records.filter((r) => r.valid);
  if (!valid.length) return;

  const xml = buildXml(state.mode, state.records, accounts, year, month, start);
  const filename = `${state.mode.category}_${year}${month}_${Date.now()}.xml`;
  const last = start + valid.length - 1;

  el.xmlOutput.textContent = xml;
  if (state.downloadUrl) URL.revokeObjectURL(state.downloadUrl);
  state.downloadUrl = URL.createObjectURL(new Blob([xml], { type: "application/xml" }));
  el.downloadLink.href = state.downloadUrl;
  el.downloadLink.download = filename;
  el.resultText.textContent = `${valid.length} jurnal · ${jvNumber(state.mode, year, month, start)} s/d ${jvNumber(
    state.mode, year, month, last
  )}`;
  el.resultCard.classList.remove("hidden");
  el.convertHint.classList.add("hidden");

  store(lastIndexKey(), Math.max(last, load(lastIndexKey(), 0)));
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
    const m = /_(\d{2})(\d{2})_\d+\.xml$/.exec(item.filename);
    const title = m ? `Periode ${periodLabel(m[1], m[2])}` : item.filename;
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
