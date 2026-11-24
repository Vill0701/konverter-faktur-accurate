# Konverter Faktur Pajak → Accurate 5

Mengubah file Excel faktur pajak dan lampiran SPT Masa PPh 21 (unduhan DJP/Coretax) menjadi file XML jurnal
umum yang bisa diimpor ke **Accurate 5**.

Aplikasi ini adalah **web statis**: tidak butuh server maupun database. Semua data diproses di browser pengguna
dan tidak dikirim ke mana pun. Riwayat file XML tersimpan di browser masing-masing (IndexedDB).

Jenis data yang didukung:

| Jenis | Untuk | Nomor jurnal |
| --- | --- | --- |
| Keluaran A2 | Penjualan | `PENJ.YYYY.MM.001` |
| Masukan B1 | Impor (PIB) | `PIB.YYYY.MM.001` |
| Masukan B2 | Pembelian dalam negeri, PPN dikreditkan | `PEMB.YYYY.MM.001` |
| Masukan B3 | Pembelian, PPN tidak dikreditkan | `PEMB.NON.YYYY.MM.001` |
| PPh 21 Gaji Karyawan | Biaya gaji, hutang gaji & hutang PPh 21 (lampiran SPT Masa PPh 21) | `GAJI.YYYY.MM.001` |

`YYYY.MM` diambil otomatis dari kolom **Masa Pajak** di Excel (mis. `Januari 2026` → `2026.01`), nomor urut
(`001`) diisi pengguna. Jika satu file berisi beberapa masa pajak, nomor urut dihitung terpisah per masa pajak.
File tanpa kolom Masa Pajak tetap bisa dipakai dengan memilih bulan & tahun secara manual.

Kolom Excel yang dibaca mengikuti lampiran SPT Masa PPN Coretax (A2, B1, B2, B3):
`Masa Pajak`, `Nomor Faktur/Dokumen`, `Tanggal Faktur/Dokumen`, `NPWP/NIK Pembeli` atau `NPWP/NIK Penjual`,
`Nama Pembeli` atau `Nama Penjual`, `Harga Jual / Penggantian / DPP`, dan `PPN`. Nama kolom format lama
(e-Faktur/DJP) tetap dikenali.

Format daftar faktur Keluaran/Masukan juga didukung:
`Nomor Faktur`, `Tanggal Faktur`, `NPWP Pembeli`/`NPWP Penjual`, `Nama Pembeli`/`Nama Penjual`,
`Kode Transaksi`, `Harga Jual / Penggantian`, `DPP Nilai Lain`, `PPN`, `PPnBM`, `Status`, `Referensi`
(Masukan: ditambah `Status Pengkreditan` dan `Masa Pengkreditan`). DPP jurnal diambil dari
`Harga Jual / Penggantian`; kolom `DPP Nilai Lain`, `PPnBM`, `Status`, `Referensi`, `Status Pengkreditan`,
dan `Masa Pengkreditan` tidak dipakai. Format ini tidak punya kolom Masa Pajak, jadi bulan & tahun jurnal
dipilih manual (terisi otomatis dari bulan terbanyak pada Tanggal Faktur).

### Faktur retur (nilai minus)

Baris dengan **DPP atau PPN bernilai minus** dianggap retur. Saat file seperti ini diunggah, aplikasi meminta
nomor **Akun Retur Penjualan** (Keluaran A2) atau **Akun Retur Pembelian** (Masukan B1/B2/B3); file XML baru
bisa dibuat setelah akun retur diisi. Jurnal retur memakai akun retur sebagai pengganti akun Penjualan/Pembelian,
dan karena nilainya minus, sisi debit/kredit otomatis terbalik:

| Jenis | Debit | Kredit |
| --- | --- | --- |
| Retur penjualan (A2) | Retur Penjualan (DPP), PPN Keluaran (PPN) | Piutang Usaha (DPP + PPN) |
| Retur pembelian (B1/B2) | Hutang Usaha (DPP + PPN) | Retur Pembelian (DPP), PPN Masukan (PPN) |
| Retur pembelian (B3) | Hutang Usaha | Retur Pembelian |

### PPh 21 — biaya & hutang gaji

Pilih **PPh 21 Gaji Karyawan**, lalu unggah lampiran SPT Masa PPh 21 dari Coretax. Setiap bukti potong menjadi
satu jurnal:

| Akun | Debit | Kredit |
| --- | --- | --- |
| Biaya Gaji | Penghasilan Bruto | |
| Hutang PPh 21 | | PPh |
| Hutang Gaji | | Penghasilan Bruto − PPh |

- Kolom yang dibaca: `Nomor Bukti Potong`/`WithholdingNumber`, `Tanggal Bukti Potong`/`WithholdingDate`,
  `Nama Penerima Penghasilan`, `Penghasilan Bruto`, `PPh`, serta `Masa Pajak` untuk `YYYY.MM` nomor jurnal.
  `NPWP/NIK Penerima Penghasilan` dan `Kode Objek Pajak` boleh kosong.
- **Semua sheet** yang memuat kolom tersebut dibaca sekaligus, mis. `Tabel 1A` (pegawai tetap) dan
  `Lampiran 3 - BP21` (bukan pegawai, mis. tenaga ahli); keduanya memakai akun yang sama. Sheet lain
  (`Info`, sheet berisi "Tidak ada data") dan baris `TOTAL` di akhir tabel diabaikan.
- Baris Hutang PPh 21 hanya dibuat jika PPh-nya lebih dari 0 (penghasilan dengan tarif TER 0% cukup dua baris).
- Tanggal jurnal = Tanggal Bukti Potong. Keterangan jurnal memakai sebutan sesuai Kode Objek Pajak:
  `Gaji` (21-100-01), `Pensiun` (21-100-02), `Upah` (21-100-03), `Honorarium` (21-100-07), selain itu `Imbalan`,
  mis. `Gaji BUDI SANTOSO - Januari 2026 - <nomor bukti potong>`.

## Cara pakai

1. Pilih jenis data (jenis faktur atau PPh 21).
2. Isi nomor urut awal jurnal (bulan & tahun terisi otomatis dari Masa Pajak setelah file diunggah).
3. Isi nomor akun sesuai daftar akun di database Accurate Anda (isian diingat otomatis).
4. Unggah file Excel, periksa tabel pratinjau (baris yang dilewati ditandai beserta alasannya).
5. Klik **Buat File XML**, unduh, lalu impor ke Accurate 5.

> Riwayat hanya ada di browser tempat file dibuat dan akan hilang jika data browser dibersihkan.
> Simpan file XML penting di folder Anda sendiri. Tombol **Hapus semua** di kartu Riwayat menghapus seluruh
> riwayat jenis data yang sedang dipilih (bisa dicentang untuk ikut menghapus riwayat jenis data lain).

## Deploy ke GitHub Pages (gratis)

1. Buat repository baru di GitHub (harus **Public** untuk GitHub Pages gratis).
2. Unggah isi folder ini ke root repository (`index.html` harus berada di paling atas).
   File `database_faktur.db` sudah dikecualikan lewat `.gitignore` — **jangan diunggah** karena berisi data pajak.
3. Buka **Settings → Pages**, pada *Build and deployment* pilih **Deploy from a branch**,
   branch `main`, folder `/ (root)`, lalu **Save**.
4. Tunggu 1–2 menit. Aplikasi bisa dibuka di `https://<username>.github.io/<nama-repo>/`.

Setiap kali ada perubahan yang di-push ke branch `main`, situs akan diperbarui otomatis.

## Menjalankan tanpa internet hosting

Cukup buka `index.html` di browser (Chrome/Edge). Koneksi internet tetap diperlukan untuk memuat
pembaca Excel dan font.
