# Konverter Faktur Pajak → Accurate 5

Mengubah file Excel faktur pajak (unduhan DJP/Coretax) menjadi file XML jurnal umum yang bisa diimpor ke **Accurate 5**.

Aplikasi ini adalah **web statis**: tidak butuh server maupun database. Semua data diproses di browser pengguna
dan tidak dikirim ke mana pun. Riwayat file XML tersimpan di browser masing-masing (IndexedDB).

Jenis faktur yang didukung:

| Jenis | Untuk | Nomor jurnal |
| --- | --- | --- |
| Keluaran A2 | Penjualan | `PENJ.YY.MM.001` |
| Masukan B1 | Impor (PIB) | `PIB.YY.MM.001` |
| Masukan B2 | Pembelian dalam negeri, PPN dikreditkan | `PEMB.YY.MM.001` |
| Masukan B3 | Pembelian, PPN tidak dikreditkan | `PEMB.NON.YY.MM.001` |

`YY.MM` diambil otomatis dari kolom **Masa Pajak** di Excel (mis. `Januari 2026` → `26.01`), nomor urut
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

## Cara pakai

1. Pilih jenis faktur.
2. Isi nomor urut awal jurnal (bulan & tahun terisi otomatis dari Masa Pajak setelah file diunggah).
3. Isi nomor akun sesuai daftar akun di database Accurate Anda (isian diingat otomatis).
4. Unggah file Excel, periksa tabel pratinjau (baris yang dilewati ditandai beserta alasannya).
5. Klik **Buat File XML**, unduh, lalu impor ke Accurate 5.

> Riwayat hanya ada di browser tempat file dibuat dan akan hilang jika data browser dibersihkan.
> Simpan file XML penting di folder Anda sendiri. Tombol **Hapus semua** di kartu Riwayat menghapus seluruh
> riwayat jenis faktur yang sedang dipilih (bisa dicentang untuk ikut menghapus riwayat jenis faktur lain).

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
