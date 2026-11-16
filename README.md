# Ujian Syafawi Al-Quran

Laman `index.html` (GitHub Pages) membaca & menyimpan markah terus ke Google Sheet melalui Apps Script.

## Sambungan ke Google Sheet (sekali sahaja)
1. Buka spreadsheet → **Extensions → Apps Script**.
2. Padam kod sedia ada, tampal isi `apps-script/Code.gs`, tukar `PASSCODE` kepada kod rahsia awak.
3. **Deploy → New deployment → Web app**: *Execute as* = **Me**, *Who has access* = **Anyone**. Benarkan akses bila diminta.
4. Salin URL Web App (berakhir `/exec`), tampal pada `const API=''` dalam `index.html`, commit & push.
5. Aktifkan **Settings → Pages** pada repo (branch ikut yang dipilih, folder root).
6. Buka laman, masukkan kod akses. Markah yang disimpan terus masuk ke sheet `ISI MARKAH` (kolum D:K, M:N, Q); formula jumlah kekal.

Selepas ubah `Code.gs`, deploy semula (**Manage deployments → Edit → New version**).
