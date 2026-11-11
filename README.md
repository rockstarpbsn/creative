# Creative Request Desk

Web statis (GitHub Pages) + Firebase Firestore + Firebase Auth (Google) untuk supervisor.

## Setup Firebase
1. https://console.firebase.google.com → Add project.
2. Build → **Firestore Database** → Create database (production mode, pilih region asia-southeast2/Jakarta).
3. Build → **Authentication** → Get started → Sign-in method → aktifkan **Google**.
4. Authentication → Settings → **Authorized domains** → tambahkan `USERNAME.github.io`.
5. Project settings (ikon roda gigi) → Your apps → Web `</>` → daftar app → salin `firebaseConfig`.
6. Tempel config ke `index.html` (bagian `firebaseConfig`) dan isi `ADMIN_EMAILS` dengan email Google supervisor.
7. Firestore → **Rules** → tempel isi `firestore.rules` (ganti `EMAIL_SUPERVISOR@gmail.com` dengan email yang sama) → Publish.

## Deploy ke GitHub Pages
1. Buat repo baru, upload `index.html` (dan file lain di folder ini).
2. Settings → Pages → Source: *Deploy from a branch* → `main` / root → Save.
3. Buka `https://USERNAME.github.io/NAMA-REPO/`.

Catatan: `apiKey` Firebase memang boleh publik; keamanan data diatur oleh `firestore.rules`.
