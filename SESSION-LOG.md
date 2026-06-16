# Session-Log · 3D-Space

## 2026-06-16 — Projekt angelegt
- Neues Projekt `3d-space` nach dem DIHAG-App-Muster (Static Site + GitHub Pages + MSAL).
- Dateien: `index.html`, `app.js`, `style.css`, `.github/workflows/jekyll-gh-pages.yml`, `README.md`.
- **Zugriff**: MSAL-Login + Allowlist über SharePoint-Liste `AppPermissions` (`App = 3d-space`).
  Allowlist dient gleichzeitig als Teilnehmer-Verzeichnis.
- **3D-Raum**: Three.js (CDN) – Tisch, Stühle, Fensterfront mit Skyline, Pflanzen, Wandbildschirm,
  Deckenleuchten, RoomEnvironment-Reflexionen.
- **Sprache**: serverloses PeerJS-Mesh; Peer-IDs aus den freigegebenen Mails abgeleitet.
- **Räumliches Audio**: `PositionalAudio` (näher = lauter), Sprech-Indikator + Melden + Mikro-Stumm.
- Bewegung per WASD, Third-Person-Kamera (OrbitControls), Übersicht-Modus.

### Offen / Setup durch IT
- Azure: Redirect-URI `https://dfedorov12.github.io/3d-space/` zur App-Registrierung hinzufügen.
- GitHub-Repo `3d-space` (dfedorov12) anlegen + GitHub Pages (Source: Actions) aktivieren.
- `AppPermissions`-Zeilen für die gewünschten Mail-Adressen anlegen.
