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
- GitHub-Repo `3d-space` (dfedorov12) anlegen + GitHub Pages (Source: Actions) aktivieren. ✓ erledigt
- `AppPermissions`-Zeilen für die gewünschten Mail-Adressen anlegen.

## 2026-06-16 — Login-Fix + Feature-Ausbau
- Login von Popup auf **Redirect** umgestellt (loginRedirect/logoutRedirect) → kein Popup-Blocker mehr.
- **Feature-Ausbau** (alles ohne eigene Server-Infra):
  - 🖥 Bildschirm teilen → erscheint als VideoTexture auf der Wand-Leinwand (separater PeerJS-Call mit metadata.kind='screen'); Presenter-Banner; Neuzugänge bekommen laufende Übertragung automatisch.
  - 🟢 Verbindungs-Ampel pro Person (ICE-State der Voice-Verbindung).
  - 👆 Push-to-Talk (Leertaste), 🎤 Mikrofon-Auswahl (replaceTrack ohne Renegotiation; aktiviert auch nachträglich Mikro bei „nur zuhören").
  - 💬 Text-Chat über Daten-Kanal, mit Ungelesen-Badge.
  - 📱 Mobiler Touch-Joystick, ⚙️ Performance-Sparmodus (Schatten/PixelRatio).
  - 🌐 TURN/ICE-Block in `app.js` vorbereitet (Platzhalter), STUN aktiv.
- **Auto-Cache-Busting** im Deploy-Workflow: `sed` ersetzt `?v=…` durch Commit-SHA beim Build (kein Rück-Commit, keine Endlosschleife). Manuelles Hochzählen entfällt.
- Verifiziert per Preview (Eval/Pixel-Readback): Szene + alle neuen Funktionen/IDs vorhanden, Toggles fehlerfrei, Screen-Share-Texturwechsel + Banner ok, keine Konsolenfehler.

### NICHT umgesetzt (brauchen Server-Infrastruktur)
- SFU-Umstieg (für >8 Personen) und echte WebRTC-Authentifizierung der Peers.
