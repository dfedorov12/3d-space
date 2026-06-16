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

## 2026-06-16 — Freigabe-Fix + externe Gäste + neue Figuren
- **Freigabe** auf Code-Liste `ALLOWED` in app.js umgestellt (AppPermissions existiert nicht / war im Ticketsystem deaktiviert). administrator@ + fedorov@ freigegeben. SharePoint-Liste nur noch optionaler Zusatz.
- **Externe Gäste**: „Als Gast beitreten" (Name + `GUEST_PASSCODE`, Standard `dihag-3d`, `''`=ohne) ODER Link `?guest`. Kein Microsoft-Konto nötig. Gäste belegen Slot-IDs `…-guest-0..N` (PeerJS `unavailable-id` → nächster Slot). Discovery-Set = interne Mail-IDs ∪ Gäste-Slots → alle finden alle, serverlos.
- **Neue Figuren**: stilisierte Low-Poly-Menschen (Beine/Schuhe, Torso, Arme/Hände, Hals, Kopf mit Augen, Haare; weiblich mit Rock + langen Haaren, männlich mit Schultern/Hose). Geschlechtswahl im Mikro-Gate (`setGender`), via Metadaten + State an alle übertragen; `rebuildAvatar` bei späterer Änderung.
- Verifiziert (Eval/Pixel-Readback): Gast-UI + Code-Prüfung (leer/falsch) greifen, beide Geschlechter bauen (je 19 Meshes) + rendern (Kleidung in Personenfarbe, Haare), keine Konsolenfehler.

## 2026-06-16 — Einladen-Button + Figur-Vorschau
- **Problem:** angemeldete interne Nutzer sehen den Boot-Screen (mit „Als Gast beitreten") nie → kein Weg, Externe einzuladen.
- **Lösung:** 🔗 „Einladen"-Button in der Top-Leiste → Dialog mit Gast-Link (`…/?guest`) + Code + „Link kopieren" (Clipboard). Funktionen `openInvite/closeInvite/copyInvite/inviteLink`.
- **Figur sichtbar:** kleine, sich drehende **3D-Vorschau** (`gender-preview`, eigener Mini-Renderer `pvRenderer/pvScene/pvCam/pvAvatar`) im Mikro-Gate; `setGender` wechselt sie live; beim Beitreten `stopGenderPreview`.
- Verifiziert: Vorschau rendert opak (Kleidungsfarbe), wechselt m↔w, Auswahl markiert; Dialog zeigt Link+Code; keine Konsolenfehler.
