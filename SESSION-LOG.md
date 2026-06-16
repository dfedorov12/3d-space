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

## 2026-06-16 — Getrennte-Räume-Bug + Räume + Tischtennis
- **Bug 1 (Hauptursache getrennte Räume):** Peer-ID-Präfix war `dihag3dspace--` → der PeerJS-Broker lehnt IDs mit `--` ab. Auf einzelne Bindestriche umgestellt (`dihag3dspace-`).
- **Bug 2:** Gäste bekamen die interne Adressliste nie (afterLogin wird für Gäste nicht aufgerufen) → suchten nie nach internen Teilnehmern. In `guestJoin` jetzt `roster` aus ALLOWED+SUPER_ADMINS befüllt.
- **Räume:** `ROOMS` (hauptraum/pingpong), Peer-IDs jetzt RAUM-bezogen (`uidFor`/`guestSlotId` enthalten roomKey) → man trifft nur Leute im selben Raum. Raumwechsel-Schalter in der Top-Leiste, `switchRoom` (trennt Peers, baut Szene um via Gruppen `gConf`/`gPing`, vernetzt im neuen Raum neu). Szene = gemeinsame Hülle (Sky/Wände/Licht) + umschaltbare Möbelgruppen.
- **Tischtennis:** eigener Raum mit Tisch/Netz/Linien; jeder Avatar hat einen Schläger (nur im Raum sichtbar). Ball hub-los, **Autorität = kleinste Peer-ID** im Raum: simuliert Schwerkraft + Tisch-/Schläger-Abprall + Aufschlag-Reset, sendet Ballposition ~20/s (`{t:'ball'}`); andere lerpen dazu. Ballwechsel-Zähler (`rally`). Spielen = Avatar zum Ball laufen.
- Verifiziert (Eval/Pixel): IDs ohne `--` & raum-bezogen, Räume schalten Sichtbarkeit korrekt, Ball bewegt sich & bleibt endlich (kein NaN), Schläger/Ball sichtbar nur im Tischtennis-Raum, Tisch rendert, keine Konsolenfehler.

## 2026-06-16 — Phantom-Teilnehmer-Bug ("12× Admin DIHAG")
- **Bug:** `bindCall` legte SOFORT bei jedem ausgehenden Anruf einen Teilnehmer+Avatar an – auch zu offline-Zielen (fedorov + 10 Gäste-Slots) → ~12 Phantome, alle „verbindet…". Name war fälschlich der EIGENE, weil bei ausgehenden Calls `call.metadata` = die eigenen Metadaten sind.
- **Fix:** Teilnehmer entstehen nur noch bei ECHTER Verbindung — `bindCall` ruft `ensurePeer` erst in `call.on('stream')`; `bindDataConn` erst in `conn.on('open')`. Neuer Parameter `inbound`: nur bei EINGEHENDEN Verbindungen werden Metadaten als Name genutzt; bei ausgehenden kommt der Name per Datenpaket (sonst E-Mail-Fallback). Zusätzlich `connectPeer`-Guard: bereits verbundene Peers nicht erneut anrufen (keine Doppelverbindungen/Rescan-Stau).
- Verifiziert: ausgehender Call/Datenconn zu offline → 0 Teilnehmer; Daten-`open` → genau 1 (Name aus E-Mail, nicht „Admin DIHAG"); eingehend → Name der Gegenseite; keine Konsolenfehler.
