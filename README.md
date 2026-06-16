# DIHAG · 3D-Space 🪑

Virtueller 3D-Konferenzraum für DIHAG. Freigegebene Personen treffen sich als Avatare in einem
schönen Konferenzraum und **sprechen per Sprachverbindung** miteinander – mit **räumlichem Audio**
(näher = lauter). Reine Static-Site auf GitHub Pages, **ohne eigenen Server**.

## Wie es funktioniert

| Baustein | Technik |
|---|---|
| Anmeldung | MSAL gegen den DIHAG-Tenant (gleiche App-Registrierung wie die übrigen Apps) |
| Freigabe | SharePoint-Liste **`AppPermissions`** auf `/sites/ticket`, Spalte `App = 3d-space` |
| 3D-Raum | Three.js (CDN) – prozedural: Tisch, Stühle, Fensterfront, Skyline, Pflanzen, Bildschirm |
| Sprache | PeerJS-WebRTC-Mesh (serverlos, öffentlicher Broker) |
| Räuml. Audio | Web Audio `PositionalAudio` – Lautstärke nach Distanz |

Da GitHub Pages rein statisch ist, gibt es **keinen Signaling-Server**. Trick: Die Peer-IDs werden
**deterministisch aus den freigegebenen Mail-Adressen** abgeleitet. Jeder Client kennt die Allowlist
(= Teilnehmer-Verzeichnis) und ruft alle Erlaubten direkt an. Wer offline ist, antwortet einfach nicht.

## Bedienung

- **WASD / Pfeiltasten** (oder Touch-Joystick am Handy) – im Raum gehen
- **Maus ziehen** – umsehen · **Scrollen** – Zoom
- Näher an eine Person gehen = sie wird lauter
- Untere Leiste:
  - 🎤 **Mikro** stummschalten · ▾ **Mikrofon wählen**
  - 👆 **Push-to-Talk** (dann Leertaste halten zum Sprechen)
  - ✋ **Melden** · 🖥 **Bildschirm teilen** (erscheint auf der Wand-Leinwand) · 💬 **Chat**
  - 🎥 **Ansicht** (Folgen/Übersicht) · ⚙️ **Qualität** (Sparmodus) · 🚪 **Verlassen**
- In der Teilnehmerliste zeigt eine **Ampel** die Verbindung (grün = verbunden, gelb = verbindet, rot = keine Verbindung)

## Einmalige Einrichtung (durch IT/Owner)

1. **Azure-App-Registrierung** `75e627e8-…`: Redirect-URI (SPA) hinzufügen
   `https://dfedorov12.github.io/3d-space/`
2. **GitHub-Pages** für das Repo `3d-space` aktivieren (Source: GitHub Actions – Workflow liegt bei).
3. **Interne Personen freigeben** – Mail-Adresse in `app.js` → Liste `ALLOWED` eintragen und pushen:

   ```js
   const ALLOWED = ['fedorov@dihag.com', 'administrator@dihag.com', 'neue.person@dihag.com'];
   ```

   (Die optionale SharePoint-Liste `AppPermissions` wird zusätzlich gelesen, falls vorhanden – ist aber keine Voraussetzung.)

## Externe Gäste einladen

- Externe ohne Microsoft-Konto klicken auf der Startseite **„Als Gast beitreten"** (oder öffnen den Link
  `…/3d-space/?guest`), geben ihren Namen + den **Einladungs-Code** ein und sind im Raum.
- Code festlegen/ändern in `app.js`: `const GUEST_PASSCODE = 'dihag-3d';` (auf `''` setzen = ohne Code).
  ⚠️ Der Code steht im öffentlichen JS → leichter Schutz vor Zufallsbesuchern, **kein echtes Geheimnis**.
- Bis zu `MAX_GUESTS` (Standard 10) Gäste gleichzeitig.

## Figuren

- Beim Beitreten wählt jede:r eine Figur: **männlich / weiblich** (stilisierte Low-Poly-Menschen,
  Kleidung in der persönlichen Farbe). Auswahl wird an alle übertragen.

## Hinweise

- Der öffentliche PeerJS-Broker reicht für interne Nutzung; bei Bedarf später eigenen PeerServer
  hinterlegen (`new Peer(id, { host, port, path })` in `app.js`).
- Erstes Sprechen erfordert eine Nutzergeste (Browser-Policy) → „Beitreten & sprechen".
- **TURN-Server (wichtig im Firmennetz):** rein P2P (nur STUN) scheitert oft an Firewall/NAT –
  dann sieht man sich zwar, hört sich aber nicht. Lösung: in `app.js` den `ICE_SERVERS`-TURN-Block
  ausfüllen (Host/User/Passwort, idealerweise `turns:…:5349` über TLS/443).
- **Cache-Busting** ist automatisch: der Deploy-Workflow ersetzt `?v=…` durch den Commit-SHA –
  manuelles Hochzählen entfällt.
- **Skaliert** als Voll-Mesh komfortabel bis ~6–8 Personen; für größere Runden wäre ein SFU nötig.
