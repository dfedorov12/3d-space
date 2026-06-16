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

- **WASD / Pfeiltasten** – im Raum gehen
- **Maus ziehen** – umsehen · **Scrollen** – Zoom
- **Mikro / Melden / Ansicht / Verlassen** – untere Leiste
- Näher an eine Person gehen = sie wird lauter

## Einmalige Einrichtung (durch IT/Owner)

1. **Azure-App-Registrierung** `75e627e8-…`: Redirect-URI (SPA) hinzufügen
   `https://dfedorov12.github.io/3d-space/`
2. **GitHub-Pages** für das Repo `3d-space` aktivieren (Source: GitHub Actions – Workflow liegt bei).
3. **Personen freigeben** – Zeilen in der Liste `AppPermissions` (`/sites/ticket`) anlegen:

   | UserEmail | App | Role |
   |---|---|---|
   | person@dihag.com | `3d-space` | `viewer` |

   `Role = none` sperrt wieder. `App = *` gilt für alle Apps.
   Owner in `SUPER_ADMINS` (siehe `app.js`) kommen immer rein.

## Hinweise

- Der öffentliche PeerJS-Broker reicht für interne Nutzung; bei Bedarf später eigenen PeerServer
  hinterlegen (`new Peer(id, { host, port, path })` in `app.js`).
- Erstes Sprechen erfordert eine Nutzergeste (Browser-Policy) → „Beitreten & sprechen".
