# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the Application

```bash
# Local development
python app.py          # Starts Flask dev server on localhost:5000, auto-creates brewery.db with demo data

# Production
gunicorn app:app --timeout 120
```

The app auto-seeds demo users and activity data on first startup. No build step is required. There are no automated tests or linting tools configured.

## Architecture

**Aktionstracker** is a field sales activity tracker for German-speaking sales teams. The entire backend lives in a single file: `app.py` (~3,500 lines). There is no separate service layer, ORM, or routing module — everything from DB helpers to route handlers to scheduled jobs is in that one file.

```
Browser (PWA, offline-capable)
    │
Flask app.py
    ├── 45+ @app.route handlers (server-side Jinja2 rendering + JSON APIs)
    ├── SQLite helpers: query() / execute() thin wrappers
    ├── APScheduler background jobs (backups, photo cleanup, email reports)
    └── brewery.db (SQLite)
```

Templates live in `templates/` and extend `base.html`. Static assets including `sw.js` (Service Worker) are in `static/`.

### Role-Based Access

Three roles enforced via session checks at every route:

| Role | German | Access |
|---|---|---|
| `admin` | Admin | Full system, user management |
| `verkaufsleiter` | VKL | Team overview, reports, territory assignment |
| `rep` | Außendienst | Own activities only |

The admin account is authenticated against the `ADMIN_PASSWORD` env var (not the DB). Regular users authenticate via email, shorthand (`kuerzel`), or full name.

### Database

SQLite accessed directly with `sqlite3` (no ORM). All schema migrations run automatically at startup via idempotent `ALTER TABLE ... IF NOT EXISTS` checks. Foreign keys are enabled (`PRAGMA foreign_keys = ON`). Soft deletes use an `aktiv` column (0/1) — records are never hard-deleted when they have history.

Key tables:
- `mitarbeiter` — users/employees (name, kuerzel, role, team_id, password hash)
- `aktivitaet` — sales visit logs (date, rep, location, displays, notes, photo_path)
- `verkaufsstelle` — customer locations (address, coordinates, type)
- `bestellposition` — order line items per visit (product, quantity)
- `zielzahlen` — annual sales targets per employee
- `vertretung` — rep coverage/substitution date ranges

### Team Filtering

VKLs only see data for their own team. This is enforced in SQL via two helper functions that inject `WHERE` clauses: `_team_ma_clause()` and `_team_m_clause()`. Any new route that exposes per-rep data must use these helpers.

### Offline / PWA

`static/sw.js` implements a Service Worker with cache-first for CDN assets and network-first for app pages. `static/offline.js` queues activity submissions in IndexedDB when offline and replays them via `POST /api/aktivitaet/offline-sync` (JSON with base64 photo) when reconnected.

## Key Conventions

**Language:** All UI strings, DB column names, route paths, and code comments are in German. New code should follow this convention (e.g. `verkaufsstelle`, not `sales_location`).

**Route naming:** Use German kebab-case paths (`/aktivitaet/neu`, `/einstellungen/wochenbericht`). Admin routes are prefixed `/admin/`, APIs prefixed `/api/`.

**JSON API responses:** Return `{"ok": true, "data": ...}` on success and `{"ok": false, "error": "..."}` on failure.

**Photos:** Uploaded via form POST, compressed to JPEG via Pillow, stored in the filesystem. Auto-deleted after 4 weeks by a scheduled job. The column `photo_path` stores relative paths.

**Excel exports:** Generated in-memory by `_build_excel_bytes()` using `openpyxl`. Multi-sheet workbooks with styled headers — extend this function to add new report sheets rather than creating new export functions.

**Email:** Prefer Resend API (`RESEND_API_KEY`) over SMTP. Fallback to SMTP only when Resend key is absent. Both paths share the same call sites.

## Configuration

All configuration is via environment variables — no config files. Key variables:

| Variable | Purpose |
|---|---|
| `DATABASE_PATH` | SQLite file path (default: `brewery.db`) |
| `ADMIN_PASSWORD` | Admin account password |
| `DEFAULT_PASSWORD` | Initial password for new user accounts |
| `COMPANY_NAME` / `COMPANY_SHORT` | Branding displayed in UI and emails |
| `LOGO_URL` | External logo URL; empty = serve local `static/logo.png` |
| `KARTE_MODUS` | Map feature level: `aus`, `basis`, or `heatmap` |
| `UNIT_LABEL` | Display unit label (e.g. `Kisten`, `Kartons`) |
| `MAX_MITARBEITER` | Employee cap (0 = unlimited) |
| `INIT_DEMO_USERS` | Auto-seed demo accounts on startup |
| `RESEND_API_KEY` | Resend email service key |
| `MAIL_SERVER` / `MAIL_PORT` / `MAIL_USERNAME` / `MAIL_PASSWORD` | SMTP fallback |
| `EXPORT_EMAIL` | Recipient for automatic 4-week Excel exports |
| `APP_BASE_URL` | Full URL used in outbound email links |

## Vor jedem Deploy: Lokaler Pflicht-Test

Vor jedem `git push` (Railway deployt automatisch) den lokalen Server starten und diese Routen prüfen — alle müssen 200 zurückgeben (außer bekannte 404er):

```bash
# Server starten (venv + Demo-Env)
TOUREN_MODUS=an INIT_DEMO_USERS=true venv/Scripts/python app.py &
sleep 6

# Admin einloggen
COOKIE=$(mktemp)
curl -sc "$COOKIE" http://127.0.0.1:5000/ -X POST -d "email=admin&passwort=admin123" > /dev/null

# Pflicht-Routen
curl -so /dev/null -w "Dashboard (admin):       %{http_code}\n" -b "$COOKIE" http://127.0.0.1:5000/dashboard
curl -so /dev/null -w "Tourenplanung (admin):   %{http_code}\n" -b "$COOKIE" http://127.0.0.1:5000/tourenplanung
curl -so /dev/null -w "Wochenbericht-Vorschau:  %{http_code}\n" -b "$COOKIE" http://127.0.0.1:5000/einstellungen/wochenbericht/vorschau
curl -so /dev/null -w "Monatsbericht-Vorschau:  %{http_code}\n" -b "$COOKIE" http://127.0.0.1:5000/einstellungen/monatsbericht/vorschau

# Rep einloggen
COOKIE_REP=$(mktemp)
curl -sc "$COOKIE_REP" http://127.0.0.1:5000/ -X POST -d "email=MM&passwort=start123" > /dev/null
curl -so /dev/null -w "Dashboard (rep):         %{http_code}\n" -b "$COOKIE_REP" http://127.0.0.1:5000/dashboard
```

Schlägt eine Route mit 500 fehl → erst fixen, dann pushen.

**Erweiterte Variante (agent-browser):** `./pflichttest_browser.sh` rendert die Routen echt im Browser (Admin + Rep) statt nur den HTTP-Status zu prüfen — fängt damit auch JS-Fehler und kaputte Interaktionen (z.B. Karte, Formulare) die curl nicht sieht. Screenshots landen in `_pflichttest_shots/` (gitignored). Ergänzt den curl-Test, ersetzt ihn nicht (curl bleibt der schnelle erste Check).

## Sicherheits-Konventionen für neue Features

Diese Regeln bei JEDER neuen Route/JEDEM neuen Feature direkt mitdenken, nicht erst bei einem späteren Audit nachbessern:

- **CSRF:** Jedes neue `<form method="POST">` braucht `<input type="hidden" name="csrf_token" value="{{ csrf_token() }}">`. JS-`fetch()`-POSTs brauchen keinen manuellen Header — `window.fetch` ist in `base.html` global gepatcht und hängt `X-CSRFToken` automatisch an (gleiche Origin, schreibende Methoden). Nur bei echten Sonderfällen (z.B. Offline-Sync-Queue mit potenziell sehr altem Token) `@csrf.exempt` einsetzen und in einem Kommentar begründen, warum `login_required`/die eigentliche Zugriffskontrolle das Risiko trotzdem abdeckt.
- **Autorisierung ist rollenbasiert UND team-/eigentümer-basiert:** Reine Rollenprüfung (`rolle in ('admin','verkaufsleiter')`) reicht bei einer VKL-Route mit einer ID aus URL/Body NICHT — zusätzlich prüfen, ob die betroffene Ressource (Aktivität, Verkaufsstelle, Mitarbeiter, Tagesplan-Eintrag) auch zum eigenen Team/Gebiet gehört (`_team_ma_clause()` / `_team_m_clause()`-Muster). Dieses Muster (Rolle geprüft, Team-Zugehörigkeit vergessen) war die mit Abstand häufigste Schwachstellenklasse in bisherigen Audits — bei Arco explizit gefunden und gefixt, bei Demo von Anfang an mitdenken.
- **Freitext in Excel-Zellen:** immer durch `_excel_formel_sicher()` schicken, bevor er in `ws.cell(...)` landet (verhindert Formel-Injection über z.B. `=HYPERLINK(...)` in Notizen).
- **Freie E-Mail-Empfänger-Felder** (Report-Konfiguration o.ä.): niemals ungeprüft speichern — gegen bekannte Mitarbeiter-E-Mails whitelisten (Muster siehe `einstellungen_wochenbericht()`), sonst kann sich jemand dauerhaft interne Berichte an eine private/externe Adresse schicken lassen.
- **Fehlerantworten an den Client:** nie `str(e)` oder `traceback.format_exc()` in eine JSON-/HTTP-Antwort schreiben — nur `app.logger.error(...)`/`app.logger.warning(...)` mit vollem Trace, dem Client nur eine generische Meldung.
- **Neue Konfigurations-Flags für Sicherheitsverhalten** (z.B. `FORCE_HTTPS`) immer als explizite, dokumentierte Env-Var einführen statt sich nur auf eine Plattform-Heuristik (`RAILWAY_ENVIRONMENT`) zu verlassen — Heuristik als Fallback behalten, nicht ersetzen.
- Die globalen HTTP-Security-Header (`X-Content-Type-Options`, `X-Frame-Options`, `Referrer-Policy`, bedingt `Strict-Transport-Security`) laufen zentral über `@app.after_request` — bei neuen Routen nichts zusätzlich nötig.

## Background Jobs (APScheduler)

All jobs run in Europe/Berlin timezone:

- **Daily:** DB backup (keeps last 7), photo cleanup (>4 weeks old)
- **Every 4 weeks:** Excel + photo ZIP sent to `EXPORT_EMAIL`
- **Mondays 07:00:** Weekly HTML KPI report emailed to VKLs and configured recipients
- **Sundays 23:30:** Demo data refresh (adds 5 activities per rep to keep demos current)
