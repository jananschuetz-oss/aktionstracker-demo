from flask import Flask, render_template, request, redirect, url_for, session, send_file, g, flash, jsonify
import sqlite3
import os
import uuid
import secrets
import shutil
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import zipfile
from datetime import datetime, date, timedelta
import io
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from werkzeug.utils import secure_filename
import urllib.request
import urllib.parse
import time as _time
import threading

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'aktionstracker_geheim_xK9m')

# ── Session-Sicherheit ────────────────────────────────────────────────────────
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=8)
app.config['SESSION_COOKIE_HTTPONLY']    = True
app.config['SESSION_COOKIE_SAMESITE']   = 'Lax'
app.config['SESSION_COOKIE_SECURE']     = os.environ.get('RAILWAY_ENVIRONMENT') is not None

DATABASE = os.environ.get('DATABASE_PATH', 'brewery.db')
LOGO_VERSION = '3'  # cache-bust

# ── Branding (pro-Kunde via Railway ENV Variables anpassbar) ─────────────────
# Railway: Settings → Variables → diese Variablen setzen
COMPANY_NAME   = os.environ.get('COMPANY_NAME',   'Ihre Firma GmbH')
COMPANY_SHORT  = os.environ.get('COMPANY_SHORT',  'Demo')
LOGO_URL       = os.environ.get('LOGO_URL',       '')    # externe Bild-URL oder leer → lokale Datei
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'admin123')
EXPORT_EMAIL   = os.environ.get('EXPORT_EMAIL',   '')        # E-Mail für automatischen 4-Wochen-Export
KARTE_MODUS    = os.environ.get('KARTE_MODUS',   'basis')   # 'aus' | 'basis' | 'heatmap'

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), 'static', 'uploads')
ALLOWED_EXTENSIONS = {'jpg', 'jpeg', 'png', 'gif', 'webp', 'heic', 'heif'}
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB max upload

BACKUP_FOLDER = os.path.join(os.path.dirname(__file__), 'backups')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(BACKUP_FOLDER, exist_ok=True)

# ── E-Mail-Konfiguration (via Umgebungsvariablen setzen) ──────────────────────
# Railway: Settings → Variables → diese Variablen eintragen
MAIL_SERVER   = os.environ.get('MAIL_SERVER',   '')          # z.B. smtp.gmail.com
MAIL_PORT     = int(os.environ.get('MAIL_PORT',  587))
MAIL_USE_TLS  = os.environ.get('MAIL_USE_TLS',  'true').lower() == 'true'
MAIL_USERNAME = os.environ.get('MAIL_USERNAME', '')
MAIL_PASSWORD = os.environ.get('MAIL_PASSWORD', '')          # App-Passwort bei Gmail
MAIL_FROM     = os.environ.get('MAIL_FROM',     MAIL_USERNAME)
APP_BASE_URL  = os.environ.get('APP_BASE_URL',  '')          # z.B. https://mein-tool.up.railway.app

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Saisonaler Verteilungsschlüssel für Produktverkäufe (Jan–Dez, Summe = 1.0)
# Sommer-Peak Juli/August, Herbst-Spike September/Oktober
SONNENSCHLUESSEL = [0.05, 0.05, 0.07, 0.08, 0.10, 0.12, 0.13, 0.12, 0.09, 0.08, 0.06, 0.05]
M_NAMEN = ['Jan', 'Feb', 'Mär', 'Apr', 'Mai', 'Jun', 'Jul', 'Aug', 'Sep', 'Okt', 'Nov', 'Dez']

@app.template_filter('from_json')
def from_json_filter(s):
    import json
    return json.loads(s)

@app.template_filter('todatetime')
def todatetime_filter(s):
    """Konvertiert 'YYYY-MM-DD' String zu date-Objekt für Datumsvergleiche im Template."""
    try:
        return date.fromisoformat(str(s))
    except Exception:
        return date.today()

@app.before_request
def check_session_lifetime():
    """Leert die Session wenn sie abgelaufen ist (PERMANENT_SESSION_LIFETIME)."""
    session.modified = False   # kein unnötiges Re-Schreiben

@app.context_processor
def inject_now():
    ctx = {
        'now':           datetime.now(),
        'company_name':  COMPANY_NAME,
        'company_short': COMPANY_SHORT,
        'logo_url':      LOGO_URL or '/static/logo.svg',
        'karte_modus':   KARTE_MODUS,
        'meine_vertretungen': [],
        'alle_kollegen':      [],
    }
    if session.get('user_id'):
        try:
            ctx['meine_vertretungen'] = query(
                '''SELECT v.id, v.von, v.bis, m.name AS vertreter_name
                   FROM vertretung v
                   JOIN mitarbeiter m ON m.id = v.vertreter_id
                   WHERE v.abwesender_id = ?
                   ORDER BY v.von DESC''',
                (session['user_id'],)
            )
            ctx['alle_kollegen'] = query(
                "SELECT id, name FROM mitarbeiter WHERE rolle IN ('rep','verkaufsleiter') AND id != ? ORDER BY name",
                (session['user_id'],)
            )
        except Exception:
            pass
    return ctx


# ─── DB Helpers ───────────────────────────────────────────────────────────────

def get_db():
    db = getattr(g, '_database', None)
    if db is None:
        db = g._database = sqlite3.connect(DATABASE)
        db.row_factory = sqlite3.Row
        db.execute("PRAGMA foreign_keys = ON")
    return db

@app.teardown_appcontext
def close_connection(exception):
    db = getattr(g, '_database', None)
    if db is not None:
        db.close()

def query(sql, args=(), one=False):
    cur = get_db().execute(sql, args)
    rv = cur.fetchall()
    return (rv[0] if rv else None) if one else rv

def execute(sql, args=()):
    db = get_db()
    cur = db.execute(sql, args)
    db.commit()
    return cur.lastrowid


FOTO_AUFBEWAHRUNG_WOCHEN = 4   # Fotos werden nach 4 Wochen gelöscht (werden vorher per Auto-Export an Kunden geschickt)

# ── E-Mail versenden ──────────────────────────────────────────────────────────

def send_email(to: str, subject: str, body_html: str) -> bool:
    """Sendet eine HTML-E-Mail. Gibt True bei Erfolg zurück."""
    if not MAIL_SERVER or not MAIL_USERNAME:
        app.logger.warning("E-Mail nicht konfiguriert (MAIL_SERVER / MAIL_USERNAME fehlen).")
        return False
    try:
        msg = MIMEMultipart('alternative')
        msg['Subject'] = subject
        msg['From']    = MAIL_FROM
        msg['To']      = to
        msg.attach(MIMEText(body_html, 'html', 'utf-8'))
        with smtplib.SMTP(MAIL_SERVER, MAIL_PORT, timeout=15) as smtp:
            if MAIL_USE_TLS:
                smtp.starttls()
            smtp.login(MAIL_USERNAME, MAIL_PASSWORD)
            smtp.send_message(msg)
        return True
    except Exception as e:
        app.logger.error(f"E-Mail-Fehler: {e}")
        return False


def send_email_with_attachments(to: str, subject: str, body_html: str,
                                attachments: list) -> bool:
    """Sendet eine HTML-E-Mail mit Dateianhängen.
    attachments: Liste von (dateiname, bytes_daten, content_type) Tupeln."""
    if not MAIL_SERVER or not MAIL_USERNAME:
        app.logger.warning("E-Mail nicht konfiguriert – Auto-Export nicht möglich.")
        return False
    try:
        msg = MIMEMultipart('mixed')
        msg['Subject'] = subject
        msg['From']    = MAIL_FROM
        msg['To']      = to
        msg.attach(MIMEText(body_html, 'html', 'utf-8'))
        for dateiname, daten, content_type in attachments:
            haupttyp, untertyp = content_type.split('/', 1)
            part = MIMEBase(haupttyp, untertyp)
            part.set_payload(daten)
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment', filename=dateiname)
            msg.attach(part)
        with smtplib.SMTP(MAIL_SERVER, MAIL_PORT, timeout=60) as smtp:
            if MAIL_USE_TLS:
                smtp.starttls()
            smtp.login(MAIL_USERNAME, MAIL_PASSWORD)
            smtp.send_message(msg)
        return True
    except Exception as e:
        app.logger.error(f"E-Mail-Fehler (Anhang): {e}")
        return False


# ── Datenbank-Backup ──────────────────────────────────────────────────────────

def backup_db():
    """Erstellt ein tägliches Backup der DB. Behält die letzten 7 Tage."""
    if not os.path.exists(DATABASE):
        return
    heute       = date.today().isoformat()
    backup_pfad = os.path.join(BACKUP_FOLDER, f'brewery_{heute}.db')
    if not os.path.exists(backup_pfad):
        shutil.copy2(DATABASE, backup_pfad)
        # Nur die letzten 7 Backups behalten
        alle = sorted([
            f for f in os.listdir(BACKUP_FOLDER)
            if f.startswith('brewery_') and f.endswith('.db')
        ])
        for alt in alle[:-7]:
            try:
                os.remove(os.path.join(BACKUP_FOLDER, alt))
            except Exception:
                pass


def cleanup_alte_fotos():
    """Löscht Foto-Dateien die älter als FOTO_AUFBEWAHRUNG_WOCHEN Wochen sind
    und setzt foto_pfad in der DB auf NULL. Gibt Anzahl gelöschter Fotos zurück."""
    from datetime import timedelta
    grenzwert = (date.today() - timedelta(weeks=FOTO_AUFBEWAHRUNG_WOCHEN)).isoformat()
    db = get_db()
    alte_akte = db.execute(
        "SELECT id, foto_pfad FROM aktivitaet WHERE foto_pfad IS NOT NULL AND foto_pfad != '' AND datum < ?",
        (grenzwert,)
    ).fetchall()
    count = 0
    for akt in alte_akte:
        pfad = os.path.join(UPLOAD_FOLDER, akt['foto_pfad'])
        if os.path.exists(pfad):
            os.remove(pfad)
            count += 1
        db.execute("UPDATE aktivitaet SET foto_pfad = NULL WHERE id = ?", (akt['id'],))
    if alte_akte:
        db.commit()
    return count


def init_db():
    with app.app_context():
        db = get_db()
        db.executescript('''
            CREATE TABLE IF NOT EXISTS mitarbeiter (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                kuerzel TEXT NOT NULL UNIQUE,
                rolle TEXT DEFAULT 'rep',
                passwort TEXT DEFAULT 'brauerei'
            );

            CREATE TABLE IF NOT EXISTS verkaufsstelle (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                ort TEXT,
                typ TEXT,
                aktiv INTEGER DEFAULT 1
            );

            CREATE TABLE IF NOT EXISTS biersorte (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                einheit TEXT DEFAULT 'Kiste (20x0.5L)',
                aktiv INTEGER DEFAULT 1
            );

            CREATE TABLE IF NOT EXISTS aktivitaet (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                datum DATE NOT NULL,
                mitarbeiter_id INTEGER NOT NULL,
                verkaufsstelle_id INTEGER NOT NULL,
                anzahl_displays INTEGER DEFAULT 0,
                notizen TEXT,
                foto_pfad TEXT,
                erstellt_am TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (mitarbeiter_id) REFERENCES mitarbeiter(id),
                FOREIGN KEY (verkaufsstelle_id) REFERENCES verkaufsstelle(id)
            );

            CREATE TABLE IF NOT EXISTS bestellposition (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                aktivitaet_id INTEGER NOT NULL,
                biersorte_id INTEGER NOT NULL,
                kisten_anzahl INTEGER NOT NULL,
                FOREIGN KEY (aktivitaet_id) REFERENCES aktivitaet(id) ON DELETE CASCADE,
                FOREIGN KEY (biersorte_id) REFERENCES biersorte(id)
            );

            CREATE TABLE IF NOT EXISTS zielzahlen (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                mitarbeiter_id INTEGER,
                jahr INTEGER NOT NULL,
                displays_ziel INTEGER DEFAULT 0,
                kisten_ziel INTEGER DEFAULT 0,
                UNIQUE(mitarbeiter_id, jahr),
                FOREIGN KEY (mitarbeiter_id) REFERENCES mitarbeiter(id)
            );

            CREATE TABLE IF NOT EXISTS displaysorte (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                aktiv INTEGER DEFAULT 1
            );

            CREATE TABLE IF NOT EXISTS displayposition (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                aktivitaet_id INTEGER NOT NULL,
                displaysorte_id INTEGER NOT NULL,
                anzahl INTEGER NOT NULL DEFAULT 0,
                FOREIGN KEY (aktivitaet_id) REFERENCES aktivitaet(id) ON DELETE CASCADE,
                FOREIGN KEY (displaysorte_id) REFERENCES displaysorte(id)
            );

            CREATE TABLE IF NOT EXISTS mitarbeiter_verkaufsstelle (
                mitarbeiter_id    INTEGER NOT NULL,
                verkaufsstelle_id INTEGER NOT NULL,
                PRIMARY KEY (mitarbeiter_id, verkaufsstelle_id),
                FOREIGN KEY (mitarbeiter_id)    REFERENCES mitarbeiter(id)    ON DELETE CASCADE,
                FOREIGN KEY (verkaufsstelle_id) REFERENCES verkaufsstelle(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS vertretung (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                abwesender_id INTEGER NOT NULL,
                vertreter_id  INTEGER NOT NULL,
                von           DATE NOT NULL,
                bis           DATE NOT NULL,
                FOREIGN KEY (abwesender_id) REFERENCES mitarbeiter(id) ON DELETE CASCADE,
                FOREIGN KEY (vertreter_id)  REFERENCES mitarbeiter(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS wochenbericht_config (
                id             INTEGER PRIMARY KEY CHECK (id = 1),
                aktiv          INTEGER DEFAULT 0,
                empfaenger_2   TEXT    DEFAULT '',
                empfaenger_3   TEXT    DEFAULT '',
                zuletzt_gesendet TEXT  DEFAULT ''
            );
            INSERT OR IGNORE INTO wochenbericht_config (id) VALUES (1);
        ''')

        # Migrationen für bestehende DBs
        for migration in [
            "ALTER TABLE aktivitaet    ADD COLUMN foto_pfad          TEXT",
            "ALTER TABLE mitarbeiter   ADD COLUMN email               TEXT",
            "ALTER TABLE mitarbeiter   ADD COLUMN reset_token         TEXT",
            "ALTER TABLE mitarbeiter   ADD COLUMN reset_token_ablauf  DATETIME",
            "ALTER TABLE verkaufsstelle ADD COLUMN strasse             TEXT",
            "ALTER TABLE verkaufsstelle ADD COLUMN ansprechpartner    TEXT",
            "ALTER TABLE verkaufsstelle ADD COLUMN lat                REAL",
            "ALTER TABLE verkaufsstelle ADD COLUMN lng                REAL",
            "ALTER TABLE mitarbeiter   ADD COLUMN karte_benachrichtigung TEXT",
        ]:
            try:
                db.execute(migration)
                db.commit()
            except Exception:
                pass  # Spalte existiert bereits

        # Admin + Verkaufsleiter (Passwort via ENV ADMIN_PASSWORD konfigurierbar)
        db.execute("INSERT OR IGNORE INTO mitarbeiter (name, kuerzel, rolle, passwort) VALUES ('Administrator', 'ADMIN', 'admin', ?)", (ADMIN_PASSWORD,))
        db.execute("UPDATE mitarbeiter SET passwort=? WHERE kuerzel='ADMIN'", (ADMIN_PASSWORD,))
        db.execute("INSERT OR IGNORE INTO mitarbeiter (name, kuerzel, rolle, passwort) VALUES ('Verkaufsleiter', 'VKL', 'verkaufsleiter', 'demo123')")

        # Beispiel-Mitarbeiter (nur bei INIT_DEMO_USERS=true)
        if os.environ.get('INIT_DEMO_USERS', 'true').lower() == 'true':
            reps = [
                ('Max Müller',     'MM', 'demo123'),
                ('Anna Schmidt',   'AS', 'demo123'),
                ('Thomas Weber',   'TW', 'demo123'),
                ('Lisa Fischer',   'LF', 'demo123'),
                ('Klaus Hoffmann', 'KH', 'demo123'),
            ]
            for name, kuerzel, pw in reps:
                db.execute("INSERT OR IGNORE INTO mitarbeiter (name, kuerzel, passwort) VALUES (?, ?, ?)", (name, kuerzel, pw))

        # Displaysorten – nur einfügen wenn Tabelle leer
        if not db.execute("SELECT 1 FROM displaysorte LIMIT 1").fetchone():
            for ds_name in ['Regal-Display', 'Eingangs-Display',
                            'Counter-Display', 'Schaufenster', 'Außenwerbung']:
                db.execute("INSERT OR IGNORE INTO displaysorte (name) VALUES (?)", (ds_name,))

        # Produkte – nur einfügen wenn Tabelle leer
        if not db.execute("SELECT 1 FROM biersorte LIMIT 1").fetchone():
            produkte = [
                ('Produkt A',  'Karton (12 Stück)'),
                ('Produkt B',  'Karton (12 Stück)'),
                ('Produkt C',  'Karton (24 Stück)'),
                ('Produkt D',  'Karton (24 Stück)'),
                ('Produkt E',  'Palette (100 Stück)'),
                ('Produkt F',  'Palette (100 Stück)'),
            ]
            for name, einheit in produkte:
                db.execute("INSERT INTO biersorte (name, einheit) VALUES (?, ?)", (name, einheit))

        # Beispiel-Kunden – nur einfügen wenn Tabelle leer
        if not db.execute("SELECT 1 FROM verkaufsstelle LIMIT 1").fetchone():
            stellen = [
                ('Supermarkt Mitte',       'Berlin',   'Einzelhandel'),
                ('Fachmarkt Nord',         'Hamburg',  'Einzelhandel'),
                ('Restaurant Zur Post',    'München',  'Gastronomie'),
                ('Hotel Stadtblick',       'Frankfurt','Hotel'),
                ('Großhandel Meyer',       'Köln',     'Getränkehandel'),
                ('Kiosk am Bahnhof',       'Düsseldorf','Kiosk'),
                ('Sportverein 1902',       'Stuttgart','Verein'),
                ('Café Central',           'Leipzig',  'Gastronomie'),
            ]
            for name, ort, typ in stellen:
                db.execute("INSERT INTO verkaufsstelle (name, ort, typ) VALUES (?, ?, ?)", (name, ort, typ))

        db.commit()

        # Beispieldaten einfügen wenn DB noch leer
        if not db.execute("SELECT 1 FROM aktivitaet LIMIT 1").fetchone():
            seed_demo_data(db)

        # Beispielfotos einmalig zuweisen – NACH seed_demo_data (Aktivitäten müssen existieren)
        fotos_in_db = db.execute(
            "SELECT COUNT(*) FROM aktivitaet WHERE foto_pfad IS NOT NULL AND foto_pfad != ''"
        ).fetchone()[0]
        if fotos_in_db == 0 and os.path.isdir(UPLOAD_FOLDER):
            upload_files = sorted([
                f for f in os.listdir(UPLOAD_FOLDER)
                if f.startswith('akt_') and f.endswith('.jpg')
            ])
            if upload_files:
                mai_akte = db.execute("""
                    SELECT id FROM aktivitaet
                    WHERE strftime('%Y-%m', datum) = '2026-05'
                      AND (foto_pfad IS NULL OR foto_pfad = '')
                    ORDER BY datum, id
                    LIMIT ?
                """, (len(upload_files),)).fetchall()
                for akt_row, dateiname in zip(mai_akte, upload_files):
                    db.execute("UPDATE aktivitaet SET foto_pfad = ? WHERE id = ?",
                               (dateiname, akt_row['id']))
                db.commit()

        # Stationszuordnung: alle noch nicht zugeordneten aktiven Stationen gleichmäßig verteilen
        unzugeordnet = db.execute("""
            SELECT v.id FROM verkaufsstelle v
            LEFT JOIN mitarbeiter_verkaufsstelle mv ON mv.verkaufsstelle_id = v.id
            WHERE v.aktiv = 1 AND mv.mitarbeiter_id IS NULL
        """).fetchall()
        if unzugeordnet:
            import random as _rnd_assign
            _rnd_assign.seed(99)
            pool    = [s['id'] for s in unzugeordnet]
            _rnd_assign.shuffle(pool)
            reps_ma = db.execute("SELECT id FROM mitarbeiter WHERE rolle='rep'").fetchall()
            vkls_ma = db.execute("SELECT id FROM mitarbeiter WHERE rolle='verkaufsleiter'").fetchall()
            # VKLs: je bis zu 3 Stationen (nur wenn sie noch keine haben)
            for vkl in vkls_ma:
                hat_bereits = db.execute(
                    "SELECT COUNT(*) FROM mitarbeiter_verkaufsstelle WHERE mitarbeiter_id=?",
                    (vkl['id'],)
                ).fetchone()[0]
                fehlend = max(0, 3 - hat_bereits)
                for _ in range(min(fehlend, len(pool))):
                    db.execute(
                        "INSERT OR IGNORE INTO mitarbeiter_verkaufsstelle (mitarbeiter_id, verkaufsstelle_id) VALUES (?,?)",
                        (vkl['id'], pool.pop(0))
                    )
            # Reps: Rest gleichmäßig
            if reps_ma:
                for idx, st_id in enumerate(pool):
                    rep = reps_ma[idx % len(reps_ma)]
                    db.execute(
                        "INSERT OR IGNORE INTO mitarbeiter_verkaufsstelle (mitarbeiter_id, verkaufsstelle_id) VALUES (?,?)",
                        (rep['id'], st_id)
                    )
            db.commit()
            app.logger.info(f"Stationszuordnung: {len(unzugeordnet)} Stationen automatisch verteilt.")

        # Demo-Koordinaten: Stationen ohne lat/lng anhand der Stadt direkt setzen
        ohne_coords = db.execute(
            "SELECT id, ort FROM verkaufsstelle WHERE aktiv=1 AND (lat IS NULL OR lng IS NULL)"
        ).fetchall()
        if ohne_coords:
            import random as _rnd_geo
            _rnd_geo.seed(77)
            STADT_COORDS = {
                'Berlin':     (52.5200, 13.4050),
                'Hamburg':    (53.5753, 10.0153),
                'München':    (48.1374, 11.5755),
                'Frankfurt':  (50.1109,  8.6821),
                'Köln':       (50.9333,  6.9500),
                'Düsseldorf': (51.2217,  6.7762),
                'Stuttgart':  (48.7758,  9.1829),
                'Leipzig':    (51.3397, 12.3731),
                'Nürnberg':   (49.4521, 11.0767),
                'Hannover':   (52.3759,  9.7320),
                'Mannheim':   (49.4875,  8.4660),
                'Dortmund':   (51.5135,  7.4653),
                'Bremen':     (53.0793,  8.8017),
                'Wiesbaden':  (50.0800,  8.2400),
                'Bonn':       (50.7374,  7.0982),
                'Freiburg':   (47.9990,  7.8421),
                'Essen':      (51.4508,  7.0131),
                'Augsburg':   (48.3717, 10.8983),
                'Starnberg':  (47.9986, 11.3381),
                'Dachau':     (48.2604, 11.4335),
            }
            gesetzt = 0
            for vs in ohne_coords:
                base = STADT_COORDS.get(vs['ort'])
                if base:
                    lat = base[0] + _rnd_geo.uniform(-0.05, 0.05)
                    lng = base[1] + _rnd_geo.uniform(-0.07, 0.07)
                    db.execute("UPDATE verkaufsstelle SET lat=?, lng=? WHERE id=?", (lat, lng, vs['id']))
                    gesetzt += 1
            db.commit()
            app.logger.info(f"Demo-Koordinaten: {gesetzt}/{len(ohne_coords)} Stationen mit Stadtkoordinaten gesetzt.")

        # Alte Fotos beim Start bereinigen
        cleanup_alte_fotos()

        # Tägliches DB-Backup
        backup_db()


def seed_demo_data(db):
    """Füllt die DB mit realistischen Beispieldaten für KW 01–22, 2026."""
    import random as rnd
    from datetime import date, timedelta
    rnd.seed(42)

    # Zusätzliche Verkaufsstellen
    extra = [
        ('Restaurant Zum Marktplatz',    'Nürnberg',          'Gastronomie'),
        ('Bistro Central',               'Hannover',          'Gastronomie'),
        ('Ristorante Bella Vista',       'Mannheim',          'Gastronomie'),
        ('Steakhouse Westend',           'Frankfurt',         'Gastronomie'),
        ('Café Metropol',                'Dortmund',          'Gastronomie'),
        ('Pizzeria Napoli',              'Bremen',            'Gastronomie'),
        ('Imbiss Am Stadtpark',          'Essen',             'Gastronomie'),
        ('Gasthaus Lindenhof',           'Wiesbaden',         'Gastronomie'),
        ('Stadthotel am Ring',           'Bonn',              'Hotel'),
        ('Pension Garni Sonnenhof',      'Freiburg',          'Hotel'),
        ('Supermarkt Stadtmitte',        'Nürnberg',          'Einzelhandel'),
        ('Verbrauchermarkt Nord',        'Hannover',          'Einzelhandel'),
        ('Discountmarkt Westend',        'Mannheim',          'Einzelhandel'),
        ('Großhandel Fischer',           'Frankfurt',         'Getränkehandel'),
        ('Cash & Carry Zentrum',         'Dortmund',          'Getränkehandel'),
        ('Handelskontor Weber',          'Bremen',            'Getränkehandel'),
        ('Sportverein Blau-Weiß',        'Essen',             'Verein'),
        ('Schützengesellschaft 1888',    'Wiesbaden',         'Verein'),
        ('TSG Vereinsheim',              'Bonn',              'Verein'),
        ('Stadionkiosk SV Mitte',        'Freiburg',          'Verein'),
    ]
    for name, ort, typ in extra:
        if not db.execute("SELECT 1 FROM verkaufsstelle WHERE name=?", (name,)).fetchone():
            db.execute("INSERT INTO verkaufsstelle (name,ort,typ) VALUES (?,?,?)", (name, ort, typ))

    reps    = db.execute("SELECT id, kuerzel FROM mitarbeiter WHERE rolle='rep'").fetchall()
    stellen = db.execute("SELECT id, typ FROM verkaufsstelle WHERE aktiv=1").fetchall()
    biere   = db.execute("SELECT id, name FROM biersorte WHERE aktiv=1").fetchall()
    bier_by = {b['name']: b['id'] for b in biere}

    PREF = {
        'Gastronomie':    ['Produkt A', 'Produkt B', 'Produkt C'],
        'Einzelhandel':   ['Produkt A', 'Produkt B', 'Produkt C', 'Produkt D'],
        'Getränkehandel': ['Produkt A', 'Produkt B', 'Produkt C', 'Produkt D', 'Produkt E', 'Produkt F'],
        'Hotel':          ['Produkt A', 'Produkt B', 'Produkt C'],
        'Verein':         ['Produkt A', 'Produkt C', 'Produkt D'],
        'Kiosk':          ['Produkt A', 'Produkt B'],
    }
    NOTIZEN = [
        '', '', '', '',
        'Sonderaktion vereinbart', 'Kunde sehr zufrieden',
        'Neues Kühlregal besprochen', 'Probierpaket mitgenommen',
        'Konkurrenzprodukte gesichtet', 'Rückgabe 3 leere Displays',
        'Termin für Herbstaktion vereinbart', 'Preiserhöhung kommuniziert',
        'Stammkunde, läuft sehr gut', 'Beschwerden über Lieferzeit',
    ]

    def positionen(typ):
        namen = PREF.get(typ, [b['name'] for b in biere])
        auswahl = rnd.sample(namen, k=rnd.randint(2, min(5, len(namen))))
        return [(bier_by[n], rnd.randint(3, 50)) for n in auswahl if n in bier_by]

    for kw in range(1, 23):
        monday = date.fromisocalendar(2026, kw, 1)
        for rep in reps:
            for tag in sorted(rnd.sample(range(5), k=rnd.randint(2, 4))):
                vs = rnd.choice(stellen)
                displays = rnd.choices([0,1,2,3,4,5], weights=[30,25,20,12,8,5])[0]
                cur = db.execute(
                    "INSERT INTO aktivitaet (datum,mitarbeiter_id,verkaufsstelle_id,anzahl_displays,notizen) "
                    "VALUES (?,?,?,?,?)",
                    ((monday + timedelta(days=tag)).isoformat(),
                     rep['id'], vs['id'], displays, rnd.choice(NOTIZEN))
                )
                for bier_id, menge in positionen(vs['typ']):
                    db.execute(
                        "INSERT INTO bestellposition (aktivitaet_id,biersorte_id,kisten_anzahl) VALUES (?,?,?)",
                        (cur.lastrowid, bier_id, menge)
                    )

    # Zielzahlen 2026 – ambitionierte Jahresziele (Reps bei ~55 % zur Jahresmitte)
    ZIELE = {'MM':(200,12000),'AS':(185,11500),'TW':(175,11000),'LF':(175,11000),'KH':(165,10500)}
    for rep in reps:
        if rep['kuerzel'] in ZIELE:
            d, k = ZIELE[rep['kuerzel']]
            db.execute('''INSERT INTO zielzahlen (mitarbeiter_id,jahr,displays_ziel,kisten_ziel)
                VALUES (?,2026,?,?) ON CONFLICT(mitarbeiter_id,jahr) DO UPDATE SET
                displays_ziel=excluded.displays_ziel, kisten_ziel=excluded.kisten_ziel''',
                (rep['id'], d, k))
    db.execute('''INSERT INTO zielzahlen (mitarbeiter_id,jahr,displays_ziel,kisten_ziel)
        VALUES (NULL,2026,900,56000) ON CONFLICT(mitarbeiter_id,jahr) DO UPDATE SET
        displays_ziel=excluded.displays_ziel, kisten_ziel=excluded.kisten_ziel''')

    # Verkaufsstellen gleichmäßig auf Reps verteilen (für Karten-Demo)
    alle_stellen = db.execute("SELECT id FROM verkaufsstelle WHERE aktiv=1 ORDER BY id").fetchall()
    stellen_ids  = [s['id'] for s in alle_stellen]
    rnd.shuffle(stellen_ids)
    for idx, stelle_id in enumerate(stellen_ids):
        rep = reps[idx % len(reps)]
        db.execute(
            "INSERT OR IGNORE INTO mitarbeiter_verkaufsstelle (mitarbeiter_id, verkaufsstelle_id) VALUES (?,?)",
            (rep['id'], stelle_id)
        )
    db.commit()


# ─── Auth ─────────────────────────────────────────────────────────────────────

def login_required(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

def admin_required(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        if session.get('rolle') != 'admin':
            flash('Zugriff verweigert – nur für Administratoren.', 'danger')
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated

def manager_required(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        if session.get('rolle') not in ('admin', 'verkaufsleiter'):
            flash('Zugriff verweigert – nur für Verkaufsleiter und Administratoren.', 'danger')
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated


# ─── Routes: Auth ─────────────────────────────────────────────────────────────

@app.route('/', methods=['GET', 'POST'])
def login():
    if 'user_id' in session:
        return redirect(url_for('dashboard'))

    if request.method == 'POST':
        email_input = request.form.get('email', '').strip()
        passwort    = request.form.get('passwort', '')

        # ADMIN-Direktlogin: Passwort aus ENV, DB-unabhängig
        _admin_pw = ADMIN_PASSWORD.strip()
        if email_input.upper() == 'ADMIN' and passwort == _admin_pw:
            db = get_db()
            db.execute("INSERT OR IGNORE INTO mitarbeiter (name,kuerzel,rolle,passwort) VALUES ('Administrator','ADMIN','admin',?)", (_admin_pw,))
            db.execute("UPDATE mitarbeiter SET passwort=? WHERE kuerzel='ADMIN'", (_admin_pw,))
            db.commit()
            admin = db.execute("SELECT * FROM mitarbeiter WHERE kuerzel='ADMIN'").fetchone()
            if admin:
                session.permanent = True
                session['user_id'] = admin['id']
                session['name']    = admin['name']
                session['kuerzel'] = admin['kuerzel']
                session['rolle']   = admin['rolle']
                return redirect(url_for('dashboard'))

        # Normale Login-Logik für alle anderen
        user = query("SELECT * FROM mitarbeiter WHERE LOWER(email) = LOWER(?)", (email_input,), one=True)
        if not user:
            user = query("SELECT * FROM mitarbeiter WHERE UPPER(kuerzel) = UPPER(?)", (email_input,), one=True)
        if user and user['passwort'] == passwort:
            session.permanent  = True          # läuft nach PERMANENT_SESSION_LIFETIME ab
            session['user_id'] = user['id']
            session['name']    = user['name']
            session['kuerzel'] = user['kuerzel']
            session['rolle']   = user['rolle']
            # Karte-Benachrichtigungen in Session laden (für Login-Notification)
            benachrichtigung = user['karte_benachrichtigung'] if 'karte_benachrichtigung' in user.keys() else None
            if benachrichtigung:
                session['karte_benachrichtigung'] = benachrichtigung
            return redirect(url_for('dashboard'))
        flash('Ungültige E-Mail-Adresse oder falsches Passwort.', 'danger')

    return render_template('login.html')


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


# ─── Passwort-Reset per E-Mail ────────────────────────────────────────────────

@app.route('/passwort-vergessen', methods=['GET', 'POST'])
def passwort_vergessen():
    mail_konfiguriert = bool(MAIL_SERVER and MAIL_USERNAME)
    if request.method == 'POST':
        eingabe = request.form.get('eingabe', '').strip()
        ma = query(
            "SELECT * FROM mitarbeiter WHERE (UPPER(kuerzel)=UPPER(?) OR email=?) AND rolle!='admin'",
            (eingabe, eingabe), one=True
        )
        if ma and ma['email']:
            token  = secrets.token_urlsafe(32)
            ablauf = (datetime.now() + timedelta(hours=1)).isoformat()
            execute("UPDATE mitarbeiter SET reset_token=?, reset_token_ablauf=? WHERE id=?",
                    (token, ablauf, ma['id']))
            base = APP_BASE_URL or request.host_url.rstrip('/')
            reset_url = f"{base}{url_for('passwort_reset', token=token)}"
            html = f"""
            <div style="font-family:Arial,sans-serif;max-width:500px;margin:0 auto">
              <div style="background:#1a3a5c;padding:20px 30px;border-radius:8px 8px 0 0">
                <h2 style="color:#fff;margin:0">📊 Aktions Tracker</h2>
              </div>
              <div style="background:#f9f9f9;padding:30px;border:1px solid #ddd;border-radius:0 0 8px 8px">
                <p>Hallo <strong>{ma['name']}</strong>,</p>
                <p>Sie haben eine Passwort-Zurücksetzung angefordert. Klicken Sie auf den Button:</p>
                <p style="text-align:center;margin:30px 0">
                  <a href="{reset_url}"
                     style="background:#1a3a5c;color:#fff;padding:14px 32px;border-radius:6px;
                            text-decoration:none;font-weight:bold;font-size:16px">
                    Passwort zurücksetzen
                  </a>
                </p>
                <p style="color:#888;font-size:13px">
                  Der Link ist <strong>1 Stunde</strong> gültig.<br>
                  Falls Sie diese Anfrage nicht gestellt haben, ignorieren Sie diese E-Mail.
                </p>
                <hr style="border:none;border-top:1px solid #ddd;margin:20px 0">
                <p style="color:#aaa;font-size:11px">
                  Aktions Tracker &mdash; automatisch generiert
                </p>
              </div>
            </div>"""
            send_email(ma['email'], 'Passwort zurücksetzen – Aktions Tracker', html)
        # Immer dieselbe Meldung (Sicherheit: kein Hinweis ob Konto existiert)
        flash('Falls ein Konto mit diesen Daten existiert, wurde eine E-Mail gesendet.', 'info')
        return redirect(url_for('login'))
    return render_template('passwort_vergessen.html', mail_konfiguriert=mail_konfiguriert)


@app.route('/passwort-reset/<token>', methods=['GET', 'POST'])
def passwort_reset(token):
    ma = query(
        "SELECT * FROM mitarbeiter WHERE reset_token=? AND reset_token_ablauf > ?",
        (token, datetime.now().isoformat()), one=True
    )
    if not ma:
        flash('Der Reset-Link ist ungültig oder abgelaufen. Bitte neu anfordern.', 'danger')
        return redirect(url_for('login'))
    if request.method == 'POST':
        neues_pw = request.form.get('passwort', '').strip()
        bestaet  = request.form.get('passwort2', '').strip()
        if len(neues_pw) < 6:
            flash('Passwort muss mindestens 6 Zeichen haben.', 'danger')
            return render_template('passwort_reset.html', token=token, name=ma['name'])
        if neues_pw != bestaet:
            flash('Passwörter stimmen nicht überein.', 'danger')
            return render_template('passwort_reset.html', token=token, name=ma['name'])
        execute("UPDATE mitarbeiter SET passwort=?, reset_token=NULL, reset_token_ablauf=NULL WHERE id=?",
                (neues_pw, ma['id']))
        flash('Passwort erfolgreich geändert! Bitte jetzt anmelden.', 'success')
        return redirect(url_for('login'))
    return render_template('passwort_reset.html', token=token, name=ma['name'])


# ─── Rechtliche Seiten ────────────────────────────────────────────────────────

@app.route('/impressum')
def impressum():
    return render_template('impressum.html')

@app.route('/datenschutz')
def datenschutz():
    return render_template('datenschutz.html')

@app.route('/agb')
def agb():
    return render_template('agb.html')

@app.route('/avv')
def avv():
    return render_template('avv.html')


# ─── Logo direkt ausliefern (umgeht Static-Cache) ────────────────────────────

@app.route('/logo.png')
def serve_logo():
    if LOGO_URL:
        from flask import redirect as _redir
        return _redir(LOGO_URL, code=302)
    logo_path = os.path.join(os.path.dirname(__file__), 'static', 'logo.png')
    return send_file(logo_path, mimetype='image/png',
                     max_age=3600, conditional=True)


# ─── Admin: DB-Backup herunterladen ──────────────────────────────────────────

@app.route('/admin/backup/herunterladen')
@login_required
def admin_backup_herunterladen():
    if session.get('rolle') != 'admin':
        flash('Keine Berechtigung.', 'danger')
        return redirect(url_for('dashboard'))
    if not os.path.exists(DATABASE):
        flash('Datenbank nicht gefunden.', 'danger')
        return redirect(url_for('admin'))
    heute = date.today().isoformat()
    return send_file(DATABASE, as_attachment=True,
                     download_name=f'brewery_backup_{heute}.db',
                     mimetype='application/x-sqlite3')


# ─── Routes: Dashboard ────────────────────────────────────────────────────────

@app.route('/dashboard')
@login_required
def dashboard():
    jahr      = request.args.get('jahr', date.today().year, type=int)
    is_admin  = session.get('rolle') == 'admin'
    is_manager = session.get('rolle') in ('admin', 'verkaufsleiter')
    ma_filter = request.args.get('ma', '', type=str)
    ma_clause = "AND a.mitarbeiter_id = ?" if ma_filter else ""
    ma_params = (ma_filter,) if ma_filter else ()

    # KW-Daten (Wochenübersicht)
    # Subquery: Kisten pro Aktivität voraggregieren → verhindert Duplikation von anzahl_displays
    BP = "(SELECT aktivitaet_id, SUM(kisten_anzahl) AS kisten_total FROM bestellposition GROUP BY aktivitaet_id)"

    if is_manager:
        kw_data = query(f'''
            SELECT strftime('%W', a.datum) AS kw,
                   CAST(strftime('%W', a.datum) AS INTEGER) AS kw_int,
                   SUM(a.anzahl_displays) AS displays,
                   COALESCE(SUM(b.kisten_total), 0) AS kisten,
                   COUNT(a.id) AS besuche
            FROM aktivitaet a
            LEFT JOIN {BP} b ON b.aktivitaet_id = a.id
            WHERE strftime('%Y', a.datum) = ? {ma_clause}
            GROUP BY kw
            ORDER BY kw
        ''', (str(jahr),) + ma_params)
    else:
        kw_data = query(f'''
            SELECT strftime('%W', a.datum) AS kw,
                   CAST(strftime('%W', a.datum) AS INTEGER) AS kw_int,
                   SUM(a.anzahl_displays) AS displays,
                   COALESCE(SUM(b.kisten_total), 0) AS kisten,
                   COUNT(a.id) AS besuche
            FROM aktivitaet a
            LEFT JOIN {BP} b ON b.aktivitaet_id = a.id
            WHERE strftime('%Y', a.datum) = ? AND a.mitarbeiter_id = ?
            GROUP BY kw
            ORDER BY kw
        ''', (str(jahr), session['user_id']))

    # Jahresgesamtwerte
    if is_manager:
        jahres = query(f'''
            SELECT SUM(a.anzahl_displays) AS displays,
                   COALESCE(SUM(b.kisten_total), 0) AS kisten,
                   COUNT(a.id) AS besuche,
                   COUNT(DISTINCT a.mitarbeiter_id) AS mitarbeiter_aktiv
            FROM aktivitaet a
            LEFT JOIN {BP} b ON b.aktivitaet_id = a.id
            WHERE strftime('%Y', a.datum) = ? {ma_clause}
        ''', (str(jahr),) + ma_params, one=True)
    else:
        jahres = query(f'''
            SELECT SUM(a.anzahl_displays) AS displays,
                   COALESCE(SUM(b.kisten_total), 0) AS kisten,
                   COUNT(a.id) AS besuche
            FROM aktivitaet a
            LEFT JOIN {BP} b ON b.aktivitaet_id = a.id
            WHERE strftime('%Y', a.datum) = ? AND a.mitarbeiter_id = ?
        ''', (str(jahr), session['user_id']), one=True)

    # Top Biersorten – direkt über bestellposition, kein Display-Problem hier
    if is_manager:
        top_bier = query(f'''
            SELECT bs.name, SUM(bp.kisten_anzahl) AS kisten
            FROM bestellposition bp
            JOIN biersorte bs ON bs.id = bp.biersorte_id
            JOIN aktivitaet a ON a.id = bp.aktivitaet_id
            WHERE strftime('%Y', a.datum) = ? {ma_clause}
            GROUP BY bs.id ORDER BY kisten DESC LIMIT 6
        ''', (str(jahr),) + ma_params)
    else:
        top_bier = query('''
            SELECT bs.name, SUM(bp.kisten_anzahl) AS kisten
            FROM bestellposition bp
            JOIN biersorte bs ON bs.id = bp.biersorte_id
            JOIN aktivitaet a ON a.id = bp.aktivitaet_id
            WHERE strftime('%Y', a.datum) = ? AND a.mitarbeiter_id = ?
            GROUP BY bs.id ORDER BY kisten DESC LIMIT 6
        ''', (str(jahr), session['user_id']))

    # Mitarbeiter-Ranking (Manager-Sicht, nur ohne Einzelfilter)
    rep_stats = []
    if is_manager and not ma_filter:
        rep_stats = query(f'''
            SELECT m.name, m.kuerzel,
                   SUM(a.anzahl_displays) AS displays,
                   COALESCE(SUM(b.kisten_total), 0) AS kisten,
                   COUNT(a.id) AS besuche
            FROM mitarbeiter m
            JOIN aktivitaet a ON a.mitarbeiter_id = m.id
            LEFT JOIN {BP} b ON b.aktivitaet_id = a.id
            WHERE strftime('%Y', a.datum) = ?
            GROUP BY m.id ORDER BY kisten DESC
        ''', (str(jahr),))

    # Letzte Aktivitäten
    if is_manager:
        letzte = query(f'''
            SELECT a.id, a.datum, m.name AS mitarbeiter, v.name AS verkaufsstelle,
                   a.anzahl_displays, COALESCE(SUM(b.kisten_anzahl), 0) AS kisten
            FROM aktivitaet a
            JOIN mitarbeiter m ON m.id = a.mitarbeiter_id
            JOIN verkaufsstelle v ON v.id = a.verkaufsstelle_id
            LEFT JOIN bestellposition b ON b.aktivitaet_id = a.id
            WHERE strftime('%Y', a.datum) = ? {ma_clause}
            GROUP BY a.id ORDER BY a.datum DESC, a.erstellt_am DESC LIMIT 10
        ''', (str(jahr),) + ma_params)
    else:
        letzte = query('''
            SELECT a.id, a.datum, m.name AS mitarbeiter, v.name AS verkaufsstelle,
                   a.anzahl_displays, COALESCE(SUM(b.kisten_anzahl), 0) AS kisten
            FROM aktivitaet a
            JOIN mitarbeiter m ON m.id = a.mitarbeiter_id
            JOIN verkaufsstelle v ON v.id = a.verkaufsstelle_id
            LEFT JOIN bestellposition b ON b.aktivitaet_id = a.id
            WHERE strftime('%Y', a.datum) = ? AND a.mitarbeiter_id = ?
            GROUP BY a.id ORDER BY a.datum DESC, a.erstellt_am DESC LIMIT 10
        ''', (str(jahr), session['user_id']))

    alle_ma = query("SELECT id, name FROM mitarbeiter WHERE rolle IN ('rep','verkaufsleiter') ORDER BY name") if is_manager else []

    verfuegbare_jahre = [r[0] for r in query(
        "SELECT DISTINCT CAST(strftime('%Y', datum) AS INTEGER) FROM aktivitaet ORDER BY 1 DESC"
    )]
    if not verfuegbare_jahre:
        verfuegbare_jahre = [date.today().year]
    if date.today().year not in verfuegbare_jahre:
        verfuegbare_jahre.insert(0, date.today().year)

    chart_kw   = [f"KW {int(r['kw_int']):02d}" for r in kw_data]
    chart_disp = [r['displays'] or 0 for r in kw_data]
    chart_kist = [r['kisten'] or 0 for r in kw_data]
    bier_namen = [r['name'] for r in top_bier]
    bier_kist  = [r['kisten'] for r in top_bier]

    return render_template('dashboard.html',
        jahr=jahr, kw_data=kw_data, jahres=jahres,
        rep_stats=rep_stats, letzte=letzte,
        verfuegbare_jahre=verfuegbare_jahre,
        chart_kw=json.dumps(chart_kw),
        chart_disp=json.dumps(chart_disp),
        chart_kist=json.dumps(chart_kist),
        bier_namen=json.dumps(bier_namen),
        bier_kist=json.dumps(bier_kist),
        is_admin=is_admin,
        is_manager=is_manager,
        ma_filter=ma_filter,
        alle_ma=alle_ma,
    )


# ─── API: Letzter Besuch ─────────────────────────────────────────────────────

@app.route('/api/letzter-besuch/<int:vs_id>')
@login_required
def api_letzter_besuch(vs_id):
    """Gibt die letzten 3 Besuche bei einer Verkaufsstelle zurück.
    Manager sehen alle Reps, normale Reps nur ihre eigenen."""
    ma_id = session['user_id']
    is_manager = session.get('rolle') in ('admin', 'verkaufsleiter')
    if is_manager:
        rows = query('''
            SELECT a.datum, a.anzahl_displays, a.notizen,
                   m.name AS mitarbeiter,
                   COALESCE((SELECT SUM(bp.kisten_anzahl) FROM bestellposition bp
                             WHERE bp.aktivitaet_id = a.id), 0) AS kisten_gesamt
            FROM aktivitaet a
            JOIN mitarbeiter m ON m.id = a.mitarbeiter_id
            WHERE a.verkaufsstelle_id = ?
            ORDER BY a.datum DESC, a.id DESC LIMIT 3
        ''', (vs_id,))
    else:
        rows = query('''
            SELECT a.datum, a.anzahl_displays, a.notizen,
                   NULL AS mitarbeiter,
                   COALESCE((SELECT SUM(bp.kisten_anzahl) FROM bestellposition bp
                             WHERE bp.aktivitaet_id = a.id), 0) AS kisten_gesamt
            FROM aktivitaet a
            WHERE a.verkaufsstelle_id = ? AND a.mitarbeiter_id = ?
            ORDER BY a.datum DESC, a.id DESC LIMIT 3
        ''', (vs_id, ma_id))
    if not rows:
        return jsonify({'besuche': []})
    besuche = []
    for row in rows:
        try:
            tage = (date.today() - date.fromisoformat(row['datum'])).days
        except Exception:
            tage = '?'
        try:
            d = date.fromisoformat(row['datum'])
            datum_fmt = f"{d.day:02d}.{d.month:02d}.{d.year}"
        except Exception:
            datum_fmt = row['datum']
        notizen = (row['notizen'] or '').strip()
        if len(notizen) > 60:
            notizen = notizen[:60] + '…'
        besuche.append({
            'datum':       datum_fmt,
            'tage_ago':    tage,
            'displays':    row['anzahl_displays'] or 0,
            'kisten':      row['kisten_gesamt']   or 0,
            'notizen':     notizen,
            'mitarbeiter': row['mitarbeiter'] or None
        })
    return jsonify({'besuche': besuche})


# ─── Routes: Aktivitäten ──────────────────────────────────────────────────────

@app.route('/aktivitaet/neu', methods=['GET', 'POST'])
@login_required
def neue_aktivitaet():
    today = date.today().isoformat()

    # Reps und VKL sehen nur ihre zugeordneten Verkaufsstellen (wenn Zuordnung gesetzt)
    if session.get('rolle') in ('rep', 'verkaufsleiter'):
        assigned = query(
            "SELECT verkaufsstelle_id FROM mitarbeiter_verkaufsstelle WHERE mitarbeiter_id=?",
            (session['user_id'],)
        )
        if assigned:
            vs_ids = [r['verkaufsstelle_id'] for r in assigned]
            ph = ','.join('?' * len(vs_ids))
            verkaufsstellen = query(
                f"SELECT * FROM verkaufsstelle WHERE aktiv=1 AND id IN ({ph}) ORDER BY name",
                vs_ids
            )
        else:
            verkaufsstellen = query("SELECT * FROM verkaufsstelle WHERE aktiv=1 ORDER BY name")
    else:
        verkaufsstellen = query("SELECT * FROM verkaufsstelle WHERE aktiv=1 ORDER BY name")

    # Aktive Vertretungen: zusätzliche VS des Abwesenden einblenden
    vertretungs_gruppen = []  # [{name, vs: [...]}, ...]
    aktive_vtr = query(
        '''SELECT v.id, v.abwesender_id, m.name AS abwesender_name
           FROM vertretung v
           JOIN mitarbeiter m ON m.id = v.abwesender_id
           WHERE v.vertreter_id = ? AND v.von <= ? AND v.bis >= ?''',
        (session['user_id'], today, today)
    )
    eigene_vs_ids = {vs['id'] for vs in verkaufsstellen}
    for vtr in aktive_vtr:
        vtr_assigned = query(
            "SELECT verkaufsstelle_id FROM mitarbeiter_verkaufsstelle WHERE mitarbeiter_id=?",
            (vtr['abwesender_id'],)
        )
        if vtr_assigned:
            vtr_ids = [r['verkaufsstelle_id'] for r in vtr_assigned]
            ph = ','.join('?' * len(vtr_ids))
            vtr_vs = query(
                f"SELECT * FROM verkaufsstelle WHERE aktiv=1 AND id IN ({ph}) ORDER BY name",
                vtr_ids
            )
        else:
            vtr_vs = []
        # Nur VS hinzufügen, die nicht schon im eigenen Gebiet sind
        extra = [v for v in vtr_vs if v['id'] not in eigene_vs_ids]
        if extra:
            vertretungs_gruppen.append({'name': vtr['abwesender_name'], 'vs': extra})

    biersorten      = query("SELECT * FROM biersorte      WHERE aktiv=1 ORDER BY name")
    displaysorte    = query("SELECT * FROM displaysorte   WHERE aktiv=1 ORDER BY name")

    if request.method == 'POST':
        datum   = request.form.get('datum')
        vs_id   = request.form.get('verkaufsstelle_id')
        notizen = request.form.get('notizen', '')

        foto_file = request.files.get('foto')
        if not foto_file or not foto_file.filename:
            flash('Bitte ein Foto hochladen – das Foto ist ein Pflichtfeld.', 'danger')
            return render_template('neue_aktivitaet.html',
                verkaufsstellen=verkaufsstellen, biersorten=biersorten,
                displaysorte=displaysorte, heute=date.today().isoformat())

        if not datum or not vs_id:
            flash('Datum und Verkaufsstelle sind Pflichtfelder.', 'danger')
            return render_template('neue_aktivitaet.html',
                verkaufsstellen=verkaufsstellen, biersorten=biersorten,
                displaysorte=displaysorte, heute=date.today().isoformat())

        # Displaypositionen sammeln + Gesamtzahl berechnen
        anzahl_displays  = 0
        disp_positionen  = []
        for ds in displaysorte:
            menge_str = request.form.get(f'disp_{ds["id"]}', '').strip()
            if menge_str and menge_str.isdigit() and int(menge_str) > 0:
                menge = int(menge_str)
                anzahl_displays += menge
                disp_positionen.append((ds['id'], menge))

        # Foto verarbeiten
        foto_pfad = None
        if foto_file and foto_file.filename and allowed_file(foto_file.filename):
            ext = foto_file.filename.rsplit('.', 1)[1].lower()
            dateiname = f"akt_{uuid.uuid4().hex}.{ext}"
            foto_file.save(os.path.join(UPLOAD_FOLDER, dateiname))
            foto_pfad = dateiname

        akt_id = execute(
            "INSERT INTO aktivitaet (datum, mitarbeiter_id, verkaufsstelle_id, anzahl_displays, notizen, foto_pfad) VALUES (?,?,?,?,?,?)",
            (datum, session['user_id'], vs_id, anzahl_displays, notizen, foto_pfad)
        )

        # Displaypositionen speichern
        for ds_id, menge in disp_positionen:
            execute(
                "INSERT INTO displayposition (aktivitaet_id, displaysorte_id, anzahl) VALUES (?,?,?)",
                (akt_id, ds_id, menge)
            )

        # Bestellpositionen speichern
        for bier in biersorten:
            menge = request.form.get(f'bier_{bier["id"]}', '').strip()
            if menge and menge.isdigit() and int(menge) > 0:
                execute(
                    "INSERT INTO bestellposition (aktivitaet_id, biersorte_id, kisten_anzahl) VALUES (?,?,?)",
                    (akt_id, bier['id'], int(menge))
                )

        if foto_pfad:
            cleanup_alte_fotos()

        flash('Aktivität erfolgreich gespeichert!', 'success')
        return redirect(url_for('aktivitaeten_liste'))

    preselect_vs = request.args.get('vs_id', '', type=str)
    return render_template('neue_aktivitaet.html',
        verkaufsstellen=verkaufsstellen, biersorten=biersorten,
        displaysorte=displaysorte, vertretungs_gruppen=vertretungs_gruppen,
        heute=date.today().isoformat(), preselect_vs=preselect_vs)


@app.route('/aktivitaeten')
@login_required
def aktivitaeten_liste():
    is_admin   = session.get('rolle') == 'admin'
    is_manager = session.get('rolle') in ('admin', 'verkaufsleiter')
    jahr       = request.args.get('jahr', date.today().year, type=int)
    kw_filter  = request.args.get('kw',    '', type=str)
    mo_filter  = request.args.get('monat', '', type=str)
    mo_ids     = [x.strip().zfill(2) for x in mo_filter.split(',') if x.strip()] if mo_filter else []
    ma_filter  = request.args.get('ma',    '', type=str)
    ma_ids     = [x.strip() for x in ma_filter.split(',') if x.strip()] if ma_filter else []
    vs_filter  = request.args.get('vs',    '', type=str)
    vs_ids     = [x.strip() for x in vs_filter.split(',') if x.strip()] if vs_filter else []
    typ_filter = request.args.get('typ',   '', type=str)   # kommagetrennte Typen
    typ_ids    = [x.strip() for x in typ_filter.split(',') if x.strip()] if typ_filter else []

    sql = '''
        SELECT a.id, a.datum, m.name AS mitarbeiter, m.id AS mitarbeiter_id,
               v.name AS verkaufsstelle, v.id AS verkaufsstelle_id,
               v.ort, v.typ, a.anzahl_displays, a.notizen, a.erstellt_am,
               a.foto_pfad,
               COALESCE(SUM(b.kisten_anzahl), 0) AS kisten_gesamt
        FROM aktivitaet a
        JOIN mitarbeiter m ON m.id = a.mitarbeiter_id
        JOIN verkaufsstelle v ON v.id = a.verkaufsstelle_id
        LEFT JOIN bestellposition b ON b.aktivitaet_id = a.id
        WHERE 1=1
    '''
    params = []

    vs_history_mode = is_manager and bool(vs_ids)
    if not vs_history_mode:
        sql += " AND strftime('%Y', a.datum) = ?"
        params.append(str(jahr))

    if not is_manager:
        sql += " AND a.mitarbeiter_id = ?"
        params.append(session['user_id'])
    elif ma_ids:
        _ph = ','.join('?' * len(ma_ids))
        sql += f" AND a.mitarbeiter_id IN ({_ph})"
        params.extend(ma_ids)

    if is_manager and vs_ids:
        _ph = ','.join('?' * len(vs_ids))
        sql += f" AND a.verkaufsstelle_id IN ({_ph})"
        params.extend(vs_ids)

    if mo_ids:
        _ph = ','.join('?' * len(mo_ids))
        sql += f" AND strftime('%m', a.datum) IN ({_ph})"
        params.extend(mo_ids)

    if kw_filter:
        sql += " AND CAST(strftime('%W', a.datum) AS INTEGER) = ?"
        params.append(int(kw_filter))

    if typ_ids:
        _ph = ','.join('?' * len(typ_ids))
        sql += f" AND v.typ IN ({_ph})"
        params.extend(typ_ids)

    sql += " GROUP BY a.id ORDER BY a.datum DESC, a.erstellt_am DESC"

    aktivitaeten = query(sql, params)

    # Bestellpositionen für jede Aktivität
    detail = {}
    for a in aktivitaeten:
        positionen = query('''
            SELECT bs.name, bp.kisten_anzahl, bs.einheit
            FROM bestellposition bp JOIN biersorte bs ON bs.id = bp.biersorte_id
            WHERE bp.aktivitaet_id = ?
        ''', (a['id'],))
        detail[a['id']] = positionen

    # Displaypositionen für jede Aktivität
    disp_detail = {}
    for a in aktivitaeten:
        dp = query('''
            SELECT ds.name, dp.anzahl
            FROM displayposition dp JOIN displaysorte ds ON ds.id = dp.displaysorte_id
            WHERE dp.aktivitaet_id = ? AND dp.anzahl > 0
            ORDER BY ds.name
        ''', (a['id'],))
        if dp:
            disp_detail[a['id']] = dp

    alle_ma = query("SELECT id, name FROM mitarbeiter WHERE rolle IN ('rep','verkaufsleiter') ORDER BY name") if is_manager else []
    # Alle VS für Dropdown (inkl. inaktive – für historische Suche)
    alle_vs = query(
        "SELECT id, name, ort, aktiv FROM verkaufsstelle ORDER BY aktiv DESC, name"
    ) if is_manager else []
    alle_typen = [r[0] for r in query(
        "SELECT DISTINCT typ FROM verkaufsstelle WHERE typ IS NOT NULL AND typ != '' ORDER BY typ"
    )]
    jahre = [r[0] for r in query("SELECT DISTINCT CAST(strftime('%Y', datum) AS INTEGER) FROM aktivitaet ORDER BY 1 DESC")]
    if not jahre:
        jahre = [date.today().year]

    return render_template('aktivitaeten.html',
        aktivitaeten=aktivitaeten, detail=detail, disp_detail=disp_detail,
        jahr=jahr, jahre=jahre, kw_filter=kw_filter,
        mo_filter=mo_filter, mo_ids=mo_ids,
        ma_filter=ma_filter, ma_ids=ma_ids,
        vs_filter=vs_filter, vs_ids=vs_ids, vs_history_mode=vs_history_mode,
        typ_filter=typ_filter, typ_ids=typ_ids, alle_typen=alle_typen,
        alle_ma=alle_ma, alle_vs=alle_vs,
        is_admin=is_admin, is_manager=is_manager)


@app.route('/aktivitaet/<int:akt_id>/loeschen', methods=['POST'])
@login_required
def aktivitaet_loeschen(akt_id):
    a = query("SELECT * FROM aktivitaet WHERE id=?", (akt_id,), one=True)
    if not a:
        flash('Aktivität nicht gefunden.', 'danger')
        return redirect(url_for('aktivitaeten_liste'))

    if session.get('rolle') != 'admin' and a['mitarbeiter_id'] != session['user_id']:
        flash('Keine Berechtigung.', 'danger')
        return redirect(url_for('aktivitaeten_liste'))

    # Foto-Datei mitlöschen
    if a['foto_pfad']:
        foto_path = os.path.join(UPLOAD_FOLDER, a['foto_pfad'])
        if os.path.exists(foto_path):
            os.remove(foto_path)

    execute("DELETE FROM bestellposition  WHERE aktivitaet_id=?", (akt_id,))
    execute("DELETE FROM displayposition  WHERE aktivitaet_id=?", (akt_id,))
    execute("DELETE FROM aktivitaet       WHERE id=?",            (akt_id,))
    flash('Aktivität gelöscht.', 'success')
    return redirect(url_for('aktivitaeten_liste'))


# ─── Excel-Hilfsfunktion (für Route + Auto-Export) ────────────────────────────

def _build_excel_bytes(jahr: int, is_admin: bool = True, mitarbeiter_id: int = None) -> bytes:
    """Erstellt die Excel-Auswertung und gibt sie als Bytes zurück."""
    wb = openpyxl.Workbook()

    HEADER_FILL = PatternFill("solid", fgColor="1a3a5c")
    SUB_FILL    = PatternFill("solid", fgColor="2e6da4")
    ALT_FILL    = PatternFill("solid", fgColor="eef4fb")
    HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
    SUB_FONT    = Font(bold=True, color="FFFFFF", size=10)
    TITLE_FONT  = Font(bold=True, size=14, color="1a3a5c")
    BOLD        = Font(bold=True)
    BORDER      = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC'),
    )
    CENTER = Alignment(horizontal='center', vertical='center')
    LEFT   = Alignment(horizontal='left', vertical='center')

    def style_header(ws, row, cols):
        for c in range(1, cols+1):
            cell = ws.cell(row=row, column=c)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = CENTER
            cell.border = BORDER

    def style_subheader(ws, row, cols):
        for c in range(1, cols+1):
            cell = ws.cell(row=row, column=c)
            cell.fill = SUB_FILL
            cell.font = SUB_FONT
            cell.alignment = CENTER
            cell.border = BORDER

    # ── Sheet 1: KW-Übersicht ──────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = f"KW-Übersicht {jahr}"
    ws1.column_dimensions['A'].width = 10
    ws1.column_dimensions['B'].width = 18
    ws1.column_dimensions['C'].width = 18
    ws1.column_dimensions['D'].width = 15

    ws1.merge_cells('A1:D1')
    ws1['A1'] = f"Wochenübersicht {jahr} – Displays & Kisten"
    ws1['A1'].font = TITLE_FONT
    ws1['A1'].alignment = CENTER

    headers = ['Kalenderwoche', 'Anzahl Displays', 'Kisten gesamt', 'Besuche']
    for col, h in enumerate(headers, 1):
        ws1.cell(row=2, column=col, value=h)
    style_header(ws1, 2, 4)

    _BP = "(SELECT aktivitaet_id, SUM(kisten_anzahl) AS kisten_total FROM bestellposition GROUP BY aktivitaet_id)"
    if is_admin:
        kw_data = query(f'''
            SELECT CAST(strftime('%W', a.datum) AS INTEGER) AS kw,
                   SUM(a.anzahl_displays) AS displays,
                   COALESCE(SUM(b.kisten_total), 0) AS kisten,
                   COUNT(a.id) AS besuche
            FROM aktivitaet a
            LEFT JOIN {_BP} b ON b.aktivitaet_id = a.id
            WHERE strftime('%Y', a.datum) = ?
            GROUP BY kw ORDER BY kw
        ''', (str(jahr),))
    else:
        kw_data = query(f'''
            SELECT CAST(strftime('%W', a.datum) AS INTEGER) AS kw,
                   SUM(a.anzahl_displays) AS displays,
                   COALESCE(SUM(b.kisten_total), 0) AS kisten,
                   COUNT(a.id) AS besuche
            FROM aktivitaet a
            LEFT JOIN {_BP} b ON b.aktivitaet_id = a.id
            WHERE strftime('%Y', a.datum) = ? AND a.mitarbeiter_id = ?
            GROUP BY kw ORDER BY kw
        ''', (str(jahr), mitarbeiter_id))

    total_disp = total_kist = total_bes = 0
    for i, row in enumerate(kw_data):
        r = i + 3
        ws1.cell(r, 1, f"KW {row['kw']:02d}")
        ws1.cell(r, 2, row['displays'] or 0)
        ws1.cell(r, 3, row['kisten'] or 0)
        ws1.cell(r, 4, row['besuche'] or 0)
        total_disp += row['displays'] or 0
        total_kist += row['kisten'] or 0
        total_bes  += row['besuche'] or 0
        fill = ALT_FILL if i % 2 == 0 else None
        for c in range(1, 5):
            cell = ws1.cell(r, c)
            cell.alignment = CENTER
            cell.border = BORDER
            if fill:
                cell.fill = fill

    # Summenzeile
    sr = len(kw_data) + 3
    ws1.cell(sr, 1, 'GESAMT').font = BOLD
    ws1.cell(sr, 2, total_disp).font = BOLD
    ws1.cell(sr, 3, total_kist).font = BOLD
    ws1.cell(sr, 4, total_bes).font = BOLD
    for c in range(1, 5):
        ws1.cell(sr, c).fill = PatternFill("solid", fgColor="FFD700")
        ws1.cell(sr, c).alignment = CENTER
        ws1.cell(sr, c).border = BORDER

    # ── Sheet 2: Jahreswerte nach Mitarbeiter (nur Admin) ─────────────────
    if is_admin:
        ws2 = wb.create_sheet(f"Mitarbeiter {jahr}")
        ws2.column_dimensions['A'].width = 22
        ws2.column_dimensions['B'].width = 10
        ws2.column_dimensions['C'].width = 18
        ws2.column_dimensions['D'].width = 16
        ws2.column_dimensions['E'].width = 12

        ws2.merge_cells('A1:E1')
        ws2['A1'] = f"Jahresübersicht nach Mitarbeiter {jahr}"
        ws2['A1'].font = TITLE_FONT
        ws2['A1'].alignment = CENTER

        h2 = ['Mitarbeiter', 'Kürzel', 'Displays gesamt', 'Kisten gesamt', 'Besuche']
        for col, h in enumerate(h2, 1):
            ws2.cell(2, col, h)
        style_header(ws2, 2, 5)

        ma_data = query(f'''
            SELECT m.name, m.kuerzel,
                   SUM(a.anzahl_displays) AS displays,
                   COALESCE(SUM(b.kisten_total), 0) AS kisten,
                   COUNT(a.id) AS besuche
            FROM mitarbeiter m
            JOIN aktivitaet a ON a.mitarbeiter_id = m.id
            LEFT JOIN {_BP} b ON b.aktivitaet_id = a.id
            WHERE strftime('%Y', a.datum) = ?
            GROUP BY m.id ORDER BY kisten DESC
        ''', (str(jahr),))

        for i, row in enumerate(ma_data):
            r = i + 3
            ws2.cell(r, 1, row['name'])
            ws2.cell(r, 2, row['kuerzel'])
            ws2.cell(r, 3, row['displays'] or 0)
            ws2.cell(r, 4, row['kisten'] or 0)
            ws2.cell(r, 5, row['besuche'] or 0)
            fill = ALT_FILL if i % 2 == 0 else None
            for c in range(1, 6):
                cell = ws2.cell(r, c)
                cell.border = BORDER
                cell.alignment = CENTER if c > 1 else LEFT
                if fill:
                    cell.fill = fill

    # ── Sheet 3: Alle Aktivitäten ─────────────────────────────────────────
    ws3 = wb.create_sheet("Aktivitäten-Detail")
    cols3 = ['Datum', 'KW', 'Mitarbeiter', 'Verkaufsstelle', 'Ort', 'Typ',
             'Displays', 'Produkt', 'Kisten', 'Notizen']
    widths = [14, 6, 20, 28, 16, 16, 10, 20, 10, 35]
    for i, w in enumerate(widths, 1):
        ws3.column_dimensions[get_column_letter(i)].width = w

    ws3.merge_cells(f'A1:{get_column_letter(len(cols3))}1')
    ws3['A1'] = f"Aktivitäten-Detail {jahr}"
    ws3['A1'].font = TITLE_FONT
    ws3['A1'].alignment = CENTER

    for col, h in enumerate(cols3, 1):
        ws3.cell(2, col, h)
    style_header(ws3, 2, len(cols3))

    if is_admin:
        aktivitaeten = query('''
            SELECT a.id, a.datum, m.name AS mitarbeiter,
                   v.name AS verkaufsstelle, v.ort, v.typ,
                   a.anzahl_displays, a.notizen
            FROM aktivitaet a
            JOIN mitarbeiter m ON m.id = a.mitarbeiter_id
            JOIN verkaufsstelle v ON v.id = a.verkaufsstelle_id
            WHERE strftime('%Y', a.datum) = ?
            ORDER BY a.datum DESC
        ''', (str(jahr),))
    else:
        aktivitaeten = query('''
            SELECT a.id, a.datum, m.name AS mitarbeiter,
                   v.name AS verkaufsstelle, v.ort, v.typ,
                   a.anzahl_displays, a.notizen
            FROM aktivitaet a
            JOIN mitarbeiter m ON m.id = a.mitarbeiter_id
            JOIN verkaufsstelle v ON v.id = a.verkaufsstelle_id
            WHERE strftime('%Y', a.datum) = ? AND a.mitarbeiter_id = ?
            ORDER BY a.datum DESC
        ''', (str(jahr), mitarbeiter_id))

    r = 3
    for i, a in enumerate(aktivitaeten):
        positionen = query('''
            SELECT bs.name, bp.kisten_anzahl
            FROM bestellposition bp JOIN biersorte bs ON bs.id = bp.biersorte_id
            WHERE bp.aktivitaet_id = ?
        ''', (a['id'],))

        import datetime as dt
        d = dt.date.fromisoformat(a['datum'])
        kw = int(d.strftime('%W'))
        fill = ALT_FILL if i % 2 == 0 else None

        if not positionen:
            ws3.cell(r, 1, a['datum'])
            ws3.cell(r, 2, f"KW {kw:02d}")
            ws3.cell(r, 3, a['mitarbeiter'])
            ws3.cell(r, 4, a['verkaufsstelle'])
            ws3.cell(r, 5, a['ort'])
            ws3.cell(r, 6, a['typ'])
            ws3.cell(r, 7, a['anzahl_displays'])
            ws3.cell(r, 8, '–')
            ws3.cell(r, 9, 0)
            ws3.cell(r, 10, a['notizen'] or '')
            for c in range(1, 11):
                ws3.cell(r, c).border = BORDER
                if fill:
                    ws3.cell(r, c).fill = fill
            r += 1
        else:
            for j, pos in enumerate(positionen):
                ws3.cell(r, 1, a['datum'] if j == 0 else '')
                ws3.cell(r, 2, f"KW {kw:02d}" if j == 0 else '')
                ws3.cell(r, 3, a['mitarbeiter'] if j == 0 else '')
                ws3.cell(r, 4, a['verkaufsstelle'] if j == 0 else '')
                ws3.cell(r, 5, a['ort'] if j == 0 else '')
                ws3.cell(r, 6, a['typ'] if j == 0 else '')
                ws3.cell(r, 7, a['anzahl_displays'] if j == 0 else '')
                ws3.cell(r, 8, pos['name'])
                ws3.cell(r, 9, pos['kisten_anzahl'])
                ws3.cell(r, 10, a['notizen'] if j == 0 else '')
                for c in range(1, 11):
                    ws3.cell(r, c).border = BORDER
                    if fill:
                        ws3.cell(r, c).fill = fill
                r += 1

    # ── Sheet 4: Produkt-Übersicht ─────────────────────────────────────────
    ws4 = wb.create_sheet(f"Produkte {jahr}")
    ws4.column_dimensions['A'].width = 22
    ws4.column_dimensions['B'].width = 20
    ws4.column_dimensions['C'].width = 16

    ws4.merge_cells('A1:C1')
    ws4['A1'] = f"Bestellte Produkte {jahr}"
    ws4['A1'].font = TITLE_FONT
    ws4['A1'].alignment = CENTER

    for col, h in enumerate(['Produkt', 'Einheit', 'Kisten gesamt'], 1):
        ws4.cell(2, col, h)
    style_header(ws4, 2, 3)

    if is_admin:
        bier_data = query('''
            SELECT bs.name, bs.einheit, SUM(bp.kisten_anzahl) AS kisten
            FROM bestellposition bp
            JOIN biersorte bs ON bs.id = bp.biersorte_id
            JOIN aktivitaet a ON a.id = bp.aktivitaet_id
            WHERE strftime('%Y', a.datum) = ?
            GROUP BY bs.id ORDER BY kisten DESC
        ''', (str(jahr),))
    else:
        bier_data = query('''
            SELECT bs.name, bs.einheit, SUM(bp.kisten_anzahl) AS kisten
            FROM bestellposition bp
            JOIN biersorte bs ON bs.id = bp.biersorte_id
            JOIN aktivitaet a ON a.id = bp.aktivitaet_id
            WHERE strftime('%Y', a.datum) = ? AND a.mitarbeiter_id = ?
            GROUP BY bs.id ORDER BY kisten DESC
        ''', (str(jahr), mitarbeiter_id))

    for i, row in enumerate(bier_data):
        r2 = i + 3
        ws4.cell(r2, 1, row['name'])
        ws4.cell(r2, 2, row['einheit'])
        ws4.cell(r2, 3, row['kisten'])
        fill = ALT_FILL if i % 2 == 0 else None
        for c in range(1, 4):
            ws4.cell(r2, c).border = BORDER
            ws4.cell(r2, c).alignment = CENTER if c > 1 else LEFT
            if fill:
                ws4.cell(r2, c).fill = fill

    # Output
    output = io.BytesIO()
    # Output als Bytes
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.read()


# ─── Route: Excel Export ──────────────────────────────────────────────────────

@app.route('/export/excel')
@login_required
def export_excel():
    jahr     = request.args.get('jahr', date.today().year, type=int)
    is_admin = session.get('rolle') in ('admin', 'verkaufsleiter')
    ma_id    = None if is_admin else session.get('user_id')
    data     = _build_excel_bytes(jahr, is_admin=is_admin, mitarbeiter_id=ma_id)
    fname    = f"Aktions_Tracker_{jahr}.xlsx"
    return send_file(io.BytesIO(data), as_attachment=True, download_name=fname,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ─── Routes: Admin ────────────────────────────────────────────────────────────

@app.route('/admin')
@admin_required
def admin():
    mitarbeiter     = query("SELECT * FROM mitarbeiter ORDER BY rolle, name")
    verkaufsstellen = query("SELECT * FROM verkaufsstelle ORDER BY aktiv DESC, name")
    biersorten      = query("SELECT * FROM biersorte ORDER BY name")
    displaysorte    = query("SELECT * FROM displaysorte ORDER BY name")
    mail_konfiguriert = bool(MAIL_SERVER and MAIL_USERNAME)

    # Zuordnungen: {mitarbeiter_id: set(verkaufsstelle_id, ...)}
    zuordnungen_raw = query("SELECT mitarbeiter_id, verkaufsstelle_id FROM mitarbeiter_verkaufsstelle")
    zuordnungen = {}
    for z in zuordnungen_raw:
        zuordnungen.setdefault(z['mitarbeiter_id'], set()).add(z['verkaufsstelle_id'])

    # Besitzer je VS: {vs_id: mitarbeiter_name} – für Anzeige im Zuordnungs-Modal
    ma_namen = {m['id']: m['name'] for m in mitarbeiter}
    vs_besitzer = {}
    for ma_id_loop, vs_set in zuordnungen.items():
        for vs_id in vs_set:
            vs_besitzer[vs_id] = ma_namen.get(ma_id_loop, '')

    # Vertretungsregelungen
    vertretungen = query('''
        SELECT v.id, v.von, v.bis,
               a.name AS abwesender, r.name AS vertreter
        FROM vertretung v
        JOIN mitarbeiter a ON a.id = v.abwesender_id
        JOIN mitarbeiter r ON r.id = v.vertreter_id
        ORDER BY v.von DESC
    ''')
    # Alle Außendienst-Mitarbeiter für Dropdowns
    alle_ad = query("SELECT id, name FROM mitarbeiter WHERE rolle IN ('rep','verkaufsleiter') ORDER BY name")

    return render_template('admin.html',
        mitarbeiter=mitarbeiter,
        verkaufsstellen=verkaufsstellen,
        biersorten=biersorten,
        displaysorte=displaysorte,
        zuordnungen=zuordnungen,
        vs_besitzer=vs_besitzer,
        vertretungen=vertretungen,
        alle_ad=alle_ad,
        mail_konfiguriert=mail_konfiguriert)


@app.route('/admin/mitarbeiter/neu', methods=['POST'])
@admin_required
def admin_mitarbeiter_neu():
    name    = request.form.get('name',    '').strip()
    kuerzel = request.form.get('kuerzel', '').strip().upper()
    passwort = request.form.get('passwort', 'brauerei').strip()
    email   = request.form.get('email',   '').strip().lower() or None
    if name and kuerzel:
        execute(
            "INSERT OR IGNORE INTO mitarbeiter (name, kuerzel, passwort, email) VALUES (?,?,?,?)",
            (name, kuerzel, passwort, email)
        )
        flash(f'Mitarbeiter "{name}" angelegt.', 'success')
    return redirect(url_for('admin'))


@app.route('/admin/mitarbeiter/<int:ma_id>/email', methods=['POST'])
@admin_required
def admin_mitarbeiter_email(ma_id):
    ma    = query("SELECT * FROM mitarbeiter WHERE id=?", (ma_id,), one=True)
    email = request.form.get('email', '').strip().lower() or None
    if not ma:
        flash('Mitarbeiter nicht gefunden.', 'danger')
        return redirect(url_for('admin'))
    # Doppelte E-Mail prüfen
    if email:
        existing = query(
            "SELECT id FROM mitarbeiter WHERE LOWER(email)=? AND id!=?",
            (email, ma_id), one=True
        )
        if existing:
            flash(f'Die E-Mail „{email}" ist bereits einem anderen Mitarbeiter zugeordnet.', 'danger')
            return redirect(url_for('admin'))
    execute("UPDATE mitarbeiter SET email=? WHERE id=?", (email, ma_id))
    if email:
        flash(f'E-Mail für „{ma["name"]}" auf {email} gesetzt.', 'success')
    else:
        flash(f'E-Mail für „{ma["name"]}" entfernt.', 'info')
    return redirect(url_for('admin'))


@app.route('/admin/mitarbeiter/<int:ma_id>/zuordnung', methods=['POST'])
@admin_required
def admin_mitarbeiter_zuordnung(ma_id):
    ma = query("SELECT * FROM mitarbeiter WHERE id=?", (ma_id,), one=True)
    if not ma:
        flash('Mitarbeiter nicht gefunden.', 'danger')
        return redirect(url_for('admin'))
    vs_ids = [int(x) for x in request.form.getlist('vs_ids')]
    # Exklusive Zuordnung: jede VS gehört nur einer Person
    # → zuerst diese VS bei allen anderen entfernen
    for vs_id in vs_ids:
        execute(
            "DELETE FROM mitarbeiter_verkaufsstelle WHERE verkaufsstelle_id=? AND mitarbeiter_id!=?",
            (vs_id, ma_id)
        )
    execute("DELETE FROM mitarbeiter_verkaufsstelle WHERE mitarbeiter_id=?", (ma_id,))
    for vs_id in vs_ids:
        execute(
            "INSERT OR IGNORE INTO mitarbeiter_verkaufsstelle (mitarbeiter_id, verkaufsstelle_id) VALUES (?,?)",
            (ma_id, vs_id)
        )
    anzahl = len(vs_ids)
    flash(f'Zuordnung für „{ma["name"]}" gespeichert: {anzahl} Verkaufsstelle(n).', 'success')
    return redirect(url_for('admin'))


@app.route('/admin/mitarbeiter/<int:ma_id>/loeschen', methods=['POST'])
@admin_required
def admin_mitarbeiter_loeschen(ma_id):
    if ma_id == session.get('user_id'):
        flash('Sie können sich nicht selbst löschen.', 'danger')
        return redirect(url_for('admin'))
    ma = query("SELECT * FROM mitarbeiter WHERE id=?", (ma_id,), one=True)
    if not ma:
        flash('Mitarbeiter nicht gefunden.', 'danger')
        return redirect(url_for('admin'))
    if ma['rolle'] == 'admin':
        flash('Admin-Konten können nicht gelöscht werden.', 'danger')
        return redirect(url_for('admin'))
    count = query("SELECT COUNT(*) AS c FROM aktivitaet WHERE mitarbeiter_id=?", (ma_id,), one=True)['c']
    if count > 0:
        flash(f'„{ma["name"]}" hat noch {count} Aktivität(en) und kann nicht gelöscht werden. '
              f'Bitte zuerst alle Aktivitäten dieses Mitarbeiters löschen.', 'danger')
        return redirect(url_for('admin'))
    execute("DELETE FROM mitarbeiter WHERE id=?", (ma_id,))
    flash(f'Mitarbeiter „{ma["name"]}" wurde gelöscht.', 'success')
    return redirect(url_for('admin'))


@app.route('/admin/mitarbeiter/<int:ma_id>/passwort', methods=['POST'])
@admin_required
def admin_mitarbeiter_passwort(ma_id):
    ma = query("SELECT * FROM mitarbeiter WHERE id=?", (ma_id,), one=True)
    if not ma:
        flash('Mitarbeiter nicht gefunden.', 'danger')
        return redirect(url_for('admin'))
    neues_pw = request.form.get('passwort', '').strip()
    if len(neues_pw) < 4:
        flash('Passwort muss mindestens 4 Zeichen haben.', 'danger')
        return redirect(url_for('admin'))
    execute("UPDATE mitarbeiter SET passwort=? WHERE id=?", (neues_pw, ma_id))
    flash(f'Passwort für „{ma["name"]}" wurde geändert.', 'success')
    return redirect(url_for('admin'))


@app.route('/profil/passwort', methods=['POST'])
@login_required
def profil_passwort():
    altes_pw  = request.form.get('altes_passwort', '')
    neues_pw  = request.form.get('neues_passwort', '').strip()
    neues_pw2 = request.form.get('neues_passwort2', '').strip()
    user = query("SELECT * FROM mitarbeiter WHERE id=?", (session['user_id'],), one=True)
    if user['passwort'] != altes_pw:
        flash('Aktuelles Passwort ist falsch.', 'danger')
    elif len(neues_pw) < 4:
        flash('Neues Passwort muss mindestens 4 Zeichen haben.', 'danger')
    elif neues_pw != neues_pw2:
        flash('Passwörter stimmen nicht überein.', 'danger')
    else:
        execute("UPDATE mitarbeiter SET passwort=? WHERE id=?", (neues_pw, session['user_id']))
        flash('Passwort erfolgreich geändert.', 'success')
    return redirect(request.referrer or url_for('dashboard'))


@app.route('/admin/vertretung/neu', methods=['POST'])
@manager_required
def admin_vertretung_neu():
    abwesender_id = request.form.get('abwesender_id', type=int)
    vertreter_id  = request.form.get('vertreter_id',  type=int)
    von           = request.form.get('von', '').strip()
    bis           = request.form.get('bis', '').strip()
    if not all([abwesender_id, vertreter_id, von, bis]):
        flash('Alle Felder sind Pflichtfelder.', 'danger')
        return redirect(url_for('admin'))
    if abwesender_id == vertreter_id:
        flash('Abwesender und Vertreter dürfen nicht dieselbe Person sein.', 'danger')
        return redirect(url_for('admin'))
    execute(
        "INSERT INTO vertretung (abwesender_id, vertreter_id, von, bis) VALUES (?,?,?,?)",
        (abwesender_id, vertreter_id, von, bis)
    )
    flash('Vertretungsregelung gespeichert.', 'success')
    return redirect(url_for('admin'))


@app.route('/admin/vertretung/<int:vtr_id>/loeschen', methods=['POST'])
@manager_required
def admin_vertretung_loeschen(vtr_id):
    execute("DELETE FROM vertretung WHERE id=?", (vtr_id,))
    flash('Vertretungsregelung gelöscht.', 'success')
    return redirect(url_for('admin'))


@app.route('/profil/vertretung/neu', methods=['POST'])
@login_required
def profil_vertretung_neu():
    vertreter_id = request.form.get('vertreter_id', type=int)
    von          = request.form.get('von', '').strip()
    bis          = request.form.get('bis', '').strip()
    if not all([vertreter_id, von, bis]):
        flash('Alle Felder sind Pflichtfelder.', 'danger')
        return redirect(request.referrer or url_for('dashboard'))
    if vertreter_id == session['user_id']:
        flash('Sie können sich nicht selbst als Vertreter eintragen.', 'danger')
        return redirect(request.referrer or url_for('dashboard'))
    execute(
        "INSERT INTO vertretung (abwesender_id, vertreter_id, von, bis) VALUES (?,?,?,?)",
        (session['user_id'], vertreter_id, von, bis)
    )
    flash('Vertretung eingetragen.', 'success')
    return redirect(request.referrer or url_for('dashboard'))


@app.route('/profil/vertretung/<int:vtr_id>/loeschen', methods=['POST'])
@login_required
def profil_vertretung_loeschen(vtr_id):
    vtr = query("SELECT * FROM vertretung WHERE id=?", (vtr_id,), one=True)
    if not vtr or vtr['abwesender_id'] != session['user_id']:
        flash('Nicht gefunden oder keine Berechtigung.', 'danger')
        return redirect(request.referrer or url_for('dashboard'))
    execute("DELETE FROM vertretung WHERE id=?", (vtr_id,))
    flash('Vertretung gelöscht.', 'success')
    return redirect(request.referrer or url_for('dashboard'))


@app.route('/admin/verkaufsstelle/neu', methods=['POST'])
@admin_required
def admin_vs_neu():
    name             = request.form.get('name',             '').strip()
    strasse          = request.form.get('strasse',          '').strip()
    ort              = request.form.get('ort',              '').strip()
    typ              = request.form.get('typ',              '').strip()
    ansprechpartner  = request.form.get('ansprechpartner',  '').strip()
    if name:
        execute("INSERT INTO verkaufsstelle (name, strasse, ort, typ, ansprechpartner) VALUES (?,?,?,?,?)",
                (name, strasse, ort, typ, ansprechpartner))
        flash(f'Verkaufsstelle "{name}" angelegt.', 'success')
    return redirect(url_for('admin'))


@app.route('/admin/verkaufsstelle/<int:vs_id>/loeschen', methods=['POST'])
@admin_required
def admin_vs_loeschen(vs_id):
    vs = query("SELECT * FROM verkaufsstelle WHERE id=?", (vs_id,), one=True)
    if not vs:
        flash('Verkaufsstelle nicht gefunden.', 'danger')
        return redirect(url_for('admin'))
    count = query("SELECT COUNT(*) AS c FROM aktivitaet WHERE verkaufsstelle_id=?", (vs_id,), one=True)['c']
    if count > 0:
        execute("UPDATE verkaufsstelle SET aktiv=0 WHERE id=?", (vs_id,))
        flash(f'„{vs["name"]}" hat {count} verknüpfte Aktivität(en) und wurde deaktiviert '
              f'(erscheint nicht mehr in neuen Aktivitäten, historische Daten bleiben erhalten).', 'warning')
    else:
        execute("DELETE FROM verkaufsstelle WHERE id=?", (vs_id,))
        flash(f'Verkaufsstelle „{vs["name"]}" wurde gelöscht.', 'success')
    return redirect(url_for('admin'))


@app.route('/admin/verkaufsstelle/<int:vs_id>/reaktivieren', methods=['POST'])
@admin_required
def admin_vs_reaktivieren(vs_id):
    vs = query("SELECT * FROM verkaufsstelle WHERE id=?", (vs_id,), one=True)
    if vs:
        execute("UPDATE verkaufsstelle SET aktiv=1 WHERE id=?", (vs_id,))
        flash(f'Verkaufsstelle „{vs["name"]}" wurde reaktiviert.', 'success')
    return redirect(url_for('admin'))


@app.route('/verkaufsstelle/neu', methods=['POST'])
@login_required
def vs_neu_rep():
    name            = request.form.get('name',            '').strip()
    strasse         = request.form.get('strasse',         '').strip()
    ort             = request.form.get('ort',             '').strip()
    typ             = request.form.get('typ',             '').strip()
    ansprechpartner = request.form.get('ansprechpartner', '').strip()
    if name:
        new_id = execute(
            "INSERT INTO verkaufsstelle (name, strasse, ort, typ, ansprechpartner) VALUES (?,?,?,?,?)",
            (name, strasse, ort, typ, ansprechpartner)
        )
        # Reps/VKL: neue Verkaufsstelle direkt dem Ersteller zuordnen
        if session.get('rolle') in ('rep', 'verkaufsleiter'):
            execute(
                "INSERT OR IGNORE INTO mitarbeiter_verkaufsstelle (mitarbeiter_id, verkaufsstelle_id) VALUES (?,?)",
                (session['user_id'], new_id)
            )
        flash(f'Verkaufsstelle "{name}" wurde angelegt und ausgewählt.', 'success')
        return redirect(url_for('neue_aktivitaet', vs_id=new_id))
    flash('Bitte einen Namen eingeben.', 'danger')
    return redirect(url_for('neue_aktivitaet'))


@app.route('/admin/biersorte/<int:b_id>/loeschen', methods=['POST'])
@admin_required
def admin_bier_loeschen(b_id):
    b = query("SELECT * FROM biersorte WHERE id=?", (b_id,), one=True)
    if not b:
        flash('Biersorte nicht gefunden.', 'danger')
        return redirect(url_for('admin'))
    count = query(
        "SELECT COUNT(*) AS c FROM bestellposition WHERE biersorte_id=?", (b_id,), one=True
    )['c']
    if count > 0:
        execute("UPDATE biersorte SET aktiv=0 WHERE id=?", (b_id,))
        flash(
            f'„{b["name"]}" hat {count} verknüpfte Bestellung(en) und wurde deaktiviert '
            f'(erscheint nicht mehr in neuen Aktivitäten, historische Daten bleiben erhalten).',
            'warning'
        )
    else:
        execute("DELETE FROM biersorte WHERE id=?", (b_id,))
        flash(f'Biersorte „{b["name"]}" wurde gelöscht.', 'success')
    return redirect(url_for('admin'))


@app.route('/admin/biersorte/<int:b_id>/reaktivieren', methods=['POST'])
@admin_required
def admin_bier_reaktivieren(b_id):
    b = query("SELECT * FROM biersorte WHERE id=?", (b_id,), one=True)
    if b:
        execute("UPDATE biersorte SET aktiv=1 WHERE id=?", (b_id,))
        flash(f'Biersorte „{b["name"]}" wurde reaktiviert.', 'success')
    return redirect(url_for('admin'))


@app.route('/admin/biersorte/neu', methods=['POST'])
@admin_required
def admin_bier_neu():
    name   = request.form.get('name', '').strip()
    einheit = request.form.get('einheit', 'Kiste (20x0.5L)').strip()
    if name:
        execute("INSERT OR IGNORE INTO biersorte (name, einheit) VALUES (?,?)", (name, einheit))
        flash(f'Biersorte "{name}" angelegt.', 'success')
    return redirect(url_for('admin'))


@app.route('/admin/displaysorte/neu', methods=['POST'])
@admin_required
def admin_display_neu():
    name = request.form.get('name', '').strip()
    if name:
        execute("INSERT OR IGNORE INTO displaysorte (name) VALUES (?)", (name,))
        flash(f'Displaysorte "{name}" angelegt.', 'success')
    return redirect(url_for('admin'))


@app.route('/admin/displaysorte/<int:ds_id>/loeschen', methods=['POST'])
@admin_required
def admin_display_loeschen(ds_id):
    d = query("SELECT * FROM displaysorte WHERE id=?", (ds_id,), one=True)
    if d:
        execute("UPDATE displaysorte SET aktiv=0 WHERE id=?", (ds_id,))
        flash(f'Display-Typ „{d["name"]}" wurde deaktiviert.', 'warning')
    return redirect(url_for('admin'))


@app.route('/admin/displaysorte/<int:ds_id>/reaktivieren', methods=['POST'])
@admin_required
def admin_display_reaktivieren(ds_id):
    d = query("SELECT * FROM displaysorte WHERE id=?", (ds_id,), one=True)
    if d:
        execute("UPDATE displaysorte SET aktiv=1 WHERE id=?", (ds_id,))
        flash(f'Display-Typ „{d["name"]}" wurde reaktiviert.', 'success')
    return redirect(url_for('admin'))


# ─── Routes: Excel-Import ────────────────────────────────────────────────────

@app.route('/admin/import-excel/vorlage')
@admin_required
def admin_import_vorlage():
    """Liefert die stammdaten_vorlage.xlsx zum Download."""
    vorlage = os.path.join(os.path.dirname(__file__), 'stammdaten_vorlage.xlsx')
    if not os.path.exists(vorlage):
        flash('Vorlagen-Datei nicht gefunden. Bitte erstelle_vorlage.py ausführen.', 'danger')
        return redirect(url_for('admin'))
    return send_file(vorlage, as_attachment=True,
                     download_name='stammdaten_vorlage.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/admin/import-excel', methods=['POST'])
@admin_required
def admin_import_excel():
    """Importiert Mitarbeiter und Verkaufsstellen aus einer hochgeladenen Excel-Datei."""
    if 'datei' not in request.files or request.files['datei'].filename == '':
        flash('Keine Datei ausgewählt.', 'danger')
        return redirect(url_for('admin'))

    datei = request.files['datei']
    if not datei.filename.lower().endswith(('.xlsx', '.xls')):
        flash('Nur Excel-Dateien (.xlsx) werden unterstützt.', 'danger')
        return redirect(url_for('admin'))

    try:
        import openpyxl
        wb = openpyxl.load_workbook(datei, data_only=True)
    except Exception as e:
        flash(f'Fehler beim Lesen der Excel-Datei: {e}', 'danger')
        return redirect(url_for('admin'))

    def _col(row, hmap, *namen):
        for name in namen:
            idx = hmap.get(name.lower())
            if idx is not None and idx < len(row):
                val = row[idx].value
                return str(val).strip() if val is not None else ''
        return ''

    def _hmap(sheet):
        headers = {}
        first_row = next(sheet.iter_rows(min_row=1, max_row=1))
        for i, cell in enumerate(first_row):
            if cell.value:
                headers[str(cell.value).strip().lower()] = i
        return headers

    stats = {'ma_neu': 0, 'ma_skip': 0, 'vs_neu': 0, 'vs_skip': 0, 'zuweis': 0, 'fehler': []}

    # ── Blatt Mitarbeiter ──────────────────────────────────────────────────────
    MA_NAMES = ['mitarbeiter', 'ma', 'reps']
    ma_sheet = next((wb[n] for n in wb.sheetnames if n.lower() in MA_NAMES), None)

    if ma_sheet:
        hmap = _hmap(ma_sheet)
        for row in ma_sheet.iter_rows(min_row=2):
            name     = _col(row, hmap, 'name', 'Name')
            kuerzel  = _col(row, hmap, 'kuerzel', 'Kürzel', 'kuerzel').upper()
            rolle    = _col(row, hmap, 'rolle', 'Rolle') or 'rep'
            passwort = _col(row, hmap, 'passwort', 'Passwort') or 'brauerei'
            email    = _col(row, hmap, 'email', 'Email', 'E-Mail').lower() or None

            if not name or not kuerzel:
                continue

            exists = query("SELECT id FROM mitarbeiter WHERE UPPER(kuerzel)=?", (kuerzel,), one=True)
            if exists:
                stats['ma_skip'] += 1
            else:
                execute("INSERT INTO mitarbeiter (name, kuerzel, rolle, passwort, email) VALUES (?,?,?,?,?)",
                        (name, kuerzel, rolle, passwort, email))
                stats['ma_neu'] += 1

    # ── Blatt Verkaufsstellen ──────────────────────────────────────────────────
    VS_NAMES = ['verkaufsstellen', 'kunden', 'vs']
    vs_sheet = next((wb[n] for n in wb.sheetnames if n.lower() in VS_NAMES), None)

    if vs_sheet:
        hmap = _hmap(vs_sheet)
        for row in vs_sheet.iter_rows(min_row=2):
            vs_name  = _col(row, hmap, 'name', 'Name')
            ort      = _col(row, hmap, 'ort', 'Ort', 'Stadt') or None
            strasse  = _col(row, hmap, 'strasse', 'Straße', 'Strasse', 'Adresse') or None
            typ      = _col(row, hmap, 'typ', 'Typ', 'Kategorie') or None
            ansprech = _col(row, hmap, 'ansprechpartner', 'Ansprechpartner', 'Kontakt') or None
            ma_raw   = _col(row, hmap, 'mitarbeiter', 'Mitarbeiter', 'rep', 'Rep')

            if not vs_name:
                continue

            existing = query("SELECT id FROM verkaufsstelle WHERE LOWER(name)=LOWER(?)", (vs_name,), one=True)
            if existing:
                vs_id = existing['id']
                stats['vs_skip'] += 1
            else:
                vs_id = execute(
                    "INSERT INTO verkaufsstelle (name, ort, strasse, typ, ansprechpartner, aktiv) VALUES (?,?,?,?,?,1)",
                    (vs_name, ort, strasse, typ, ansprech))
                stats['vs_neu'] += 1

            if ma_raw:
                for kuerzel in [k.strip().upper() for k in ma_raw.split(',') if k.strip()]:
                    ma = query("SELECT id FROM mitarbeiter WHERE UPPER(kuerzel)=?", (kuerzel,), one=True)
                    if not ma:
                        stats['fehler'].append(f'Kürzel „{kuerzel}" nicht gefunden (Zuweisung für „{vs_name}" übersprungen)')
                        continue
                    already = query("SELECT 1 FROM mitarbeiter_verkaufsstelle WHERE mitarbeiter_id=? AND verkaufsstelle_id=?",
                                    (ma['id'], vs_id), one=True)
                    if not already:
                        execute("INSERT INTO mitarbeiter_verkaufsstelle (mitarbeiter_id, verkaufsstelle_id) VALUES (?,?)",
                                (ma['id'], vs_id))
                        stats['zuweis'] += 1

    # ── Ergebnismeldung ────────────────────────────────────────────────────────
    teile = []
    if stats['ma_neu']:    teile.append(f"{stats['ma_neu']} Mitarbeiter neu")
    if stats['ma_skip']:   teile.append(f"{stats['ma_skip']} Mitarbeiter bereits vorhanden")
    if stats['vs_neu']:    teile.append(f"{stats['vs_neu']} Verkaufsstellen neu")
    if stats['vs_skip']:   teile.append(f"{stats['vs_skip']} Verkaufsstellen bereits vorhanden")
    if stats['zuweis']:    teile.append(f"{stats['zuweis']} Zuweisungen neu")

    if not teile:
        flash('Keine neuen Daten gefunden – alle Einträge bereits vorhanden oder Datei leer.', 'warning')
    else:
        flash('Import erfolgreich: ' + ' · '.join(teile) + '.', 'success')

    for fehler in stats['fehler']:
        flash(fehler, 'warning')

    return redirect(url_for('admin'))


# ─── Routes: Vergleich & Zielzahlen ──────────────────────────────────────────

@app.route('/vergleich')
@login_required
def vergleich():
    is_admin   = session.get('rolle') == 'admin'
    is_manager = session.get('rolle') in ('admin', 'verkaufsleiter')
    jahr       = request.args.get('jahr', date.today().year, type=int)

    # Reps sehen nur eigene Ziele; VKL/Admin können per Tab filtern
    if not is_manager:
        ma_filter = str(session['user_id'])
    else:
        ma_filter = request.args.get('ma', '', type=str)

    alle_jahre = [r[0] for r in query(
        "SELECT DISTINCT CAST(strftime('%Y', datum) AS INTEGER) FROM aktivitaet ORDER BY 1 DESC"
    )]
    if not alle_jahre:
        alle_jahre = [date.today().year]
    if date.today().year not in alle_jahre:
        alle_jahre.insert(0, date.today().year)

    # IST-Werte aller Reps – Subquery verhindert Doppelung von anzahl_displays
    _BP = "(SELECT aktivitaet_id, SUM(kisten_anzahl) AS kisten_total FROM bestellposition GROUP BY aktivitaet_id)"
    ist = query(f'''
        SELECT m.id, m.name, m.kuerzel,
               COALESCE(SUM(a.anzahl_displays), 0) AS displays_ist,
               COALESCE(SUM(b.kisten_total), 0)    AS kisten_ist,
               COUNT(a.id)                          AS besuche
        FROM mitarbeiter m
        LEFT JOIN aktivitaet a
               ON a.mitarbeiter_id = m.id
              AND strftime('%Y', a.datum) = ?
        LEFT JOIN {_BP} b ON b.aktivitaet_id = a.id
        WHERE m.rolle IN ('rep','verkaufsleiter')
        GROUP BY m.id
        ORDER BY m.name
    ''', (str(jahr),))

    # Zielzahlen
    ziele_raw = query(
        "SELECT mitarbeiter_id, displays_ziel, kisten_ziel FROM zielzahlen WHERE jahr = ?",
        (str(jahr),)
    )
    ziele = {r['mitarbeiter_id']: dict(r) for r in ziele_raw}

    # Teamziel (mitarbeiter_id IS NULL)
    teamziel_row = query(
        "SELECT displays_ziel, kisten_ziel FROM zielzahlen WHERE mitarbeiter_id IS NULL AND jahr = ?",
        (str(jahr),), one=True
    )
    teamziel = dict(teamziel_row) if teamziel_row else None

    # Daten für Charts aufbereiten (ohne saison_soll – zielkurs noch nicht befüllt)
    reps_namen  = [r['name'] for r in ist]
    kisten_ist  = [r['kisten_ist'] for r in ist]
    kisten_soll = [ziele.get(r['id'], {}).get('kisten_ziel', 0) or 0 for r in ist]
    disp_ist    = [r['displays_ist'] for r in ist]
    disp_soll   = [ziele.get(r['id'], {}).get('displays_ziel', 0) or 0 for r in ist]

    # KW-Verlauf je Mitarbeiter (für Liniendiagramm)
    kw_verlauf = query('''
        SELECT m.name,
               CAST(strftime('%W', a.datum) AS INTEGER) AS kw,
               COALESCE(SUM(bp.kisten_anzahl), 0) AS kisten
        FROM mitarbeiter m
        JOIN aktivitaet a ON a.mitarbeiter_id = m.id
        LEFT JOIN bestellposition bp ON bp.aktivitaet_id = a.id
        WHERE strftime('%Y', a.datum) = ? AND m.rolle IN ('rep','verkaufsleiter')
        GROUP BY m.id, kw
        ORDER BY m.name, kw
    ''', (str(jahr),))

    # KW-Verlauf als Dict {rep_name: [(kw, kisten), ...]}
    from collections import defaultdict
    kw_by_rep = defaultdict(list)
    for row in kw_verlauf:
        kw_by_rep[row['name']].append({'kw': row['kw'], 'kisten': row['kisten']})

    # ── Saisonaler Zielkurs (Sonnenschlüssel) ────────────────────────────────
    aktueller_monat = date.today().month  # 1–12

    # Monatliche IST-Werte je Rep (mit korrekter BP-Subquery)
    _BPm = "(SELECT aktivitaet_id, SUM(kisten_anzahl) AS kisten_total FROM bestellposition GROUP BY aktivitaet_id)"
    monatlich_raw = query(f'''
        SELECT m.id AS rep_id,
               CAST(strftime('%m', a.datum) AS INTEGER) AS monat,
               COALESCE(SUM(a.anzahl_displays), 0) AS displays_ist,
               COALESCE(SUM(b.kisten_total),   0) AS kisten_ist
        FROM mitarbeiter m
        LEFT JOIN aktivitaet a
               ON a.mitarbeiter_id = m.id
              AND strftime('%Y', a.datum) = ?
        LEFT JOIN {_BPm} b ON b.aktivitaet_id = a.id
        WHERE m.rolle IN ('rep','verkaufsleiter')
        GROUP BY m.id, monat
    ''', (str(jahr),))

    # {rep_id: {monat(1–12): {displays, kisten}}}
    monatlich_dict = {}
    for row in monatlich_raw:
        rid = row['rep_id']
        if rid not in monatlich_dict:
            monatlich_dict[rid] = {}
        if row['monat']:
            monatlich_dict[rid][row['monat']] = {
                'displays': row['displays_ist'],
                'kisten':   row['kisten_ist'],
            }

    monatsziele   = {}   # {rep_id: list[12 dicts]}
    zielkurs      = {}   # {rep_id: {disp_kurs, kist_kurs, kum_*}}
    kum_ist_data  = {}   # {kuerzel: [kum_kisten pro vergangenen Monat]}
    kum_ziel_data = {}   # {kuerzel: [kum_kisten_ziel pro vergangenen Monat]}

    for rep_row in ist:
        rid     = rep_row['id']
        kuerzel = rep_row['kuerzel']
        ziel    = ziele.get(rid, {})
        d_jz    = ziel.get('displays_ziel', 0) or 0
        k_jz    = ziel.get('kisten_ziel',   0) or 0

        monate = []
        kum_d_ziel = kum_k_ziel = kum_d_ist = kum_k_ist = 0
        kum_k_series = []
        kum_kz_series = []
        lauf_k = lauf_kz = 0

        for idx in range(12):
            m_num      = idx + 1
            mon_d_ziel = round(d_jz * SONNENSCHLUESSEL[idx])
            mon_k_ziel = round(k_jz * SONNENSCHLUESSEL[idx])
            mon_ist    = monatlich_dict.get(rid, {}).get(m_num, {'displays': 0, 'kisten': 0})
            vergangen  = (m_num <= aktueller_monat)

            d_pct_m = round(mon_ist['displays'] / mon_d_ziel * 100) if (mon_d_ziel and vergangen) else None
            k_pct_m = round(mon_ist['kisten']   / mon_k_ziel * 100) if (mon_k_ziel and vergangen) else None

            if vergangen:
                kum_d_ziel += mon_d_ziel;  kum_k_ziel += mon_k_ziel
                kum_d_ist  += mon_ist['displays']
                kum_k_ist  += mon_ist['kisten']
                lauf_k  += mon_ist['kisten'];  lauf_kz += mon_k_ziel
                kum_k_series.append(lauf_k)
                kum_kz_series.append(lauf_kz)

            monate.append({
                'num':           m_num,
                'name':          M_NAMEN[idx],
                'displays_ziel': mon_d_ziel,
                'kisten_ziel':   mon_k_ziel,
                'displays_ist':  mon_ist['displays'],
                'kisten_ist':    mon_ist['kisten'],
                'displays_pct':  d_pct_m,
                'kisten_pct':    k_pct_m,
                'vergangen':     vergangen,
            })

        monatsziele[rid] = monate
        zielkurs[rid] = {
            'disp_kurs': round(kum_d_ist / kum_d_ziel * 100) if kum_d_ziel else None,
            'kist_kurs': round(kum_k_ist / kum_k_ziel * 100) if kum_k_ziel else None,
            'kum_d_ist':  kum_d_ist,  'kum_d_ziel': kum_d_ziel,
            'kum_k_ist':  kum_k_ist,  'kum_k_ziel': kum_k_ziel,
        }
        kum_ist_data[kuerzel]  = kum_k_series
        kum_ziel_data[kuerzel] = kum_kz_series

    # Saisonaler Anteilszielkurs bis aktueller Monat – jetzt nach zielkurs-Loop
    kisten_saison_soll = [zielkurs.get(r['id'], {}).get('kum_k_ziel', 0) or 0 for r in ist]
    disp_saison_soll   = [zielkurs.get(r['id'], {}).get('kum_d_ziel', 0) or 0 for r in ist]

    sk_pct  = [int(x * 100) for x in SONNENSCHLUESSEL]
    alle_ma = query("SELECT id, name, kuerzel FROM mitarbeiter WHERE rolle IN ('rep','verkaufsleiter') ORDER BY name") if is_manager else []

    # Einzelner Rep für Detailansicht (Tabs)
    selected_rep = None
    if ma_filter:
        for rep_row in ist:
            if str(rep_row['id']) == ma_filter:
                selected_rep = dict(rep_row)
                break

    return render_template('vergleich.html',
        jahr=jahr, alle_jahre=alle_jahre,
        ist=ist, ziele=ziele, teamziel=teamziel,
        reps_namen=json.dumps(reps_namen),
        kisten_ist=json.dumps(kisten_ist),
        kisten_soll=json.dumps(kisten_soll),
        disp_ist=json.dumps(disp_ist),
        disp_soll=json.dumps(disp_soll),
        kw_by_rep=json.dumps(kw_by_rep),
        monatsziele=monatsziele,
        zielkurs=zielkurs,
        kum_ist_data=json.dumps(kum_ist_data),
        kum_ziel_data=json.dumps(kum_ziel_data),
        kisten_saison_soll=json.dumps(kisten_saison_soll),
        disp_saison_soll=json.dumps(disp_saison_soll),
        aktueller_monat=aktueller_monat,
        m_namen=M_NAMEN,
        sk_pct=sk_pct,
        is_manager=is_manager,
        is_admin=is_admin,
        ma_filter=ma_filter,
        alle_ma=alle_ma,
        selected_rep=selected_rep,
    )


@app.route('/zielzahlen', methods=['GET', 'POST'])
@manager_required
def zielzahlen():
    jahr = request.args.get('jahr', date.today().year, type=int)

    if request.method == 'POST':
        jar = request.form.get('jahr', date.today().year, type=int)
        reps = query("SELECT id FROM mitarbeiter WHERE rolle IN ('rep','verkaufsleiter')")

        for rep in reps:
            d_ziel = request.form.get(f'disp_{rep["id"]}', 0) or 0
            k_ziel = request.form.get(f'kist_{rep["id"]}', 0) or 0
            execute('''
                INSERT INTO zielzahlen (mitarbeiter_id, jahr, displays_ziel, kisten_ziel)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(mitarbeiter_id, jahr) DO UPDATE SET
                    displays_ziel = excluded.displays_ziel,
                    kisten_ziel   = excluded.kisten_ziel
            ''', (rep['id'], jar, int(d_ziel), int(k_ziel)))

        # Teamziel
        td = request.form.get('team_disp', 0) or 0
        tk = request.form.get('team_kist', 0) or 0
        execute('''
            INSERT INTO zielzahlen (mitarbeiter_id, jahr, displays_ziel, kisten_ziel)
            VALUES (NULL, ?, ?, ?)
            ON CONFLICT(mitarbeiter_id, jahr) DO UPDATE SET
                displays_ziel = excluded.displays_ziel,
                kisten_ziel   = excluded.kisten_ziel
        ''', (jar, int(td), int(tk)))

        flash(f'Zielzahlen für {jar} gespeichert.', 'success')
        return redirect(url_for('zielzahlen', jahr=jar))

    reps = query("SELECT id, name, kuerzel FROM mitarbeiter WHERE rolle IN ('rep','verkaufsleiter') ORDER BY name")
    ziele_raw = query(
        "SELECT mitarbeiter_id, displays_ziel, kisten_ziel FROM zielzahlen WHERE jahr = ?",
        (str(jahr),)
    )
    ziele = {r['mitarbeiter_id']: dict(r) for r in ziele_raw}
    teamziel = ziele.get(None)

    alle_jahre = list(range(date.today().year, date.today().year - 3, -1))

    return render_template('zielzahlen.html',
        reps=reps, ziele=ziele, teamziel=teamziel,
        jahr=jahr, alle_jahre=alle_jahre)


# ─── PWA ─────────────────────────────────────────────────────────────────────

@app.route('/sw.js')
def service_worker():
    from flask import send_from_directory
    return send_from_directory('static', 'sw.js', mimetype='application/javascript')


# ─── API: Autocomplete ────────────────────────────────────────────────────────

@app.route('/api/verkaufsstellen')
@login_required
def api_verkaufsstellen():
    q = request.args.get('q', '')
    rows = query("SELECT id, name, ort, typ FROM verkaufsstelle WHERE aktiv=1 AND name LIKE ? ORDER BY name LIMIT 20",
                 (f'%{q}%',))
    return jsonify([dict(r) for r in rows])


# ─── Auto-Export ──────────────────────────────────────────────────────────────

def erstelle_fotos_zip_bytes(wochen: int = 4):
    """Erstellt ZIP-Archiv aller Fotos der letzten `wochen` Wochen.
    Gibt (zip_bytes, anzahl) zurück."""
    grenzwert = (date.today() - timedelta(weeks=wochen)).isoformat()
    fotos = query(
        "SELECT a.datum, m.kuerzel, a.foto_pfad "
        "FROM aktivitaet a JOIN mitarbeiter m ON m.id = a.mitarbeiter_id "
        "WHERE a.foto_pfad IS NOT NULL AND a.foto_pfad != '' AND a.datum >= ? "
        "ORDER BY a.datum",
        (grenzwert,)
    )
    buf   = io.BytesIO()
    count = 0
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for f in fotos:
            pfad = os.path.join(UPLOAD_FOLDER, f['foto_pfad'])
            if os.path.exists(pfad):
                ext     = os.path.splitext(f['foto_pfad'])[1]
                arcname = f"{f['datum']}_{f['kuerzel']}_{count+1:03d}{ext}"
                zf.write(pfad, arcname)
                count += 1
    buf.seek(0)
    return buf.read(), count


def auto_export_job():
    """Automatischer 4-Wochen-Export: Excel-Jahresauswertung + Foto-ZIP per E-Mail."""
    if not EXPORT_EMAIL:
        app.logger.info("AUTO_EXPORT: EXPORT_EMAIL nicht gesetzt – übersprungen.")
        return
    with app.app_context():
        try:
            jahr      = date.today().year
            heute_str = date.today().strftime('%d.%m.%Y')

            excel_bytes           = _build_excel_bytes(jahr, is_admin=True)
            zip_bytes, foto_count = erstelle_fotos_zip_bytes(wochen=4)

            body = f"""
            <div style="font-family:sans-serif;max-width:600px;color:#222">
              <h2 style="color:#1a3a5c">Aktions Tracker – Automatischer Export</h2>
              <p>Sehr geehrte Damen und Herren,</p>
              <p>anbei erhalten Sie den automatischen Datenexport vom <strong>{heute_str}</strong>.</p>
              <ul>
                <li><strong>Excel-Auswertung:</strong> Jahresübersicht {jahr}
                    (KW-Übersicht, Mitarbeiter-Ranking, Aktivitäten-Detail, Produktübersicht)</li>
                <li><strong>Foto-Archiv:</strong> {foto_count} Foto(s) der letzten 4 Wochen als ZIP</li>
              </ul>
              <p style="color:#666;font-size:.9em">
                Die Fotos werden im System automatisch nach {FOTO_AUFBEWAHRUNG_WOCHEN} Wochen gelöscht –
                dieses Archiv enthält alle Aufnahmen des abgelaufenen Zeitraums.<br>
                Die Excel-Datei enthält den vollständigen Jahresstand zum Exportzeitpunkt.
              </p>
              <hr style="border:none;border-top:1px solid #eee;margin:1.5rem 0">
              <p style="font-size:.85em;color:#888">
                Aktions Tracker · Jan Anschütz · anschuetz.info@gmail.com<br>
                Automatisch generiert – bitte nicht auf diese E-Mail antworten.
              </p>
            </div>
            """

            attachments = [
                (f"Aktions_Tracker_{jahr}.xlsx", excel_bytes,
                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            ]
            if foto_count > 0:
                zip_name = f"Fotos_{date.today().strftime('%Y-%m-%d')}.zip"
                attachments.append((zip_name, zip_bytes, "application/zip"))

            ok = send_email_with_attachments(
                EXPORT_EMAIL,
                f"Aktions Tracker – Automatischer Export {heute_str}",
                body,
                attachments
            )
            if ok:
                app.logger.info(f"AUTO_EXPORT: Gesendet an {EXPORT_EMAIL} ({foto_count} Foto(s))")
            else:
                app.logger.error("AUTO_EXPORT: E-Mail-Versand fehlgeschlagen")
        except Exception as e:
            app.logger.error(f"AUTO_EXPORT Fehler: {e}", exc_info=True)


# ─── Wochenbericht ───────────────────────────────────────────────────────────

APP_URL = os.environ.get('APP_URL', '')

def send_wochenbericht(force=False):
    """Wöchentlichen Bericht generieren und an konfigurierte Empfänger senden.
    force=True überspringt aktiv- und Duplikat-Prüfung (für Test-Versand)."""
    with app.app_context():
        try:
            config = query("SELECT * FROM wochenbericht_config WHERE id=1", one=True)
            if not config:
                return False, "Keine Konfiguration gefunden."
            if not force and not config['aktiv']:
                return False, "Wochenbericht ist deaktiviert."

            # Nicht zwei Mal in derselben Woche senden (außer bei force)
            kw_key = date.today().strftime('%Y-W%V')
            if not force and config['zuletzt_gesendet'] == kw_key:
                app.logger.info("WOCHENBERICHT: Diese Woche bereits gesendet – übersprungen.")
                return False, "Diese Woche bereits gesendet."

            # Empfänger: VKL-E-Mail + bis zu 2 weitere
            vkl = query(
                "SELECT email, name FROM mitarbeiter WHERE rolle IN ('verkaufsleiter','admin') "
                "AND email IS NOT NULL AND email != '' ORDER BY rolle='verkaufsleiter' DESC LIMIT 1",
                one=True
            )
            empfaenger = []
            if vkl and vkl['email']:
                empfaenger.append(vkl['email'])
            if config['empfaenger_2']:
                empfaenger.append(config['empfaenger_2'])
            if config['empfaenger_3']:
                empfaenger.append(config['empfaenger_3'])
            if not empfaenger:
                app.logger.warning("WOCHENBERICHT: Keine Empfänger konfiguriert – übersprungen.")
                return False, "Keine Empfänger konfiguriert. Bitte E-Mail-Adresse des Verkaufsleiters im Admin-Panel hinterlegen oder einen zusätzlichen Empfänger eintragen."

            # Zeiträume
            heute          = date.today()
            montag_diese   = heute - timedelta(days=heute.weekday())
            sonntag_diese  = montag_diese + timedelta(days=6)
            montag_letzte  = montag_diese - timedelta(days=7)
            sonntag_letzte = montag_letzte + timedelta(days=6)
            kw_nr          = montag_diese.strftime('%V')
            datum_von      = montag_diese.strftime('%d.%m.')
            datum_bis      = sonntag_diese.strftime('%d.%m.%Y')

            def stats(von, bis):
                return query('''
                    SELECT COUNT(DISTINCT a.id)          AS besuche,
                           COALESCE(SUM(bp.kisten_anzahl),0) AS kisten,
                           COALESCE(SUM(a.anzahl_displays),0) AS displays
                    FROM aktivitaet a
                    LEFT JOIN bestellposition bp ON bp.aktivitaet_id = a.id
                    WHERE a.datum BETWEEN ? AND ?
                ''', (von.isoformat(), bis.isoformat()), one=True)

            diese = stats(montag_diese,  sonntag_diese)
            letzte = stats(montag_letzte, sonntag_letzte)

            rep_stats = query('''
                SELECT m.name,
                       COUNT(DISTINCT a.id)              AS besuche,
                       COALESCE(SUM(bp.kisten_anzahl),0) AS kisten
                FROM aktivitaet a
                JOIN mitarbeiter m ON m.id = a.mitarbeiter_id
                LEFT JOIN bestellposition bp ON bp.aktivitaet_id = a.id
                WHERE a.datum BETWEEN ? AND ? AND m.rolle = 'rep'
                GROUP BY m.id, m.name
                ORDER BY kisten DESC
            ''', (montag_diese.isoformat(), sonntag_diese.isoformat()))

            def trend_str(neu, alt):
                diff = neu - alt
                if diff > 0: return f'+{diff}'
                if diff < 0: return str(diff)
                return '±0'

            def trend_col(neu, alt):
                if neu > alt: return '#2d8a4e'
                if neu < alt: return '#c0392b'
                return '#888'

            rep_rows = ''.join(f'''
                <tr>
                  <td style="padding:9px 14px;border-bottom:1px solid #f0f0f0">{r["name"]}</td>
                  <td style="padding:9px 14px;border-bottom:1px solid #f0f0f0;text-align:center">{r["besuche"]}</td>
                  <td style="padding:9px 14px;border-bottom:1px solid #f0f0f0;text-align:center;font-weight:600;color:#c8860a">{r["kisten"]}</td>
                </tr>''' for r in rep_stats) or \
                '<tr><td colspan="3" style="padding:12px 14px;color:#999;text-align:center">Keine Aktivitäten diese Woche</td></tr>'

            dashboard_link = APP_URL or '#'

            html = f'''<!DOCTYPE html>
<html><head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;background:#f0f4f8;font-family:Arial,Helvetica,sans-serif">
<div style="max-width:600px;margin:32px auto;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,0.10)">

  <div style="background:#1a3a5c;padding:26px 32px">
    <div style="color:#fff;font-size:20px;font-weight:bold;letter-spacing:.3px">Aktions Tracker</div>
    <div style="color:#90b8d8;font-size:13px;margin-top:5px">Wochenbericht KW {kw_nr} &nbsp;·&nbsp; {datum_von} – {datum_bis}</div>
  </div>

  <div style="padding:28px 32px 8px">
    <div style="font-size:15px;font-weight:bold;color:#1a3a5c;margin-bottom:16px">Gesamtübersicht</div>
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <td style="text-align:center;padding:18px 10px;background:#f4f8fc;border-radius:8px">
          <div style="font-size:30px;font-weight:bold;color:#1a3a5c">{diese["besuche"]}</div>
          <div style="font-size:12px;color:#666;margin-top:3px">Besuche</div>
          <div style="font-size:11px;font-weight:bold;color:{trend_col(diese["besuche"],letzte["besuche"])};margin-top:5px">{trend_str(diese["besuche"],letzte["besuche"])} ggü. Vorwoche</div>
        </td>
        <td width="12"></td>
        <td style="text-align:center;padding:18px 10px;background:#f4f8fc;border-radius:8px">
          <div style="font-size:30px;font-weight:bold;color:#c8860a">{diese["kisten"]}</div>
          <div style="font-size:12px;color:#666;margin-top:3px">Kisten</div>
          <div style="font-size:11px;font-weight:bold;color:{trend_col(diese["kisten"],letzte["kisten"])};margin-top:5px">{trend_str(diese["kisten"],letzte["kisten"])} ggü. Vorwoche</div>
        </td>
        <td width="12"></td>
        <td style="text-align:center;padding:18px 10px;background:#f4f8fc;border-radius:8px">
          <div style="font-size:30px;font-weight:bold;color:#2e6da4">{diese["displays"]}</div>
          <div style="font-size:12px;color:#666;margin-top:3px">Displays</div>
          <div style="font-size:11px;font-weight:bold;color:{trend_col(diese["displays"],letzte["displays"])};margin-top:5px">{trend_str(diese["displays"],letzte["displays"])} ggü. Vorwoche</div>
        </td>
      </tr>
    </table>
  </div>

  <div style="padding:24px 32px">
    <div style="font-size:15px;font-weight:bold;color:#1a3a5c;margin-bottom:12px">Mitarbeiter diese Woche</div>
    <table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e4eaf0;border-radius:8px;overflow:hidden">
      <thead>
        <tr style="background:#edf2f7">
          <th style="padding:9px 14px;text-align:left;font-size:11px;color:#666;font-weight:600;letter-spacing:.5px">MITARBEITER</th>
          <th style="padding:9px 14px;text-align:center;font-size:11px;color:#666;font-weight:600;letter-spacing:.5px">BESUCHE</th>
          <th style="padding:9px 14px;text-align:center;font-size:11px;color:#666;font-weight:600;letter-spacing:.5px">KISTEN</th>
        </tr>
      </thead>
      <tbody>{rep_rows}</tbody>
    </table>
  </div>

  <div style="padding:16px 32px 24px;text-align:center">
    <a href="{dashboard_link}" style="display:inline-block;background:#1a3a5c;color:#fff;text-decoration:none;padding:10px 24px;border-radius:6px;font-size:13px;font-weight:bold">→ Zum Dashboard</a>
  </div>

  <div style="padding:14px 32px;background:#f4f8fc;border-top:1px solid #e4eaf0;text-align:center">
    <div style="font-size:11px;color:#aaa">Aktions Tracker · Automatischer Wochenbericht jeden Montag<br>
    Einstellungen unter <em>Einstellungen → Wochenbericht</em></div>
  </div>

</div>
</body></html>'''

            betreff = f'Wochenbericht KW {kw_nr} – Aktions Tracker'
            ok_count = 0
            for mail in empfaenger:
                if send_email(mail, betreff, html):
                    app.logger.info(f"WOCHENBERICHT KW {kw_nr}: Gesendet an {mail}")
                    ok_count += 1
                else:
                    app.logger.error(f"WOCHENBERICHT KW {kw_nr}: Versand an {mail} fehlgeschlagen")

            if ok_count > 0:
                execute("UPDATE wochenbericht_config SET zuletzt_gesendet=? WHERE id=1", (kw_key,))
                return True, f"Gesendet an {ok_count} Empfänger: {', '.join(empfaenger)}"
            else:
                return False, "E-Mail-Versand fehlgeschlagen – SMTP-Konfiguration prüfen (Railway-Umgebungsvariablen MAIL_SERVER, MAIL_USERNAME, MAIL_PASSWORD)."

        except Exception as e:
            app.logger.error(f"WOCHENBERICHT Fehler: {e}", exc_info=True)
            return False, f"Fehler: {e}"


@app.route('/einstellungen/wochenbericht', methods=['GET', 'POST'])
@login_required
def einstellungen_wochenbericht():
    if session.get('rolle') not in ('admin', 'verkaufsleiter'):
        return redirect(url_for('dashboard'))

    vkl = query(
        "SELECT email, name FROM mitarbeiter WHERE rolle IN ('verkaufsleiter','admin') "
        "AND email IS NOT NULL AND email != '' ORDER BY rolle='verkaufsleiter' DESC LIMIT 1",
        one=True
    )
    config = query("SELECT * FROM wochenbericht_config WHERE id=1", one=True)

    if request.method == 'POST':
        aktiv        = 1 if request.form.get('aktiv') else 0
        empfaenger_2 = request.form.get('empfaenger_2', '').strip()
        empfaenger_3 = request.form.get('empfaenger_3', '').strip()
        execute(
            "UPDATE wochenbericht_config SET aktiv=?, empfaenger_2=?, empfaenger_3=? WHERE id=1",
            (aktiv, empfaenger_2, empfaenger_3)
        )
        if request.form.get('jetzt_senden'):
            ok, msg = send_wochenbericht(force=True)
            flash(msg, 'success' if ok else 'danger')
        else:
            flash('Einstellungen gespeichert.', 'success')
        return redirect(url_for('einstellungen_wochenbericht'))

    return render_template('einstellungen_wochenbericht.html',
                           config=config, vkl=vkl,
                           is_manager=True, is_admin=session.get('rolle')=='admin')


# ─── Karte ────────────────────────────────────────────────────────────────────

@app.route('/karte')
@login_required
def karte():
    if KARTE_MODUS == 'aus':
        flash('Die Karten-Funktion ist in Ihrem aktuellen Paket nicht verfügbar.', 'warning')
        return redirect(url_for('dashboard'))
    is_manager = session.get('rolle') in ('admin', 'verkaufsleiter')
    reps = query("SELECT id, name, kuerzel FROM mitarbeiter WHERE rolle IN ('rep','verkaufsleiter') ORDER BY name")
    return render_template('karte.html', reps=reps, is_manager=is_manager, karte_modus=KARTE_MODUS)


@app.route('/api/karte/daten')
@login_required
def api_karte_daten():
    if KARTE_MODUS == 'aus':
        return jsonify({'error': 'Nicht verfügbar'}), 403
    is_manager = session.get('rolle') in ('admin', 'verkaufsleiter')

    if is_manager:
        stellen = query("""
            SELECT v.id, v.name, v.ort, v.typ, v.strasse, v.ansprechpartner, v.lat, v.lng,
                   GROUP_CONCAT(m.id || ':' || m.name || ':' || m.kuerzel, '|') AS zuordnungen
            FROM verkaufsstelle v
            LEFT JOIN mitarbeiter_verkaufsstelle mv ON mv.verkaufsstelle_id = v.id
            LEFT JOIN mitarbeiter m ON m.id = mv.mitarbeiter_id AND m.rolle IN ('rep', 'verkaufsleiter')
            WHERE v.aktiv = 1
            GROUP BY v.id
            ORDER BY v.name
        """)
    else:
        stellen = query("""
            SELECT v.id, v.name, v.ort, v.typ, v.strasse, v.ansprechpartner, v.lat, v.lng,
                   m.id || ':' || m.name || ':' || m.kuerzel AS zuordnungen
            FROM verkaufsstelle v
            JOIN mitarbeiter_verkaufsstelle mv ON mv.verkaufsstelle_id = v.id
            JOIN mitarbeiter m ON m.id = mv.mitarbeiter_id
            WHERE v.aktiv = 1 AND mv.mitarbeiter_id = ?
            ORDER BY v.name
        """, (session['user_id'],))

    result = []
    for s in stellen:
        zuordnung_list = []
        if s['zuordnungen']:
            for z in s['zuordnungen'].split('|'):
                parts = z.split(':', 2)
                if len(parts) == 3:
                    zuordnung_list.append({
                        'id': int(parts[0]), 'name': parts[1], 'kuerzel': parts[2]
                    })
        result.append({
            'id': s['id'],
            'name': s['name'],
            'ort':  s['ort'] or '',
            'typ':  s['typ'] or '',
            'strasse': s['strasse'] or '',
            'ansprechpartner': s['ansprechpartner'] or '',
            'lat': s['lat'],
            'lng': s['lng'],
            'zuordnungen': zuordnung_list,
        })

    reps = query(
        "SELECT id, name, kuerzel FROM mitarbeiter WHERE rolle IN ('rep','verkaufsleiter') ORDER BY name"
    ) if is_manager else []

    return jsonify({
        'stellen': result,
        'reps': [{'id': r['id'], 'name': r['name'], 'kuerzel': r['kuerzel']} for r in reps],
    })


@app.route('/api/karte/zuordnung-aendern', methods=['POST'])
@manager_required
def api_karte_zuordnung_aendern():
    if KARTE_MODUS == 'aus':
        return jsonify({'error': 'Nicht verfügbar'}), 403
    data      = request.get_json() or {}
    stelle_id = data.get('stelle_id')
    neue_ids  = set(int(i) for i in data.get('rep_ids', []))

    if not stelle_id:
        return jsonify({'error': 'stelle_id fehlt'}), 400

    stelle = query("SELECT name FROM verkaufsstelle WHERE id=?", (stelle_id,), one=True)
    if not stelle:
        return jsonify({'error': 'Station nicht gefunden'}), 404
    stelle_name = stelle['name']

    alte_rows = query(
        "SELECT mitarbeiter_id FROM mitarbeiter_verkaufsstelle WHERE verkaufsstelle_id=?",
        (stelle_id,)
    )
    alte_ids = {r['mitarbeiter_id'] for r in alte_rows}

    entfernt     = alte_ids - neue_ids
    hinzugefuegt = neue_ids - alte_ids

    db = get_db()
    db.execute("DELETE FROM mitarbeiter_verkaufsstelle WHERE verkaufsstelle_id=?", (stelle_id,))
    for rep_id in neue_ids:
        db.execute(
            "INSERT OR IGNORE INTO mitarbeiter_verkaufsstelle (mitarbeiter_id, verkaufsstelle_id) VALUES (?,?)",
            (rep_id, stelle_id)
        )

    heute = date.today().strftime('%d.%m.%Y')
    for rep_id in entfernt | hinzugefuegt:
        ma_rolle = db.execute("SELECT rolle FROM mitarbeiter WHERE id=?", (rep_id,)).fetchone()
        if not ma_rolle or ma_rolle['rolle'] != 'rep':
            continue
        if rep_id in entfernt:
            msg = f'{heute}: Station "{stelle_name}" wurde aus Ihrem Gebiet entfernt.'
        else:
            msg = f'{heute}: Station "{stelle_name}" wurde Ihrem Gebiet hinzugefügt.'
        bestehend = db.execute(
            "SELECT karte_benachrichtigung FROM mitarbeiter WHERE id=?", (rep_id,)
        ).fetchone()
        alt = bestehend['karte_benachrichtigung'] if bestehend and bestehend['karte_benachrichtigung'] else ''
        neu = (alt + '\n' + msg).strip()
        db.execute("UPDATE mitarbeiter SET karte_benachrichtigung=? WHERE id=?", (neu, rep_id))

    db.commit()
    return jsonify({'ok': True, 'stelle_name': stelle_name})


@app.route('/api/karte/geocode', methods=['POST'])
@manager_required
def api_karte_geocode():
    if KARTE_MODUS == 'aus':
        return jsonify({'error': 'Nicht verfügbar'}), 403

    stellen = query(
        "SELECT id, name, strasse, ort FROM verkaufsstelle WHERE aktiv=1 AND (lat IS NULL OR lng IS NULL)"
    )
    if not stellen:
        return jsonify({'geocoded': 0, 'total': 0, 'msg': 'Alle Stationen haben bereits Koordinaten.'})

    stellen_list = [dict(s) for s in stellen]

    def geocode_worker(items):
        import sqlite3 as _sqlite3
        conn = _sqlite3.connect(DATABASE)
        try:
            for stelle in items:
                teile = [stelle['name']]
                if stelle.get('strasse'):
                    teile.append(stelle['strasse'])
                if stelle.get('ort'):
                    teile.append(stelle['ort'])
                teile.append('Deutschland')
                url = ('https://nominatim.openstreetmap.org/search?q='
                       + urllib.parse.quote(', '.join(teile))
                       + '&format=json&limit=1&countrycodes=de')
                try:
                    req = urllib.request.Request(
                        url,
                        headers={'User-Agent': 'AktionsTracker/1.0 (anschuetz.info@gmail.com)'}
                    )
                    with urllib.request.urlopen(req, timeout=8) as resp:
                        hits = json.loads(resp.read().decode())
                    if hits:
                        conn.execute(
                            "UPDATE verkaufsstelle SET lat=?, lng=? WHERE id=?",
                            (float(hits[0]['lat']), float(hits[0]['lon']), stelle['id'])
                        )
                        conn.commit()
                except Exception as exc:
                    app.logger.warning(f"Geocode {stelle['name']}: {exc}")
                _time.sleep(1.1)
        finally:
            conn.close()
        app.logger.info(f"Geocodierung abgeschlossen: {len(items)} Stationen verarbeitet.")

    t = threading.Thread(target=geocode_worker, args=(stellen_list,), daemon=True)
    t.start()

    warte = max(30, round(len(stellen_list) * 1.2))
    return jsonify({
        'total': len(stellen_list),
        'msg':   (f'Geocodierung für {len(stellen_list)} Stationen gestartet – '
                  f'läuft im Hintergrund (~{warte} Sek.). '
                  f'Karte wird danach automatisch aktualisiert.'),
        'warte': warte,
    })


@app.route('/api/karte/heatmap')
@login_required
def api_karte_heatmap():
    if KARTE_MODUS != 'heatmap':
        return jsonify({'error': 'Heatmap nicht verfügbar'}), 403
    jahr = request.args.get('jahr', date.today().year, type=int)

    stellen = query("""
        SELECT v.id, v.name, v.ort, v.lat, v.lng,
               COUNT(a.id) AS anzahl
        FROM verkaufsstelle v
        LEFT JOIN aktivitaet a
          ON a.verkaufsstelle_id = v.id
         AND strftime('%Y', a.datum) = ?
        WHERE v.aktiv = 1
          AND v.lat IS NOT NULL AND v.lng IS NOT NULL
        GROUP BY v.id
        ORDER BY anzahl DESC
    """, (str(jahr),))

    jahre_raw = query("""
        SELECT DISTINCT strftime('%Y', datum) AS jahr
        FROM aktivitaet
        ORDER BY jahr DESC
    """)

    return jsonify({
        'stellen': [{'id': s['id'], 'name': s['name'], 'ort': s['ort'] or '',
                     'lat': s['lat'], 'lng': s['lng'], 'anzahl': s['anzahl']} for s in stellen],
        'jahre':   [j['jahr'] for j in jahre_raw],
        'jahr':    jahr,
    })


@app.route('/api/karte/benachrichtigung-quittieren', methods=['POST'])
@login_required
def api_karte_benachrichtigung_quittieren():
    execute("UPDATE mitarbeiter SET karte_benachrichtigung=NULL WHERE id=?", (session['user_id'],))
    session.pop('karte_benachrichtigung', None)
    return jsonify({'ok': True})


# ─── Main ─────────────────────────────────────────────────────────────────────

# Wird von gunicorn (Railway) beim Import ausgeführt
init_db()

# ── Hintergrund-Scheduler ─────────────────────────────────────────────────────
try:
    from apscheduler.schedulers.background import BackgroundScheduler
    _scheduler = BackgroundScheduler(daemon=True, timezone='Europe/Berlin')
    _scheduler.add_job(backup_db,           'interval', days=1,  id='backup_db',      replace_existing=True)
    _scheduler.add_job(cleanup_alte_fotos,  'interval', days=1,  id='cleanup_fotos',  replace_existing=True)
    _scheduler.add_job(auto_export_job,     'interval', weeks=4, id='auto_export',    replace_existing=True)
    _scheduler.add_job(send_wochenbericht,  'cron', day_of_week='mon', hour=7, minute=0,
                       id='wochenbericht', replace_existing=True, timezone='Europe/Berlin')
    _scheduler.start()
    app.logger.info("Scheduler gestartet (Backup täglich, Foto-Cleanup täglich, Auto-Export 4-wöchentlich, Wochenbericht montags 07:00)")
except ImportError:
    app.logger.warning("APScheduler nicht installiert – automatische Jobs deaktiviert.")

if __name__ == '__main__':
    init_db()
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('FLASK_ENV') != 'production'
    print("\n" + "="*55)
    print("  Aktions Tracker gestartet!")
    print(f"  http://127.0.0.1:{port}")
    print("  Admin-Login: ADMIN / admin123")
    print("  Rep-Login:   z.B. MM / demo123")
    print("="*55 + "\n")
    app.run(debug=debug, host='0.0.0.0', port=port)
