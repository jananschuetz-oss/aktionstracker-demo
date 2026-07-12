#!/usr/bin/env bash
# Browser-basierter Pflicht-Test vor jedem Deploy (agent-browser statt curl).
# Rendert Seiten echt (inkl. JS) und fängt damit auch kaputte Interaktionen,
# die reine HTTP-200-Checks übersehen (siehe CLAUDE.md "Vor jedem Deploy").
#
# Nutzung: im Repo-Root ausführen (wo app.py liegt): ./pflichttest_browser.sh
set -uo pipefail

SESSION="pflichttest-demo-$$"
BASE="http://127.0.0.1:5000"
SHOTDIR="_pflichttest_shots"
mkdir -p "$SHOTDIR"
FAIL=0
SERVER_PID=""

cleanup() {
  agent-browser --session "$SESSION" close --all >/dev/null 2>&1
  [ -n "$SERVER_PID" ] && kill "$SERVER_PID" >/dev/null 2>&1
}
trap cleanup EXIT

echo "== Server starten (Demo-Env) =="
TOUREN_MODUS=an INIT_DEMO_USERS=true venv/Scripts/python app.py > _pflichttest_server.log 2>&1 &
SERVER_PID=$!
sleep 6

check_route() {
  local label="$1" path="$2"
  agent-browser --session "$SESSION" open "$BASE$path" >/dev/null 2>&1
  agent-browser --session "$SESSION" wait --load networkidle >/dev/null 2>&1
  local shotname
  shotname=$(echo "$label" | tr -c 'A-Za-z0-9' '_')
  agent-browser --session "$SESSION" screenshot "$SHOTDIR/${shotname}.png" >/dev/null 2>&1

  local body page_errs console_errs
  body=$(agent-browser --session "$SESSION" read 2>/dev/null)
  page_errs=$(agent-browser --session "$SESSION" errors 2>/dev/null)
  console_errs=$(agent-browser --session "$SESSION" console 2>/dev/null | grep -i "error" || true)

  if echo "$body" | grep -qi "Internal Server Error"; then
    printf "%-30s FAIL (500 Internal Server Error)\n" "$label"
    FAIL=1
  elif [ -n "$page_errs" ]; then
    printf "%-30s WARN  JS-Fehler: %s\n" "$label" "$(echo "$page_errs" | head -1)"
  elif [ -n "$console_errs" ]; then
    printf "%-30s WARN  Console-Error: %s\n" "$label" "$(echo "$console_errs" | head -1)"
  else
    printf "%-30s OK\n" "$label"
  fi
}

login() {
  local user="$1" pass="$2"
  agent-browser --session "$SESSION" open "$BASE/logout" >/dev/null 2>&1
  agent-browser --session "$SESSION" open "$BASE/" >/dev/null 2>&1
  agent-browser --session "$SESSION" wait --load domcontentloaded >/dev/null 2>&1
  agent-browser --session "$SESSION" fill "input[name=email]" "$user" >/dev/null
  agent-browser --session "$SESSION" fill "input[name=passwort]" "$pass" >/dev/null
  agent-browser --session "$SESSION" click "button[type=submit]" >/dev/null
  agent-browser --session "$SESSION" wait --url "**/dashboard" >/dev/null 2>&1
}

echo
echo "== Admin (admin/admin123) =="
login "admin" "admin123"
check_route "Dashboard (admin)"          "/dashboard"
check_route "Tourenplanung (admin)"      "/tourenplanung"
check_route "Admin-Panel"                "/admin"
check_route "Karte"                      "/karte"
check_route "Aktivitäten"                "/aktivitaeten"
check_route "Neue Aktivität (Formular)"  "/aktivitaet/neu"
check_route "Wochenbericht-Vorschau"     "/einstellungen/wochenbericht/vorschau"
check_route "Monatsbericht-Vorschau"     "/einstellungen/monatsbericht/vorschau"

echo
echo "== Rep (MM/start123) =="
login "MM" "start123"
check_route "Dashboard (rep)"            "/dashboard"
check_route "Neue Aktivität (rep)"       "/aktivitaet/neu"

echo
if [ "$FAIL" -eq 1 ]; then
  echo "FEHLGESCHLAGEN: mindestens eine Route liefert 500 -- erst fixen, dann pushen."
  exit 1
else
  echo "OK: keine 500er. WARN-Zeilen (falls vorhanden) manuell prüfen. Screenshots: $SHOTDIR/"
  exit 0
fi
