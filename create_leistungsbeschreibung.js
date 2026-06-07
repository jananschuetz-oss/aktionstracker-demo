// create_leistungsbeschreibung.js – Anlage B zum IT-Dienstleistungsvertrag
// Leistungsbeschreibung Aktions Tracker – Arcobräu
// Ausgabe: Leistungsbeschreibung_Aktions_Tracker_Arcobräu.docx

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, BorderStyle, WidthType,
  VerticalAlign, PageNumber
} = require('docx');
const fs = require('fs');

// ── Farben ───────────────────────────────────────────────────────────────────
const BLUE   = '1A3A5C';
const ORANGE = 'C8860A';
const GREY   = 'F2F6FA';
const MID    = '4A6FA5';
const GREEN  = '1E7E34';

// ── Konfiguration ────────────────────────────────────────────────────────────
const DEMO   = true;
const DATEI  = DEMO
  ? 'Leistungsbeschreibung_Aktions_Tracker_Blanko.docx'
  : 'Leistungsbeschreibung_Aktions_Tracker_Arcobräu.docx';

const KUNDE_FIRMA = DEMO ? '[Firma des Auftraggebers]' : 'Arcobräu Gräfliches Brauhaus GmbH & Co. KG';
const PAKET       = DEMO ? '[Paketname]'               : 'Pro';
const MAX_REPS    = DEMO ? '[Anzahl]'                  : '5';
const MONATLICH   = DEMO ? '[Betrag] €'               : '229,00 €';
const EINRICHTUNG = DEMO ? '[Betrag] €'               : '699,00 €';

// ── Helfer ───────────────────────────────────────────────────────────────────
function border(col) { return { style: BorderStyle.SINGLE, size: 4, color: col || 'BBCCDD' }; }
function borders(col) { const b = border(col); return { top:b, bottom:b, left:b, right:b }; }
function noBorder()   { const n = { style: BorderStyle.NIL, size:0, color:'FFFFFF' }; return { top:n,bottom:n,left:n,right:n }; }

function p(text, opts) {
  opts = opts || {};
  return new Paragraph({
    spacing: { after: opts.afterPt ? opts.afterPt*20 : 160, before: opts.beforePt ? opts.beforePt*20 : 0 },
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.JUSTIFIED,
    children: [new TextRun({
      text: text, font: 'Arial',
      size: opts.size || 22,
      bold: opts.bold || false,
      color: opts.color || '000000',
      italics: opts.italic || false,
    })]
  });
}

function sectionHead(text) {
  return new Paragraph({
    spacing: { before: 320, after: 120 },
    children: [new TextRun({ text: text, font: 'Arial', size: 24, bold: true, color: BLUE })]
  });
}

function subHead(text) {
  return new Paragraph({
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text: text, font: 'Arial', size: 22, bold: true, color: MID })]
  });
}

function bullet(text, opts) {
  opts = opts || {};
  return new Paragraph({
    spacing: { after: 80, before: 0 },
    indent: { left: 480, hanging: 260 },
    children: [
      new TextRun({ text: '•  ', font: 'Arial', size: 22, color: opts.color || MID }),
      new TextRun({ text: text, font: 'Arial', size: 22, bold: opts.bold||false, color: '000000' })
    ]
  });
}

function bulletNein(text) {
  return new Paragraph({
    spacing: { after: 80, before: 0 },
    indent: { left: 480, hanging: 260 },
    children: [
      new TextRun({ text: '✗  ', font: 'Arial', size: 22, color: 'C0392B' }),
      new TextRun({ text: text, font: 'Arial', size: 22, color: '000000' })
    ]
  });
}

function ph(text) {
  return new TextRun({ text: text, font: 'Arial', size: 22, bold: true, color: ORANGE });
}

function empty(n) {
  var arr = [];
  for (var i = 0; i < n; i++) arr.push(new Paragraph({ spacing:{after:0}, children:[new TextRun({text:''})] }));
  return arr;
}

function hr() {
  return new Paragraph({
    spacing: { before: 160, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'BBCCDD' } },
    children: []
  });
}

// ── Paket-Vergleichstabelle ───────────────────────────────────────────────────
function paketTable() {
  var bCol = 'BBCCDD';

  function hCell(text, w, highlight) {
    return new TableCell({
      width:{size:w,type:WidthType.DXA},
      borders: borders(highlight ? BLUE : bCol),
      shading:{fill: highlight ? BLUE : MID, type:'clear'},
      margins:{top:100,bottom:100,left:120,right:120},
      children:[new Paragraph({alignment:AlignmentType.CENTER, children:[
        new TextRun({text:text,font:'Arial',size:20,bold:true,color:'FFFFFF'})
      ]})]
    });
  }

  function labelCell(text, w) {
    return new TableCell({
      width:{size:w,type:WidthType.DXA},
      borders:borders(bCol),
      shading:{fill:GREY,type:'clear'},
      margins:{top:60,bottom:60,left:120,right:120},
      children:[new Paragraph({children:[new TextRun({text:text,font:'Arial',size:20,bold:true,color:BLUE})]})]
    });
  }

  function valCell(text, w, bold, highlight) {
    return new TableCell({
      width:{size:w,type:WidthType.DXA},
      borders:borders(bCol),
      shading:{fill: highlight ? 'E8F4F8' : 'FFFFFF', type:'clear'},
      margins:{top:60,bottom:60,left:120,right:120},
      children:[new Paragraph({alignment:AlignmentType.CENTER, children:[
        new TextRun({text:text,font:'Arial',size:20,bold:bold||false,color: highlight ? BLUE : '333333'})
      ]})]
    });
  }

  var rows = [
    new TableRow({ children: [
      hCell('Merkmal', 3360, false),
      hCell('Starter', 2000, false),
      hCell('Team', 2000, false),
      hCell('Pro', 2000, true),
    ]}),
  ];

  var data = [
    ['Max. Repräsentanten',     'bis 5',        'bis 12',       'bis 20'],
    ['Max. Verkaufsleiter',     '1',            '2',            '3'],
    ['Aktivitätserfassung',     'Ja',           'Ja',           'Ja'],
    ['Foto-Upload pro Besuch',  'Ja',           'Ja',           'Ja'],
    ['Dashboard / KPIs',        'Ja',           'Ja',           'Ja'],
    ['IST/SOLL-Vergleich',      'Ja',           'Ja',           'Ja'],
    ['Excel-Export',            'Ja',           'Ja',           'Ja'],
    ['Vertretungsregelung',     '–',            'Ja',           'Ja'],
    ['Mehrere Display-Typen',   '–',            'Ja',           'Ja'],
    ['Passwort-Reset per Mail', 'Ja',           'Ja',           'Ja'],
    ['Mobile PWA (iOS/Android)','Ja',           'Ja',           'Ja'],
    ['Einrichtungsgebühr',      '299,00 €',     '499,00 €',     '699,00 €'],
    ['Monatspauschale',         '99,00 €',      '159,00 €',     '229,00 €'],
  ];

  data.forEach(function(row) {
    rows.push(new TableRow({ children: [
      labelCell(row[0], 3360),
      valCell(row[1], 2000, false, false),
      valCell(row[2], 2000, false, false),
      valCell(row[3], 2000, true,  true),
    ]}));
  });

  return new Table({ width:{size:9360,type:WidthType.DXA}, columnWidths:[3360,2000,2000,2000], rows:rows });
}

// ── SLA-Tabelle ───────────────────────────────────────────────────────────────
function slaTable() {
  var bCol = 'BBCCDD';
  var headerRow = new TableRow({ children: [
    new TableCell({ width:{size:2400,type:WidthType.DXA}, borders:borders(BLUE), shading:{fill:BLUE,type:'clear'}, margins:{top:80,bottom:80,left:120,right:120},
      children:[new Paragraph({children:[new TextRun({text:'Priorität',font:'Arial',size:20,bold:true,color:'FFFFFF'})]})] }),
    new TableCell({ width:{size:2960,type:WidthType.DXA}, borders:borders(BLUE), shading:{fill:BLUE,type:'clear'}, margins:{top:80,bottom:80,left:120,right:120},
      children:[new Paragraph({children:[new TextRun({text:'Beschreibung',font:'Arial',size:20,bold:true,color:'FFFFFF'})]})] }),
    new TableCell({ width:{size:2000,type:WidthType.DXA}, borders:borders(BLUE), shading:{fill:BLUE,type:'clear'}, margins:{top:80,bottom:80,left:120,right:120},
      children:[new Paragraph({children:[new TextRun({text:'Reaktionszeit',font:'Arial',size:20,bold:true,color:'FFFFFF'})]})] }),
    new TableCell({ width:{size:2000,type:WidthType.DXA}, borders:borders(BLUE), shading:{fill:BLUE,type:'clear'}, margins:{top:80,bottom:80,left:120,right:120},
      children:[new Paragraph({children:[new TextRun({text:'Behebungsziel',font:'Arial',size:20,bold:true,color:'FFFFFF'})]})] }),
  ]});

  var slaDaten = [
    ['Kritisch',  'System komplett nicht erreichbar',             '2 Stunden',    '8 Stunden'],
    ['Hoch',      'Kerffunktion eingeschränkt oder fehlerhaft',   '4 Stunden',    '1 Werktag'],
    ['Mittel',    'Einzelne Funktion eingeschränkt',              '1 Werktag',    '3 Werktage'],
    ['Niedrig',   'Kosmetischer Fehler, Verbesserungswunsch',     '3 Werktage',   'nach Absprache'],
  ];

  var rows = [headerRow];
  slaDaten.forEach(function(row, i) {
    var bg = (i % 2 === 0) ? 'FFFFFF' : GREY;
    rows.push(new TableRow({ children: [
      new TableCell({ width:{size:2400,type:WidthType.DXA}, borders:borders(bCol), shading:{fill:bg,type:'clear'}, margins:{top:60,bottom:60,left:120,right:120},
        children:[new Paragraph({children:[new TextRun({text:row[0],font:'Arial',size:20,bold:true,color:MID})]})] }),
      new TableCell({ width:{size:2960,type:WidthType.DXA}, borders:borders(bCol), shading:{fill:bg,type:'clear'}, margins:{top:60,bottom:60,left:120,right:120},
        children:[new Paragraph({children:[new TextRun({text:row[1],font:'Arial',size:20,color:'333333'})]})] }),
      new TableCell({ width:{size:2000,type:WidthType.DXA}, borders:borders(bCol), shading:{fill:bg,type:'clear'}, margins:{top:60,bottom:60,left:120,right:120},
        children:[new Paragraph({children:[new TextRun({text:row[2],font:'Arial',size:20,color:'333333'})]})] }),
      new TableCell({ width:{size:2000,type:WidthType.DXA}, borders:borders(bCol), shading:{fill:bg,type:'clear'}, margins:{top:60,bottom:60,left:120,right:120},
        children:[new Paragraph({children:[new TextRun({text:row[3],font:'Arial',size:20,color:'333333'})]})] }),
    ]}));
  });
  return new Table({ width:{size:9360,type:WidthType.DXA}, columnWidths:[2400,2960,2000,2000], rows:rows });
}

// ════════════════════════════════════════════════════════════════════════════
// DOKUMENT
// ════════════════════════════════════════════════════════════════════════════

var headerTitle = DEMO
  ? 'Leistungsbeschreibung  |  Aktions Tracker'
  : 'Leistungsbeschreibung  |  Aktions Tracker – Arcobräu';

var doc = new Document({
  sections: [{
    properties: {
      page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 1134, left: 1134, right: 1134 } }
    },
    headers: {
      default: new Header({ children: [
        new Paragraph({
          spacing: { after: 80 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'BBCCDD' } },
          children: [
            new TextRun({ text: headerTitle, font: 'Arial', size: 18, bold: true, color: BLUE }),
            new TextRun({ text: '    Anlage B zum IT-Dienstleistungsvertrag', font: 'Arial', size: 18, color: '888888' }),
          ]
        })
      ]})
    },
    footers: {
      default: new Footer({ children: [
        new Paragraph({
          spacing: { before: 80 },
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'BBCCDD' } },
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: 'Seite ', font: 'Arial', size: 16, color: '888888' }),
            new TextRun({ children: [PageNumber.CURRENT], font: 'Arial', size: 16, color: '888888' }),
            new TextRun({ text: ' von ', font: 'Arial', size: 16, color: '888888' }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], font: 'Arial', size: 16, color: '888888' }),
            new TextRun({ text: '  ·  Anlage B – Leistungsbeschreibung', font: 'Arial', size: 16, color: 'AAAAAA', italics: true }),
          ]
        })
      ]})
    },
    children: [
      // ── Titel ─────────────────────────────────────────────────────────
      new Paragraph({ spacing:{after:80,before:200}, alignment:AlignmentType.CENTER,
        children:[new TextRun({text:'LEISTUNGSBESCHREIBUNG', font:'Arial', size:44, bold:true, color:BLUE})] }),
      new Paragraph({ spacing:{after:80}, alignment:AlignmentType.CENTER,
        children:[new TextRun({text: DEMO ? 'Aktions Tracker' : 'Aktions Tracker – Arcobräu', font:'Arial', size:28, bold:true, color:MID})] }),
      new Paragraph({ spacing:{after:80}, alignment:AlignmentType.CENTER,
        children:[new TextRun({text:'Webbasierte Außendienst-Software (SaaS)  ·  Anlage B zum IT-Dienstleistungsvertrag', font:'Arial', size:20, color:'555555', italics:true})] }),
      new Paragraph({ spacing:{after:40}, alignment:AlignmentType.CENTER,
        children: DEMO
          ? [new TextRun({text:'Auftraggeber: ',font:'Arial',size:20,color:'555555'}), ph('[Firma des Auftraggebers]'),
             new TextRun({text:'   ·   Paket: ',font:'Arial',size:20,color:'555555'}), ph('[Paketname]')]
          : [new TextRun({text:'Auftraggeber: ',font:'Arial',size:20,color:'555555'}),
             new TextRun({text:KUNDE_FIRMA,font:'Arial',size:20,bold:true,color:BLUE}),
             new TextRun({text:'   ·   Paket: ',font:'Arial',size:20,color:'555555'}),
             new TextRun({text:PAKET,font:'Arial',size:20,bold:true,color:BLUE}),
             new TextRun({text:'   ·   Datum: ',font:'Arial',size:20,color:'555555'}),
             ph('[DATUM]')]
      }),
      ...empty(1),
      hr(),
      ...empty(1),

      // ── § 1 Vertragsgegenstand ─────────────────────────────────────────
      sectionHead('§ 1  Vertragsgegenstand'),
      p('Der Auftragnehmer (AN) entwickelt, betreibt und pflegt für den Auftraggeber (AG) eine webbasierte Außendienst-Software (SaaS) unter der Bezeichnung „Aktions Tracker". Die Software dient der digitalen Erfassung, Auswertung und Steuerung von Verkaufsaktivitäten des Außendienstteams des AG. Diese Leistungsbeschreibung definiert verbindlich, welche Leistungen im vereinbarten Paket enthalten sind und welche nicht.'),
      ...empty(1),

      // ── § 2 Funktionsumfang ────────────────────────────────────────────
      sectionHead('§ 2  Funktionsumfang der Software'),

      subHead('2.1  Aktivitätserfassung'),
      bullet('Tagesaktuelle Erfassung von Kundenbesuchen: Datum, Verkaufsstelle, Mitarbeiter'),
      bullet('Erfassung von Displays nach Typ und Anzahl (konfigurierbare Display-Kategorien)'),
      bullet('Erfassung von Produktbestellungen nach Sorte und Menge'),
      bullet('Freitextnotizen pro Besuch'),
      bullet('Foto-Upload pro Besuch (Regalbilder, Aktionsflächen) – max. 16 MB/Foto, Formate: JPG, PNG, GIF, WebP, HEIC'),
      bullet('Vertretungsregelung: Mitarbeiter kann Verkaufsstellen abwesender Kollegen betreuen (im zugewiesenen Paket)'),
      ...empty(1),

      subHead('2.2  Dashboard und Auswertungen'),
      bullet('Echtzeit-KPI-Dashboard: Displays gesamt, Kisten gesamt, Besuche gesamt, aktive Mitarbeiter'),
      bullet('Balkendiagramm Kalenderwochen-Verlauf (Displays + Kisten, zwei Y-Achsen)'),
      bullet('Doughnut-Diagramm Top-6-Produkte nach Absatzmenge'),
      bullet('Wochenübersichtstabelle mit KW-Vergleich'),
      bullet('Mitarbeiter-Ranking (nur Verkaufsleiter/Admin)'),
      bullet('Jahres-Filterung, Mitarbeiter-Filterung (Verkaufsleiter-Ansicht)'),
      ...empty(1),

      subHead('2.3  IST/SOLL-Vergleich (Vergleichsansicht)'),
      bullet('Individuelle Jahresziele je Mitarbeiter (Displays und Kisten)'),
      bullet('Saisonaler Zielkurs basierend auf monatsgewichtetem Verteilungsschlüssel'),
      bullet('Ampel-Badge: „Auf Zielkurs" / „Leicht hinter Plan" / „Ziel gefährdet"'),
      bullet('Kumulative Verlaufskurven (IST vs. SOLL) je Mitarbeiter'),
      bullet('Teamziel gesamt'),
      ...empty(1),

      subHead('2.4  Stammdaten- und Benutzerverwaltung (Admin)'),
      bullet('Mitarbeiterverwaltung: Anlegen, Bearbeiten, Passwort setzen, Rolle zuweisen'),
      bullet('Verkaufsstellenverwaltung: Anlegen, Bearbeiten, exklusive Zuweisung (1 Mitarbeiter pro Stelle)'),
      bullet('Produktverwaltung (Biersorten / Produkte): Anlegen, Bearbeiten, Aktivieren/Deaktivieren'),
      bullet('Display-Typ-Verwaltung: Kategorien anlegen und verwalten'),
      bullet('Vertretungsregelungen anlegen und löschen'),
      bullet('Excel-Import für Stammdaten (Mitarbeiter, Verkaufsstellen, Produkte)'),
      bullet('Datenbankdownload (Admin-Backup)'),
      ...empty(1),

      subHead('2.5  Datensicherheit und Infrastruktur'),
      bullet('HTTPS-verschlüsselte Verbindung (TLS)'),
      bullet('Rollenbasierte Zugriffskontrolle: Admin → Verkaufsleiter → Repräsentant'),
      bullet('Session-Timeout nach 8 Stunden Inaktivität'),
      bullet('Automatisches tägliches Datenbank-Backup (7 Tage Aufbewahrung)'),
      bullet('Betrieb auf Railway-Infrastruktur (EU-Region, persistentes Volume)'),
      bullet('Jeder Kunde erhält eigene, isolierte Instanz (keine gemeinsame Datenbank)'),
      ...empty(1),

      subHead('2.6  Export'),
      bullet('Excel-Export mit 4 Tabellenblättern: KW-Übersicht, Mitarbeiter-Ranking, Aktivitäten-Detail, Produkt-Übersicht'),
      bullet('Dateiname: Aktions_Tracker_{Jahr}.xlsx'),
      ...empty(1),

      subHead('2.7  Mobile Nutzung (Progressive Web App)'),
      bullet('Vollständig mobil optimiert – nutzbar auf Smartphone und Tablet'),
      bullet('Installierbar auf iOS und Android als PWA (kein App Store erforderlich)'),
      bullet('Offline-Indikation; alle Kernfunktionen erfordern Internetverbindung'),
      ...empty(1),

      subHead('2.8  Passwort-Management'),
      bullet('Admin setzt Passwörter für Mitarbeiter'),
      bullet('Mitarbeiter ändern eigenes Passwort (aktuelles + neues + Wiederholung)'),
      bullet('Passwort vergessen: Reset-Link per E-Mail (1 Stunde gültig), erfordert konfiguriertes SMTP-Konto'),
      ...empty(1),

      // ── § 3 Pakete ────────────────────────────────────────────────────
      sectionHead('§ 3  Pakete und Nutzergrenzen'),
      p('Die Software wird in drei Paketen angeboten. Der AG hat das nachfolgend markierte Paket gebucht:'),
      ...empty(1),
      paketTable(),
      ...empty(1),
      p('Das gebuchte Paket des AG ist in der Tabelle farblich hervorgehoben. Alle Preise sind Nettobeträge gemäß § 19 UStG (Kleinunternehmer, keine Umsatzsteuer ausgewiesen).', { italic: true, color: '666666', size: 20 }),
      ...empty(1),

      // ── § 4 Nicht enthaltene Leistungen ──────────────────────────────
      sectionHead('§ 4  Nicht enthaltene Leistungen'),
      p('Folgende Leistungen sind ausdrücklich nicht Bestandteil des vereinbarten Pakets und können gegen gesonderte Vergütung beauftragt werden (Anlage C):'),
      bulletNein('Native Mobile Apps (iOS App Store / Google Play Store)'),
      bulletNein('ERP-, CRM- oder Warenwirtschafts-Integration (SAP, Salesforce, etc.)'),
      bulletNein('Schnittstellen / REST-API für Drittsysteme'),
      bulletNein('Custom-Berichte oder individuell programmierte Auswertungen'),
      bulletNein('Mehrsprachigkeit der Oberfläche (aktuell nur Deutsch)'),
      bulletNein('Single Sign-On (SSO) / Active Directory / LDAP-Anbindung'),
      bulletNein('White-Label-Lizenzierung oder Weiterverkauf der Software'),
      bulletNein('Individuelle Schulungen über die Einrichtungsschulung hinaus (auf Anfrage buchbar)'),
      bulletNein('Buchhaltungs- oder Kassenfunktionen'),
      bulletNein('Automatisierte Rechnungserstellung für den AG-internen Gebrauch'),
      ...empty(1),

      // ── § 5 Technische Voraussetzungen ────────────────────────────────
      sectionHead('§ 5  Technische Voraussetzungen (Kundenseite)'),
      p('Der AG stellt auf eigene Kosten sicher:'),
      bullet('Aktueller Webbrowser (Chrome 90+, Firefox 88+, Safari 14+, Edge 90+)'),
      bullet('Stabile Internetverbindung (mind. 5 Mbit/s empfohlen)'),
      bullet('Endgeräte der Mitarbeiter: Smartphone, Tablet oder PC mit aktuellem Browser'),
      bullet('E-Mail-Adresse pro Benutzer (optional, nur für Passwort-Reset-Funktion erforderlich)'),
      bullet('Für SMTP/Passwort-Reset: Konfiguration eines Gmail-Kontos mit App-Passwort (Anleitung durch AN)'),
      ...empty(1),

      // ── § 6 Einrichtungsleistungen ────────────────────────────────────
      sectionHead('§ 6  Einrichtungsleistungen (einmalig)'),
      p('Im Rahmen der einmaligen Einrichtungsgebühr (' + (DEMO ? ph('[Betrag] €') : new TextRun({text: EINRICHTUNG, font:'Arial', size:22, bold:true, color:BLUE})) + ') erbringt der AN folgende Leistungen:'),
      bullet('Einrichten der Software-Instanz (Railway-Deployment, Datenbank, Domain)'),
      bullet('Konfiguration der Firmenbezeichnung, des Logos und der Branding-Farben'),
      bullet('Import der Stammdaten (Mitarbeiter, Verkaufsstellen, Produkte) per Excel-Vorlage oder manuell'),
      bullet('Einrichtung der Benutzerkonten und Zuweisung von Rollen und Verkaufsstellen'),
      bullet('Konfiguration des E-Mail-SMTP-Kontos für Passwort-Reset'),
      bullet('Hinterlegung der Jahresziele (Displays und Kisten pro Mitarbeiter)'),
      bullet('Eine Einführungsschulung (Video-Call, ca. 60 Minuten) für Admin und Verkaufsleiter'),
      bullet('Bereitstellung einer Kurzanleitung für Außendienstmitarbeiter'),
      ...empty(1),
      p('Die Einrichtung ist abgeschlossen, wenn der AG den Zugang erfolgreich getestet hat. Der AN strebt eine Bereitstellung innerhalb von 5 Werktagen nach Vertragsunterzeichnung an.'),
      ...empty(1),

      // ── § 7 SLA ───────────────────────────────────────────────────────
      sectionHead('§ 7  Verfügbarkeit und Reaktionszeiten (SLA)'),
      p('Der AN strebt eine Systemverfügbarkeit von mind. 99 % im Monatsdurchschnitt an (gemessen ohne geplante Wartungsfenster). Die folgenden Reaktions- und Behebungszeiten gelten während der Servicezeiten:'),
      ...empty(1),
      slaTable(),
      ...empty(1),

      new Paragraph({ spacing:{after:160}, children: [
        new TextRun({text:'Servicezeiten: ', font:'Arial', size:22, bold:true, color:BLUE}),
        DEMO
          ? new TextRun({text:'[Servicezeiten eintragen, z. B. Mo–Fr 08:00–18:00 Uhr]', font:'Arial', size:22, bold:true, color:ORANGE})
          : new TextRun({text:'Mo–Fr 09:00–18:00 Uhr (außer gesetzliche Feiertage Bayern und NRW)', font:'Arial', size:22, color:'333333'}),
      ]}),
      p('Geplante Wartungsarbeiten werden mind. 48 Stunden im Voraus per E-Mail angekündigt und finden möglichst außerhalb der Servicezeiten statt.'),
      ...empty(1),

      // ── § 8 Änderungen ────────────────────────────────────────────────
      sectionHead('§ 8  Änderungen des Leistungsumfangs'),
      p('(1) Der AN entwickelt die Software kontinuierlich weiter. Verbesserungen und neue Funktionen innerhalb des vereinbarten Pakets werden ohne zusätzliche Kosten bereitgestellt.'),
      p('(2) Wesentliche Änderungen des Funktionsumfangs, die bestehende Abläufe des AG betreffen, werden mindestens 2 Wochen im Voraus angekündigt.'),
      p('(3) Individuelle Anpassungen außerhalb des Paketumfangs werden auf Anfrage angeboten und gesondert berechnet. Ein schriftliches Angebot wird vorab eingeholt.'),
      p('(4) Preisanpassungen der Monatspauschale erfolgen gemäß § 2 Abs. 4 des IT-Dienstleistungsvertrags.'),
      ...empty(2),
      hr(),

      // ── Unterschriften ────────────────────────────────────────────────
      sectionHead('Unterzeichnung'),
      p('Mit ihrer Unterschrift bestätigen beide Parteien, dass diese Leistungsbeschreibung verbindlicher Bestandteil des IT-Dienstleistungsvertrags ist.'),
      ...empty(1),
      new Table({
        width: { size: 9360, type: WidthType.DXA }, columnWidths: [4320, 720, 4320],
        rows: [new TableRow({ children: [
          (function() {
            var nb2 = noBorder();
            return new TableCell({
              width:{size:4320,type:WidthType.DXA}, borders:nb2, margins:{top:200,bottom:100,left:0,right:0},
              children:[
                new Paragraph({spacing:{after:800},children:[new TextRun({text:'',font:'Arial',size:22})]}),
                new Paragraph({border:{top:{style:BorderStyle.SINGLE,size:4,color:'888888'}},spacing:{after:80},children:[new TextRun({text:'Ort, Datum AG',font:'Arial',size:18,color:'888888',italics:true})]}),
                new Paragraph({spacing:{after:40},children:[new TextRun({text:'Name, Funktion',font:'Arial',size:20,bold:true,color:'333333'})]}),
                new Paragraph({spacing:{after:0}, children:[new TextRun({text: DEMO ? '[Firma des Auftraggebers]' : KUNDE_FIRMA, font:'Arial',size:18,color:'666666'})]})
              ]
            });
          })(),
          new TableCell({width:{size:720,type:WidthType.DXA},borders:noBorder(),children:[new Paragraph({children:[]})]}),
          (function() {
            var nb2 = noBorder();
            return new TableCell({
              width:{size:4320,type:WidthType.DXA}, borders:nb2, margins:{top:200,bottom:100,left:0,right:0},
              children:[
                new Paragraph({spacing:{after:800},children:[new TextRun({text:'',font:'Arial',size:22})]}),
                new Paragraph({border:{top:{style:BorderStyle.SINGLE,size:4,color:'888888'}},spacing:{after:80},children:[new TextRun({text:'Ort, Datum AN',font:'Arial',size:18,color:'888888',italics:true})]}),
                new Paragraph({spacing:{after:40},children:[new TextRun({text:'Name, Inhaber',font:'Arial',size:20,bold:true,color:'333333'})]}),
                new Paragraph({spacing:{after:0}, children:[new TextRun({text:'[Name oder Firma AN]',font:'Arial',size:18,color:ORANGE,bold:true})]})
              ]
            });
          })(),
        ]})]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(function(buffer) {
  fs.writeFileSync(DATEI, buffer);
  console.log('OK: ' + DATEI + ' erstellt (' + buffer.length + ' Bytes)');
}).catch(function(err) {
  console.error('ERROR:', err);
  process.exit(1);
});
