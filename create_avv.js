// create_avv.js – Auftragsverarbeitungsvertrag (AVV) gemäß Art. 28 DSGVO
// Anlage A zum IT-Dienstleistungsvertrag – Aktions Tracker – Arcobräu
// Ausgabe: AVV_Auftragsverarbeitungsvertrag_Arcobräu.docx

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
const RED    = 'C0392B';

// ── Kunden-Konfiguration (Arcobräu) ─────────────────────────────────────────
const KUNDE = {
  firma:    'Arcobräu Gräfliches Brauhaus GmbH & Co. KG',
  strasse:  'Brauhausstrasse 1',
  plzOrt:   '83529 Grafling',
  land:     'Deutschland',
  vertreter: '[Name des Geschäftsführers]',
  email:    '[E-Mail AG]',
};
const DATEI = 'AVV_Auftragsverarbeitungsvertrag_Blanko.docx';
const DEMO  = true;

// ── Border-Helfer ────────────────────────────────────────────────────────────
function border(col) {
  return { style: BorderStyle.SINGLE, size: 4, color: col || 'BBCCDD' };
}
function borders(col) { const b = border(col); return { top:b, bottom:b, left:b, right:b }; }
function noBorder()   { const n = { style: BorderStyle.NIL, size:0, color:'FFFFFF' }; return { top:n,bottom:n,left:n,right:n }; }
const nb = noBorder();

// ── Paragraph-Helfer ─────────────────────────────────────────────────────────
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

function pRuns(runs, opts) {
  opts = opts || {};
  return new Paragraph({
    spacing: { after: opts.afterPt ? opts.afterPt*20 : 160, before: opts.beforePt ? opts.beforePt*20 : 0 },
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.JUSTIFIED,
    children: runs.map(function(r) {
      return new TextRun({ text: r.text, font: 'Arial', size: r.size||opts.size||22,
        bold: r.bold||false, color: r.color||'000000', italics: r.italic||false });
    })
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
      new TextRun({ text: '•  ', font: 'Arial', size: 22, color: MID }),
      new TextRun({ text: text, font: 'Arial', size: 22, bold: opts.bold||false, color: opts.color||'000000' })
    ]
  });
}

function ph(text) {
  return new TextRun({ text: text, font: 'Arial', size: 22, bold: true, color: ORANGE });
}

function empty(n) {
  var arr = [];
  for (var i = 0; i < n; i++) arr.push(new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: '' })] }));
  return arr;
}

function hr() {
  return new Paragraph({
    spacing: { before: 160, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'BBCCDD' } },
    children: []
  });
}

// ── Einfache Tabellenzelle ───────────────────────────────────────────────────
function tc(content, w, opts) {
  opts = opts || {};
  var children = Array.isArray(content) ? content : [p(content, { size: 20 })];
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders: borders(opts.borderCol || 'BBCCDD'),
    shading: opts.shade ? { fill: opts.shade, type: 'clear' } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.TOP,
    children: children,
  });
}

// ── Vertragsparteien-Box ─────────────────────────────────────────────────────
function partnerBox() {
  function cell(title, lines, w) {
    var children = [
      new Paragraph({ spacing: { after: 100 }, children: [
        new TextRun({ text: title, font: 'Arial', size: 22, bold: true, color: BLUE })
      ]})
    ].concat(lines.map(function(l) {
      return new Paragraph({ spacing: { after: 60 }, children: l });
    }));
    return new TableCell({
      width: { size: w, type: WidthType.DXA },
      borders: borders('BBCCDD'),
      shading: { fill: GREY, type: 'clear' },
      margins: { top: 140, bottom: 140, left: 200, right: 200 },
      children: children
    });
  }

  function lv(label, val, isOrange) {
    return [
      new TextRun({ text: label + ': ', font: 'Arial', size: 20, bold: true, color: '333333' }),
      isOrange
        ? new TextRun({ text: val, font: 'Arial', size: 20, bold: true, color: ORANGE })
        : new TextRun({ text: val, font: 'Arial', size: 20, color: '333333' })
    ];
  }

  var agLines = DEMO ? [
    lv('Firma',    '[Firma des Auftraggebers]', true),
    lv('Anschrift','[Straße, PLZ Ort]', true),
    lv('Vertr. d.','[Name des Geschäftsführers]', true),
    lv('E-Mail',   '[E-Mail AG]', true),
  ] : [
    lv('Firma',    KUNDE.firma, false),
    lv('Anschrift',KUNDE.strasse + ', ' + KUNDE.plzOrt, false),
    lv('Vertr. d.', KUNDE.vertreter, true),
    lv('E-Mail',   KUNDE.email, true),
  ];

  var anLines = [
    lv('Name / Firma', '[Name oder Firma AN]', true),
    lv('Anschrift',    '[Straße, PLZ Ort AN]', true),
    lv('Vertr. d.',    '[Name Inhaber / GF]', true),
    lv('E-Mail',       '[E-Mail AN]', true),
  ];

  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [4560, 4800],
    rows: [new TableRow({ children: [
      cell('Verantwortlicher (AG) – Auftraggeber', agLines, 4560),
      cell('Auftragsverarbeiter (AN) – Auftragnehmer', anLines, 4800),
    ]})]
  });
}

// ── Unterauftragsverarbeiter-Tabelle ─────────────────────────────────────────
function subProcessorTable() {
  var bCol = 'BBCCDD';
  var headerRow = new TableRow({
    children: [
      new TableCell({ width:{size:2800,type:WidthType.DXA}, borders:borders(BLUE), shading:{fill:BLUE,type:'clear'}, margins:{top:80,bottom:80,left:120,right:120},
        children:[new Paragraph({children:[new TextRun({text:'Unternehmen',font:'Arial',size:20,bold:true,color:'FFFFFF'})]})] }),
      new TableCell({ width:{size:2800,type:WidthType.DXA}, borders:borders(BLUE), shading:{fill:BLUE,type:'clear'}, margins:{top:80,bottom:80,left:120,right:120},
        children:[new Paragraph({children:[new TextRun({text:'Leistung',font:'Arial',size:20,bold:true,color:'FFFFFF'})]})] }),
      new TableCell({ width:{size:1800,type:WidthType.DXA}, borders:borders(BLUE), shading:{fill:BLUE,type:'clear'}, margins:{top:80,bottom:80,left:120,right:120},
        children:[new Paragraph({children:[new TextRun({text:'Server-Standort',font:'Arial',size:20,bold:true,color:'FFFFFF'})]})] }),
      new TableCell({ width:{size:1960,type:WidthType.DXA}, borders:borders(BLUE), shading:{fill:BLUE,type:'clear'}, margins:{top:80,bottom:80,left:120,right:120},
        children:[new Paragraph({children:[new TextRun({text:'Datenschutz-Info',font:'Arial',size:20,bold:true,color:'FFFFFF'})]})] }),
    ]
  });

  function dataRow(name, leistung, standort, dsInfo, shade) {
    var bg = shade ? GREY : 'FFFFFF';
    function dc(text, w) { return new TableCell({ width:{size:w,type:WidthType.DXA}, borders:borders(bCol), shading:{fill:bg,type:'clear'}, margins:{top:60,bottom:60,left:120,right:120},
      children:[new Paragraph({children:[new TextRun({text:text,font:'Arial',size:20,color:'333333'})]})] }); }
    return new TableRow({ children: [dc(name,2800), dc(leistung,2800), dc(standort,1800), dc(dsInfo,1960)] });
  }

  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2800, 2800, 1800, 1960],
    rows: [
      headerRow,
      dataRow('Railway Inc.', 'Hosting, Datenbankbetrieb, Server-Infrastruktur', 'EU-Region (Frankfurt)', 'privacy.railway.app', false),
      dataRow('Google LLC', 'E-Mail-Versand via SMTP (Gmail, nur Passwort-Reset)', 'USA (SCCs)', 'policies.google.com/privacy', true),
    ]
  });
}

// ── TOM-Tabelle ───────────────────────────────────────────────────────────────
function tomTable() {
  var bCol = 'BBCCDD';
  var headerRow = new TableRow({
    children: [
      new TableCell({ width:{size:2600,type:WidthType.DXA}, borders:borders(BLUE), shading:{fill:BLUE,type:'clear'}, margins:{top:80,bottom:80,left:120,right:120},
        children:[new Paragraph({children:[new TextRun({text:'Maßnahme',font:'Arial',size:20,bold:true,color:'FFFFFF'})]})] }),
      new TableCell({ width:{size:6760,type:WidthType.DXA}, borders:borders(BLUE), shading:{fill:BLUE,type:'clear'}, margins:{top:80,bottom:80,left:120,right:120},
        children:[new Paragraph({children:[new TextRun({text:'Umsetzung',font:'Arial',size:20,bold:true,color:'FFFFFF'})]})] }),
    ]
  });

  var toms = [
    ['Zutrittskontrolle', 'Betrieb auf Cloud-Infrastruktur (Railway EU-Region); kein physischer Serverzugang durch AN-Personal; Zutrittssicherung obliegt Railway Inc. gemäß deren TOMs.'],
    ['Zugangskontrolle', 'Authentifizierung per Benutzername/Passwort; automatischer Sitzungsablauf nach 8 Stunden; ausschließlich HTTPS/TLS-Verbindung (kein HTTP); Secure- und HttpOnly-Cookies; SameSite-Lax-Schutz.'],
    ['Zugriffskontrolle', 'Rollenbasiertes Berechtigungsmodell (Admin / Verkaufsleiter / Repräsentant); Repräsentanten sehen ausschließlich eigene Aktivitätsdaten; Admin-Bereiche durch Rollenwächter (Decorator) geschützt.'],
    ['Weitergabekontrolle', 'Alle Datenübertragungen TLS-verschlüsselt; API-Endpunkte nur für authentifizierte Sessions erreichbar; Excel/CSV-Exporte nur für autorisierte Nutzer; keine Weitergabe an Dritte außer genannten Unterauftragsverarbeitern.'],
    ['Eingabekontrolle', 'Jede Aktivität wird mit Erstellungszeitpunkt (Timestamp) und Mitarbeiter-ID gespeichert; Änderungen an Stammdaten nur durch Administratoren.'],
    ['Auftragskontrolle', 'Datenbankzugriff ausschließlich durch AN im Rahmen dieses Vertrags; keine Nutzung der Kundendaten für eigene Zwecke; Parametrisierte SQL-Abfragen (kein SQL-Injection-Risiko).'],
    ['Verfügbarkeitskontrolle', 'Automatisiertes tägliches Datenbank-Backup; Aufbewahrung der letzten 7 Tage; persistentes Railway-Volume (Daten überleben Deployments); Railway-Infrastruktur mit eigener Redundanz.'],
    ['Trennungskontrolle', 'Jeder Kunde erhält eine eigene Software-Instanz (eigene Railway-Deployment-Umgebung) mit separater SQLite-Datenbank; keine mandantenfähige Architektur; vollständige Datenisolation zwischen Kunden.'],
  ];

  var rows = [headerRow];
  toms.forEach(function(tom, i) {
    var bg = (i % 2 === 0) ? 'FFFFFF' : GREY;
    rows.push(new TableRow({ children: [
      new TableCell({ width:{size:2600,type:WidthType.DXA}, borders:borders(bCol), shading:{fill:bg,type:'clear'}, margins:{top:80,bottom:80,left:120,right:120},
        children:[new Paragraph({children:[new TextRun({text:tom[0],font:'Arial',size:20,bold:true,color:BLUE})]})] }),
      new TableCell({ width:{size:6760,type:WidthType.DXA}, borders:borders(bCol), shading:{fill:bg,type:'clear'}, margins:{top:80,bottom:80,left:120,right:120},
        children:[new Paragraph({children:[new TextRun({text:tom[1],font:'Arial',size:20,color:'333333'})]})] }),
    ]}));
  });
  return new Table({ width:{size:9360,type:WidthType.DXA}, columnWidths:[2600,6760], rows:rows });
}

// ── Unterschriftenblock ───────────────────────────────────────────────────────
function sigTable() {
  var nb2 = noBorder();
  function sigCell(role, ortDatum, name, firma) {
    return new TableCell({
      width: { size: 4320, type: WidthType.DXA }, borders: nb2,
      margins: { top: 200, bottom: 100, left: 0, right: 0 },
      children: [
        new Paragraph({ spacing:{after:800}, children:[new TextRun({text:'',font:'Arial',size:22})] }),
        new Paragraph({ border:{top:{style:BorderStyle.SINGLE,size:4,color:'888888'}}, spacing:{after:80},
          children:[new TextRun({text:ortDatum,font:'Arial',size:18,color:'888888',italics:true})] }),
        new Paragraph({ spacing:{after:40}, children:[new TextRun({text:name,font:'Arial',size:20,bold:true,color:'333333'})] }),
        new Paragraph({ spacing:{after:40}, children:[new TextRun({text:firma,font:'Arial',size:18,color:'666666'})] }),
        new Paragraph({ spacing:{after:0},  children:[new TextRun({text:role,font:'Arial',size:18,bold:true,color:BLUE})] }),
      ]
    });
  }
  return new Table({
    width: { size: 9360, type: WidthType.DXA }, columnWidths: [4320, 720, 4320],
    rows: [new TableRow({ children: [
      sigCell('Verantwortlicher (AG)', 'Ort, Datum AG', 'Name, Funktion', DEMO ? '[Firma des Auftraggebers]' : KUNDE.firma),
      new TableCell({ width:{size:720,type:WidthType.DXA}, borders:nb2, children:[new Paragraph({children:[]})] }),
      sigCell('Auftragsverarbeiter (AN)', 'Ort, Datum AN', 'Name, Inhaber', '[Name oder Firma AN]'),
    ]})]
  });
}

// ════════════════════════════════════════════════════════════════════════════
// DOKUMENT
// ════════════════════════════════════════════════════════════════════════════

var headerTitle = DEMO
  ? 'AVV – Auftragsverarbeitungsvertrag  |  Aktions Tracker'
  : 'AVV – Auftragsverarbeitungsvertrag  |  Aktions Tracker – Arcobräu';

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
            new TextRun({ text: '    Anlage A zum IT-Dienstleistungsvertrag vom  ', font: 'Arial', size: 18, color: '888888' }),
            new TextRun({ text: '[DATUM]', font: 'Arial', size: 18, bold: true, color: ORANGE }),
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
            new TextRun({ text: '  ·  Vertraulich – nur für Vertragsparteien', font: 'Arial', size: 16, color: 'AAAAAA', italics: true }),
          ]
        })
      ]})
    },
    children: [
      // ── Titel ──────────────────────────────────────────────────────────
      new Paragraph({ spacing:{after:80,before:200}, alignment:AlignmentType.CENTER,
        children:[new TextRun({text:'AUFTRAGSVERARBEITUNGSVERTRAG', font:'Arial', size:44, bold:true, color:BLUE})] }),
      new Paragraph({ spacing:{after:80}, alignment:AlignmentType.CENTER,
        children:[new TextRun({text: DEMO ? 'Aktions Tracker' : 'Aktions Tracker – Arcobräu', font:'Arial', size:28, bold:true, color:MID})] }),
      new Paragraph({ spacing:{after:80}, alignment:AlignmentType.CENTER,
        children:[new TextRun({text:'gemäß Art. 28 DSGVO  ·  Anlage A zum IT-Dienstleistungsvertrag', font:'Arial', size:20, color:'555555', italics:true})] }),
      new Paragraph({ spacing:{after:40}, alignment:AlignmentType.CENTER,
        children:[
          new TextRun({text:'Datum: ', font:'Arial', size:20, color:'555555'}),
          ph('[DATUM]'),
        ] }),
      ...empty(1),
      hr(),
      ...empty(1),

      // ── Parteien ────────────────────────────────────────────────────────
      subHead('Vertragsparteien'),
      ...empty(1),
      partnerBox(),
      ...empty(1),

      // ── Präambel ────────────────────────────────────────────────────────
      subHead('Präambel'),
      p('Dieser Auftragsverarbeitungsvertrag (im Folgenden „AVV") konkretisiert die datenschutzrechtlichen Pflichten der Parteien im Zusammenhang mit dem IT-Dienstleistungsvertrag (im Folgenden „Hauptvertrag") und gilt als dessen Anlage A. Er regelt die Verarbeitung personenbezogener Daten durch den Auftragnehmer (Auftragsverarbeiter/AN) im Auftrag des Auftraggebers (Verantwortlicher/AG) gemäß Art. 28 der Verordnung (EU) 2016/679 (DSGVO).'),
      ...empty(1),
      hr(),

      // ── § 1 ─────────────────────────────────────────────────────────────
      sectionHead('§ 1  Gegenstand und Dauer der Verarbeitung'),
      p('(1) Der AN verarbeitet im Auftrag des AG personenbezogene Daten im Rahmen des Betriebs der webbasierten Software „Aktions Tracker" gemäß dem Hauptvertrag.'),
      p('(2) Die Verarbeitung beginnt mit der Freischaltung des Zugangs zur Software und endet mit dem Ende des Hauptvertrags, vorbehaltlich der Regelungen zur Datenlöschung in § 9.'),
      ...empty(1),

      // ── § 2 ─────────────────────────────────────────────────────────────
      sectionHead('§ 2  Art und Zweck der Verarbeitung'),
      p('Die Verarbeitung personenbezogener Daten umfasst folgende Tätigkeiten:'),
      bullet('Hosting und Betrieb der Software auf Servern des Unterauftragsverarbeiters (EU-Region)'),
      bullet('Speicherung und Abruf von Aktivitätsdaten der Außendienstmitarbeiter des AG'),
      bullet('Automatisierte E-Mail-Benachrichtigungen (Passwort-Reset-Verfahren)'),
      bullet('Erstellen von Auswertungen, Berichten und Exporten auf Anfrage des AG'),
      bullet('Automatisierte Datensicherung (Backup) zum Schutz vor Datenverlust'),
      p('Zweck der Verarbeitung ist ausschließlich die Erbringung der vertraglich vereinbarten Leistungen gemäß Hauptvertrag. Eine Verarbeitung zu eigenen Zwecken des AN ist ausgeschlossen.'),
      ...empty(1),

      // ── § 3 ─────────────────────────────────────────────────────────────
      sectionHead('§ 3  Art der personenbezogenen Daten und Kategorien betroffener Personen'),
      subHead('3.1  Kategorien betroffener Personen'),
      bullet('Außendienstmitarbeiter (Repräsentanten) des AG'),
      bullet('Verkaufsleiter des AG'),
      bullet('Systemadministratoren des AG (sofern benannt)'),
      ...empty(1),
      subHead('3.2  Art der verarbeiteten personenbezogenen Daten'),
      bullet('Stammdaten: Vor- und Nachname, internes Kürzel, Rolle im System (Repräsentant / Verkaufsleiter / Admin)'),
      bullet('Kontaktdaten: E-Mail-Adresse (sofern hinterlegt, ausschließlich für Passwort-Reset-Funktion)'),
      bullet('Aktivitätsdaten: Datum, Uhrzeit, besuchte Verkaufsstellen, Anzahl platzierter Displays nach Typ, bestellte Produktmengen nach Sorte, persönliche Notizen des Mitarbeiters'),
      bullet('Mediendaten: Fotos von Verkaufsregalen oder Aktionsflächen (können Personen im Hintergrund zeigen)'),
      bullet('Technische Daten: Login-Zeitstempel, Session-Metadaten (8-Stunden-Ablauf), IP-Adressen (serverseitig durch Infrastrukturanbieter)'),
      bullet('Zugangsdaten: Passwörter (zugriffsgeschützt gespeichert)'),
      p('Es werden keine besonderen Kategorien personenbezogener Daten gemäß Art. 9 DSGVO verarbeitet.'),
      ...empty(1),

      // ── § 4 ─────────────────────────────────────────────────────────────
      sectionHead('§ 4  Pflichten des Auftragsverarbeiters'),
      p('(1) Weisungsgebundenheit: Der AN verarbeitet personenbezogene Daten ausschließlich auf dokumentierte Weisung des AG. Hält der AN eine Weisung für rechtswidrig, informiert er den AG unverzüglich schriftlich.'),
      p('(2) Vertraulichkeit: Der AN stellt sicher, dass alle zur Datenverarbeitung befugten Personen zur Vertraulichkeit verpflichtet sind oder einer gesetzlichen Verschwiegenheitspflicht unterliegen.'),
      p('(3) Sicherheit der Verarbeitung: Der AN trifft die in Anlage 1 dieses AVV beschriebenen technischen und organisatorischen Maßnahmen gemäß Art. 32 DSGVO.'),
      p('(4) Datenpannenmeldung: Bei einer Verletzung des Schutzes personenbezogener Daten informiert der AN den AG unverzüglich, spätestens jedoch innerhalb von 24 Stunden nach Bekanntwerden, und stellt alle Informationen für eine etwaige Meldepflicht nach Art. 33/34 DSGVO bereit.'),
      p('(5) Beendigung: Nach Beendigung des Hauptvertrags handelt der AN gemäß § 9 dieses AVV (Löschung und Rückgabe).'),
      ...empty(1),

      // ── § 5 ─────────────────────────────────────────────────────────────
      sectionHead('§ 5  Weisungsbefugnis des Verantwortlichen'),
      p('(1) Der AG ist allein für die Rechtmäßigkeit der Verarbeitungszwecke und der ihm überlassenen Daten verantwortlich.'),
      p('(2) Weisungen erteilt der AG in der Regel durch Nutzung der Software-Funktionen (Anlegen, Ändern, Löschen von Daten im System). Darüber hinausgehende Weisungen (z. B. zur vollständigen Datenlöschung) erfolgen schriftlich per E-Mail.'),
      p('(3) Der AG stellt sicher, dass seine Mitarbeiter gemäß Art. 13/14 DSGVO über die Verarbeitung ihrer Daten informiert sind.'),
      p('(4) Der AG benennt intern eine verantwortliche Person für die Kommunikation mit dem AN in datenschutzrechtlichen Fragen.'),
      ...empty(1),

      // ── § 6 ─────────────────────────────────────────────────────────────
      sectionHead('§ 6  Technische und Organisatorische Maßnahmen'),
      p('Die vom AN getroffenen technischen und organisatorischen Maßnahmen (TOMs) gemäß Art. 32 DSGVO sind in Anlage 1 dieses AVV beschrieben. Der AN ist berechtigt, die TOMs anzupassen oder zu verbessern, solange das vereinbarte Schutzniveau nicht unterschritten wird.'),
      ...empty(1),

      // ── § 7 ─────────────────────────────────────────────────────────────
      sectionHead('§ 7  Unterauftragsverarbeiter'),
      p('(1) Der AG erteilt seine generelle Genehmigung für den Einsatz folgender Unterauftragsverarbeiter:'),
      ...empty(1),
      subProcessorTable(),
      ...empty(1),
      p('(2) Bei Wechsel oder Hinzunahme weiterer Unterauftragsverarbeiter informiert der AN den AG schriftlich mit einer Ankündigungsfrist von 14 Tagen. Der AG kann innerhalb dieser Frist schriftlich Einwände erheben; schweigt er nach Fristablauf, gilt die Zustimmung als erteilt.'),
      p('(3) Der AN stellt sicher, dass Unterauftragsverarbeiter, die in Drittländern außerhalb des EWR tätig sind, durch geeignete Garantien gemäß Art. 46 DSGVO (z. B. EU-Standardvertragsklauseln) gebunden sind.'),
      p('(4) Der AN stellt sicher, dass Unterauftragsverarbeiter mindestens gleichwertige Datenschutzverpflichtungen eingehen wie in diesem AVV vereinbart.'),
      ...empty(1),

      // ── § 8 ─────────────────────────────────────────────────────────────
      sectionHead('§ 8  Unterstützung des Verantwortlichen'),
      p('(1) Der AN unterstützt den AG durch geeignete technische und organisatorische Maßnahmen bei der Erfüllung seiner Pflichten bei Anfragen betroffener Personen gemäß Art. 15–22 DSGVO (Auskunft, Berichtigung, Löschung, Einschränkung, Datenübertragbarkeit, Widerspruch).'),
      p('(2) Auf Anfrage exportiert oder löscht der AN die Daten einzelner betroffener Personen aus dem System. Anfragen sind schriftlich per E-Mail zu stellen und werden innerhalb von 5 Werktagen bearbeitet.'),
      p('(3) Der AN unterstützt den AG bei der Erfüllung der Pflichten gemäß Art. 32–36 DSGVO (Datensicherheit, Datenschutz-Folgenabschätzung, Meldepflichten).'),
      ...empty(1),

      // ── § 9 ─────────────────────────────────────────────────────────────
      sectionHead('§ 9  Löschung und Rückgabe von Daten'),
      p('(1) Nach Beendigung des Hauptvertrags stellt der AN dem AG auf Anfrage innerhalb von 14 Tagen einen vollständigen Datenexport (SQLite-Datenbank und/oder CSV-Dateien) kostenlos zur Verfügung.'),
      p('(2) Spätestens 30 Tage nach Vertragsende löscht der AN alle personenbezogenen Daten aus der Produktivumgebung.'),
      p('(3) Backup-Daten werden innerhalb von weiteren 7 Tagen nach Ablauf ihrer regulären Aufbewahrungsfrist gelöscht.'),
      p('(4) Auf Wunsch des AG bestätigt der AN die vollständige Datenlöschung schriftlich.'),
      p('(5) Gesetzliche Aufbewahrungspflichten (z. B. nach HGB oder AO) bleiben unberührt und können einer sofortigen Löschung entgegenstehen.'),
      ...empty(1),

      // ── § 10 ────────────────────────────────────────────────────────────
      sectionHead('§ 10  Kontroll- und Nachweispflichten'),
      p('(1) Der AG hat das Recht, die Einhaltung der DSGVO-Anforderungen durch den AN zu überprüfen – durch Einholung von Nachweisdokumenten oder durch angekündigte Kontrollen (mind. 5 Werktage Vorankündigung).'),
      p('(2) Kontrollen dürfen den Betrieb des AN nicht unverhältnismäßig beeinträchtigen. Entstehende Kosten trägt der AG, sofern kein Verschulden des AN vorliegt.'),
      p('(3) Der AN kann Nachweise durch Vorlage von Drittprüfungszertifikaten oder Prüfberichten der eingesetzten Unterauftragsverarbeiter erbringen.'),
      ...empty(1),

      // ── § 11 ────────────────────────────────────────────────────────────
      sectionHead('§ 11  Schlussbestimmungen'),
      p('(1) Dieser AVV ist Bestandteil des Hauptvertrags. Bei Widersprüchen in datenschutzrechtlichen Fragen hat dieser AVV Vorrang.'),
      p('(2) Änderungen dieses AVV bedürfen der Schriftform.'),
      p('(3) Es gilt das Recht der Bundesrepublik Deutschland. Gerichtsstand ist der Sitz des AN, soweit gesetzlich zulässig.'),
      p('(4) Sollten einzelne Bestimmungen dieses AVV unwirksam sein oder werden, bleibt die Gültigkeit der übrigen Bestimmungen unberührt.'),
      ...empty(2),
      hr(),

      // ── Unterschriften ───────────────────────────────────────────────────
      sectionHead('Unterschriften'),
      p('Mit ihrer Unterschrift bestätigen beide Parteien, diesen Auftragsverarbeitungsvertrag zur Kenntnis genommen zu haben und diesem zuzustimmen.'),
      ...empty(1),
      sigTable(),
      ...empty(2),
      hr(),

      // ── Anlage 1: TOMs ───────────────────────────────────────────────────
      new Paragraph({ spacing:{before:400, after:120},
        children:[new TextRun({text:'Anlage 1 zum AVV: Technische und Organisatorische Maßnahmen (TOMs)', font:'Arial', size:26, bold:true, color:BLUE})] }),
      p('Die nachfolgenden Maßnahmen gemäß Art. 32 DSGVO wurden vom AN zum Zeitpunkt des Vertragsabschlusses getroffen. Der AN ist berechtigt, Maßnahmen zu aktualisieren oder zu verbessern, sofern das Schutzniveau insgesamt nicht abgesenkt wird.'),
      ...empty(1),
      tomTable(),
      ...empty(1),
      p('Stand: ' + new Date().toLocaleDateString('de-DE', {day:'2-digit',month:'2-digit',year:'numeric'}) + ' – Der AN informiert den AG bei wesentlichen Änderungen der TOMs.', { italic: true, color: '666666' }),
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
