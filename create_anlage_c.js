// create_anlage_c.js – Anlage C zum IT-Dienstleistungsvertrag
// Preisliste für Zusatz- und Sonderleistungen – Blanko-Version
// Ausgabe: Anlage_C_Preisliste_Blanko.docx

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
const GREEN  = '27AE60';
const RED    = 'C0392B';

// ── Dateiname ────────────────────────────────────────────────────────────────
const DATEI  = 'Anlage_C_Preisliste_Blanko.docx';

// ── Helfer ───────────────────────────────────────────────────────────────────
function border(col) { return { style: BorderStyle.SINGLE, size: 4, color: col || 'BBCCDD' }; }
function borders(col) { var b = border(col); return { top:b, bottom:b, left:b, right:b }; }
function noBorder()   { var n = { style: BorderStyle.NIL, size:0, color:'FFFFFF' }; return { top:n,bottom:n,left:n,right:n }; }

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

function bullet(text) {
  return new Paragraph({
    spacing: { after: 80, before: 0 },
    indent: { left: 480, hanging: 260 },
    children: [
      new TextRun({ text: '•  ', font: 'Arial', size: 22, color: MID }),
      new TextRun({ text: text, font: 'Arial', size: 22, color: '000000' })
    ]
  });
}

// Oranger Platzhalter-TextRun
function ph(text) {
  return new TextRun({ text: text, font: 'Arial', size: 22, bold: true, color: ORANGE });
}

// Platzhalter-Paragraph (nur orange)
function phP(text, opts) {
  opts = opts || {};
  return new Paragraph({
    spacing: { after: opts.afterPt ? opts.afterPt*20 : 160, before: 0 },
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
    children: [ph(text)]
  });
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

// ── Tabellenhelfer ────────────────────────────────────────────────────────────
var bCol = 'BBCCDD';

function hCell(text, w) {
  return new TableCell({
    width:{size:w,type:WidthType.DXA},
    borders: borders(BLUE),
    shading:{fill:BLUE, type:'clear'},
    margins:{top:80,bottom:80,left:120,right:120},
    children:[new Paragraph({children:[new TextRun({text:text,font:'Arial',size:20,bold:true,color:'FFFFFF'})]})]
  });
}

function lCell(text, w) {
  return new TableCell({
    width:{size:w,type:WidthType.DXA},
    borders:borders(bCol),
    shading:{fill:GREY,type:'clear'},
    margins:{top:60,bottom:60,left:120,right:120},
    children:[new Paragraph({children:[new TextRun({text:text,font:'Arial',size:20,bold:true,color:BLUE})]})]
  });
}

function tCell(children, w, center) {
  return new TableCell({
    width:{size:w,type:WidthType.DXA},
    borders:borders(bCol),
    shading:{fill:'FFFFFF',type:'clear'},
    margins:{top:60,bottom:60,left:120,right:120},
    children:[new Paragraph({
      alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT,
      children: Array.isArray(children) ? children : [new TextRun({text:children,font:'Arial',size:20,color:'333333'})]
    })]
  });
}

function phCell(text, w) {
  return new TableCell({
    width:{size:w,type:WidthType.DXA},
    borders:borders(bCol),
    shading:{fill:'FFF8EC',type:'clear'},
    margins:{top:60,bottom:60,left:120,right:120},
    children:[new Paragraph({alignment:AlignmentType.CENTER, children:[ph(text)]})]
  });
}

// ── Tabelle 1: Stundensätze ───────────────────────────────────────────────────
function stundenTable() {
  var rows = [
    new TableRow({ children: [
      hCell('Leistungsart', 5200),
      hCell('Einheit', 1600),
      hCell('Netto-Preis', 2560),
    ]}),
  ];
  var data = [
    ['Softwareentwicklung & Implementierung (neue Funktionen, Anpassungen)',    'pro Stunde', '[Betrag] €'],
    ['Beratung & Konzeption (Anforderungsanalyse, Planung, Dokumentation)',     'pro Stunde', '[Betrag] €'],
    ['Schulung & Training (Video-Call oder Vor-Ort)',                           'pro Stunde', '[Betrag] €'],
    ['Datenmigration & Datenaufbereitung (Import/Bereinigung bestehender Daten)', 'pro Stunde', '[Betrag] €'],
    ['Notfall-Support außerhalb der Servicezeiten (nach Absprache)',            'pro Stunde', '[Betrag] €'],
  ];
  data.forEach(function(row) {
    rows.push(new TableRow({ children: [
      tCell(row[0], 5200),
      tCell(row[1], 1600, true),
      phCell(row[2], 2560),
    ]}));
  });
  return new Table({ width:{size:9360,type:WidthType.DXA}, columnWidths:[5200,1600,2560], rows:rows });
}

// ── Tabelle 2: Pauschalen ─────────────────────────────────────────────────────
function pauschalTable() {
  var rows = [
    new TableRow({ children: [
      hCell('Leistung', 4800),
      hCell('Beschreibung', 2640),
      hCell('Pauschalpreis', 1920),
    ]}),
  ];
  var data = [
    ['Zusätzliche Einführungsschulung',
     '60 Min. Video-Call für neue Mitarbeiter oder Auffrischung',
     '[Betrag] €'],
    ['Erweiterte Einrichtung / Rebranding',
     'Anpassung Logo, Farben, Firmenbezeichnung nach Ersteinrichtung',
     '[Betrag] €'],
    ['Individuelle Excel-Exportvorlage',
     'Angepasstes Export-Layout mit Ihren Feldern und Spalten',
     '[Betrag] €'],
    ['Datenmigration (bis 500 Datensätze)',
     'Import vorhandener Daten aus Excel, CSV o.ä.',
     '[Betrag] €'],
    ['Datenmigration (501–2.000 Datensätze)',
     'Größerer Import inkl. Bereinigung und Mapping',
     '[Betrag] €'],
    ['Datenbank-Wiederherstellung (Backup-Restore)',
     'Wiederherstellung eines früheren Datenstands auf Kundenwunsch',
     '[Betrag] €'],
    ['Jahres-Reporterstellung (PDF)',
     'Professionelle Auswertung auf Basis der Aktivitätsdaten',
     '[Betrag] €'],
  ];
  data.forEach(function(row) {
    rows.push(new TableRow({ children: [
      tCell(row[0], 4800),
      tCell(row[1], 2640),
      phCell(row[2], 1920),
    ]}));
  });
  return new Table({ width:{size:9360,type:WidthType.DXA}, columnWidths:[4800,2640,1920], rows:rows });
}

// ── Tabelle 3: Paket-Upgrades ────────────────────────────────────────────────
function upgradeTable() {
  var rows = [
    new TableRow({ children: [
      hCell('Upgrade', 3600),
      hCell('Einmalige Wechselgebühr', 2880),
      hCell('Neue Monatspauschale', 2880),
    ]}),
  ];
  var data = [
    ['Starter → Team  (bis 5 → bis 12 Nutzer)',  '[Betrag] €',  '159,00 €'],
    ['Starter → Pro   (bis 5 → bis 20 Nutzer)',  '[Betrag] €',  '229,00 €'],
    ['Team → Pro      (bis 12 → bis 20 Nutzer)', '[Betrag] €',  '229,00 €'],
  ];
  data.forEach(function(row) {
    rows.push(new TableRow({ children: [
      tCell(row[0], 3600),
      phCell(row[1], 2880),
      tCell(row[2], 2880, true),
    ]}));
  });
  return new Table({ width:{size:9360,type:WidthType.DXA}, columnWidths:[3600,2880,2880], rows:rows });
}

// ════════════════════════════════════════════════════════════════════════════
// DOKUMENT
// ════════════════════════════════════════════════════════════════════════════
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
            new TextRun({ text: 'Preisliste Zusatzleistungen  |  Aktions Tracker', font: 'Arial', size: 18, bold: true, color: BLUE }),
            new TextRun({ text: '    Anlage C zum IT-Dienstleistungsvertrag', font: 'Arial', size: 18, color: '888888' }),
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
            new TextRun({ text: '  ·  Anlage C – Preisliste Zusatzleistungen', font: 'Arial', size: 16, color: 'AAAAAA', italics: true }),
          ]
        })
      ]})
    },
    children: [

      // ── Titel ───────────────────────────────────────────────────────────
      new Paragraph({ spacing:{after:80,before:200}, alignment:AlignmentType.CENTER,
        children:[new TextRun({text:'PREISLISTE ZUSATZLEISTUNGEN', font:'Arial', size:44, bold:true, color:BLUE})] }),
      new Paragraph({ spacing:{after:80}, alignment:AlignmentType.CENTER,
        children:[new TextRun({text:'Aktions Tracker', font:'Arial', size:28, bold:true, color:MID})] }),
      new Paragraph({ spacing:{after:80}, alignment:AlignmentType.CENTER,
        children:[new TextRun({text:'Anlage C zum IT-Dienstleistungsvertrag  ·  Gilt für Leistungen außerhalb des vereinbarten Pakets', font:'Arial', size:20, color:'555555', italics:true})] }),
      new Paragraph({ spacing:{after:40}, alignment:AlignmentType.CENTER,
        children:[
          new TextRun({text:'Auftraggeber: ', font:'Arial', size:20, color:'555555'}),
          ph('[Firma des Auftraggebers]'),
          new TextRun({text:'   ·   Datum: ', font:'Arial', size:20, color:'555555'}),
          ph('[DATUM]'),
        ]
      }),
      ...empty(1),
      hr(),
      ...empty(1),

      // ── Hinweisbox ──────────────────────────────────────────────────────
      new Paragraph({
        spacing:{after:160},
        shading:{fill:'EEF4FB',type:'clear'},
        border:{
          top:{style:BorderStyle.SINGLE,size:4,color:MID},
          bottom:{style:BorderStyle.SINGLE,size:4,color:MID},
          left:{style:BorderStyle.SINGLE,size:12,color:MID},
          right:{style:BorderStyle.SINGLE,size:4,color:MID},
        },
        indent:{left:200,right:200},
        children:[
          new TextRun({text:'Hinweis: ', font:'Arial', size:20, bold:true, color:BLUE}),
          new TextRun({text:'Diese Anlage C gilt ergänzend zum IT-Dienstleistungsvertrag und zur Leistungsbeschreibung (Anlage B). Alle hier aufgeführten Leistungen sind optional und werden nur auf ausdrücklichen schriftlichen Auftrag hin erbracht. Orange markierte Felder ', font:'Arial', size:20, color:'333333'}),
          new TextRun({text:'[Platzhalter]', font:'Arial', size:20, bold:true, color:ORANGE}),
          new TextRun({text:' sind vor Unterzeichnung durch die vereinbarten Werte zu ersetzen. Alle Preise sind Nettobeträge gemäß § 19 UStG.', font:'Arial', size:20, color:'333333'}),
        ]
      }),
      ...empty(1),

      // ── § 1 Geltungsbereich ──────────────────────────────────────────────
      sectionHead('§ 1  Geltungsbereich'),
      p('Diese Preisliste regelt die Vergütung von Leistungen, die nicht im vereinbarten Paketumfang gemäß Anlage B (Leistungsbeschreibung) enthalten sind. Sie wird Bestandteil des IT-Dienstleistungsvertrags durch beiderseitige Unterzeichnung oder durch schriftliche Beauftragung einzelner Positionen.'),
      p('Für jede Zusatzleistung ist eine schriftliche Auftragserteilung (E-Mail genügt) mit Angabe der gewünschten Leistung und Kostenschätzung erforderlich. Leistungen über 3 Arbeitsstunden werden vorab durch ein Angebot mit Zeitschätzung abgestimmt.'),
      ...empty(1),

      // ── § 2 Stundensätze ─────────────────────────────────────────────────
      sectionHead('§ 2  Stundensätze (zeitbasierte Leistungen)'),
      p('Zeitbasierte Leistungen werden nach tatsächlichem Aufwand abgerechnet. Die Abrechnung erfolgt in 15-Minuten-Schritten. Vor Beginn wird der voraussichtliche Aufwand mitgeteilt.'),
      ...empty(1),
      stundenTable(),
      ...empty(1),

      // ── § 3 Pauschalangebote ──────────────────────────────────────────────
      sectionHead('§ 3  Pauschalangebote (Festpreisleistungen)'),
      p('Die nachfolgenden Leistungen werden zu einem Festpreis angeboten. Etwaige Mehraufwände durch veränderten Umfang oder Kundenwunsch werden gesondert vereinbart.'),
      ...empty(1),
      pauschalTable(),
      ...empty(1),

      // ── § 4 Paket-Upgrades ────────────────────────────────────────────────
      sectionHead('§ 4  Paket-Upgrades'),
      p('Ein Upgrade auf ein größeres Paket ist jederzeit möglich und wird zum nächsten Monatsbeginn aktiviert. Die einmalige Wechselgebühr deckt Konfigurationsanpassungen ab. Die neue Monatspauschale gilt ab dem Aktivierungsmonat.'),
      ...empty(1),
      upgradeTable(),
      ...empty(1),

      // ── § 5 Auf Anfrage ──────────────────────────────────────────────────
      sectionHead('§ 5  Leistungen auf individuelle Anfrage'),
      p('Die folgenden Leistungen erfordern eine individuelle Aufwandsschätzung und werden auf Anfrage angeboten. Kosten und Zeitrahmen werden vorab schriftlich vereinbart.'),
      bullet('REST-API / Schnittstellen-Entwicklung für Drittsysteme (CRM, ERP, Warenwirtschaft)'),
      bullet('Native Mobile App (iOS App Store / Google Play Store)'),
      bullet('Mehrsprachigkeit der Benutzeroberfläche'),
      bullet('Single Sign-On (SSO) / Active Directory / LDAP-Anbindung'),
      bullet('White-Label-Lizenzierung oder Weiterverkauf der Software an Dritte'),
      bullet('Individuelle Dashboard-Erweiterungen oder Custom-KPIs'),
      bullet('Automatisierte Berichte (E-Mail-Versand von Reports)'),
      bullet('Vor-Ort-Schulung beim Auftraggeber (zzgl. Reisekosten nach Aufwand)'),
      ...empty(1),

      // ── § 6 Auftragserteilung ─────────────────────────────────────────────
      sectionHead('§ 6  Auftragserteilung und Ablauf'),
      p('(1) Jede Zusatzleistung wird durch schriftliche Beauftragung per E-Mail ausgelöst. Der AN bestätigt den Auftrag innerhalb von 2 Werktagen mit Zeitschätzung und ggf. Angebot.'),
      p('(2) Bei Aufträgen über 5 Arbeitsstunden wird zunächst ein schriftliches Angebot erstellt. Die Leistung beginnt nach schriftlicher Annahme durch den AG.'),
      p('(3) Änderungen am Leistungsumfang während der Umsetzung werden gesondert bewertet und nur auf schriftliche Freigabe hin umgesetzt.'),
      ...empty(1),

      // ── § 7 Zahlungsbedingungen ───────────────────────────────────────────
      sectionHead('§ 7  Zahlungsbedingungen'),
      p('(1) Stundensatzleistungen werden monatlich oder nach Projektabschluss abgerechnet. Pauschalangebote werden nach Fertigstellung und Abnahme in Rechnung gestellt.'),
      p('(2) Das Zahlungsziel beträgt 7 Tage ab Rechnungsstellung.'),
      p('(3) Alle Preise sind Nettobeträge gemäß § 19 UStG (Kleinunternehmerregelung). Es wird keine Umsatzsteuer ausgewiesen.'),
      p('(4) Reisekosten (Fahrtkosten, Unterkunft) werden nach tatsächlichem Aufwand separat in Rechnung gestellt und vorab abgestimmt.'),
      ...empty(2),
      hr(),

      // ── Unterschriften ────────────────────────────────────────────────────
      sectionHead('Unterzeichnung'),
      p('Mit ihrer Unterschrift erkennen beide Parteien diese Preisliste als verbindliche Anlage C zum IT-Dienstleistungsvertrag an.'),
      ...empty(1),
      new Table({
        width: { size: 9360, type: WidthType.DXA }, columnWidths: [4320, 720, 4320],
        rows: [new TableRow({ children: [
          (function() {
            return new TableCell({
              width:{size:4320,type:WidthType.DXA}, borders:noBorder(), margins:{top:200,bottom:100,left:0,right:0},
              children:[
                new Paragraph({spacing:{after:800},children:[new TextRun({text:'',font:'Arial',size:22})]}),
                new Paragraph({border:{top:{style:BorderStyle.SINGLE,size:4,color:'888888'}},spacing:{after:80},
                  children:[new TextRun({text:'Ort, Datum  –  Auftraggeber',font:'Arial',size:18,color:'888888',italics:true})]}),
                new Paragraph({spacing:{after:40},children:[new TextRun({text:'Name, Funktion',font:'Arial',size:20,bold:true,color:'333333'})]}),
                new Paragraph({spacing:{after:0},children:[ph('[Firma des Auftraggebers]')]})
              ]
            });
          })(),
          new TableCell({width:{size:720,type:WidthType.DXA},borders:noBorder(),children:[new Paragraph({children:[]})]}),
          (function() {
            return new TableCell({
              width:{size:4320,type:WidthType.DXA}, borders:noBorder(), margins:{top:200,bottom:100,left:0,right:0},
              children:[
                new Paragraph({spacing:{after:800},children:[new TextRun({text:'',font:'Arial',size:22})]}),
                new Paragraph({border:{top:{style:BorderStyle.SINGLE,size:4,color:'888888'}},spacing:{after:80},
                  children:[new TextRun({text:'Ort, Datum  –  Auftragnehmer',font:'Arial',size:18,color:'888888',italics:true})]}),
                new Paragraph({spacing:{after:40},children:[new TextRun({text:'Name, Inhaber',font:'Arial',size:20,bold:true,color:'333333'})]}),
                new Paragraph({spacing:{after:0},children:[new TextRun({text:'[Name oder Firma AN]',font:'Arial',size:18,color:ORANGE,bold:true})]})
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
