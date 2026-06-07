const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  PageNumber, Header, Footer, TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

// ── Farben & Schrift ──────────────────────────────────────────────────────────
const ACCENT   = '1F4E79';
const BG_HEAD  = '1F4E79';
const BG_LIGHT = 'DEEAF1';
const BG_WHITE = 'FFFFFF';
const BG_GELB  = 'FFF2CC';
const BG_GRUEN = 'E2EFDA';
const FONT     = 'Calibri';

// ── Hilfsfunktionen ───────────────────────────────────────────────────────────
const border = (color = 'BFBFBF') => ({ style: BorderStyle.SINGLE, size: 4, color });
const cellBorders = (c = 'BFBFBF') => ({ top: border(c), bottom: border(c), left: border(c), right: border(c) });

function headCell(text, width) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: cellBorders(ACCENT),
    shading: { fill: BG_HEAD, type: ShadingType.CLEAR },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      children: [new TextRun({ text, font: FONT, size: 18, bold: true, color: BG_WHITE })]
    })]
  });
}

function dataCell(text, width, opts = {}) {
  const { bold = false, bg = BG_WHITE, color = '000000', center = false, italics = false } = opts;
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: cellBorders(),
    shading: { fill: bg, type: ShadingType.CLEAR },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 70, bottom: 70, left: 120, right: 120 },
    children: [new Paragraph({
      alignment: center ? AlignmentType.CENTER : AlignmentType.LEFT,
      children: [new TextRun({ text, font: FONT, size: 18, bold, color, italics })]
    })]
  });
}

function multiLineCell(lines, width, opts = {}) {
  const { bg = BG_WHITE } = opts;
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: cellBorders(),
    shading: { fill: bg, type: ShadingType.CLEAR },
    verticalAlign: VerticalAlign.TOP,
    margins: { top: 70, bottom: 70, left: 120, right: 120 },
    children: lines.map((l, i) => new Paragraph({
      spacing: i < lines.length - 1 ? { after: 40 } : {},
      children: [new TextRun({ text: l, font: FONT, size: 18 })]
    }))
  });
}

function para(text, opts = {}) {
  const { bold = false, size = 20, spBefore = 80, spAfter = 80, color = '000000', italics = false } = opts;
  return new Paragraph({
    spacing: { before: spBefore, after: spAfter },
    children: [new TextRun({ text, font: FONT, size, bold, color, italics })]
  });
}

function spacer(pt = 80) {
  return new Paragraph({ spacing: { before: pt, after: pt }, children: [new TextRun('')] });
}

function sectionTitle(text) {
  return new Paragraph({
    spacing: { before: 280, after: 100 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 4 } },
    children: [new TextRun({ text, font: FONT, size: 22, bold: true, color: ACCENT })]
  });
}

function infoBox(lines, bg = BG_LIGHT) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({ children: [
      new TableCell({
        width: { size: 9360, type: WidthType.DXA },
        borders: cellBorders(ACCENT),
        shading: { fill: bg, type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 200, right: 200 },
        children: lines.map(l => new Paragraph({
          spacing: { after: 40 },
          children: [new TextRun({ text: l, font: FONT, size: 18, italics: true, color: '2F5496' })]
        }))
      })
    ]})]
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
//  § 2: Modul-Änderungstabelle
// ═══════════════════════════════════════════════════════════════════════════════
// Spalten: Modul | Beschreibung | Monatspreis | Jahrespreis | Änderung
const MW = [1700, 3160, 1300, 1300, 1900];  // = 9360

const tabelleModule = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: MW,
  rows: [
    new TableRow({ children: [
      headCell('Modul', MW[0]),
      headCell('Beschreibung', MW[1]),
      headCell('+Preis/Monat', MW[2]),
      headCell('+Preis/Jahr', MW[3]),
      headCell('Änderung', MW[4]),
    ]}),

    // Gebiets-Karte
    new TableRow({ children: [
      dataCell('Gebiets-Karte', MW[0], { bold: true }),
      multiLineCell([
        'Interaktive Stationskarte',
        'Gebiets-Zuordnung & -verwaltung',
        'Geocodierung & Benachrichtigung',
      ], MW[1]),
      dataCell('+39 €', MW[2], { center: true }),
      dataCell('+390 €', MW[3], { center: true }),
      new TableCell({
        width: { size: MW[4], type: WidthType.DXA },
        borders: cellBorders(),
        shading: { fill: BG_WHITE, type: ShadingType.CLEAR },
        margins: { top: 70, bottom: 70, left: 120, right: 120 },
        children: [
          new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: '☐  wird hinzugefügt', font: FONT, size: 17 })] }),
          new Paragraph({ children: [new TextRun({ text: '☐  wird entfernt', font: FONT, size: 17 })] }),
        ]
      }),
    ]}),

    // Aktivitäten-Heatmap
    new TableRow({ children: [
      new TableCell({
        width: { size: MW[0], type: WidthType.DXA },
        borders: cellBorders(),
        shading: { fill: BG_WHITE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [
          new Paragraph({ children: [new TextRun({ text: 'Aktivitäten-Heatmap', font: FONT, size: 18, bold: true })] }),
          new Paragraph({ children: [new TextRun({ text: '(+ Gebiets-Karte)', font: FONT, size: 15, italics: true, color: '666666' })] }),
        ]
      }),
      multiLineCell([
        'Heatmap-Overlay (Aktivitäten/Station)',
        'Jahresfilter & historischer Vergleich',
        'Kombinierte Ansicht',
      ], MW[1]),
      dataCell('+29 €', MW[2], { center: true }),
      dataCell('+290 €', MW[3], { center: true }),
      new TableCell({
        width: { size: MW[4], type: WidthType.DXA },
        borders: cellBorders(),
        shading: { fill: BG_WHITE, type: ShadingType.CLEAR },
        margins: { top: 70, bottom: 70, left: 120, right: 120 },
        children: [
          new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: '☐  wird hinzugefügt', font: FONT, size: 17 })] }),
          new Paragraph({ children: [new TextRun({ text: '☐  wird entfernt', font: FONT, size: 17 })] }),
        ]
      }),
    ]}),

    // Freifeld
    new TableRow({ children: [
      dataCell('[Sonstiges]', MW[0], { italics: true, color: '888888' }),
      dataCell('[Beschreibung]', MW[1], { italics: true, color: '888888' }),
      dataCell('[+X €]', MW[2], { center: true, color: '888888', italics: true }),
      dataCell('[+X €]', MW[3], { center: true, color: '888888', italics: true }),
      new TableCell({
        width: { size: MW[4], type: WidthType.DXA },
        borders: cellBorders(),
        shading: { fill: BG_WHITE, type: ShadingType.CLEAR },
        margins: { top: 70, bottom: 70, left: 120, right: 120 },
        children: [
          new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: '☐  wird hinzugefügt', font: FONT, size: 17 })] }),
          new Paragraph({ children: [new TextRun({ text: '☐  wird entfernt', font: FONT, size: 17 })] }),
        ]
      }),
    ]}),
  ]
});

// ═══════════════════════════════════════════════════════════════════════════════
//  § 3: Vergütungstabelle (monatlich + jährlich)
// ═══════════════════════════════════════════════════════════════════════════════
const tabelleVerguetung = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [4560, 2400, 2400],
  rows: [
    new TableRow({ children: [
      headCell('Position', 4560),
      headCell('Monatlich', 2400),
      headCell('Jährlich (Voraus)', 2400),
    ]}),
    new TableRow({ children: [
      dataCell('Bisherige monatliche Pauschale', 4560, { bg: BG_LIGHT }),
      dataCell('[BISHERIGER BETRAG] €/Mon.', 2400, { bg: BG_LIGHT }),
      dataCell('[BISHERIGER BETRAG × 12] €/Jahr', 2400, { bg: BG_LIGHT }),
    ]}),
    new TableRow({ children: [
      dataCell('Aufpreis für hinzugefügte Module', 4560),
      dataCell('+ [AUFPREIS] €/Mon.', 2400),
      dataCell('+ [AUFPREIS_JAHR] €/Jahr', 2400),
    ]}),
    new TableRow({ children: [
      dataCell('Abzug für entfernte Module', 4560),
      dataCell('– [ABZUG] €/Mon.', 2400),
      dataCell('– [ABZUG_JAHR] €/Jahr', 2400),
    ]}),
    new TableRow({ children: [
      new TableCell({
        width: { size: 4560, type: WidthType.DXA },
        borders: cellBorders(ACCENT),
        shading: { fill: BG_LIGHT, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 120, right: 120 },
        children: [new Paragraph({
          children: [new TextRun({ text: 'Neue Gesamtpauschale (ab Inkrafttreten)', font: FONT, size: 18, bold: true })]
        })]
      }),
      dataCell('[NEUER BETRAG] €/Mon.', 2400, { bold: true, bg: BG_LIGHT }),
      dataCell('[NEUER BETRAG_JAHR] €/Jahr', 2400, { bold: true, bg: BG_LIGHT }),
    ]}),
  ]
});

// ── Zahlungsrhythmus-Box ──────────────────────────────────────────────────────
const zahlungsrhythmus = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [9360],
  rows: [new TableRow({ children: [
    new TableCell({
      width: { size: 9360, type: WidthType.DXA },
      borders: cellBorders(ACCENT),
      shading: { fill: BG_GELB, type: ShadingType.CLEAR },
      margins: { top: 140, bottom: 140, left: 200, right: 200 },
      children: [
        new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text: 'Zahlungsrhythmus für die geänderten Zusatzmodule:', font: FONT, size: 18, bold: true })] }),
        new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: '☐  Monatlich  – Abrechnung ab dem Inkrafttretensdatum mit der monatlichen Gesamtpauschale', font: FONT, size: 18 })] }),
        new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text: '☐  Jährlich im Voraus  – Der Jahresaufpreis für das hinzugefügte Modul ist bei Unterzeichnung fällig', font: FONT, size: 18 })] }),
        new Paragraph({
          children: [
            new TextRun({ text: 'Hinweis: ', font: FONT, size: 17, bold: true, italics: true }),
            new TextRun({ text: 'Beim Jahrestarif werden 2 Monate gespart (10 statt 12 Monatsbeiträge). Der Jahresbetrag ist als Einmalzahlung im Voraus fällig. Eine vorzeitige Stornierung des Moduls berechtigt nicht zur anteiligen Rückerstattung.', font: FONT, size: 17, italics: true, color: '555555' }),
          ]
        }),
      ]
    })
  ]})]
});

// ── Unterschriften ────────────────────────────────────────────────────────────
const unterschriften = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [4580, 200, 4580],
  rows: [new TableRow({ children: [
    new TableCell({
      width: { size: 4580, type: WidthType.DXA },
      borders: cellBorders(),
      shading: { fill: BG_WHITE, type: ShadingType.CLEAR },
      margins: { top: 120, bottom: 120, left: 140, right: 140 },
      children: [
        new Paragraph({ spacing: { after: 400 }, children: [new TextRun({ text: 'Auftraggeber (AG)', font: FONT, size: 18, bold: true })] }),
        new Paragraph({ spacing: { after: 600 }, children: [new TextRun({ text: 'Ort, Datum: __________________________', font: FONT, size: 18 })] }),
        new Paragraph({ border: { top: { style: BorderStyle.SINGLE, size: 4, color: '000000' } }, spacing: { before: 20, after: 60 }, children: [new TextRun('')] }),
        new Paragraph({ children: [new TextRun({ text: 'Unterschrift, Stempel', font: FONT, size: 16, italics: true, color: '666666' })] }),
        spacer(40),
        new Paragraph({ children: [new TextRun({ text: '[[Firma des Auftraggebers]]', font: FONT, size: 18 })] }),
      ]
    }),
    new TableCell({
      width: { size: 200, type: WidthType.DXA },
      borders: { top: border('FFFFFF'), bottom: border('FFFFFF'), left: border('FFFFFF'), right: border('FFFFFF') },
      children: [new Paragraph({ children: [new TextRun('')] })]
    }),
    new TableCell({
      width: { size: 4580, type: WidthType.DXA },
      borders: cellBorders(),
      shading: { fill: BG_WHITE, type: ShadingType.CLEAR },
      margins: { top: 120, bottom: 120, left: 140, right: 140 },
      children: [
        new Paragraph({ spacing: { after: 400 }, children: [new TextRun({ text: 'Auftragnehmer (AN)', font: FONT, size: 18, bold: true })] }),
        new Paragraph({ spacing: { after: 600 }, children: [new TextRun({ text: 'Ort, Datum: __________________________', font: FONT, size: 18 })] }),
        new Paragraph({ border: { top: { style: BorderStyle.SINGLE, size: 4, color: '000000' } }, spacing: { before: 20, after: 60 }, children: [new TextRun('')] }),
        new Paragraph({ children: [new TextRun({ text: 'Unterschrift', font: FONT, size: 16, italics: true, color: '666666' })] }),
        spacer(40),
        new Paragraph({ children: [new TextRun({ text: '[Name oder Firma AN]', font: FONT, size: 18 })] }),
      ]
    }),
  ]})]
});

// ═══════════════════════════════════════════════════════════════════════════════
//  Dokument
// ═══════════════════════════════════════════════════════════════════════════════
const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 20 } } } },
  sections: [{
    properties: {
      page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } }
    },
    headers: {
      default: new Header({ children: [new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: ACCENT, space: 4 } },
        spacing: { after: 120 },
        tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
        children: [
          new TextRun({ text: 'Aktions Tracker – Vertragsergänzung', font: FONT, size: 16, bold: true, color: ACCENT }),
          new TextRun({ text: '\t Nachtrag zum IT-Dienstleistungsvertrag', font: FONT, size: 16, color: '666666' }),
        ]
      })] })
    },
    footers: {
      default: new Footer({ children: [new Paragraph({
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: ACCENT, space: 4 } },
        spacing: { before: 120 },
        tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
        children: [
          new TextRun({ text: 'Vertraulich – nur für Vertragsparteien', font: FONT, size: 16, color: '666666' }),
          new TextRun({ text: '\tSeite ', font: FONT, size: 16, color: '666666' }),
          new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: '666666' }),
        ]
      })] })
    },
    children: [
      // Titel
      new Paragraph({ spacing: { before: 0, after: 40 }, children: [new TextRun({ text: 'Nachtrag Nr. [X]', font: FONT, size: 32, bold: true, color: ACCENT })] }),
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [new TextRun({ text: 'Vertragsergänzung / Änderungsvereinbarung', font: FONT, size: 24, bold: true })] }),
      new Paragraph({ spacing: { before: 0, after: 220 }, children: [new TextRun({ text: 'zum IT-Dienstleistungsvertrag Aktions Tracker', font: FONT, size: 20, italics: true, color: '444444' })] }),

      // Bezugstabelle
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2000, 7360],
        rows: [
          new TableRow({ children: [dataCell('Auftraggeber:', 2000, { bold: true, bg: BG_LIGHT }), dataCell('[[Firma des Auftraggebers]], [Anschrift]', 7360, { bg: BG_LIGHT })] }),
          new TableRow({ children: [dataCell('Auftragnehmer:', 2000, { bold: true, bg: BG_LIGHT }), dataCell('[Name oder Firma AN], [Anschrift]', 7360, { bg: BG_LIGHT })] }),
          new TableRow({ children: [dataCell('Hauptvertrag:', 2000, { bold: true, bg: BG_LIGHT }), dataCell('IT-Dienstleistungsvertrag Aktions Tracker vom [DATUM DES HAUPTVERTRAGS]', 7360, { bg: BG_LIGHT })] }),
          new TableRow({ children: [dataCell('Nachtragsdatum:', 2000, { bold: true, bg: BG_LIGHT }), dataCell('[DATUM DIESES NACHTRAGS]', 7360, { bg: BG_LIGHT })] }),
        ]
      }),

      spacer(160),
      infoBox([
        'Dieser Nachtrag ergänzt und ändert den oben genannten Hauptvertrag. Alle nicht ausdrücklich geänderten Bestimmungen',
        'des Hauptvertrags (inkl. bisheriger Anlagen und Nachträge) bleiben unverändert in Kraft.',
      ]),
      spacer(100),

      // § 1
      sectionTitle('§ 1  Gegenstand dieses Nachtrags'),
      para('Die Parteien vereinbaren einvernehmlich folgende Änderungen und Ergänzungen zum oben genannten Vertrag:'),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [9360],
        rows: [new TableRow({ children: [
          new TableCell({
            width: { size: 9360, type: WidthType.DXA },
            borders: cellBorders(),
            shading: { fill: BG_WHITE, type: ShadingType.CLEAR },
            margins: { top: 120, bottom: 280, left: 160, right: 160 },
            children: [
              new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: 'Kurzbeschreibung der Änderung:', font: FONT, size: 18, bold: true })] }),
              new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: '[z. B. Hinzubuchung Modul „Gebiets-Karte" ab 01.07.2026, monatliche Abrechnung]', font: FONT, size: 18, italics: true, color: '888888' })] }),
              new Paragraph({ children: [new TextRun({ text: '', font: FONT, size: 18 })] }),
            ]
          })
        ]})]
      }),
      spacer(160),

      // § 2
      sectionTitle('§ 2  Geänderte / neue Leistungsmodule'),
      para('Folgende Module werden dem Vertragsumfang hinzugefügt oder daraus entfernt (zutreffendes ankreuzen). Alle Preise netto; AN ist Kleinunternehmer gem. § 19 UStG.'),
      spacer(80),
      tabelleModule,
      spacer(160),

      // § 3
      sectionTitle('§ 3  Angepasste Vergütung'),
      para('Durch die in § 2 vereinbarten Änderungen ergibt sich folgende neue Vergütungsstruktur:'),
      spacer(80),
      tabelleVerguetung,
      spacer(120),
      zahlungsrhythmus,
      spacer(80),
      para('Für die Basismonatspauschale (ohne Zusatzmodule) gilt unverändert der im Hauptvertrag vereinbarte Zahlungsrhythmus.', { size: 18, italics: true, color: '444444', spBefore: 0 }),
      spacer(160),

      // § 4
      sectionTitle('§ 4  Inkrafttreten'),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [9360],
        rows: [new TableRow({ children: [
          new TableCell({
            width: { size: 9360, type: WidthType.DXA },
            borders: cellBorders(),
            shading: { fill: BG_WHITE, type: ShadingType.CLEAR },
            margins: { top: 120, bottom: 120, left: 160, right: 160 },
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: 'Die Änderungen treten zum  ', font: FONT, size: 18 }),
                  new TextRun({ text: '[INKRAFTTRETEN-DATUM]', font: FONT, size: 18, bold: true, color: ACCENT }),
                  new TextRun({ text: '  in Kraft.', font: FONT, size: 18 }),
                ]
              }),
              spacer(60),
              para('☐  Sofort mit Unterzeichnung durch beide Parteien', { size: 18, spBefore: 40, spAfter: 20 }),
              para('☐  Zum 1. des auf die Unterzeichnung folgenden Kalendermonats', { size: 18, spBefore: 20, spAfter: 20 }),
              para('☐  Abweichendes Datum (s. o.)', { size: 18, spBefore: 20, spAfter: 40 }),
            ]
          })
        ]})]
      }),
      spacer(160),

      // § 5
      sectionTitle('§ 5  Schlussbestimmungen'),
      para('(1)  Alle übrigen Bestimmungen des Hauptvertrags und aller bisherigen Anlagen und Nachträge bleiben unberührt und gelten unverändert fort.'),
      para('(2)  Mündliche Nebenabreden bestehen nicht. Änderungen und Ergänzungen dieses Nachtrags bedürfen der Schriftform (E-Mail genügt).'),
      para('(3)  Im Falle von Widersprüchen zwischen diesem Nachtrag und dem Hauptvertrag hat der Nachtrag Vorrang.'),
      spacer(200),

      // Unterschriften
      sectionTitle('Unterzeichnung'),
      para('Mit ihrer Unterschrift erklären beide Parteien ihr Einverständnis mit den vorstehenden Änderungen, den vereinbarten Preisen und dem gewählten Zahlungsrhythmus.', { spAfter: 120 }),
      unterschriften,
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('Nachtrag_Vorlage_Aktions_Tracker.docx', buf);
  console.log('Nachtragsvorlage erstellt.');
});
