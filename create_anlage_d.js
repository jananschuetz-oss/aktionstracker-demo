const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  PageNumber, Header, Footer, TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

// ── Farben & Schrift ───────────────────────────────────────────────────────────
const ACCENT   = '1F4E79';
const BG_HEAD  = '1F4E79';
const BG_LIGHT = 'DEEAF1';
const BG_WHITE = 'FFFFFF';
const BG_GRUEN = 'E2EFDA';
const BG_GELB  = 'FFF2CC';
const FONT     = 'Calibri';

// ── Hilfsfunktionen ──────────────────────────────────────────────────────────
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
      children: typeof l === 'string'
        ? [new TextRun({ text: l, font: FONT, size: 18 })]
        : l
    }))
  });
}

function preisCell(monat, jahr, width) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: cellBorders(),
    shading: { fill: BG_WHITE, type: ShadingType.CLEAR },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 70, bottom: 70, left: 120, right: 120 },
    children: [
      new Paragraph({
        spacing: { after: 30 },
        children: [new TextRun({ text: monat, font: FONT, size: 18, bold: true })]
      }),
      new Paragraph({
        children: [new TextRun({ text: 'oder ' + jahr, font: FONT, size: 17, italics: true, color: '2F5496' })]
      }),
      new Paragraph({
        children: [new TextRun({ text: '(2 Monate gespart)', font: FONT, size: 15, color: '375623', italics: true })]
      }),
    ]
  });
}

function buchungCell(width) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: cellBorders(),
    shading: { fill: BG_WHITE, type: ShadingType.CLEAR },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 70, bottom: 70, left: 120, right: 120 },
    children: [
      new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: '☐  Nicht gebucht', font: FONT, size: 17 })] }),
      new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: '☐  Monatlich', font: FONT, size: 17 })] }),
      new Paragraph({ children: [new TextRun({ text: '☐  Jährlich', font: FONT, size: 17 })] }),
    ]
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
    spacing: { before: 240, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 4 } },
    children: [new TextRun({ text, font: FONT, size: 24, bold: true, color: ACCENT })]
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

// ─── Spaltenbreiten ────────────────────────────────────────────────────────────
//  Modul | Funktionen | Aufpreis | Buchung
const W = [1900, 3560, 2300, 1600];  // Summe = 9360

// ═══════════════════════════════════════════════════════════════════════════════
//  Tabelle 1: Grundmodul (immer enthalten)
// ═══════════════════════════════════════════════════════════════════════════════
const tabelleGrund = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [2400, 5600, 1360],
  rows: [
    new TableRow({ children: [
      headCell('Modul', 2400),
      headCell('Enthaltene Funktionen', 5600),
      headCell('Status', 1360),
    ]}),
    new TableRow({ children: [
      new TableCell({
        width: { size: 2400, type: WidthType.DXA },
        borders: cellBorders(),
        shading: { fill: BG_WHITE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [
          new Paragraph({ children: [new TextRun({ text: 'Basis', font: FONT, size: 18, bold: true })] }),
          new Paragraph({ children: [new TextRun({ text: '(immer enthalten)', font: FONT, size: 16, italics: true, color: '666666' })] }),
        ]
      }),
      multiLineCell([
        '▸  Dashboard & KPI-Auswertung (Wochen-/Jahresansicht)',
        '▸  Aktivitäten-Erfassung (Besuche, Platzierungen, Aktionen)',
        '▸  Foto-Dokumentation mit Galerie',
        '▸  Zielplanung & Tracking (Display / Kisten)',
        '▸  Rollen- und Rechteverwaltung (Rep / VKL / Admin)',
        '▸  Excel-Export aller Vertriebsdaten',
        '▸  Progressive Web App (PWA) für mobile Nutzung',
      ], 5600),
      dataCell('✓  Inklusive', 1360, { bold: true, bg: BG_GRUEN, color: '375623', center: true }),
    ]}),
  ]
});

// ═══════════════════════════════════════════════════════════════════════════════
//  Tabelle 2: Optionale Module (mit Preisen + Buchungsauswahl)
// ═══════════════════════════════════════════════════════════════════════════════
const tabelleOptional = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: W,
  rows: [
    new TableRow({ children: [
      headCell('Modul', W[0]),
      headCell('Enthaltene Funktionen', W[1]),
      headCell('Aufpreis zur Monatspauschale', W[2]),
      headCell('Buchung', W[3]),
    ]}),

    // ── Gebiets-Karte ─────────────────────────────────────────────────────────
    new TableRow({ children: [
      new TableCell({
        width: { size: W[0], type: WidthType.DXA },
        borders: cellBorders(),
        shading: { fill: BG_WHITE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [
          new Paragraph({ children: [new TextRun({ text: 'Gebiets-Karte', font: FONT, size: 18, bold: true })] }),
        ]
      }),
      multiLineCell([
        '▸  Interaktive Stationskarte (OpenStreetMap)',
        '▸  Gebiets-Zuordnung (Mitarbeiter ↔ Station)',
        '▸  Verwaltung der Außendienst-Gebiete',
        '▸  Benachrichtigung bei Gebietsänderungen',
        '▸  Geocodierung aller Verkaufsstellen',
      ], W[1]),
      preisCell('+39 €/Monat', '+390 €/Jahr', W[2]),
      buchungCell(W[3]),
    ]}),

    // ── Aktivitäten-Heatmap ───────────────────────────────────────────────────
    new TableRow({ children: [
      new TableCell({
        width: { size: W[0], type: WidthType.DXA },
        borders: cellBorders(),
        shading: { fill: BG_WHITE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [
          new Paragraph({ children: [new TextRun({ text: 'Aktivitäten-Heatmap', font: FONT, size: 18, bold: true })] }),
          new Paragraph({ children: [new TextRun({ text: '(erfordert Gebiets-Karte)', font: FONT, size: 15, italics: true, color: '666666' })] }),
        ]
      }),
      multiLineCell([
        '▸  Heatmap-Overlay auf der Stationskarte',
        '▸  Aktivitätsintensität pro Station visualisiert',
        '▸  Jahresfilter & historischer Vergleich',
        '▸  Kombinierte Ansicht mit Gebiets-Zuordnung',
      ], W[1]),
      preisCell('+29 €/Monat', '+290 €/Jahr', W[2]),
      buchungCell(W[3]),
    ]}),
  ]
});

// ═══════════════════════════════════════════════════════════════════════════════
//  Summentabelle
// ═══════════════════════════════════════════════════════════════════════════════
const tabelleSumme = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [5460, 2300, 1600],
  rows: [
    new TableRow({ children: [
      new TableCell({
        width: { size: 5460, type: WidthType.DXA },
        borders: cellBorders(ACCENT),
        shading: { fill: BG_LIGHT, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 120, right: 120 },
        children: [new Paragraph({
          children: [new TextRun({ text: 'Monatliche Gesamtpauschale (Basis + gebuchte Zusatzmodule)', font: FONT, size: 18, bold: true })]
        })]
      }),
      new TableCell({
        width: { size: 2300, type: WidthType.DXA },
        borders: cellBorders(ACCENT),
        shading: { fill: BG_LIGHT, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 120, right: 120 },
        children: [
          new Paragraph({ spacing: { after: 30 }, children: [new TextRun({ text: '[GESAMT] €/Monat', font: FONT, size: 18, bold: true })] }),
          new Paragraph({ children: [new TextRun({ text: 'oder [GESAMT_JAHR] €/Jahr', font: FONT, size: 16, italics: true, color: '2F5496' })] }),
        ]
      }),
      dataCell('', 1600, { bg: BG_LIGHT }),
    ]})
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
        new Paragraph({
          spacing: { after: 100 },
          children: [new TextRun({ text: 'Zahlungsrhythmus für Zusatzmodule:', font: FONT, size: 18, bold: true })]
        }),
        new Paragraph({
          spacing: { after: 40 },
          children: [new TextRun({ text: '☐  Monatlich  – Abrechnung mit der regulären Monatspauschale', font: FONT, size: 18 })]
        }),
        new Paragraph({
          spacing: { after: 100 },
          children: [new TextRun({ text: '☐  Jährlich im Voraus  – Jahresbetrag bei Buchung fällig, entspricht ca. 1 Freimo­nat', font: FONT, size: 18 })]
        }),
        new Paragraph({
          children: [
            new TextRun({ text: 'Hinweis: ', font: FONT, size: 17, bold: true, italics: true }),
            new TextRun({ text: 'Beim Jahrestarif werden 2 Monate gespart (10 statt 12 Monatsbeiträge). Der Jahresbetrag ist im Voraus fällig. Eine vorzeitige Kündigung des Moduls berechtigt nicht zur anteiligen Rückerstattung.', font: FONT, size: 17, italics: true, color: '555555' }),
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
          new TextRun({ text: 'Aktions Tracker – Leistungsmodule', font: FONT, size: 16, bold: true, color: ACCENT }),
          new TextRun({ text: '\t Anlage D zum IT-Dienstleistungsvertrag', font: FONT, size: 16, color: '666666' }),
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
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [new TextRun({ text: 'Anlage D', font: FONT, size: 32, bold: true, color: ACCENT })] }),
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [new TextRun({ text: 'Leistungsmodule & optionale Erweiterungen', font: FONT, size: 26, bold: true })] }),
      new Paragraph({ spacing: { before: 0, after: 200 }, children: [new TextRun({ text: 'zum IT-Dienstleistungsvertrag Aktions Tracker', font: FONT, size: 20, italics: true, color: '444444' })] }),

      // Vertragsreferenz
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [1800, 7560],
        rows: [
          new TableRow({ children: [dataCell('Auftraggeber:', 1800, { bold: true, bg: BG_LIGHT }), dataCell('[[Firma des Auftraggebers]]', 7560, { bg: BG_LIGHT })] }),
          new TableRow({ children: [dataCell('Auftragnehmer:', 1800, { bold: true, bg: BG_LIGHT }), dataCell('[Name oder Firma AN]', 7560, { bg: BG_LIGHT })] }),
          new TableRow({ children: [dataCell('Hauptvertrag vom:', 1800, { bold: true, bg: BG_LIGHT }), dataCell('[DATUM DES HAUPTVERTRAGS]', 7560, { bg: BG_LIGHT })] }),
        ]
      }),

      spacer(160),
      infoBox([
        'Diese Anlage D ist Bestandteil des oben genannten IT-Dienstleistungsvertrags. Die konkret gebuchten Module sowie der',
        'Zahlungsrhythmus werden bei Vertragsschluss durch Ankreuzen markiert. Änderungen erfordern eine schriftliche Vereinbarung.',
      ]),
      spacer(160),

      // Abschnitt 1: Grundmodul
      sectionTitle('1.  Grundmodul (immer im Vertragsumfang enthalten)'),
      tabelleGrund,
      spacer(200),

      // Abschnitt 2: Optionale Module
      sectionTitle('2.  Optionale Zusatzmodule'),
      para('Die folgenden Module können bei Vertragsschluss oder per Nachtrag gebucht werden. Nicht angekreuzte Module sind nicht Bestandteil des Vertrags. Alle Preise netto (AN ist Kleinunternehmer gem. § 19 UStG).', { size: 18, spBefore: 40, spAfter: 80 }),
      tabelleOptional,
      spacer(120),
      tabelleSumme,
      spacer(140),
      zahlungsrhythmus,
      spacer(200),

      // Abschnitt 3: Hinweise
      sectionTitle('3.  Hinweise & Bedingungen'),
      new Paragraph({ spacing: { before: 80, after: 40 }, children: [new TextRun({ text: 'Moduländerungen während der Vertragslaufzeit', font: FONT, size: 18, bold: true })] }),
      para('Modulbuchungen und -stornierungen sind per schriftlichem Nachtrag zu vereinbaren. Änderungen werden zum Ersten des Folgemonats wirksam, sofern nichts anderes vereinbart wird.', { size: 18, spBefore: 0, spAfter: 80 }),
      new Paragraph({ spacing: { before: 0, after: 40 }, children: [new TextRun({ text: 'Technische Voraussetzungen', font: FONT, size: 18, bold: true })] }),
      para('Das Modul „Aktivitäten-Heatmap" setzt das Modul „Gebiets-Karte" voraus und kann nicht isoliert gebucht werden.', { size: 18, spBefore: 0, spAfter: 80 }),
      new Paragraph({ spacing: { before: 0, after: 40 }, children: [new TextRun({ text: 'Preisanpassungen', font: FONT, size: 18, bold: true })] }),
      para('Preisanpassungen für Zusatzmodule werden mit einer Ankündigungsfrist von 8 Wochen mitgeteilt und gelten als genehmigt, wenn der AG nicht innerhalb von 4 Wochen widerspricht.', { size: 18, spBefore: 0, spAfter: 0 }),
      spacer(200),

      // Unterschriften
      sectionTitle('4.  Unterzeichnung'),
      para('Mit ihrer Unterschrift erklären beide Parteien, die gebuchten Module, den Zahlungsrhythmus und die Preise zur Kenntnis genommen und akzeptiert zu haben.', { size: 18, spAfter: 120 }),
      unterschriften,
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('Anlage_D_Leistungsmodule_Blanko.docx', buf);
  console.log('Anlage D erstellt.');
});
