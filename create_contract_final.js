const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageNumber, HeadingLevel
} = require('docx');
const fs = require('fs');

// ── colour palette ──────────────────────────────────────────────────────────
const BLUE   = '1A3A5C';
const ORANGE = 'C8860A';
const GREY   = 'F2F6FA';
const MID    = '4A6FA5';

// ── helper: thin border object ───────────────────────────────────────────────
function border(col) {
  return { style: BorderStyle.SINGLE, size: 4, color: col || 'BBCCDD' };
}
function borders(col) {
  const b = border(col);
  return { top: b, bottom: b, left: b, right: b };
}
function noBorder() {
  const n = { style: BorderStyle.NIL, size: 0, color: 'FFFFFF' };
  return { top: n, bottom: n, left: n, right: n };
}

// ── helper: plain paragraph ──────────────────────────────────────────────────
function p(text, opts) {
  opts = opts || {};
  return new Paragraph({
    spacing: { after: opts.afterPt ? opts.afterPt * 20 : 160, before: opts.beforePt ? opts.beforePt * 20 : 0 },
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.JUSTIFIED,
    children: [new TextRun({
      text: text,
      font: 'Arial',
      size: opts.size || 22,
      bold: opts.bold || false,
      color: opts.color || '000000',
      italics: opts.italic || false,
    })]
  });
}

// ── helper: mixed-run paragraph ─────────────────────────────────────────────
function pRuns(runs, opts) {
  opts = opts || {};
  return new Paragraph({
    spacing: { after: opts.afterPt ? opts.afterPt * 20 : 160, before: opts.beforePt ? opts.beforePt * 20 : 0 },
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.JUSTIFIED,
    children: runs.map(function(r) {
      return new TextRun({
        text: r.text,
        font: 'Arial',
        size: r.size || opts.size || 22,
        bold: r.bold || false,
        color: r.color || '000000',
        italics: r.italic || false,
      });
    })
  });
}

// ── helper: section heading (§ N Titel) ─────────────────────────────────────
function sectionHead(text) {
  return new Paragraph({
    spacing: { before: 320, after: 120 },
    children: [new TextRun({ text: text, font: 'Arial', size: 24, bold: true, color: BLUE })]
  });
}

// ── helper: bullet paragraph ────────────────────────────────────────────────
function bullet(text, opts) {
  opts = opts || {};
  return new Paragraph({
    spacing: { after: 100, before: 0 },
    indent: { left: 480, hanging: 260 },
    children: [
      new TextRun({ text: '•  ', font: 'Arial', size: 22, color: MID }),
      new TextRun({ text: text, font: 'Arial', size: 22, bold: opts.bold || false, color: opts.color || '000000' })
    ]
  });
}

// ── helper: placeholder run ─────────────────────────────────────────────────
function ph(label) {
  return new TextRun({ text: '[' + label + ']', font: 'Arial', size: 22, bold: true, color: ORANGE });
}

// ── helper: label+value pair in table cell ──────────────────────────────────
function labelVal(label, value, isPlaceholder) {
  return new Paragraph({
    spacing: { after: 100, before: 0 },
    children: [
      new TextRun({ text: label + ': ', font: 'Arial', size: 20, bold: true, color: BLUE }),
      isPlaceholder
        ? new TextRun({ text: '[' + value + ']', font: 'Arial', size: 20, bold: true, color: ORANGE })
        : new TextRun({ text: value, font: 'Arial', size: 20, color: '333333' })
    ]
  });
}

// ── helper: two-column table (for Vertragsparteien) ─────────────────────────
function twoColTable(leftChildren, rightChildren) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [4500, 4500],
    borders: { ...noBorder(), insideH: border('DDDDDD'), insideV: border('DDDDDD') },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 4500, type: WidthType.DXA },
            shading: { fill: GREY, type: ShadingType.CLEAR },
            margins: { top: 200, bottom: 200, left: 240, right: 200 },
            borders: borders('BBCCDD'),
            children: leftChildren
          }),
          new TableCell({
            width: { size: 4500, type: WidthType.DXA },
            shading: { fill: GREY, type: ShadingType.CLEAR },
            margins: { top: 200, bottom: 200, left: 240, right: 200 },
            borders: borders('BBCCDD'),
            children: rightChildren
          })
        ]
      })
    ]
  });
}

// ── SLA table ────────────────────────────────────────────────────────────────
function slaTable() {
  const hdr = { style: BorderStyle.SINGLE, size: 6, color: BLUE };
  const hBorders = { top: hdr, bottom: hdr, left: hdr, right: hdr };
  const dBorder  = border('BBCCDD');
  const dBorders = { top: dBorder, bottom: dBorder, left: dBorder, right: dBorder };

  function hCell(txt, w) {
    return new TableCell({
      width: { size: w, type: WidthType.DXA },
      shading: { fill: BLUE, type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 160, right: 100 },
      borders: hBorders,
      children: [new Paragraph({ children: [new TextRun({ text: txt, font: 'Arial', size: 20, bold: true, color: 'FFFFFF' })] })]
    });
  }
  function dCell(txt, w) {
    return new TableCell({
      width: { size: w, type: WidthType.DXA },
      margins: { top: 80, bottom: 80, left: 160, right: 100 },
      borders: dBorders,
      children: [new Paragraph({ children: [new TextRun({ text: txt, font: 'Arial', size: 20, color: '333333' })] })]
    });
  }
  function row(a, b, c, d) {
    return new TableRow({ children: [dCell(a, 2800), dCell(b, 2200), dCell(c, 2200), dCell(d, 2160)] });
  }

  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2800, 2200, 2200, 2160],
    rows: [
      new TableRow({ children: [hCell('Kategorie', 2800), hCell('Reaktionszeit', 2200), hCell('Behebungszeit', 2200), hCell('Verfügbarkeit', 2160)] }),
      row('Kritisch (System down)', '2 Stunden', '8 Stunden', '99,5 %'),
      row('Hoch (Kernfunktion)', '4 Stunden', '24 Stunden', '99,0 %'),
      row('Normal (Komfort)', '1 Werktag', '5 Werktage', '–'),
      row('Wartung (geplant)', 'nach Absprache', 'nach Absprache', '–'),
    ]
  });
}

// ── horizontal rule (bottom border on paragraph) ────────────────────────────
function hr() {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 1 } },
    children: []
  });
}

// ── empty line ───────────────────────────────────────────────────────────────
function empty(n) {
  const arr = [];
  for (let i = 0; i < (n || 1); i++) {
    arr.push(new Paragraph({ spacing: { after: 0, before: 0 }, children: [new TextRun('')] }));
  }
  return arr;
}

// ── SIGNATURE TABLE ──────────────────────────────────────────────────────────
function signatureTable() {
  const nb = noBorder();
  function sigCell(partyLabel, cityDate, nameLabel, funcLabel) {
    return new TableCell({
      width: { size: 4320, type: WidthType.DXA },
      margins: { top: 200, bottom: 200, left: 200, right: 200 },
      borders: nb,
      children: [
        new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: partyLabel, font: 'Arial', size: 22, bold: true, color: BLUE })] }),
        new Paragraph({ spacing: { after: 60 }, children: [
          new TextRun({ text: 'Ort, Datum: ', font: 'Arial', size: 20, bold: true, color: '555555' }),
          ph(cityDate)
        ]}),
        empty(1)[0],
        new Paragraph({
          spacing: { after: 0 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BLUE, space: 1 } },
          children: [new TextRun({ text: '', font: 'Arial', size: 22 })]
        }),
        new Paragraph({ spacing: { after: 40, before: 80 }, children: [new TextRun({ text: nameLabel, font: 'Arial', size: 18, color: '888888', italics: true })] }),
        new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: funcLabel, font: 'Arial', size: 18, color: '888888', italics: true })] }),
      ]
    });
  }

  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [4320, 720, 4320],
    rows: [
      new TableRow({
        children: [
          sigCell('Auftraggeber (AG)', 'Ort, Datum AG', 'Name, Funktion', '[Firma des Auftraggebers]'),
          new TableCell({ width: { size: 720, type: WidthType.DXA }, borders: nb, children: [new Paragraph({ children: [] })] }),
          sigCell('Auftragnehmer (AN)', 'Ort, Datum AN', 'Name, Funktion', 'IT-Dienstleister')
        ]
      })
    ]
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// DOCUMENT
// ═══════════════════════════════════════════════════════════════════════════
const doc = new Document({
  styles: {
    default: { document: { run: { font: 'Arial', size: 22 } } }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 }
      }
    },

    headers: {
      default: new Header({
        children: [
          new Paragraph({
            spacing: { after: 60 },
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 2 } },
            children: [
              new TextRun({ text: 'IT-Dienstleistungsvertrag  |  Aktions Tracker', font: 'Arial', size: 18, bold: true, color: BLUE }),
              new TextRun({ text: '    ', font: 'Arial', size: 18 }),
              new TextRun({ text: 'Vertragsnummer: ', font: 'Arial', size: 18, color: '888888' }),
              new TextRun({ text: '[VERTRAGSNR]', font: 'Arial', size: 18, bold: true, color: ORANGE }),
            ]
          })
        ]
      })
    },

    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            spacing: { before: 60 },
            border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'BBCCDD', space: 2 } },
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: 'Seite ', font: 'Arial', size: 18, color: '888888' }),
              new TextRun({ children: [PageNumber.CURRENT], font: 'Arial', size: 18, color: '888888' }),
              new TextRun({ text: ' von ', font: 'Arial', size: 18, color: '888888' }),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], font: 'Arial', size: 18, color: '888888' }),
              new TextRun({ text: '   |   Vertraulich – nur für Vertragsparteien', font: 'Arial', size: 18, color: 'BBBBBB' }),
            ]
          })
        ]
      })
    },

    children: [
      // ── TITLE BLOCK ─────────────────────────────────────────────────────
      new Paragraph({
        spacing: { before: 480, after: 120 },
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: 'IT-DIENSTLEISTUNGSVERTRAG', font: 'Arial', size: 56, bold: true, color: BLUE })]
      }),
      new Paragraph({
        spacing: { after: 80 },
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: 'Aktions Tracker', font: 'Arial', size: 32, bold: true, color: MID })]
      }),
      new Paragraph({
        spacing: { after: 80 },
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: 'Webbasierte Außendienst-Software (SaaS)', font: 'Arial', size: 22, color: '666666', italics: true })]
      }),
      new Paragraph({
        spacing: { after: 480 },
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: 'Vertragsversion: ', font: 'Arial', size: 20, color: '888888' }),
          new TextRun({ text: '[VERSION]', font: 'Arial', size: 20, bold: true, color: ORANGE }),
          new TextRun({ text: '    |    Datum: ', font: 'Arial', size: 20, color: '888888' }),
          new TextRun({ text: '[DATUM]', font: 'Arial', size: 20, bold: true, color: ORANGE }),
        ]
      }),
      hr(),

      // ── VERTRAGSPARTEIEN ─────────────────────────────────────────────────
      sectionHead('Vertragsparteien'),
      twoColTable(
        [
          new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: 'Auftraggeber (AG)', font: 'Arial', size: 22, bold: true, color: BLUE })] }),
          labelVal('Firma', '[Firma des Auftraggebers]', true),
          labelVal('Anschrift', '[Straße, PLZ Ort]', true),
          labelVal('Vertr. durch', 'Name des Geschäftsführers', true),
          labelVal('E-Mail', 'E-Mail AG', true),
          labelVal('USt-IdNr.', 'Ust-Id AG', true),
        ],
        [
          new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: 'Auftragnehmer (AN)', font: 'Arial', size: 22, bold: true, color: BLUE })] }),
          labelVal('Name / Firma', 'Name oder Firma AN', true),
          labelVal('Anschrift', 'Straße, PLZ Ort AN', true),
          labelVal('Vertr. durch', 'Name Inhaber / GF', true),
          labelVal('E-Mail', 'E-Mail AN', true),
          labelVal('USt-IdNr.', 'Ust-Id AN', true),
        ]
      ),
      ...empty(1),

      // ── § 1 ──────────────────────────────────────────────────────────────
      sectionHead('§ 1  Leistungsgegenstand'),
      p('(1) Der AN entwickelt, betreibt und pflegt eine webbasierte Außendienst-Software („Aktions Tracker“) für die Erfassung und Auswertung von Verkaufsaktivitäten des Außendienstteams des AG.'),
      p('(2) Der Funktionsumfang der Software umfasst insbesondere:'),
      bullet('Erfassung von Kundenbesuchen, Platzierungen und Promotionmaßnahmen'),
      bullet('Abbildung von Produktbestellungen und Bestellmengen nach Sortiment'),
      bullet('KPI-Dashboard mit Wochen- und Jahresauswertung (IST / SOLL)'),
      bullet('Rollen- und Rechteverwaltung (Repräsentant, Verkaufsleiter, Admin)'),
      bullet('Progressive Web App (PWA) für mobile Nutzung'),
      bullet('Excel-Export für Vertriebsdaten'),
      p('(3) Änderungen des Leistungsumfangs bedürfen einer schriftlichen Vereinbarung und können gesondert vergütet werden.'),
      ...empty(1),

      // ── § 2 ──────────────────────────────────────────────────────────────
      sectionHead('§ 2  Vergütung'),
      p('(1) Die Vergütung setzt sich aus zwei Komponenten zusammen:'),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [3800, 2280, 3280],
        rows: [
          new TableRow({
            children: [
              new TableCell({
                width: { size: 3800, type: WidthType.DXA },
                shading: { fill: BLUE, type: ShadingType.CLEAR },
                margins: { top: 100, bottom: 100, left: 160, right: 100 },
                borders: borders(BLUE),
                children: [new Paragraph({ children: [new TextRun({ text: 'Komponente', font: 'Arial', size: 20, bold: true, color: 'FFFFFF' })] })]
              }),
              new TableCell({
                width: { size: 2280, type: WidthType.DXA },
                shading: { fill: BLUE, type: ShadingType.CLEAR },
                margins: { top: 100, bottom: 100, left: 160, right: 100 },
                borders: borders(BLUE),
                children: [new Paragraph({ children: [new TextRun({ text: 'Betrag (netto)', font: 'Arial', size: 20, bold: true, color: 'FFFFFF' })] })]
              }),
              new TableCell({
                width: { size: 3280, type: WidthType.DXA },
                shading: { fill: BLUE, type: ShadingType.CLEAR },
                margins: { top: 100, bottom: 100, left: 160, right: 100 },
                borders: borders(BLUE),
                children: [new Paragraph({ children: [new TextRun({ text: 'Fälligkeit / Turnus', font: 'Arial', size: 20, bold: true, color: 'FFFFFF' })] })]
              })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({
                width: { size: 3800, type: WidthType.DXA },
                shading: { fill: GREY, type: ShadingType.CLEAR },
                margins: { top: 100, bottom: 100, left: 160, right: 100 },
                borders: borders('BBCCDD'),
                children: [new Paragraph({ children: [new TextRun({ text: 'Einrichtungsgebühr (einmalig)', font: 'Arial', size: 20, bold: true, color: BLUE })] })]
              }),
              new TableCell({
                width: { size: 2280, type: WidthType.DXA },
                shading: { fill: GREY, type: ShadingType.CLEAR },
                margins: { top: 100, bottom: 100, left: 160, right: 100 },
                borders: borders('BBCCDD'),
                children: [new Paragraph({ children: [ph('BETRAG') , new TextRun({ text: ' €', font: 'Arial', size: 20, color: '333333' })] })]
              }),
              new TableCell({
                width: { size: 3280, type: WidthType.DXA },
                shading: { fill: GREY, type: ShadingType.CLEAR },
                margins: { top: 100, bottom: 100, left: 160, right: 100 },
                borders: borders('BBCCDD'),
                children: [new Paragraph({ children: [new TextRun({ text: 'Einmalig bei Vertragsabschluss', font: 'Arial', size: 20, color: '333333' })] })]
              })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({
                width: { size: 3800, type: WidthType.DXA },
                margins: { top: 100, bottom: 100, left: 160, right: 100 },
                borders: borders('BBCCDD'),
                children: [new Paragraph({ children: [new TextRun({ text: 'Monatliche Pauschale (laufend)', font: 'Arial', size: 20, bold: true, color: BLUE })] })]
              }),
              new TableCell({
                width: { size: 2280, type: WidthType.DXA },
                margins: { top: 100, bottom: 100, left: 160, right: 100 },
                borders: borders('BBCCDD'),
                children: [new Paragraph({ children: [ph('BETRAG') , new TextRun({ text: ' € / Monat', font: 'Arial', size: 20, color: '333333' })] })]
              }),
              new TableCell({
                width: { size: 3280, type: WidthType.DXA },
                margins: { top: 100, bottom: 100, left: 160, right: 100 },
                borders: borders('BBCCDD'),
                children: [new Paragraph({ children: [new TextRun({ text: 'Monatlich im Voraus, zum 1. des Monats', font: 'Arial', size: 20, color: '333333' })] })]
              })
            ]
          })
        ]
      }),
      ...empty(1),
      p('(2) Der AN ist Kleinunternehmer im Sinne des § 19 UStG. Es wird daher keine Umsatzsteuer berechnet und ausgewiesen. Die genannten Beträge sind zugleich Endbeträge (Brutto = Netto).'),
      p('(3) Rechnungen sind innerhalb von 7 Tagen nach Rechnungsdatum ohne Abzug zu begleichen. Als Verwendungszweck ist die aufgedruckte Rechnungsnummer anzugeben. Details zum Zahlungsverzug regelt § 3.'),
      p('(4) Der AN ist berechtigt, die monatliche Pauschale mit einer Ankündigungsfrist von 8 Wochen zum Ende des laufenden Quartals anzupassen. Erhöhungen um mehr als 10 % pro Jahr bedürfen der schriftlichen Zustimmung des AG.'),
      p('(5) Sollte der AN die Kleinunternehmerregelung künftig nicht mehr in Anspruch nehmen (z. B. bei Überschreiten der Umsatzgrenzen gemäß § 19 UStG), wird er den AG rechtzeitig informieren. Ab diesem Zeitpunkt versteht sich die vereinbarte Pauschale zuzüglich der dann gesetzlich gültigen Umsatzsteuer.'),
      p('(6) Bei Vertragsbeginn werden zwei getrennte Rechnungen gestellt: eine Rechnung über die einmalige Einrichtungsgebühr sowie eine Rechnung über die anteilige Monatspauschale für den Startmonat. Die anteilige Berechnung richtet sich nach folgendem Staffelmodell: Beginnt die Nutzung am 1.–8. eines Monats, wird die volle Monatspauschale fällig; beginnt sie am 9.–25., wird der halbe Monatsbeitrag (50 %) berechnet; beginnt sie am 26. oder später, beträgt die Gebühr 25 % der monatlichen Pauschale. Ab dem Folgemonat wird die volle Monatspauschale in Rechnung gestellt.'),
      ...empty(1),

      // ── § 3 ──────────────────────────────────────────────────────────────
      sectionHead('§ 3  Zahlungsverzug und Mahnwesen'),
      p('(1) Gerät der AG mit einer Zahlung in Verzug (d. h. die Rechnung ist nach Ablauf des Zahlungsziels unbeglichen), gilt folgendes Mahnverfahren:'),
      bullet('Zahlungserinnerung (3 Tage nach Fälligkeit): kostenlose Erinnerung per E-Mail'),
      bullet('1. Mahnung (10 Tage nach Fälligkeit): Mahngebühr 5,00 EUR'),
      bullet('2. Mahnung (20 Tage nach Fälligkeit): Mahngebühr 10,00 EUR'),
      bullet('Letzte Mahnung (30 Tage nach Fälligkeit): Mahngebühr 15,00 EUR, Ankündigung rechtlicher Schritte'),
      p('(2) Ab dem Tag nach Eintritt des Verzugs werden Verzugszinsen gemäß § 288 Abs. 2 BGB in Höhe von 9 Prozentpunkten über dem jeweils geltenden Basiszinssatz der Deutschen Bundesbank pro Jahr erhoben. Die Mahngebühren nach Abs. 1 stellen pauschalierten Aufwendungsersatz dar.'),
      p('(3) Bleibt die Forderung nach der letzten Mahnung unbeglichen, ist der AN berechtigt: (a) einen gerichtlichen Mahnbescheid gemäß §§ 688 ff. ZPO zu beantragen; (b) ein Inkassounternehmen zu beauftragen – anfallende Inkassokosten trägt der AG; (c) den Vertrag außerordentlich zu kündigen (vgl. § 4 Abs. 5).'),
      p('(4) Leistet der AG Teilzahlungen, werden diese in folgender Reihenfolge verrechnet: zuerst auf Mahngebühren, dann auf Verzugszinsen, zuletzt auf die Hauptforderung (§ 367 Abs. 1 BGB).'),
      p('(5) Der AN ist berechtigt, den Zugang zur Software bei Zahlungsverzug von mehr als 14 Tagen nach der 1. Mahnung einzuschränken (Service-Pause), bis alle offenen Beträge vollständig beglichen sind. Eine Service-Pause entbindet den AG nicht von seiner Zahlungspflicht.'),
      p('(6) Mahnungen werden per E-Mail an die im Vertrag hinterlegte Rechnungsadresse zugestellt und gelten innerhalb von 24 Stunden als zugegangen.'),
      ...empty(1),

      // ── § 4 ──────────────────────────────────────────────────────────────
      sectionHead('§ 4  Laufzeit und Kündigung'),
      pRuns([
        { text: '(1) Der Vertrag beginnt am ' },
        { text: '[STARTDATUM]', bold: true, color: ORANGE },
        { text: ' und läuft auf unbestimmte Zeit.' }
      ]),
      p('(2) Der AG kann den Vertrag jederzeit mit einer Frist von 4 Wochen zum Ende des laufenden Kalendermonats kündigen. Es gilt keine Mindestlaufzeit.'),
      p('(3) Der AN kann den Vertrag mit einer Frist von 3 Monaten zum Monatsende kündigen, um dem AG ausreichend Zeit für die Umstellung auf eine alternative Lösung zu gewähren.'),
      p('(4) Wurde die monatliche Pauschale als Jahresbetrag im Voraus entrichtet, läuft der Vertrag bis zum Ende des bezahlten Jahreszeitraums weiter und verlängert sich automatisch um ein weiteres Jahr, sofern er nicht spätestens 4 Wochen vor Ablauf des jeweiligen Vertragsjahres in Textform (E-Mail genügt) gekündigt wird. Eine vorzeitige Kündigung durch den AG während des laufenden Jahreszeitraums berechtigt nicht zur anteiligen Rückerstattung des Jahresbetrags. Kündigt der AN das Verhältnis vor Ablauf des bezahlten Zeitraums, erstattet er dem AG die verbleibenden vollen Kalendermonate anteilig zurück.'),
      p('(5) Das Recht beider Parteien zur außerordentlichen Kündigung aus wichtigem Grund bleibt unberührt. Ein wichtiger Grund für den AN liegt insbesondere vor, wenn der AG nach Durchlaufen des Mahnverfahrens gemäß § 3 mit mehr als zwei Monatsbeträgen in Zahlungsverzug bleibt.'),
      p('(6) Kündigungen bedürfen der Schriftform (E-Mail genügt).'),
      ...empty(1),

      // ── § 5 ──────────────────────────────────────────────────────────────
      sectionHead('§ 5  Nutzungsrechte und geistiges Eigentum'),
      p('(1) Der AN gewährt dem AG für die Dauer des Vertrags ein einfaches, nicht exklusives und nicht übertragbares Nutzungsrecht an der Software zum Betrieb im eigenen Unternehmen.'),
      p('(2) Sämtliche Rechte am Quellcode, der Softwarearchitektur, dem Design sowie allen damit verbundenen Materialien verbleiben ausschließlich beim AN. Der AG erwirbt keinerlei Eigentum am Quellcode oder an den geistigen Eigentumsrechten der Software.'),
      p('(3) Die Software wird nicht exklusiv für den AG entwickelt oder betrieben. Der AN ist ausdrücklich berechtigt, die Software – ggf. angepasst oder unter einem anderen Namen – auch anderen Kunden oder Branchen anzubieten, zu lizenzieren und zu betreiben. Der AG erhält durch diesen Vertrag kein Recht auf eine exklusive Nutzung.'),
      p('(4) Der AG ist nicht berechtigt, die Software zu kopieren, weiterzuverkaufen, unterzulizenzieren oder Dritten zugänglich zu machen.'),
      p('(5) Individuelle Anpassungen, die ausschließlich auf Wunsch des AG entwickelt und gesondert vergütet wurden, bleiben zwar im Quellcode-Eigentum des AN; der AG erhält jedoch ein dauerhaftes, einfaches Nutzungsrecht an diesen spezifischen Anpassungen, das über eine etwaige Vertragskündigung hinaus gilt.'),
      ...empty(1),

      // ── § 6 ──────────────────────────────────────────────────────────────
      sectionHead('§ 6  Datenschutz und Auftragsverarbeitung'),
      p('(1) Soweit der AN im Rahmen der Leistungserbringung personenbezogene Daten des AG verarbeitet, handelt er als Auftragsverarbeiter gemäß Art. 28 DSGVO.'),
      p('(2) Beide Parteien verpflichten sich, einen Auftragsverarbeitungsvertrag (AVV) gemäß Art. 28 DSGVO abzuschließen. Der AVV wird als Anlage A beigefügt und ist Bestandteil dieses Vertrags.'),
      p('(3) Der AN setzt die Software auf einer sicheren Infrastruktur (aktuell: Railway / Fly.io oder vergleichbar, Server in der EU) ein und gewährleistet angemessene technische und organisatorische Maßnahmen (TOMs) gemäß Art. 32 DSGVO.'),
      ...empty(1),

      // ── § 7 ──────────────────────────────────────────────────────────────
      sectionHead('§ 7  Support und Wartung'),
      p('(1) Der AN erbringt Support während folgender Zeiten:'),
      pRuns([
        { text: 'Servicezeiten: ' , bold: true },
        { text: '[SERVICEZEITEN, z. B. Mo–Fr 08:00–18:00 Uhr]', bold: true, color: ORANGE }
      ], { afterPt: 8 }),
      p('(2) Supportanfragen können per E-Mail an ' + '[SUPPORT-EMAIL]'.replace('[', '\u{200B}[') + ' oder telefonisch gestellt werden.'),
      p('(3) Geplante Wartungsarbeiten werden mind. 48 Stunden im Voraus angekündigt und finden möglichst außerhalb der Kernarbeitszeiten statt.'),
      ...empty(1),

      // ── § 8 ──────────────────────────────────────────────────────────────
      sectionHead('§ 8  Verfügbarkeit und Service-Level (SLA)'),
      p('(1) Der AN strebt eine Systemverfügbarkeit von mind. 99 % im Monatsdurchschnitt an (gemessen ohne geplante Wartungsfenster).'),
      p('(2) Die folgenden Reaktions- und Behebungszeiten gelten:'),
      slaTable(),
      ...empty(1),

      // ── § 9 ──────────────────────────────────────────────────────────────
      sectionHead('§ 9  Haftung'),
      p('(1) Der AN haftet unbeschränkt für Schäden aus der Verletzung des Lebens, des Körpers oder der Gesundheit sowie für Schäden, die auf Vorsatz oder grober Fahrlässigkeit beruhen.'),
      p('(2) Für leicht fahrlässige Verletzungen wesentlicher Vertragspflichten (Kardinalpflichten) haftet der AN auf den vertragstypisch vorhersehbaren Schaden, maximal jedoch auf die in den letzten 12 Monaten geleisteten Monatspauschalen.'),
      p('(3) Eine weitergehende Haftung ist ausgeschlossen. Der AN haftet insbesondere nicht für mittelbare Schäden, entgangenen Gewinn oder Datenverlust, sofern dieser nicht auf grober Fahrlässigkeit oder Vorsatz beruht.'),
      ...empty(1),

      // ── § 10 ─────────────────────────────────────────────────────────────
      sectionHead('§ 10  Datensicherung und Rückgabe bei Vertragsende'),
      p('(1) Der AN führt regelmäßige automatisierte Datensicherungen durch (mind. täglich).'),
      p('(2) Nach Vertragsbeendigung stellt der AN dem AG auf Anfrage eine vollständige Datensicherung (SQLite-Datenbank und/oder CSV-Export) kostenlos zur Verfügung.'),
      pRuns([
        { text: '(3) Der AN löscht alle Kundendaten spätestens ' },
        { text: '[30 / 60 / 90]', bold: true, color: ORANGE },
        { text: ' Tage nach Vertragsende endgültig, sofern keine gesetzliche Aufbewahrungspflicht entgegensteht.' }
      ]),
      ...empty(1),

      // ── § 11 ─────────────────────────────────────────────────────────────
      sectionHead('§ 11  Schlussbestimmungen'),
      p('(1) Änderungen und Ergänzungen dieses Vertrags bedürfen der Schriftform. Dies gilt auch für die Aufhebung dieser Klausel.'),
      p('(2) Sollten einzelne Bestimmungen dieses Vertrags unwirksam sein oder werden, berührt dies die Gültigkeit der übrigen Bestimmungen nicht.'),
      p('(3) Es gilt das Recht der Bundesrepublik Deutschland. Erfüllungsort und ausschließlicher Gerichtsstand für sämtliche Streitigkeiten ist – soweit gesetzlich zulässig – der Sitz des AN.'),
      p('(4) Dieser Vertrag einschließlich sämtlicher Anlagen stellt die vollständige Vereinbarung der Parteien dar und ersetzt alle früheren Absprachen zum Vertragsgegenstand.'),
      ...empty(2),
      hr(),

      // ── SIGNATURES ───────────────────────────────────────────────────────
      sectionHead('Unterschriften'),
      p('Mit ihrer Unterschrift erklären beide Parteien, den Inhalt dieses Vertrags zur Kenntnis genommen zu haben und damit einverstanden zu sein.'),
      ...empty(1),
      signatureTable(),
      ...empty(2),
      hr(),

      // ── ANLAGEN ──────────────────────────────────────────────────────────
      sectionHead('Anlagenverzeichnis'),
      bullet('Anlage A – Auftragsverarbeitungsvertrag (AVV) gemäß Art. 28 DSGVO'),
      bullet('Anlage B – Leistungsbeschreibung (detaillierter Funktionsumfang)'),
      bullet('Anlage C – Preisliste zusätzlicher Leistungen'),
      ...empty(1),
    ]
  }]
});

// ── WRITE FILE ───────────────────────────────────────────────────────────────
Packer.toBuffer(doc).then(function(buffer) {
  fs.writeFileSync('Dienstleistungsvertrag_Aktions_Tracker_Blanko.docx', buffer);
  console.log('OK: Dienstleistungsvertrag_Aktions_Tracker_Blanko.docx created (' + buffer.length + ' bytes)');
}).catch(function(err) {
  console.error('ERROR:', err);
  process.exit(1);
});
