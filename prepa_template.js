const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, AlignmentType, ShadingType,
  VerticalAlign, TableLayoutType, LevelFormat,
  HeadingLevel
} = require('docx');
const fs = require('fs');

// ── Dimensions A4 ─────────────────────────────────────────────────────────────
const TW = 10106;

// ── helpers ───────────────────────────────────────────────────────────────────

function h1(text) {
  return new Paragraph({
    text,
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 240, after: 80 },
  });
}

function h2(text) {
  return new Paragraph({
    text,
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 160, after: 60 },
  });
}

function p(children, opts = {}) {
  const kids = typeof children === 'string'
    ? [new TextRun(children)]
    : Array.isArray(children) ? children : [children];
  return new Paragraph({
    children: kids,
    style: 'Normal',
    alignment: opts.align ?? AlignmentType.LEFT,
    spacing: {
      before: opts.before ?? 0,
      after:  opts.after  ?? 60,
      line: 276, lineRule: 'auto',
    },
    indent: opts.indent ? { left: opts.indent } : undefined,
  });
}

function spacer(twips = 120) {
  return p([new TextRun('')], { before: twips });
}

function dotLine(n = 1) {
  return Array.from({ length: n }, () =>
    new Paragraph({
      children: [new TextRun(' ')],
      style: 'Normal',
      spacing: { before: 0, after: 0 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' } },
    })
  );
}

function chk(text, opts = {}) {
  return p([
    new TextRun({ text: '☐  ' }),
    new TextRun({ text, bold: opts.bold }),
  ], { before: 40, after: 40, indent: opts.indent ?? 240 });
}

// ── Helpers tableaux ──────────────────────────────────────────────────────────

const BORDER = { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' };
const BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER, insideH: BORDER, insideV: BORDER };

function tc(content, opts = {}) {
  let children;
  if (typeof content === 'string') {
    children = [p([new TextRun({ text: content, bold: opts.bold })], { before: 60, after: 60 })];
  } else if (Array.isArray(content)) {
    children = content;
  } else {
    children = [content];
  }
  return new TableCell({
    children,
    width: opts.w ? { size: opts.w, type: WidthType.DXA } : undefined,
    columnSpan: opts.span ?? 1,
    shading: opts.fill ? { type: ShadingType.CLEAR, color: 'auto', fill: opts.fill } : undefined,
    verticalAlign: VerticalAlign.TOP,
    margins: { top: 80, bottom: 80, left: 160, right: 160 },
    borders: BORDERS,
  });
}

function emptyRow(widths, h = 500) {
  return new TableRow({
    children: widths.map(w => tc('', { w })),
    height: { value: h, rule: 'atLeast' },
  });
}

function headerRow(labels, widths) {
  return new TableRow({
    children: labels.map((lbl, i) => tc(lbl, { bold: true, w: widths[i], fill: 'EEEEEE' })),
    tableHeader: true,
  });
}

function mkTable(rows, colWidths) {
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: colWidths,
    layout: TableLayoutType.FIXED,
    borders: BORDERS,
    rows,
  });
}

function doubleEntreeTable(nRows = 3) {
  const W = Math.floor(TW / 2);
  const W2 = TW - W;
  return mkTable([
    headerRow(["Actions / Interventions de l'enseignant(e)", 'Activités / Productions des élèves'], [W, W2]),
    ...Array.from({ length: nRows }, () => emptyRow([W, W2], 560)),
  ], [W, W2]);
}

// ── DOCUMENT ──────────────────────────────────────────────────────────────────

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: 'Calibri', size: 20 } },
      heading1: {
        run: { font: 'Calibri', size: 26, bold: true, color: '1F3864' },
        paragraph: {
          spacing: { before: 240, after: 80 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '1F3864' } },
        },
      },
      heading2: {
        run: { font: 'Calibri', size: 22, bold: true, color: '2E4A87' },
        paragraph: { spacing: { before: 160, after: 60 } },
      },
      normal: {
        run: { font: 'Calibri', size: 20 },
        paragraph: { spacing: { after: 60 } },
      },
    },
  },
  numbering: {
    config: [{
      reference: 'bullets',
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: '•',
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } },
      }],
    }],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 720, bottom: 800, left: 900, right: 900 },
      },
    },
    children: [

      // ── TITRE ────────────────────────────────────────────────────────────────
      // Ligne 1 : "Fiche de Préparation OCT.X.X" (numero en fin de titre)
      new Paragraph({
        children: [new TextRun({ text: 'Fiche de Préparation', bold: true, size: 40, font: 'Calibri', color: '1F3864' })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 20 },
      }),
      // Ligne 2 : titre de l'activité (sous-titre)
      new Paragraph({
        children: [new TextRun({ text: '[Titre de l’activité]', bold: true, size: 30, font: 'Calibri', color: '1F3864' })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 40 },
      }),
      new Paragraph({
        children: [new TextRun({ text: 'Enseignement catholique - SEGEC / CECP', italics: true, size: 19, font: 'Calibri', color: '555555' })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 280 },
      }),

      // ── 1. EN-TETE ───────────────────────────────────────────────────────────
      h1('1. En-tête administratif'),
      mkTable([
        headerRow(["Nom de l'enseignant(e)", 'École', 'Classe / Degré'], [3332, 3442, 3332]),
        new TableRow({ children: [
          tc('Isabelle Giboux',                   { w: 3332 }),
          tc('Ste-Ursule - Maternelles - Laeken', { w: 3442 }),
          tc('M1 - Les abeilles',                 { w: 3332 }),
        ]}),
        headerRow(['Date', 'Durée de la période', 'N° de la fiche'], [3332, 3442, 3332]),
        // Date au format : Mois YYYY (ex. Octobre 2026) - Mois avec majuscule
        new TableRow({ children: [
          tc('[Mois AAAA]', { w: 3332 }),
          tc('[X] minutes', { w: 3442 }),
          tc('[OCT.X.X]',   { w: 3332 }),
        ]}),
      ], [3332, 3442, 3332]),
      spacer(120),
      mkTable([
        headerRow(['Matière / Discipline', 'Titre de la leçon / séquence', 'Niveau'], [2580, 4946, 2580]),
        new TableRow({
          children: [
            tc('[Domaine(s)]', { w: 2580 }),
            tc('[Titre de l’activité]', { w: 4946 }),
            tc([
              chk('Accueil', { indent: 80 }),
              chk('M1',      { indent: 80 }),
              chk('M2',      { indent: 80 }),
              chk('M3',      { indent: 80 }),
            ], { w: 2580 }),
          ],
          height: { value: 560, rule: 'atLeast' },
        }),
      ], [2580, 4946, 2580]),
      spacer(160),

      // ── 2. REFERENCES ────────────────────────────────────────────────────────
      // Colonnes : Code [1800] | Référentiel / Sous-domaine [3000] | Intitulé officiel [5306]
      // Codes autonomie avec préfixe AUT (ex. AUT 4.1.9)
      // Colonne 2 : nom du domaine + numéro de sous-domaine (ex. Autonomie - 2.1 Se mouvoir)
      h1('2. Références aux programmes (référentiels FWB)'),
      mkTable([
        headerRow(['Code', 'Référentiel / Sous-domaine', 'Intitulé officiel'], [1800, 3000, 5306]),
        ...Array.from({ length: 8 }, () => emptyRow([1800, 3000, 5306], 460)),
      ], [1800, 3000, 5306]),
      spacer(160),

      // ── 3. DEMARCHE ──────────────────────────────────────────────────────────
      h1("3. Démarche d'investigation"),
      p([new TextRun({ text: "Cochez l'étape travaillée lors de cette activité :", italics: true })], { before: 0, after: 80 }),
      ...[
        'Recueillir les conceptions / représentations initiales',
        'Amener le questionnement et/ou émettre des hypothèses',
        "S'initier à la recherche, s'investiguer",
        'Comparer - confronter',
        "Garder des traces (mise en évidence de ce qu'ils ont appris)",
      ].map(t => chk(t)),
      spacer(160),

      // ── 4. SITUATION MOBILISATRICE ───────────────────────────────────────────
      h1("4. Situation mobilisatrice et type d'activité"),
      mkTable([
        headerRow(['Situation mobilisatrice', "Type d'activité"], [5053, 5053]),
        new TableRow({
          children: [
            tc([
              "Découverte d'un matériel", 'Situation-problème', 'Défi', 'Jeu', 'Récit',
              'Projet : ..............................', 'Autre : ..............................',
            ].map(t => chk(t, { indent: 120 })), { w: 5053 }),
            tc([
              'Découverte', 'Pratiquer, essayer', 'Consolider, vérifier', 'Rappeler',
            ].map(t => chk(t, { indent: 120 })), { w: 5053 }),
          ],
        }),
      ], [5053, 5053]),
      // Description directement apres le tableau, titre en gras sur la meme ligne que le texte
      p([
        new TextRun({ text: 'Description de la situation : ', bold: true }),
        new TextRun('[Décrire le rôle de l’enseignante et ce que les enfants doivent faire]'),
      ], { before: 100, after: 80 }),
      spacer(80),

      // ── 5. OBJECTIFS ─────────────────────────────────────────────────────────
      h1("5. Objectifs d'apprentissage"),
      p([new TextRun({ text: "Formulez chaque objectif : « L'enfant sera capable de... »", italics: true })], { before: 0, after: 80 }),
      mkTable([
        headerRow(['Prérequis identifiés', "L'enfant sera capable de :"], [3332, 6774]),
        ...Array.from({ length: 5 }, () => emptyRow([3332, 6774], 460)),
      ], [3332, 6774]),
      spacer(160),

      // ── 6. DEROULEMENT ───────────────────────────────────────────────────────
      h1('6. Déroulement pédagogique'),

      h2("A. Activité initiale - Phase d'accroche"),
      doubleEntreeTable(2),
      p([new TextRun({ text: 'Rappel des prérequis / acquis antérieurs :', bold: true })], { before: 120, after: 60 }),
      ...dotLine(2),
      p([new TextRun({ text: 'Situation-problème / Élément déclencheur :', bold: true })], { before: 120, after: 60 }),
      ...dotLine(2),
      p([new TextRun({ text: 'Annonce du sujet (ce que nous allons apprendre et pourquoi) :', bold: true })], { before: 120, after: 60 }),
      ...dotLine(2),

      h2('B. Activité principale - Coeur de la leçon'),
      doubleEntreeTable(4),

      h2('C. Synthèse / Institutionnalisation - Trace écrite'),
      ...dotLine(4),

      // ── 6D : cases a cocher (pas de tirets) ──────────────────────────────────
      h2('D. Évaluation formative'),
      p([new TextRun({ text: "Observations à noter pendant l'activité :", italics: true })], { before: 0, after: 60 }),
      chk("[L'enfant réalise-t-il ... ?]"),
      chk("[L'enfant attend-il ... ?]"),
      chk("[L'enfant accepte-t-il ... ?]"),
      chk("[L'enfant prend-il plaisir à ... ?]"),
      spacer(160),

      // ── 7. DIFFERENCIATION ───────────────────────────────────────────────────
      h1('7. Différenciation pédagogique'),
      p([new TextRun({ text: 'Cochez la ou les stratégie(s) utilisée(s) :', italics: true })], { before: 0, after: 80 }),
      ...[
        ['Différencier les contenus',                   'offrir du matériel supplémentaire, varier les supports'],
        ["Différencier le processus d'apprentissage",   'procéder par étapes, adapter le temps, reformuler les consignes'],
        ["Revenir à l'étape précédente", 'passer par le concret, la découverte par le corps'],
        ['Adapter la tâche',                            'simplifier ou complexifier selon les réussites / difficultés'],
      ].map(([titre, detail]) =>
        p([
          new TextRun('☐  '),
          new TextRun({ text: titre + ' : ', bold: true }),
          new TextRun({ text: detail, italics: true }),
        ], { before: 50, after: 50, indent: 240 })
      ),
      spacer(80),
      mkTable([
        headerRow(["L'environnement / Le dispositif", 'Mode de travail pour les élèves'], [5053, 5053]),
        new TableRow({
          children: [
            tc([
              'Leçon collective : au coin rassemblement', 'Coin entraide',
              'Activité libre avec choix', 'Atelier différencié',
              "S'aider entre eux", 'Co-enseignement',
            ].map(t => chk(t, { indent: 120 })), { w: 5053 }),
            tc([
              'Travail individuel', 'Travail collectif', 'Travail autonome',
              'Travail collaboratif', 'Travail en binôme', 'Autre : ................................',
            ].map(t => chk(t, { indent: 120 })), { w: 5053 }),
          ],
        }),
      ], [5053, 5053]),
      spacer(160),

      // ── 8. MATERIEL ──────────────────────────────────────────────────────────
      h1('8. Matériel / Ressources nécessaires'),
      ...dotLine(5),
      spacer(160),

      // ── 9. CROISEMENTS ───────────────────────────────────────────────────────
      h1("9. Croisements avec d'autres disciplines"),
      mkTable([
        new TableRow({
          children: [
            tc([
              'Français', 'Mathématiques', 'Formation technique et manuelle',
              'Expression plastique', 'Expression musicale',
              'Expression française et corporelle', 'Les autonomies',
            ].map(t =>
              p([
                new TextRun('☐  '),
                new TextRun({ text: t + '  ', bold: true }),
                new TextRun({ text: 'Code : .................', color: '888888' }),
              ], { before: 50, after: 50 })
            ), { w: 5053 }),
            tc([
              'Les visées transversales',
              'Éducation à la philosophie et à la citoyenneté',
              'Éveil aux langues',
              'Formation humaine et sociale',
              'Psychomotricité',
            ].map(t =>
              p([
                new TextRun('☐  '),
                new TextRun({ text: t + '  ', bold: true }),
                new TextRun({ text: 'Code : .................', color: '888888' }),
              ], { before: 50, after: 50 })
            ), { w: 5053 }),
          ],
        }),
      ], [5053, 5053]),
      spacer(280),

      // ── PIED DE PAGE ─────────────────────────────────────────────────────────
      new Paragraph({
        children: [new TextRun({
          text: 'Isabelle Giboux   -   Ste-Ursule - Maternelles - Laeken   -   M1 - Les abeilles',
          size: 17, italics: true, color: '999999', font: 'Calibri',
        })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 100, after: 0 },
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC' } },
      }),
    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/home/claude/prepa_vierge.docx', buf);
  console.log('OK: prepa_vierge.docx');
});
