const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, AlignmentType, ShadingType,
  VerticalAlign, TableLayoutType, LevelFormat,
  HeadingLevel
} = require('docx');
const fs = require('fs');

// -- Dimensions A4 -------------------------------------------------------------
const TW = 10106;

// -- helpers -------------------------------------------------------------------

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

function chkChecked(text, opts = {}) {
  return p([
    new TextRun({ text: '☑  ' }),
    new TextRun({ text, bold: opts.bold }),
  ], { before: 40, after: 40, indent: opts.indent ?? 240 });
}

// -- Helpers tableaux ----------------------------------------------------------

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

// -- DOCUMENT ------------------------------------------------------------------

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

      // -- TITRE ---------------------------------------------------------------
      new Paragraph({
        children: [new TextRun({ text: 'Fiche de Préparation', bold: true, size: 40, font: 'Calibri', color: '1F3864' })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 20 },
      }),
      new Paragraph({
        children: [new TextRun({ text: 'Préparation d’une soupe d’Halloween (Potimarron et Butternut)', bold: true, size: 30, font: 'Calibri', color: '1F3864' })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 40 },
      }),
      new Paragraph({
        children: [new TextRun({ text: 'Enseignement catholique - SEGEC / CECP', italics: true, size: 19, font: 'Calibri', color: '555555' })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 280 },
      }),

      // -- 1. EN-TÊTE ----------------------------------------------------------
      h1('1. En-tête administratif'),
      mkTable([
        headerRow(["Nom de l'enseignant(e)", 'École', 'Classe / Degré'], [3332, 3442, 3332]),
        new TableRow({ children: [
          tc('Isabelle Giboux',                   { w: 3332 }),
          tc('Ste-Ursule - Maternelles - Laeken', { w: 3442 }),
          tc('M1 - Les abeilles',                 { w: 3332 }),
        ]}),
        headerRow(['Date', 'Durée de la période', 'N° de la fiche'], [3332, 3442, 3332]),
        new TableRow({ children: [
          tc('Octobre 2026', { w: 3332 }),
          tc('50 minutes',   { w: 3442 }),
          tc('OCT.13.1',     { w: 3332 }),
        ]}),
      ], [3332, 3442, 3332]),
      spacer(120),
      mkTable([
        headerRow(['Matière / Discipline', 'Titre de la leçon / séquence', 'Niveau'], [2580, 4946, 2580]),
        new TableRow({
          children: [
            tc('Autonomie - Sciences - Formation manuelle - Français', { w: 2580 }),
            tc('Préparation d’une soupe d’Halloween (Potimarron et Butternut) avec les 3ème maternelle', { w: 4946 }),
            tc([
              chk('Accueil', { indent: 80 }),
              chkChecked('M1', { indent: 80 }),
              chk('M2',      { indent: 80 }),
              chk('M3',      { indent: 80 }),
            ], { w: 2580 }),
          ],
          height: { value: 560, rule: 'atLeast' },
        }),
      ], [2580, 4946, 2580]),
      spacer(160),

      // -- 2. RÉFÉRENCES -------------------------------------------------------
      h1('2. Références aux programmes (référentiels FWB)'),
      mkTable([
        headerRow(['Code', 'Référentiel / Sous-domaine', 'Intitulé officiel'], [1800, 3000, 5306]),
        new TableRow({ children: [
          tc('AUT 4.1.1',  { w: 1800 }),
          tc('Autonomie - 4.1 Développer des compétences cognitives transversales', { w: 3000 }),
          tc('Observer une situation pour se poser des questions : Qu’est-ce que c’est ? Comment ? Pourquoi ? Où ? Qui ?...', { w: 5306 }),
        ]}),
        new TableRow({ children: [
          tc('AUT 4.1.9',  { w: 1800 }),
          tc('Autonomie - 4.1 Développer des compétences cognitives transversales', { w: 3000 }),
          tc('Explorer (tester, essayer, tâtonner, manipuler, expérimenter).', { w: 5306 }),
        ]}),
        new TableRow({ children: [
          tc('AUT 4.1.10', { w: 1800 }),
          tc('Autonomie - 4.1 Développer des compétences cognitives transversales', { w: 3000 }),
          tc('Observer en suivant une démarche guidée par l’adulte.', { w: 5306 }),
        ]}),
        new TableRow({ children: [
          tc('AUT 3.1.9',  { w: 1800 }),
          tc('Autonomie - 3.1 Vivre la relation avec les autres', { w: 3000 }),
          tc('Participer aux activités de groupe.', { w: 5306 }),
        ]}),
        new TableRow({ children: [
          tc('FSM 2.1.9',  { w: 1800 }),
          tc('Formation scientifique - 2.1 La matière et les matériaux', { w: 3000 }),
          tc('Dégager des qualités de matières/matériaux rencontrés dans le vécu de la classe à l’aide de certains sens. Distinguer une matière d’une autre en fonction d’une perception sensorielle.', { w: 5306 }),
        ]}),
        new TableRow({ children: [
          tc('FMT 1.2.5',  { w: 1800 }),
          tc('Formation manuelle et technique - 1.2 Utiliser des appareils techniques/technologiques', { w: 3000 }),
          tc('Recourir de manière pertinente à certains appareils techniques/technologiques présents dans l’école et les utiliser.', { w: 5306 }),
        ]}),
        new TableRow({ children: [
          tc('FREC 2.2.13', { w: 1800 }),
          tc('Français Écouter - 2.2', { w: 3000 }),
          tc('Manifester sa compréhension d’une consigne entendue (lue/dite par un tiers). Construire le sens du message entendu.', { w: 5306 }),
        ]}),
        new TableRow({ children: [
          tc('FRP 1.2.6',  { w: 1800 }),
          tc('Français Parler - 1.2', { w: 3000 }),
          tc('Enrichir le bagage lexical pour exprimer un message (vocabulaire spécifique aux domaines) : la cuisine (soupes de légumes...).', { w: 5306 }),
        ]}),
      ], [1800, 3000, 5306]),
      spacer(160),

      // -- 3. DÉMARCHE ---------------------------------------------------------
      h1("3. Démarche d'investigation"),
      p([new TextRun({ text: "Cochez l'étape travaillée lors de cette activité :", italics: true })], { before: 0, after: 80 }),
      chkChecked('Recueillir les conceptions / représentations initiales'),
      chk('Amener le questionnement et/ou émettre des hypothèses'),
      chkChecked("S'initier à la recherche, s'investiguer"),
      chk('Comparer - confronter'),
      chk("Garder des traces (mise en évidence de ce qu'ils ont appris)"),
      spacer(160),

      // -- 4. SITUATION MOBILISATRICE ------------------------------------------
      h1("4. Situation mobilisatrice et type d'activité"),
      mkTable([
        headerRow(['Situation mobilisatrice', "Type d'activité"], [5053, 5053]),
        new TableRow({
          children: [
            tc([
              chk("Découverte d'un matériel",   { indent: 120 }),
              chk('Situation-problème',          { indent: 120 }),
              chk('Défi',                        { indent: 120 }),
              chk('Jeu',                         { indent: 120 }),
              chk('Récit',                       { indent: 120 }),
              chkChecked('Projet : Soupe d’Halloween avec la 3ème maternelle', { indent: 120 }),
              chk('Autre : ..............................', { indent: 120 }),
            ], { w: 5053 }),
            tc([
              chk('Découverte',                  { indent: 120 }),
              chkChecked('Pratiquer, essayer',   { indent: 120 }),
              chk('Consolider, vérifier',        { indent: 120 }),
              chk('Rappeler',                    { indent: 120 }),
            ], { w: 5053 }),
          ],
        }),
      ], [5053, 5053]),
      p([
        new TextRun({ text: 'Description de la situation : ', bold: true }),
        new TextRun('L’enseignante présente à la classe réunie (M1 et 3ème maternelle) un potimarron et un butternut. Les enfants les observent, les touchent, les sentent et essayent de les nommer. Ensuite, avec l’aide des collègues et de deux parents, les enfants participent aux différentes étapes : observer les légumes, éplucher (les adultes se chargent de la découpe au couteau), retirer les pépins et mettre les morceaux dans le soupmaker. La soupe préparée est servie en gobelet et dégustée ensemble.'),
      ], { before: 100, after: 80 }),
      spacer(80),

      // -- 5. OBJECTIFS --------------------------------------------------------
      h1("5. Objectifs d'apprentissage"),
      p([new TextRun({ text: "Formulez chaque objectif : « L'enfant sera capable de... »", italics: true })], { before: 0, after: 80 }),
      mkTable([
        headerRow(['Prérequis identifiés', "L'enfant sera capable de :"], [3332, 6774]),
        new TableRow({ children: [
          tc('Connaissance de légumes courants', { w: 3332 }),
          tc('Observer et nommer le potimarron et le butternut (couleur, forme, texture).', { w: 6774 }),
        ], height: { value: 460, rule: 'atLeast' } }),
        new TableRow({ children: [
          tc('Capacité à écouter une consigne', { w: 3332 }),
          tc('Écouter et suivre les consignes données par un adulte lors des différentes étapes de la recette.', { w: 6774 }),
        ], height: { value: 460, rule: 'atLeast' } }),
        new TableRow({ children: [
          tc('Manipulation simple d’objets', { w: 3332 }),
          tc('Éplucher un légume et retirer les pépins avec l’aide d’un adulte.', { w: 6774 }),
        ], height: { value: 460, rule: 'atLeast' } }),
        new TableRow({ children: [
          tc('Connaissance de l’hygiène de base', { w: 3332 }),
          tc('Se laver les mains correctement avant et après l’activité.', { w: 6774 }),
        ], height: { value: 460, rule: 'atLeast' } }),
        new TableRow({ children: [
          tc('Expérience de travail en groupe', { w: 3332 }),
          tc('Coopérer avec ses camarades et les adultes présents lors de la préparation.', { w: 6774 }),
        ], height: { value: 460, rule: 'atLeast' } }),
      ], [3332, 6774]),
      spacer(160),

      // -- 6. DÉROULEMENT ------------------------------------------------------
      h1('6. Déroulement pédagogique'),

      h2("A. Activité initiale - Phase d'accroche"),
      (() => {
        const W = Math.floor(TW / 2); const W2 = TW - W;
        return mkTable([
          headerRow(["Actions / Interventions de l'enseignant(e)", 'Activités / Productions des élèves'], [W, W2]),
          new TableRow({ children: [
            tc('Présente le potimarron et le butternut à la classe réunie. Pose des questions : "Qui connaît ce légume ? Quelle couleur ? Comment s\'appelle-t-il ?"', { w: W }),
            tc('Observent, touchent, sentent les légumes. Tentent de les nommer et partagent leurs observations avec le groupe.', { w: W2 }),
          ], height: { value: 560, rule: 'atLeast' } }),
          new TableRow({ children: [
            tc('Explique les étapes de l\'activité et les règles de sécurité. Invite les enfants à se laver les mains.', { w: W }),
            tc('Écoutent les consignes. Se lavent les mains avant de commencer l\'activité.', { w: W2 }),
          ], height: { value: 560, rule: 'atLeast' } }),
        ], [W, W2]);
      })(),
      p([new TextRun({ text: 'Rappel des prérequis / acquis antérieurs :', bold: true })], { before: 120, after: 60 }),
      p('Les enfants ont déjà manipulé des légumes lors d\'activités précédentes. Certains ont peut-être cuisiné à la maison.', { before: 0, after: 60 }),
      p([new TextRun({ text: 'Situation-problème / Élément déclencheur :', bold: true })], { before: 120, after: 60 }),
      p('Arrivée des légumes (potimarron et butternut) dans la classe : "Qu\'est-ce que c\'est ? Comment peut-on les manger ?"', { before: 0, after: 60 }),
      p([new TextRun({ text: 'Annonce du sujet (ce que nous allons apprendre et pourquoi) :', bold: true })], { before: 120, after: 60 }),
      p('Aujourd\'hui, nous allons préparer une soupe d\'Halloween avec nos amis de 3ème maternelle et deux parents. Nous allons éplucher les légumes et utiliser le soupmaker.', { before: 0, after: 60 }),

      h2('B. Activité principale - Cœur de la leçon'),
      (() => {
        const W = Math.floor(TW / 2); const W2 = TW - W;
        return mkTable([
          headerRow(["Actions / Interventions de l'enseignant(e)", 'Activités / Productions des élèves'], [W, W2]),
          new TableRow({ children: [
            tc('Organise les groupes avec les collègues et les deux parents. Montre comment utiliser l\'éplucheur en faisant une démonstration.', { w: W }),
            tc('Observent la démonstration avec attention.', { w: W2 }),
          ], height: { value: 560, rule: 'atLeast' } }),
          new TableRow({ children: [
            tc('Supervise l\'épluchage avec les adultes présents. Guide le geste, rassure, encourage.', { w: W }),
            tc('Épluchent les légumes avec l\'éplucheur, sous supervision adulte.', { w: W2 }),
          ], height: { value: 560, rule: 'atLeast' } }),
          new TableRow({ children: [
            tc('Montre comment retirer les pépins à la cuillère. Les adultes encadrent chaque groupe.', { w: W }),
            tc('Retirent les pépins à la cuillère et les déposent dans le bassin.', { w: W2 }),
          ], height: { value: 560, rule: 'atLeast' } }),
          new TableRow({ children: [
            tc('Place les morceaux dans le soupmaker avec les enfants. Ajoute l\'eau, l\'oignon et le cube bouillon. Lance la cuisson.', { w: W }),
            tc('Déposent les morceaux dans le soupmaker. Observent et attendent la cuisson. Sentent les odeurs qui se dégagent.', { w: W2 }),
          ], height: { value: 560, rule: 'atLeast' } }),
        ], [W, W2]);
      })(),

      h2('C. Synthèse / Institutionnalisation - Trace écrite'),
      p('La soupe est servie en gobelet. L\'enseignante invite les enfants à décrire le goût, la couleur et l\'odeur. Retour collectif oral sur les différentes étapes réalisées ensemble : "Qu\'avons-nous fait en premier ? Et ensuite ?"', { before: 60, after: 60 }),

      // -- 6D : cases à cocher (pas de tirets) ---------------------------------
      h2('D. Évaluation formative'),
      p([new TextRun({ text: "Observations à noter pendant l'activité :", italics: true })], { before: 0, after: 60 }),
      chk("L'enfant identifie-t-il le potimarron et le butternut ?"),
      chk("L'enfant respecte-t-il les consignes de sécurité (couteaux, chaleur) ?"),
      chk("L'enfant participe-t-il activement à chacune des étapes ?"),
      chk("L'enfant prend-il plaisir à préparer et à déguster la soupe ?"),
      spacer(160),

      // -- 7. DIFFÉRENCIATION --------------------------------------------------
      h1('7. Différenciation pédagogique'),
      p([new TextRun({ text: 'Cochez la ou les stratégie(s) utilisée(s) :', italics: true })], { before: 0, after: 80 }),
      ...[
        ['Différencier les contenus',                    'offrir du matériel supplémentaire, varier les supports'],
        ["Différencier le processus d'apprentissage",    'procéder par étapes, adapter le temps, reformuler les consignes'],
        ["Revenir à l'étape précédente",                 'passer par le concret, la découverte par le corps'],
        ['Adapter la tâche',                             'simplifier ou complexifier selon les réussites / difficultés'],
      ].map(([titre, detail], i) =>
        p([
          new TextRun(i === 3 ? '☑  ' : '☐  '),
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
              chkChecked('Leçon collective : au coin rassemblement', { indent: 120 }),
              chk('Coin entraide',                                    { indent: 120 }),
              chk('Activité libre avec choix',                        { indent: 120 }),
              chkChecked('Atelier différencié',                       { indent: 120 }),
              chk("S'aider entre eux",                               { indent: 120 }),
              chkChecked('Co-enseignement',                           { indent: 120 }),
            ], { w: 5053 }),
            tc([
              chk('Travail individuel',                               { indent: 120 }),
              chkChecked('Travail collectif',                         { indent: 120 }),
              chk('Travail autonome',                                 { indent: 120 }),
              chkChecked('Travail collaboratif',                      { indent: 120 }),
              chk('Travail en binôme',                                { indent: 120 }),
              chk('Autre : ................................',          { indent: 120 }),
            ], { w: 5053 }),
          ],
        }),
      ], [5053, 5053]),
      spacer(160),

      // -- 8. MATÉRIEL ---------------------------------------------------------
      h1('8. Matériel / Ressources nécessaires'),
      p('Butternut, potimarron, oignon, cube bouillon, eau, récipient, couteaux (adultes), cuillère, planche à découper, soupmaker, louche, gobelets, bassin, savon pour les mains, essuie-mains, rallonge électrique, éplucheur.', { before: 60, after: 80 }),
      ...dotLine(4),
      spacer(160),

      // -- 9. CROISEMENTS ------------------------------------------------------
      h1("9. Croisements avec d'autres disciplines"),
      mkTable([
        new TableRow({
          children: [
            tc([
              p([
                new TextRun('☑  '),
                new TextRun({ text: 'Français  ', bold: true }),
                new TextRun({ text: 'Code : FRP 1.2.6 / FREC 2.2.13', color: '888888' }),
              ], { before: 50, after: 50 }),
              p([
                new TextRun('☐  '),
                new TextRun({ text: 'Mathématiques  ', bold: true }),
                new TextRun({ text: 'Code : .................', color: '888888' }),
              ], { before: 50, after: 50 }),
              p([
                new TextRun('☑  '),
                new TextRun({ text: 'Formation technique et manuelle  ', bold: true }),
                new TextRun({ text: 'Code : FMT 1.2.5', color: '888888' }),
              ], { before: 50, after: 50 }),
              p([
                new TextRun('☐  '),
                new TextRun({ text: 'Expression plastique  ', bold: true }),
                new TextRun({ text: 'Code : .................', color: '888888' }),
              ], { before: 50, after: 50 }),
              p([
                new TextRun('☐  '),
                new TextRun({ text: 'Expression musicale  ', bold: true }),
                new TextRun({ text: 'Code : .................', color: '888888' }),
              ], { before: 50, after: 50 }),
              p([
                new TextRun('☐  '),
                new TextRun({ text: 'Expression française et corporelle  ', bold: true }),
                new TextRun({ text: 'Code : .................', color: '888888' }),
              ], { before: 50, after: 50 }),
              p([
                new TextRun('☑  '),
                new TextRun({ text: 'Les autonomies  ', bold: true }),
                new TextRun({ text: 'Code : AUT 4.1.1 / AUT 4.1.9 / AUT 3.1.9', color: '888888' }),
              ], { before: 50, after: 50 }),
            ], { w: 5053 }),
            tc([
              p([
                new TextRun('☐  '),
                new TextRun({ text: 'Les visées transversales  ', bold: true }),
                new TextRun({ text: 'Code : .................', color: '888888' }),
              ], { before: 50, after: 50 }),
              p([
                new TextRun('☐  '),
                new TextRun({ text: 'Éducation à la philosophie et à la citoyenneté  ', bold: true }),
                new TextRun({ text: 'Code : .................', color: '888888' }),
              ], { before: 50, after: 50 }),
              p([
                new TextRun('☐  '),
                new TextRun({ text: 'Éveil aux langues  ', bold: true }),
                new TextRun({ text: 'Code : .................', color: '888888' }),
              ], { before: 50, after: 50 }),
              p([
                new TextRun('☐  '),
                new TextRun({ text: 'Formation humaine et sociale  ', bold: true }),
                new TextRun({ text: 'Code : .................', color: '888888' }),
              ], { before: 50, after: 50 }),
              p([
                new TextRun('☐  '),
                new TextRun({ text: 'Psychomotricité  ', bold: true }),
                new TextRun({ text: 'Code : .................', color: '888888' }),
              ], { before: 50, after: 50 }),
            ], { w: 5053 }),
          ],
        }),
      ], [5053, 5053]),
      spacer(280),

      // -- PIED DE PAGE --------------------------------------------------------
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
  fs.writeFileSync('/home/claude/ste-ursule/OCT.13.1_SoupeHalloween.docx', buf);
  console.log('OK: OCT.13.1_SoupeHalloween.docx');
});
