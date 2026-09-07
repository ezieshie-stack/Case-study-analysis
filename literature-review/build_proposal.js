const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, convertMillimetersToTwip,
} = require('docx');

const FONT = 'Times New Roman';
const SIZE = 24;          // 12 pt (half-points)
const LINE = 360;         // 1.5 line spacing (240 = single)

const body = (text, opts = {}) => new Paragraph({
  alignment: opts.alignment || AlignmentType.JUSTIFIED,
  spacing: { line: LINE, after: 0 },
  indent: opts.first === false ? undefined : { firstLine: convertMillimetersToTwip(10) },
  children: (Array.isArray(text) ? text : [{ text }]).map(r => new TextRun({
    text: r.text, bold: r.bold, italics: r.italics, font: FONT, size: SIZE,
  })),
});

const heading = text => new Paragraph({
  alignment: AlignmentType.LEFT,
  spacing: { line: LINE, before: 180, after: 0 },
  children: [new TextRun({ text, bold: true, font: FONT, size: SIZE })],
});

const centred = (text, opts = {}) => new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { line: LINE, after: opts.after === undefined ? 0 : opts.after },
  children: [new TextRun({ text, bold: opts.bold, italics: opts.italics, font: FONT, size: SIZE })],
});

const refEntry = runs => new Paragraph({
  alignment: AlignmentType.LEFT,
  spacing: { line: LINE, after: 60 },
  indent: { left: convertMillimetersToTwip(12.7), hanging: convertMillimetersToTwip(12.7) },
  children: runs.map(r => new TextRun({ text: r.text, italics: r.italics, font: FONT, size: SIZE })),
});

const doc = new Document({
  creator: 'David Ezieshi',
  title: 'The Paradox of Strategic Authenticity — Proposal',
  styles: {
    default: {
      document: { run: { font: FONT, size: SIZE }, paragraph: { spacing: { line: LINE } } },
    },
  },
  sections: [{
    properties: {
      page: {
        margin: {
          top: convertMillimetersToTwip(25), right: convertMillimetersToTwip(25),
          bottom: convertMillimetersToTwip(25), left: convertMillimetersToTwip(25),
        },
      },
    },
    children: [
      // ---- Title block -------------------------------------------------
      centred('The Paradox of Strategic Authenticity:', { bold: true }),
      centred('A Critical Review of Authenticity in Strategic Organisational Communication', { bold: true, after: 180 }),
      centred('Proposal for the literature review assignment (Examination A006)', { italics: true }),
      centred('Module 1: Strategic Communication \u2014 Theory, Practice and Critique'),
      centred('MA Strategic Communication, \u00d6rebro University'),
      centred('David Ezieshi \u00b7 14 September 2026', { after: 240 }),

      // ---- Introduction ------------------------------------------------
      heading('Introduction'),

      body('Authenticity has become one of the qualities organisations most want to be credited with. Companies describe their brands as honest and real, public agencies promise transparency, and leaders are urged to bring their whole selves to work. The quality is attractive because audiences are assumed to reward it with trust.', { first: false }),

      body('Authenticity nevertheless sits awkwardly with the field that pursues it. Strategic communication is the purposeful use of communication by an organisation to achieve its mission (Hallahan et al., 2007); purpose and intent are what make it strategic. Authenticity, in ordinary usage, implies the absence of exactly that. An organisation that plans a campaign in order to be seen as authentic has arguably already forfeited what it is pursuing. This tension is the problem the proposed review takes as its starting point.'),

      body('The problem sharpens where internal and external communication meet. Under the banner of internal branding, employees are asked to live the brand and to act as its ambassadors, while that same brand is presented externally as an expression of what the organisation genuinely is (Davis, 2013). What, then, is asked of an employee required to embody an identity designed for them, and what becomes of the external claim when internal experience contradicts it? The stakes are not only organisational. As public agencies, universities and political actors adopt promotional logics, claims to authenticity become claims to credibility in public life.'),

      body('Research has approached this ground from several directions without joining them up. Lehman, O\u2019Connor, Kov\u00e1cs and Newman (2019) review the concept in management studies, but communication is not their object. Molleda (2010) reviews it in public relations in order to propose an index of perceived authenticity, treating it as a property to be measured rather than a relation that can be contested. Li et al. (2024) map twenty-one years of brand authenticity research, but their corpus is consumer-facing and excludes critical organisational scholarship. That scholarship is pointed: Fleming and Sturdy (2009) read managerial invitations to \u201cjust be yourself\u201d as a form of neo-normative control, and M\u00fcller (2017) shows how internal branding enlists external audiences to discipline employees. No existing review reads the managerial and critical strands against each other within strategic communication, with the internal\u2013external relationship at the centre. That is the gap addressed here.'),

      // ---- Purpose -----------------------------------------------------
      heading('Purpose and guiding questions'),

      body('The purpose of this review is to map and critically synthesise how research in strategic and organisational communication conceptualises organisational authenticity, and to examine how that research accounts for the tension between authenticity as genuineness and strategic communication as intentional, managed meaning-making, with particular attention to the relationship between authenticity claims directed inward at employees and outward at external publics.', { first: false }),

      body('Two questions guide the synthesis. Where does the literature locate authenticity: in the organisation, in the judgements of audiences, in the act of communication, or in the machinery of organisational control? And how does each position handle the contradiction between internal and external claims? The framing is one of mapping rather than adjudication, so that the paradox emerges as a finding rather than a premise.'),

      // ---- Organisation ------------------------------------------------
      heading('Organising principle, delimitation and outline'),

      body('The results section is planned around where each strand of the literature locates authenticity: as a property of the organisation, consistent with its heritage and values; as a judgement attributed by audiences; as an effect produced in communication itself; and as an instrument of normative control over employees. The first two strands treat the paradox as a solvable problem of consistency, the latter two as constitutive and unresolved.', { first: false }),

      body('Given the ten-page limit, the review addresses organisational-level authenticity in strategic communication, public relations, corporate and internal communication, and branding, from approximately 2005 onwards. Authenticity of products, places and heritage is excluded as taking a different referent, and authentic leadership is represented through its critique rather than surveyed in full.'),

      body('The paper proceeds in four parts: this introduction; a method section reporting databases, search terms and inclusion criteria; a results section following the organising principle above; and a conclusion comparing the findings with the earlier reviews and identifying openings for further research.'),

      // ---- References --------------------------------------------------
      heading('Preliminary references'),

      refEntry([
        { text: 'Fleming, P., & Sturdy, A. (2009). \u201cJust be yourself!\u201d: Towards neo-normative control in organisations? ' },
        { text: 'Employee Relations, 31', italics: true },
        { text: '(6), 569\u2013583.' },
      ]),
      refEntry([
        { text: 'Lehman, D. W., O\u2019Connor, K., Kov\u00e1cs, B., & Newman, G. E. (2019). Authenticity. ' },
        { text: 'Academy of Management Annals, 13', italics: true },
        { text: '(1), 1\u201342.' },
      ]),
      refEntry([
        { text: 'Li, X., Lim, M.-F., Ramlee, A. N. A., & Chekima, B. (2024). Brand authenticity: A 21-year bibliometric review and future outlook. ' },
        { text: 'SAGE Open', italics: true },
        { text: '.' },
      ]),
      refEntry([
        { text: 'Molleda, J.-C. (2010). Authenticity and the construct\u2019s dimensions in public relations and communication research. ' },
        { text: 'Journal of Communication Management, 14', italics: true },
        { text: '(3), 223\u2013236.' },
      ]),
      refEntry([
        { text: 'M\u00fcller, M. (2017). \u201cBrand-centred control\u201d: A study of internal branding and normative control. ' },
        { text: 'Organization Studies, 38', italics: true },
        { text: '(7).' },
      ]),
    ],
  }],
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync('/home/user/Case-study-analysis/literature-review/Proposal_Strategic_Authenticity.docx', b);
  console.log('written');
});
