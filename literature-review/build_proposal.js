const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, convertMillimetersToTwip,
} = require('docx');

const FONT = 'Times New Roman';
const SIZE = 24;   // 12 pt
const LINE = 360;  // 1.5 spacing

const body = (text, opts = {}) => new Paragraph({
  alignment: AlignmentType.JUSTIFIED,
  spacing: { line: LINE, after: 0 },
  indent: opts.first === false ? undefined : { firstLine: convertMillimetersToTwip(10) },
  children: (Array.isArray(text) ? text : [{ text }]).map(r => new TextRun({
    text: r.text, bold: r.bold, italics: r.italics, font: FONT, size: SIZE,
  })),
});

const heading = text => new Paragraph({
  spacing: { line: LINE, before: 180, after: 0 },
  children: [new TextRun({ text, bold: true, font: FONT, size: SIZE })],
});

const centred = (text, opts = {}) => new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { line: LINE, after: opts.after === undefined ? 0 : opts.after },
  children: [new TextRun({ text, bold: opts.bold, italics: opts.italics, font: FONT, size: SIZE })],
});

const refEntry = runs => new Paragraph({
  spacing: { line: LINE, after: 0 },
  indent: { left: convertMillimetersToTwip(12.7), hanging: convertMillimetersToTwip(12.7) },
  children: runs.map(r => new TextRun({ text: r.text, italics: r.italics, font: FONT, size: SIZE })),
});

const doc = new Document({
  creator: 'David Ezieshi',
  title: 'Manufacturing Sincerity — Proposal',
  styles: { default: { document: {
    run: { font: FONT, size: SIZE },
    paragraph: { spacing: { line: LINE } },
  } } },
  sections: [{
    properties: { page: { margin: {
      top: convertMillimetersToTwip(25), right: convertMillimetersToTwip(25),
      bottom: convertMillimetersToTwip(25), left: convertMillimetersToTwip(25),
    } } },
    children: [
      centred('Manufacturing Sincerity:', { bold: true }),
      centred('Authenticity Claims Across Internal and External Communication', { bold: true, after: 180 }),
      centred('Proposal for the literature review assignment (Examination A006)', { italics: true }),
      centred('Module 1: Strategic Communication — Theory, Practice and Critique'),
      centred('MA Strategic Communication, Örebro University'),
      centred('David Ezieshi · 14 September 2026', { after: 240 }),

      heading('Introduction'),

      body('Authenticity has become one of the qualities organisations most want to be credited with. Companies describe their brands as honest and real, public agencies promise transparency, and leaders are urged to bring their whole selves to work, on the assumption that audiences reward it with trust.', { first: false }),

      body('Authenticity nevertheless sits awkwardly with the field that pursues it. Strategic communication is the purposeful use of communication by an organisation to achieve its mission (Hallahan et al., 2007); intent is what makes it strategic. Authenticity implies the absence of exactly that. An organisation that plans a campaign in order to be seen as authentic has arguably already forfeited what it pursues.'),

      body('That tension becomes concrete where internal and external communication meet. An organisation presenting itself externally as honest, human or genuinely committed must have that presentation produced somewhere, and it is produced inside: by employees asked to live the brand and act as its ambassadors (Davis, 2013). The external claim is therefore not only a message about the organisation but a demand upon it. This is the review’s starting point. And as public agencies, universities and political actors adopt promotional logics, claims to authenticity become claims to credibility in public life.'),

      body('Research has approached this ground from several directions without joining them up. Lehman, O’Connor, Kovács and Newman (2019) review the concept in management studies, but communication is not their object. Molleda (2010) reviews it in public relations in order to propose an index of perceived authenticity, treating it as a property to be measured rather than a relation that can be contested. Li et al. (2024) map twenty-one years of brand authenticity research, but their corpus is consumer-facing and excludes critical organisational scholarship. That scholarship is pointed: Fleming and Sturdy (2009) read managerial invitations to “just be yourself” as a form of neo-normative control, and Müller (2017) shows how internal branding enlists external audiences to discipline employees. No existing review reads the managerial and critical strands against each other within strategic communication, and none takes the movement of authenticity claims between internal and external communication as its organising concern. That is the gap addressed here.'),

      heading('Purpose and guiding questions'),

      body('The purpose of this review is to map and critically synthesise how research in strategic communication conceptualises organisational authenticity, and to examine what happens to authenticity claims as they travel between internal and external communication.', { first: false }),

      body([
        { text: 'One question drives the review: ' },
        { text: 'what happens when the authenticity an organisation communicates externally has to be produced, performed and maintained internally?', italics: true },
        { text: ' Two further questions guide the synthesis. Where does the literature locate authenticity: in the organisation, in the judgements of audiences, in the act of communication, or in organisational control? And how does each position account for contradiction between what is communicated inside and claimed outside? The framing is one of mapping rather than adjudication, so that the paradox emerges as a finding rather than a premise.' },
      ]),

      heading('Organising principle, delimitation and outline'),

      body('The results section is organised by where each strand locates authenticity: as a property of the organisation, consistent with its heritage and values; as a judgement attributed by audiences; as an effect produced in communication itself; and as an instrument of normative control. Each strand is then read for what it says about the internal–external relationship. The first two treat the paradox as a solvable problem of consistency; in the latter two it becomes constitutive, and the internal production of external claims appears as a problem rather than a technique.', { first: false }),

      body('Given the ten-page limit, the review addresses organisational-level authenticity in strategic communication, public relations, corporate and internal communication, and branding, from approximately 2005 onwards. Studies whose object is employee attitude or retention rather than communication itself are excluded, as is the authenticity of products, places and heritage. Authentic leadership is represented through its critique rather than surveyed in full.'),

      body('The paper proceeds in four parts: this introduction; a method section reporting databases, search terms and inclusion criteria; a results section following the organising principle above; and a conclusion comparing the findings with the earlier reviews and identifying openings for further research.'),

      heading('Preliminary references'),

      refEntry([
        { text: 'Fleming, P., & Sturdy, A. (2009). “Just be yourself!”: Towards neo-normative control in organisations? ' },
        { text: 'Employee Relations, 31', italics: true }, { text: '(6), 569–583.' },
      ]),
      refEntry([
        { text: 'Lehman, D. W., O’Connor, K., Kovács, B., & Newman, G. E. (2019). Authenticity. ' },
        { text: 'Academy of Management Annals, 13', italics: true }, { text: '(1), 1–42.' },
      ]),
      refEntry([
        { text: 'Li, X., Lim, M.-F., Ramlee, A. N. A., & Chekima, B. (2024). Brand authenticity: A 21-year bibliometric review and future outlook. ' },
        { text: 'SAGE Open', italics: true }, { text: '.' },
      ]),
      refEntry([
        { text: 'Molleda, J.-C. (2010). Authenticity and the construct’s dimensions in public relations and communication research. ' },
        { text: 'Journal of Communication Management, 14', italics: true }, { text: '(3), 223–236.' },
      ]),
      refEntry([
        { text: 'Müller, M. (2017). “Brand-centred control”: A study of internal branding and normative control. ' },
        { text: 'Organization Studies, 38', italics: true }, { text: '(7).' },
      ]),
    ],
  }],
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync('/home/user/Case-study-analysis/literature-review/Proposal_Manufacturing_Sincerity.docx', b);
  console.log('written');
});
