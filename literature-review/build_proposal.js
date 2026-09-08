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
  spacing: { line: LINE, before: 120, after: 0 },
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

      body('That tension becomes concrete where internal and external communication meet. An organisation presenting itself externally as honest, human or genuinely committed must have that presentation produced somewhere, and it is produced inside, by employees asked to live the brand (Davis, 2013). The external claim is therefore not only a message about the organisation but a demand upon it. This is the review’s starting point. And as public agencies, universities and political actors adopt promotional logics, claims to authenticity become claims to credibility in public life.'),

      body('Research has approached this ground from several directions without joining them up. Lehman, O’Connor, Kovács and Newman (2019) review the concept in management studies, but communication is not their object. Molleda (2010) reviews it in public relations to propose an index of perceived authenticity, treating it as a property to be measured rather than a relation that can be contested. Li et al. (2024) map twenty-one years of brand authenticity research, but the studies they review concern consumers rather than organisations, and critical work on what authenticity does inside the organisation falls outside their scope.'),

      body('That work is uncomfortable reading. Fleming and Sturdy (2009), studying a call centre where staff were encouraged to “just be yourself”, found that inviting employees to bring their real selves to work turned those selves into a resource for the employer. Müller (2017) found that internal branding is watched not only by managers but by customers and the public, so that the external audience becomes a further means of holding employees to the brand. No existing review puts these two bodies of work side by side: the research that treats authenticity as something an organisation can achieve and demonstrate, and the research that treats it as something demanded of employees. Nor does any of them follow an authenticity claim as it moves between the inside of the organisation and the outside. That is the gap this review addresses.'),

      heading('Purpose and guiding questions'),

      body('The purpose of this review is to map and critically synthesise how research in strategic communication conceptualises organisational authenticity, and to examine what happens to authenticity claims as they travel between internal and external communication.', { first: false }),

      body([
        { text: 'One question drives the review: ' },
        { text: 'what happens when the authenticity an organisation communicates externally has to be produced, performed and maintained internally?', italics: true },
        { text: ' Two further questions guide the synthesis. Where does the literature locate authenticity: in the organisation, in the judgements of audiences, in the act of communication, or in organisational control? And how does each position account for contradiction between what is communicated inside and claimed outside? The review maps the debate rather than settling it, so that the paradox emerges as a finding rather than a premise.' },
      ]),

      heading('Organising principle, delimitation and outline'),

      body('The results section is organised by where each strand locates authenticity: as a property of the organisation, consistent with its heritage and values; as a judgement attributed by audiences; as an effect produced in communication itself; and as an instrument of normative control. Each strand is then read for what it says about the internal–external relationship. The first two treat the paradox as a solvable problem of consistency; in the latter two it becomes constitutive, and the internal production of external claims appears as a problem rather than a technique.', { first: false }),

      body('Given the ten-page limit, the review addresses organisational-level authenticity in strategic communication, public relations, corporate and internal communication, and branding, from 2005 onwards. Studies whose object is employee attitude or retention rather than communication itself are excluded, as is the authenticity of products, places and heritage. Authentic leadership is represented through its critique rather than surveyed in full.'),

      body('The paper proceeds in four parts: this introduction; a method section reporting databases, search terms and inclusion criteria; a results section following the organising principle; and a conclusion comparing the findings with the earlier reviews and identifying openings for research.'),

      heading('Preliminary references'),

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
    ],
  }],
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync('/home/user/Case-study-analysis/literature-review/Proposal_Manufacturing_Sincerity.docx', b);
  console.log('written');
});
