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

// Numbered research question, hanging indent, as in both student examples
const rq = (n, text) => new Paragraph({
  alignment: AlignmentType.LEFT,
  spacing: { line: LINE, after: 0 },
  indent: { left: convertMillimetersToTwip(12), hanging: convertMillimetersToTwip(6) },
  children: [new TextRun({ text: `${n}.  ${text}`, font: FONT, size: SIZE })],
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
  title: 'Can Authenticity Be Strategic? Proposal',
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
      centred('Can Authenticity Be Strategic?', { bold: true }),
      centred('A Critical Review of Authenticity Claims Across Internal and External Communication', { bold: true, after: 180 }),
      centred('Proposal for the literature review assignment (Examination A006)', { italics: true }),
      centred('Module 1: Strategic Communication - Theory, Practice and Critique'),
      centred('MA Strategic Communication, Örebro University'),
      centred('David Ezieshi · 14 September 2026', { after: 200 }),

      heading('Introduction'),

      body('Authenticity has become one of the qualities organizations most want to be credited with. Companies describe their brands as honest, public agencies promise transparency, and leaders are urged to bring their whole selves to work, on the assumption that audiences reward it with trust.', { first: false }),

      body('Authenticity nevertheless sits awkwardly with the field that pursues it. Strategic communication is the purposeful use of communication by an organization to achieve its mission (Hallahan et al., 2007); intent is what makes it strategic. Authenticity implies the absence of exactly that. An organization that plans a campaign to be seen as authentic has arguably already forfeited what it pursues.'),

      body('That tension becomes concrete where internal and external communication meet. An organization presenting itself externally as honest, human or genuinely committed must have that presentation produced somewhere, and it is produced inside, by employees asked to live the brand (Davis, 2013). The external claim is therefore not only a message about the organization but a demand upon it. And as public agencies, universities and political actors adopt promotional logics, claims to authenticity become claims to credibility in public life.'),

      body('Research has approached this ground from several directions without joining them up. Lehman, O’Connor, Kovács and Newman (2019) review the concept in management studies, but communication is not their object. Molleda (2010) reviews it in public relations to propose an index of perceived authenticity, treating it as a property to be measured rather than a relation that can be contested. Li et al. (2024) map twenty-one years of brand authenticity research, but the studies they review concern consumers rather than organizations, and critical work on what authenticity does inside the organization falls outside their scope.'),

      body('That work is uncomfortable reading. Fleming and Sturdy (2009), studying a call center where staff were encouraged to “just be yourself”, found that inviting employees to bring their real selves to work turned those selves into a resource for the employer. Müller (2017) found that internal branding is watched not only by managers but by customers and the public, so that the external audience becomes a further means of holding employees to the brand. Christensen and Cornelissen (2011) bridge corporate and organizational communication, criticizing the ideal of an organization that speaks with one voice and proposing polyphony as an alternative to control-based integration. Their object, however, is communication theory rather than authenticity, and the reviews that do take authenticity as their object ignore that critique. That is the gap this review addresses: what has to happen inside an organization for a claim made outside it to hold.'),

      heading('Purpose and research questions'),

      body('The purpose of this review is to map and critically synthesize how research in strategic communication conceptualizes organizational authenticity, and to examine what happens to authenticity claims as they travel between internal and external communication. Given the ten-page limit, it addresses authenticity at the level of the organization, in public relations, corporate and internal communication, and branding, published from 2005 onwards.', { first: false }),

      body('Three questions guide the review:'),

      rq(1, 'What happens when the authenticity an organization communicates externally has to be produced, performed and maintained internally?'),
      rq(2, 'Who or what makes an organization authentic: the organization itself, its audiences, the communication that produces the claim, or the managers who put that claim to work?'),
      rq(3, 'What does each of those answers say happens when the inside and the outside do not match?'),

      body('Those answers disagree with one another, and that disagreement is the review’s result rather than its starting assumption.', { first: false }),

      heading('Outline'),

      body('The review is structured as follows. The next section describes how the literature was searched and selected, reporting the databases, search terms and inclusion criteria used. The results are then presented in four parts, according to where each strand places authenticity: as a property of the organization; as a judgment made by audiences; as an effect of communication; and as an instrument of control over employees. Each part is read for what it says about the relationship between internal and external communication. The first two strands treat the tension as a solvable problem of consistency; in the second two it becomes constitutive, and the internal production of external claims appears as a problem rather than a technique. The conclusion returns to the research questions, compares the findings with the earlier reviews, and identifies openings for further research.', { first: false }),

      heading('Preliminary references'),

      refEntry([
        { text: 'Christensen, L. T., & Cornelissen, J. (2011). Bridging corporate and organizational communication: Review, development and a look to the future. ' },
        { text: 'Management Communication Quarterly, 25', italics: true }, { text: '(3), 383–414.' },
      ]),
      refEntry([
        { text: 'Lehman, D. W., O’Connor, K., Kovács, B., & Newman, G. E. (2019). Authenticity. ' },
        { text: 'Academy of Management Annals, 13', italics: true }, { text: '(1), 1–42.' },
      ]),
      refEntry([
        { text: 'Molleda, J.-C. (2010). Authenticity and the construct’s dimensions in public relations and communication research. ' },
        { text: 'Journal of Communication Management, 14', italics: true }, { text: '(3), 223–236.' },
      ]),
    ],
  }],
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync('/home/user/Case-study-analysis/literature-review/Proposal_Can_Authenticity_Be_Strategic.docx', b);
  console.log('written');
});
