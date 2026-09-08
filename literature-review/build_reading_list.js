const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, convertMillimetersToTwip, HeadingLevel,
} = require('docx');

const FONT = 'Times New Roman';
const S = 22;      // 11 pt — working document, not a submission
const LINE = 276;  // ~1.15 spacing

const p = (runs, opts = {}) => new Paragraph({
  alignment: opts.align || AlignmentType.LEFT,
  spacing: { line: LINE, before: opts.before || 0, after: opts.after === undefined ? 60 : opts.after },
  indent: opts.indent,
  children: (Array.isArray(runs) ? runs : [{ text: runs }]).map(r => new TextRun({
    text: r.text, bold: r.bold, italics: r.italics, font: FONT, size: opts.size || S,
  })),
});

const h1 = text => new Paragraph({
  spacing: { line: LINE, before: 240, after: 80 },
  children: [new TextRun({ text, bold: true, font: FONT, size: 26 })],
});

const h2 = text => new Paragraph({
  spacing: { line: LINE, before: 200, after: 60 },
  children: [new TextRun({ text, bold: true, font: FONT, size: S })],
});

// One source: citation, link, what it argues, where it goes
const src = ({ cite, link, argues, use, open, flag }) => {
  const out = [p(cite, { after: 20 })];
  if (link) out.push(p([{ text: link, italics: true }], { after: 20, indent: { left: convertMillimetersToTwip(6) } }));
  out.push(p([{ text: 'Argues: ', bold: true }, { text: argues }], { after: 20, indent: { left: convertMillimetersToTwip(6) } }));
  out.push(p([{ text: 'Use it for: ', bold: true }, { text: use }], { after: open || flag ? 20 : 120, indent: { left: convertMillimetersToTwip(6) } }));
  if (open) out.push(p([{ text: 'Free PDF: ', bold: true }, { text: open, italics: true }], { after: flag ? 20 : 120, indent: { left: convertMillimetersToTwip(6) } }));
  if (flag) out.push(p([{ text: 'Check before citing: ', bold: true }, { text: flag }], { after: 120, indent: { left: convertMillimetersToTwip(6) } }));
  return out;
};

const doc = new Document({
  creator: 'David Ezieshi',
  title: 'Manufacturing Sincerity — Reading List',
  styles: { default: { document: { run: { font: FONT, size: S }, paragraph: { spacing: { line: LINE } } } } },
  sections: [{
    properties: { page: { margin: {
      top: convertMillimetersToTwip(20), right: convertMillimetersToTwip(20),
      bottom: convertMillimetersToTwip(20), left: convertMillimetersToTwip(20),
    } } },
    children: [
      p('Manufacturing Sincerity: Reading List', { align: AlignmentType.CENTER, size: 30, after: 40 }),
      p([{ text: 'Authenticity Claims Across Internal and External Communication', italics: true }], { align: AlignmentType.CENTER, after: 40 }),
      p('Literature review, Module 1 · MA Strategic Communication, Örebro University', { align: AlignmentType.CENTER, after: 240 }),

      p([{ text: 'How to use this. ', bold: true }, { text: 'Seventeen sources against a required minimum of twelve, grouped by where each sits in the review. Start with the five marked READ FIRST; they carry the argument. Bibliographic details were checked against publisher records, but page ranges and full author lists were not all confirmed — anything uncertain is flagged, and you should pull the record from the database before it goes in your reference list. Access everything through the Örebro library proxy; several are open access and the free link is given.' }]),

      // ---------------------------------------------------------------
      h1('1. The reviews you are writing against'),
      p('Your introduction has to answer “are there existing reviews, and what does mine add?” These are the three. None of them does what you are doing, and the proposal says why.'),

      ...src({
        cite: 'Lehman, D. W., O’Connor, K., Kovács, B., & Newman, G. E. (2019). Authenticity. Academy of Management Annals, 13(1), 1–42.   [READ FIRST]',
        link: 'https://journals.aom.org/doi/abs/10.5465/annals.2017.0047',
        open: 'https://spinup-000d1a-wp-offload-media.s3.amazonaws.com/faculty/wp-content/uploads/sites/72/2020/02/Annals-Authenticity.pdf',
        argues: 'Authenticity means different things depending on three questions: what is the reference point, what is being judged, and who is judging. They call this the 3C view. A 42-page integrative review.',
        use: 'The backbone of your “what is already known” paragraph, and a ready-made vocabulary for your organising principle. Your gap: it is management, not communication.',
      }),
      ...src({
        cite: 'Molleda, J.-C. (2010). Authenticity and the construct’s dimensions in public relations and communication research. Journal of Communication Management, 14(3), 223–236.   [READ FIRST]',
        link: 'https://www.emerald.com/jcom/article-abstract/14/3/223/435169/',
        argues: 'Reviews authenticity across advertising, marketing and PR, then proposes an index for measuring the perceived authenticity of organisational messages and actions.',
        use: 'The closest existing review to yours, and the clearest contrast: it treats authenticity as a property to be measured. Your review treats it as a relation that can be contested. Say so explicitly.',
      }),
      ...src({
        cite: 'Li, X., Lim, M.-F., Ramlee, A. N. A., & Chekima, B. (2024). Brand authenticity: A 21-year bibliometric review and future outlook. SAGE Open.',
        link: 'https://journals.sagepub.com/doi/10.1177/21582440241268847',
        argues: 'Maps two decades of brand authenticity research bibliometrically and sets out where the field is heading.',
        use: 'Shows the marketing side of the field and its blind spot: the corpus is consumer-facing and leaves out critical organisational work. Useful for one paragraph, not more.',
        flag: 'Volume and article number not confirmed.',
      }),

      ...src({
        cite: 'Christensen, L. T., & Cornelissen, J. (2011). Bridging corporate and organizational communication: Review, development and a look to the future. Management Communication Quarterly.   [READ FIRST]',
        link: 'https://journals.sagepub.com/doi/abs/10.1177/0893318910390194',
        argues: 'Reviews and joins two fields that had grown apart: corporate communication, which looks outward at markets and publics, and organizational communication, which looks inward at employees and culture.',
        use: 'The review closest to your framing, and the one you must name. It bridges the same two fields you do, but it is not about authenticity and it predates the critical work on brand control. Saying that precisely is what makes your gap defensible.',
        flag: 'Volume, issue and pages need confirming.',
      }),
      ...src({
        cite: 'Employees as a second audience: The effect of external communication on internal brand management outcomes (2018). Journal of Brand Management.',
        link: 'https://link.springer.com/article/10.1057/s41262-018-0135-z',
        argues: 'Employees are an audience for their employer’s external communication, not only its producers. Where external messages match internal communication and actual practice, employees understand the brand better; where they do not, the effect weakens.',
        use: 'Your mechanism, already tested. Not a review, so it does not compete with your gap — it is evidence that the question you are asking is one researchers take seriously.',
        flag: 'Authors and page range not confirmed.',
      }),

      // ---------------------------------------------------------------
      h1('2. The four strands (your results section)'),

      h2('Strand 1 — Authenticity as a property of the organisation'),
      p('The essentialist position: an organisation is authentic when what it says matches its heritage, origins and stated values. This strand treats your paradox as a solvable consistency problem.'),
      ...src({
        cite: 'Napoli, J., Dickinson, S. J., Beverland, M. B., & Farrelly, F. (2014). Measuring consumer-based brand authenticity. Journal of Business Research.',
        link: 'https://www.researchgate.net/publication/260609203_Measuring_consumer-based_brand_authenticity',
        argues: 'Builds brand authenticity from three components: commitment to quality, sincerity, and heritage.',
        use: 'A clean example of the essentialist logic. One or two paragraphs; do not survey this literature in depth.',
        flag: 'Author order, volume and pages need confirming.',
      }),

      h2('Strand 2 — Authenticity as a judgement made by audiences'),
      p('The receiver-side position: authenticity is not in the object, it is attributed by those perceiving it. Also treats the paradox as manageable — get the signals right and the attribution follows.'),
      ...src({
        cite: 'Morhart, F., Malär, L., Guèvremont, A., Girardin, F., & Grohmann, B. (2015). Brand authenticity: An integrative framework and measurement scale. Journal of Consumer Psychology, 25(2), 200–218.',
        link: 'https://myscp.onlinelibrary.wiley.com/doi/abs/10.1016/j.jcps.2014.11.006',
        open: 'https://www2.novasbe.unl.pt/Portals/0/Research/documents/lucia.pdf',
        argues: 'Perceived brand authenticity has four dimensions: credibility, integrity, symbolism and continuity. Validated across brands and cultures.',
        use: 'The standard reference for the measurement approach. Pair it with Molleda to show how the field tried to operationalise a contested concept.',
        flag: 'Confirm the full author list and page range.',
      }),

      h2('Strand 3 — Authenticity as something produced in communication'),
      p('The performative position, and the pivot of your review: authenticity is not found, it is made — in stories, images and talk. This is where the internal/external axis starts to matter.'),
      ...src({
        cite: 'Christensen, L. T., Morsing, M., & Thyssen, O. (2013). CSR as aspirational talk. Organization, 20(3).   [READ FIRST]',
        link: 'https://journals.sagepub.com/doi/10.1177/1350508413478310',
        argues: 'The gap between what organisations say and what they do is not necessarily a failure. Aspirational talk can pull an organisation towards its stated ideals. Communication is performative, not merely descriptive.',
        use: 'The strongest counter-argument to your own paradox, and you must engage it. If talk can produce the reality it describes, then strategic authenticity is not automatically a contradiction. Your review is better for taking this seriously than for dismissing it.',
        flag: 'Confirm page range.',
      }),
      ...src({
        cite: 'Christensen, L. T., Morsing, M., & Thyssen, O. Timely hypocrisy? Hypocrisy temporalities in CSR communication. Journal of Business Research.',
        link: 'https://www.sciencedirect.com/science/article/pii/S0148296319304370',
        argues: 'Four temporal modes of organisational hypocrisy — aspiration, deferment, evasion and re-narration — each with different consequences.',
        use: 'Gives you precise vocabulary for distinguishing a promise not yet kept from a lie. Directly useful when the internal and external stories diverge.',
        flag: 'Year not confirmed (approximately 2020); check volume and pages.',
      }),
      ...src({
        cite: 'Chen, A., & Eriksson, G. (2019). The making of healthy and moral snacks: A multimodal critical discourse analysis of corporate storytelling. Discourse, Context & Media, 32.   [COURSE LITERATURE]',
        argues: 'Analyses how a company builds a moral, healthy identity through the multimodal design of its storytelling.',
        use: 'Your method exemplar. Göran Eriksson teaches on this module and co-wrote it, which matters if you take this into a thesis in the spring. Shows what close analysis of an authenticity claim looks like.',
      }),
      ...src({
        cite: 'Banet-Weiser, S. (2012). Authentic™: The Politics of Ambivalence in a Brand Culture. New York: NYU Press.   [COURSE LITERATURE]',
        argues: 'Authenticity has itself become a brand strategy. The supposedly non-commercial and the commercial have merged in contemporary brand culture.',
        use: 'The cultural framing for your introduction. Available as an e-book through the university library. Read the introduction and one case chapter; you do not need the whole book.',
      }),

      h2('Strand 4 — Authenticity as organisational control'),
      p('The critical position, and where your central question lives. Here the paradox stops being a problem to solve and becomes the point.'),
      ...src({
        cite: 'Müller, M. (2017). “Brand-centred control”: A study of internal branding and normative control. Organization Studies, 38(7).   [READ FIRST]',
        link: 'https://journals.sagepub.com/doi/10.1177/0170840616663238',
        argues: 'Internal branding extends culture management into a form of normative control. Crucially, brand-centred control recruits an external audience — customers, fans, the public — as an additional source of discipline over employees, blurring work and private life.',
        use: 'The single most important source for your review. It is the empirical demonstration of your driving question: the external claim is produced internally, and the external audience is then turned back on employees. Build your fourth strand around this.',
        flag: 'Confirm page range.',
      }),
      ...src({
        cite: 'Fleming, P., & Sturdy, A. (2009). “Just be yourself!”: Towards neo-normative control in organisations? Employee Relations, 31(6), 569–583.   [READ FIRST]',
        link: 'https://www.emerald.com/insight/content/doi/10.1108/01425450910991730/full/html',
        argues: 'Based on a call-centre study: telling employees to “be themselves” looks liberating but captures their personality, sociality and non-work selves as emotional labour. Control by authenticity rather than despite it.',
        use: 'The sharpest statement of the critical position, and your ethics section writes itself from it. One caution: this sits in an HR journal. Frame it as critical organisation theory, not HR practice, so your review stays inside communication.',
      }),
      ...src({
        cite: 'Alvesson, M., & Einola, K. (2019). Warning for excessive positivity: Authentic leadership and other traps in leadership studies. The Leadership Quarterly, 30(4), 383–395.',
        link: 'https://www.semanticscholar.org/paper/f1f3ef22194d775a3c27e68f699658acb9aeaee8',
        argues: 'Authentic leadership theory rests on shaky philosophical foundations, circular reasoning and weak measurement.',
        use: 'Lets you handle authentic leadership in one honest paragraph and exclude it with justification, rather than pretending the literature does not exist. Connects to Fairhurst & Connaughton on your reading list.',
      }),
      ...src({
        cite: 'Einola, K., & Alvesson, M. (2021). The perils of authentic leadership theory. Leadership.',
        link: 'https://journals.sagepub.com/doi/10.1177/17427150211004059',
        open: 'https://openaccess.city.ac.uk/id/eprint/34150/1/einola-alvesson-2021-the-perils-of-authentic-leadership-theory.pdf',
        argues: 'Extends the critique: the theory is not harmlessly wrong but actively misleading.',
        use: 'Optional. Only if you want a second voice on the leadership exclusion.',
      }),

      // ---------------------------------------------------------------
      h1('3. Recent work in strategic communication'),
      p('These keep the review current and prove the conversation is live in your own field, not only in management and marketing.'),
      ...src({
        cite: 'The authenticity of organizational-level visual identity in the context of strategic communication (2024). International Journal of Strategic Communication.',
        link: 'https://www.tandfonline.com/doi/full/10.1080/1553118X.2024.2352114',
        argues: 'Treats authenticity as central to how organisations construct visual identity.',
        use: 'Evidence that authenticity is an active concern in the flagship strategic communication journal. Good for your introduction.',
        flag: 'Authors not confirmed — get the full citation from the journal page.',
      }),
      ...src({
        cite: 'Strategic authenticity: Signaling authenticity without undermining professional image in workplace interactions. Organization Science.',
        link: 'https://pubsonline.informs.org/doi/10.1287/orsc.2020.14807',
        argues: 'How people signal authenticity at work without damaging their professional image.',
        use: 'Individual-level rather than organisational, so use it carefully — but the title names your exact paradox, and it shows the tension is recognised as real.',
        flag: 'Authors and year not confirmed.',
      }),
      ...src({
        cite: 'Müller, N. (2024). Challenge or resist dominant discourses: Authenticity as a strategic component of activist public relations. Public Relations Inquiry.',
        link: 'https://journals.sagepub.com/doi/10.1177/2046147X241232753',
        argues: 'Authenticity functions as a deliberate strategic resource for activist organisations.',
        use: 'Widens your review beyond corporations and shows authenticity is claimed by organisations with very different aims. Note this is a different Müller from the Organization Studies one — cite carefully.',
      }),

      // ---------------------------------------------------------------
      h1('4. Course literature that carries weight here'),
      p('Using these is required, and they are genuine anchors rather than decoration.'),
      ...src({
        cite: 'Davis, A. (2013). Promotional Cultures: The Rise and Spread of Advertising, Public Relations, Marketing and Branding. Cambridge: Polity.',
        argues: 'Promotional logic has spread from advertising into public relations, politics and public institutions.',
        use: 'Your societal-stakes paragraph, and the justification for why this matters beyond one company. Chapters 1–3 and 10.',
      }),
      ...src({
        cite: 'Gulbrandsen, I. T., & Just, S. (2020). Strategizing Communication: Theory and Practice. Lund: Studentlitteratur.',
        argues: 'The module’s framework for strategic communication as theory and as practice.',
        use: 'Definitional grounding. Chapters 5–8 are the relevant ones for branding and internal communication.',
      }),
      ...src({
        cite: 'Verčič, A. T., & Vokić, N. P. (2017). Engaging employees through internal communication. Public Relations Review, 43(5), 885–893.',
        argues: 'Internal communication drives employee engagement.',
        use: 'Handle with care. This is the closest thing on your list to an HR framing. Use it to establish that internal communication is a recognised sub-field of strategic communication, then move past it — your object is the claim, not the engagement score.',
      }),

      // ---------------------------------------------------------------
      h1('5. Leads worth a search, not yet verified'),
      p('I have not confirmed these against publisher records. Search them; if they hold up they would strengthen the review.'),
      p('•  Alvesson, M., & Willmott, H. (2002). Identity regulation as organizational control. Journal of Management Studies. — the theoretical parent of the Strand 4 argument.', { indent: { left: convertMillimetersToTwip(4) } }),
      p('•  Peterson, R. A. (2005). In search of authenticity. Journal of Management Studies. — authenticity as a claim that is fabricated rather than found.', { indent: { left: convertMillimetersToTwip(4) } }),
      p('•  Beverland, M. B. — on crafting authenticity in luxury brands; several papers, pick one.', { indent: { left: convertMillimetersToTwip(4) } }),
      p('•  Heide, M., & Simonsson, C. — coworkership and internal communication in the Nordic context. Useful and close to home.', { indent: { left: convertMillimetersToTwip(4) } }),
      p('•  Vallaster, C., & de Chernatony, L. — internal brand building and leadership.', { indent: { left: convertMillimetersToTwip(4) } }),

      // ---------------------------------------------------------------
      h1('6. Running the search yourself'),
      p('Your method section has to describe this, so keep a record as you go: which database, which string, how many hits, what you kept and why.'),
      p([{ text: 'Databases: ', bold: true }, { text: 'Communication & Mass Media Complete, Scopus, Web of Science, Business Source Premier. Book a session with Peder Bergenwall, the library’s database search instructor named in the course guide — twenty minutes with him will produce a better method section than anything I can give you.' }]),
      p([{ text: 'Search strings: ', bold: true }, { text: 'authenticity AND (“strategic communication” OR “public relations” OR “corporate communication”) · “brand authenticity” AND (review OR framework) · authenticity AND (“internal branding” OR “brand ambassador”) · “organizational hypocrisy” OR “aspirational talk” · “normative control” AND brand' }]),
      p([{ text: 'Journals to watch: ', bold: true }, { text: 'International Journal of Strategic Communication · Public Relations Review · Public Relations Inquiry · Journal of Communication Management · Corporate Communications: An International Journal · Management Communication Quarterly · Organization · Organization Studies' }]),
      p([{ text: 'Screening rule: ', bold: true }, { text: 'peer-reviewed articles and academic book chapters only. Trade press, consultancy reports and white papers stay out — grade F specifically penalises confusing research with other kinds of text.' }], { after: 200 }),

      h1('7. Adjacent reviews — know these exist'),
      p('These review nearby ground. None does what you are doing, but an examiner may know them, so name the closest one (Christensen & Cornelissen, above) rather than claiming the field is empty.'),
      p('•  A systematic review of internal and external brand management (MDPI Encyclopedia, 2026) — managerial framing, employee brand equity as the outcome.', { indent: { left: convertMillimetersToTwip(4) } }),
      p('•  Recommendations for internal communication to strengthen the employer brand: a systematic literature review (Administrative Sciences, 2023).', { indent: { left: convertMillimetersToTwip(4) } }),
      p('•  Internal branding: conceptualization from a literature review and opportunities for future research (Journal of Brand Management).', { indent: { left: convertMillimetersToTwip(4) } }),
      p('•  An employer branding systematic review covering 145 articles, 2000–2024.', { indent: { left: convertMillimetersToTwip(4) } }),
      p([{ text: 'Two of these are MDPI journals, where quality varies. Check them before citing; the rubric penalises weak sourcing. ', italics: true }], { after: 200 }),

      p([{ text: 'One last thing. ', bold: true }, { text: 'Do not cite anything on this list you have not opened. If a source turns out to say something different from what is written here, trust the source and tell me — the note is my summary, not the paper.' }], { after: 0 }),
    ],
  }],
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync('/home/user/Case-study-analysis/literature-review/Reading_List_Manufacturing_Sincerity.docx', b);
  console.log('written');
});
