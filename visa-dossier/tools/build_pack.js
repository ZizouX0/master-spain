const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Header, Footer,
  AlignmentType, LevelFormat, HeadingLevel, BorderStyle, WidthType, ShadingType,
  PageNumber, PageBreak, TabStopType,
} = require("docx");

const W = 9638; // A4 content width with 2 cm margins
const ACCENT = "1F4E79";

// ---------- helpers ----------
// Text with [blanks] -> blanks rendered as yellow-highlighted runs.
function runs(text, opts = {}) {
  const out = [];
  const re = /(\[[^\]]+\])|(\*\*[^*]+\*\*)/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), ...opts }));
    if (m[1]) out.push(new TextRun({ text: m[1], highlight: "yellow", ...opts }));
    else out.push(new TextRun({ text: m[2].slice(2, -2), ...opts, bold: true }));
    last = m.index + m[0].length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), ...opts }));
  return out;
}
const p = (text, o = {}) => new Paragraph({ children: runs(text, o.run || {}), spacing: { after: o.after ?? 100, line: o.line ?? 259 }, alignment: o.align, indent: o.indent });
const blank = () => new Paragraph({ children: [], spacing: { after: 60 } });
const h1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
const h2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const h3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(t)] });
const bullet = (t, lvl = 0) => new Paragraph({ numbering: { reference: "bullets", level: lvl }, children: runs(t), spacing: { after: 60 } });
const num = (ref, t) => new Paragraph({ numbering: { reference: ref, level: 0 }, children: runs(t), spacing: { after: 60 } });
const check = (t) => new Paragraph({ children: [new TextRun({ text: "☐  ", size: 24 }), ...runs(t)], spacing: { after: 80 }, indent: { left: 360, hanging: 360 } });
const pageBreak = () => new Paragraph({ children: [new PageBreak()] });
const note = (t) => new Paragraph({
  children: runs(t, { size: 20, color: "444444" }),
  shading: { fill: "EEF3F8", type: ShadingType.CLEAR },
  border: { left: { style: BorderStyle.SINGLE, size: 18, color: ACCENT, space: 6 } },
  spacing: { before: 60, after: 160 }, indent: { left: 120, right: 120 },
});

const bd = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const borders = { top: bd, bottom: bd, left: bd, right: bd };
function table(widths, header, rows, opts = {}) {
  const cell = (txt, i, isHead) => new TableCell({
    borders, width: { size: widths[i], type: WidthType.DXA },
    shading: isHead ? { fill: ACCENT, type: ShadingType.CLEAR } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ children: runs(txt, isHead ? { bold: true, color: "FFFFFF", size: 19 } : { size: 19 }), alignment: (opts.center || []).includes(i) ? AlignmentType.CENTER : AlignmentType.LEFT })],
  });
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: header.map((h, i) => cell(h, i, true)) }),
      ...rows.map(r => new TableRow({ children: r.map((c, i) => cell(c, i, false)) })),
    ],
  });
}

// ---------- content ----------
const children = [];
const numS = (ref, t) => new Paragraph({ numbering: { reference: ref, level: 0 }, children: runs(t, { size: 18 }), spacing: { after: 10, line: 240 } });
const trNote = () => note("English translation, for your understanding only. **Do not print or submit this page.** The Spanish page before it is the one you submit.");

// Cover
children.push(
  new Paragraph({ spacing: { before: 1600, after: 200 }, children: [new TextRun({ text: "Spain Study Visa", size: 56, bold: true, color: ACCENT })] }),
  new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "Application pack — Embassy of Spain in Tunis (BLS)", size: 32 })] }),
  new Paragraph({ spacing: { after: 600 }, children: [new TextRun({ text: "Intensive Spanish course at Estudio Sampere → Master’s degree in Spain, September 2027", size: 24, italics: true, color: "555555" })] }),
  p("Applicant: [Nom et prénom]"),
  p("Version 2 — 26 September 2026 — corrected after a four-reviewer mock consular review"),
  blank(),
  h3("How to use this pack"),
  bullet("Text highlighted in [yellow] is a personal fact only you can supply. Replace every highlighted blank, then remove the highlight."),
  bullet("Every fact must be true and backed by a document in your file. The consulate checks bank statements, school letters and employer letters, and a false statement leads to refusal (art. 54.8.b)."),
  bullet("Each Spanish text is followed by an **English translation page**. That page is for your understanding only: don’t print or submit it."),
  bullet("Print each letter on its own page. Before printing, delete the section title and the blue guidance box above it. Sign in blue ink and date it."),
  bullet("Part 1 explains the file and the review results. Part 2 lists what only you can get. Parts 3 to 11 are what you prepare and submit."),
  pageBreak(),
);

// Contents
children.push(
  h1("Contents"),
  num("toc", "Analysis and mock consular review results"),
  num("toc", "Your to-do list: what only you can get"),
  num("toc", "Final checklist, in the order BLS wants it"),
  num("toc", "Your stay dates: the 45 extra days, money, insurance and TIE"),
  num("toc", "Motivation letter (carta de motivación, Spanish) + English translation"),
  num("toc", "Cover letter and document index (carta de presentación, Spanish) + English translation"),
  num("toc", "Parent’s sponsorship letter (attestation de prise en charge, French)"),
  num("toc", "Text for the doctor: medical certificate (French)"),
  num("toc", "Email to Estudio Sampere: requesting the visa letter (Spanish) + English translation"),
  num("toc", "Email requesting a national-visa appointment (French)"),
  num("toc", "Interview preparation, and the final check before you submit"),
  pageBreak(),
);

// Part 1 — Analysis
children.push(
  h1("1. Analysis and review results"),
  p("This pack is built from four sources: the embassy’s official requirements page, an agency checklist for Tunis, two posts from the “BLS visa espagne TUNISIE” group (June–July 2026), and Spain’s immigration regulation (Real Decreto 1155/2024). Version 2 then went through a mock review by four simulated reviewers: a visa officer, a legal and BLS-counter checker, a native Spanish and French reviewer, and a skeptical “find every reason to refuse” reviewer. They are simulations, not embassy staff, and their figures (income bands, 20 hours a week, exam result delays) are estimates rather than published rules."),
  h2("1.1 What the mock review found"),
  table([3000, 6638], ["Reviewer", "Verdict on version 1"], [
    ["**Visa officer**", "More documents requested, leaning towards refusal. Reason: “finalidad no acreditada” (the purpose isn’t proven). Scores out of 5: documents 4, money 3, letters 3, believable purpose 2, coherent plan 2."],
    ["**Skeptical reviewer**", "Medium refusal risk. Medium-high if your level is A2, you’re unemployed, or there’s a recent large deposit."],
    ["**Legal and BLS check**", "Mostly sound, but two legal errors and six must-fix points."],
    ["**Language reviewer**", "The plan convinces, but the motivation letter was too long and read like a template."],
  ]),
  blank(),
  p("**Everything they asked to fix is corrected in this version.** The single biggest risk that remains is one only you can remove: **proof of the master’s** and **a certificate of your current Spanish level** (Part 2)."),
  h2("1.2 What every approved file has in common"),
  table([2600, 7038], ["Pattern", "What it means for you"], [
    ["**Complete and in order**", "BLS checks your file against its own sheet at the counter. Most setbacks come from a missing document or translation, not from the applicant’s profile. Stack each item: original → copy → Spanish translation → copy of the translation."],
    ["**Apostille + sworn translation**", "Every Tunisian public document is apostilled by an authorised notary, then translated by a sworn translator, **apostille included**. If the translator isn’t a sworn translator accredited by Spain’s foreign ministry (MAEC), the translation needs its own apostille too. **Every document not in Spanish is translated**, including your CIN and bank statements."],
    ["**Money blocked at the bank**", "A student file at the bank (dossier scolaire) with authorisation to transfer to Spain, and a certificate stating the money is irrevocably blocked. The agency uses €700 a month (the legal minimum is €600), counted over your **whole stay** plus the trip home. Only you or your parents can be sponsors."],
    ["**Recent documents**", "Medical certificate, birth certificate, B3, payslips, attestation de travail and blocking certificate: all less than 3 months old on the day of your appointment."],
    ["**Paid studies**", "Fees paid in full, with a letter from the school confirming it received the money. An invoice alone isn’t enough."],
    ["**Short supporting documents**", "A short motivation letter, a CV and a language certificate."],
    ["**Interview**", "The group says interviews are rare, but the officer warned that language-course files do get called in. Prepare Part 11."],
  ]),
  blank(),
  h2("1.3 How your case differs from the posts"),
  p("The posts and the agency checklist were written for **university** admissions. You are applying for a **language school**, which the regulation treats as a separate category: “actividad formativa”, art. 52.1.e.2º. Four consequences:"),
  num("diff", "**The school must be accredited by Instituto Cervantes and the course must be in person.** The school’s letter must say so, with the hours per week in 60-minute hours (Part 9)."),
  num("diff", "**The officer will ask: “Is the language course the real purpose?”** The answer is a chain you can prove: certified current level → intensive course → SIELE/DELE → named master’s that requires B2 → September 2027. Proof of the master’s is what turns this chain from a promise into a fact."),
  num("diff", "**The permit lasts one year at most, allows only one extension, and gives no automatic right to work** (art. 55.1, 55.5, 57)."),
  num("diff", "**Switching to the master’s from inside Spain is legally possible** (art. 54.1 ¶2 and 54.3). Once admitted, you file a new higher-education application at the immigration office (oficina de extranjería), or the university files it for you (art. 54.6). Deadline: at least 2 months before your current permit expires **and** at least 2 months before the master’s starts. The alternative is to return to Tunisia and apply at the embassy. Your letter states the in-Spain route and a fallback. **Version 1 of this pack wrongly said the in-Spain switch was impossible.**"),
  h2("1.4 Why files get refused, and how this pack prevents it"),
  table([4200, 5438], ["Reason for refusal", "Prevented by"], [
    ["Purpose not proven (the officer’s #1 reason)", "Proof of the master’s and a certified Spanish level (Part 2); motivation letter built on them (Part 5)"],
    ["Missing apostille or translation", "Checklist (Part 3) with a translation and an apostille column for every item"],
    ["Too little money, or an unexplained deposit", "€700 × months of the whole stay + trip home (Part 4); 6 months of statements; proof of source for any large deposit"],
    ["Insurance not accepted", "Health insurance from an insurer authorised in Spain, with SNS-equivalent cover, no co-payment and no waiting period, covering the whole stay"],
    ["Medical certificate invalid", "Exact wording (Part 8), legalised by the regional health directorate, less than 3 months old"],
    ["School not eligible", "Visa letter with Cervantes accreditation, in-person attendance and hours per week (Part 9)"],
    ["Inconsistent dates", "One rule: course dates on the school letter; stay dates everywhere else (Part 4)"],
    ["Late application", "Submit at least 2 months before the course starts"],
  ]),
  blank(),
  h2("1.5 Where the sources disagree, and what to do"),
  bullet("**Insurance.** The embassy requires health insurance from an insurer authorised in Spain, with cover equivalent to Spain’s public health system (SNS). The agency offers travel insurance (about 500 TND). **Follow the embassy.**"),
  bullet("**Amount.** Legal minimum €600 a month; the agency uses €700. **Block €700 a month over the whole stay**, plus the trip home."),
  bullet("**Accommodation.** Not required by law. The agency asks for 3 months; two reviewers found 1 month thin (“where will you live afterwards?”). **Book at least 3 months**, or prepay the whole stay, which also reduces the money you must block. A cheaper city (Salamanca, Alicante) helps."),
  bullet("**Flight.** Not required, but your money must cover the trip home. **Add a return reservation** (not a purchased ticket) that matches your stay dates."),
  bullet("**Bank statements.** The agency says no translation; the official rule says every document not in Spanish must be translated. **Translate them.**"),
  bullet("**Appointment.** Confirmed on the embassy’s contact page: national-visa appointments are requested at **emb.tunez.cit@maec.es** (Part 10)."),
  pageBreak(),
);

// Part 2 — To-do (things only you can get)
children.push(
  h1("2. Your to-do list: what only you can get"),
  note("The reviewers agreed: these items decide whether the officer believes your plan. Nothing in the letters can replace them."),
  table([500, 4900, 4238], ["#", "What to get", "Why it matters"], [
    ["1", "**Proof of the master’s**: a pre-registration receipt (justificante de preinscripción), a conditional admission (“admisión condicionada a acreditar B2”), or at least an email from the programme coordinator confirming you’re eligible and that B2 is required. Apply in the first admission rounds (often November to February).", "The officer’s #1 reason for refusal. Printed web pages alone don’t prove anything."],
    ["2", "**A certificate of your current Spanish level**: Instituto Cervantes de Túnez, a SIELE result, or Sampere’s graded placement test.", "A self-declared level counts as none. Rough guide from the officer: certified A2 fits 5–6 months; B1 fits 3–4 months; A0/A1 or an existing B2 is a red flag."],
    ["3", "**Choose the campus**: Madrid, Salamanca or Alicante. Ideally the city of your target university, or a cheaper one.", "An undecided campus makes the plan look unfinished."],
    ["4", "**The school’s visa letter** with hours per week in **60-minute hours** (20 or more), in person, full-time, Cervantes-accredited, paid in full (use Part 9).", "“20 lessons” of 50 minutes is only about 17 hours, which the officer would question."],
    ["5", "**Course dates**: plan the course so your master’s application fits. In-Spain route: file at least 2 months before your permit ends. Return route: finish by mid-May so you can apply in Tunis in time.", "A July return with new medical, B3, payslips and blocking during peak season is risky."],
    ["6", "**Language exam**: register for the **SIELE** (results within about 3 weeks) or an early DELE session.", "DELE results take about 3 months, too late for most admissions. Registration also qualifies you for the one allowed extension (art. 55.5)."],
    ["7", "**Money proof**: the parent’s income documents, 6 months of stamped statements, proof of source for any large deposit (sale deed, loan contract), the bank’s transfer authorisation, and ideally evidence the family can also pay for the master’s year.", "The officer checks the sponsor’s capacity, not just the balance."],
    ["8", "**An answer to “Why not the Instituto Cervantes in Tunis?”**, based on true facts (hours per week, immersion, a certificate of hours already done in Tunis if you have one).", "The officer will ask it. The letter (Part 5) and interview (Part 11) use your answer."],
  ]),
  pageBreak(),
);

// Part 3 — Checklist
const Y = "Oui", N = "Non", R = "Conseillé", S = "Si pas en ES";
children.push(
  h1("3. Final checklist (in BLS order)"),
  p("**Trad.** = sworn Spanish translation (it must include the apostille). **Apost.** = apostille from an authorised notary, on the original first. “Conseillé” = recommended. “Si pas en ES” = translate only if not already in Spanish. For each item bring: original → copy → translation → copy of the translation."),
  table([500, 5738, 1000, 1000, 700, 700], ["#", "Document", "Trad.", "Apost.", "Prêt", "Copie"], [
    ["1", "Formulaire de visa national, tapé, signé à l’encre bleue et daté ; dates d’entrée et de séjour = vos dates de séjour (Partie 4)", N, N, "☐", "☐"],
    ["2", "Photo 3,5 × 4,5 cm, fond blanc, moins de 6 mois, collée sur le formulaire (+ une photo de réserve)", N, N, "☐", "☐"],
    ["3", "Passeport délivré il y a **moins de 10 ans**, valide au moins 1 an, 2 pages vierges + copie des 5 premières pages + anciens passeports et visas", N, N, "☐", "☐"],
    ["4", "Extrait de naissance en français, de moins de 3 mois", Y, Y, "☐", "☐"],
    ["5", "CIN (preuve de résidence en Tunisie)", Y, N, "☐", "☐"],
    ["6", "Lettre d’admission Estudio Sampere : présentiel, temps plein, heures de 60 min par semaine (≥ 20), accréditation Instituto Cervantes (Partie 9) + page imprimée de l’accréditation", N, N, "☐", "☐"],
    ["7", "Preuve de paiement intégral : SWIFT + lettre de l’école confirmant la réception", S, N, "☐", "☐"],
    ["8", "Attestation de non-boursier", Y, R, "☐", "☐"],
    ["9", "Dossier scolaire bancaire avec autorisation de transfert + attestation de blocage irrévocable : ≥ 700 € × mois de séjour (arrivée → fin du cours + 15 j) + retour + frais non payés ; moins de 3 mois", Y, Y, "☐", "☐"],
    ["10", "Attestation de prise en charge du/des parent(s), signature légalisée à la municipalité (Partie 7)", Y, Y, "☐", "☐"],
    ["11", "CIN ou passeport du parent (copie)", Y, N, "☐", "☐"],
    ["12", "Revenus du parent (moins de 3 mois) : salarié = 3 fiches de paie + attestation de travail ; commerçant = RNE + patente + déclaration d’ouverture ; retraité = pension + historique CNRPS/CNSS", Y, R, "☐", "☐"],
    ["13", "Relevés bancaires du parent, 6 mois, cachetés par la banque (+ justificatif de l’origine de tout gros dépôt)", Y, N, "☐", "☐"],
    ["14", "Assurance santé : assureur autorisé en Espagne, couverture équivalente au SNS (cartera común básica), sans copago ni carencia, de l’arrivée à la fin du cours + 15 j", S, N, "☐", "☐"],
    ["15", "Certificat médical de moins de 3 mois, légalisé par la Direction Régionale de la Santé (Partie 8)", Y, Y, "☐", "☐"],
    ["16", "Bulletin n°3 (B3) : **obligatoire** (séjour > 6 mois) ; un pour chaque pays de résidence des 5 dernières années ; moins de 3 mois ; demande en ligne à lancer tôt", Y, Y, "☐", "☐"],
    ["17", "Hébergement (recommandé) : au moins 3 mois, ou tout le séjour payé d’avance (réduit le montant à bloquer) + preuve de paiement ; entrée = date d’arrivée", S, N, "☐", "☐"],
    ["18", "Réservation de vol aller-retour (recommandé ; ne pas acheter avant le visa) aux dates du séjour", N, N, "☐", "☐"],
    ["19", "Bac (diplôme + relevé) et diplôme universitaire (ou attestation de réussite) + TOUS les relevés de notes", Y, Y, "☐", "☐"],
    ["20", "Attestations de stage / de travail : papier à en-tête, signées et cachetées, avec dates et missions", Y, N, "☐", "☐"],
    ["21", "**Preuve du master** : préinscription, admission conditionnelle ou e-mail du coordinateur + pages du programme (niveau exigé, calendrier)", S, N, "☐", "☐"],
    ["22", "**Certificat du niveau actuel d’espagnol** (Cervantes Tunis, SIELE ou test de placement Sampere) + inscription SIELE/DELE", S, N, "☐", "☐"],
    ["23", "Lettre de motivation en espagnol (Partie 5)", N, N, "☐", "☐"],
    ["24", "CV en espagnol (format Europass)", N, N, "☐", "☐"],
    ["25", "Frais : ≈ 372 TND au total, frais BLS inclus, en espèces (à vérifier sur la grille du trimestre en cours)", "—", "—", "☐", "—"],
  ], { center: [0, 2, 3, 4, 5] }),
  blank(),
  note("Timing: submit at least 2 months before the course starts (course on 1 February 2027 → by 1 December 2026). Start with the B3, the bank file and the master’s pre-registration: they take the longest. Request the appointment early (Part 10)."),
  pageBreak(),
);

// Part 4 — Stay dates
children.push(
  h1("4. Your stay dates: the 45 extra days"),
  p("By law, your permit **starts 1 month before the course** and **ends 15 days after it** (art. 55.2): “La vigencia de la autorización deberá comenzar con una antelación de un mes con respecto al comienzo de la actividad […]. La vigencia se extenderá quince días más allá de la finalización de la actividad.” You don’t need a special form, but you should request it clearly so the visa is issued with the right dates."),
  h2("4.1 How to ask for them"),
  num("idx4", "**Apply at least 2 months before the course starts.** If you apply later, the permit may start on the day of approval instead of 1 month before."),
  num("idx4", "**On the visa form**, enter your intended arrival date (up to 1 month before the course) and your intended length of stay (until the course end + 15 days)."),
  num("idx4", "**In the cover letter**, one sentence requests it, with a reason (already written in Part 6): “Les agradecería que la vigencia del visado abarcara del [fecha de llegada] al [fecha de fin del curso + 15 días]…”"),
  num("idx4", "**Make every other document cover the same period**: insurance, blocked money, accommodation (from arrival), flight reservation, sponsorship letter."),
  num("idx4", "**When you collect the passport, check the visa sticker** (“desde / hasta”). If the dates are shorter, say so at the counter immediately."),
  h2("4.2 The one date rule"),
  table([3200, 6438], ["Which dates", "Where they go"], [
    ["**Course dates** (first class → last class)", "The school’s letter, the payment confirmation, the motivation letter"],
    ["**Stay dates** (arrival → course end + 15 days)", "The visa form, insurance, blocked-money calculation, accommodation, flight reservation, sponsorship letter, cover letter request"],
  ]),
  blank(),
  h2("4.3 Worked example"),
  table([3600, 6038], ["Item", "Example (change it to your real dates)"], [
    ["Course", "1 February – 25 June 2027"],
    ["Stay (permit)", "1 January – 10 July 2027 (about 6.3 months)"],
    ["Blocked money", "7 months × €700 = **€4,900**, plus the return trip and any unpaid fees"],
    ["Insurance", "1 January – 10 July 2027"],
    ["Submit the visa file", "By **1 December 2026** at the latest"],
    ["TIE (foreigner ID card)", "Required, because the stay is over 6 months: apply at the police within 1 month of arriving (art. 54.9)"],
    ["Master’s, in-Spain route", "File at the oficina de extranjería by **10 May 2027** (2 months before the permit ends), with admission and fees paid"],
    ["Master’s, return route", "Back in Tunis and file at the embassy at least 2 months before September. A course ending by mid-May makes this realistic."],
  ]),
  note("Trade-off: the 45 days help you settle in and sit the exam, but mean more money blocked, longer insurance, and a TIE if the stay exceeds 6 months."),
  pageBreak(),
);

// Letter header helper
const letterHead = (extra) => [
  p("[Nombre y apellidos]", { after: 0 }),
  p("[Dirección], [Código postal] [Ciudad], Túnez", { after: 0 }),
  p("Pasaporte n.º [número] · Tel.: [+216 …] · [correo electrónico]", { after: 160 }),
  p("Sección Consular", { after: 0 }),
  p("Embajada de España en Túnez", { after: 160 }),
  p("[Ciudad], a [día] de [mes] de 2026", { align: AlignmentType.RIGHT, after: 160 }),
  ...extra,
];
const letterHeadEn = (extra) => [
  p("[Full name]", { after: 0 }),
  p("[Address], [Postcode] [City], Tunisia", { after: 0 }),
  p("Passport no. [number] · Tel.: [+216 …] · [email]", { after: 160 }),
  p("Consular Section", { after: 0 }),
  p("Embassy of Spain in Tunis", { after: 160 }),
  p("[City], [day] [month] 2026", { align: AlignmentType.RIGHT, after: 160 }),
  ...extra,
];
const J = AlignmentType.JUSTIFIED;

// Part 5 — Carta de motivación
children.push(
  h1("5. Carta de motivación"),
  note("About 280 words. Replace every blank with a true fact; delete any sentence you can’t prove. If your Spanish is below B1, be ready to discuss the letter in French at an interview, or submit a French version with a sworn translation."),
  ...letterHead([p("**Asunto: Carta de motivación – Visado nacional de estudios (curso intensivo de español)**", { after: 160 })]),
  p("Señoras y señores:"),
  p("Tengo [edad] años, resido en [ciudad] y solicito un visado nacional de estudios para seguir un curso intensivo y presencial de lengua española en Estudio Sampere [ciudad], centro acreditado por el Instituto Cervantes, del [fecha] al [fecha], con [N] horas semanales.", { align: J }),
  p("En [año] obtuve el título de [titulación] por la Universidad de [nombre] (nota media: [nota]/20). [Desde (año) trabajo como (puesto) en (empresa) / En (año) realicé unas prácticas en (empresa)], experiencia que ha reforzado mi interés por especializarme en [área].", { align: J }),
  p("Mi objetivo es cursar, a partir de septiembre de 2027, el [Máster en … de la Universidad de …] ([preinscripción n.º … / admisión condicionada]). El programa se imparte en español y exige acreditar un nivel [B2]; mi nivel actual es [B1] ([certificado]). [He seguido (N) horas de clase en el Instituto Cervantes de Túnez;] para alcanzar el B2 a tiempo necesito un curso de [N] horas semanales en inmersión, [que no encuentro en Túnez]. He elegido Estudio Sampere [ciudad] por [motivo concreto] y por su preparación para los exámenes oficiales.", { align: J }),
  p("Calendario previsto:", { after: 40 }),
  bullet("[fecha] – [fecha]: curso intensivo."),
  bullet("[mes] de 2027: examen SIELE [o DELE B2]."),
  bullet("[meses] de 2027: admisión definitiva en el máster."),
  bullet("Antes del [fecha]: solicitud de la autorización de estancia por estudios del máster, conforme a la normativa vigente."),
  bullet("Septiembre de 2027: inicio del máster."),
  p("Si no obtuviera la admisión, regresaré a Túnez al terminar el curso, el [fecha], y [plan concreto].", { align: J }),
  p("Mi [padre/madre], [nombre], [profesión] en [empleador], sufragará todos mis gastos. He bloqueado de forma irrevocable [importe] € en [banco], que cubren toda mi estancia y el regreso. El curso está íntegramente abonado y el alojamiento, pagado hasta el [fecha].", { align: J }),
  p("Tras el máster, mi intención es [plan profesional concreto] en Túnez.", { align: J }),
  p("Quedo a su disposición para facilitarles cualquier documentación adicional."),
  p("Atentamente,", { after: 360 }),
  p("[Firma]", { after: 0 }),
  p("[Nombre y apellidos]"),
  pageBreak(),
  h2("5b. Motivation letter: English translation"),
  trNote(),
  ...letterHeadEn([p("**Subject: Motivation letter – National study visa (intensive Spanish course)**", { after: 160 })]),
  p("Dear Sir or Madam,"),
  p("I am [age] years old, I live in [city], and I am applying for a national study visa to attend an intensive, in-person Spanish language course at Estudio Sampere [city], a centre accredited by Instituto Cervantes, from [date] to [date], with [N] hours per week.", { align: J }),
  p("In [year] I obtained a [degree] from the University of [name] (average grade: [grade]/20). [Since (year) I have worked as (position) at (company) / In (year) I did an internship at (company)], an experience that strengthened my interest in specialising in [field].", { align: J }),
  p("My goal is to study the [Master’s in … at the University of …] from September 2027 ([pre-registration no. … / conditional admission]). The programme is taught in Spanish and requires a certified [B2] level; my current level is [B1] ([certificate]). [I have completed (N) hours of classes at the Instituto Cervantes in Tunis;] to reach B2 in time I need a course of [N] hours per week in an immersion setting, [which I cannot find in Tunisia]. I chose Estudio Sampere [city] for [specific reason] and for its preparation for the official exams.", { align: J }),
  p("Planned timeline:", { after: 40 }),
  bullet("[date] – [date]: intensive course."),
  bullet("[month] 2027: SIELE exam [or DELE B2]."),
  bullet("[months] 2027: final admission to the master’s."),
  bullet("Before [date]: application for the study-stay permit for the master’s, in accordance with the regulations in force."),
  bullet("September 2027: start of the master’s."),
  p("If I am not admitted, I will return to Tunisia at the end of the course, on [date], and [concrete plan].", { align: J }),
  p("My [father/mother], [name], [profession] at [employer], will cover all my expenses. I have irrevocably blocked €[amount] at [bank], which covers my entire stay and my return. The course is fully paid and my accommodation is paid until [date].", { align: J }),
  p("After the master’s, I intend to [concrete career plan] in Tunisia.", { align: J }),
  p("I remain at your disposal to provide any additional documentation."),
  p("Yours faithfully,", { after: 360 }),
  p("[Signature]", { after: 0 }),
  p("[Full name]"),
  pageBreak(),
);

// Part 6 — Carta de presentación
const idxES_req = [
  "Formulario de solicitud de visado nacional, cumplimentado y firmado, con fotografía.",
  "Pasaporte, copia de las cinco primeras páginas, y pasaportes y visados anteriores.",
  "Certificado de nacimiento, apostillado y con traducción jurada.",
  "Documento de identidad tunecino (CIN), con traducción jurada, como justificante de residencia.",
  "Carta de admisión de Estudio Sampere y justificante del pago íntegro del curso.",
  "Certificado de no becario, con traducción jurada.",
  "Expediente escolar bancario y certificado de bloqueo irrevocable de [importe] €, apostillado y con traducción jurada.",
  "Carta de compromiso de manutención de mi [padre/madre], con firma legalizada, apostillada y con traducción jurada; copia de su CIN; [nóminas y certificado de trabajo] y extractos bancarios de los últimos 6 meses.",
  "Seguro de enfermedad con [aseguradora], póliza n.º [número], válido del [fecha] al [fecha].",
  "Certificado médico, legalizado, apostillado y con traducción jurada.",
  "Certificado de antecedentes penales (Bulletin n.º 3), apostillado y con traducción jurada.",
  "Justificante del abono de la tasa de visado.",
];
const idxES_comp = [
  "Carta de motivación.",
  "Justificante de preinscripción en el [Máster en …] [o correo del coordinador del programa].",
  "Certificado de mi nivel actual de español e inscripción en el examen [SIELE/DELE].",
  "Títulos y expedientes académicos, apostillados y con traducción jurada.",
  "Certificados de prácticas y de trabajo, con traducción jurada.",
  "Reserva y justificante de pago del alojamiento.",
  "Reserva de vuelo de ida y vuelta.",
  "Currículum vítae.",
];
const idxEN_req = [
  "National visa application form, completed and signed, with photograph.",
  "Passport, copy of the first five pages, and previous passports and visas.",
  "Birth certificate, apostilled and with sworn translation.",
  "Tunisian identity card (CIN), with sworn translation, as proof of residence.",
  "Admission letter from Estudio Sampere and proof of full payment of the course.",
  "Certificate of non-scholarship, with sworn translation.",
  "Bank student file and certificate of an irrevocable block of €[amount], apostilled and with sworn translation.",
  "Letter of financial support from my [father/mother], with legalised signature, apostilled and with sworn translation; copy of their CIN; [payslips and employment certificate] and bank statements for the last 6 months.",
  "Health insurance with [insurer], policy no. [number], valid from [date] to [date].",
  "Medical certificate, legalised, apostilled and with sworn translation.",
  "Criminal record certificate (Bulletin no. 3), apostilled and with sworn translation.",
  "Proof of payment of the visa fee.",
];
const idxEN_comp = [
  "Motivation letter.",
  "Proof of pre-registration for the [Master’s in …] [or email from the programme coordinator].",
  "Certificate of my current Spanish level and registration for the [SIELE/DELE] exam.",
  "Diplomas and academic transcripts, apostilled and with sworn translation.",
  "Internship and work certificates, with sworn translation.",
  "Accommodation booking and proof of payment.",
  "Return flight reservation.",
  "Curriculum vitae.",
];
children.push(
  h1("6. Carta de presentación e índice"),
  note("Put this letter at the front of the file and arrange the documents in the same order as the index. The last paragraph requests the 45 extra days (Part 4)."),
  ...letterHead([p("**Asunto: Solicitud de visado nacional de estudios – Curso intensivo de lengua española en Estudio Sampere ([ciudad]), del [fecha] al [fecha]**", { after: 160 })]),
  p("Señoras y señores:"),
  p("Presento mi solicitud de visado nacional de estudios para seguir un curso intensivo y presencial de lengua española en Estudio Sampere ([ciudad]), centro acreditado por el Instituto Cervantes, del [fecha de inicio] al [fecha de fin]. Adjunto la documentación, ordenada según el siguiente índice:", { align: J }),
  p("**Documentación exigida**", { after: 40 }),
  ...idxES_req.map(t => numS("idx", t)),
  p("**Documentación complementaria**", { after: 40 }),
  ...idxES_comp.map(t => numS("idx", t)),
  p("Les agradecería que la vigencia del visado abarcara del [fecha de llegada] al [fecha de fin del curso + 15 días], para instalarme antes del inicio de las clases y presentarme al examen [SIELE] al finalizar el curso.", { align: J }),
  p("Quedo a su disposición para cualquier aclaración o documentación adicional."),
  p("Atentamente,", { after: 240 }),
  p("[Firma]", { after: 0 }),
  p("[Nombre y apellidos]"),
  pageBreak(),
  h2("6b. Cover letter: English translation"),
  trNote(),
  ...letterHeadEn([p("**Subject: National study visa application – Intensive Spanish language course at Estudio Sampere ([city]), from [date] to [date]**", { after: 160 })]),
  p("Dear Sir or Madam,"),
  p("I hereby submit my application for a national study visa to attend an intensive, in-person Spanish language course at Estudio Sampere ([city]), a centre accredited by Instituto Cervantes, from [start date] to [end date]. I enclose the documents, arranged according to the following index:", { align: J }),
  p("**Required documents**", { after: 40 }),
  ...idxEN_req.map(t => numS("idxen", t)),
  p("**Supporting documents**", { after: 40 }),
  ...idxEN_comp.map(t => numS("idxen", t)),
  p("I would be grateful if the visa could be valid from [arrival date] to [course end date + 15 days], so that I can settle in before classes begin and sit the [SIELE] exam at the end of the course.", { align: J }),
  p("I remain at your disposal for any clarification or additional documentation."),
  p("Yours faithfully,", { after: 240 }),
  p("[Signature]", { after: 0 }),
  p("[Full name]"),
  pageBreak(),
);

// Part 7 — Prise en charge
children.push(
  h1("7. Attestation de prise en charge"),
  note("Follows the official BLS Tunisia model, plus two sentences that make the file easier to read. Your parent signs in front of the municipality (légalisation de signature); then apostille by a notary and sworn Spanish translation. The dates are your **stay dates** (Part 4)."),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 240, after: 360 }, children: [new TextRun({ text: "ATTESTATION DE PRISE EN CHARGE", bold: true, size: 30, underline: {} })] }),
  p("Je soussigné(e) [Nom et prénom du parent],", { after: 160 }),
  p("en qualité de [père / mère],", { after: 160 }),
  p("titulaire de la CIN n° [numéro], délivrée le [date] à [lieu],", { after: 160 }),
  p("N° de téléphone : [numéro],", { after: 240 }),
  p("m’engage à subvenir à toutes les dépenses relatives à l’hébergement, aux transports et aux frais d’hospitalisation ou soins médicaux de :", { align: J, after: 240 }),
  p("mon fils / ma fille [Nom et prénom de l’étudiant(e)], né(e) le [date],", { after: 160 }),
  p("titulaire du passeport tunisien sous le N° [numéro],", { after: 240 }),
  p("pendant toute la durée de son séjour en Espagne du [date d’arrivée] au [date de fin du cours + 15 jours], sans avoir recours aux aides publiques.", { align: J, after: 240 }),
  p("Cette prise en charge, d’un montant minimum de 700 € par mois, s’inscrit dans le cadre de son cours intensif de langue espagnole à l’école Estudio Sampere, [ville] (Espagne).", { align: J, after: 240 }),
  p("Je déclare exercer la profession de [profession] au sein de [employeur] depuis le [date] et percevoir un revenu mensuel net de [montant] TND.", { align: J, after: 480 }),
  p("Fait à [ville], le [date]", { align: AlignmentType.RIGHT, after: 720 }),
  p("Signature (à légaliser auprès de la Municipalité)"),
  pageBreak(),
);

// Part 8 — Medical certificate
children.push(
  h1("8. Certificat médical: text to give your doctor"),
  note("Issued less than 3 months before your appointment. From a private doctor it must be legalised by the Direction Régionale de la Santé of the doctor’s region (or get it from a public health centre). Then apostille and sworn Spanish translation. If BLS or your translator gives you their own model, use theirs."),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 240, after: 360 }, children: [new TextRun({ text: "CERTIFICAT MÉDICAL", bold: true, size: 30 })] }),
  p("Je soussigné(e), Docteur [Nom du médecin], [spécialité], exerçant à [adresse du cabinet],", { align: J, after: 200 }),
  p("certifie avoir examiné ce jour [M. / Mme] [Nom et prénom], né(e) le [date de naissance] à [lieu de naissance], titulaire du passeport n° [numéro],", { align: J, after: 200 }),
  p("et atteste que l’intéressé(e) ne souffre d’aucune des maladies pouvant avoir des répercussions graves sur la santé publique, conformément aux dispositions du Règlement Sanitaire International de 2005.", { align: J, after: 200, run: { bold: true } }),
  p("Certificat délivré à l’intéressé(e) pour servir et valoir ce que de droit.", { after: 360 }),
  p("Fait à [ville], le [date]", { align: AlignmentType.RIGHT, after: 480 }),
  p("Signature et cachet du médecin"),
  pageBreak(),
);

// Part 9 — Email to Sampere
const sampereES = [
  "Mi nombre completo, fecha de nacimiento y número de pasaporte ([número]).",
  "Las fechas exactas de inicio y fin del curso.",
  "Que el curso es intensivo, a tiempo completo y presencial, con el número de horas lectivas semanales expresado en horas de 60 minutos.",
  "Que Estudio Sampere [ciudad] es un centro acreditado por el Instituto Cervantes.",
  "Que el curso está pagado en su totalidad (importe y fecha de recepción).",
  "En su caso, los datos del alojamiento reservado a través de la escuela (dirección, fechas y confirmación de pago).",
  "Si procede, que el programa incluye preparación para los exámenes SIELE o DELE.",
];
const sampereEN = [
  "My full name, date of birth and passport number ([number]).",
  "The exact start and end dates of the course.",
  "That the course is intensive, full-time and in person, with the number of teaching hours per week expressed in 60-minute hours.",
  "That Estudio Sampere [city] is a centre accredited by Instituto Cervantes.",
  "That the course has been paid in full (amount and date received).",
  "Where applicable, the details of accommodation booked through the school (address, dates and payment confirmation).",
  "If applicable, that the programme includes preparation for the SIELE or DELE exams.",
];
children.push(
  h1("9. Email to Estudio Sampere: visa letter"),
  note("Send this after you have paid. The school’s letter proves you fit the legal category, so every point matters, especially the hours in 60-minute hours."),
  p("**Asunto:** Carta de admisión para visado de estudios – [Nombre y apellidos] – Curso intensivo [fecha de inicio]"),
  p("Estimado equipo de Estudio Sampere:"),
  p("He realizado el pago del curso intensivo de español en su centro de [ciudad] del [fecha de inicio] al [fecha de fin] (referencia de pago: [referencia]). Para solicitar mi visado nacional de estudios en la Embajada de España en Túnez, les agradecería que me enviaran una carta de admisión firmada y sellada en la que conste:", { align: J }),
  ...sampereES.map(t => bullet(t)),
  p("Muchas gracias de antemano."),
  p("Un cordial saludo,", { after: 120 }),
  p("[Nombre y apellidos] · [teléfono] · [correo electrónico]"),
  pageBreak(),
  h2("9b. Email to Estudio Sampere: English translation"),
  trNote(),
  p("**Subject:** Admission letter for a study visa – [Full name] – Intensive course [start date]"),
  p("Dear Estudio Sampere team,"),
  p("I have paid for the intensive Spanish course at your [city] centre from [start date] to [end date] (payment reference: [reference]). To apply for my national study visa at the Embassy of Spain in Tunis, I would be grateful if you could send me a signed and stamped admission letter stating:", { align: J }),
  ...sampereEN.map(t => bullet(t)),
  p("Many thanks in advance."),
  p("Kind regards,", { after: 120 }),
  p("[Full name] · [phone] · [email]"),
  pageBreak(),
);

// Part 10 — Appointment email
children.push(
  h1("10. Email requesting an appointment"),
  note("Send to **emb.tunez.cit@maec.es**, the address the embassy’s contact page lists for national-visa appointment requests. Send it when your file is about 15 days from ready, and state the date it will be complete."),
  p("**À :** emb.tunez.cit@maec.es"),
  p("**Objet :** Demande de rendez-vous – Visa national d’études – [Nom et prénom] – Passeport n° [numéro]"),
  p("Madame, Monsieur,"),
  p("Titulaire du passeport n° [numéro], je suis inscrit(e) à un cours intensif de langue espagnole à l’école Estudio Sampere de [ville], centre accrédité par l’Instituto Cervantes, du [date] au [date].", { align: J }),
  p("Je souhaiterais obtenir un rendez-vous afin de déposer ma demande de visa national d’études. Mon dossier sera complet, traductions et apostilles comprises, à partir du [date].", { align: J }),
  p("Vous trouverez ci-joint la copie de mon passeport et ma lettre d’admission."),
  p("Je vous prie d’agréer, Madame, Monsieur, l’expression de mes salutations distinguées.", { after: 160 }),
  p("[Nom et prénom] · [téléphone] · [e-mail]"),
  p("**Pièces jointes :** copie_passeport.pdf, lettre_admission_sampere.pdf", { run: { size: 18 } }),
  pageBreak(),
);

// Part 11 — Interview + final check
children.push(
  h1("11. Interview preparation and final check"),
  p("Language-course files are more likely to be called in. Your answers must match your documents exactly. Keep each answer to 2 or 3 sentences, and practise questions 1, 2, 12 and 13 in Spanish."),
  table([4000, 5638], ["Likely question", "What a strong answer contains"], [
    ["1. Pourquoi l’Espagne ?", "The named master’s, and a concrete reason for that field and that university."],
    ["2. Pourquoi un cours de langue et pas directement le master ?", "The master’s requires B2; your certified level is [X]; the course is the step that gets you there."],
    ["3. Pourquoi Estudio Sampere, et cette ville ?", "Cervantes accreditation, exam preparation, and your personal reason for the campus."],
    ["4. Combien d’heures par semaine ? Quelles dates ?", "The exact figures from the school letter, in 60-minute hours."],
    ["5. Quel master, quelle université, quel niveau exigé ?", "Names, pre-registration number or coordinator email, and the language requirement."],
    ["6. Qui finance ? Quelle profession, quel revenu ?", "The parent’s job, employer and income, and the blocked amount."],
    ["7. Où allez-vous loger ?", "The address and dates from the booking, and whether it’s paid."],
    ["8. Que ferez-vous à la fin du cours ?", "SIELE/DELE → admission → master’s permit application before [date] (from Spain or Tunis); if not admitted, return on [date]."],
    ["9. Et après le master ?", "Your concrete professional plan."],
    ["10. Avez-vous de la famille en Espagne ?", "The truthful answer. Hiding it counts as bad faith."],
    ["11. Avez-vous déjà eu un refus de visa ?", "The truthful answer, and what has changed since."],
    ["12. Pourquoi ne pas apprendre l’espagnol à l’Institut Cervantes de Tunis ?", "Your true reason from Part 2, item 8: hours per week, immersion, hours already done in Tunis."],
    ["13. ¿Puede presentarse en español?", "3 or 4 simple sentences: name, studies, plan."],
  ]),
  blank(),
  h2("Final check before you submit"),
  check("My name is spelled identically everywhere (passport, form, school letter, insurance, translations)."),
  check("Course dates are on the school letter; stay dates (arrival → course end + 15 days) are on the form, insurance, sponsorship letter, accommodation and flight."),
  check("The blocked amount is at least 700 € × months of stay, plus the trip home and any unpaid fees."),
  check("The bank statements show no unexplained recent deposit, or I have proof of its source."),
  check("Every Tunisian public document has its apostille; every non-Spanish document has a sworn translation that includes the apostille."),
  check("Medical certificate, birth certificate, B3, payslips, attestation de travail and blocking certificate are less than 3 months old."),
  check("My passport was issued less than 10 years ago and is valid for at least 1 year."),
  check("I have proof of the master’s and a certificate of my current Spanish level."),
  check("The school letter says: in person, full-time, X hours a week in 60-minute hours, Centro acreditado por el Instituto Cervantes, paid in full."),
  check("The insurance is from an insurer authorised in Spain, with SNS-equivalent cover and no co-payment or waiting period."),
  check("Each item is stacked: original → copy → translation → copy."),
  check("I am submitting at least 2 months before the course starts, and I know I must apply for a TIE within 1 month of arriving if my stay exceeds 6 months."),
  check("I kept a complete copy of the file for myself."),
);

// ---------- document ----------
const doc = new Document({
  creator: "Visa dossier",
  title: "Spain Study Visa - Application Pack (Tunis)",
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: ACCENT },
        paragraph: { spacing: { before: 120, after: 200 }, outlineLevel: 0,
          border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ACCENT, space: 4 } } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 25, bold: true, font: "Arial", color: "222222" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial", color: ACCENT },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [
        { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } },
        { level: 1, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1080, hanging: 270 } } } } ] },
      ...["toc", "diff", "idx", "idxen", "idx4"].map(ref => ({ reference: ref, levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 360 } } } } ] })),
    ],
  },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } } },
    footers: { default: new Footer({ children: [new Paragraph({
      tabStops: [{ type: TabStopType.RIGHT, position: W }],
      children: [
        new TextRun({ text: "Spain study visa — application pack (Tunis)", size: 16, color: "888888" }),
        new TextRun({ children: ["\t", PageNumber.CURRENT], size: 16, color: "888888" }),
      ] })] }) },
    children,
  }],
});

const out = process.argv[2];
Packer.toBuffer(doc).then(buf => { fs.writeFileSync(out, buf); console.log("written", out, buf.length); });
