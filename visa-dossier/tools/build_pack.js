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

// Cover
children.push(
  new Paragraph({ spacing: { before: 1800, after: 200 }, children: [new TextRun({ text: "Spain Study Visa", size: 56, bold: true, color: ACCENT })] }),
  new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "Application pack — Embassy of Spain in Tunis (BLS)", size: 32 })] }),
  new Paragraph({ spacing: { after: 600 }, children: [new TextRun({ text: "Intensive Spanish course at Estudio Sampere → Master’s degree in Spain, September 2027", size: 24, italics: true, color: "555555" })] }),
  p("Applicant: [Nom et prénom]"),
  p("Prepared: 26 September 2026"),
  blank(),
  h3("How to use this pack"),
  bullet("Text highlighted in [yellow] is a personal fact only you can supply. Replace every highlighted blank, then remove the highlight."),
  bullet("Every fact must be true and backed by a document in your file. The consulate checks bank statements, school letters and employer letters, and a false statement leads to refusal."),
  bullet("Print each letter on its own page. Before printing, delete the section title and the blue guidance box above the letter. Sign in blue ink and date it."),
  bullet("Part 1 explains why the file is built this way. Parts 2 to 9 are what you prepare and submit."),
  pageBreak(),
);

// Contents
children.push(
  h1("Contents"),
  num("toc", "Analysis: what approved files contain, and why files get refused"),
  num("toc", "Final checklist, in the order BLS wants it"),
  num("toc", "Motivation letter (carta de motivación, Spanish)"),
  num("toc", "Cover letter and document index (carta de presentación, Spanish)"),
  num("toc", "Parent’s sponsorship letter (attestation de prise en charge, French)"),
  num("toc", "Text for the doctor: medical certificate (French)"),
  num("toc", "Email to Estudio Sampere: requesting the visa letter (Spanish)"),
  num("toc", "Email requesting a study-visa appointment (French)"),
  num("toc", "Interview preparation, and the final check before you submit"),
  pageBreak(),
);

// Part 1 — Analysis
children.push(
  h1("1. Analysis"),
  p("I compared four sources: the embassy’s official requirements page, an agency’s checklist for Tunis, and two posts from the “BLS visa espagne TUNISIE” group written by people who went through the process (June–July 2026). Here is what they agree on, where they differ, and what that means for your file."),
  h2("1.1 What every approved file has in common"),
  table([2600, 7038], ["Pattern", "What it means for you"], [
    ["**Complete and in order**", "BLS checks your file against its own checklist at the counter. Most setbacks come from a missing document or a missing translation, not from the applicant’s profile. Each document is stacked: original → copy → Spanish translation → copy of the translation."],
    ["**Apostille + sworn translation**", "Every Tunisian public document (birth certificate, B3, medical certificate, blocking certificate, sponsorship letter, diplomas) is apostilled by a notary and translated by a sworn translator from the embassy’s list. Where required, the apostille goes on both the original and the translation."],
    ["**Money blocked at the bank**", "Applicants from Tunis prove their funds with a student file at the bank (dossier scolaire) and a blocking certificate showing the money is “irrevocably blocked”. The agency uses €700 a month; the legal minimum is €600. Only you or your parents can be sponsors."],
    ["**Recent documents**", "The medical certificate, birth certificate, payslips and attestation de travail must all be less than 3 months old on the day of your appointment."],
    ["**Paid studies**", "A place reservation or the full fees, always with a letter from the school confirming it received the payment. If you pay in instalments, the unpaid balance goes into the blocked account."],
    ["**Short supporting documents**", "A short motivation letter (“don’t make it long”), a CV and a language certificate. None are mandatory, but they are consistently recommended."],
    ["**Rarely an interview**", "“99% without interview”, unless something in the file needs explaining. A clean, consistent file is how you avoid one."],
  ]),
  blank(),
  h2("1.2 How your case differs from the posts"),
  p("The group posts and the agency checklist were written for **university** admissions. You are applying for a **language school**, which the regulation treats as a separate category (“actividad formativa”, Real Decreto 1155/2024, art. 52.1.e.2º). That changes four things:"),
  num("diff", "**The school must be accredited by Instituto Cervantes, and the course must be in person.** The school’s letter has to say so explicitly (see Part 7)."),
  num("diff", "**You need an attestation de non-boursier** (a certificate that you hold no scholarship). The agency requires it for language schools only."),
  num("diff", "**The officer will ask themselves: “Is a language course the real purpose?”** Your answer is a coherent plan: your current Spanish level → the intensive course → the DELE B2 → named master’s programmes that require B2 → September 2027. The motivation letter (Part 3) is built around that chain."),
  num("diff", "**The permit ends 15 days after the course.** It lasts one year at most, and there is no automatic right to work. The regulation does not provide for switching to a master’s from inside Spain, so the plan states that you will return to Tunisia and apply for the master’s visa. That is honest, and it also reassures the officer."),
  h2("1.3 Why files get refused, and how this pack prevents it"),
  table([4200, 5438], ["Reason for refusal", "Prevented by"], [
    ["Missing apostille or translation", "Checklist (Part 2) with a translation and an apostille column for every item"],
    ["Too little money, or a recent unexplained deposit", "Blocked account of at least €700 × months; 6 months of the parent’s bank statements; proof of the source for any large deposit"],
    ["Insurance not accepted", "Health insurance from an insurer authorised in Spain, covering the whole stay, with no co-payment and no waiting period"],
    ["Medical certificate invalid", "Exact wording (Part 6), legalised by the regional health directorate, less than 3 months old"],
    ["School not eligible", "Visa letter showing Cervantes accreditation, in-person attendance and hours per week (Part 7)"],
    ["Purpose not credible", "Motivation letter with named master’s programmes, a DELE date and a timeline (Part 3)"],
    ["Dates or names inconsistent", "Final consistency check (Part 9)"],
    ["Late application", "Submit at least 2 months before the course starts"],
  ]),
  blank(),
  h2("1.4 Conflicts between the sources, and my recommendations"),
  bullet("**Insurance.** The embassy requires health insurance from an insurer authorised in Spain, with cover similar to Spain’s public health system. The agency offers travel insurance (about 500 TND, €30,000 cover). **Follow the embassy.** If you use the agency’s policy anyway, get written confirmation that it is accepted for a national study visa."),
  bullet("**Amount.** The legal minimum is €600 a month; the agency uses €700. **Block €700 a month**, which satisfies both."),
  bullet("**Flight and accommodation.** The embassy requires **neither**. The agency asks for a return flight reservation and 3 months of paid accommodation; the group lists the flight as optional. **A flight reservation is optional:** a simple reservation is enough, and don\u2019t buy the ticket before you have the visa. **One month of paid accommodation is enough** to show where you will live when you arrive. Only accommodation prepaid for the whole stay reduces the amount of money you must show."),
  bullet("**Appointment.** The embassy page only describes Schengen (tourist) appointments. The group reports requesting a study-visa appointment by email. **Ask BLS** which applies (info.tun@blshelpline.com, +216 71 138 748), and use Part 8 if email is the channel."),
  pageBreak(),
);

// Part 2 — Checklist
const Y = "Oui", N = "Non", R = "Conseillé", S = "Si pas en ES";
children.push(
  h1("2. Final checklist (in BLS order)"),
  p("Columns: **Trad.** = sworn Spanish translation. **Apost.** = apostille from a notary. “Conseillé” = recommended. “Si pas en ES” = translate only if the document isn’t already in Spanish. For each item, bring: original → copy → translation → copy of the translation."),
  table([500, 5738, 1000, 1000, 700, 700], ["#", "Document", "Trad.", "Apost.", "Prêt", "Copie"], [
    ["1", "Formulaire de visa national, tapé, signé à l’encre bleue et daté (dernière page)", N, N, "☐", "☐"],
    ["2", "Photo 3,5 × 4,5 cm, fond blanc, moins de 6 mois, collée sur le formulaire", N, N, "☐", "☐"],
    ["3", "Passeport (valide au moins 1 an, 2 pages vierges) + copie des 5 premières pages + anciens passeports et visas", N, N, "☐", "☐"],
    ["4", "Extrait de naissance en français, de moins de 3 mois", Y, Y, "☐", "☐"],
    ["5", "CIN (preuve de résidence en Tunisie)", R, N, "☐", "☐"],
    ["6", "Lettre d’admission Estudio Sampere (voir Partie 7)", N, N, "☐", "☐"],
    ["7", "Preuve de paiement : SWIFT + lettre de l’école confirmant la réception", S, N, "☐", "☐"],
    ["8", "Attestation de non-boursier (centres de langues uniquement)", N, N, "☐", "☐"],
    ["9", "Dossier scolaire bancaire + attestation de blocage irrévocable (≥ 700 € × mois + frais non payés)", Y, Y, "☐", "☐"],
    ["10", "Attestation de prise en charge du/des parent(s), signature légalisée à la municipalité (Partie 5)", Y, Y, "☐", "☐"],
    ["11", "CIN ou passeport du parent (copie)", R, N, "☐", "☐"],
    ["12", "Revenus du parent (moins de 3 mois) : salarié = 3 fiches de paie + attestation de travail ; commerçant = RNE + patente + déclaration d’ouverture ; retraité = pension + historique CNRPS/CNSS", Y, R, "☐", "☐"],
    ["13", "Relevés bancaires du parent, 6 mois, cachetés par la banque", N, N, "☐", "☐"],
    ["14", "Assurance santé (assureur autorisé en Espagne), toute la durée + 15 jours, sans copago ni carencia", S, N, "☐", "☐"],
    ["15", "Certificat médical de moins de 3 mois, légalisé par la Direction Régionale de la Santé (Partie 6)", Y, Y, "☐", "☐"],
    ["16", "Bulletin n°3 (B3) — demande en ligne, à lancer tôt", Y, Y, "☐", "☐"],
    ["17", "Hébergement (facultatif) : réservation + preuve de paiement ; 1 mois suffit (l’agence conseille 3 mois) ; date d’entrée = date d’arrivée", S, N, "☐", "☐"],
    ["18", "Réservation de vol (facultatif, non exigé par l’ambassade ; ne pas acheter le billet avant le visa)", N, N, "☐", "☐"],
    ["19", "Diplômes + relevés de notes (bac, licence, etc.)", Y, Y, "☐", "☐"],
    ["20", "Attestations de stage / de travail", Y, R, "☐", "☐"],
    ["21", "Lettre de motivation en espagnol (Partie 3)", N, N, "☐", "☐"],
    ["22", "CV en espagnol (format Europass)", N, N, "☐", "☐"],
    ["23", "Justificatif du niveau actuel d’espagnol (certificat ou test de placement)", N, N, "☐", "☐"],
    ["24", "Pages imprimées des masters visés (niveau d’espagnol exigé, calendrier d’admission)", N, N, "☐", "☐"],
    ["25", "Frais de visa : 372 TND en espèces + frais BLS 57,810 TND (à vérifier)", "—", "—", "☐", "—"],
  ], { center: [0, 2, 3, 4, 5] }),
  blank(),
  note("Timing: submit at least 2 months before the course starts (course on 1 February 2027 → by 1 December 2026 at the latest). Start with the B3 and the bank file, since those take the longest."),
  pageBreak(),
);

// Part 3 — Carta de motivación
const letterHead = (extra) => [
  p("[Nombre y apellidos]", { after: 0 }),
  p("[Dirección], [Código postal] [Ciudad], Túnez", { after: 0 }),
  p("Pasaporte n.º [número] · Tel.: [+216 …] · [correo electrónico]", { after: 160 }),
  p("Sección Consular", { after: 0 }),
  p("Embajada de España en Túnez", { after: 160 }),
  p("[Ciudad], a [día] de [mes] de 2026", { align: AlignmentType.RIGHT, after: 160 }),
  ...extra,
];
children.push(
  h1("3. Carta de motivación"),
  note("One page. Replace every highlighted blank with a true fact. If a sentence doesn’t apply to you, delete it rather than keep something you can’t prove."),
  ...letterHead([
    p("**Asunto: Carta de motivación – Solicitud de visado de estudios para un curso intensivo de lengua española**", { after: 200 }),
  ]),
  p("Señoras y señores:"),
  p("Me llamo [nombre], tengo [edad] años y resido en [ciudad], Túnez. Solicito un visado de estudios para realizar un curso intensivo y presencial de lengua española en Estudio Sampere de [Madrid/Salamanca/Alicante], centro acreditado por el Instituto Cervantes, del [fecha de inicio] al [fecha de fin], con [N] horas lectivas semanales.", { align: AlignmentType.JUSTIFIED }),
  p("**Mi trayectoria.** En [año] obtuve el título de [titulación] en la [universidad], con una nota media de [nota]. [Desde (año) trabajo como (puesto) en (empresa) / Realicé unas prácticas en (empresa) en (año)], experiencia que confirmó mi interés por especializarme en [área].", { align: AlignmentType.JUSTIFIED }),
  p("**Mi objetivo: un máster en España en septiembre de 2027.** Quiero cursar el [Máster en … de la Universidad …] o el [Máster en … de la Universidad …]. Ambos programas se imparten en español y exigen un nivel [B2] acreditado. Mi nivel actual es [A2/B1] ([justificante]). Por eso este curso intensivo no es un fin en sí mismo, sino el paso previo imprescindible para mi proyecto académico.", { align: AlignmentType.JUSTIFIED }),
  p("**¿Por qué Estudio Sampere?** Es un centro acreditado por el Instituto Cervantes, con una larga trayectoria desde 1956 y una preparación específica para el examen DELE. [Motivo personal y verdadero: p. ej., grupos reducidos, ciudad de la universidad objetivo]. Mi meta es obtener el DELE [B2] en la convocatoria de [mes] de 2027, título que presentaré en mis solicitudes de admisión.", { align: AlignmentType.JUSTIFIED }),
  p("**Mi calendario:**", { after: 40 }),
  bullet("[fecha] – [fecha]: curso intensivo de español en Estudio Sampere."),
  bullet("[mes] de 2027: examen DELE [B2]."),
  bullet("[meses] de 2027: solicitudes de admisión a los másteres."),
  bullet("Al terminar el curso: regreso a Túnez para solicitar ante esta Embajada el visado de estudios del máster."),
  p("**Financiación.** Mi [padre/madre], [nombre], [profesión] en [empleador], se compromete a cubrir todos mis gastos (carta de compromiso adjunta). He abierto un expediente escolar bancario con un bloqueo irrevocable de [importe] € (700 € al mes durante [N] meses), por encima del mínimo exigido. El curso está pagado en su totalidad y mi alojamiento en [ciudad] está reservado y pagado [hasta (fecha)].", { align: AlignmentType.JUSTIFIED }),
  p("**Mi futuro.** Tras el máster, mi objetivo es [plan profesional concreto] en Túnez, donde vive mi familia.", { align: AlignmentType.JUSTIFIED }),
  p("Me comprometo a cumplir la legislación española y las condiciones del visado durante toda mi estancia. Quedo a su disposición para cualquier documento adicional o entrevista.", { align: AlignmentType.JUSTIFIED }),
  p("Atentamente,", { after: 360 }),
  p("[Firma]", { after: 0 }),
  p("[Nombre y apellidos]"),
  pageBreak(),
);

// Part 4 — Carta de presentación
children.push(
  h1("4. Carta de presentación e índice"),
  note("Put this letter at the front of the file, and arrange the documents in the same order as the index."),
  ...letterHead([
    p("**Asunto: Solicitud de visado nacional de estudios – Curso de lengua española (actividad formativa, art. 52.1.e.2.º RD 1155/2024)**", { after: 200 }),
  ]),
  p("Señoras y señores:"),
  p("Presento mi solicitud de visado nacional de estudios para realizar un curso intensivo y presencial de lengua española en Estudio Sampere ([ciudad]), centro acreditado por el Instituto Cervantes, del [fecha de inicio] al [fecha de fin]. Adjunto la documentación ordenada según el siguiente índice:", { align: AlignmentType.JUSTIFIED }),
  p("**Documentación exigida**", { after: 40 }),
  ...[
    "Formulario de solicitud de visado nacional, cumplimentado y firmado, con fotografía.",
    "Pasaporte, copia de las cinco primeras páginas y pasaportes/visados anteriores.",
    "Certificado de nacimiento, apostillado y traducido.",
    "Documento de identidad tunecino (CIN) como justificante de residencia.",
    "Carta de admisión de Estudio Sampere y justificante de pago de la matrícula.",
    "Certificado de no becario.",
    "Expediente escolar bancario y certificado de bloqueo irrevocable por [importe] €, apostillado y traducido.",
    "Carta de compromiso de manutención de mi [padre/madre], legalizada, apostillada y traducida; copia de su CIN; [nóminas y certificado de trabajo] y extractos bancarios de los últimos 6 meses.",
    "Seguro de enfermedad [aseguradora], póliza n.º [número], válido del [fecha] al [fecha].",
    "Certificado médico, legalizado, apostillado y traducido.",
    "Certificado de antecedentes penales (Bulletin n.º 3), apostillado y traducido.",
    "Justificante del abono de la tasa de visado.",
  ].map(t => num("idx", t)),
  p("**Documentación complementaria**", { after: 40 }),
  ...[
    "Carta de motivación.",
    "Títulos y expedientes académicos, apostillados y traducidos.",
    "Certificados de prácticas y de trabajo, traducidos.",
    "Justificante del nivel actual de español.",
    "Información de los másteres objetivo.",
    "Reserva de alojamiento [para el primer mes] y justificante de pago.",
    "Reserva de vuelo.",
    "Currículum vítae.",
  ].map(t => num("idx", t)),
  p("Quedo a su disposición para cualquier aclaración, documentación adicional o entrevista."),
  p("Atentamente,", { after: 360 }),
  p("[Firma]", { after: 0 }),
  p("[Nombre y apellidos]"),
  pageBreak(),
);

// Part 5 — Prise en charge
children.push(
  h1("5. Attestation de prise en charge"),
  note("This follows the official BLS Tunisia model, plus two optional lines that make the file easier to read. Your parent signs it in front of the municipality (légalisation de signature). Then it gets an apostille from a notary and a sworn Spanish translation. The dates must match the school letter exactly."),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 240, after: 360 }, children: [new TextRun({ text: "ATTESTATION DE PRISE EN CHARGE", bold: true, size: 30, underline: {} })] }),
  p("Je soussigné(e) [Nom et prénom du parent],", { after: 160 }),
  p("en qualité de [père / mère],", { after: 160 }),
  p("titulaire de la CIN n° [numéro], délivrée le [date] à [lieu],", { after: 160 }),
  p("N° de téléphone : [numéro],", { after: 240 }),
  p("m’engage à subvenir à toutes les dépenses relatives à l’hébergement, aux transports et aux frais d’hospitalisation ou soins médicaux de :", { align: AlignmentType.JUSTIFIED, after: 240 }),
  p("mon fils / ma fille [Nom et prénom de l’étudiant(e)], né(e) le [date],", { after: 160 }),
  p("titulaire du passeport tunisien sous le N° [numéro],", { after: 240 }),
  p("pendant toute la durée de son séjour en Espagne du [date de début] au [date de fin], sans avoir recours aux aides publiques,", { align: AlignmentType.JUSTIFIED, after: 240 }),
  p("à raison d’un montant minimum de 700 € par mois, dans le cadre de son cours intensif de langue espagnole à Estudio Sampere, [ville], Espagne. (ligne optionnelle)", { align: AlignmentType.JUSTIFIED, after: 240 }),
  p("Je déclare exercer la profession de [profession] auprès de [employeur] depuis le [date], avec un revenu mensuel net de [montant] TND. (ligne optionnelle)", { align: AlignmentType.JUSTIFIED, after: 480 }),
  p("Fait à [ville], le [date]", { align: AlignmentType.RIGHT, after: 720 }),
  p("Signature (à légaliser auprès de la Municipalité)"),
  pageBreak(),
);

// Part 6 — Medical certificate
children.push(
  h1("6. Certificat médical: text to give your doctor"),
  note("The certificate must be issued less than 3 months before your appointment. If it comes from a private doctor, it must be legalised by the Direction Régionale de la Santé of the doctor’s region (or you can get it from a public health centre). Then it needs an apostille and a sworn Spanish translation. It must show your name, date of birth, CIN or passport number, and the doctor’s signature and stamp."),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 240, after: 360 }, children: [new TextRun({ text: "CERTIFICAT MÉDICAL", bold: true, size: 30 })] }),
  p("Je soussigné(e), Docteur [Nom du médecin], [spécialité], exerçant à [adresse du cabinet],", { align: AlignmentType.JUSTIFIED, after: 200 }),
  p("certifie avoir examiné ce jour M. / Mme [Nom et prénom], né(e) le [date de naissance], titulaire de la CIN n° [numéro] / du passeport n° [numéro],", { align: AlignmentType.JUSTIFIED, after: 200 }),
  p("et atteste que l’intéressé(e) ne souffre d’aucune des maladies pouvant avoir des répercussions graves sur la santé publique, conformément aux dispositions du Règlement Sanitaire International de 2005.", { align: AlignmentType.JUSTIFIED, after: 200, run: { bold: true } }),
  p("Certificat délivré à l’intéressé(e) pour servir et valoir ce que de droit.", { after: 360 }),
  p("Fait à [ville], le [date]", { align: AlignmentType.RIGHT, after: 480 }),
  p("Signature et cachet du médecin"),
  pageBreak(),
);

// Part 7 — Email to Sampere
children.push(
  h1("7. Email to Estudio Sampere: visa letter"),
  note("Send this after you have paid. The school’s letter is the document that proves you fit the legal category, so every point listed below matters."),
  p("**Asunto:** Carta de admisión para visado de estudios – [Nombre y apellidos] – Curso intensivo [fecha de inicio]"),
  p("Estimados señores:"),
  p("He realizado el pago del curso intensivo de español en su centro de [ciudad] del [fecha de inicio] al [fecha de fin] (referencia de pago: [referencia]). Para solicitar mi visado nacional de estudios en la Embajada de España en Túnez, les agradecería que me enviaran una carta de admisión firmada y sellada que indique:", { align: AlignmentType.JUSTIFIED }),
  ...[
    "Mi nombre completo, fecha de nacimiento y número de pasaporte ([número]).",
    "Las fechas exactas de inicio y fin del curso.",
    "Que el curso es intensivo, a tiempo completo y presencial, con el número de horas lectivas semanales.",
    "Que Estudio Sampere [ciudad] es un centro acreditado por el Instituto Cervantes.",
    "Que la matrícula y el curso están pagados en su totalidad (importe y fecha de recepción).",
    "En su caso, los datos del alojamiento reservado a través de la escuela (dirección, fechas y confirmación de pago).",
    "Si es posible, la preparación para el examen DELE incluida en el programa.",
  ].map(t => bullet(t)),
  p("Muchas gracias de antemano."),
  p("Un cordial saludo,", { after: 120 }),
  p("[Nombre y apellidos] · [teléfono] · [correo electrónico]"),
  pageBreak(),
);

// Part 8 — Appointment email
children.push(
  h1("8. Email requesting an appointment"),
  note("According to the group, study-visa applicants email the embassy once their file is about 15 days from ready, attaching their passport and admission letter as PDFs. Confirm the right address with BLS first (info.tun@blshelpline.com)."),
  p("**Objet :** Rendez-vous visa d’étude (long séjour) – [Nom et prénom] – Passeport [numéro]"),
  p("Madame, Monsieur,"),
  p("Je m’appelle [Nom et prénom], titulaire du passeport n° [numéro]. J’ai été admis(e) à un cours intensif de langue espagnole à Estudio Sampere ([ville]), centre accrédité par l’Instituto Cervantes, du [date de début] au [date de fin].", { align: AlignmentType.JUSTIFIED }),
  p("Je vous contacte afin de demander un rendez-vous pour déposer ma demande de visa national d’études. Mon dossier est complet (traductions et apostilles comprises).", { align: AlignmentType.JUSTIFIED }),
  p("Vous trouverez ci-joint mon passeport et ma lettre d’admission au format PDF."),
  p("Je vous remercie par avance et reste à votre disposition."),
  p("Cordialement,", { after: 120 }),
  p("[Nom et prénom] · [téléphone] · [e-mail]"),
  p("**Pièces jointes :** passeport.pdf, lettre_admission_sampere.pdf", { run: { size: 20 } }),
  pageBreak(),
);

// Part 9 — Interview + final check
children.push(
  h1("9. Interview preparation and final check"),
  p("Interviews are rare, but if you get one, your answers must match your documents exactly. Keep each answer short: 2 or 3 sentences. Practise answering questions 1, 2 and 11 in Spanish."),
  table([4000, 5638], ["Likely question", "What a strong answer contains"], [
    ["1. Pourquoi l’Espagne ?", "The named master’s programmes, and why that field and that university (a concrete reason)."],
    ["2. Pourquoi un cours de langue et pas directement le master ?", "The programmes require B2; your current level is [X]; the course is the step that gets you there, not the goal."],
    ["3. Pourquoi Estudio Sampere, et cette ville ?", "Cervantes accreditation, DELE preparation, and your personal reason."],
    ["4. Combien d’heures par semaine ? Quelles dates ?", "The exact figures from the school letter."],
    ["5. Quel master, quelle université, quel niveau exigé ?", "Names, and the language requirement you printed."],
    ["6. Qui finance ? Quelle profession, quel revenu ?", "The parent’s job, employer and income, and the blocked amount."],
    ["7. Où allez-vous loger ?", "The address and dates from the booking, and whether it’s paid."],
    ["8. Que ferez-vous à la fin du cours ?", "DELE exam → master’s admission → return to Tunisia to apply for the master’s visa."],
    ["9. Et après le master ?", "Your concrete professional plan."],
    ["10. Avez-vous de la famille en Espagne ?", "The truthful answer."],
    ["11. ¿Puede presentarse en español?", "3 or 4 simple sentences: name, studies, plan."],
  ]),
  blank(),
  h2("Final check before you submit"),
  check("My name is spelled identically everywhere (passport, form, school letter, insurance, translations)."),
  check("The dates match on the form, school letter, insurance, sponsorship letter, and the accommodation and flight reservations if I include them."),
  check("The insurance covers from my arrival until at least 15 days after the course ends."),
  check("The blocked amount is at least 700 € × months, plus any unpaid fees."),
  check("The bank statements show no unexplained recent deposit (or I have proof of its source)."),
  check("Every Tunisian public document has its apostille and sworn Spanish translation."),
  check("The medical certificate, birth certificate, payslips and attestation de travail are less than 3 months old."),
  check("I have the attestation de non-boursier."),
  check("The school letter says: in person, full-time, X hours a week, and Centro acreditado por el Instituto Cervantes."),
  check("Each item is stacked: original → copy → translation → copy."),
  check("I am submitting at least 2 months before the course starts."),
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
      ...["toc", "diff", "idx"].map(ref => ({ reference: ref, levels: [
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
