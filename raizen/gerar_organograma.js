// Organograma societário: Raízen > Blueway > Neolubes (lubrificantes)
const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/798eb7cc-768c-4b8d-80d1-a2a0bc6f3467_d1b63ba7-5a89-4b75-b773-b3f2d018e30f/pptx/scripts/apply_theme.js");
const OUT = process.argv[2] || "Raizen_Blueway_Neolubes_Organograma.pptx";

const P = "772583", P2 = "A35BB0", PL = "F1E8F4", PG = "D9C3E0", Y = "F6B600", R = "DD1D21";
const INK = "2B2B2B", MUTE = "6B6B6B", GREY = "F2F2F2", GLINE = "B5B5B5";
const THEME = { name: "Raizen Lubes", headFontFace: "Arial", bodyFontFace: "Arial",
  colors: { dk1: "000000", lt1: "FFFFFF", dk2: INK, lt2: "F4F4F4", accent1: P, accent2: Y, accent3: R, accent4: P2, accent5: "9E9E9E", accent6: "4A1550", hlink: P, folHlink: P2 } };

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
  pres.theme = { headFontFace: "Arial", bodyFontFace: "Arial" };
  pres.title = "Raízen – Blueway e Neolubes";
  pres.defineSlideMaster({ title: "ORG", background: { color: "FFFFFF" }, objects: [] });
  pres.addSection({ title: "Organograma" });
  const s = pres.addSlide({ masterName: "ORG", sectionTitle: "Organograma" });

  const T = (t, o) => s.addText(t, Object.assign({ isTextBox: true, fontFace: "Arial", fontSize: 9, color: INK, margin: 0, valign: "top" }, o));
  const rect = (x, y, w, h, fill, line, name, dash) => s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: line, width: 1, dashType: dash || "solid" }, objectName: name });
  const ln = (x1, y1, x2, y2, dash) => s.addShape(pres.shapes.LINE, { x: x1, y: y1, w: x2 - x1, h: y2 - y1, line: { color: P2, width: 1.25, dashType: dash || "solid" } });
  const tag = (t, x, y, w, fill, color) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.24, rectRadius: 0.05, fill: { color: fill }, line: { color: fill }, objectName: "Etiqueta " + t });
    T(t, { x, y, w, h: 0.24, fontSize: 8, bold: true, color, align: "center", valign: "middle" });
  };
  const b = (t) => ({ text: t, options: { bold: true } });
  const n = (t, br = true) => ({ text: t, options: { breakLine: br } });

  // Título
  T([{ text: "raízen", options: { bold: true, color: P } }], { x: 0.5, y: 0.25, w: 1.6, h: 0.5, fontSize: 24, valign: "middle" });
  T("Estrutura societária do negócio de lubrificantes: Blueway e Neolubes", { x: 2.2, y: 0.25, w: 10.6, h: 0.5, fontSize: 20, bold: true, color: INK, valign: "middle" });
  T("Situação em nov/2026 · participações, sócios e administradores conforme Formulário de Referência, ITR de 30/09/2026 e cadastro CNPJ da Receita Federal", { x: 0.5, y: 0.78, w: 12.3, h: 0.25, fontSize: 10, color: MUTE });

  // ---------- Nível 1: acionistas ----------
  const sh = [["Cosan S.A.", "44,0%", "1B3A6B", "FFFFFF", 0.6], ["Shell Brazil Holding BV", "44,0%", Y, R, 3.25], ["Free float (PN)", "11,9%", PG, P, 5.9]];
  sh.forEach(([nm, pc, f, c, x]) => {
    rect(x, 1.25, 2.4, 0.55, f, f, "Acionista " + nm);
    T([{ text: nm, options: { bold: true, breakLine: true } }, { text: pc + " do capital", options: {} }], { x, y: 1.25, w: 2.4, h: 0.55, fontSize: 10, color: c, align: "center", valign: "middle" });
    ln(x + 1.2, 1.8, x + 1.2, 1.98);
  });
  ln(1.8, 1.98, 7.1, 1.98);
  ln(4.45, 1.98, 4.45, 2.12);

  // ---------- Nível 2: Raízen S.A. ----------
  rect(1.75, 2.12, 5.4, 0.86, P, P, "Raízen S.A.");
  T([
    { text: "Raízen S.A. ", options: { bold: true, fontSize: 12 } }, { text: "(B3: RAIZ4 / RAIZ55) · CNPJ 33.453.598/0001-23", options: { fontSize: 8.5, breakLine: true } },
    { text: "CEO: Nelson Gomes · VP Mobilidade (combustíveis, lubrificantes e logística): Leonardo Gadotti", options: { fontSize: 8.5, breakLine: true } },
    { text: "Após a cisão do plano: concentra a distribuição de combustíveis e lubrificantes Shell", options: { fontSize: 8.5, italic: true } },
  ], { x: 1.85, y: 2.15, w: 5.2, h: 0.8, color: "FFFFFF", align: "center", valign: "middle" });
  ln(4.45, 2.98, 4.45, 3.3);
  tag("99,99% (FRE) · 100% (ITR)", 4.6, 3.03, 1.85, PL, P);

  // ---------- Nível 3: Blueway ----------
  rect(0.5, 3.3, 5.3, 2.0, PL, P, "Blueway");
  T([
    { text: "Blueway Trading Importação e Exportação S.A.", options: { bold: true, color: P, fontSize: 11.5, breakLine: true } },
    { text: "CNPJ 04.958.554/0001-57 · S.A. fechada · ativa · aberta em 13/03/2002", options: { color: MUTE, fontSize: 8.5, breakLine: true } },
    { text: "Papel: ", options: { bold: true } }, n("importadora do grupo (compras de derivados no exterior) e, desde ago/2023, operadora do negócio de lubrificantes com o nome fantasia “Raízen Lubrificantes”, incluindo a Raízen Base Oil."),
    { text: "Capital social: ", options: { bold: true } }, { text: "R$ 468,5 mi · ", options: {} }, { text: "Investimento na Raízen: ", options: { bold: true } }, n("R$ 1,01 bi (30/09/2026)"),
    { text: "Diretoria (Receita Federal): ", options: { bold: true } }, n("Juliano Tamaso (desde 04/2025; VP de Supply Chain da Raízen), Celso Henrique Martins Silva (12/2024), Sandro de Deus Lima Brito (08/2025) e Gustavo Vieiralves Koeler (10/2026)", false),
  ], { x: 0.62, y: 3.38, w: 5.08, h: 1.86, fontSize: 9, paraSpaceAfter: 2 });

  // Filiais
  ln(5.8, 4.3, 6.0, 4.3);
  rect(6.0, 3.3, 2.65, 2.0, "FFFFFF", PG, "Filiais Blueway");
  T([
    { text: "Filiais “Raízen Lubrificantes”", options: { bold: true, color: P, fontSize: 10, breakLine: true } },
    { text: "Ilha do Governador (RJ): ", options: { bold: true } }, n("fábrica · desde 23/08/2023"),
    { text: "Duque de Caxias (RJ): ", options: { bold: true } }, n("terminal de básicos e fabricação · 08/2023"),
    { text: "São Paulo (SP): ", options: { bold: true } }, n("escritório · 08/2023"),
    { text: "Cabo de Sto. Agostinho (PE): ", options: { bold: true } }, n("atacado · 08/2023"),
    { text: "S. José dos Campos (SP): ", options: { bold: true } }, n("atacado · 04/2024"),
    { text: "Amostra das filiais com dados confirmados na Receita.", options: { italic: true, color: MUTE, fontSize: 7.5 } },
  ], { x: 6.1, y: 3.38, w: 2.48, h: 1.88, fontSize: 8.3, paraSpaceAfter: 1.5 });

  // ---------- Nível 4: Neolubes ----------
  ln(3.15, 5.3, 3.15, 5.62, "dash");
  tag("100% das quotas (mai/2022 a mar/2026)", 3.3, 5.35, 2.5, GREY, MUTE);
  rect(0.5, 5.62, 5.3, 1.25, GREY, GLINE, "Neolubes", "dash");
  T([
    { text: "Neolubes Indústria de Lubrificantes Ltda.", options: { bold: true, color: MUTE, fontSize: 11.5, breakLine: true } },
    { text: "CNPJ 41.778.008/0001-60 · Ltda. · capital R$ 7,7 mi · sede na fábrica (RJ)", options: { color: MUTE, fontSize: 8.5, breakLine: true } },
    n("Criada pela Shell em 30/04/2021 para receber o negócio de lubrificantes. A Raízen comprou 100% das quotas em 02/05/2022 e a Blueway virou sócia em 31/05/2022. Últimos administradores: Adriano Bello, Sandro de Deus Lima Brito e Vinicius Nobrega da Rocha.", false),
  ], { x: 0.62, y: 5.68, w: 3.85, h: 1.16, fontSize: 8.2, color: INK });
  tag("BAIXADA", 4.75, 5.7, 0.95, R, "FFFFFF");
  T("incorporação · 03/03/2026", { x: 4.55, y: 5.97, w: 1.2, h: 0.32, fontSize: 7.5, color: R, align: "center" });

  // Leitura
  rect(6.0, 5.62, 2.65, 1.25, "FFF6DB", "F2D27A", "Leitura");
  T([
    { text: "Leitura", options: { bold: true, color: INK, fontSize: 10, breakLine: true } },
    n("Pelo cadastro da Receita, só a Blueway segue ativa: a Neolubes foi o veículo de compra. O negócio de lubrificantes roda dentro da Blueway, 100% da Raízen S.A., que fica com combustíveis e lubrificantes após a cisão.", false),
  ], { x: 6.1, y: 5.68, w: 2.48, h: 1.15, fontSize: 8.3 });

  // ---------- Linha do tempo ----------
  const TX = 9.0, TW = 3.85;
  rect(TX, 1.25, TW, 5.62, "FFFFFF", PG, "Linha do tempo");
  T("Linha do tempo", { x: TX + 0.15, y: 1.33, w: TW - 0.3, h: 0.3, fontSize: 12, bold: true, color: P });
  const ev = [
    ["2011", "Raízen vira agente exclusiva de vendas dos lubrificantes Shell (contrato de 10 anos)"],
    ["30/04/2021", "Shell cria a Neolubes para segregar o negócio de lubrificantes"],
    ["07/06/2021", "Raízen anuncia a compra (R$ 750 mi, segundo o 20-F da Cosan)"],
    ["02/05/2022", "Fechamento: Raízen compra 100% das quotas da Neolubes (fábrica, terminal de básicos, Shell Marine e distribuição)"],
    ["31/05/2022", "Blueway passa a ser a sócia da Neolubes"],
    ["ago/2023", "Blueway abre filiais “Raízen Lubrificantes”, incluindo a fábrica e Duque de Caxias. No mesmo ano nasce a Raízen Base Oil"],
    ["03/03/2026", "Neolubes é baixada por incorporação"],
    ["18/12/2026", "Capitalização prevista. Na cisão (até dez/2027), a Blueway fica na Raízen S.A."],
  ];
  const ty0 = 1.75, step = 0.62;
  s.addShape(pres.shapes.LINE, { x: TX + 0.3, y: ty0 + 0.08, w: 0, h: step * (ev.length - 1), line: { color: PG, width: 1.5 } });
  ev.forEach(([d, t], i) => {
    const yy = ty0 + i * step;
    s.addShape(pres.shapes.OVAL, { x: TX + 0.23, y: yy + 0.01, w: 0.14, h: 0.14, fill: { color: i === 6 ? R : P }, line: { color: "FFFFFF", width: 1 }, objectName: "Marco " + d });
    T([{ text: d, options: { bold: true, color: i === 6 ? R : P, breakLine: true } }, { text: t, options: {} }], { x: TX + 0.5, y: yy - 0.02, w: TW - 0.65, h: step - 0.02, fontSize: 8 });
  });

  T("Fontes: Raízen – Formulário de Referência (itens 1.1 e 6.4) e ITR de 30/09/2026 (notas 2.4, 11 e 14); RI Raízen (administração, fatos relevantes de 30/10/2026); Receita Federal via minhareceita.org (CNPJs 04.958.554 e 41.778.008, consulta em 25/11/2026); Cosan 20-F 2021 (preço); Forbes/MundoLogística (cargo de J. Tamaso). A Receita informa o motivo “incorporação” mas não o nome da incorporadora; como a Blueway era a única sócia, a incorporação por ela é a leitura mais provável.",
    { x: 0.5, y: 6.97, w: 12.35, h: 0.42, fontSize: 7, italic: true, color: MUTE });

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
