// One-page competitor profile: Raízen (lubrificantes) — mesmo modelo do estudo da Iconic
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const gi = require("react-icons/gi");
const { applyTheme } = require("/root/.claude/skills/synced/798eb7cc-768c-4b8d-80d1-a2a0bc6f3467_d1b63ba7-5a89-4b75-b773-b3f2d018e30f/pptx/scripts/apply_theme.js");

const OUT = process.argv[2] || "Raizen_Lubrificantes_OnePage.pptx";

// Paleta Raízen (roxo institucional) + amarelo/vermelho Shell como apoio
const P = "772583";   // roxo Raízen
const P2 = "A35BB0";  // roxo médio
const PL = "F1E8F4";  // lavanda claro (fundos de card)
const PG = "D9C3E0";  // lavanda médio (bordas)
const Y = "F6B600";   // amarelo Shell
const R = "DD1D21";   // vermelho Shell
const INK = "2B2B2B";
const MUTE = "6B6B6B";
const GRID = "E3E3E3";

const THEME = {
  name: "Raizen Lubes",
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  colors: {
    dk1: "000000", lt1: "FFFFFF", dk2: INK, lt2: "F4F4F4",
    accent1: P, accent2: Y, accent3: R, accent4: P2, accent5: "9E9E9E", accent6: "4A1550",
    hlink: P, folHlink: P2,
  },
};

const W = 8.27, H = 15.2, M = 0.3, CW = W - 2 * M;

async function icon(Comp, color) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.defineLayout({ name: "ONEPAGE", width: W, height: H });
  pres.layout = "ONEPAGE";
  pres.theme = { headFontFace: "Arial", bodyFontFace: "Arial" };
  pres.title = "Raízen – Perfil de concorrente em lubrificantes";
  pres.author = "Estudo de concorrência";

  pres.defineSlideMaster({
    title: "ONEPAGE",
    background: { color: "FFFFFF" },
    objects: [],
  });
  pres.addSection({ title: "Raízen" });
  const s = pres.addSlide({ masterName: "ONEPAGE", sectionTitle: "Raízen" });

  const T = (text, o) => s.addText(text, Object.assign({ isTextBox: true, fontFace: "Arial", fontSize: 7.5, color: INK, margin: 0, valign: "top" }, o));
  const bar = (label, x, y, w) => {
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.22, fill: { color: P }, line: { color: P }, objectName: "Barra " + label });
    T(label, { x: x + 0.08, y, w: w - 0.1, h: 0.22, fontSize: 9, bold: true, color: "FFFFFF", valign: "middle" });
  };
  const box = (x, y, w, h, fill = PL, line = PG, name) =>
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: line, width: 0.75 }, objectName: name });
  const chartTitle = (t, x, y, w) => T(t, { x, y, w, h: 0.2, fontSize: 8, bold: true, color: INK });
  const src = (t, x, y, w) => T(t, { x, y, w, h: 0.14, fontSize: 5.5, italic: true, color: MUTE });

  // ---------------- Cabeçalho ----------------
  T([{ text: "raízen", options: { bold: true, color: P } }], { x: M, y: 0.12, w: 2.4, h: 0.5, fontSize: 30, valign: "middle" });
  T([
    { text: "Perfil de concorrente | Lubrificantes", options: { bold: true, color: INK, breakLine: true } },
    { text: "Raízen Lubrificantes (marca Shell) · atualizado em nov/2026", options: { color: MUTE } },
  ], { x: 4.2, y: 0.18, w: W - M - 4.2, h: 0.42, fontSize: 8, align: "right", valign: "middle" });

  // ---------------- Sumário executivo ----------------
  let y = 0.72;
  bar("Sumário executivo", M, y, CW);
  T([
    { text: "Fundada em 2011 como joint venture entre Cosan e Shell, a Raízen é uma das maiores empresas de energia do Brasil e líder global na produção de etanol de cana-de-açúcar. Atua em distribuição de combustíveis (6.734 postos Shell, mais de 70 terminais e 68 bases em aeroportos), açúcar, etanol e bioenergia (24 usinas) e lubrificantes. ", options: {} },
    { text: "Em lubrificantes, foi agente exclusiva da Shell de 2011 a 2021 e, em maio/2022, comprou todo o negócio da Shell Brasil: ", options: {} },
    { text: "fábrica da Ilha do Governador (RJ), terminal de óleos básicos de Campos Elíseos (Duque de Caxias), Shell Marine e a distribuição da marca Shell no país", options: { bold: true } },
    { text: ". A operação roda pela Blueway Trading (nome fantasia Raízen Lubrificantes). Em 2026, com dívida de R$ 65 bi, a companhia entrou em recuperação extrajudicial e será dividida em duas empresas. Combustíveis e lubrificantes Shell ficam na Raízen S.A., com aporte de R$ 3,5 bi da Shell.", options: {} },
  ], { x: M, y: y + 0.27, w: CW, h: 0.92, fontSize: 7.5, align: "justify", lineSpacingMultiple: 1.05 });

  // ---------------- Estrutura societária + Organograma ----------------
  y = 1.98;
  const ESW = 2.85;
  bar("Estrutura societária (capital aberto · RAIZ4)", M, y, ESW);
  bar("Organograma (gestão de transição · nov/2026)", M + ESW + 0.12, y, CW - ESW - 0.12);

  // Estrutura: 3 círculos
  const cy = y + 0.32;
  const circ = (cx, label, pct, fill, txt) => {
    s.addShape(pres.shapes.OVAL, { x: cx, y: cy, w: 0.66, h: 0.66, fill: { color: fill }, line: { color: fill }, objectName: "Acionista " + label });
    T(label, { x: cx - 0.1, y: cy, w: 0.86, h: 0.66, fontSize: 8, bold: true, color: txt, align: "center", valign: "middle" });
    T(pct, { x: cx - 0.15, y: cy + 0.69, w: 0.96, h: 0.22, fontSize: 11, bold: true, color: P, align: "center" });
  };
  circ(M + 0.12, "Cosan", "44%", "1B3A6B", "FFFFFF");
  circ(M + 1.08, "Shell", "44%", Y, R);
  circ(M + 2.04, "Free float", "12%", PG, P);
  T("Bloco de controle: 88% (Cosan e Shell, 50/50) até mar/26", { x: M, y: cy + 0.93, w: ESW, h: 0.16, fontSize: 6.5, color: MUTE, align: "center" });
  box(M, cy + 1.12, ESW, 0.5, "FFF6DB", "F2D27A", "Caixa reestruturação");
  T([
    { text: "Após a recuperação extrajudicial: ", options: { bold: true } },
    { text: "credores convertem 45% dos créditos em ações; a Shell aporta R$ 3,5 bi e a Aguassanta R$ 0,5 bi; a Cosan não acompanha e é diluída. O bloco de controle deve cair para cerca de 30%." },
  ], { x: M + 0.06, y: cy + 1.15, w: ESW - 0.12, h: 0.45, fontSize: 6.5 });

  // Organograma
  const OX = M + ESW + 0.12, OW = CW - ESW - 0.12; // ~4.8
  const ob = (t, x, yy, w, h, fill, color, bold = true, fs = 6.5) => {
    s.addShape(pres.shapes.RECTANGLE, { x, y: yy, w, h, fill: { color: fill }, line: { color: P, width: 0.75 }, objectName: "Org " + (Array.isArray(t) ? "caixa" : t) });
    T(t, { x: x + 0.03, y: yy, w: w - 0.06, h, fontSize: fs, bold, color, align: "center", valign: "middle" });
  };
  const ln = (x1, y1, x2, y2) => s.addShape(pres.shapes.LINE, { x: x1, y: y1, w: x2 - x1, h: y2 - y1, line: { color: P2, width: 0.75 } });
  const oy = y + 0.32;
  ob("Conselho de Administração", OX, oy + 0.3, 0.9, 0.42, "FFFFFF", INK);
  ob([{ text: "CEO", options: { bold: true, breakLine: true } }, { text: "Nelson Gomes", options: { bold: false } }], OX + 1.05, oy + 0.3, 0.8, 0.42, P, "FFFFFF");
  ln(OX + 0.9, oy + 0.51, OX + 1.05, oy + 0.51);
  ln(OX + 1.85, oy + 0.51, OX + 2.0, oy + 0.51);
  // duas empresas
  const ex = OX + 2.0, ew = OW - 2.0;
  ln(ex, oy + 0.17, ex, oy + 1.17);
  ln(ex, oy + 0.17, ex + 0.08, oy + 0.17);
  ln(ex, oy + 1.17, ex + 0.08, oy + 1.17);
  ob([{ text: "Raízen S.A.", options: { bold: true, breakLine: true } }, { text: "Combustíveis e lubrificantes Shell · CEO indicado: Lorival Luz", options: { bold: false } }], ex + 0.08, oy - 0.04, ew - 0.08, 0.42, P, "FFFFFF");
  const sub = ["Distribuição de combustíveis", "Raízen Lubrificantes (Blueway)", "Raízen Base Oil", "Aviação e Shell Select"];
  sub.forEach((t, i) => {
    const sx = ex + 0.08 + i * ((ew - 0.08) / 4);
    ob(t, sx + 0.01, oy + 0.42, (ew - 0.08) / 4 - 0.02, 0.4, i === 1 || i === 2 ? "FFF6DB" : PL, INK, i === 1 || i === 2, 5.8);
  });
  ob([{ text: "Raízen Energia", options: { bold: true, breakLine: true } }, { text: "Etanol, açúcar e bioenergia · indicado: Geovane Consul", options: { bold: false } }], ex + 0.08, oy + 0.99, ew - 0.08, 0.36, PL, P);
  T("A troca de gestão só vale depois que a capitalização da Raízen S.A. for aprovada (previsão: 18/dez/2026). Até lá, a diretoria atual continua.", { x: OX, y: oy + 1.42, w: OW, h: 0.24, fontSize: 6, italic: true, color: MUTE });

  // ---------------- Posicionamento estratégico ----------------
  y = 4.05;
  bar("Posicionamento estratégico", M, y, CW);
  const PW = 5.25;
  T([
    { text: "A Raízen é a ", options: {} },
    { text: "3ª/4ª maior distribuidora de lubrificantes do Brasil, com cerca de 11% de participação", options: { bold: true } },
    { text: ": atrás de Iconic, Vibra e Moove em 2023-24, à frente da Moove em 2025 (Lubes em Foco). Une fábrica própria, marca e tecnologia Shell (sob licença) e a rede de postos Shell. Desde a compra, a fatia caiu de ~15% (1T22) para 11-12%, com a transição de gestão e a pressão de Petronas, Castrol e independentes.", options: { breakLine: true } },
    { text: "Em 2023 verticalizou o suprimento com a ", options: {} },
    { text: "Raízen Base Oil (RBO)", options: { bold: true } },
    { text: ", que importa básicos dos grupos I, II e III via Shell Trading e revende a terceiros. Hoje a prioridade é proteger a marca durante a reestruturação: a Raízen S.A. busca investidor e o plano admite vender o controle.", options: {} },
  ], { x: M, y: y + 0.28, w: PW, h: 1.32, fontSize: 7.3, align: "justify", paraSpaceAfter: 3 });
  // KPI tiles 2x2
  const KX = M + PW + 0.12, KW = (CW - PW - 0.12 - 0.08) / 2, KH = 0.62;
  const kpis = [
    ["~11%", "participação de mercado (2025e)"],
    ["~171 mil m³", "volume estimado de acabados (2025e)"],
    ["220 mi L/ano", "capacidade da fábrica da Ilha do Governador"],
    ["+60 mil", "pontos de venda Shell no Brasil"],
  ];
  kpis.forEach(([v, l], i) => {
    const kx = KX + (i % 2) * (KW + 0.08), ky = y + 0.3 + Math.floor(i / 2) * (KH + 0.06);
    box(kx, ky, KW, KH, PL, PG, "KPI " + l);
    T(v, { x: kx + 0.05, y: ky + 0.05, w: KW - 0.1, h: 0.26, fontSize: 11, bold: true, color: P, align: "center", valign: "middle" });
    T(l, { x: kx + 0.05, y: ky + 0.32, w: KW - 0.1, h: 0.27, fontSize: 6, color: INK, align: "center" });
  });

  // ---------------- Segmentos ----------------
  y = 5.72;
  bar("Segmentos de atuação e frentes de negócio", M, y, CW);
  const segs = [
    [fa.FaCar, "Automotivo leve"], [fa.FaTruck, "Pesados e frotas"], [fa.FaMotorcycle, "Motocicletas"], [fa.FaTractor, "Agrícola"], [gi.GiMining, "Mineração"],
    [fa.FaIndustry, "Indústria"], [fa.FaShip, "Marítimo"], [fa.FaHardHat, "Construção"], [fa.FaBolt, "Geração de energia"], [fa.FaGasPump, "Postos e varejo"],
  ];
  const GX = M, GW = 3.55, cell = GW / 5;
  for (let i = 0; i < segs.length; i++) {
    const [C, lbl] = segs[i];
    const cx = GX + (i % 5) * cell, cyy = y + 0.32 + Math.floor(i / 5) * 0.74;
    s.addShape(pres.shapes.OVAL, { x: cx + cell / 2 - 0.22, y: cyy, w: 0.44, h: 0.44, fill: { color: PL }, line: { color: PG, width: 0.75 }, objectName: "Ícone " + lbl });
    s.addImage({ data: await icon(C, P), x: cx + cell / 2 - 0.13, y: cyy + 0.09, w: 0.26, h: 0.26, altText: lbl });
    T(lbl, { x: cx, y: cyy + 0.46, w: cell, h: 0.2, fontSize: 6, align: "center", color: INK });
  }
  const FX = M + GW + 0.15, FW = CW - GW - 0.15;
  const front = (t, d) => [
    { text: t, options: { bold: true, color: P, fontSize: 7.5, breakLine: true } },
    { text: d, options: { color: INK, fontSize: 6.6, breakLine: true } },
  ];
  T([
    ...front("Lubrificantes | Produção e comercialização", "Fabrica e distribui lubrificantes Shell (automotivos, industriais e graxas) a partir da Ilha do Governador (RJ), via distribuidores, postos Shell, oficinas e autopeças."),
    ...front("Óleos básicos | Raízen Base Oil (desde 2023)", "Importa e revende básicos grupos I, II e III (leve, médio e pesado), com tancagem na fábrica e em Duque de Caxias."),
    ...front("Lubrificantes marítimos | Shell Marine", "Produzidos pela Raízen. Desde out/2024 a Bunker One tem exclusividade de venda por 5 anos (meta: ~1.000 m³/ano)."),
    ...front("Serviços técnicos B2B", "Shell LubeAnalyst (análise de óleo e manutenção preditiva) e suporte a frotas, mineração e agro."),
  ], { x: FX, y: y + 0.28, w: FW, h: 1.5, paraSpaceAfter: 1.5 });

  // ---------------- Portfólio e marcas ----------------
  y = 7.55;
  bar("Portfólio e marcas", M, y, CW);
  const c1 = 2.95, c2 = 2.3, c3 = CW - c1 - c2 - 0.24;
  const lineItem = (b, d, last) => [
    { text: b + " ", options: { bold: true, color: P } },
    { text: d, options: { color: INK, breakLine: !last } },
  ];
  T([
    { text: "Lubrificantes acabados (marca Shell)", options: { bold: true, breakLine: true, fontSize: 7.5 } },
    ...lineItem("Shell Helix", "carros de passeio (Ultra, HX8, HX7, HX5). Lançou o Helix Ultra Professional AG 0W-20 (dexos1 Gen 3)"),
    ...lineItem("Shell Rimula", "diesel pesado e frotas (R6 LME Plus, R5, R4)"),
    ...lineItem("Shell Advance", "motocicletas 4T e 2T"),
    ...lineItem("Shell Spirax / Gadus", "transmissões e eixos / graxas"),
    ...lineItem("Shell Tellus / Omala / Turbo", "hidráulicos, engrenagens industriais e turbinas"),
    ...lineItem("Shell Marine", "lubrificantes marítimos (canal Bunker One)", true),
  ], { x: M, y: y + 0.28, w: c1, h: 1.48, fontSize: 6.6, paraSpaceAfter: 1.5 });
  T([
    { text: "Óleos básicos e serviços", options: { bold: true, breakLine: true, fontSize: 7.5 } },
    ...lineItem("RBO", "básicos grupos I, II e III importados, para misturadores terceiros"),
    ...lineItem("LubeAnalyst", "análise de óleo usado e manutenção preditiva"),
    ...lineItem("Assistência técnica", "recomendação de produtos e consultoria em campo para B2B"),
    ...lineItem("Embalagens", "frascos de 1 L em plástico reciclado nas linhas Helix, Rimula e Advance (-536 t de plástico/ano, jan/25)", true),
  ], { x: M + c1 + 0.12, y: y + 0.28, w: c2, h: 1.48, fontSize: 6.6, paraSpaceAfter: 1.5 });
  // Parcerias
  const PX = M + c1 + c2 + 0.24;
  T("Parcerias", { x: PX, y: y + 0.28, w: c3, h: 0.16, fontSize: 7.5, bold: true });
  const partners = [["Shell", "marca, tecnologia e Shell Trading", Y, R], ["Bunker One", "venda exclusiva Shell Marine", "1B3A6B", "FFFFFF"], ["Cosan", "sócia fundadora (44%)", "E9ECF2", "1B3A6B"]];
  partners.forEach(([n, d, f, c], i) => {
    const py = y + 0.47 + i * 0.3;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: PX, y: py, w: 0.78, h: 0.25, rectRadius: 0.05, fill: { color: f }, line: { color: f }, objectName: "Parceiro " + n });
    T(n, { x: PX, y: py, w: 0.78, h: 0.25, fontSize: 6.8, bold: true, color: c, align: "center", valign: "middle" });
    T(d, { x: PX + 0.84, y: py, w: c3 - 0.84, h: 0.25, fontSize: 6, color: INK, valign: "middle" });
  });
  T("Certificações", { x: PX, y: y + 1.38, w: c3, h: 0.16, fontSize: 7.5, bold: true });
  ["IATF 16949*", "ISO 14001**"].forEach((c, i) => {
    const bx = PX + i * 0.98;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: bx, y: y + 1.57, w: 0.92, h: 0.22, rectRadius: 0.05, fill: { color: "FFFFFF" }, line: { color: P, width: 0.75 }, objectName: "Cert " + c });
    T(c, { x: bx, y: y + 1.57, w: 0.92, h: 0.22, fontSize: 6.5, bold: true, color: P, align: "center", valign: "middle" });
  });

  // ---------------- Atuação no Brasil ----------------
  y = 9.5;
  bar("Atuação no Brasil", M, y, CW);
  const chY = y + 0.28, chH = 1.55;
  const aw1 = 2.2, aw2 = 2.75, aw3 = CW - aw1 - aw2 - 0.24;
  const axis = { catAxisLabelColor: MUTE, valAxisLabelColor: MUTE, catAxisLabelFontSize: 6.5, valAxisLabelFontSize: 6, catAxisLabelFontFace: "Arial", valAxisLabelFontFace: "Arial", dataLabelFontFace: "Arial" };

  chartTitle("Vendas de acabados | mil m³ (estimativa)", M, chY, aw1);
  s.addChart(pres.charts.BAR, [{ name: "Raízen", labels: ["2021", "2023e", "2024e", "2025e"], values: [174.0, 170.8, 165.1, 170.8] }], Object.assign({
    x: M, y: chY + 0.2, w: aw1, h: chH - 0.2, barDir: "col", chartColors: [P, P2, P2, P2], barGapWidthPct: 55,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 6.5, dataLabelColor: INK, dataLabelFormatCode: "0.0",
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false, valAxisMinVal: 0, valAxisMaxVal: 210,
  }, axis));

  chartTitle("Participação de mercado | %", M + aw1 + 0.12, chY, aw2);
  s.addChart(pres.charts.LINE, [{ name: "Raízen/Shell", labels: ["2021", "1T22", "2023", "1T24", "9M24", "2025e"], values: [14.9, 15.0, 11.7, 11.75, 10.9, 11.2] }], Object.assign({
    x: M + aw1 + 0.12, y: chY + 0.2, w: aw2, h: chH - 0.2, chartColors: [P], lineSize: 2, lineDataSymbol: "circle", lineDataSymbolSize: 5,
    showValue: true, dataLabelPosition: "t", dataLabelFontSize: 6.5, dataLabelColor: P, dataLabelFormatCode: "0.0",
    valAxisHidden: true, valAxisMinVal: 8, valAxisMaxVal: 17, valGridLine: { color: GRID, size: 0.5 }, catGridLine: { style: "none" }, showLegend: false,
  }, axis));

  const X3 = M + aw1 + aw2 + 0.24;
  chartTitle("Top 5 distribuidoras | % (9M24)", X3, chY, aw3);
  s.addChart(pres.charts.BAR, [{ name: "Participação", labels: ["Iconic", "Moove", "Vibra", "Raízen", "Petronas"], values: [18.1, 16.1, 15.5, 10.9, 8.9] }], Object.assign({
    x: X3, y: chY + 0.2, w: aw3, h: chH - 0.2, barDir: "bar", chartColors: ["BDBDBD", "BDBDBD", "BDBDBD", P, "BDBDBD"], barGapWidthPct: 40,
    catAxisOrientation: "maxMin", showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 6.5, dataLabelColor: INK, dataLabelFormatCode: "0.0",
    valAxisHidden: true, valAxisMaxVal: 23, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false,
  }, axis));
  src("2021: Sindicom (14,5 mil m³/mês). 2023-25: participação × mercado total da Lubes em Foco (1.524.596 m³ em 2025, +1,6%).", M, chY + chH + 0.02, aw1 + aw2 + 0.12);
  src("Fonte: Lubes em Foco. 2025: ~11,2% (MarqStats).", X3, chY + chH + 0.02, aw3);

  // ---------------- Perfil importador ----------------
  y = 11.55;
  bar("Perfil importador (2025)", M, y, CW);
  const iy = y + 0.28, ih = 1.45;
  const iw1 = 2.2, iw2 = 2.2, iw3 = CW - iw1 - iw2 - 0.24;
  chartTitle("Oferta de óleo básico no Brasil | %", M, iy, iw1);
  s.addChart(pres.charts.DOUGHNUT, [{ name: "Oferta", labels: ["Importado", "Produção nacional", "Rerrefino"], values: [44.1, 36.3, 19.6] }], {
    x: M, y: iy + 0.2, w: iw1, h: ih - 0.2, holeSize: 55, chartColors: [P, P2, Y],
    showPercent: false, showValue: true, dataLabelFormatCode: "0.0", dataLabelColor: "FFFFFF", dataLabelFontSize: 6.5, dataLabelFontFace: "Arial",
    showLegend: true, legendPos: "r", legendFontSize: 6.5, legendFontFace: "Arial", legendColor: INK,
  });
  chartTitle("Importação brasileira de básicos | mil t", M + iw1 + 0.12, iy, iw2);
  s.addChart(pres.charts.BAR, [{ name: "Importações", labels: ["2024", "2025"], values: [979, 931] }], Object.assign({
    x: M + iw1 + 0.12, y: iy + 0.2, w: iw2, h: ih - 0.2, barDir: "col", chartColors: [P2, P], barGapWidthPct: 70,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 6.5, dataLabelColor: INK, dataLabelFormatCode: "#,##0",
    valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 1150, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false,
  }, axis));
  const IX = M + iw1 + iw2 + 0.24;
  box(IX, iy, iw3, ih + 0.02, PL, PG, "Modelo de importação Raízen");
  T([
    { text: "Como a Raízen importa", options: { bold: true, color: P, fontSize: 7.5, breakLine: true } },
    { text: "Navios negociados todo mês com a Shell Trading, com acesso a refinarias do mundo todo", options: { bullet: { indent: 7 }, breakLine: true } },
    { text: "Píer próprio alfandegado (Receita Federal) dentro da fábrica da Ilha do Governador", options: { bullet: { indent: 7 }, breakLine: true } },
    { text: "Tancagem na fábrica e em Duque de Caxias, com carregamento rodoviário dedicado aos básicos", options: { bullet: { indent: 7 }, breakLine: true } },
    { text: "Origens prováveis: grupo II dos EUA (~3/4 das importações do país) e grupo III do Catar (Shell Pearl GTL), da Coreia e do Bahrein", options: { bullet: { indent: 7 } } },
  ], { x: IX + 0.08, y: iy + 0.05, w: iw3 - 0.14, h: ih - 0.06, fontSize: 6.4, paraSpaceAfter: 1.5 });
  src("Fontes: Lubes em Foco (751 mil m³ importados = 44,1% da oferta; rerrefino 19,6%), Lube Report/LNG (mil t), folheto RBO.", M, iy + ih + 0.04, CW);
  src("Volume importado pela própria Raízen e fornecedores por empresa exigem base de comércio exterior paga (ex.: LogComex).", M, iy + ih + 0.15, CW);

  // ---------------- Comentários estratégicos ----------------
  y = 13.66;
  bar("Comentários estratégicos", M, y, CW);
  const half = (CW - 0.15) / 2;
  const bl = (t, last) => ({ text: t, options: { bullet: { indent: 7 }, breakLine: !last } });
  T([
    bl("Escala e marca premium: Shell é top 3-4 no Brasil e tem a vantagem da rede de postos e do canal B2B."),
    bl("Perda de participação após a compra (de ~15% para ~11%): ainda não converteu a integração em ganho de mercado."),
    bl("Uso estimado de ~75-80% da capacidade da fábrica (171 mil m³ para 220 mi L/ano) dá espaço para crescer sem investir em ativos.", true),
  ], { x: M, y: y + 0.28, w: half, h: 1.05, fontSize: 6.6, paraSpaceAfter: 2 });
  T([
    bl("A RBO transforma a importação de básicos em fonte de receita e concorre diretamente com trading e misturadores independentes."),
    bl("Risco de governança: com a recuperação extrajudicial, a cisão e a possível venda do controle da Raízen S.A., a área de lubrificantes pode ter novo dono em 2027."),
    bl("O que monitorar: comprador da Raízen S.A., renovação da licença Shell, preço de básicos importados e expansão da RBO.", true),
  ], { x: M + half + 0.15, y: y + 0.28, w: half, h: 1.05, fontSize: 6.6, paraSpaceAfter: 2 });

  src("* IATF 16949 citada em vaga para a fábrica da Ilha do Governador. ** ISO 14001 vale para terminais de distribuição da Raízen. e = estimativa. Demais fontes: Raízen (Relatório Integrado 24/25 e 25/26, sala de imprensa), Money Times, Click Petróleo e Gás, AgFeed, Novacana, CNN Brasil.", M, H - 0.3, CW);

  s.addNotes("Fontes e premissas: participação Lubes em Foco (2023: 11,7%; 1T24: 11,75%; 9M24: 10,9%); 2025 ~11,2% MarqStats; mercado 2025 1.524.596 m³ (+1,6%) => 2024 ≈ 1.500,6 mil m³; 2023 ≈ 1.459,7 mil m³ (crescimento de 2,8% nos 9M24). Volumes Raízen = participação × mercado. 2021: Sindicom 14,5 mil m³/mês (14,9%). Capacidade: site Raízen Lubs 220 mi L (notícias de 2022 citavam 280 mil m³). Básicos: Lubes em Foco (importação 751.469 m³ = 44,1%; rerrefino 19,6%; nacional = restante), Lube Report (931 mil t em 2025 vs 979 mil t em 2024).");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
