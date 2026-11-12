// Testes do verificador: npm test
const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");
const { DOMParser } = require("@xmldom/xmldom");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Footer, Header, PageNumber, AlignmentType,
  FootnoteReferenceRun, Table, TableRow, TableCell, WidthType, PageOrientation, SectionType,
} = require("docx");
const { verificar } = require("./verificador");

const cm = (v) => Math.round(v * 567);
const rodar = async (buf) => verificar(buf, { JSZip, DOMParser });
const item = (rel, id) => rel.itens.find((x) => x.id === id);
const detalhes = (rel, id) => item(rel, id).ocorrencias.map((o) => o.detalhe).join(" | ");

// ------------------------------------------------------------ o modelo gerado
const MODELO = path.join(__dirname, "..", "dist", "Modelo-TCC-Monografia-ABNT.docx");

test("o próprio modelo não tem erros, só avisos de conteúdo a preencher", { skip: !fs.existsSync(MODELO) }, async () => {
  const r = await rodar(fs.readFileSync(MODELO));
  const erros = r.itens.filter((x) => x.nivel === "erro");
  assert.deepEqual(erros.map((x) => `${x.id}: ${detalhes(r, x.id)}`), []);
  assert.equal(item(r, "modelo").nivel, "aviso");   // caixas amarelas e [colchetes]
  assert.equal(item(r, "resumo").nivel, "aviso");   // resumo de exemplo é curto
  for (const id of ["margens", "papel", "paginacao", "fonte-tamanho", "espacamento", "citacao-longa", "titulo-pagina", "titulo-numero", "titulo-centralizado", "referencias", "ilustracoes", "notas"]) {
    assert.equal(item(r, id).nivel, "ok", id);
  }
});

// ------------------------------------------------------------ o modelo de artigo (NBR 6022)
const ARTIGO = path.join(__dirname, "..", "dist", "Modelo-Artigo-ABNT.docx");

test("o modelo de artigo é reconhecido como artigo e não tem erros", { skip: !fs.existsSync(ARTIGO) }, async () => {
  const r = await rodar(fs.readFileSync(ARTIGO));
  assert.equal(r.tipo, "artigo");
  assert.equal(r.tipoDetectado, true);
  assert.deepEqual(r.itens.filter((x) => x.nivel === "erro").map((x) => `${x.id}: ${detalhes(r, x.id)}`), []);
  assert.equal(item(r, "titulo-pagina").nivel, "info");            // texto corrido
  assert.match(detalhes(r, "resumo"), /de 100 a 250/);

  // forçando "monografia", as exigências da NBR 14724 voltam
  const m = await verificar(fs.readFileSync(ARTIGO), { JSZip, DOMParser, tipo: "monografia" });
  assert.equal(m.tipo, "monografia");
  assert.match(m.itens.find((x) => x.id === "estrutura").ocorrencias.map((o) => o.trecho).join(","), /Sumário/);
  assert.equal(m.itens.find((x) => x.id === "titulo-pagina").nivel, "erro");
});

test("o modelo de monografia é reconhecido como monografia", { skip: !fs.existsSync(MODELO) }, async () => {
  const r = await rodar(fs.readFileSync(MODELO));
  assert.equal(r.tipo, "monografia");
});

test("toda regra informa a norma em que se baseia e explica a exigência", { skip: !fs.existsSync(MODELO) }, async () => {
  const { NORMAS } = require("./verificador");
  for (const [arq, artigo] of [[MODELO, false], [ARTIGO, true]]) {
    if (!fs.existsSync(arq)) continue;
    const r = await rodar(fs.readFileSync(arq));
    for (const it of r.itens) {
      assert.ok(it.norma && it.explicacao, `${it.id} sem norma/explicação`);
      if (/NBR/.test(it.norma)) assert.ok(it.normas.length, `${it.id}: ${it.norma} não está na lista de normas`);
    }
    assert.equal(r.itens.find((x) => x.id === "estrutura").norma, artigo ? "NBR 6022:2018" : "NBR 14724:2011, 4");
  }
  assert.ok(NORMAS.length >= 7 && NORMAS.every((n) => /^NBR \d+:\d{4}$/.test(n.codigo) && n.titulo && n.trata));
});

// ------------------------------------------------------------ um TCC cheio de erros
async function documentoComErros() {
  const corpo = { font: "Times New Roman", size: 22 }; // 11 pt
  const p = (texto, extra = {}) => new Paragraph({ children: [new TextRun({ text: texto, ...corpo })], spacing: { line: 480 }, ...extra });
  const titulo = (texto, extra = {}) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(texto)], ...extra });
  const doc = new Document({
    styles: { default: { document: { run: corpo } } },
    footnotes: { 1: { children: [new Paragraph({ children: [new TextRun({ text: "Nota em tamanho 12.", size: 24 })] })] } },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: cm(2.5), bottom: cm(2.5), left: cm(2.5), right: cm(2.5) } } },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT] })] })] }) },
      children: [
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "RESUMO", bold: true })] }),
        p("Resumo muito curto."),
        p("Segundo parágrafo do resumo."),
        p("Palavras-chave: ensino, escola, aprendizagem"),
        new Paragraph({ pageBreakBefore: true, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "SUMÁRIO", bold: true })] }),
        titulo("1 INTRODUÇÃO", { pageBreakBefore: true }),
        p("Como afirma SILVA (2020), o tema é relevante, o que também se vê em outros autores (SOUZA, 2019, p. 10), segundo dados do (IBGE, 2021).", { alignment: AlignmentType.LEFT }),
        new Paragraph({ indent: { left: cm(2) }, spacing: { line: 360 }, children: [new TextRun({ text: "“Citação longa com recuo errado, fonte 12 e aspas, que ocupa mais de três linhas no texto do trabalho acadêmico.” (Lima, 2018, p. 3).", size: 24 })] }),
        p("Texto com nota", { children: [new TextRun({ text: "Texto com nota", ...corpo }), new FootnoteReferenceRun(1)] }),
        titulo("METODOLOGIA"), // sem número e sem quebra de página
        new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun("2.1.1 Nível pulado")] }),
        new Table({
          width: { size: 9000, type: WidthType.DXA }, columnWidths: [9000],
          rows: [new TableRow({ children: [new TableCell({ width: { size: 9000, type: WidthType.DXA }, children: [new Paragraph("dado")] })] })],
        }),
        new Paragraph({ children: [new TextRun({ text: "Tabela 1 – Legenda embaixo e sem fonte", size: 24 })] }),
        p("Texto depois da tabela."),
        new Paragraph({ pageBreakBefore: true, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "REFERÊNCIAS", bold: true })] }),
        new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { line: 360 }, children: [new TextRun("SOUZA, Ana. Livro. São Paulo: Atlas, 2019.")] }),
        new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun("ALVES, Bruno. Outro livro. Rio de Janeiro: Record, 2020.")] }),
      ],
    }],
  });
  return Packer.toBuffer(doc);
}

test("aponta cada erro de um documento fora da norma", async () => {
  const r = await rodar(await documentoComErros());
  const nivel = (id) => item(r, id).nivel;

  assert.equal(nivel("estrutura"), "erro");
  assert.match(detalhes(r, "estrutura"), /não encontrado/);
  assert.match(item(r, "estrutura").ocorrencias.map((o) => o.trecho).join(","), /Abstract/);

  assert.equal(nivel("margens"), "erro");
  assert.match(detalhes(r, "margens"), /superior 2,5 cm \(deveria ser 3 cm\)/);
  assert.match(detalhes(r, "margens"), /direita 2,5 cm \(deveria ser 2 cm\)/);

  assert.equal(nivel("paginacao"), "erro");
  assert.match(detalhes(r, "paginacao"), /rodapé/);

  assert.equal(nivel("fonte-tamanho"), "erro");
  assert.match(detalhes(r, "fonte-tamanho"), /tamanho 11/);

  assert.equal(nivel("espacamento"), "erro");
  assert.match(detalhes(r, "espacamento"), /duplo/);

  assert.equal(nivel("alinhamento"), "aviso");

  assert.equal(nivel("citacao-longa"), "erro");
  assert.match(detalhes(r, "citacao-longa"), /recuo de 2 cm/);
  assert.match(detalhes(r, "citacao-longa"), /tamanho 12/);
  assert.match(detalhes(r, "citacao-longa"), /aspas/);

  assert.equal(nivel("titulo-pagina"), "erro");
  assert.match(item(r, "titulo-pagina").ocorrencias.map((o) => o.trecho).join(","), /METODOLOGIA/);
  assert.equal(nivel("titulo-numero"), "erro");
  assert.match(item(r, "titulo-numero").ocorrencias.map((o) => o.trecho).join(","), /METODOLOGIA/);
  assert.equal(nivel("titulo-centralizado"), "erro");
  assert.match(item(r, "titulo-centralizado").ocorrencias.map((o) => o.trecho).join(","), /SUMÁRIO/);
  assert.equal(nivel("titulo-hierarquia"), "aviso");

  assert.equal(nivel("referencias"), "erro");
  assert.match(detalhes(r, "referencias"), /justificado/);
  assert.match(detalhes(r, "referencias"), /espaçamento 1,5/);
  assert.match(detalhes(r, "referencias"), /ordem alfabética/);

  assert.equal(nivel("citacao-2023"), "aviso");
  const cits = item(r, "citacao-2023").ocorrencias.map((o) => o.trecho);
  assert.ok(cits.some((c) => c.startsWith("SILVA (2020")), cits.join(" / "));
  assert.ok(cits.some((c) => c.startsWith("(SOUZA, 2019")), cits.join(" / "));

  assert.equal(nivel("ilustracoes"), "erro");
  assert.match(detalhes(r, "ilustracoes"), /acima/);
  assert.match(detalhes(r, "ilustracoes"), /Fonte/);

  assert.equal(nivel("notas"), "erro");
  assert.match(detalhes(r, "notas"), /tamanho 12/);

  // como corrigir cada problema (usado pelo suplemento do Word)
  const corr = (id) => item(r, id).ocorrencias.map((o) => o.correcao).filter(Boolean);
  assert.deepEqual(corr("espacamento")[0], { formato: { linha: 1.5 } });
  assert.deepEqual(corr("fonte-tamanho")[0], { formato: { tamanho: 12 } });
  assert.deepEqual(corr("citacao-longa")[0].formato, { recuoEsq: 4, primeiraLinha: 0, tamanho: 10, linha: 1 });
  assert.ok(corr("titulo-pagina").every((c) => c.quebraAntes));
  assert.equal(corr("titulo-centralizado")[0].formato.alinhamento, "centralizado");
  assert.equal(corr("referencias")[0].formato.alinhamento, "esquerda");
  const ord = corr("referencias").find((c) => c.ordenar).ordenar;
  assert.deepEqual(ord.alvos.map((a) => a.trecho.slice(0, 5)), ["SOUZA", "ALVES"]);            // ordem atual
  assert.deepEqual(ord.conteudo.map((segs) => segs.map((s) => s.t).join("").slice(0, 5)), ["ALVES", "SOUZA"]); // ordem certa
  const cit = corr("citacao-2023").map((c) => c.texto.para);
  assert.ok(cit.includes("Silva (2020") && cit.includes("(Souza, 2019"), cit.join(" / "));
  assert.ok(corr("citacao-2023").every((c) => c.individual));
  assert.deepEqual(corr("resumo")[0].texto, { de: "ensino, escola, aprendizagem", para: "ensino; escola; aprendizagem." });
  if (process.env.SALVAR_DOC) fs.writeFileSync(process.env.SALVAR_DOC, await documentoComErros());

  assert.equal(nivel("resumo"), "aviso");
  assert.match(detalhes(r, "resumo"), /parágrafo único/);
  assert.match(detalhes(r, "resumo"), /ponto e vírgula/);
  assert.match(detalhes(r, "resumo"), /finalize/);
});

test("arquivo que não é .docx gera erro claro", async () => {
  await assert.rejects(() => rodar(Buffer.from("não é um docx")));
});

// ------------------------------------------------------------ títulos e fontes alterados à mão
test("aponta título numerado centralizado, título em tamanho 23 e parágrafo em outra fonte", async () => {
  const a4 = { size: { width: 11906, height: 16838 }, margin: { top: cm(3), left: cm(3), bottom: cm(2), right: cm(2) } };
  const texto = (t, run = {}) => new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { line: 360 }, children: [new TextRun({ text: t, font: "Arial", size: 24, ...run })] });
  const doc = new Document({
    styles: { default: { document: { run: { font: "Arial", size: 24 } } } },
    sections: [{ properties: { page: a4 }, children: [
      new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: true, alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "1 INTRODUÇÃO", font: "Aharoni", size: 46 })] }),
      texto("Primeiro parágrafo do texto, em Arial como o resto do trabalho, para que Arial seja a fonte principal."),
      texto("Segundo parágrafo também em Arial, com bastante texto para pesar na contagem das fontes do documento."),
      texto("Parágrafo escrito em outra fonte.", { font: "Agency FB" }),
    ] }],
  });
  const r = await rodar(await Packer.toBuffer(doc));
  assert.equal(item(r, "titulo-esquerda").nivel, "erro");
  assert.deepEqual(item(r, "titulo-esquerda").ocorrencias[0].correcao, { formato: { alinhamento: "esquerda" } });
  assert.equal(item(r, "titulo-tamanho").nivel, "erro");
  assert.match(detalhes(r, "titulo-tamanho"), /tamanho 23/);
  const fontes = item(r, "fonte-familia").ocorrencias;
  assert.ok(fontes.some((o) => /Agency FB/.test(o.detalhe) && o.correcao.formato.fonte === "Arial"), detalhes(r, "fonte-familia"));
  assert.ok(fontes.some((o) => /Aharoni/.test(o.detalhe)), detalhes(r, "fonte-familia"));
});

// ------------------------------------------------------------ posição dos parágrafos (para o suplemento)
test("informa a posição de cada parágrafo, mesmo com textos repetidos e tabelas", async () => {
  const t = (x, size = 24) => new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { line: 360 }, children: [new TextRun({ text: x, font: "Arial", size })] });
  const doc = new Document({
    styles: { default: { document: { run: { font: "Arial", size: 24 } } } },
    sections: [{ children: [
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "1 INTRODUÇÃO", size: 24 })] }), // 0
      t("[Texto.]"), t("[Texto.]", 32), t("[Texto.]"),                                                                   // 1, 2, 3
      new Table({ width: { size: 9000, type: WidthType.DXA }, columnWidths: [9000],
        rows: [new TableRow({ children: [new TableCell({ width: { size: 9000, type: WidthType.DXA }, children: [t("célula")] })] })] }), // 4
      t("[Texto.]"), t("[Texto.]", 26),                                                                                  // 5, 6
    ] }],
  });
  const r = await rodar(await Packer.toBuffer(doc));
  assert.deepEqual(item(r, "fonte-tamanho").ocorrencias.map((o) => o.posicao), [2, 6]);
});

// ------------------------------------------------------------ casos vistos num TCC real
test("não acusa paisagem com margens giradas, seção 'Lista de ...' no texto nem SUMÁRIO no estilo do sumário", async () => {
  const a4 = { size: { width: 11906, height: 16838 }, margin: { top: cm(3), left: cm(3), bottom: cm(2), right: cm(2) } };
  const t = (texto, extra = {}) => new Paragraph({ children: [new TextRun(texto)], ...extra });
  const centro = (texto) => t(texto, { pageBreakBefore: true, alignment: AlignmentType.CENTER });
  const doc = new Document({
    styles: { paragraphStyles: [{ id: "Sumrio1", name: "Sumário 1", basedOn: "Normal", paragraph: { alignment: AlignmentType.CENTER } }] },
    sections: [
      { properties: { page: a4 }, children: [
        centro("RESUMO"), centro("ABSTRACT"),
        new Paragraph({ style: "Sumrio1", pageBreakBefore: true, children: [new TextRun("SUMÁRIO")] }),
        new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: true, children: [new TextRun("1 INTRODUÇÃO")] }),
        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("1.1 LISTA DE RISCOS")] }),
        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("1.2 PRIORIZAÇÃO DOS RISCOS")] }),
        t("Lista de itens e equipamentos."),
      ] },
      { properties: { type: SectionType.NEXT_PAGE, page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: cm(3), right: cm(3), bottom: cm(2), left: cm(2) } } },
        children: [t("Quadro grande em paisagem.")] },
      { properties: { type: SectionType.NEXT_PAGE, page: a4 }, children: [
        centro("REFERÊNCIAS"), // "REFERÊNCIAS" com acento decomposto (NFD)
        t("SILVA, Ana. Livro. São Paulo: Atlas, 2019.", { alignment: AlignmentType.LEFT, spacing: { line: 240 } }),
        centro("APÊNDICE A: QUESTIONÁRIO"),
      ] },
    ],
  });
  const r = await rodar(await Packer.toBuffer(doc));
  for (const id of ["estrutura", "margens", "titulo-pagina", "titulo-centralizado", "titulo-hierarquia"]) {
    assert.equal(item(r, id).nivel, "ok", `${id}: ${detalhes(r, id)}`);
  }
  assert.equal(item(r, "titulo-numero").nivel, "aviso");
  assert.match(detalhes(r, "titulo-numero"), /travessão/);
});
