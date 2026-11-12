// Esqueleto de monografia (TCC) segundo a NBR 14724:2011.
// Ordem dos elementos: pré-textuais -> textuais -> pós-textuais.

const {
  Document, Paragraph, TextRun, HeadingLevel, Header, PageNumber, TableOfContents,
  SequentialIdentifier, FootnoteReferenceRun, Table, TableRow, TableCell, WidthType,
  BorderStyle, AlignmentType, SectionType,
} = require("docx");
const { cm, PAGINA, LARGURA_UTIL, estilos, numeracao } = require("./abnt");

// ---------------------------------------------------------------- auxiliares
// Partes de texto: "texto simples" | { n: "negrito" } | { i: "itálico" }
const runs = (partes) =>
  (Array.isArray(partes) ? partes : [partes]).map((x) => {
    if (typeof x === "string") return new TextRun(x);
    if (x.n) return new TextRun({ text: x.n, bold: true });
    if (x.i) return new TextRun({ text: x.i, italics: true });
    return x; // já é um Run (campo, nota de rodapé...)
  });

const p = (partes, style = "Normal", extra = {}) =>
  new Paragraph({ style, children: runs(partes), ...extra });

const instr = (texto) => p(texto, "Instrucao");
const vazio = (style = "Centralizado") => new Paragraph({ style, children: [] });
const espaco = (twips) => ({ spacing: { before: twips } });

const H = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3,
  HeadingLevel.HEADING_4, HeadingLevel.HEADING_5];
const secao = (nivel, texto, extra = {}) =>
  new Paragraph({ heading: H[nivel - 1], children: [new TextRun(texto)], ...extra });

const pre = (texto) => p(texto, "TituloPreTextual");
const pos = (texto) => p(texto, "TituloPosTextual");

const NOVA = { pageBreakBefore: true };

const legenda = (rotulo, titulo) => new Paragraph({
  style: "Caption",
  children: [new TextRun(`${rotulo} `), new SequentialIdentifier(rotulo), new TextRun(` – ${titulo}`)],
});

// ---------------------------------------------------------------- pré-textuais
function capa() {
  return [
    p("[NOME DA INSTITUIÇÃO]", "CapaDestaque"),
    p("[Nome do curso]", "Centralizado"),
    p("[NOME COMPLETO DO AUTOR]", "CapaDestaque", espaco(cm(3))),
    p("[TÍTULO DO TRABALHO]:", "CapaDestaque", espaco(cm(5))),
    p("[subtítulo, se houver]", "Centralizado"),
    p("[Cidade]", "Centralizado", espaco(cm(9))),
    p("[Ano de entrega]", "Centralizado"),
  ];
}

function folhaDeRosto() {
  return [
    p("[NOME COMPLETO DO AUTOR]", "CapaDestaque"),
    p("[TÍTULO DO TRABALHO]:", "CapaDestaque", espaco(cm(6.5))),
    p("[subtítulo, se houver]", "Centralizado", { spacing: { after: cm(1.5) } }),
    p("Trabalho de Conclusão de Curso apresentado ao Curso de [Nome do Curso] da [Nome da Instituição], como requisito parcial para obtenção do título de [Bacharel/Licenciado] em [Área].", "Natureza"),
    p("Orientador(a): Prof.(ª) [Titulação] [Nome do orientador]", "Natureza"),
    p("Coorientador(a): Prof.(ª) [Titulação] [Nome] (se houver)", "Natureza"),
    p("[Cidade]", "Centralizado", espaco(cm(4.5))),
    p("[Ano de entrega]", "Centralizado"),
  ];
}

function fichaCatalografica() {
  return [
    p("FICHA CATALOGRÁFICA (verso da folha de rosto) – elemento obrigatório. Deve ser elaborada pela biblioteca da sua instituição, que fornece o quadro pronto. Cole-o na parte inferior desta página. Em versões só digitais, esta página fica logo após a folha de rosto.", "Instrucao", NOVA),
    p("[Insira aqui a ficha catalográfica fornecida pela biblioteca]", "Centralizado", espaco(cm(14))),
  ];
}

function folhaDeAprovacao() {
  const assinatura = (nome, inst) => [
    p("_______________________________________", "Centralizado", espaco(cm(1.2))),
    p(nome, "Centralizado", { spacing: { line: 240 } }),
    p(inst, "Centralizado", { spacing: { line: 240 } }),
  ];
  return [
    p("[NOME COMPLETO DO AUTOR]", "CapaDestaque", NOVA),
    p("[TÍTULO DO TRABALHO]: [subtítulo]", "CapaDestaque", { spacing: { before: cm(1.5), after: cm(1.5) } }),
    p("Trabalho de Conclusão de Curso apresentado ao Curso de [Nome do Curso] da [Nome da Instituição], como requisito parcial para obtenção do título de [Bacharel/Licenciado] em [Área].", "Natureza"),
    p("Aprovado em: ____ de _______________ de ______.", "SemRecuo", espaco(cm(1))),
    p("BANCA EXAMINADORA", "CapaDestaque", espaco(cm(1))),
    ...assinatura("Prof.(ª) [Titulação] [Nome do orientador] (Orientador)", "[Instituição]"),
    ...assinatura("Prof.(ª) [Titulação] [Nome do avaliador]", "[Instituição]"),
    ...assinatura("Prof.(ª) [Titulação] [Nome do avaliador]", "[Instituição]"),
  ];
}

function dedicatoria() {
  return [
    p("DEDICATÓRIA – elemento opcional (sem título). Apague esta página se não for usar.", "Instrucao", NOVA),
    p("[Texto da dedicatória.]", "Epigrafe", espaco(cm(16))),
  ];
}

function agradecimentos() {
  return [
    pre("AGRADECIMENTOS"),
    instr("Elemento opcional. Apague esta página se não for usar."),
    p("[Texto dos agradecimentos.]"),
  ];
}

function epigrafe() {
  return [
    p("EPÍGRAFE – elemento opcional (sem título). Citação seguida da indicação de autoria. Apague esta página se não for usar.", "Instrucao", NOVA),
    p("“[Texto da epígrafe.]”", "Epigrafe", espaco(cm(15))),
    p([new TextRun({ text: "([Autor], [ano])", italics: false })], "Epigrafe"),
  ];
}

function resumo() {
  return [
    pre("RESUMO"),
    instr("Elemento obrigatório (NBR 6028:2021). Parágrafo único, em voz ativa e na terceira pessoa, com 150 a 500 palavras. Apresente objetivo, método, resultados e conclusões. Não use citações."),
    p("[Texto do resumo em parágrafo único.]", "SemRecuo"),
    p([{ n: "Palavras-chave: " }, "[palavra 1]; [palavra 2]; [palavra 3]."], "SemRecuo", espaco(360)),
    instr("Palavras-chave: de 3 a 5 termos, separados por ponto e vírgula e finalizados por ponto."),
  ];
}

function abstract() {
  return [
    pre("ABSTRACT"),
    instr("Elemento obrigatório. Tradução do resumo para língua estrangeira (inglês: Abstract; espanhol: Resumen)."),
    p("[Abstract text in a single paragraph.]", "SemRecuo"),
    p([{ n: "Keywords: " }, "[keyword 1]; [keyword 2]; [keyword 3]."], "SemRecuo", espaco(360)),
  ];
}

function listaDeIlustracoes() {
  return [
    pre("LISTA DE ILUSTRAÇÕES"),
    instr("Elemento opcional (recomendado com 5 ou mais ilustrações). Atualiza sozinho: clique com o botão direito na lista > Atualizar campo."),
    new TableOfContents("Lista de ilustrações", { captionLabelIncludingNumbers: "Figura", hyperlink: true, hideTabAndPageNumbersInWebView: true }),
  ];
}

function listaDeTabelas() {
  return [
    pre("LISTA DE TABELAS"),
    instr("Elemento opcional. Atualiza da mesma forma que a lista de ilustrações."),
    new TableOfContents("Lista de tabelas", { captionLabelIncludingNumbers: "Tabela", hyperlink: true, hideTabAndPageNumbersInWebView: true }),
  ];
}

function listaDeSiglas() {
  return [
    pre("LISTA DE ABREVIATURAS E SIGLAS"),
    instr("Elemento opcional. Ordem alfabética: sigla, tabulação, significado por extenso."),
    p("ABNT\tAssociação Brasileira de Normas Técnicas", "ItemLista"),
    p("NBR\tNorma Brasileira", "ItemLista"),
    p("TCC\tTrabalho de Conclusão de Curso", "ItemLista"),
  ];
}

function sumario() {
  return [
    pre("SUMÁRIO"),
    instr("Elemento obrigatório (NBR 6027). Gerado automaticamente a partir dos títulos. Depois de editar o trabalho: clique com o botão direito no sumário > Atualizar campo > Atualizar o índice inteiro."),
    new TableOfContents("Sumário", {
      headingStyleRange: "1-5", useAppliedParagraphOutlineLevel: true,
      hyperlink: true, hideTabAndPageNumbersInWebView: true,
    }),
  ];
}

// ---------------------------------------------------------------- textuais
function introducao() {
  return [
    secao(1, "INTRODUÇÃO"),
    instr("Apresente o tema, a delimitação, o problema de pesquisa, os objetivos (geral e específicos) e a justificativa. Os títulos numerados já seguem a NBR 6024: aplique os estilos “Título 1” a “Título 5” e a numeração se ajusta sozinha."),
    p("[Texto da introdução. O estilo Normal já aplica fonte 12, espaçamento 1,5, texto justificado e recuo de 1,25 cm na primeira linha.]"),
    secao(2, "OBJETIVOS"),
    secao(3, "Objetivo geral"),
    p("[Texto.]"),
    secao(3, "Objetivos específicos"),
    p("[Texto.]"),
    secao(2, "JUSTIFICATIVA"),
    p("[Texto.]"),
  ];
}

function referencialTeorico() {
  return [
    secao(1, "REFERENCIAL TEÓRICO"),
    instr("Exemplos de citação conforme a NBR 10520:2023. Desde 2023, o sobrenome do autor é escrito apenas com a inicial maiúscula, tanto no texto quanto entre parênteses."),
    p([
      "Citação indireta com o autor no texto: segundo Sobrenome (2020), [ideia parafraseada]. ",
      "Citação direta curta, de até três linhas, entre aspas: “[trecho copiado]” (Sobrenome, 2020, p. 15). ",
      "Dois autores: (Sobrenome; Sobrenome, 2019). Três ou mais autores: (Sobrenome ",
      { i: "et al." }, ", 2021).",
    ]),
    p("Citação direta com mais de três linhas: use o estilo “Citação longa” (recuo de 4 cm, fonte 10, espaçamento simples, sem aspas):", "Normal"),
    p("[Trecho citado com mais de três linhas. Trecho citado com mais de três linhas. Trecho citado com mais de três linhas. Trecho citado com mais de três linhas. Trecho citado com mais de três linhas.] (Sobrenome, 2020, p. 32).", "CitacaoLonga"),
    p([
      "Notas de rodapé usam fonte 10 e espaçamento simples automaticamente",
      new FootnoteReferenceRun(1),
      ".",
    ]),

    secao(2, "SEÇÃO SECUNDÁRIA"),
    p("[Texto.]"),
    secao(3, "Seção terciária"),
    p("[Texto.]"),
    secao(4, "Seção quaternária"),
    p("[Texto.]"),
    secao(5, "Seção quinária"),
    p("[Texto.]"),

    secao(2, "ILUSTRAÇÕES E TABELAS"),
    instr("Ilustrações (NBR 14724, 5.8): título na parte superior (Figura n – Título) e fonte na parte inferior, ambos em fonte 10. Para inserir uma nova legenda numerada: Referências > Inserir Legenda > rótulo “Figura”."),
    p("Conforme a Figura 1, [texto que menciona a figura]."),
    legenda("Figura", "[Título da figura]"),
    new Paragraph({
      style: "Ilustracao",
      border: ["top", "bottom", "left", "right"].reduce((b, k) => ({ ...b, [k]: { style: BorderStyle.DASHED, size: 6, color: "999999", space: 8 } }), {}),
      spacing: { before: cm(1.5), after: cm(1.5) },
      indent: { left: cm(3), right: cm(3) },
      children: [new TextRun({ text: "[Insira a imagem aqui]", color: "808080" })],
    }),
    p("Fonte: Elaborado pelo autor (2026).", "FonteIlustracao"),

    instr("Tabelas (normas do IBGE): bordas laterais abertas, título em cima, fonte embaixo. Use “Texto de tabela” para o conteúdo das células."),
    p("A Tabela 1 apresenta [texto que menciona a tabela]."),
    legenda("Tabela", "[Título da tabela]"),
    tabelaExemplo(),
    p("Fonte: [Autor] (ano).", "FonteIlustracao"),
  ];
}

function tabelaExemplo() {
  const linha = { style: BorderStyle.SINGLE, size: 6, color: "000000" };
  const nada = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const larguras = [6000, 3000, 3000];
  const total = larguras.reduce((a, b) => a + b, 0);
  const dados = [
    ["Categoria", "Quantidade", "Percentual (%)"],
    ["[Item A]", "[00]", "[00,0]"],
    ["[Item B]", "[00]", "[00,0]"],
    ["Total", "[00]", "100,0"],
  ];
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: larguras,
    alignment: AlignmentType.CENTER,
    rows: dados.map((cols, i) => new TableRow({
      tableHeader: i === 0,
      children: cols.map((txt, j) => new TableCell({
        width: { size: larguras[j], type: WidthType.DXA },
        borders: {
          top: i <= 1 || i === dados.length - 1 ? linha : nada,
          bottom: i === 0 || i === dados.length - 1 ? linha : nada,
          left: nada, right: nada,
        },
        children: [new Paragraph({ style: "TextoTabela", children: [new TextRun({ text: txt, bold: i === 0 })] })],
      })),
    })),
  });
}

function metodologia() {
  return [
    secao(1, "METODOLOGIA"),
    instr("Descreva o tipo de pesquisa, a abordagem, os procedimentos, os instrumentos de coleta e a forma de análise dos dados."),
    p("[Texto.]"),
  ];
}

function resultados() {
  return [
    secao(1, "RESULTADOS E DISCUSSÃO"),
    p("[Texto.]"),
  ];
}

function conclusao() {
  return [
    secao(1, "CONCLUSÃO"),
    instr("Retome o problema e os objetivos, mostre como foram atingidos, aponte limitações e sugira trabalhos futuros. Não traga informações novas nem citações."),
    p("[Texto.]"),
  ];
}

// ---------------------------------------------------------------- pós-textuais
function referencias() {
  const ref = (partes) => p(partes, "Referencia");
  return [
    pos("REFERÊNCIAS"),
    instr("Elemento obrigatório (NBR 6023:2018). Ordem alfabética, alinhadas à esquerda, espaçamento simples e uma linha em branco entre elas (o estilo “Referência” já faz isso). O destaque (negrito) vai no título da obra ou no nome da revista. Modelos de cada tipo (apague os exemplos):"),
    ref(["NOME DA INSTITUIÇÃO. ", { n: "Título da página" }, ". Cidade, 2024. Disponível em: https://www.exemplo.com.br. Acesso em: 12 nov. 2026."]),
    ref(["SOBRENOME, Prenome. Título do capítulo. ", { i: "In" }, ": SOBRENOME, Prenome (org.). ", { n: "Título do livro" }, ". Cidade: Editora, 2019. p. 10-25."]),
    ref(["SOBRENOME, Prenome. ", { n: "Título do livro" }, ": subtítulo. 2. ed. Cidade: Editora, 2020."]),
    ref(["SOBRENOME, Prenome. ", { n: "Título do trabalho" }, ". 2022. Trabalho de Conclusão de Curso (Graduação em Área) – Nome da Instituição, Cidade, 2022."]),
    ref(["SOBRENOME, Prenome; SOBRENOME, Prenome. Título do artigo. ", { n: "Nome da Revista" }, ", Cidade, v. 12, n. 3, p. 45-60, jul./set. 2021. DOI: 10.0000/xxxx."]),
  ];
}

function apendice() {
  return [
    pos("APÊNDICE A – [TÍTULO DO APÊNDICE]"),
    instr("Elemento opcional. Apêndice = material elaborado pelo próprio autor (questionários, roteiros). Identifique com letras: APÊNDICE A, APÊNDICE B..."),
    p("[Conteúdo.]"),
  ];
}

function anexo() {
  return [
    pos("ANEXO A – [TÍTULO DO ANEXO]"),
    instr("Elemento opcional. Anexo = material de outra autoria (leis, documentos, mapas). Identifique com letras: ANEXO A, ANEXO B..."),
    p("[Conteúdo.]"),
  ];
}

// ---------------------------------------------------------------- documento
function pagina(extra = {}) {
  return {
    page: {
      size: { width: PAGINA.largura, height: PAGINA.altura },
      margin: {
        top: PAGINA.margem.superior, left: PAGINA.margem.esquerda,
        bottom: PAGINA.margem.inferior, right: PAGINA.margem.direita,
        header: PAGINA.cabecalho, footer: PAGINA.rodape,
      },
      ...extra,
    },
  };
}

function construir() {
  return new Document({
    creator: "Modelo ABNT",
    title: "Trabalho de Conclusão de Curso",
    description: "Modelo de monografia segundo a NBR 14724:2011",
    styles: estilos(),
    numbering: numeracao(),
    features: { updateFields: false },
    footnotes: {
      1: { children: [new Paragraph({ style: "FootnoteText", children: [new TextRun(" [Texto da nota de rodapé.]")] })] },
    },
    sections: [
      // 1. Capa: não é contada nem numerada
      { properties: pagina(), children: capa() },

      // 2. Pré-textuais: contados a partir da folha de rosto, sem número visível
      {
        properties: { ...pagina({ pageNumbers: { start: 1 } }), type: SectionType.NEXT_PAGE },
        children: [
          ...folhaDeRosto(),
          ...fichaCatalografica(),
          ...folhaDeAprovacao(),
          ...dedicatoria(),
          ...agradecimentos(),
          ...epigrafe(),
          ...resumo(),
          ...abstract(),
          ...listaDeIlustracoes(),
          ...listaDeTabelas(),
          ...listaDeSiglas(),
          ...sumario(),
        ],
      },

      // 3. Textuais e pós-textuais: número no canto superior direito
      {
        properties: { ...pagina(), type: SectionType.NEXT_PAGE },
        headers: {
          default: new Header({
            children: [new Paragraph({ style: "Header", children: [new TextRun({ children: [PageNumber.CURRENT] })] })],
          }),
        },
        children: [
          ...introducao(),
          ...referencialTeorico(),
          ...metodologia(),
          ...resultados(),
          ...conclusao(),
          ...referencias(),
          ...apendice(),
          ...anexo(),
        ],
      },
    ],
  });
}


module.exports = {
  construir,
  // partes reaproveitadas pelo modelo de artigo
  runs, p, instr, secao, pre, pos, legenda, tabelaExemplo, pagina, referencias, apendice, anexo,
};
