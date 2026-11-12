// Esqueleto de artigo científico segundo a NBR 6022:2018.
// Sem capa nem sumário; texto corrido (as seções não começam em página nova).
// Ordem: pré-textuais (título, autoria, resumos, datas) -> textuais -> pós-textuais.

const { Document, Paragraph, TextRun, Header, PageNumber, FootnoteReferenceRun, AlignmentType } = require("docx");
const { cm, estilos, numeracao } = require("./abnt");
const { p, instr, secao, pos, legenda, tabelaExemplo, pagina, referencias, apendice, anexo } = require("./monografia");

const SIMPLES = { spacing: { line: 240 } };

// ---------------------------------------------------------------- pré-textuais
function cabecalhoDoArtigo() {
  return [
    instr("Modelo de artigo (NBR 6022:2018). Revistas e cursos costumam ter normas próprias de tamanho e formatação: confira as do seu orientador ou do periódico. As caixas amarelas são orientações; apague-as antes de entregar."),
    p("[TÍTULO DO ARTIGO]: [subtítulo, se houver]", "CapaDestaque"),
    p([{ i: "[Title in English]: [subtitle]" }], "Centralizado", { spacing: { after: cm(0.8) } }),
    instr("Título e subtítulo no idioma do artigo, separados por dois-pontos (obrigatório). O título em outro idioma é opcional e vem logo abaixo, em itálico."),
    p(["[Nome Completo do Autor]", new FootnoteReferenceRun(1)], "SemRecuo", { alignment: AlignmentType.RIGHT }),
    p(["[Nome Completo do Coautor ou Orientador]", new FootnoteReferenceRun(2)], "SemRecuo", { alignment: AlignmentType.RIGHT, spacing: { after: cm(0.8) } }),
    instr("Autoria (obrigatório): nome de cada autor, alinhado à direita. Em nota de rodapé, um breve currículo com vinculação institucional e o e-mail de contato."),
  ];
}

function resumos() {
  return [
    p([{ n: "RESUMO" }], "Centralizado", { spacing: { after: 240 } }),
    instr("Obrigatório (NBR 6028:2021). Parágrafo único, de 100 a 250 palavras, na terceira pessoa e em voz ativa: objetivo, método, resultados e conclusões. Sem citações."),
    p("[Texto do resumo em parágrafo único, com espaçamento simples.]", "SemRecuo", SIMPLES),
    p([{ n: "Palavras-chave: " }, "[palavra 1]; [palavra 2]; [palavra 3]."], "SemRecuo", { spacing: { before: 240, after: cm(0.8), line: 240 } }),

    p([{ n: "ABSTRACT" }], "Centralizado", { spacing: { after: 240 } }),
    instr("Resumo em outro idioma: opcional na NBR 6022, mas exigido pela maioria das revistas e cursos."),
    p("[Abstract text in a single paragraph.]", "SemRecuo", SIMPLES),
    p([{ n: "Keywords: " }, "[keyword 1]; [keyword 2]; [keyword 3]."], "SemRecuo", { spacing: { before: 240, after: cm(0.8), line: 240 } }),

    p("Data de submissão: [dia mês ano]. Data de aprovação: [dia mês ano].", "SemRecuo", SIMPLES),
    instr("Datas de submissão e de aprovação: obrigatórias no artigo publicado em periódico. Em artigo de TCC, siga a orientação do curso (podem ser removidas). A identificação e disponibilidade (DOI, ISSN, link) é opcional."),
  ];
}

// ---------------------------------------------------------------- textuais
function introducao() {
  return [
    secao(1, "INTRODUÇÃO"),
    instr("Parte inicial (obrigatória): tema, problema, objetivos e justificativa. No artigo as seções não começam em página nova; os estilos “Título 1” a “Título 5” já numeram e espaçam."),
    p("[Texto da introdução. O estilo Normal já aplica fonte 12, espaçamento 1,5, texto justificado e recuo de 1,25 cm na primeira linha.]"),
  ];
}

function desenvolvimento() {
  return [
    secao(1, "REFERENCIAL TEÓRICO"),
    instr("Desenvolvimento (obrigatório): divida em quantas seções forem necessárias (referencial, metodologia, resultados e discussão)."),
    p([
      "Citação indireta: segundo Sobrenome (2020), [ideia parafraseada]. ",
      "Citação direta curta: “[trecho copiado]” (Sobrenome, 2020, p. 15). ",
      "Três ou mais autores: (Sobrenome ", { i: "et al." }, ", 2021).",
    ]),
    p("[Trecho citado com mais de três linhas. Trecho citado com mais de três linhas. Trecho citado com mais de três linhas. Trecho citado com mais de três linhas. Trecho citado com mais de três linhas.] (Sobrenome, 2020, p. 32).", "CitacaoLonga"),
    secao(2, "SEÇÃO SECUNDÁRIA"),
    p("[Texto.]"),

    secao(1, "METODOLOGIA"),
    p("[Texto.]"),

    secao(1, "RESULTADOS E DISCUSSÃO"),
    p("Conforme a Tabela 1, [texto que menciona a tabela]."),
    legenda("Tabela", "[Título da tabela]"),
    tabelaExemplo(),
    p("Fonte: [Autor] (ano).", "FonteIlustracao"),
    p("[Texto.]"),
  ];
}

function consideracoesFinais() {
  return [
    secao(1, "CONSIDERAÇÕES FINAIS"),
    instr("Parte final (obrigatória): retome os objetivos e mostre as conclusões, as limitações e as sugestões de trabalhos futuros."),
    p("[Texto.]"),
  ];
}

// ---------------------------------------------------------------- pós-textuais
function agradecimentos() {
  return [
    pos("AGRADECIMENTOS"),
    instr("Opcional. Na NBR 6022:2018, os agradecimentos vêm depois das referências e dos apêndices e anexos. Apague se não for usar."),
    p("[Texto dos agradecimentos.]"),
  ];
}

// ---------------------------------------------------------------- documento
function construir() {
  const nota = (texto) => ({ children: [new Paragraph({ style: "FootnoteText", children: [new TextRun(texto)] })] });
  return new Document({
    creator: "Modelo ABNT",
    title: "Artigo científico",
    description: "Modelo de artigo segundo a NBR 6022:2018",
    styles: estilos({ novaPagina: false }),
    numbering: numeracao(),
    features: { updateFields: false },
    footnotes: {
      1: nota(" [Curso/titulação, Instituição]. E-mail: [autor@email.com]."),
      2: nota(" [Titulação, Instituição]. E-mail: [orientador@email.com]."),
    },
    sections: [{
      properties: pagina(),
      headers: {
        default: new Header({
          children: [new Paragraph({ style: "Header", children: [new TextRun({ children: [PageNumber.CURRENT] })] })],
        }),
      },
      children: [
        ...cabecalhoDoArtigo(),
        ...resumos(),
        ...introducao(),
        ...desenvolvimento(),
        ...consideracoesFinais(),
        ...referencias(),
        ...apendice(),
        ...anexo(),
        ...agradecimentos(),
      ],
    }],
  });
}

module.exports = { construir };
