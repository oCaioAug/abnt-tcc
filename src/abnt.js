// Regras de formatação ABNT compartilhadas por todos os modelos
// (monografia hoje; artigo/NBR 6022 no futuro).
//
// Normas de referência:
//   NBR 14724:2011  Trabalhos acadêmicos – apresentação
//   NBR 6024:2012   Numeração progressiva das seções
//   NBR 6027:2012   Sumário
//   NBR 6028:2021   Resumo
//   NBR 10520:2023  Citações
//   NBR 6023:2018   Referências
//   Normas de apresentação tabular do IBGE (tabelas)

const {
  AlignmentType, LevelFormat, LevelSuffix, TabStopType, LeaderType,
} = require("docx");

// ---------------------------------------------------------------- unidades
const cm = (v) => Math.round(v * 567); // 1 cm = 567 twips (DXA)
const pt = (v) => v * 2;               // tamanho de fonte em meio-pontos

// ---------------------------------------------------------------- constantes
// A ABNT pede "fonte tamanho 12" mas não fixa a família: Arial e
// Times New Roman são as mais aceitas. Troque aqui se preferir.
const FONTE = "Arial";

const PAGINA = {
  largura: 11906, altura: 16838,              // A4
  margem: { superior: cm(3), esquerda: cm(3), inferior: cm(2), direita: cm(2) },
  cabecalho: cm(2),                           // nº de página a 2 cm da borda superior
  rodape: cm(1.25),
};
const LARGURA_UTIL = PAGINA.largura - PAGINA.margem.esquerda - PAGINA.margem.direita; // 16 cm

const LINHA_1_5 = 360;   // espaçamento 1,5
const LINHA_SIMPLES = 240;
const UMA_LINHA = 360;   // "um espaço de 1,5" entre título e texto
const RECUO_PARAGRAFO = cm(1.25);
const RECUO_CITACAO = cm(4);
const RECUO_NATUREZA = cm(8); // do meio da mancha gráfica para a margem direita

// ---------------------------------------------------------------- estilos
// novaPagina: na monografia (NBR 14724) seções primárias e títulos sem número começam em
// página nova; no artigo (NBR 6022) o texto é corrido e o título só ganha uma linha em branco antes.
function estilos({ novaPagina = true } = {}) {
  const corpo = { font: FONTE, size: pt(12) };
  const menor = { font: FONTE, size: pt(10) };
  const semRecuo = { indent: { left: 0, firstLine: 0 } };

  const tituloSecao = (nivel, run, extra = {}) => ({
    id: `Heading${nivel + 1}`,
    name: `heading ${nivel + 1}`,
    basedOn: "Normal",
    next: "Normal",
    quickFormat: true,
    run: { ...corpo, bold: false, italics: false, allCaps: false, color: "000000", ...run },
    paragraph: {
      alignment: AlignmentType.LEFT,
      spacing: { before: nivel === 0 && novaPagina ? 0 : UMA_LINHA, after: UMA_LINHA, line: LINHA_1_5 },
      ...semRecuo,
      keepNext: true,
      keepLines: true,
      outlineLevel: nivel,
      numbering: { reference: "secoes", level: nivel },
      ...extra,
    },
  });

  const tituloSemNumero = (id, name, outline) => ({
    id, name, basedOn: "Normal", next: "Normal", quickFormat: true,
    run: { ...corpo, bold: true, allCaps: true },
    paragraph: {
      alignment: AlignmentType.CENTER,
      spacing: { before: novaPagina ? 0 : UMA_LINHA, after: UMA_LINHA, line: LINHA_1_5 },
      ...semRecuo,
      keepNext: true,
      ...(novaPagina ? { pageBreakBefore: true } : {}),
      ...(outline === undefined ? {} : { outlineLevel: outline }),
    },
  });

  const sumario = (nivel, run) => ({
    id: `TOC${nivel}`,
    name: `toc ${nivel}`,
    basedOn: "Normal",
    next: "Normal",
    run: { ...corpo, ...run },
    paragraph: {
      alignment: AlignmentType.LEFT,
      spacing: { before: 0, after: 0, line: LINHA_1_5 },
      ...semRecuo,
      tabStops: [{ type: TabStopType.RIGHT, position: LARGURA_UTIL, leader: LeaderType.DOT }],
    },
  });

  return {
    default: {
      document: {
        run: { ...corpo, language: { value: "pt-BR" } },
        paragraph: { spacing: { before: 0, after: 0, line: LINHA_1_5 } },
      },
    },
    paragraphStyles: [
      {
        id: "Normal", name: "Normal", quickFormat: true,
        run: corpo,
        paragraph: {
          alignment: AlignmentType.JUSTIFIED,
          spacing: { before: 0, after: 0, line: LINHA_1_5 },
          indent: { firstLine: RECUO_PARAGRAFO },
          widowControl: true,
        },
      },

      // Seções numeradas (NBR 6024): a tipografia diferencia os níveis
      // e é repetida no sumário.
      tituloSecao(0, { bold: true, allCaps: true }, novaPagina ? { pageBreakBefore: true } : {}), // 1 SEÇÃO PRIMÁRIA
      tituloSecao(1, { allCaps: true }),                                       // 1.1 SEÇÃO SECUNDÁRIA
      tituloSecao(2, { bold: true }),                                          // 1.1.1 Seção terciária
      tituloSecao(3, {}),                                                      // 1.1.1.1 Seção quaternária
      tituloSecao(4, { italics: true }),                                       // 1.1.1.1.1 Seção quinária

      // Títulos sem indicativo numérico, centralizados (NBR 14724, 5.2.2)
      tituloSemNumero("TituloPreTextual", "Título pré-textual"),          // fora do sumário
      tituloSemNumero("TituloPosTextual", "Título pós-textual", 0),       // entra no sumário

      {
        id: "CitacaoLonga", name: "Citação longa", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: menor,
        paragraph: {
          alignment: AlignmentType.JUSTIFIED,
          spacing: { before: UMA_LINHA, after: UMA_LINHA, line: LINHA_SIMPLES },
          indent: { left: RECUO_CITACAO, firstLine: 0 },
        },
      },
      {
        id: "Caption", name: "caption", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { ...menor, bold: false, italics: false, color: "000000" },
        paragraph: {
          alignment: AlignmentType.CENTER,
          spacing: { before: UMA_LINHA, after: 60, line: LINHA_SIMPLES },
          ...semRecuo, keepNext: true,
        },
      },
      {
        id: "FonteIlustracao", name: "Fonte da ilustração", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: menor,
        paragraph: {
          alignment: AlignmentType.CENTER,
          spacing: { before: 60, after: UMA_LINHA, line: LINHA_SIMPLES },
          ...semRecuo,
        },
      },
      {
        id: "Ilustracao", name: "Ilustração", basedOn: "Normal", next: "FonteIlustracao", quickFormat: true,
        paragraph: { alignment: AlignmentType.CENTER, ...semRecuo, keepNext: true, spacing: { line: LINHA_SIMPLES } },
      },
      {
        id: "TextoTabela", name: "Texto de tabela", basedOn: "Normal", quickFormat: true,
        run: menor,
        paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 40, after: 40, line: LINHA_SIMPLES }, ...semRecuo },
      },
      {
        id: "Referencia", name: "Referência", basedOn: "Normal", next: "Referencia", quickFormat: true,
        run: corpo,
        paragraph: {
          alignment: AlignmentType.LEFT,
          spacing: { before: 0, after: LINHA_SIMPLES, line: LINHA_SIMPLES },
          ...semRecuo,
        },
      },
      {
        id: "Natureza", name: "Natureza do trabalho", basedOn: "Normal", quickFormat: true,
        paragraph: {
          alignment: AlignmentType.JUSTIFIED,
          spacing: { before: 0, after: LINHA_SIMPLES, line: LINHA_SIMPLES },
          indent: { left: RECUO_NATUREZA, firstLine: 0 },
        },
      },
      {
        id: "Centralizado", name: "Centralizado", basedOn: "Normal", quickFormat: true,
        paragraph: { alignment: AlignmentType.CENTER, ...semRecuo },
      },
      {
        id: "CapaDestaque", name: "Capa – destaque", basedOn: "Centralizado", quickFormat: true,
        run: { bold: true, allCaps: true },
      },
      {
        id: "SemRecuo", name: "Texto sem recuo", basedOn: "Normal", quickFormat: true,
        paragraph: { ...semRecuo },
      },
      {
        id: "Epigrafe", name: "Epígrafe", basedOn: "Normal", quickFormat: true,
        run: { italics: true },
        paragraph: { alignment: AlignmentType.RIGHT, indent: { left: RECUO_NATUREZA, firstLine: 0 } },
      },
      {
        id: "ItemLista", name: "Item de lista de siglas", basedOn: "Normal", quickFormat: true,
        paragraph: {
          alignment: AlignmentType.LEFT, ...semRecuo,
          tabStops: [{ type: TabStopType.LEFT, position: cm(2.5) }],
        },
      },
      {
        id: "Instrucao", name: "Instrução (apagar)", basedOn: "Normal", quickFormat: true,
        run: { ...menor, italics: true, color: "7F6000" },
        paragraph: {
          alignment: AlignmentType.LEFT,
          spacing: { before: 60, after: 120, line: LINHA_SIMPLES },
          ...semRecuo,
          shading: { type: "clear", fill: "FFF2CC", color: "auto" },
        },
      },

      // Sumário (NBR 6027): mesma tipografia dos títulos no texto
      sumario(1, { bold: true, allCaps: true }),
      sumario(2, { allCaps: true }),
      sumario(3, { bold: true }),
      sumario(4, {}),
      sumario(5, { italics: true }),
      {
        id: "TableofFigures", name: "table of figures", basedOn: "Normal", next: "Normal",
        paragraph: {
          alignment: AlignmentType.LEFT, ...semRecuo,
          tabStops: [{ type: TabStopType.RIGHT, position: LARGURA_UTIL, leader: LeaderType.DOT }],
        },
      },
      {
        id: "FootnoteText", name: "footnote text", basedOn: "Normal",
        run: menor,
        paragraph: { alignment: AlignmentType.JUSTIFIED, spacing: { line: LINHA_SIMPLES }, ...semRecuo },
      },
      {
        id: "Header", name: "header", basedOn: "Normal",
        run: menor,
        paragraph: { alignment: AlignmentType.RIGHT, ...semRecuo, spacing: { line: LINHA_SIMPLES } },
      },
    ],
  };
}

// Numeração progressiva 1, 1.1, 1.1.1 ... ligada aos estilos Título 1–5
function numeracao() {
  const textos = ["%1", "%1.%2", "%1.%2.%3", "%1.%2.%3.%4", "%1.%2.%3.%4.%5"];
  return {
    config: [{
      reference: "secoes",
      levels: textos.map((text, level) => ({
        level,
        format: LevelFormat.DECIMAL,
        text,
        alignment: AlignmentType.LEFT,
        suffix: LevelSuffix.SPACE,
        style: { paragraph: { indent: { left: 0, hanging: 0 } } },
      })),
    }],
  };
}

module.exports = {
  cm, pt, FONTE, PAGINA, LARGURA_UTIL, estilos, numeracao,
};
