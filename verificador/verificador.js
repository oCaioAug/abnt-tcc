/*
 * Verificador de formatação ABNT para trabalhos acadêmicos (.docx).
 *
 * Lê o XML do documento (Word ou LibreOffice salvando em .docx), calcula a
 * formatação efetiva de cada parágrafo (padrões do documento -> estilos ->
 * formatação direta) e confere com:
 *   NBR 14724:2011 (apresentação), 6024:2012 (numeração), 6027:2012 (sumário),
 *   6028:2021 (resumo), 10520:2023 (citações), 6023:2018 (referências).
 *
 * Sem dependências próprias: recebe um DOMParser e um JSZip, para rodar
 * no navegador e no Node (testes).
 */
(function (raiz, fabrica) {
  if (typeof module === "object" && module.exports) module.exports = fabrica();
  else raiz.VerificadorABNT = fabrica();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const A = "http://schemas.openxmlformats.org/drawingml/2006/main";

  // ---------------------------------------------------------------- unidades
  const CM = 567;                       // twips por centímetro
  const cm = (tw) => (tw / CM).toFixed(2).replace(".", ",").replace(/,?0+$/, "") + " cm";
  const perto = (v, alvo, tol) => v != null && Math.abs(v - alvo) <= tol;

  // ---------------------------------------------------------------- XML
  const filhos = (el, nome) => (el ? Array.from(el.childNodes).filter((n) => n.nodeType === 1 && n.localName === nome) : []);
  const filho = (el, nome) => filhos(el, nome)[0];
  const todos = (el, nome) => (el ? Array.from(el.getElementsByTagNameNS(W, nome)) : []);
  const at = (el, nome) => {
    if (!el) return null;
    const v = el.getAttributeNS(W, nome) || el.getAttribute("w:" + nome);
    return v === "" ? null : v;
  };
  const num = (el, nome) => { const v = at(el, nome); return v == null ? undefined : Number(v); };
  const liga = (el) => (el ? !["0", "false", "off"].includes(at(el, "val") || "") : undefined);

  const definidos = (o) => Object.fromEntries(Object.entries(o).filter(([, v]) => v !== undefined));
  const junta = (...objs) => Object.assign({}, ...objs.map((o) => definidos(o || {})));

  // ---------------------------------------------------------------- propriedades
  function lerRPr(rPr, tema) {
    if (!rPr) return {};
    const f = filho(rPr, "rFonts");
    let fonte = at(f, "ascii") || at(f, "hAnsi") || undefined;
    const tf = at(f, "asciiTheme") || at(f, "hAnsiTheme");
    if (!fonte && tf && tema) fonte = /^major/.test(tf) ? tema.major : tema.minor;
    const sz = num(filho(rPr, "sz"), "val");
    return {
      fonte,
      tamanho: sz === undefined ? undefined : sz / 2,
      negrito: liga(filho(rPr, "b")),
      italico: liga(filho(rPr, "i")),
      caixaAlta: liga(filho(rPr, "caps")),
      sobrescrito: filho(rPr, "vertAlign") ? at(filho(rPr, "vertAlign"), "val") !== "baseline" : undefined,
      oculto: liga(filho(rPr, "vanish")),
    };
  }

  function lerPPr(pPr) {
    if (!pPr) return {};
    const sp = filho(pPr, "spacing");
    const ind = filho(pPr, "ind");
    const numPr = filho(pPr, "numPr");
    const numId = numPr ? at(filho(numPr, "numId"), "val") : undefined;
    const ol = num(filho(pPr, "outlineLvl"), "val");
    const jc = at(filho(pPr, "jc"), "val");
    return {
      alinhamento: jc ? ({ both: "justificado", distribute: "justificado", center: "centralizado", right: "direita", end: "direita", left: "esquerda", start: "esquerda" }[jc] || jc) : undefined,
      linha: num(sp, "line"),
      regraLinha: sp ? at(sp, "lineRule") || (num(sp, "line") !== undefined ? "auto" : undefined) : undefined,
      antes: num(sp, "before"),
      depois: num(sp, "after"),
      recuoEsq: num(ind, "left") ?? num(ind, "start"),
      primeiraLinha: num(ind, "firstLine") ?? (num(ind, "hanging") !== undefined ? -num(ind, "hanging") : undefined),
      nivelTopico: ol === undefined || ol >= 9 ? undefined : ol,
      numerado: numId === undefined ? undefined : numId !== "0",
      quebraAntes: liga(filho(pPr, "pageBreakBefore")),
    };
  }

  // ---------------------------------------------------------------- estilos
  function lerEstilos(doc, tema) {
    const estilos = {};
    let padraoP = null;
    const raiz = doc ? doc.documentElement : null;
    const dd = filho(raiz, "docDefaults");
    const padrao = {
      r: lerRPr(filho(filho(dd, "rPrDefault"), "rPr"), tema),
      p: lerPPr(filho(filho(dd, "pPrDefault"), "pPr")),
    };
    for (const s of filhos(raiz, "style")) {
      const id = at(s, "styleId");
      estilos[id] = {
        id, tipo: at(s, "type"), nome: at(filho(s, "name"), "val") || id,
        base: at(filho(s, "basedOn"), "val"),
        p: lerPPr(filho(s, "pPr")), r: lerRPr(filho(s, "rPr"), tema),
      };
      if (at(s, "type") === "paragraph" && ["1", "true"].includes(at(s, "default"))) padraoP = id;
    }
    const memo = {};
    function cadeia(id, prof = 0) {
      if (!id || !estilos[id] || prof > 20) return { p: {}, r: {}, nomes: [] };
      if (memo[id]) return memo[id];
      const s = estilos[id];
      const b = cadeia(s.base, prof + 1);
      return (memo[id] = { p: junta(b.p, s.p), r: junta(b.r, s.r), nomes: [s.nome, ...b.nomes] });
    }
    return { estilos, padrao, padraoP, cadeia };
  }

  function lerTema(doc) {
    if (!doc) return null;
    const pega = (nome) => {
      const f = doc.getElementsByTagNameNS(A, nome)[0];
      const l = f && f.getElementsByTagNameNS(A, "latin")[0];
      return l ? l.getAttribute("typeface") : undefined;
    };
    return { major: pega("majorFont"), minor: pega("minorFont") };
  }

  // ---------------------------------------------------------------- parágrafos
  function textoRun(r) {
    let t = "";
    for (const n of Array.from(r.childNodes)) {
      if (n.nodeType !== 1) continue;
      if (n.localName === "t") t += n.textContent;
      else if (n.localName === "tab") t += "\t";
      else if (n.localName === "noBreakHyphen") t += "-";
    }
    return t;
  }

  function analisarParagrafo(p, est, ctx) {
    const pPr = filho(p, "pPr");
    const estiloId = at(filho(pPr, "pStyle"), "val") || est.padraoP;
    const cad = est.cadeia(estiloId);
    const props = junta(est.padrao.p, cad.p, lerPPr(pPr));
    const rBase = junta(est.padrao.r, cad.r, lerRPr(filho(pPr, "rPr"), ctx.tema));

    const runs = [];
    let texto = "";
    let quebraInicio = false, quebraFim = false;
    for (const r of todos(p, "r")) {
      // ignora texto excluído em revisão
      let pai = r.parentNode, excluido = false;
      while (pai && pai !== p) { if (pai.localName === "del" || pai.localName === "moveFrom") excluido = true; pai = pai.parentNode; }
      if (excluido) continue;
      const rPr = filho(r, "rPr");
      const cs = est.cadeia(at(filho(rPr, "rStyle"), "val"));
      const pr = junta(rBase, cs.r, lerRPr(rPr, ctx.tema));
      for (const br of filhos(r, "br")) {
        if (at(br, "type") === "page") { if (texto.trim()) quebraFim = true; else quebraInicio = true; }
      }
      const t = textoRun(r);
      if (!t) continue;
      texto += t;
      if (!pr.oculto) runs.push({ texto: t, props: pr });
    }
    const instr = todos(p, "instrText").map((x) => x.textContent).join(" ") +
      " " + todos(p, "fldSimple").map((x) => at(x, "instr") || "").join(" ");
    const nomes = cad.nomes.join("|");
    let nivel;
    if (props.nivelTopico !== undefined) nivel = props.nivelTopico + 1;
    else {
      const m = /^(heading|t[íi]tulo|titulo)\s*(\d)$/i.exec(cad.nomes[cad.nomes.length - 1] || "") || /^(heading|t[íi]tulo)\s*(\d)$/i.exec(cad.nomes[0] || "");
      if (m) nivel = Number(m[2]);
    }
    return {
      el: p, texto: texto.replace(/\s+/g, " ").trim(), props, runs, estilo: cad.nomes[0] || estiloId || "",
      nivel,
      emTabela: ctx.emTabela, emSumario: ctx.emSumario || /^(toc|contents|sum[áa]rio|[íi]ndice)\s*\d/i.test(cad.nomes[0] || "") || /table of figures|[íi]ndice de ilustra/i.test(nomes) || /\bTOC\b/.test(instr),
      instrucao: /instru[çc][ãa]o|apagar/i.test(nomes),
      desenho: todos(p, "drawing").length + todos(p, "pict").length + todos(p, "object").length > 0,
      sectPr: filho(pPr, "sectPr") || null,
      quebraInicio, quebraFim,
    };
  }

  function lerCorpo(doc, est, tema) {
    const corpo = filho(doc.documentElement, "body");
    const blocos = [];
    let posicao = 0; // índice do parágrafo na ordem de leitura (a mesma de body.paragraphs no Word)
    // linhas e células podem vir dentro de controles de conteúdo (sdt)
    const conteudo = (el) => filho(el, "sdtContent") || el;
    const celulas = (el, ctx) => {
      for (const n of Array.from(el.childNodes)) {
        if (n.nodeType !== 1) continue;
        if (n.localName === "tr") celulas(n, ctx);
        else if (n.localName === "tc") andar(n, ctx);
        else if (n.localName === "sdt" || n.localName === "customXml") celulas(conteudo(n), ctx);
      }
    };
    const andar = (el, ctx) => {
      for (const n of Array.from(el.childNodes)) {
        if (n.nodeType !== 1) continue;
        if (n.localName === "p") blocos.push({ tipo: "p", posicao: posicao++, ...analisarParagrafo(n, est, ctx) });
        else if (n.localName === "tbl") {
          blocos.push({ tipo: "tabela" });
          celulas(n, { ...ctx, emTabela: true });
          blocos.push({ tipo: "fimTabela" });
        } else if (n.localName === "sdt") {
          const g = todos(filho(n, "sdtPr"), "docPartGallery").map((x) => at(x, "val")).join(" ");
          andar(filho(n, "sdtContent") || n, { ...ctx, emSumario: ctx.emSumario || /contents|figures/i.test(g) });
        } else if (n.localName === "sectPr") blocos.push({ tipo: "secao", sectPr: n });
        else if (["customXml", "smartTag", "ins"].includes(n.localName)) andar(n, ctx);
      }
    };
    andar(corpo, { tema });
    return blocos;
  }

  // ---------------------------------------------------------------- classificação
  const normaliza = (t) => t.normalize("NFC").toLocaleUpperCase("pt-BR").replace(/^\d+(\.\d+)*\.?\s+/, "").replace(/\s+/g, " ").trim();
  const PRE_TEXTUAIS = /^(ERRATA|AGRADECIMENTOS?|RESUMO|ABSTRACT|RESUMEN|RÉSUMÉ|RESUME|LISTA DE .+|SUM[ÁA]RIO)$/;
  const POS_TEXTUAIS = /^(REFER[ÊE]NCIAS( BIBLIOGR[ÁA]FICAS)?|GLOSS[ÁA]RIO|AP[ÊE]NDICES?( [A-Z])?(\s*[–—:-] .*)?|ANEXOS?( [A-Z])?(\s*[–—:-] .*)?|[ÍI]NDICE)$/;
  const LEGENDA = /^(Figura|Tabela|Quadro|Gr[áa]fico|Fotografia|Imagem|Mapa|Desenho|Esquema|Fluxograma|Organograma|Planta|Ilustra[çc][ãa]o)\s+\d+/i;
  // "Lista de ...", "Resumo" etc. só são títulos pré-textuais antes da Introdução
  // (no texto, "4.1 LISTA DE RISCOS" é uma seção comum); títulos não terminam em ponto.
  const preTextual = (b) => !b.posIntro && PRE_TEXTUAIS.test(normaliza(b.texto));
  const semNumeroTitulo = (b) => !/\.$/.test(b.texto.trim()) && (preTextual(b) || POS_TEXTUAIS.test(normaliza(b.texto)));
  const ehTitulo = (b) => b.nivel !== undefined || (b.texto.length <= 80 && semNumeroTitulo(b) && !b.emSumario);
  // Formatação da lista de referências (NBR 6023): esquerda, simples, sem recuo, linha em branco depois
  const FORMATO_REF = { alinhamento: "esquerda", linha: 1, recuoEsq: 0, primeiraLinha: 0, depois: 12 };
  // "SILVA" -> "Silva", "SOUZA DE ALMEIDA" -> "Souza de Almeida"
  const capitalizaNome = (nome) => nome.toLocaleLowerCase("pt-BR").replace(/(^|[\s'’-])(\p{L})/gu, (m, sep, l) => sep + l.toLocaleUpperCase("pt-BR"))
    .replace(/ (De|Da|Do|Dos|Das|E)(?= )/g, (m, p) => " " + p.toLowerCase());
  const temNumero = (b) => b.props.numerado || /^\d+(\.\d+)*\.?\s/.test(b.texto);

  // ---------------------------------------------------------------- relatório
  function novoRelatorio() {
    const itens = [];
    const regra = (id, titulo, norma) => {
      const it = { id, titulo, norma, nivel: "ok", mensagem: "", ocorrencias: [] };
      itens.push(it);
      return {
        it,
        // correcao (opcional) descreve como corrigir, sem depender do Word:
        //   { formato: {alinhamento, linha, recuoEsq, primeiraLinha, depois, tamanho} }  (cm / pt / múltiplo de linha)
        //   { quebraAntes: true }                       página nova antes do parágrafo
        //   { texto: { de, para }, individual: true }   troca de texto, só item a item
        // trecho: texto, ou o próprio parágrafo ({texto, posicao}) para o suplemento achá-lo pela posição
        problema(nivel, trecho, detalhe, correcao) {
          if (nivel === "erro" || it.nivel === "ok") it.nivel = nivel;
          const bloco = trecho && typeof trecho === "object";
          const t = bloco ? trecho.texto : trecho;
          const o = { trecho: t == null ? "" : String(t).slice(0, 110), detalhe };
          if (bloco && trecho.posicao !== undefined) o.posicao = trecho.posicao;
          if (correcao) o.correcao = correcao;
          it.ocorrencias.push(o);
        },
        msg(m) { it.mensagem = m; return this; },
        info(m) { if (it.nivel === "ok") it.nivel = "info"; it.mensagem = m; },
      };
    };
    return { itens, regra };
  }

  const descreveLinha = (pr) => {
    if (pr.linha === undefined) return "simples";
    if (pr.regraLinha === "auto") {
      const m = pr.linha / 240;
      return ({ 1: "simples", 1.5: "1,5", 2: "duplo" }[m] || m.toFixed(2).replace(".", ",")) + (m in { 1: 1, 1.5: 1, 2: 1 } ? "" : " linhas");
    }
    return `${pr.regraLinha === "exact" ? "exatamente" : "pelo menos"} ${(pr.linha / 20).toFixed(1).replace(".", ",")} pt`;
  };
  const alinhado = (a) => ({ esquerda: "alinhado à esquerda", direita: "alinhado à direita", centralizado: "centralizado", justificado: "justificado" }[a || "esquerda"] || a);
  // Ordena só pela entrada (sobrenome/entidade até a vírgula ou o ponto): empates
  // seguem regras próprias da NBR 6023 (autor sozinho antes de coautoria etc.).
  const entradaRef = (t) => t.split(/[,.;]/)[0].replace(/[^\p{L}\p{N} ]/gu, "").trim();
  const linhaSimples = (pr) => pr.linha === undefined || (pr.regraLinha === "auto" && pr.linha <= 252);
  const linha15 = (pr) => pr.regraLinha === "auto" && perto(pr.linha, 360, 12);

  // tamanhos de fonte efetivos do texto visível (ignora sobrescritos: chamadas de nota)
  const tamanhos = (b) => [...new Set(b.runs.filter((r) => r.texto.trim() && !r.props.sobrescrito).map((r) => r.props.tamanho ?? 10))];

  // ---------------------------------------------------------------- normas
  // Normas usadas como base. As explicações são resumos com palavras próprias (o texto das
  // normas é protegido e vendido pela ABNT); o número do item só aparece quando é certo.
  const NORMAS = [
    { codigo: "NBR 14724:2011", titulo: "Trabalhos acadêmicos — Apresentação",
      trata: "Estrutura de TCCs, dissertações e teses (elementos pré-textuais, textuais e pós-textuais) e regras de formatação: papel, margens, fonte, espaçamento, paginação, ilustrações e tabelas." },
    { codigo: "NBR 6022:2018", titulo: "Artigo em publicação periódica técnica e/ou científica — Apresentação",
      trata: "Estrutura do artigo: título, autoria, resumos, datas, introdução, desenvolvimento, considerações finais e pós-textuais." },
    { codigo: "NBR 6023:2018", titulo: "Referências — Elaboração",
      trata: "Como escrever cada tipo de referência e como apresentar a lista (alinhamento, espaçamento e ordenação)." },
    { codigo: "NBR 6024:2012", titulo: "Numeração progressiva das seções de um documento — Apresentação",
      trata: "Numeração das seções (1, 1.1, 1.1.1...), até a seção quinária, e destaque gráfico dos títulos de cada nível." },
    { codigo: "NBR 6027:2012", titulo: "Sumário — Apresentação",
      trata: "Como montar o sumário: títulos e numeração iguais aos do texto, com a página de início de cada seção." },
    { codigo: "NBR 6028:2021", titulo: "Resumo, resenha e recensão — Apresentação",
      trata: "Resumo em parágrafo único, extensão por tipo de documento e palavras-chave logo abaixo." },
    { codigo: "NBR 10520:2023", titulo: "Citações em documentos — Apresentação",
      trata: "Citações diretas e indiretas, citações longas e o formato da chamada de autoria (desde 2023, sobrenome só com a inicial maiúscula)." },
  ];

  // Para cada regra: a norma em que se baseia e a exigência explicada em poucas palavras.
  // Funções recebem `artigo` (true quando o documento foi verificado como artigo).
  const REGRAS = {
    estrutura: {
      norma: (artigo) => (artigo ? "NBR 6022:2018" : "NBR 14724:2011, 4"),
      explica: (artigo) => (artigo
        ? "O artigo tem como obrigatórios: título, autoria, resumo no idioma do texto, introdução, desenvolvimento, considerações finais e referências. O resumo em outro idioma é opcional na norma."
        : "Entre os elementos obrigatórios da monografia estão folha de rosto, resumo, resumo em língua estrangeira, sumário, introdução, desenvolvimento, conclusão e referências."),
    },
    papel: { norma: "NBR 14724:2011, 5.1", explica: "O trabalho é apresentado em papel A4 (21 × 29,7 cm)." },
    margens: {
      norma: "NBR 14724:2011, 5.1",
      explica: "Margens de 3 cm na parte superior e na esquerda e de 2 cm na inferior e na direita. Em páginas em paisagem, as margens giram junto com a página.",
    },
    paginacao: {
      norma: "NBR 14724:2011, 5.3",
      explica: "As folhas são contadas a partir da folha de rosto, mas o número só aparece a partir da primeira folha do texto (Introdução), no canto superior direito.",
    },
    "fonte-tamanho": {
      norma: "NBR 14724:2011, 5.1",
      explica: "Fonte tamanho 12 em todo o texto. Citações longas, notas de rodapé, paginação, legendas e fontes das ilustrações e tabelas usam tamanho menor e uniforme.",
    },
    "fonte-familia": {
      norma: "Recomendação",
      explica: "A ABNT não define qual fonte usar, mas o trabalho deve ter uma só. Arial e Times New Roman são as mais aceitas pelas instituições.",
    },
    espacamento: {
      norma: "NBR 14724:2011",
      explica: "O texto usa espaçamento de 1,5 entre linhas. Citações longas, notas de rodapé, referências, legendas e a natureza do trabalho usam espaçamento simples.",
    },
    alinhamento: {
      norma: "Prática recomendada",
      explica: "A ABNT não exige texto justificado, mas é o padrão adotado pelas instituições e pelos modelos.",
    },
    "citacao-longa": {
      norma: "NBR 10520:2023 e NBR 14724:2011",
      explica: "Citação direta com mais de três linhas: recuo de 4 cm da margem esquerda, fonte menor que a do texto, espaçamento simples e sem aspas.",
    },
    "titulo-pagina": {
      norma: (artigo) => (artigo ? "NBR 6022:2018" : "NBR 14724:2011"),
      explica: (artigo) => (artigo
        ? "No artigo o texto é corrido: as seções não começam em página nova."
        : "Seções primárias (1 INTRODUÇÃO, 2 ...) e títulos sem número (RESUMO, SUMÁRIO, REFERÊNCIAS...) começam em página nova, no alto da página."),
    },
    "titulo-numero": {
      norma: "NBR 6024:2012",
      explica: "As seções do texto são numeradas em algarismos arábicos (1, 1.1, 1.1.1...). Elementos como resumo, sumário, referências, apêndices e anexos não têm número; apêndices e anexos são identificados por letra, travessão e título.",
    },
    "titulo-centralizado": {
      norma: "NBR 14724:2011",
      explica: "Títulos sem indicativo numérico (errata, agradecimentos, listas, resumos, sumário, referências, glossário, apêndices, anexos e índices) são centralizados.",
    },
    "titulo-hierarquia": {
      norma: "NBR 6024:2012",
      explica: "Uma seção só existe dentro da anterior: 1.1.1 fica dentro de 1.1, que fica dentro de 1. A numeração vai até a seção quinária (1.1.1.1.1).",
    },
    "titulo-esquerda": {
      norma: "NBR 6024:2012",
      explica: "O título da seção vem depois do número, separado por um espaço, e fica alinhado à margem esquerda.",
    },
    "titulo-tamanho": {
      norma: "NBR 14724:2011, 5.1",
      explica: "Os títulos usam o mesmo tamanho 12 do texto. A diferença entre os níveis é feita com negrito, itálico ou maiúsculas, igual no texto e no sumário.",
    },
    referencias: {
      norma: "NBR 6023:2018",
      explica: "As referências ficam alinhadas à margem esquerda, em espaçamento simples, separadas umas das outras por uma linha em branco de espaço simples, e ordenadas alfabeticamente pela entrada (sobrenome ou instituição).",
    },
    "citacao-2023": {
      norma: "NBR 10520:2023",
      explica: "Desde a versão de 2023, o sobrenome do autor na citação é escrito só com a inicial maiúscula, tanto no texto quanto entre parênteses: (Silva, 2020). Siglas de instituições continuam em maiúsculas.",
    },
    ilustracoes: {
      norma: "NBR 14724:2011, 5.8 e 5.9",
      explica: "A identificação da ilustração ou tabela (Figura 1 – Título) fica acima dela, e a fonte (Fonte: Autor, ano) fica abaixo, ambas em tamanho menor. Tabelas seguem as Normas de apresentação tabular do IBGE.",
    },
    notas: {
      norma: "NBR 14724:2011",
      explica: "Notas de rodapé ficam dentro da margem, em fonte menor que a do texto e espaçamento simples.",
    },
    resumo: {
      norma: "NBR 6028:2021",
      explica: (artigo) => `Resumo em parágrafo único, com ${artigo ? "100 a 250 palavras em artigos" : "150 a 500 palavras em trabalhos acadêmicos"}, seguido das palavras-chave, separadas por ponto e vírgula e finalizadas por ponto.`,
    },
    modelo: {
      norma: "Do modelo",
      explica: "As caixas amarelas e os textos entre colchetes são orientações do modelo e devem sair antes da entrega.",
    },
  };
  const valor = (x, artigo) => (typeof x === "function" ? x(artigo) : x);

  // ---------------------------------------------------------------- verificação
  function analisar(arquivos, DOMParserCls, opcoes = {}) {
    const parser = new DOMParserCls();
    const xml = (nome) => (arquivos[nome] ? parser.parseFromString(arquivos[nome], "application/xml") : null);
    const docXml = xml("word/document.xml");
    if (!docXml) throw new Error("O arquivo não parece ser um documento .docx válido.");

    const tema = lerTema(xml("word/theme/theme1.xml"));
    const est = lerEstilos(xml("word/styles.xml"), tema);
    const blocos = lerCorpo(docXml, est, tema);
    // título "SUMÁRIO" (ou "LISTA DE ...") escrito com o estilo "Sumário 1": é o título, não uma entrada
    blocos.forEach((b) => { if (b.emSumario && b.tipo === "p" && /^(SUM[ÁA]RIO|LISTA DE [^\d]+)$/.test(normaliza(b.texto || ""))) b.emSumario = false; });
    const ps = blocos.filter((b) => b.tipo === "p");
    const rel = novoRelatorio();

    // ----- partes do trabalho
    const idx = (re, desde = 0) => blocos.findIndex((b, i) => i >= desde && b.tipo === "p" && !b.emSumario && b.texto.length <= 80 && re.test(normaliza(b.texto)));
    const iIntro = idx(/^INTRODU[ÇC][ÃA]O$/);
    if (iIntro >= 0) blocos.forEach((b, i) => { if (i >= iIntro) b.posIntro = true; });
    const iRef = idx(/^REFER[ÊE]NCIAS( BIBLIOGR[ÁA]FICAS)?$/, Math.max(iIntro, 0));
    let iFimRef = iRef < 0 ? -1 : blocos.findIndex((b, i) => i > iRef && b.tipo === "p" && ehTitulo(b) && !b.emSumario);
    if (iRef >= 0 && iFimRef < 0) iFimRef = blocos.length;
    // Artigo (NBR 6022) ou monografia (NBR 14724)? O artigo não tem sumário e chega
    // rápido à introdução; a monografia tem capa, folha de rosto e outros pré-textuais antes.
    const temSumario = ps.some((b) => b.texto.length <= 80 && /^SUM[ÁA]RIO$/.test(normaliza(b.texto)));
    const antesDaIntro = iIntro < 0 ? Infinity : blocos.slice(0, iIntro).filter((b) => b.tipo === "p" && b.texto && !b.instrucao).length;
    const tipo = opcoes.tipo === "artigo" || opcoes.tipo === "monografia" ? opcoes.tipo : (!temSumario && antesDaIntro <= 40 ? "artigo" : "monografia");
    const artigo = tipo === "artigo";
    const inicioTexto = iIntro >= 0 ? iIntro : 0;
    const fimTexto = iRef >= 0 ? iRef : blocos.length;
    const naParteTextual = (i) => i >= inicioTexto && i < fimTexto;

    // ----- 1. elementos obrigatórios
    {
      const r = rel.regra("estrutura", "Elementos obrigatórios", artigo ? "NBR 6022:2018, 4" : "NBR 14724, 4");
      const presentes = (re) => ps.some((b) => !b.emSumario && b.texto.length <= 80 && re.test(normaliza(b.texto)));
      // no artigo, o resumo em outro idioma é opcional e não há sumário
      const exigidos = artigo
        ? [["Resumo", /^RESUMO$/], ["Introdução", /^INTRODU[ÇC][ÃA]O$/], ["Referências", /^REFER[ÊE]NCIAS/]]
        : [
          ["Resumo", /^RESUMO$/], ["Resumo em língua estrangeira (Abstract)", /^(ABSTRACT|RESUMEN|RÉSUMÉ|RESUME)$/],
          ["Sumário", /^SUM[ÁA]RIO$/], ["Introdução", /^INTRODU[ÇC][ÃA]O$/], ["Referências", /^REFER[ÊE]NCIAS/],
        ];
      for (const [nome, re] of exigidos) if (!presentes(re)) r.problema("erro", nome, "não encontrado no documento");
      r.msg(r.it.nivel === "ok" ? (artigo ? "Resumo, introdução e referências encontrados." : "Resumo, abstract, sumário, introdução e referências encontrados.") : "Faltam elementos obrigatórios (ou o título está escrito de outra forma).");
    }

    // ----- 2. papel e margens
    {
      const rp = rel.regra("papel", "Papel A4", "NBR 14724, 5.1");
      const rm = rel.regra("margens", "Margens (3 cm superior/esquerda, 2 cm inferior/direita)", "NBR 14724, 5.1");
      const secoes = [...blocos.filter((b) => b.sectPr).map((b) => b.sectPr), ...blocos.filter((b) => b.tipo === "secao").map((b) => b.sectPr)];
      secoes.forEach((s, i) => {
        const nome = secoes.length > 1 ? `Seção ${i + 1} do documento` : "Documento";
        const pg = filho(s, "pgSz");
        const w = num(pg, "w"), h = num(pg, "h");
        const paisagem = at(pg, "orient") === "landscape" || (w > h);
        const [lw, lh] = paisagem ? [h, w] : [w, h];
        if (w !== undefined && !(perto(lw, 11906, 40) && perto(lh, 16838, 40))) rp.problema("erro", nome, `papel de ${cm(lw)} × ${cm(lh)} (A4 = 21 × 29,7 cm)`);
        const m = filho(s, "pgMar");
        if (!m) return;
        const v = { sup: num(m, "top"), inf: num(m, "bottom"), esq: num(m, "left"), dir: num(m, "right") };
        const esperado = { sup: 3 * CM, inf: 2 * CM, esq: 3 * CM, dir: 2 * CM };
        const confere = (alvo) => Object.entries(alvo).every(([k, x]) => perto(Math.abs(v[k]), x, 30));
        // Página em paisagem: as margens podem girar junto com a página (a margem de 3 cm
        // da encadernação passa para cima ou para baixo).
        if (paisagem && (confere({ sup: 3 * CM, dir: 3 * CM, inf: 2 * CM, esq: 2 * CM }) || confere({ inf: 3 * CM, esq: 3 * CM, sup: 2 * CM, dir: 2 * CM }))) return;
        const ruins = Object.entries(esperado).filter(([k, alvo]) => !perto(Math.abs(v[k]), alvo, 30))
          .map(([k, alvo]) => `${{ sup: "superior", inf: "inferior", esq: "esquerda", dir: "direita" }[k]} ${cm(Math.abs(v[k]))} (deveria ser ${cm(alvo)})`);
        if (ruins.length) rm.problema("erro", nome, ruins.join("; "));
      });
      rp.msg(rp.it.nivel === "ok" ? "Todas as páginas em A4." : "Configure o papel em Layout > Tamanho > A4.");
      rm.msg(rm.it.nivel === "ok" ? "Margens corretas." : "Ajuste em Layout > Margens > Margens personalizadas.");
    }

    // ----- 3. numeração de páginas
    {
      const r = rel.regra("paginacao", "Número da página no canto superior direito", "NBR 14724, 5.3");
      const temPagina = (nome) => {
        const d = xml(nome);
        if (!d) return null;
        const instr = todos(d.documentElement, "instrText").map((x) => x.textContent).join(" ") + " " +
          todos(d.documentElement, "fldSimple").map((x) => at(x, "instr") || "").join(" ");
        if (!/\bPAGE\b/i.test(instr)) return null;
        const p = todos(d.documentElement, "p").find((pp) => /\bPAGE\b/i.test(todos(pp, "instrText").map((x) => x.textContent).join(" ") + todos(pp, "fldSimple").map((x) => at(x, "instr") || "").join(" ")));
        const ap = p ? analisarParagrafo(p, est, { tema }) : null;
        return { alinhamento: ap ? ap.props.alinhamento || "esquerda" : "?", posTab: p && todos(p, "ptab").length > 0 };
      };
      const cab = Object.keys(arquivos).filter((n) => /^word\/header\d*\.xml$/.test(n)).map(temPagina).filter(Boolean);
      const rod = Object.keys(arquivos).filter((n) => /^word\/footer\d*\.xml$/.test(n)).map(temPagina).filter(Boolean);
      if (!cab.length && !rod.length) r.problema("erro", "Cabeçalho", "o documento não tem número de página");
      if (rod.length) r.problema("erro", "Rodapé", "o número de página está no rodapé; a ABNT pede no cabeçalho (canto superior direito)");
      for (const c of cab) if (c.alinhamento !== "direita" && !c.posTab) r.problema("erro", "Cabeçalho", `número de página ${alinhado(c.alinhamento)}; deve ficar à direita`);
      r.msg(r.it.nivel === "ok" ? "Número de página no cabeçalho, alinhado à direita. Lembre que ele só deve aparecer a partir da Introdução." : "Insira o número em Inserir > Número de Página > Início da Página > Número simples 3.");
    }

    // ----- o que fica entre a legenda e a "Fonte:" é a própria ilustração
    blocos.forEach((b, i) => {
      if (b.tipo !== "p" || !LEGENDA.test(b.texto) || b.texto.length > 160) return;
      const j = blocos.slice(i + 1, i + 5).findIndex((x) => x.tipo === "p" && /^fonte\s*:/i.test(x.texto));
      if (j > 0) blocos.slice(i + 1, i + 1 + j).forEach((x) => { x.ilustracao = true; });
    });

    // ----- parágrafos de texto comum da parte textual
    const corpo = [];
    const citacoesLongas = [];
    blocos.forEach((b, i) => {
      if (b.tipo !== "p" || !naParteTextual(i) || !b.texto || b.emTabela || b.emSumario || b.instrucao) return;
      if (ehTitulo(b) || LEGENDA.test(b.texto) || /^fonte\s*:/i.test(b.texto) || b.desenho || b.ilustracao) return;
      const recuo = b.props.recuoEsq || 0;
      if (recuo >= 1.5 * CM && !b.props.numerado && b.props.alinhamento !== "centralizado") citacoesLongas.push(b);
      else if (b.props.numerado) return; // itens de lista: espaçamento livre
      else corpo.push(b);
    });

    // ----- 4. fonte 12
    {
      const r = rel.regra("fonte-tamanho", "Fonte tamanho 12 no texto", "NBR 14724, 5.1");
      for (const b of corpo) {
        const t = tamanhos(b).filter((x) => x !== 12);
        if (t.length) r.problema("erro", b, `tamanho ${t.join(", ").replace(/\./g, ",")}`, { formato: { tamanho: 12 } });
      }
      r.msg(r.it.nivel === "ok" ? `Os ${corpo.length} parágrafos de texto estão em tamanho 12.` : "O texto deve estar em tamanho 12 (exceto citações longas, notas, legendas e fontes, que usam tamanho menor).");
    }

    // ----- 5. família da fonte uniforme
    {
      const r = rel.regra("fonte-familia", "Mesma fonte em todo o texto", "NBR 14724, 5.1 (recomendação)");
      const cont = {};
      let total = 0;
      for (const b of ps) {
        if (b.emSumario || b.instrucao) continue;
        for (const x of b.runs) {
          const n = x.texto.replace(/\s/g, "").length;
          if (!n) continue;
          const f = x.props.fonte || "Times New Roman";
          cont[f] = (cont[f] || 0) + n; total += n;
        }
      }
      const usadas = Object.entries(cont).sort((a, b) => b[1] - a[1]).filter(([, n]) => n / total > 0.01);
      const principal = usadas.length ? usadas[0][0] : null;
      // cada parágrafo com outra fonte, para poder ir até ele e corrigir
      let qtd = 0;
      for (const b of ps) {
        if (!principal || b.emSumario || b.instrucao || !b.texto.trim()) continue;
        const outras = [...new Set(b.runs.filter((x) => x.texto.trim() && (x.props.fonte || "Times New Roman") !== principal).map((x) => x.props.fonte || "Times New Roman"))];
        if (outras.length && qtd++ < 300) r.problema("aviso", b, `fonte ${outras.join(", ")} (o resto do trabalho usa ${principal})`, { formato: { fonte: principal } });
      }
      if (principal && !/^(Arial|Times New Roman)$/i.test(principal)) r.problema("aviso", principal, "fonte pouco usual em trabalhos acadêmicos; prefira Arial ou Times New Roman");
      if (r.it.nivel !== "ok") r.msg(`A norma não fixa a família, mas o trabalho deve usar uma só${principal ? ` (a principal aqui é ${principal})` : ""}. Arial e Times New Roman são as mais aceitas.`);
      else r.msg(principal ? `Fonte usada: ${principal}.` : "");
    }

    // ----- 6. espaçamento 1,5
    {
      const r = rel.regra("espacamento", "Espaçamento 1,5 entre linhas no texto", "NBR 14724, 5.2");
      for (const b of corpo) if (!linha15(b.props)) r.problema("erro", b, `espaçamento ${descreveLinha(b.props)}`, { formato: { linha: 1.5 } });
      r.msg(r.it.nivel === "ok" ? "Texto com espaçamento 1,5." : "Selecione o texto > Parágrafo > Espaçamento entre linhas: 1,5 linha.");
    }

    // ----- 7. alinhamento justificado
    {
      const r = rel.regra("alinhamento", "Texto justificado", "Prática recomendada");
      for (const b of corpo) if (b.props.alinhamento !== "justificado") r.problema("aviso", b, alinhado(b.props.alinhamento), { formato: { alinhamento: "justificado" } });
      r.msg(r.it.nivel === "ok" ? "Texto justificado." : "A NBR 14724 não exige, mas o texto justificado é o padrão adotado pelas instituições.");
    }

    // ----- 8. citações longas
    {
      const r = rel.regra("citacao-longa", "Citações com mais de 3 linhas", "NBR 10520:2023 e NBR 14724, 5.2");
      for (const b of citacoesLongas) {
        const prob = [];
        if (!perto(b.props.recuoEsq, 4 * CM, 85)) prob.push(`recuo de ${cm(b.props.recuoEsq)} (deveria ser 4 cm)`);
        const t = tamanhos(b);
        if (t.some((x) => x >= 12)) prob.push(`fonte tamanho ${t.join(", ")} (deveria ser menor que 12, ex.: 10)`);
        if (!linhaSimples(b.props)) prob.push(`espaçamento ${descreveLinha(b.props)} (deveria ser simples)`);
        if (/^[“"]/.test(b.texto)) prob.push("citação longa não leva aspas (apague as aspas à mão)");
        if (prob.length) r.problema("erro", b, prob.join("; "), { formato: { recuoEsq: 4, primeiraLinha: 0, tamanho: 10, linha: 1 } });
      }
      if (!citacoesLongas.length) r.info("Nenhuma citação longa encontrada (parágrafo com recuo à esquerda).");
      else r.msg(r.it.nivel === "ok" ? `${citacoesLongas.length} citação(ões) longa(s) formatada(s) corretamente.` : "Use recuo de 4 cm, fonte menor (10), espaçamento simples e sem aspas. No modelo: estilo “Citação longa”.");
    }

    // ----- 9 a 11. títulos
    {
      const rNova = rel.regra("titulo-pagina", "Seções primárias começam em nova página", "NBR 14724, 5.2.1");
      const rNum = rel.regra("titulo-numero", "Numeração progressiva das seções", "NBR 6024:2012");
      const rCentro = rel.regra("titulo-centralizado", "Títulos sem número centralizados", "NBR 14724, 5.2.2");
      const rHier = rel.regra("titulo-hierarquia", "Hierarquia dos títulos", "NBR 6024:2012");
      const rEsq = rel.regra("titulo-esquerda", "Títulos numerados alinhados à esquerda", "NBR 6024:2012, 4.1, e NBR 14724, 5.2.2");
      const rTam = rel.regra("titulo-tamanho", "Títulos em tamanho 12", "NBR 14724, 5.1");

      // tipos das seções para saber se a quebra de seção muda de página
      const sects = [...blocos.filter((b) => b.sectPr).map((b) => b.sectPr), ...blocos.filter((b) => b.tipo === "secao").map((b) => b.sectPr)];
      let iSect = 0, quebra = true, nivelAnterior = 0;
      blocos.forEach((b, i) => {
        if (b.tipo !== "p") return;
        if (b.quebraInicio) quebra = true;
        if (b.texto && !b.emSumario && !b.instrucao) {
          const titulo = ehTitulo(b) && b.texto.length <= 150;
          const n = normaliza(b.texto);
          const semNumero = semNumeroTitulo(b);
          if (titulo) {
            const t = tamanhos(b).filter((x) => x !== 12);
            if (t.length) rTam.problema("erro", b, `tamanho ${t.join(", ").replace(/\./g, ",")} (todo o trabalho, inclusive títulos, usa tamanho 12)`, { formato: { tamanho: 12 } });
          }
          if (!artigo && titulo && (b.nivel === 1 || semNumero) && !(b.props.quebraAntes || quebra) && i > 0) {
            rNova.problema("erro", b, "não começa em nova página", { quebraAntes: true });
          }
          if (titulo && semNumero) {
            if (temNumero(b)) rNum.problema("erro", b, "este título não deve ter número");
            const ap = b.texto.normalize("NFC").match(/^((?:AP[ÊE]NDICE|ANEXO|Ap[êe]ndice|Anexo) [A-Z])\s*[:—-]\s*/);
            if (ap) rNum.problema("aviso", b, "separe a letra do título com travessão: “APÊNDICE A – Título”", { texto: { de: ap[0], para: ap[1] + " – " }, individual: true });
            if (b.props.alinhamento !== "centralizado") rCentro.problema("erro", b, `${alinhado(b.props.alinhamento)} (deve ser centralizado)`, { formato: { alinhamento: "centralizado", primeiraLinha: 0, recuoEsq: 0 } });
            nivelAnterior = 0;
          } else if (titulo && b.nivel && naParteTextual(i)) {
            if (!temNumero(b)) rNum.problema("erro", b, "seção sem número");
            if (b.props.alinhamento === "centralizado" || b.props.alinhamento === "direita") rEsq.problema("erro", b, `${alinhado(b.props.alinhamento)} (título numerado fica à esquerda)`, { formato: { alinhamento: "esquerda" } });
            if (b.nivel > nivelAnterior + 1) rHier.problema("aviso", b, `pula do nível ${nivelAnterior} para o ${b.nivel}`);
            nivelAnterior = b.nivel;
          }
          quebra = false;
        }
        if (b.quebraFim) quebra = true;
        if (b.sectPr) {
          iSect++;
          const prox = sects[iSect];
          const tipo = prox ? at(filho(prox, "type"), "val") : null;
          if (tipo !== "continuous") quebra = true;
        }
      });
      if (artigo) rNova.info("No artigo (NBR 6022) o texto é corrido: as seções não começam em página nova.");
      else rNova.msg(rNova.it.nivel === "ok" ? "Seções primárias e títulos sem número começam em página nova." : "Use Inserir > Quebra de Página antes do título, ou ative “Quebra de página antes” no estilo Título 1.");
      rNum.msg(rNum.it.nivel === "ok" ? "Seções do texto numeradas; títulos de elementos pré e pós-textuais sem número." : rNum.it.nivel === "aviso" ? "Apêndices e anexos são identificados por letra, travessão e título." : "As seções do texto são numeradas (1, 1.1, 1.1.1...). Resumo, sumário, referências, apêndices e anexos não levam número.");
      rCentro.msg(rCentro.it.nivel === "ok" ? "Títulos sem número centralizados." : "Resumo, abstract, listas, sumário, referências, apêndices e anexos têm título centralizado.");
      rEsq.msg(rEsq.it.nivel === "ok" ? "Títulos numerados alinhados à esquerda." : "Títulos com número ficam na margem esquerda; só os títulos sem número (resumo, sumário, referências…) são centralizados.");
      rTam.msg(rTam.it.nivel === "ok" ? "Títulos em tamanho 12." : "Os títulos usam o mesmo tamanho do texto (12); destaque-os com negrito, maiúsculas ou itálico.");
      rHier.msg(rHier.it.nivel === "ok" ? "Os níveis dos títulos seguem a ordem." : "Uma seção terciária (1.1.1) deve estar dentro de uma secundária (1.1).");
    }

    // ----- 12. referências
    {
      const r = rel.regra("referencias", "Lista de referências", "NBR 6023:2018, 6");
      if (iRef < 0) r.info("Seção REFERÊNCIAS não encontrada.");
      else {
        const refs = blocos.slice(iRef + 1, iFimRef).filter((b) => b.tipo === "p" && b.texto && !b.instrucao && !b.emTabela);
        let anterior = null, ordemOk = true, foraDeOrdem = null;
        for (let k = 0; k < refs.length; k++) {
          const b = refs[k];
          const prob = [];
          if (b.props.alinhamento && b.props.alinhamento !== "esquerda") prob.push(`${alinhado(b.props.alinhamento).replace("alinhado", "alinhada")} (deve ser alinhada à esquerda)`);
          if (!linhaSimples(b.props)) prob.push(`espaçamento ${descreveLinha(b.props)} (deve ser simples)`);
          if ((b.props.recuoEsq || 0) > 60 || (b.props.primeiraLinha || 0) > 60) prob.push("tem recuo (deve começar na margem esquerda)");
          if (prob.length) r.problema("erro", b, prob.join("; "), { formato: FORMATO_REF });
          if (anterior && entradaRef(anterior.texto).localeCompare(entradaRef(b.texto), "pt-BR", { sensitivity: "base" }) > 0 && ordemOk) {
            foraDeOrdem = b;
            ordemOk = false;
          }
          anterior = b;
        }
        if (foraDeOrdem) {
          // Correção: reescrever cada parágrafo da lista com a referência que deve ocupar aquela
          // posição, preservando negrito, itálico, fonte e tamanho de cada trecho.
          const chave = (b) => entradaRef(b.texto);
          const ordenadas = [...refs].sort((a, b) => chave(a).localeCompare(chave(b), "pt-BR", { sensitivity: "base" })
            || a.texto.localeCompare(b.texto, "pt-BR", { sensitivity: "base" }));
          const segmentos = (b) => b.runs.reduce((acc, x) => {
            const seg = { t: x.texto, b: !!x.props.negrito, i: !!x.props.italico, f: x.props.fonte, s: x.props.tamanho, sup: !!x.props.sobrescrito };
            const ult = acc[acc.length - 1];
            if (ult && ult.b === seg.b && ult.i === seg.i && ult.f === seg.f && ult.s === seg.s && ult.sup === seg.sup) ult.t += seg.t;
            else acc.push(seg);
            return acc;
          }, []);
          r.problema("erro", foraDeOrdem, `fora da ordem alfabética (a lista tem ${refs.length} referências)`, {
            individual: true,
            ordenar: {
              alvos: refs.map((b) => ({ posicao: b.posicao, trecho: b.texto.slice(0, 110) })),
              conteudo: ordenadas.map(segmentos),
            },
          });
        }
        // separação entre referências: espaço depois ou parágrafo vazio
        const bl = blocos.slice(iRef + 1, iFimRef).filter((b) => b.tipo === "p" && !b.instrucao);
        const separadas = refs.length < 2 || refs.every((b) => (b.props.depois || 0) >= 180) || bl.some((b) => !b.texto);
        if (!separadas) for (const b of refs) if ((b.props.depois || 0) < 180) r.problema("aviso", b, "falta a linha em branco depois da referência", { formato: { depois: 12 } });
        r.msg(r.it.nivel === "ok" ? `${refs.length} referência(s): alinhadas à esquerda, espaçamento simples, em ordem alfabética.` : "Referências: alinhadas à esquerda, espaçamento simples, separadas por uma linha em branco e em ordem alfabética.");
      }
    }

    // ----- 13. citações no formato de 2023
    {
      const r = rel.regra("citacao-2023", "Sobrenome em citações (só a inicial maiúscula)", "NBR 10520:2023");
      const reParen = /\(([A-ZÀ-Ý][A-ZÀ-Ý'’-]{2,}(?:\s+(?:[A-ZÀ-Ý][A-ZÀ-Ý'’-]+|DE|DA|DO|DOS|DAS))*)(?:\s+et al\.)?(?:;\s*[^,()]+?)?,\s*(?:\d{4}|\[?s\.\s?d\.\]?)/g;
      const reNarr = /\b([A-ZÀ-Ý]{3,}(?:\s+[A-ZÀ-Ý]{2,})*)(?:\s+et al\.)?\s+\((\d{4})/g;
      blocos.forEach((b, i) => {
        if (b.tipo !== "p" || !naParteTextual(i) || b.emSumario || b.instrucao) return;
        const achados = [...b.texto.matchAll(reParen), ...b.texto.matchAll(reNarr)];
        for (const m of achados) {
          const para = m[0].replace(m[1], capitalizaNome(m[1]));
          r.problema("aviso", { texto: m[0], posicao: b.posicao }, "se for sobrenome, escreva só com a inicial maiúscula, ex.: (Silva, 2020). Siglas (IBGE, ABNT) podem continuar em maiúsculas.",
            para !== m[0] ? { texto: { de: m[0], para }, individual: true } : undefined);
        }
      });
      r.msg(r.it.nivel === "ok" ? "Nenhuma citação com sobrenome todo em maiúsculas." : "Desde 2023, o sobrenome nas citações vai só com a inicial maiúscula, inclusive entre parênteses.");
    }

    // ----- 14. legendas e fontes de ilustrações e tabelas
    {
      const r = rel.regra("ilustracoes", "Ilustrações e tabelas: título acima, fonte abaixo", "NBR 14724, 5.8 e 5.9");
      let qtd = 0;
      blocos.forEach((b, i) => {
        if (b.tipo !== "p" || b.emSumario || b.instrucao || !LEGENDA.test(b.texto) || b.texto.length > 220 || !naParteTextual(i)) return;
        // é uma menção no meio do texto? legendas costumam ser curtas e sem ponto final longo
        if (b.texto.length > 160 && /[.]\s+\S/.test(b.texto)) return;
        qtd++;
        const prox = blocos.slice(i + 1, i + 4);
        const ilustracao = prox.findIndex((x) => x.tipo === "tabela" || (x.tipo === "p" && x.desenho));
        const anterior = blocos.slice(Math.max(0, i - 3), i).some((x) => x.tipo === "fimTabela" || (x.tipo === "p" && x.desenho));
        if (ilustracao < 0 && anterior) r.problema("erro", b, "a identificação está abaixo da ilustração; deve ficar acima");
        const depois = ilustracao < 0 ? i : i + 1 + ilustracao;
        let j = depois + 1;
        if (blocos[depois] && blocos[depois].tipo === "tabela") while (j < blocos.length && blocos[j].tipo !== "fimTabela") j++;
        const temFonte = blocos.slice(j, j + 4).some((x) => x.tipo === "p" && /^fonte\s*:/i.test(x.texto));
        if (!temFonte) r.problema("erro", b, "falta a indicação “Fonte:” logo abaixo");
        if (tamanhos(b).some((x) => x >= 12)) r.problema("aviso", b, "legenda em tamanho 12; a norma pede tamanho menor (ex.: 10)", { formato: { tamanho: 10 } });
      });
      if (!qtd) r.info("Nenhuma ilustração ou tabela com legenda encontrada.");
      else r.msg(r.it.nivel === "ok" ? `${qtd} ilustração(ões)/tabela(s) com título acima e fonte abaixo.` : "Acima: “Figura 1 – Título”. Abaixo: “Fonte: Autor (ano).” Ambos em tamanho menor que o texto.");
    }

    // ----- 15. notas de rodapé
    {
      const r = rel.regra("notas", "Notas de rodapé em fonte menor e espaçamento simples", "NBR 14724, 5.2");
      const fn = xml("word/footnotes.xml");
      let qtd = 0;
      if (fn) {
        for (const nota of todos(fn.documentElement, "footnote")) {
          if (["separator", "continuationSeparator", "continuationNotice"].includes(at(nota, "type"))) continue;
          for (const p of filhos(nota, "p")) {
            const b = analisarParagrafo(p, est, { tema });
            if (!b.texto) continue;
            qtd++;
            const prob = [];
            if (tamanhos(b).some((x) => x >= 12)) prob.push(`tamanho ${tamanhos(b).join(", ")} (deveria ser menor, ex.: 10)`);
            if (!linhaSimples(b.props)) prob.push(`espaçamento ${descreveLinha(b.props)} (deveria ser simples)`);
            if (prob.length) r.problema("erro", b, prob.join("; "));
          }
        }
      }
      if (!qtd) r.info("Nenhuma nota de rodapé encontrada.");
      else r.msg(r.it.nivel === "ok" ? `${qtd} nota(s) de rodapé formatada(s) corretamente.` : "Notas de rodapé usam fonte menor (10) e espaçamento simples.");
    }

    // ----- 16. resumo
    {
      const r = rel.regra("resumo", `Resumo: parágrafo único, ${artigo ? "100 a 250" : "150 a 500"} palavras, palavras-chave`, "NBR 6028:2021");
      const iRes = idx(/^RESUMO$/);
      if (iRes < 0) r.info("Resumo não encontrado.");
      else {
        const fim = blocos.findIndex((b, i) => i > iRes && b.tipo === "p" && (ehTitulo(b) || /^palavras?[- ]chave/i.test(b.texto)));
        const pars = blocos.slice(iRes + 1, fim < 0 ? undefined : fim).filter((b) => b.tipo === "p" && b.texto && !b.instrucao);
        const palavras = pars.map((b) => b.texto.split(/\s+/).filter(Boolean).length).reduce((a, n) => a + n, 0);
        if (pars.length > 1) r.problema("aviso", pars[1], `o resumo tem ${pars.length} parágrafos; deve ser parágrafo único`);
        const [min, max] = artigo ? [100, 250] : [150, 500];
        if (palavras < min || palavras > max) r.problema("aviso", pars[0] || "RESUMO", `${palavras} palavras (${artigo ? "em artigos, de 100 a 250" : "o ideal para trabalhos acadêmicos é de 150 a 500"})`);
        const kw = fim >= 0 && /^palavras?[- ]chave/i.test(blocos[fim].texto) ? blocos[fim] : null;
        if (!kw) r.problema("erro", "RESUMO", "faltam as palavras-chave logo abaixo do resumo");
        else {
          const lista = kw.texto.replace(/^palavras?[- ]chave\s*:\s*/i, "");
          const virgulas = lista.includes(",") && !lista.includes(";");
          const semPonto = !/\.$/.test(lista.trim());
          const de = lista.trim();
          let para = virgulas ? de.replace(/\s*,\s*/g, "; ") : de;
          if (semPonto) para += ".";
          const corr = de.length >= 3 && de.length <= 250 && para !== de ? { texto: { de, para }, individual: true } : undefined;
          if (virgulas) r.problema("aviso", kw, "separe as palavras-chave com ponto e vírgula", corr);
          if (semPonto) r.problema("aviso", kw, "finalize as palavras-chave com ponto", virgulas ? undefined : corr);
        }
        r.msg(r.it.nivel === "ok" ? `Resumo com ${palavras} palavras, em parágrafo único, com palavras-chave.` : `Resumo: parágrafo único, ${artigo ? "100 a 250" : "150 a 500"} palavras, seguido de “Palavras-chave:” separadas por ponto e vírgula e finalizadas por ponto.`);
      }
    }

    // ----- 17. sobras do modelo
    {
      const r = rel.regra("modelo", "Orientações do modelo apagadas", "—");
      const inst = ps.filter((b) => b.instrucao && b.texto);
      const marcadores = ps.filter((b) => !b.emSumario && /\[(texto|nome|t[íi]tulo|insira|cidade|ano|subt[íi]tulo|conte[úu]do)[^\]]*\]/i.test(b.texto));
      if (inst.length) r.problema("aviso", inst[0], `${inst.length} caixa(s) amarela(s) de orientação ainda no documento`);
      for (const b of marcadores.slice(0, 50)) r.problema("aviso", b, "texto de exemplo entre colchetes ainda não substituído");
      r.msg(r.it.nivel === "ok" ? "Nenhuma sobra do modelo." : "Apague as caixas amarelas e substitua os textos entre colchetes antes de entregar.");
    }

    // norma e explicação de cada regra (fonte única: REGRAS)
    for (const it of rel.itens) {
      const meta = REGRAS[it.id];
      if (!meta) continue;
      it.norma = valor(meta.norma, artigo);
      it.explicacao = valor(meta.explica, artigo);
      it.normas = NORMAS.map((n) => n.codigo).filter((c) => it.norma.includes(c));
    }
    const resumo = { erro: 0, aviso: 0, ok: 0, info: 0 };
    for (const it of rel.itens) resumo[it.nivel]++;
    return {
      resumo,
      itens: rel.itens,
      partes: { introducao: iIntro >= 0, referencias: iRef >= 0, paragrafosDeTexto: corpo.length },
      tipo, tipoDetectado: !(opcoes.tipo === "artigo" || opcoes.tipo === "monografia"),
    };
  }

  // Lê um .docx (ArrayBuffer/Buffer) com JSZip e devolve os XMLs relevantes.
  async function lerDocx(dados, JSZip) {
    const zip = await JSZip.loadAsync(dados);
    const arquivos = {};
    for (const nome of Object.keys(zip.files)) {
      if (/^word\/(document|styles|footnotes|numbering|header\d*|footer\d*|theme\/theme\d*)\.xml$/.test(nome)) {
        arquivos[nome] = await zip.files[nome].async("string");
      }
    }
    return arquivos;
  }

  // tipo: "auto" (padrão), "monografia" (NBR 14724) ou "artigo" (NBR 6022)
  async function verificar(dados, { JSZip, DOMParser, tipo }) {
    return analisar(await lerDocx(dados, JSZip), DOMParser, { tipo });
  }

  return { verificar, analisar, lerDocx, NORMAS, REGRAS };
});
