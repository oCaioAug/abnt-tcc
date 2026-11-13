/*
 * Núcleo do gerador de referências ABNT.
 *   NBR 6023:2025  – elaboração de referências (3. ed.)
 *   NBR 10520:2023 – citações (sobrenome só com inicial maiúscula)
 *
 * Sem dependências: funciona no navegador (global `ABNTRef`) e no Node
 * (require), para poder ser reaproveitado num suplemento do Word ou
 * numa extensão do LibreOffice no futuro.
 */
(function (raiz, fabrica) {
  if (typeof module === "object" && module.exports) module.exports = fabrica();
  else raiz.ABNTRef = fabrica();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  // ------------------------------------------------------------ utilidades
  const limpa = (s) => String(s == null ? "" : s).replace(/\s+/g, " ").trim();
  const maiusc = (s) => s.toLocaleUpperCase("pt-BR");
  const minusc = (s) => s.toLocaleLowerCase("pt-BR");
  const terminaPontuado = (s) => /[.?!]$/.test(s);

  const MESES = ["jan.", "fev.", "mar.", "abr.", "maio", "jun.", "jul.", "ago.", "set.", "out.", "nov.", "dez."];
  const MESES_EXTENSO = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto",
    "setembro", "outubro", "novembro", "dezembro"];

  // "julho/setembro" -> "jul./set."; "jul./set." fica igual
  function abreviaMeses(s) {
    return limpa(s).replace(/[A-Za-zÀ-ÿ]+\.?/g, (palavra) => {
      const base = minusc(palavra.replace(/\.$/, ""));
      const i = MESES_EXTENSO.findIndex((m) => m === base || m.slice(0, 3) === base);
      return i >= 0 ? MESES[i] : palavra;
    });
  }

  // "2026-11-12" -> "12 nov. 2026"
  function dataABNT(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(limpa(iso));
    if (!m) return limpa(iso);
    return `${Number(m[3])} ${MESES[Number(m[2]) - 1]} ${m[1]}`;
  }

  function hojeISO() {
    const d = new Date();
    const z = (n) => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${z(d.getMonth() + 1)}-${z(d.getDate())}`;
  }

  // ------------------------------------------------------------ montador
  // Uma referência é uma lista de segmentos { t, b (negrito), i (itálico) }.
  class Montador {
    constructor() { this.segs = []; }
    add(t, fmt) { if (t) this.segs.push({ t, ...(fmt || {}) }); return this; }
    texto() { return this.segs.map((s) => s.t).join(""); }
    // Fecha um elemento com ponto (sem duplicar) e prepara o próximo.
    // Elemento vazio (nada desde o último ponto) não gera nada.
    ponto() {
      const txt = this.texto();
      if (!txt.trim() || /\s$/.test(txt)) return this;
      if (!terminaPontuado(txt)) this.add(".");
      return this.add(" ");
    }
    fim() {
      this.ponto();
      while (this.segs.length && !this.segs[this.segs.length - 1].t.trim()) this.segs.pop();
      const ult = this.segs[this.segs.length - 1];
      if (ult) ult.t = ult.t.trimEnd();
      return this.segs;
    }
  }

  // ------------------------------------------------------------ autoria
  const SUFIXOS = /^(j[uú]nior|jr\.?|filho|filha|neto|neta|sobrinho|sobrinha)$/i;
  const PARTICULAS = /^(de|da|do|das|dos|e|del|della|di|du|van|von|der)$/i;

  // "Maria da Silva" -> { sobrenome: "Silva", prenome: "Maria da" }
  // "José Assis Júnior" -> { sobrenome: "Assis Júnior", prenome: "José" }
  // "Silva, Maria" (já invertido) é respeitado.
  function separaNome(nome) {
    nome = limpa(nome);
    const virgula = nome.indexOf(",");
    if (virgula >= 0) {
      return { sobrenome: limpa(nome.slice(0, virgula)), prenome: limpa(nome.slice(virgula + 1)) };
    }
    const p = nome.split(" ");
    if (p.length === 1) return { sobrenome: p[0], prenome: "" };
    const n = p.length > 2 && SUFIXOS.test(p[p.length - 1]) ? 2 : 1;
    return { sobrenome: p.slice(-n).join(" "), prenome: p.slice(0, -n).join(" ") };
  }

  // "ASSIS JÚNIOR" -> "Assis Júnior"  (forma usada nas citações, NBR 10520:2023)
  function nomeCitacao(s) {
    return limpa(s).split(" ").map((w, i) => {
      if (i > 0 && PARTICULAS.test(w)) return minusc(w);
      return w.split("-").map((x) => maiusc(x.charAt(0)) + minusc(x.slice(1))).join("-");
    }).join(" ");
  }

  function listaLinhas(s) {
    return String(s || "").split(/[\n;]+/).map(limpa).filter(Boolean);
  }

  // Retorna null quando não há autor (a entrada passa a ser o título).
  function autoria(autores, opcoes) {
    const linhas = listaLinhas(autores);
    if (!linhas.length) return null;
    const o = opcoes || {};
    let ref, cit;
    if (o.tipoAutor === "entidade") {
      ref = linhas.map(maiusc);
      cit = linhas;
    } else {
      const nomes = linhas.map(separaNome);
      ref = nomes.map((x) => maiusc(x.sobrenome) + (x.prenome ? ", " + x.prenome : ""));
      cit = nomes.map((x) => nomeCitacao(x.sobrenome));
    }
    let texto = o.etAl && ref.length > 3 ? ref[0] + " et al." : ref.join("; ");
    if (o.responsabilidade) texto += ` (${o.responsabilidade})`;
    return { texto, citacao: cit };
  }

  // Sem autoria: primeira palavra em maiúsculas (com o artigo, se houver).
  const ARTIGOS = /^(o|a|os|as|um|uma|uns|umas|the|an?|el|la|los|las|le|les)$/i;
  function palavrasIniciais(titulo) {
    const p = limpa(titulo).split(" ");
    return ARTIGOS.test(p[0]) && p.length > 1 ? 2 : 1;
  }
  function tituloComoEntrada(titulo) {
    const p = limpa(titulo).split(" ");
    const n = palavrasIniciais(titulo);
    return [maiusc(p.slice(0, n).join(" ")), ...p.slice(n)].join(" ");
  }

  // ------------------------------------------------------------ elementos
  function edicao(e) {
    e = limpa(e);
    if (!e || /^1(\.|ª|a)?(\s*ed\.?)?$/i.test(e)) return "";
    if (/^\d+$/.test(e)) return `${e}. ed`;
    if (/^\d+(ª|a)$/i.test(e)) return `${parseInt(e, 10)}. ed`;
    return e;
  }

  function imprenta(local, editora, ano) {
    local = limpa(local); editora = limpa(editora);
    const le = !local && !editora ? "[S. l.: s. n.]" : `${local || "[S. l.]"}: ${editora || "[s. n.]"}`;
    return `${le}, ${limpa(ano) || "[s. d.]"}`;
  }

  const paginas = (p) => {
    p = limpa(p);
    return !p ? "" : /^p\./i.test(p) ? p : `p. ${p.replace(/\s*[-–]\s*/, "-")}`;
  };

  function entrada(m, d, destaque) {
    const aut = autoria(d.autores, d);
    if (aut) {
      m.add(aut.texto).ponto();
      m.add(limpa(d.titulo), destaque ? { b: true } : null);
    } else {
      m.add(tituloComoEntrada(d.titulo));
    }
    if (limpa(d.subtitulo)) m.add(": " + limpa(d.subtitulo));
    m.ponto();
    return aut;
  }

  function online(m, d) {
    // NBR 6023:2025 (exemplos de 7.7.6): o DOI vem como endereço completo, antes do "Disponível em"
    const doi = limpa(d.doi).replace(/^(https?:\/\/)?(dx\.)?doi\.org\//i, "").replace(/^doi:\s*/i, "");
    if (doi) m.add("DOI: https://doi.org/" + doi).ponto();
    if (limpa(d.url)) {
      m.add("Disponível em: " + limpa(d.url)).ponto();
      m.add("Acesso em: " + dataABNT(d.acesso || hojeISO())).ponto();
    }
  }

  // ------------------------------------------------------------ tipos
  const montar = {
    livro(m, d) {
      const aut = entrada(m, d, true);
      m.add(edicao(d.edicao)).ponto();
      m.add(imprenta(d.local, d.editora, d.ano)).ponto();
      if (limpa(d.paginasTotal)) m.add(limpa(d.paginasTotal).replace(/^(\d+)$/, "$1 p")).ponto();
      online(m, d);
      return aut;
    },

    capitulo(m, d) {
      const aut = entrada(m, d, false);
      m.add("In", { i: true }).add(": ");
      const autLivro = autoria(d.autoresLivro, { tipoAutor: d.tipoAutorLivro, responsabilidade: d.responsabilidadeLivro });
      if (autLivro) m.add(autLivro.texto).ponto().add(limpa(d.tituloLivro), { b: true });
      else m.add(tituloComoEntrada(d.tituloLivro));
      if (limpa(d.subtituloLivro)) m.add(": " + limpa(d.subtituloLivro));
      m.ponto();
      m.add(edicao(d.edicao)).ponto();
      m.add(imprenta(d.local, d.editora, d.ano)).ponto();
      m.add(paginas(d.paginas)).ponto();
      online(m, d);
      return aut;
    },

    artigo(m, d) {
      const aut = entrada(m, d, false);
      const partes = [
        limpa(d.localPeriodico),
        limpa(d.volume) && `v. ${limpa(d.volume).replace(/^v\.\s*/i, "")}`,
        limpa(d.numero) && `n. ${limpa(d.numero).replace(/^n\.\s*/i, "")}`,
        // NBR 6023:2025, 7.7.6: artigos só eletrônicos usam a e-localização no lugar das páginas
        limpa(d.eLocalizador) || paginas(d.paginas),
        [abreviaMeses(d.mes), limpa(d.ano)].filter(Boolean).join(" "),
      ].filter(Boolean);
      m.add(limpa(d.periodico), { b: true });
      if (partes.length) m.add(", " + partes.join(", "));
      m.ponto();
      online(m, d);
      return aut;
    },

    trabalho(m, d) {
      const aut = entrada(m, d, true);
      m.add(limpa(d.ano)).ponto();
      if (limpa(d.folhas)) m.add(limpa(d.folhas).replace(/^(\d+)$/, "$1 f")).ponto();
      const tipo = limpa(d.tipoTrabalho) + (limpa(d.grau) ? ` (${limpa(d.grau)})` : "");
      const local = [limpa(d.instituicao), limpa(d.local), limpa(d.anoDefesa) || limpa(d.ano)].filter(Boolean).join(", ");
      m.add([tipo, local].filter(Boolean).join(" – ")).ponto();
      online(m, d);
      return aut;
    },

    evento(m, d) {
      const aut = entrada(m, d, false);
      const num = limpa(d.numeroEvento).replace(/^(\d+)$/, "$1.");
      m.add("In", { i: true }).add(": ");
      m.add([maiusc(limpa(d.nomeEvento)), num, limpa(d.anoEvento), limpa(d.localEvento)].filter(Boolean).join(", ")).ponto();
      m.add(limpa(d.tituloAnais) || "Anais", { b: true }).add(" [...]").ponto();
      m.add(imprenta(d.local, d.editora, d.ano)).ponto();
      m.add(paginas(d.paginas)).ponto();
      online(m, d);
      return aut;
    },

    site(m, d) {
      const aut = entrada(m, d, true);
      m.add(`${limpa(d.local) || "[S. l.]"}, ${limpa(d.ano) || "[s. d.]"}`).ponto();
      online(m, d);
      return aut;
    },

    legislacao(m, d) {
      m.add(maiusc(limpa(d.jurisdicao))).ponto();
      m.add(limpa(d.epigrafe)).ponto();
      m.add(limpa(d.ementa)).ponto();
      const det = limpa(d.detalhes);
      m.add(limpa(d.publicacao), { b: true });
      if (det) m.add((/^se[çc][ãa]o/i.test(det) ? ": " : ", ") + det);
      m.ponto();
      online(m, d);
      return { citacao: [nomeCitacao(d.jurisdicao)] };
    },
  };

  // ------------------------------------------------------------ citação
  function anoCitacao(tipo, d) {
    if (tipo === "legislacao") {
      const anos = limpa(d.epigrafe).match(/\b(1[5-9]|20)\d{2}\b/g);
      return anos ? anos[anos.length - 1] : "[s. d.]";
    }
    return limpa(d.ano) || "[s. d.]";
  }

  function citacao(tipo, d, aut) {
    const ano = anoCitacao(tipo, d);
    let par, nar;
    if (aut && aut.citacao.length) {
      const n = aut.citacao;
      if (n.length > 3) {
        par = [{ t: n[0] + " " }, { t: "et al.", i: true }, { t: `, ${ano}` }];
        nar = [{ t: n[0] + " " }, { t: "et al.", i: true }, { t: ` (${ano})` }];
      } else {
        par = [{ t: `${n.join("; ")}, ${ano}` }];
        const e = n.length === 1 ? n[0] : n.slice(0, -1).join(", ") + " e " + n[n.length - 1];
        nar = [{ t: `${e} (${ano})` }];
      }
    } else {
      const t = limpa(tipo === "capitulo" ? d.titulo : d.titulo).split(" ");
      const ini = t.slice(0, palavrasIniciais(d.titulo)).join(" ");
      par = [{ t: `${ini} [...], ${ano}` }];
      nar = [{ t: `${ini} [...] (${ano})` }];
    }
    par = [{ t: "(" }, ...par, { t: ")" }];
    return { parenteses: saida(par), narrativa: saida(nar) };
  }

  // ------------------------------------------------------------ saída
  const escapa = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  function saida(segs) {
    return {
      segmentos: segs,
      texto: segs.map((s) => s.t).join(""),
      html: segs.map((s) => {
        let h = escapa(s.t);
        if (s.i) h = `<i>${h}</i>`;
        if (s.b) h = `<b>${h}</b>`;
        return h;
      }).join(""),
    };
  }

  // ------------------------------------------------------------ campos/tipos
  // Descrição dos campos, usada pela interface para montar o formulário.
  const CAMPOS = {
    autores: { rotulo: "Autor(es)", tipo: "textarea", dica: "Um por linha, como aparece na obra. Ex.: Maria da Silva", exemplo: "Antônio Carlos Gil" },
    tipoAutor: { rotulo: "O autor é", tipo: "select", opcoes: [["pessoa", "Pessoa(s)"], ["entidade", "Instituição / órgão / empresa"]] },
    responsabilidade: { rotulo: "Responsabilidade", tipo: "select", opcoes: [["", "Autor(es)"], ["org.", "Organizador (org.)"], ["coord.", "Coordenador (coord.)"], ["ed.", "Editor (ed.)"]] },
    etAl: { rotulo: "Com 4 ou mais autores, indicar só o primeiro + et al.", tipo: "checkbox" },
    titulo: { rotulo: "Título", tipo: "text", exemplo: "Como elaborar projetos de pesquisa" },
    subtitulo: { rotulo: "Subtítulo", tipo: "text" },
    edicao: { rotulo: "Edição", tipo: "text", dica: "Só o número. Deixe vazio se for a 1ª.", exemplo: "6" },
    local: { rotulo: "Local (cidade)", tipo: "text", exemplo: "São Paulo" },
    editora: { rotulo: "Editora", tipo: "text", exemplo: "Atlas" },
    ano: { rotulo: "Ano", tipo: "text", exemplo: "2017" },
    paginasTotal: { rotulo: "Número de páginas", tipo: "text", dica: "Opcional" },
    paginas: { rotulo: "Páginas", tipo: "text", exemplo: "45-60" },
    autoresLivro: { rotulo: "Autor(es) / organizador(es) do livro", tipo: "textarea", dica: "Um por linha" },
    tipoAutorLivro: { rotulo: "O autor do livro é", tipo: "select", opcoes: [["pessoa", "Pessoa(s)"], ["entidade", "Instituição"]] },
    responsabilidadeLivro: { rotulo: "Responsabilidade no livro", tipo: "select", opcoes: [["org.", "Organizador (org.)"], ["", "Autor(es)"], ["coord.", "Coordenador (coord.)"], ["ed.", "Editor (ed.)"]] },
    tituloLivro: { rotulo: "Título do livro", tipo: "text" },
    subtituloLivro: { rotulo: "Subtítulo do livro", tipo: "text" },
    periodico: { rotulo: "Nome da revista / periódico", tipo: "text", exemplo: "Revista Brasileira de Educação" },
    localPeriodico: { rotulo: "Local de publicação da revista", tipo: "text", dica: "Opcional" },
    volume: { rotulo: "Volume", tipo: "text" },
    numero: { rotulo: "Número / fascículo", tipo: "text" },
    mes: { rotulo: "Mês(es)", tipo: "text", dica: "Ex.: jul./set. ou julho/setembro" },
    doi: { rotulo: "DOI", tipo: "text", dica: "Opcional. Ex.: 10.1590/xxxx" },
    eLocalizador: { rotulo: "Localizador eletrônico", tipo: "text", dica: "Artigos só on-line, sem páginas. Ex.: e03304" },
    url: { rotulo: "Endereço (URL)", tipo: "text", dica: "Para documentos on-line" },
    acesso: { rotulo: "Data de acesso", tipo: "date" },
    folhas: { rotulo: "Número de folhas", tipo: "text", dica: "Opcional" },
    tipoTrabalho: { rotulo: "Tipo", tipo: "select", opcoes: [["Trabalho de Conclusão de Curso", "TCC"], ["Monografia", "Monografia (especialização)"], ["Dissertação", "Dissertação"], ["Tese", "Tese"]] },
    grau: { rotulo: "Grau e curso", tipo: "text", dica: "Ex.: Graduação em Engenharia Civil; Mestrado em Educação" },
    instituicao: { rotulo: "Instituição", tipo: "text", dica: "Ex.: Faculdade de Educação, Universidade Federal do Paraná" },
    anoDefesa: { rotulo: "Ano da defesa", tipo: "text", dica: "Se diferente do ano de depósito" },
    nomeEvento: { rotulo: "Nome do evento", tipo: "text", exemplo: "Congresso Brasileiro de Educação" },
    numeroEvento: { rotulo: "Número do evento", tipo: "text", dica: "Ex.: 10 (para 10º)" },
    anoEvento: { rotulo: "Ano do evento", tipo: "text" },
    localEvento: { rotulo: "Cidade do evento", tipo: "text", dica: "Se houver (NBR 6023:2025)" },
    tituloAnais: { rotulo: "Título da publicação", tipo: "text", dica: "Normalmente “Anais”, “Resumos” ou “Proceedings”" },
    jurisdicao: { rotulo: "Jurisdição", tipo: "text", exemplo: "Brasil", dica: "Brasil, nome do estado ou do município" },
    epigrafe: { rotulo: "Epígrafe (tipo, número e data)", tipo: "text", exemplo: "Lei nº 9.394, de 20 de dezembro de 1996" },
    ementa: { rotulo: "Ementa", tipo: "textarea", dica: "Resumo oficial do conteúdo da lei" },
    publicacao: { rotulo: "Publicado em", tipo: "text", exemplo: "Diário Oficial da União" },
    detalhes: { rotulo: "Dados da publicação", tipo: "text", dica: "Ex.: seção 1, Brasília, DF, ano 134, n. 248, p. 27833-27841, 23 dez. 1996" },
  };

  const AUTOR = ["autores", "tipoAutor", "responsabilidade", "etAl"];
  const ONLINE = ["url", "acesso"];
  const TIPOS = {
    livro: { nome: "Livro", campos: [...AUTOR, "titulo", "subtitulo", "edicao", "local", "editora", "ano", "paginasTotal", ...ONLINE], obrigatorios: ["titulo", "ano"] },
    capitulo: { nome: "Capítulo de livro", campos: [...AUTOR, "titulo", "subtitulo", "autoresLivro", "tipoAutorLivro", "responsabilidadeLivro", "tituloLivro", "subtituloLivro", "edicao", "local", "editora", "ano", "paginas", ...ONLINE], obrigatorios: ["titulo", "tituloLivro", "ano", "paginas"] },
    artigo: { nome: "Artigo de revista", campos: [...AUTOR, "titulo", "subtitulo", "periodico", "localPeriodico", "volume", "numero", "paginas", "eLocalizador", "mes", "ano", "doi", ...ONLINE], obrigatorios: ["titulo", "periodico", "ano"] },
    trabalho: { nome: "TCC, dissertação ou tese", campos: [...AUTOR, "titulo", "subtitulo", "ano", "folhas", "tipoTrabalho", "grau", "instituicao", "local", "anoDefesa", ...ONLINE], obrigatorios: ["autores", "titulo", "ano", "grau", "instituicao", "local"] },
    evento: { nome: "Trabalho em evento (anais)", campos: [...AUTOR, "titulo", "subtitulo", "nomeEvento", "numeroEvento", "anoEvento", "localEvento", "tituloAnais", "local", "editora", "ano", "paginas", "doi", ...ONLINE], obrigatorios: ["titulo", "nomeEvento", "anoEvento", "ano"] },
    site: { nome: "Site / documento on-line", campos: [...AUTOR, "titulo", "subtitulo", "local", "ano", ...ONLINE], obrigatorios: ["titulo", "url"] },
    legislacao: { nome: "Legislação", campos: ["jurisdicao", "epigrafe", "ementa", "publicacao", "detalhes", ...ONLINE], obrigatorios: ["jurisdicao", "epigrafe", "publicacao"] },
  };

  // ------------------------------------------------------------ API
  function formatar(tipo, dados) {
    if (!TIPOS[tipo]) throw new Error("Tipo de referência desconhecido: " + tipo);
    const d = dados || {};
    const m = new Montador();
    const aut = montar[tipo](m, d);
    const ref = saida(m.fim());
    return {
      ...ref,
      citacao: citacao(tipo, d, aut),
      faltando: TIPOS[tipo].obrigatorios.filter((c) => !limpa(d[c])).map((c) => CAMPOS[c].rotulo),
    };
  }

  // Ordem alfabética da lista de referências (NBR 6023, 9.1).
  function ordenar(lista, chave) {
    const k = chave || ((x) => x.texto);
    return [...lista].sort((a, b) => k(a).localeCompare(k(b), "pt-BR", { sensitivity: "base" }));
  }

  return { formatar, ordenar, TIPOS, CAMPOS, dataABNT, separaNome, hojeISO };
});
