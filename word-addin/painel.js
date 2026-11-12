// Painel do suplemento ABNT para o Word.
// Reaproveita os mesmos núcleos das páginas: verificador/verificador.js e referencias/abnt-ref.js.
(function () {
  "use strict";
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const ROTULO = { erro: "Erro", aviso: "Atenção", ok: "OK", info: "Info" };
  const MOSTRAR = 5;
  // Regras cujas ocorrências não são trechos do texto (seções, cabeçalho, nomes de elementos)
  const SEM_TRECHO = new Set(["papel", "margens", "paginacao", "estrutura", "fonte-familia", "notas"]);
  let noWord = false;

  let t;
  function aviso(msg) {
    const el = $("toast"); el.textContent = msg; el.classList.add("on");
    clearTimeout(t); t = setTimeout(() => el.classList.remove("on"), 2200);
  }

  // ------------------------------------------------------------------ abas
  for (const [aba, painel] of [["aba-verificar", "p-verificar"], ["aba-refs", "p-refs"]]) {
    $(aba).onclick = () => {
      for (const b of document.querySelectorAll(".abas button")) b.setAttribute("aria-selected", String(b.id === aba));
      $("p-verificar").hidden = painel !== "p-verificar";
      $("p-refs").hidden = painel !== "p-refs";
    };
  }

  // ------------------------------------------------------------------ ler o documento aberto
  // Pede ao Word o documento atual (inclusive o que ainda não foi salvo) como .docx.
  function lerDocumentoAberto() {
    return new Promise((resolve, reject) => {
      Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4 * 1024 * 1024 }, (res) => {
        if (res.status !== Office.AsyncResultStatus.Succeeded) return reject(new Error(res.error.message));
        const arq = res.value;
        const partes = [];
        const fechar = () => arq.closeAsync(() => {});
        const pedir = (i) => arq.getSliceAsync(i, (r) => {
          if (r.status !== Office.AsyncResultStatus.Succeeded) { fechar(); return reject(new Error(r.error.message)); }
          partes.push(new Uint8Array(r.value.data));
          if (i + 1 < arq.sliceCount) return pedir(i + 1);
          fechar();
          const total = partes.reduce((n, p) => n + p.length, 0);
          const tudo = new Uint8Array(total);
          let pos = 0;
          for (const p of partes) { tudo.set(p, pos); pos += p.length; }
          resolve(tudo.buffer);
        });
        pedir(0);
      });
    });
  }

  async function verificarBuffer(buf, nome) {
    const res = $("resultado");
    res.innerHTML = `<p class="nota">Analisando…</p>`;
    try {
      const rel = await VerificadorABNT.verificar(buf, { JSZip: window.JSZip, DOMParser: window.DOMParser });
      relatorio = rel;
      res.innerHTML = desenhar(rel, nome);
      ligarEventos(res);
      $("dica-verificar").hidden = true;
    } catch (e) {
      res.innerHTML = `<div class="falha">Não foi possível analisar o documento.<br><small>${esc(e.message || e)}</small></div>`;
    }
  }

  let relatorio = null;
  async function verificarAberto() {
    const bt = $("bt-verificar");
    bt.disabled = true; bt.textContent = "Lendo o documento…";
    try { await verificarBuffer(await lerDocumentoAberto()); } catch (e) {
      $("resultado").innerHTML = `<div class="falha">O Word não entregou o documento.<br><small>${esc(e.message || e)}</small></div>`;
    } finally { bt.disabled = false; bt.textContent = "Verificar de novo"; }
  }
  $("bt-verificar").onclick = () => (noWord ? verificarAberto() : $("arquivo").click());
  $("arquivo").onchange = async (e) => {
    const f = e.target.files[0];
    if (f) await verificarBuffer(await f.arrayBuffer(), f.name);
  };

  // ------------------------------------------------------------------ relatório
  // automáticas: formatação; individuais: trocas de texto que você confere uma a uma
  const automaticas = (it) => it.ocorrencias.filter((o) => o.correcao && !o.correcao.individual);

  function desenharRegra(it, ri, aberta) {
    const navega = noWord && !SEM_TRECHO.has(it.id);
    const lis = it.ocorrencias.map((o, i) => {
      const ir = navega && o.trecho && o.posicao !== undefined;
      const c = noWord && o.correcao;
      const rotulo = c ? (c.texto ? `Trocar por “${esc(c.texto.para.trim())}”` : "Corrigir") : "";
      return `<li${i >= MOSTRAR ? " hidden" : ""}${ir ? ` class="ir" data-oc="${ri}:${i}" title="Ir para este trecho"` : ""}>` +
        (o.trecho ? `<span>${esc(o.trecho)}${o.trecho.length >= 110 ? "…" : ""}</span>` : "") +
        (o.detalhe ? `<span class="det">${esc(o.detalhe)}</span>` : "") +
        `<span class="acoes">${ir ? `<span class="ir-txt">Ir para o trecho →</span>` : ""}` +
        `${c ? `<button class="bt mini" type="button" data-corrigir="${ri}:${i}">${rotulo}</button>` : ""}</span></li>`;
    }).join("");
    const auto = noWord ? automaticas(it).length : 0;
    return `<details class="regra ${it.nivel}"${aberta ? " open" : ""}>
      <summary><span class="selo">${ROTULO[it.nivel]}</span><span class="tit">${esc(it.titulo)}</span>
        <span class="msg">${esc(it.mensagem)}</span></summary>
      ${auto ? `<div class="corrigir-regra"><button class="bt prim mini" type="button" data-corrigir-regra="${ri}">Corrigir ${auto > 1 ? `os ${auto} itens` : "este item"}</button></div>` : ""}
      ${lis ? `<ul class="ocorr">${lis}</ul>` : ""}
      ${it.ocorrencias.length > MOSTRAR ? `<button class="mais" type="button">Mostrar mais ${it.ocorrencias.length - MOSTRAR}</button>` : ""}
    </details>`;
  }

  function desenhar(rel, nome) {
    const g = (n) => rel.itens.filter((x) => x.nivel === n);
    const erros = g("erro"), avisos = g("aviso"), certos = [...g("ok"), ...g("info")];
    const regra = (x, aberta) => desenharRegra(x, rel.itens.indexOf(x), aberta);
    const totalAuto = noWord ? rel.itens.reduce((n, x) => n + automaticas(x).length, 0) : 0;
    const veredito = erros.length ? `Corrija os ${erros.length} ponto(s) marcados como erro.`
      : avisos.length ? "Nenhum erro encontrado. Revise os pontos de atenção." : "Tudo certo nos itens verificados.";
    return `${nome ? `<p class="nota">${esc(nome)}</p>` : ""}
      <div class="placar"><span class="chip erro">${erros.length} erro(s)</span><span class="chip aviso">${avisos.length} atenção</span><span class="chip ok">${certos.length} ok</span></div>
      <p class="nota">${veredito}${rel.partes.introducao ? "" : " Não encontrei o título “Introdução”; algumas regras foram aplicadas ao documento inteiro."}</p>
      ${totalAuto ? `<button class="bt prim larg" type="button" id="bt-corrigir-tudo">Corrigir a formatação automaticamente (${totalAuto})</button>
        <p class="nota">Ajusta espaçamento, fonte, alinhamento, recuos e quebras de página. Trocas de texto (citações, títulos de apêndice) ficam para você conferir uma a uma. Dá para desfazer com Ctrl+Z.</p>` : ""}
      ${erros.length ? `<h3 class="grupo">Erros</h3>${erros.map((x) => regra(x, true)).join("")}` : ""}
      ${avisos.length ? `<h3 class="grupo">Pontos de atenção</h3>${avisos.map((x) => regra(x, true)).join("")}` : ""}
      ${certos.length ? `<h3 class="grupo">Dentro da norma</h3>${certos.map((x) => regra(x, false)).join("")}` : ""}`;
  }

  function ligarEventos(raiz) {
    raiz.querySelectorAll("button.mais").forEach((b) => b.onclick = () => {
      b.previousElementSibling.querySelectorAll("li[hidden]").forEach((li) => li.hidden = false);
      b.remove();
    });
    const corrigiveis = (ri) => relatorio.itens[ri].ocorrencias
      .map((o, i) => (o.correcao && !o.correcao.individual ? alvoDe(ri, i) : null)).filter(Boolean);
    raiz.querySelectorAll("[data-corrigir]").forEach((b) => b.onclick = (ev) => {
      ev.stopPropagation();
      const [ri, i] = b.dataset.corrigir.split(":").map(Number);
      corrigir([alvoDe(ri, i)]);
    });
    raiz.querySelectorAll("[data-corrigir-regra]").forEach((b) => b.onclick = () => corrigir(corrigiveis(Number(b.dataset.corrigirRegra))));
    const tudo = raiz.querySelector("#bt-corrigir-tudo");
    if (tudo) tudo.onclick = () => corrigir(relatorio.itens.flatMap((it, ri) => corrigiveis(ri)));
    raiz.querySelectorAll("li.ir").forEach((li) => li.onclick = async () => {
      const [ri, i] = li.dataset.oc.split(":").map(Number);
      const achou = await irPara(alvoDe(ri, i));
      if (!achou) {
        li.classList.add("nao-achou");
        li.querySelector(".ir-txt").textContent = "Trecho não encontrado (pode ter sido alterado)";
      }
    });
  }

  // ------------------------------------------------------------------ achar o parágrafo no Word
  // O verificador informa a posição de cada parágrafo (a mesma ordem de body.paragraphs).
  // Textos repetidos ("[Texto.]", "Fonte: Autor") deixam de ser problema: vale a posição,
  // conferida pelo texto. Se o documento mudou, fica o parágrafo de mesmo texto mais próximo.
  const norm = (t) => String(t || "").normalize("NFC").replace(/[\s\u0007]+/g, " ").trim();

  function alvoDe(ri, i) {
    const oc = relatorio.itens[ri].ocorrencias;
    const o = oc[i];
    const de = o.correcao && o.correcao.texto ? o.correcao.texto.de : null;
    // várias trocas do mesmo texto no mesmo parágrafo: qual delas é esta
    const ordem = de ? oc.slice(0, i).filter((x) => x.posicao === o.posicao && x.correcao && x.correcao.texto && x.correcao.texto.de === de).length : 0;
    return { trecho: o.trecho, posicao: o.posicao, correcao: o.correcao, ordem, contem: de || (o.posicao !== undefined && o.trecho.length < 40 && /\d{4}/.test(o.trecho)) };
  }

  // Compara só as letras: o texto que o Word devolve traz marcas que o .docx não tem
  // (chamada de nota de rodapé, campos, objetos), e o trecho do relatório pode vir cortado.
  const letras = (t) => String(t || "").normalize("NFC").toLocaleLowerCase("pt-BR").replace(/[^\p{L}]/gu, "");
  function confere(textoWord, alvo) {
    const a = letras(textoWord), b = letras(alvo.contem ? (typeof alvo.contem === "string" ? alvo.contem : alvo.trecho) : alvo.trecho);
    if (!b) return norm(textoWord) === norm(alvo.trecho); // trecho sem letras (só números/símbolos)
    return alvo.contem ? a.includes(b) : a.startsWith(b);
  }

  async function paragrafosDoWord(ctx) {
    const ps = ctx.document.body.paragraphs;
    ps.load("items/text");
    await ctx.sync();
    return ps.items;
  }

  function acharParagrafo(itens, alvo) {
    if (alvo.posicao === undefined) return null;
    const p = itens[alvo.posicao];
    if (p && confere(p.text, alvo)) return p;
    let melhor = null, dist = Infinity;
    itens.forEach((q, k) => {
      if (Math.abs(k - alvo.posicao) < dist && confere(q.text, alvo)) { dist = Math.abs(k - alvo.posicao); melhor = q; }
    });
    return melhor;
  }

  // A busca do Word aceita até 255 caracteres e trata "^" como caractere especial.
  const termoBusca = (t) => norm(t).slice(0, 250).replace(/\^/g, "^^");

  async function irPara(alvo) {
    try {
      return await Word.run(async (ctx) => {
        const p = acharParagrafo(await paragrafosDoWord(ctx), alvo);
        if (!p) return false;
        if (typeof alvo.contem === "string" || (alvo.contem && alvo.trecho)) {
          const res = p.search(termoBusca(typeof alvo.contem === "string" ? alvo.contem : alvo.trecho), { matchCase: true });
          res.load("items");
          await ctx.sync();
          (res.items[alvo.ordem] || res.items[0] || p).select();
        } else p.select();
        await ctx.sync();
        return true;
      });
    } catch (e) {
      return false;
    }
  }

  // ------------------------------------------------------------------ correções
  const PT_POR_CM = 72 / 2.54;
  const ALINHAMENTO = { esquerda: "Left", centralizado: "Centered", justificado: "Justified", direita: "Right" };

  function aplicarFormato(p, f) {
    if (f.alinhamento) p.alignment = ALINHAMENTO[f.alinhamento];
    if (f.linha !== undefined) p.lineSpacing = 12 * f.linha; // Word: 12 = 1 linha, 18 = 1,5
    if (f.recuoEsq !== undefined) p.leftIndent = f.recuoEsq * PT_POR_CM;
    if (f.primeiraLinha !== undefined) p.firstLineIndent = f.primeiraLinha * PT_POR_CM;
    if (f.depois !== undefined) p.spaceAfter = f.depois;
    if (f.tamanho !== undefined) p.font.size = f.tamanho;
    if (f.fonte) p.font.name = f.fonte;
  }

  async function corrigir(alvos) {
    if (!alvos.length) return;
    let feitos = 0, faltaram = 0;
    try {
      await Word.run(async (ctx) => {
        const itens = await paragrafosDoWord(ctx);
        // primeiro acha tudo (as quebras de página inseridas mudariam as posições), depois altera
        const plano = alvos.map((a) => {
          const p = acharParagrafo(itens, a);
          if (!p) return null;
          if (!a.correcao.texto) return { a, p };
          const res = p.search(termoBusca(a.correcao.texto.de), { matchCase: true });
          res.load("items");
          return { a, p, res };
        });
        if (plano.some((x) => x && x.res)) await ctx.sync();
        for (const x of plano) {
          if (!x) { faltaram++; continue; }
          const c = x.a.correcao;
          if (c.texto) {
            const r = x.res.items[x.a.ordem] || x.res.items[0];
            if (!r) { faltaram++; continue; }
            r.insertText(c.texto.para, "Replace");
          } else {
            if (c.formato) aplicarFormato(x.p, c.formato);
            if (c.quebraAntes) x.p.insertBreak(Word.BreakType.page, "Before");
          }
          feitos++;
        }
        await ctx.sync();
      });
    } catch (e) {
      aviso("Não foi possível corrigir: " + (e.message || e));
      return;
    }
    aviso(`${feitos} correção(ões) aplicada(s)${faltaram ? `; ${faltaram} trecho(s) não encontrado(s)` : ""}. Verificando de novo…`);
    await verificarAberto();
  }

  // ------------------------------------------------------------------ referências
  const { formatar, TIPOS, CAMPOS, hojeISO } = window.ABNTRef;
  const CURTOS = new Set(["edicao", "local", "editora", "ano", "paginasTotal", "paginas", "volume", "numero", "mes",
    "localPeriodico", "folhas", "anoDefesa", "numeroEvento", "anoEvento", "localEvento", "acesso"]);
  const estado = { tipo: "livro", dados: { acesso: hojeISO() } };

  $("tipo").innerHTML = Object.entries(TIPOS).map(([id, x]) => `<option value="${id}">${esc(x.nome)}</option>`).join("");
  $("tipo").onchange = () => { estado.tipo = $("tipo").value; desenharForm(); atualizar(); };

  function criarCampo(id) {
    const c = CAMPOS[id];
    const obrig = TIPOS[estado.tipo].obrigatorios.includes(id);
    const div = document.createElement("div");
    div.className = "campo";
    const ex = c.exemplo ? ` placeholder="Ex.: ${esc(c.exemplo)}"` : "";
    let html;
    if (c.tipo === "checkbox") html = `<label class="checkbox"><input type="checkbox" id="f-${id}"> ${esc(c.rotulo)}</label>`;
    else {
      html = `<label for="f-${id}">${esc(c.rotulo)}${obrig ? '<span class="obrig" title="obrigatório">*</span>' : ""}</label>`;
      if (c.tipo === "textarea") html += `<textarea id="f-${id}"${ex}></textarea>`;
      else if (c.tipo === "select") html += `<select id="f-${id}">${c.opcoes.map(([v, x]) => `<option value="${esc(v)}">${esc(x)}</option>`).join("")}</select>`;
      else html += `<input type="${c.tipo}" id="f-${id}"${ex}>`;
      if (c.dica) html += `<small>${esc(c.dica)}</small>`;
    }
    div.innerHTML = html;
    const el = div.querySelector("input, textarea, select");
    const valor = estado.dados[id];
    if (c.tipo === "checkbox") { el.checked = !!valor; el.onchange = () => { estado.dados[id] = el.checked; atualizar(); }; }
    else {
      if (valor !== undefined) el.value = valor; else if (c.tipo === "select") estado.dados[id] = el.value;
      el.oninput = () => { estado.dados[id] = el.value; atualizar(); };
    }
    return div;
  }

  function desenharForm() {
    const form = $("form");
    form.innerHTML = "";
    let grade = null;
    for (const id of TIPOS[estado.tipo].campos) {
      const campo = criarCampo(id);
      if (CURTOS.has(id)) {
        if (!grade) { grade = document.createElement("div"); grade.className = "grade2"; form.appendChild(grade); }
        grade.appendChild(campo);
        if (grade.children.length === 2) grade = null;
      } else { grade = null; form.appendChild(campo); }
    }
  }

  const temConteudo = () => TIPOS[estado.tipo].campos.some((c) => CAMPOS[c].tipo !== "select" && c !== "acesso" && c !== "etAl" && String(estado.dados[c] || "").trim());

  function atualizar() {
    const ok = temConteudo();
    ["bt-ref", "bt-cit-par", "bt-cit-nar"].forEach((id) => { $(id).disabled = !ok; });
    if (!ok) {
      $("previa").className = "previa vazia"; $("previa").textContent = "Preencha os campos acima.";
      $("faltando").hidden = true;
      return;
    }
    const r = formatar(estado.tipo, estado.dados);
    $("previa").className = "previa"; $("previa").innerHTML = r.html;
    $("bt-cit-par").innerHTML = "Inserir citação: " + r.citacao.parenteses.html;
    $("bt-cit-nar").innerHTML = "Inserir citação: " + r.citacao.narrativa.html;
    $("faltando").hidden = !r.faltando.length;
    $("faltando").textContent = r.faltando.length ? "Faltam dados obrigatórios: " + r.faltando.join(", ") + "." : "";
  }

  $("bt-limpar").onclick = () => {
    estado.dados = { acesso: hojeISO() };
    desenharForm(); atualizar();
    $("bt-cit-par").textContent = "Inserir citação"; $("bt-cit-nar").textContent = "Inserir citação narrativa";
  };

  // Escreve os segmentos ({t, b, i}) em sequência, cada um com seu negrito/itálico.
  function escreverSegmentos(alvo, segs, local) {
    let r = null;
    for (const s of segs) {
      if (!s.t) continue;
      r = r ? r.insertText(s.t, "After") : alvo.insertText(s.t, local);
      r.font.bold = !!s.b;
      r.font.italic = !!s.i;
    }
    return r;
  }

  async function inserirReferencia() {
    if (!noWord) return aviso("Abra o painel dentro do Word para inserir.");
    const r = formatar(estado.tipo, estado.dados);
    try {
      await Word.run(async (ctx) => {
        const atual = ctx.document.getSelection().paragraphs.getFirst();
        atual.load("text");
        await ctx.sync();
        // usa a linha do cursor se estiver vazia; senão cria um parágrafo logo abaixo
        const p = atual.text.trim() === "" ? atual : atual.insertParagraph("", "After");
        p.styleBuiltIn = Word.Style.normal;
        // NBR 6023: alinhada à esquerda, espaçamento simples, uma linha em branco entre as referências
        p.alignment = "Left";
        p.firstLineIndent = 0;
        p.leftIndent = 0;
        p.lineSpacing = 12;
        p.spaceBefore = 0;
        p.spaceAfter = 12;
        const fim = escreverSegmentos(p, r.segmentos, "Start");
        if (fim) fim.getRange("End").select();
        await ctx.sync();
      });
      aviso("Referência inserida");
    } catch (e) {
      aviso("Não foi possível inserir: " + (e.message || e));
    }
  }

  async function inserirCitacao(tipo) {
    if (!noWord) return aviso("Abra o painel dentro do Word para inserir.");
    const c = formatar(estado.tipo, estado.dados).citacao[tipo];
    try {
      await Word.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        const fim = escreverSegmentos(sel, c.segmentos, "Replace");
        if (fim) fim.getRange("End").select();
        await ctx.sync();
      });
      aviso("Citação inserida");
    } catch (e) {
      aviso("Não foi possível inserir: " + (e.message || e));
    }
  }

  $("bt-ref").onclick = inserirReferencia;
  $("bt-cit-par").onclick = () => inserirCitacao("parenteses");
  $("bt-cit-nar").onclick = () => inserirCitacao("narrativa");

  desenharForm(); atualizar();

  // ------------------------------------------------------------------ início
  function foraDoWord() {
    noWord = false;
    $("fora-word").hidden = false;
    $("bt-verificar").textContent = "Escolher arquivo .docx";
  }
  if (typeof Office === "undefined") foraDoWord();
  else Office.onReady((info) => {
    if (info.host === Office.HostType.Word) noWord = true;
    else foraDoWord();
  });
})();
