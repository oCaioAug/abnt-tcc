// Testes do núcleo: npm test
const test = require("node:test");
const assert = require("node:assert/strict");
const { formatar, ordenar, separaNome } = require("./abnt-ref");

const ACESSO = "2026-11-12";
const ref = (tipo, d) => formatar(tipo, d);

test("livro com um autor: título em negrito e edição", () => {
  const r = ref("livro", { autores: "Antônio Carlos Gil", titulo: "Como elaborar projetos de pesquisa", edicao: "6", local: "São Paulo", editora: "Atlas", ano: "2017" });
  assert.equal(r.texto, "GIL, Antônio Carlos. Como elaborar projetos de pesquisa. 6. ed. São Paulo: Atlas, 2017.");
  assert.equal(r.html, "GIL, Antônio Carlos. <b>Como elaborar projetos de pesquisa</b>. 6. ed. São Paulo: Atlas, 2017.");
  assert.equal(r.citacao.parenteses.texto, "(Gil, 2017)");
  assert.equal(r.citacao.narrativa.texto, "Gil (2017)");
  assert.deepEqual(r.faltando, []);
});

test("dois autores, partícula no prenome, subtítulo fora do negrito", () => {
  const r = ref("livro", { autores: "Marina de Andrade Marconi\nEva Maria Lakatos", titulo: "Metodologia científica", subtitulo: "ciência e conhecimento científico", edicao: "8", local: "São Paulo", editora: "Atlas", ano: "2022" });
  assert.equal(r.texto, "MARCONI, Marina de Andrade; LAKATOS, Eva Maria. Metodologia científica: ciência e conhecimento científico. 8. ed. São Paulo: Atlas, 2022.");
  assert.match(r.html, /<b>Metodologia científica<\/b>: ciência/);
  assert.equal(r.citacao.parenteses.texto, "(Marconi; Lakatos, 2022)");
  assert.equal(r.citacao.narrativa.texto, "Marconi e Lakatos (2022)");
});

test("primeira edição é omitida; local e editora ausentes", () => {
  const r = ref("livro", { autores: "Ana Souza", titulo: "Título", edicao: "1", ano: "2020" });
  assert.equal(r.texto, "SOUZA, Ana. Título. [S. l.: s. n.], 2020.");
  const r2 = ref("livro", { autores: "Ana Souza", titulo: "Título", editora: "Atlas", ano: "2020" });
  assert.equal(r2.texto, "SOUZA, Ana. Título. [S. l.]: Atlas, 2020.");
});

test("sobrenomes com Júnior/Filho/Neto e nome já invertido", () => {
  assert.deepEqual(separaNome("José Assis Júnior"), { sobrenome: "Assis Júnior", prenome: "José" });
  assert.deepEqual(separaNome("Paulo Freire Neto"), { sobrenome: "Freire Neto", prenome: "Paulo" });
  assert.deepEqual(separaNome("Silva, Maria da"), { sobrenome: "Silva", prenome: "Maria da" });
  const r = ref("livro", { autores: "José Assis Júnior", titulo: "T", ano: "2020" });
  assert.match(r.texto, /^ASSIS JÚNIOR, José\./);
  assert.equal(r.citacao.parenteses.texto, "(Assis Júnior, 2020)");
});

test("quatro ou mais autores: todos ou et al.", () => {
  const autores = "Ana Silva\nBruno Souza\nCarla Lima\nDaniel Rocha";
  const todos = ref("livro", { autores, titulo: "T", ano: "2020" });
  assert.match(todos.texto, /^SILVA, Ana; SOUZA, Bruno; LIMA, Carla; ROCHA, Daniel\. /);
  const abrev = ref("livro", { autores, etAl: true, titulo: "T", ano: "2020" });
  assert.match(abrev.texto, /^SILVA, Ana et al\. T\. /);
  assert.equal(todos.citacao.parenteses.texto, "(Silva et al., 2020)");
  assert.equal(todos.citacao.parenteses.html, "(Silva <i>et al.</i>, 2020)");
});

test("organizador", () => {
  const r = ref("livro", { autores: "Ana Silva", responsabilidade: "org.", titulo: "Coletânea", ano: "2019" });
  assert.match(r.texto, /^SILVA, Ana \(org\.\)\. Coletânea\./);
});

test("instituição como autora", () => {
  const r = ref("site", { autores: "Universidade Federal do Paraná", tipoAutor: "entidade", titulo: "Normas para apresentação de documentos", local: "Curitiba", ano: "2022", url: "https://www.ufpr.br", acesso: ACESSO });
  assert.equal(r.texto, "UNIVERSIDADE FEDERAL DO PARANÁ. Normas para apresentação de documentos. Curitiba, 2022. Disponível em: https://www.ufpr.br. Acesso em: 12 nov. 2026.");
  assert.equal(r.citacao.parenteses.texto, "(Universidade Federal do Paraná, 2022)");
});

test("sem autor: entrada pelo título, primeira palavra (com artigo) em maiúsculas, sem negrito", () => {
  const r = ref("livro", { titulo: "O ensino de física no Brasil", local: "Brasília", editora: "MEC", ano: "2015" });
  assert.equal(r.texto, "O ENSINO de física no Brasil. Brasília: MEC, 2015.");
  assert.doesNotMatch(r.html, /<b>/);
  assert.equal(r.citacao.parenteses.texto, "(O ensino [...], 2015)");
});

test("capítulo de livro com organizador", () => {
  const r = ref("capitulo", { autores: "Carlos Lima", titulo: "Pesquisa qualitativa", autoresLivro: "Ana Silva", responsabilidadeLivro: "org.", tituloLivro: "Métodos de pesquisa", local: "Porto Alegre", editora: "Penso", ano: "2019", paginas: "10 - 25" });
  assert.equal(r.texto, "LIMA, Carlos. Pesquisa qualitativa. In: SILVA, Ana (org.). Métodos de pesquisa. Porto Alegre: Penso, 2019. p. 10-25.");
  assert.match(r.html, /<i>In<\/i>: SILVA, Ana \(org\.\)\. <b>Métodos de pesquisa<\/b>/);
});

test("artigo de revista com DOI e meses por extenso", () => {
  const r = ref("artigo", { autores: "Maria Souza", titulo: "Título do artigo", periodico: "Revista Brasileira de Educação", localPeriodico: "Rio de Janeiro", volume: "25", numero: "3", paginas: "45-60", mes: "julho/setembro", ano: "2021", doi: "https://doi.org/10.1590/abc" });
  assert.equal(r.texto, "SOUZA, Maria. Título do artigo. Revista Brasileira de Educação, Rio de Janeiro, v. 25, n. 3, p. 45-60, jul./set. 2021. DOI: https://doi.org/10.1590/abc.");
  assert.match(r.html, /<b>Revista Brasileira de Educação<\/b>, Rio/);
});

test("artigo on-line", () => {
  const r = ref("artigo", { autores: "Maria Souza", titulo: "Artigo", periodico: "Revista X", volume: "2", ano: "2020", url: "https://revista.br/a", acesso: ACESSO });
  assert.equal(r.texto, "SOUZA, Maria. Artigo. Revista X, v. 2, 2020. Disponível em: https://revista.br/a. Acesso em: 12 nov. 2026.");
});

test("dissertação", () => {
  const r = ref("trabalho", { autores: "Ana Silva", titulo: "Educação inclusiva", ano: "2022", folhas: "120", tipoTrabalho: "Dissertação", grau: "Mestrado em Educação", instituicao: "Universidade Federal do Paraná", local: "Curitiba" });
  assert.equal(r.texto, "SILVA, Ana. Educação inclusiva. 2022. 120 f. Dissertação (Mestrado em Educação) – Universidade Federal do Paraná, Curitiba, 2022.");
});

test("trabalho em evento", () => {
  const r = ref("evento", { autores: "Ana Silva", titulo: "Ensino remoto", nomeEvento: "Congresso Brasileiro de Educação", numeroEvento: "10", anoEvento: "2020", localEvento: "Recife", local: "Recife", editora: "UFPE", ano: "2020", paginas: "12-20" });
  assert.equal(r.texto, "SILVA, Ana. Ensino remoto. In: CONGRESSO BRASILEIRO DE EDUCAÇÃO, 10., 2020, Recife. Anais [...]. Recife: UFPE, 2020. p. 12-20.");
  assert.match(r.html, /<b>Anais<\/b> \[\.\.\.\]/);
});

test("legislação", () => {
  const r = ref("legislacao", { jurisdicao: "Brasil", epigrafe: "Lei nº 9.394, de 20 de dezembro de 1996", ementa: "Estabelece as diretrizes e bases da educação nacional.", publicacao: "Diário Oficial da União", detalhes: "seção 1, Brasília, DF, ano 134, n. 248, p. 27833-27841, 23 dez. 1996" });
  assert.equal(r.texto, "BRASIL. Lei nº 9.394, de 20 de dezembro de 1996. Estabelece as diretrizes e bases da educação nacional. Diário Oficial da União: seção 1, Brasília, DF, ano 134, n. 248, p. 27833-27841, 23 dez. 1996.");
  assert.equal(r.citacao.parenteses.texto, "(Brasil, 1996)");
});

test("título terminado em interrogação não ganha ponto extra", () => {
  const r = ref("livro", { autores: "Ana Silva", titulo: "O que é ciência?", ano: "2020" });
  assert.equal(r.texto, "SILVA, Ana. O que é ciência? [S. l.: s. n.], 2020.");
});

test("campos obrigatórios faltando são informados", () => {
  const r = ref("artigo", { titulo: "X" });
  assert.deepEqual(r.faltando, ["Nome da revista / periódico", "Ano"]);
});

test("ordenação alfabética ignora acentos e negrito", () => {
  const l = ordenar([{ texto: "ÉVORA, B." }, { texto: "SOUZA, A." }, { texto: "ALVES, C." }]);
  assert.deepEqual(l.map((x) => x.texto), ["ALVES, C.", "ÉVORA, B.", "SOUZA, A."]);
});

test("NBR 6023:2025: artigo só eletrônico usa a e-localização e o DOI como endereço", () => {
  // dados inventados (não copiados de exemplos da norma)
  const r = ref("artigo", { autores: "Joana Prado" + String.fromCharCode(10) + "Rui Andrade", titulo: "Leitura digital no ensino médio", periodico: "Revista de Educação e Tecnologia",
    localPeriodico: "Curitiba", volume: "8", eLocalizador: "e2104", ano: "2024", doi: "10.0000/ret.2024.2104",
    url: "https://revista.exemplo.br/ret/2104", acesso: "2025-03-10" });
  assert.equal(r.texto, "PRADO, Joana; ANDRADE, Rui. Leitura digital no ensino médio. Revista de Educação e Tecnologia, Curitiba, v. 8, e2104, 2024. DOI: https://doi.org/10.0000/ret.2024.2104. Disponível em: https://revista.exemplo.br/ret/2104. Acesso em: 10 mar. 2025.");
});

test("NBR 6023:2025: trabalho em evento sem a cidade do evento", () => {
  const r = ref("evento", { autores: "Ana Silva", titulo: "Trabalho", nomeEvento: "Congresso Brasileiro de Educação", numeroEvento: "10", anoEvento: "2020", local: "São Paulo", editora: "SBE", ano: "2020", paginas: "1-10" });
  assert.equal(r.texto, "SILVA, Ana. Trabalho. In: CONGRESSO BRASILEIRO DE EDUCAÇÃO, 10., 2020. Anais [...]. São Paulo: SBE, 2020. p. 1-10.");
  assert.deepEqual(r.faltando, []);
});
