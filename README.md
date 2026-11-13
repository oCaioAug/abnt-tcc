# Modelo ABNT para TCC

Modelos de monografia (TCC) e de artigo científico já formatados segundo as normas da ABNT, para Word e LibreOffice.

**Site:** https://ocaioaug.github.io/abnt-tcc/ reúne o verificador, o gerador de referências, os modelos para baixar e o passo a passo para instalar o painel no Word. É o link para mandar a quem quiser usar.

## Arquivos prontos (`dist/`)

| Arquivo | Para quê |
|---|---|
| `Modelo-TCC-Monografia-ABNT.dotx` | Monografia (NBR 14724). Dar dois cliques cria um documento novo com tudo configurado. Também abre no LibreOffice. |
| `Modelo-Artigo-ABNT.dotx` | Artigo científico (NBR 6022:2018): título, autoria com nota de rodapé, resumo e abstract (100 a 250 palavras), datas de submissão e aprovação, texto corrido, referências, apêndice, anexo e agradecimentos. |
| `*.docx` | O mesmo conteúdo como documento comum, para editar direto. |
| `*.pdf` | Prévia de como fica impresso. |

No artigo, as seções não começam em página nova e não há capa nem sumário. Revistas e cursos costumam ter regras próprias de tamanho e formatação: confira as do seu orientador ou do periódico.

## Como usar

1. **Escreva usando os estilos** (Página Inicial > Estilos), sem formatar à mão:
   - `Título 1` a `Título 5`: seções numeradas (1, 1.1, 1.1.1...). A numeração se ajusta sozinha.
   - `Título pré-textual`: RESUMO, SUMÁRIO etc. (centralizado, fora do sumário).
   - `Título pós-textual`: REFERÊNCIAS, APÊNDICE, ANEXO (centralizado, entra no sumário).
   - `Citação longa`: citação com mais de 3 linhas (recuo de 4 cm, fonte 10, espaçamento simples).
   - `Referência`: cada item da lista de referências.
   - `Legenda` (Referências > Inserir Legenda) e `Fonte da ilustração`: título e fonte de figuras e tabelas.
2. **Apague as caixas amarelas** (estilo `Instrução (apagar)`): são só orientações.
3. **Atualize o sumário e as listas** antes de entregar: clique com o botão direito sobre eles > *Atualizar campo* > *Atualizar o índice inteiro* (ou `Ctrl+A` e depois `F9`).

## Gerador de referências (`referencias/index.html`)

Abra o arquivo `referencias/index.html` no navegador (dois cliques). Não precisa de internet nem de instalar nada.

1. Escolha o tipo: livro, capítulo, artigo, TCC/dissertação/tese, trabalho em evento, site ou legislação.
2. Preencha os campos. A referência (NBR 6023:2018) aparece na hora, junto com a forma de citá-la no texto (NBR 10520:2023), por exemplo `(Gil, 2017)` e `Gil (2017)`.
3. Clique em **Adicionar à lista**. A lista fica em ordem alfabética e salva neste navegador.
4. Clique em **Copiar lista completa** e cole sob o título REFERÊNCIAS. O negrito e o itálico vão junto, tanto no Word quanto no LibreOffice.

O gerador já trata sozinho:
- a inversão dos nomes (Maria da Silva → SILVA, Maria da) e sobrenomes como Júnior, Filho e Neto;
- autores que são instituições;
- 4 ou mais autores (todos ou só o primeiro + et al.);
- obras sem autor (a primeira palavra do título vai em maiúsculas);
- [S. l.] e [s. n.] quando faltam local ou editora;
- abreviação dos meses e data de acesso.

As regras ficam em `referencias/abnt-ref.js`, separadas da página, para poderem ser reaproveitadas num suplemento do Word ou numa extensão do LibreOffice. Para rodar os testes: `npm test`.

## Verificador de formatação (`verificador/index.html`)

Abra `verificador/index.html` no navegador e arraste o seu TCC em `.docx`. No LibreOffice, salve antes como “Word 2007-365 (.docx)”. O arquivo é lido no seu computador e nada é enviado para a internet.

O relatório separa **erros**, **pontos de atenção** e **o que está dentro da norma**, e mostra o trecho do texto onde está cada problema. Ele confere:

| Item | Base |
|---|---|
| Elementos obrigatórios (monografia: resumo, abstract, sumário, introdução, referências; artigo: resumo, introdução, referências) | NBR 14724:2011, 4 / NBR 6022:2018 |
| Papel A4 e margens (3/3/2/2 cm) | NBR 14724:2011, 5.1 |
| Número da página no cabeçalho, à direita | NBR 14724:2011, 5.3 |
| Texto e títulos em tamanho 12 | NBR 14724:2011, 5.1 |
| Espaçamento 1,5 no texto | NBR 14724:2011 |
| Uma só fonte no trabalho | Recomendação (a ABNT não fixa a família) |
| Texto justificado | Prática recomendada (não exigida pela ABNT) |
| Citações longas: recuo de 4 cm, fonte menor, espaçamento simples, sem aspas | NBR 10520:2023 e NBR 14724:2011 |
| Sobrenome nas citações só com a inicial maiúscula, ex.: (Silva, 2020) | NBR 10520:2023 |
| Seções primárias e títulos sem número em página nova (só na monografia) | NBR 14724:2011 |
| Seções numeradas, com hierarquia correta e título à esquerda; apêndices como “APÊNDICE A – Título” | NBR 6024:2012 |
| Títulos sem número (resumo, sumário, referências…) centralizados | NBR 14724:2011 |
| Ilustrações e tabelas com título acima e “Fonte:” abaixo | NBR 14724:2011, 5.8 e 5.9 (tabelas: IBGE) |
| Notas de rodapé em fonte menor e espaçamento simples | NBR 14724:2011 |
| Resumo em parágrafo único (150 a 500 palavras; artigo: 100 a 250), com palavras-chave | NBR 6028:2021 |
| Referências alinhadas à esquerda, espaçamento simples, separadas e em ordem alfabética | NBR 6023:2018 |
| Orientações e textos “[entre colchetes]” do modelo ainda não apagados | Do modelo |

Em cada item do relatório, **Por que?** mostra a norma e o que ela pede, resumido com palavras nossas (o texto das normas é protegido e vendido pela ABNT). O número do item da norma só aparece quando é certo; nos outros casos, vai só o nome da norma. As normas e as explicações ficam em `NORMAS` e `REGRAS`, em `verificador/verificador.js`.

O verificador detecta sozinho se o trabalho é **monografia** ou **artigo** (o artigo não tem sumário e chega logo à introdução) e mostra qual tipo usou; dá para escolher o tipo à mão no topo da página. No artigo, a regra de página nova não se aplica, o abstract e o sumário não são exigidos e o resumo vai de 100 a 250 palavras.

O verificador não enxerga a paginação final (isso depende do Word ou do LibreOffice). Por isso, conferências como “o número só aparece a partir da Introdução” continuam sendo manuais. As regras ficam em `verificador/verificador.js`.

## Suplemento do Word (`word-addin/`)

Um painel dentro do Word (Página Inicial > **Painel ABNT**) com duas abas:

- **Verificar**: confere o documento aberto, inclusive o que ainda não foi salvo, com as mesmas regras do verificador. Clicar num trecho leva até ele no texto.
- **Corrigir**: depois de verificar, o botão **Corrigir a formatação automaticamente** ajusta espaçamento, tamanho da fonte, alinhamento, recuo das citações longas, títulos sem número centralizados, quebra de página antes dos títulos e a formatação das referências. Trocas de texto (sobrenome em maiúsculas nas citações, travessão nos apêndices, palavras-chave) têm um botão em cada item, para você conferir uma a uma. Depois de corrigir, o painel verifica de novo. Tudo pode ser desfeito com `Ctrl+Z`.
  - Referências fora da ordem alfabética: o botão **Reordenar as N referências** reescreve a lista na ordem certa, mantendo negrito, itálico, fonte e tamanho de cada trecho (links clicáveis voltam como texto comum).
  - Margens e número de página no canto superior direito também entram na correção automática (requer Microsoft 365 ou Word 2021 ou mais recente; em versões antigas o painel mostra o caminho manual).
  - Elementos obrigatórios que faltam (resumo, abstract, sumário, introdução, referências) têm o botão **Criar**, que insere o título e um texto de exemplo no lugar certo; o sumário é gerado a partir dos títulos.
  - Continua manual: notas de rodapé, o conteúdo do resumo e o número só a partir da Introdução (o painel numera todas as páginas).
- **Referências**: o mesmo gerador de referências. Os botões inserem a referência num parágrafo novo, abaixo do cursor (alinhada à esquerda, espaçamento simples), e a citação, como (Gil, 2017) ou Gil (2017), no ponto do cursor.

Requer Word 2019 ou mais recente, ou Microsoft 365.

### Instalar (versão publicada, sem servidor)

O painel fica publicado em `https://ocaioaug.github.io/abnt-tcc/word-addin/painel.html`. Para instalar no Word do Windows (requer Node.js):

```bash
npx --yes office-addin-dev-settings register word-addin/manifest-publicado.xml
```

Feche e abra o Word: o botão **Painel ABNT** aparece na guia Página Inicial. Não precisa de terminal aberto nem de certificado. Para remover: `npx --yes office-addin-dev-settings unregister word-addin/manifest-publicado.xml`.

No Word para a web: Página Inicial > Suplementos > Mais suplementos > Meus suplementos > Carregar meu suplemento, e escolha o arquivo `word-addin/manifest-publicado.xml`.

### Rodar localmente (para quem for alterar o painel)

Requer Node.js:

```bash
npm run addin:certificado
npm run addin:registrar
npm run addin
```

1. `addin:certificado` (só na primeira vez) cria um certificado para `https://localhost`. O Windows pergunta se você confia nele: responda **Sim**. O Office só abre suplementos por https.
2. `addin:registrar` (só na primeira vez) cadastra o suplemento no Word deste computador.
3. `addin` liga o servidor local. Deixe a janela aberta enquanto usa o painel.
4. Abra o Word: o botão **Painel ABNT** aparece na guia Página Inicial. Se não aparecer, procure em Página Inicial > Suplementos.

Para tirar o suplemento do Word: `npm run addin:remover`.

Para publicar em outro endereço https, gere o manifest com `node word-addin/gerar-manifest.js https://seu-endereco/`. Para testar mudanças no painel sem o Word, abra `word-addin/simulador.html?docx=caminho/do/arquivo.docx` por um servidor local: ele imita o Word e registra no console (`__log`) cada alteração pedida.

## O que já está configurado

- Folha A4; margens de 3 cm (superior e esquerda) e 2 cm (inferior e direita).
- Arial 12, espaçamento 1,5, texto justificado e recuo de 1,25 cm na primeira linha.
- Paginação: a contagem começa na folha de rosto, mas o número só aparece a partir da Introdução, no canto superior direito. A capa não é contada.
- Estrutura completa da NBR 14724: capa, folha de rosto, ficha catalográfica, folha de aprovação, dedicatória, agradecimentos, epígrafe, resumo, abstract, listas, sumário, capítulos, referências, apêndice e anexo.
- Exemplos de citação (NBR 10520:2023) e de referência (NBR 6023:2018).

## Normas usadas

| Norma | Assunto |
|---|---|
| ABNT NBR 14724:2011 | Trabalhos acadêmicos — Apresentação |
| ABNT NBR 6022:2018 | Artigo em publicação periódica técnica e/ou científica — Apresentação |
| ABNT NBR 6023:2018 | Referências — Elaboração |
| ABNT NBR 6024:2012 | Numeração progressiva das seções de um documento — Apresentação |
| ABNT NBR 6027:2012 | Sumário — Apresentação |
| ABNT NBR 6028:2021 | Resumo, resenha e recensão — Apresentação |
| ABNT NBR 10520:2023 | Citações em documentos — Apresentação |

Tabelas seguem também as Normas de apresentação tabular do IBGE. O texto completo das normas é vendido pela ABNT ([catálogo da ABNT](https://www.abntcatalogo.com.br/)); muitas faculdades dão acesso gratuito pela biblioteca (ABNT Coleção ou Target GEDWeb).

## Gerar o modelo de novo (para quem for alterar)

Requer Node.js e Microsoft Word (o Word é usado para atualizar o sumário e exportar os arquivos).

```bash
npm install
node build.js
powershell -File scripts/finalizar-word.ps1
node scripts/limpar-metadados.js
```

O último passo tira dos arquivos o nome de quem os salvou, que o Word grava automaticamente.

- `src/abnt.js`: regras de formatação (margens, fonte, estilos), compartilhadas por todos os modelos.
- `src/monografia.js`: conteúdo e estrutura da monografia.
- `src/artigo.js`: conteúdo e estrutura do artigo (reaproveita partes da monografia; usa `estilos({ novaPagina: false })` para o texto corrido).
- `build.js` também remove o "recomeçar em 1" que a biblioteca docx grava nas listas: o Word às vezes aplicava esse recomeço no meio do texto e numerava uma seção como 1.
- Para trocar a fonte para Times New Roman, altere `FONTE` em `src/abnt.js`.
