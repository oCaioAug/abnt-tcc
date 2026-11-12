// Gera os modelos em build/ (.docx bruto). Depois, scripts/finalizar-word.ps1
// abre no Word, atualiza sumário/listas e salva o .dotx e a prévia em PDF em dist/.
const fs = require("fs");
const path = require("path");
const { Packer } = require("docx");
const JSZip = require("jszip");

const modelos = {
  "Modelo-TCC-Monografia-ABNT": require("./src/monografia"),
  "Modelo-Artigo-ABNT": require("./src/artigo"),        // NBR 6022:2018
};

// A biblioteca docx grava "recomeçar em 1" (startOverride) em cada lista. O Word às vezes
// aplica esse recomeço no meio do texto ao paginar (ex.: "1 CONSIDERAÇÕES FINAIS" depois
// da seção 4). As listas já começam em 1 pela definição, então o recomeço é removido.
async function semRecomecoDeLista(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const f = zip.file("word/numbering.xml");
  if (!f) return buffer;
  zip.file("word/numbering.xml", (await f.async("string")).replace(/<w:lvlOverride\b[\s\S]*?<\/w:lvlOverride>/g, ""));
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

(async () => {
  const saida = path.join(__dirname, "build");
  fs.mkdirSync(saida, { recursive: true });
  for (const [nome, modelo] of Object.entries(modelos)) {
    const arquivo = path.join(saida, `${nome}.docx`);
    fs.writeFileSync(arquivo, await semRecomecoDeLista(await Packer.toBuffer(modelo.construir())));
    console.log("gerado:", arquivo);
  }
})();
