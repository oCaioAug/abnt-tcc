// Gera os modelos em build/ (.docx bruto). Depois, scripts/finalizar-word.ps1
// abre no Word, atualiza sumário/listas e salva o .dotx e a prévia em PDF em dist/.
const fs = require("fs");
const path = require("path");
const { Packer } = require("docx");

const modelos = {
  "Modelo-TCC-Monografia-ABNT": require("./src/monografia"),
  // "Modelo-Artigo-ABNT": require("./src/artigo"),   // NBR 6022 (futuro)
};

(async () => {
  const saida = path.join(__dirname, "build");
  fs.mkdirSync(saida, { recursive: true });
  for (const [nome, modelo] of Object.entries(modelos)) {
    const arquivo = path.join(saida, `${nome}.docx`);
    fs.writeFileSync(arquivo, await Packer.toBuffer(modelo.construir()));
    console.log("gerado:", arquivo);
  }
})();
