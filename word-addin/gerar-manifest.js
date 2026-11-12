// Gera uma cópia do manifest apontando para outro endereço https (ex.: GitHub Pages).
// Uso: node word-addin/gerar-manifest.js https://usuario.github.io/abnt-tcc/
const fs = require("fs");
const path = require("path");

const base = (process.argv[2] || "").replace(/\/+$/, "");
if (!/^https:\/\/[^/]+/.test(base)) {
  console.error("Informe o endereço https onde a pasta do projeto está publicada.\nEx.: node word-addin/gerar-manifest.js https://usuario.github.io/abnt-tcc/");
  process.exit(1);
}
const origem = base.match(/^https:\/\/[^/]+/)[0];
const xml = fs.readFileSync(path.join(__dirname, "manifest.xml"), "utf8")
  .replace(/<AppDomain>https:\/\/localhost:3000<\/AppDomain>/, `<AppDomain>${origem}</AppDomain>`)
  .replace(/https:\/\/localhost:3000/g, base);
const saida = path.join(__dirname, "manifest-publicado.xml");
fs.writeFileSync(saida, xml);
console.log("Manifest gerado em", saida);
