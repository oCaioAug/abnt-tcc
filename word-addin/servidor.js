// Servidor https local para o suplemento do Word (o Office só carrega suplementos via https).
// Uso: npm run addin   (antes, uma vez: npm run addin:certificado)
const https = require("https");
const fs = require("fs");
const path = require("path");
const os = require("os");

const PORTA = 3000;
const RAIZ = path.resolve(__dirname, "..");
const PASTA_CERT = path.join(os.homedir(), ".office-addin-dev-certs");
const TIPOS = {
  ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8",
  ".png": "image/png", ".svg": "image/svg+xml", ".xml": "application/xml; charset=utf-8", ".md": "text/plain; charset=utf-8",
  ".json": "application/json; charset=utf-8", ".ico": "image/x-icon",
};
// Só as pastas que o painel usa; o resto do projeto (node_modules, dist...) não é servido.
const PERMITIDAS = ["word-addin", "verificador", "referencias"].map((p) => path.join(RAIZ, p) + path.sep);

let chave, cert;
try {
  chave = fs.readFileSync(path.join(PASTA_CERT, "localhost.key"));
  cert = fs.readFileSync(path.join(PASTA_CERT, "localhost.crt"));
} catch (e) {
  console.error("Certificado de desenvolvimento não encontrado em " + PASTA_CERT + ".\nRode uma vez: npm run addin:certificado");
  process.exit(1);
}

https.createServer({ key: chave, cert }, (req, res) => {
  const url = decodeURIComponent(new URL(req.url, "https://localhost").pathname);
  const arquivo = path.resolve(RAIZ, "." + url);
  const ok = arquivo === path.join(RAIZ, "README.md") || PERMITIDAS.some((p) => arquivo.startsWith(p));
  if (!ok) { res.writeHead(404); return res.end("Não encontrado"); }
  fs.readFile(arquivo, (erro, dados) => {
    if (erro) { res.writeHead(404); return res.end("Não encontrado"); }
    res.writeHead(200, { "Content-Type": TIPOS[path.extname(arquivo).toLowerCase()] || "application/octet-stream", "Cache-Control": "no-cache" });
    res.end(dados);
  });
}).listen(PORTA, "localhost", () => {
  console.log(`Suplemento ABNT disponível em https://localhost:${PORTA}/word-addin/painel.html`);
  console.log("Deixe esta janela aberta enquanto usa o painel no Word. Ctrl+C para parar.");
});
