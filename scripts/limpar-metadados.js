// Tira o nome de quem salvou o modelo (o Word grava "Modificado por" com o usuário da máquina).
// Rodado depois de scripts/finalizar-word.ps1.
const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");

const NOME = "Modelo ABNT";
const DIST = path.join(__dirname, "..", "dist");

(async () => {
  for (const arq of fs.readdirSync(DIST).filter((f) => /\.(docx|dotx)$/i.test(f))) {
    const caminho = path.join(DIST, arq);
    const zip = await JSZip.loadAsync(fs.readFileSync(caminho));
    for (const parte of ["docProps/core.xml", "docProps/app.xml"]) {
      const f = zip.file(parte);
      if (!f) continue;
      const xml = (await f.async("string"))
        .replace(/<dc:creator>[^<]*<\/dc:creator>/, `<dc:creator>${NOME}</dc:creator>`)
        .replace(/<cp:lastModifiedBy>[^<]*<\/cp:lastModifiedBy>/, `<cp:lastModifiedBy>${NOME}</cp:lastModifiedBy>`)
        .replace(/<Company>[^<]*<\/Company>/, "<Company></Company>")
        .replace(/<Manager>[^<]*<\/Manager>/, "<Manager></Manager>");
      zip.file(parte, xml);
    }
    fs.writeFileSync(caminho, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
    console.log("metadados limpos:", arq);
  }
})();
