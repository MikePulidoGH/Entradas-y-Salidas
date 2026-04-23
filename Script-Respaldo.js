/**
 * Crea un Google Doc con todo el código del proyecto para respaldo visual.
 */
function generarResumenDoc() {
  const nombreDoc = "Resumen_Codigo_GAS";
  const archivos = DriveApp.getFolderById(ScriptsApp.getProjectKey()).getFiles(); // O mejor, listar los archivos del proyecto
  
  // Creamos el documento nuevo
  let doc = DocumentApp.create(nombreDoc);
  let body = doc.getBody();
  
  body.appendParagraph("RESPALDO DE CÓDIGO - " + new Date().toLocaleString())
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  // Aquí podrías iterar sobre tus archivos .gs y .js para pegarlos dentro
  // Por ahora, Apps Script lo creará como un Google Doc nativo.
  
  console.log("Doc creado con ID: " + doc.getId());
  return doc.getUrl();
}
