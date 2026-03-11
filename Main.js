/**
 * Ponto de entrada do Web App e URL.
 */

/**
 * Entrada do Web App (GAS).
 * @param {Object} e - Objeto com e.parameter (ex.: view=consulta ou view=cadastro).
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function doGet(e) {
  const view = (e && e.parameter && e.parameter.view) ? String(e.parameter.view) : "consulta";
  const file = view === "cadastro" ? "Index" : "Consulta";
  const template = HtmlService.createTemplateFromFile(file);
  if (file === "Index" && e && e.parameter && e.parameter.id) {
    template.id = String(e.parameter.id);
  } else {
    template.id = "";
  }
  return template
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle("Sistema de Gestão de Cursos")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
 * Retorna a URL do Web App com query view=...
 * @param {string} visualizacao - "consulta" ou "cadastro".
 * @returns {string}
 */
function obterUrlWebApp(visualizacao) {
  const base = ScriptApp.getService().getUrl();
  if (!visualizacao) return base;
  return base + "?view=" + encodeURIComponent(String(visualizacao));
}
