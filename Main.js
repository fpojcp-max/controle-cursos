/**
 * Ponto de entrada do Web App e URL.
 * Camada Controller (entrada): SPA (uma única página Shell); deep link por view/id na URL.
 */

/**
 * Entrada do Web App (GAS). Sempre serve o Shell (SPA); view/id na URL para deep link.
 * @param {Object} e - Objeto com e.parameter (view=, id=).
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function doGet(e) {
  const view = (e && e.parameter && e.parameter.view) ? String(e.parameter.view) : "home";
  const id = (e && e.parameter && e.parameter.id) ? String(e.parameter.id) : "";
  const template = HtmlService.createTemplateFromFile("Shell");
  template.initialView = view;
  template.initialId = id;
  template.menuHtml = getMenuHtml(view, true);
  return template
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle("Sistema de Gestão de Cursos")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
 * Endpoint HTTP (OpenAPI) – Criação de eventos de agendamento.
 * Camada Controller (entrada): recebe JSON e delega ao ControllerAgendamento.
 *
 * CONTRATO PARA O CLIENTE (obrigatório):
 * - Do NOT use HTTP status codes.
 * - Do NOT rely on response.ok.
 * - ONLY use the JSON body to determine success or failure:
 *   - Sucesso: body.status === "ok" → use body.eventos, body.total.
 *   - Erro: body.status === "erro" → use body.code, body.message, body.details.
 *
 * @param {Object} e - e.postData.contents contém o JSON do request.
 * @returns {GoogleAppsScript.Content.TextOutput}
 */
function doPost(e) {
  let request = null;
  const bodyText = (e && e.postData && e.postData.contents) ? String(e.postData.contents) : "";
  try {
    request = bodyText ? JSON.parse(bodyText) : {};
  } catch (_) {
    request = null;
  }

  const resp = criarEventosAgendamentoController(request);
  return ContentService
    .createTextOutput(JSON.stringify(resp))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Retorna o HTML do menu principal. Se spa=true, links usam data-view e href="#".
 * @param {string} view - "home", "consulta" ou "cadastro".
 * @param {boolean} spa - Se true, menu para SPA (navegação client-side).
 * @returns {string}
 */
function getMenuHtml(view, spa) {
  const t = HtmlService.createTemplateFromFile("Menu");
  t.view = view || "";
  t.urlConsulta = spa ? "#" : obterUrlWebApp("consulta");
  t.spa = spa === true;
  return t.evaluate().getContent();
}

/**
 * Retorna conteúdo HTML e script de uma view para injeção no SPA (chamado pelo cliente).
 * @param {string} view - "home", "consulta" ou "cadastro".
 * @param {string} id - ID do registro (só para cadastro/edição).
 * @returns {{ html: string, script: string }}
 */
function getPageContent(view, id) {
  view = view || "home";
  id = (id && String(id).trim()) ? String(id).trim() : "";
  if (view === "home") {
    const t = HtmlService.createTemplateFromFile("HomeFragment");
    return { html: t.evaluate().getContent(), script: "" };
  }
  if (view === "consulta") {
    const t = HtmlService.createTemplateFromFile("ConsultaFragment");
    t.spa = true;
    t.parentItem = "Turma";
    t.subItem = "Consulta";
    return {
      html: t.evaluate().getContent(),
      script: HtmlService.createHtmlOutputFromFile("ConsultaJavaScript").getContent()
    };
  }
  if (view === "cadastro") {
    const t = HtmlService.createTemplateFromFile("IndexFragment");
    t.id = id;
    t.spa = true;
    t.parentItem = "Turma";
    t.subItem = id ? "Edição" : "Cadastro";
    const scriptCadastro = HtmlService.createHtmlOutputFromFile("JavaScript").getContent();
    const scriptUpdateStyle = "<script>function updateStyle(el){if(el.value&&el.value!=='')el.classList.remove('is-placeholder');else el.classList.add('is-placeholder');}<\/script>";
    return {
      html: t.evaluate().getContent(),
      script: scriptUpdateStyle + scriptCadastro
    };
  }
  return { html: "", script: "" };
}

/**
 * Retorna a URL do Web App com query view=... (ou base sem query para home).
 * @param {string} visualizacao - "home", "consulta" ou "cadastro".
 * @returns {string}
 */
function obterUrlWebApp(visualizacao) {
  const base = ScriptApp.getService().getUrl();
  if (!visualizacao || String(visualizacao) === "home") return base;
  return base + "?view=" + encodeURIComponent(String(visualizacao));
}
