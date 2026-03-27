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
  template.menuHtml = getMenuHtml(view, true, id);
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
 * @param {string} [cadastroId] - id na URL quando view=cadastro (edição); vazio = inclusão.
 * @returns {string}
 */
function getMenuHtml(view, spa, cadastroId) {
  const idNorm =
    cadastroId && String(cadastroId).trim() ? String(cadastroId).trim() : "";
  const v = view || "";
  const t = HtmlService.createTemplateFromFile("Menu");
  t.view = v;
  t.menuConsultaAtiva = v === "consulta" || (v === "cadastro" && idNorm.length > 0);
  t.menuInserirAtivo = v === "cadastro" && idNorm.length === 0;
  t.menuTurmaExcluirAtivo = v === "turma-excluir";
  t.menuTurmaEditarAtivo = v === "turma-editar";
  t.urlConsulta = spa ? "#" : obterUrlWebApp("consulta");
  t.urlCadastroInserir = spa ? "#" : obterUrlWebApp("cadastro");
  t.urlAgendamentoIncluir = spa ? "#" : obterUrlWebApp("agendamento-incluir");
  t.urlAgendamentoConsulta = spa ? "#" : obterUrlWebApp("agendamento-consulta");
  t.urlAgendamentoEditar = spa ? "#" : obterUrlWebApp("agendamento-editar");
  t.urlAgendamentoExcluir = spa ? "#" : obterUrlWebApp("agendamento-excluir");
  t.urlTurmaExcluir = spa ? "#" : obterUrlWebApp("turma-excluir");
  t.urlTurmaEditar = spa ? "#" : obterUrlWebApp("turma-editar");
  t.menuTurmaGrupoAtivo =
    v === "consulta" || v === "cadastro" || v === "turma-editar" || v === "turma-excluir";
  t.menuAgendamentoAtivo =
    v === "agendamento-incluir" ||
    v === "agendamento-consulta" ||
    v === "agendamento-editar" ||
    v === "agendamento-excluir";
  t.spa = spa === true;
  return t.evaluate().getContent();
}

/**
 * Retorna conteúdo HTML e script de uma view para injeção no SPA (chamado pelo cliente).
 * @param {string} view - "home", "consulta", "cadastro" ou "agendamento-incluir".
 * @param {string} id - ID do registro (cadastro/edição ou agendamento).
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
    t.subItem = "Consultar";
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
  if (view === "agendamento-incluir") {
    const t = HtmlService.createTemplateFromFile("AgendamentoIncluirFragment");
    t.id = id;
    t.spa = true;
    t.parentItem = "Agendamento";
    t.subItem = "Incluir";
    return {
      html: t.evaluate().getContent(),
      script: HtmlService.createHtmlOutputFromFile("AgendamentoIncluirJavaScript").getContent()
    };
  }
  if (view === "agendamento-consulta") {
    const t = HtmlService.createTemplateFromFile("AgendamentoConsultaFragment");
    t.id = id;
    t.spa = true;
    t.parentItem = "Agendamento";
    t.subItem = "Consultar";
    return {
      html: t.evaluate().getContent(),
      script: HtmlService.createHtmlOutputFromFile("AgendamentoConsultaJavaScript").getContent()
    };
  }
  if (view === "agendamento-editar") {
    const t = HtmlService.createTemplateFromFile("AgendamentoEditarFragment");
    t.id = id;
    t.spa = true;
    t.parentItem = "Agendamento";
    t.subItem = "Editar";
    return {
      html: t.evaluate().getContent(),
      script: HtmlService.createHtmlOutputFromFile("AgendamentoEditarJavaScript").getContent()
    };
  }
  if (view === "agendamento-excluir") {
    const t = HtmlService.createTemplateFromFile("AgendamentoExcluirFragment");
    t.id = id;
    t.spa = true;
    t.parentItem = "Agendamento";
    t.subItem = "Excluir";
    return {
      html: t.evaluate().getContent(),
      script: HtmlService.createHtmlOutputFromFile("AgendamentoExcluirJavaScript").getContent()
    };
  }
  if (view === "turma-excluir") {
    const t = HtmlService.createTemplateFromFile("TurmaExcluirFragment");
    t.id = id;
    t.spa = true;
    t.parentItem = "Turma";
    t.subItem = "Excluir";
    return {
      html: t.evaluate().getContent(),
      script: HtmlService.createHtmlOutputFromFile("TurmaExcluirJavaScript").getContent()
    };
  }
  if (view === "turma-editar") {
    const t = HtmlService.createTemplateFromFile("TurmaEditarFragment");
    t.id = id;
    t.spa = true;
    t.parentItem = "Turma";
    t.subItem = "Editar";
    return {
      html: t.evaluate().getContent(),
      script: HtmlService.createHtmlOutputFromFile("TurmaEditarJavaScript").getContent()
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
