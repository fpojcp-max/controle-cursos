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
  let sessaoIdentificada = false;
  let acessoSistemaOk = false;
  let mensagemSemAcessoSistema = "";
  let podeGerirPermissoesInicial = false;
  try {
    const em = Session.getActiveUser().getEmail();
    sessaoIdentificada = !!(em && String(em).trim());
    if (sessaoIdentificada) {
      try {
        const perm = PermissaoRepo.obterPermissaoPorEmail(String(em).trim().toLowerCase());
        if (perm) {
          acessoSistemaOk = true;
          try {
            podeGerirPermissoesInicial = PermissaoService.usuarioPodeGerirPermissoes();
          } catch (eAdmin) {
            podeGerirPermissoesInicial = false;
          }
        } else {
          mensagemSemAcessoSistema = PermissaoService.MSG_SEM_CADASTRO;
        }
      } catch (errPlan) {
        acessoSistemaOk = false;
        mensagemSemAcessoSistema =
          errPlan && errPlan.message
            ? String(errPlan.message)
            : "Não foi possível aceder à planilha.";
      }
    }
  } catch (errSessao) {
    sessaoIdentificada = false;
    acessoSistemaOk = false;
  }
  const template = HtmlService.createTemplateFromFile("Shell");
  template.initialView = view;
  template.initialId = id;
  template.sessaoIdentificada = sessaoIdentificada;
  template.acessoSistemaOk = acessoSistemaOk;
  template.mensagemSemAcessoSistema = mensagemSemAcessoSistema;
  template.podeGerirPermissoesInicial = podeGerirPermissoesInicial;
  template.menuHtml = getMenuHtml(view, true, id, sessaoIdentificada, acessoSistemaOk);
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
 *
 * O POST direto à URL está desativado (não há identidade de visitante em chamadas HTTP anónimas).
 * Use a interface Web ou `google.script.run`.
 */
function doPost(e) {
  const resp = {
    status: "erro",
    code: "ENDPOINT_DISABLED",
    message:
      "Endpoint HTTP desativado. Utilize a aplicação Web (interface) com sessão Google do domínio; não use POST direto à URL.",
    details: []
  };
  return ContentService.createTextOutput(JSON.stringify(resp)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Verificação antecipada: planilha associada ao script acessível pelo utilizador da Web App.
 * @returns {{ ok: true } | { ok: false, code: string, message: string }}
 */
/**
 * Gate da UI (passo 1): há e-mail de sessão identificável na Web App.
 * @returns {{ ok: boolean, message?: string }}
 */
function obterEstadoGateSessaoExecucaoWebApp() {
  let email = "";
  try {
    email = Session.getActiveUser().getEmail();
  } catch (errSessao) {
    email = "";
  }
  email = String(email || "").trim();
  if (!email) {
    return { ok: false, message: SessaoWebApp.MSG_SEM_IDENTIDADE };
  }
  return { ok: true };
}

function verificarAcessoPlanilhaWebApp() {
  let email = "";
  try {
    email = Session.getActiveUser().getEmail();
  } catch (errSessao) {
    email = "";
  }
  email = String(email || "").trim();
  if (!email) {
    return { ok: false, code: "NO_SESSION", message: SessaoWebApp.MSG_SEM_IDENTIDADE };
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return {
        ok: false,
        code: "NO_SPREADSHEET",
        message: "Não há planilha associada a este projeto Apps Script."
      };
    }
    ss.getName();
  } catch (err) {
    const raw = err && err.message ? String(err.message) : String(err);
    return { ok: false, code: "SPREADSHEET_ACCESS", message: raw };
  }
  if (!PermissaoRepo.obterPermissaoPorEmail(email.toLowerCase())) {
    return {
      ok: false,
      code: "NO_PERMISSAO",
      message: PermissaoService.MSG_SEM_CADASTRO
    };
  }
  return { ok: true };
}

/**
 * Retorna o HTML do menu principal. Se spa=true, links usam data-view e href="#".
 * @param {string} view - "home", "consulta" ou "cadastro".
 * @param {boolean} spa - Se true, menu para SPA (navegação client-side).
 * @param {string} [cadastroId] - id na URL quando view=cadastro (edição); vazio = inclusão.
 * @param {boolean} [sessaoWebAppOk] - se false, o menu não oferece navegação (sem identidade na sessão).
 * @param {boolean} [acessoSistemaOk] - se false, utilizador sem linha na aba Permissoes.
 * @returns {string}
 */
function getMenuHtml(view, spa, cadastroId, sessaoWebAppOk, acessoSistemaOk) {
  const idNorm =
    cadastroId && String(cadastroId).trim() ? String(cadastroId).trim() : "";
  const v = view || "";
  const t = HtmlService.createTemplateFromFile("Menu");
  t.sessaoWebAppOk = sessaoWebAppOk === true && acessoSistemaOk === true;
  t.mensagemSemSessaoWebApp = acessoSistemaOk === false && sessaoWebAppOk === true
    ? PermissaoService.MSG_SEM_CADASTRO
    : SessaoWebApp.MSG_SEM_IDENTIDADE;
  t.exibirMenuPermissoes = false;
  if (sessaoWebAppOk === true && acessoSistemaOk === true) {
    try {
      t.exibirMenuPermissoes = PermissaoService.usuarioPodeGerirPermissoes();
    } catch (eMenu) {
      t.exibirMenuPermissoes = false;
    }
  }
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
  t.menuPermissoesAtivo = v === "permissoes-incluir" || v === "permissoes-consultar";
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
  SessaoWebApp.exigirParaGoogleScriptRun();
  view = view || "home";
  id = (id && String(id).trim()) ? String(id).trim() : "";
  if (view === "permissoes-incluir" || view === "permissoes-consultar") {
    if (!PermissaoService.usuarioPodeGerirPermissoes()) {
      throw new Error(PermissaoService.MSG_NEGADO_GERIR);
    }
  }
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
  if (view === "permissoes-incluir") {
    const t = HtmlService.createTemplateFromFile("PermissoesIncluirFragment");
    t.spa = true;
    t.parentItem = "Permissão";
    t.subItem = "Incluir";
    return {
      html: t.evaluate().getContent(),
      script: HtmlService.createHtmlOutputFromFile("PermissoesIncluirJavaScript").getContent()
    };
  }
  if (view === "permissoes-consultar") {
    const t = HtmlService.createTemplateFromFile("PermissoesConsultaFragment");
    t.spa = true;
    t.parentItem = "Permissão";
    t.subItem = "Consultar";
    return {
      html: t.evaluate().getContent(),
      script: HtmlService.createHtmlOutputFromFile("PermissoesConsultaJavaScript").getContent()
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
  SessaoWebApp.exigirParaGoogleScriptRun();
  const base = ScriptApp.getService().getUrl();
  if (!visualizacao || String(visualizacao) === "home") return base;
  return base + "?view=" + encodeURIComponent(String(visualizacao));
}

/**
 * Identidade e capacidades do visitante (passo 2 do gate da UI).
 * Não lança erro quando o utilizador não está na aba Permissoes — devolve flags.
 * @returns {{
 *   email: string,
 *   logoutUrl: string,
 *   acessoSistemaOk: boolean,
 *   podeGerirPermissoes: boolean,
 *   perfil: string,
 *   mensagemSemAcesso: string
 * }}
 */
function obterIdentidadeCabecalhoWebApp() {
  const logoutUrl = String(Configuracoes.URL_LOGOUT_SSO || "").trim();
  let email = "";
  try {
    email = Session.getActiveUser().getEmail();
  } catch (errSessao) {
    email = "";
  }
  email = String(email || "").trim().toLowerCase();
  if (!email) {
    return {
      email: "",
      logoutUrl: logoutUrl,
      acessoSistemaOk: false,
      podeGerirPermissoes: false,
      perfil: "",
      mensagemSemAcesso: SessaoWebApp.MSG_SEM_IDENTIDADE
    };
  }
  let perm = null;
  try {
    perm = PermissaoRepo.obterPermissaoPorEmail(email);
  } catch (errPlan) {
    return {
      email: email,
      logoutUrl: logoutUrl,
      acessoSistemaOk: false,
      podeGerirPermissoes: false,
      perfil: "",
      mensagemSemAcesso:
        errPlan && errPlan.message ? String(errPlan.message) : "Não foi possível aceder à planilha."
    };
  }
  if (!perm) {
    return {
      email: email,
      logoutUrl: logoutUrl,
      acessoSistemaOk: false,
      podeGerirPermissoes: false,
      perfil: "",
      mensagemSemAcesso: PermissaoService.MSG_SEM_CADASTRO
    };
  }
  let podeGerir = false;
  try {
    podeGerir = PermissaoService.usuarioPodeGerirPermissoes();
  } catch (eGerir) {
    podeGerir = false;
  }
  return {
    email: email,
    logoutUrl: logoutUrl,
    acessoSistemaOk: true,
    podeGerirPermissoes: podeGerir,
    perfil: String(perm.perfil || "").trim(),
    mensagemSemAcesso: ""
  };
}
