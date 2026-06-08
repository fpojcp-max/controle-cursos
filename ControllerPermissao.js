/**
 * Camada Controller – Permissões (Web App).
 */

function obterOpcoesCadastroPermissoes() {
  SessaoWebApp.exigirParaGoogleScriptRun();
  try {
    PermissaoService.garantirPodeGerirPermissoes();
    return PermissaoService.opcoesCadastroPermissoes();
  } catch (e) {
    return { perfis: [], erro: e && e.message ? e.message : String(e) };
  }
}

function incluirPermissaoCadastro(dados) {
  SessaoWebApp.exigirParaGoogleScriptRun();
  try {
    return PermissaoService.incluirPermissaoCadastro(dados);
  } catch (e) {
    return { success: false, message: e && e.message ? e.message : String(e) };
  }
}

function obterOpcoesFiltroPermissoesConsulta() {
  SessaoWebApp.exigirParaGoogleScriptRun();
  return PermissaoService.obterOpcoesFiltroPermissoesConsulta();
}

function pesquisarPermissoesConsulta(filtros, ordenacao, paginacao) {
  SessaoWebApp.exigirParaGoogleScriptRun();
  return PermissaoService.pesquisarPermissoesConsulta(filtros, ordenacao, paginacao);
}

function obterPermissoesConsultaParaExportar(filtros, ordenacao) {
  SessaoWebApp.exigirParaGoogleScriptRun();
  return PermissaoService.obterPermissoesConsultaParaExportar(filtros, ordenacao);
}

function excluirPermissaoConsulta(sheetRow) {
  SessaoWebApp.exigirParaGoogleScriptRun();
  return PermissaoService.excluirPermissaoConsulta(sheetRow);
}

function usuarioPodeGerirPermissoesWebApp() {
  try {
    SessaoWebApp.exigirParaGoogleScriptRun();
    return { podeGerir: PermissaoService.usuarioPodeGerirPermissoes() };
  } catch (e) {
    return { podeGerir: false };
  }
}
