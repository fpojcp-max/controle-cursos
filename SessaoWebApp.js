/**
 * Identidade na Web App (executeAs: USER_ACCESSING).
 * Sem e-mail de sessão não há operações nem consultas via google.script.run.
 */

const SessaoWebApp = (() => {
  const MSG_SEM_IDENTIDADE =
    "Não foi possível identificar o seu utilizador. Aceda já autenticado com a conta do domínio (Web App como «quem acede»).";

  const MSG_NAO_RESPONSAVEL = "Acesso negado. Você não é o responsável pela turma.";

  const MSG_SEM_EMAIL_CRIADOR_TURMA =
    "Esta turma não tem o e-mail do responsável registado na planilha; não é possível validar o acesso.";

  const MSG_SEM_EMAIL_CRIADOR_AG =
    "Este agendamento não tem o criador registado na planilha; não é possível validar o acesso.";

  function obterEmailAtivoNormalizado_() {
    let email = "";
    try {
      email = Session.getActiveUser().getEmail();
    } catch (e) {
      email = "";
    }
    email = String(email || "").trim();
    if (!email) {
      throw new Error(MSG_SEM_IDENTIDADE);
    }
    return email.toLowerCase();
  }

  /**
   * Compara e-mail do criador (planilha) com o utilizador ativo.
   * @param {string} emailArmazenado - Valor bruto na célula.
   * @param {string} mensagemSeDiferente - Ex.: MSG_NAO_RESPONSAVEL.
   * @param {string} mensagemSeVazio - Quando não há criador registado.
   */
  function garantirMesmoUsuarioQueEmailArmazenadoOuErro_(emailArmazenado, mensagemSeDiferente, mensagemSeVazio) {
    const eu = obterEmailAtivoNormalizado_();
    const arm = String(emailArmazenado != null ? emailArmazenado : "").trim();
    if (!arm) {
      throw new Error(mensagemSeVazio || MSG_SEM_EMAIL_CRIADOR_TURMA);
    }
    if (arm.toLowerCase() !== eu) {
      throw new Error(mensagemSeDiferente || MSG_NAO_RESPONSAVEL);
    }
  }

  return {
    MSG_SEM_IDENTIDADE: MSG_SEM_IDENTIDADE,
    MSG_NAO_RESPONSAVEL: MSG_NAO_RESPONSAVEL,
    MSG_SEM_EMAIL_CRIADOR_TURMA: MSG_SEM_EMAIL_CRIADOR_TURMA,
    MSG_SEM_EMAIL_CRIADOR_AG: MSG_SEM_EMAIL_CRIADOR_AG,
    /** Primeira linha das funções expostas a `google.script.run`. */
    exigirParaGoogleScriptRun: obterEmailAtivoNormalizado_,
    obterEmailAtivoNormalizado: obterEmailAtivoNormalizado_,
    garantirMesmoUsuarioQueEmailArmazenadoOuErro: garantirMesmoUsuarioQueEmailArmazenadoOuErro_
  };
})();
