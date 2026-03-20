/**
 * Camada Controller – Agendamento: API (payload estruturado) e Web App (google.script.run).
 */

const AgendamentoController = (() => {
  function errorResponse_(code, message, details) {
    return {
      status: "erro",
      message: String(message || "Erro"),
      code: code || "INTERNAL_ERROR",
      details: Array.isArray(details) ? details : []
    };
  }

  function criarEventos_(request) {
    try {
      return AgendamentoService.criarEventos(request);
    } catch (err) {
      if (err && err.code && typeof err.message === "string") {
        return errorResponse_(err.code, err.message, err.details);
      }
      return errorResponse_("INTERNAL_ERROR", err && err.message ? err.message : "Erro interno", []);
    }
  }

  return {
    criarEventos: criarEventos_
  };
})();

function criarEventosAgendamentoController(request) {
  return AgendamentoController.criarEventos(request);
}

/**
 * Dados para montar a tela Incluir: cursos da planilha, salas, fuso.
 * @returns {{ success: boolean, cursos?: string[], salas?: Object[], timezone?: string, message?: string }}
 */
function obterDadosTelaAgendamentoIncluir() {
  try {
    const dados = AgendamentoService.obterDadosIncluir();
    return {
      success: true,
      cursos: dados.cursos,
      salas: dados.salas,
      timezone: dados.timezone
    };
  } catch (e) {
    return {
      success: false,
      message: e && e.message ? e.message : String(e)
    };
  }
}

/**
 * Turmas distintas existentes na planilha para o curso informado.
 * @param {string} curso
 * @returns {{ success: boolean, turmas?: string[], message?: string }}
 */
function obterTurmasAgendamentoPorCurso(curso) {
  try {
    const turmas = AgendamentoService.listarTurmasPorCursoIncluir(curso || "");
    return { success: true, turmas: turmas };
  } catch (e) {
    return {
      success: false,
      message: e && e.message ? e.message : String(e)
    };
  }
}

/**
 * Rode no editor (Executar) ou via google.script.run na Web App.
 * @returns {{ globalCalendarDefinido: boolean, usaServicoAvancado: boolean }}
 */
function diagnosticarCalendarAgendamento() {
  let globalOk = false;
  try {
    globalOk = typeof Calendar !== "undefined";
  } catch (e) {
    globalOk = false;
  }
  return {
    globalCalendarDefinido: globalOk,
    usaServicoAvancado: CalendarAdapter.usandoServicoAvancado()
  };
}

/**
 * Cria eventos no Calendar (um por ocorrência) e grava linhas na planilha de agendamentos.
 * @param {Object} payload
 * @returns {{ success: boolean, message?: string, ocorrencias?: number }}
 */
function criarAgendamentos(payload) {
  try {
    const r = AgendamentoService.criarAgendamentos(payload);
    return { success: true, message: r.mensagem, ocorrencias: r.ocorrencias };
  } catch (e) {
    return {
      success: false,
      message: e && e.message ? e.message : String(e)
    };
  }
}
