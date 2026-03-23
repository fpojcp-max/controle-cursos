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

/**
 * Lista agendamentos da planilha para exclusão (filtro curso + turma, paginado).
 * @param {{ curso: string, turma: string, offset?: number, limit?: number, sortCol?: number, sortDir?: string }} payload
 * sortCol -1 = data + hora início (padrão); 0..n-1 = índice da coluna no cabeçalho da planilha.
 */
function pesquisarAgendamentosExcluir(payload) {
  try {
    const p = payload && typeof payload === "object" ? payload : {};
    const off = p.offset != null ? Number(p.offset) : 0;
    const lim = p.limit != null ? Number(p.limit) : 50;
    let sortCol = -1;
    if (p.sortCol !== undefined && p.sortCol !== null && p.sortCol !== "") {
      const n = Number(p.sortCol);
      if (!isNaN(n)) sortCol = n;
    }
    const sortDir =
      p.sortDir != null && String(p.sortDir).trim() !== "" ? String(p.sortDir) : "asc";
    return AgendamentoService.pesquisarAgendamentosExcluir(
      p.curso,
      p.turma,
      off,
      lim,
      sortCol,
      sortDir
    );
  } catch (e) {
    return {
      success: false,
      message: e && e.message ? e.message : String(e)
    };
  }
}

/**
 * Exportação CSV: todos os agendamentos do filtro (curso+turma) com a ordenação atual.
 * @param {{ curso: string, turma: string, sortCol?: number, sortDir?: string }} payload
 * @returns {{ columns: { key: string, label: string }[], rows: string[][] }}
 */
function obterAgendamentosConsultaParaExportar(payload) {
  const p = payload && typeof payload === "object" ? payload : {};
  let sortCol = -1;
  if (p.sortCol !== undefined && p.sortCol !== null && p.sortCol !== "") {
    const n = Number(p.sortCol);
    if (!isNaN(n)) sortCol = n;
  }
  const sortDir =
    p.sortDir != null && String(p.sortDir).trim() !== "" ? String(p.sortDir) : "asc";
  return AgendamentoService.obterAgendamentosConsultaParaExportar(
    p.curso,
    p.turma,
    sortCol,
    sortDir
  );
}

/**
 * Todos os eventIds do filtro atual (para “Selecionar tudo”).
 */
function obterTodosEventIdsAgendamentoExcluir(curso, turma) {
  try {
    return AgendamentoService.obterTodosEventIdsExcluir(curso, turma);
  } catch (e) {
    return {
      success: false,
      message: e && e.message ? e.message : String(e)
    };
  }
}

/**
 * Exclui agendamentos (Calendar + planilha), tudo ou nada por tentativa.
 * @param {{ curso: string, turma: string, sheetRows?: number[], eventIds?: string[] }} payload
 */
function excluirAgendamentosLote(payload) {
  try {
    const p = payload && typeof payload === "object" ? payload : {};
    const r = AgendamentoService.excluirAgendamentosLote(p.curso, p.turma, p);
    return { success: true, message: r.mensagem };
  } catch (e) {
    return {
      success: false,
      message: e && e.message ? e.message : String(e)
    };
  }
}
