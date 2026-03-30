/**
 * Camada Controller – Entrada do cliente (google.script.run).
 */

/**
 * Retorna a configuração do formulário (listas para selects).
 * @returns {Object} CURSOS, TURMAS, OFERTAS, RESPONSAVEIS, PRIORIDADES, STATUS, BOOLEANOS
 */
function obterConfiguracaoFormulario() {
  return obterOpcoesFormulario();
}

/**
 * Pesquisa registros com filtros, ordenação e paginação.
 * @param {Object} filtros
 * @param {{ key?: string, dir?: string, keys?: { key: string, dir: string }[] }} ordenacao
 * @param {{ offset: number, limit: number }} paginacao
 * @returns {{ columns: any[], rows: any[][], total: number, truncated: boolean }}
 */
function pesquisarRegistros(filtros, ordenacao, paginacao) {
  return RegistroService.buscarRegistrosComFiltros(filtros, ordenacao, paginacao);
}

/**
 * Salva um novo registro (cadastro).
 * @param {Object} dados - Objeto com campos do formulário.
 * @returns {{ success: boolean, message: string }}
 */
function salvarRegistro(dados) {
  try {
    const message = RegistroService.cadastrarRegistro(dados);
    return { success: true, message: message };
  } catch (e) {
    return { success: false, message: (e && e.message) ? e.message : String(e) };
  }
}

/**
 * Atualiza um registro existente pelo ID.
 * @param {string} id - UUID do registro.
 * @param {Object} dados - Objeto com campos do formulário.
 * @returns {{ success: boolean, message: string }}
 */
function atualizarRegistro(id, dados) {
  try {
    const message = RegistroService.atualizarRegistroPorId(id, dados);
    return { success: true, message: message };
  } catch (e) {
    return { success: false, message: (e && e.message) ? e.message : String(e) };
  }
}

/**
 * Obtém um registro por ID (para edição).
 * @param {string} id - UUID do registro.
 * @returns {Object|null} Objeto com campos do registro ou null.
 */
function obterRegistroPorId(id) {
  return RegistroService.obterRegistroPorIdServico(id);
}

/**
 * Exclui um registro por ID (exclusão física).
 * @param {string} id - UUID do registro.
 * @returns {string} Mensagem de sucesso ou erro.
 */
function excluirRegistro(id) {
  try {
    return RegistroService.excluirRegistroPorId(id);
  } catch (e) {
    return "Erro ao excluir: " + e.toString();
  }
}

/**
 * Listas para filtros da tela Turma >> Excluir (valores distintos na planilha).
 * @returns {{ success: boolean, cursos?: string[], responsaveis?: string[], message?: string }}
 */
function obterOpcoesFiltroTurmaExcluir() {
  try {
    return {
      success: true,
      cursos: RegistroRepo.listarCursosDistintos(),
      responsaveis: RegistroRepo.listarResponsaveisDistintos()
    };
  } catch (e) {
    return {
      success: false,
      message: (e && e.message) ? e.message : String(e)
    };
  }
}

/**
 * Listas para filtros da tela Turma >> Editar (apenas cursos; turmas por curso via obterTurmasPlanilhaPorCurso).
 * @returns {{ success: boolean, cursos?: string[], message?: string }}
 */
function obterOpcoesFiltroTurmaEditar() {
  try {
    return { success: true, cursos: RegistroRepo.listarCursosDistintos() };
  } catch (e) {
    return {
      success: false,
      message: (e && e.message) ? e.message : String(e)
    };
  }
}

/**
 * Pesquisa turma por curso + turma (edição individual); retorno alinhado ao uso na tela Editar.
 * @param {string} curso
 * @param {string} turma
 * @returns {{ success: boolean, columns?: any[], rows?: any[][], total?: number, idColumnIndex?: number, message?: string }}
 */
function pesquisarTurmasParaEditar(curso, turma) {
  try {
    const c = String(curso || "").trim();
    const t = String(turma || "").trim();
    if (!c || !t) {
      return { success: false, message: "Selecione curso e turma para pesquisar." };
    }
    const r = RegistroService.pesquisarRegistrosTurmaEditarTela(c, t);
    return {
      success: true,
      columns: r.columns,
      rows: r.rows,
      total: r.total,
      idColumnIndex: typeof r.idColumnIndex === "number" ? r.idColumnIndex : -1
    };
  } catch (e) {
    return {
      success: false,
      message: (e && e.message) ? e.message : String(e)
    };
  }
}

/**
 * Turmas distintas na planilha para o curso (cadastro de turmas).
 * @param {string} curso
 * @returns {{ success: boolean, turmas?: string[], message?: string }}
 */
function obterTurmasPlanilhaPorCurso(curso) {
  try {
    return {
      success: true,
      turmas: RegistroRepo.listarTurmasDistintasPorCurso(curso || "")
    };
  } catch (e) {
    return {
      success: false,
      message: (e && e.message) ? e.message : String(e)
    };
  }
}

/**
 * Pesquisa paginada para exclusão de turmas + IDs de todo o conjunto filtrado (selecionar tudo).
 */
function pesquisarTurmasParaExcluir(filtros, ordenacao, paginacao) {
  return RegistroService.buscarRegistrosExcluirTurmaTela(filtros, ordenacao, paginacao);
}

/**
 * Linhas atuais da planilha para os IDs indicados (ordem preservada).
 * @param {string[]} ids
 * @returns {{ success: boolean, columns?: any[], rows?: any[][], message?: string }}
 */
function obterRegistrosTurmaPorIdsParaExcluir(ids) {
  try {
    const r = RegistroService.obterLinhasRegistroTurmaPorIdsNaOrdem(ids);
    return {
      success: true,
      columns: r.columns,
      rows: r.rows,
      idColumnIndex: typeof r.idColumnIndex === "number" ? r.idColumnIndex : -1
    };
  } catch (e) {
    return {
      success: false,
      message: (e && e.message) ? e.message : String(e)
    };
  }
}

/**
 * Exclusão em lote de turmas por ID (com limpeza de agendamentos por turma).
 * @param {string[]} ids
 * @returns {{ success: boolean, tipo?: string, mensagemTopo?: string, excluidos?: Object[], falhas?: Object[], message?: string }}
 */
function excluirRegistrosTurmaLote(ids) {
  try {
    const lista = Array.isArray(ids) ? ids : [];
    if (!lista.length) {
      return { success: false, message: "Selecione ao menos uma turma." };
    }
    const r = RegistroService.excluirRegistrosTurmaLote(lista);
    return {
      success: true,
      tipo: r.tipo,
      mensagemTopo: r.mensagemTopo,
      excluidos: r.excluidos,
      falhas: r.falhas
    };
  } catch (e) {
    return {
      success: false,
      message: (e && e.message) ? e.message : String(e)
    };
  }
}

/**
 * Retorna todos os registros (filtros + ordenação) para exportação CSV.
 * @param {Object} filtros
 * @param {{ key?: string, dir?: string, keys?: { key: string, dir: string }[] }} ordenacao
 * @returns {{ columns: any[], rows: any[][] }}
 */
function obterRegistrosParaExportar(filtros, ordenacao) {
  return RegistroService.buscarRegistrosParaExportar(filtros, ordenacao);
}

/**
 * Executa a migração que preenche a coluna ID nos registros antigos.
 * @returns {{ ok: boolean, message: string }}
 */
function executarMigracaoIds() {
  return RegistroRepo.preencherColunaId();
}
