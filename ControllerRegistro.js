/**
 * Camada Controller – Entrada do cliente (google.script.run).
 */

/**
 * Retorna a configuração do formulário (listas para selects).
 * @returns {Object} CURSOS, TURMAS, SALAS, OFERTAS, RESPONSAVEIS, PRIORIDADES, STATUS, BOOLEANOS
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
