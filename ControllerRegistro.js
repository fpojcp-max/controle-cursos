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
 * @param {{ key: string, dir: string }} ordenacao
 * @param {{ offset: number, limit: number }} paginacao
 * @returns {{ columns: any[], rows: any[][], total: number, truncated: boolean }}
 */
function pesquisarRegistros(filtros, ordenacao, paginacao) {
  return buscarRegistrosComFiltros(filtros, ordenacao, paginacao);
}

/**
 * Salva um novo registro (cadastro).
 * @param {Object} dados - Objeto com campos do formulário.
 * @returns {string} Mensagem de sucesso ou erro.
 */
function salvarRegistro(dados) {
  try {
    return cadastrarRegistro(dados);
  } catch (e) {
    return "Erro ao salvar: " + e.toString();
  }
}

/**
 * Atualiza um registro existente pelo ID.
 * @param {string} id - UUID do registro.
 * @param {Object} dados - Objeto com campos do formulário.
 * @returns {string} Mensagem de sucesso ou erro.
 */
function atualizarRegistro(id, dados) {
  try {
    return atualizarRegistroPorId(id, dados);
  } catch (e) {
    return "Erro ao atualizar: " + e.toString();
  }
}

/**
 * Obtém um registro por ID (para edição).
 * @param {string} id - UUID do registro.
 * @returns {Object|null} Objeto com campos do registro ou null.
 */
function obterRegistroPorId(id) {
  return obterRegistroPorIdServico(id);
}

/**
 * Executa a migração que preenche a coluna ID nos registros antigos.
 * @returns {{ ok: boolean, message: string }}
 */
function executarMigracaoIds() {
  return preencherColunaId();
}
