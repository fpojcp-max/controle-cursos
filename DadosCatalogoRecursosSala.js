/**
 * Camada Data – Catálogo único de recursos de sala (rótulo exibido + identificador no Google Calendar).
 * Consumido por formulários, filtros e resolução de recursos; nomes neutros em relação a telas específicas.
 */

/**
 * @returns {{ rotulo: string, identificadorCalendario: string }[]}
 */
function obterEntradasCatalogoRecursosSala() {
  const lista = Configuracoes.CATALOGO_RECURSOS_SALA || [];
  return lista.map((e) => ({
    rotulo: String(e && e.rotulo != null ? e.rotulo : "").trim(),
    identificadorCalendario: String(e && e.identificadorCalendario != null ? e.identificadorCalendario : "").trim()
  }));
}

/**
 * Rótulos distintos na ordem do catálogo (para selects).
 * @returns {string[]}
 */
function listarRotulosCatalogoRecursosSala() {
  const out = [];
  const seen = {};
  obterEntradasCatalogoRecursosSala().forEach((e) => {
    if (!e.rotulo || seen[e.rotulo]) return;
    seen[e.rotulo] = true;
    out.push(e.rotulo);
  });
  return out;
}

/**
 * Mapa rótulo → identificador de calendário (para APIs que precisam do ID do recurso).
 * @returns {Object<string, string>}
 */
function montarMapaRotuloParaIdentificadorCalendario() {
  const map = {};
  obterEntradasCatalogoRecursosSala().forEach((e) => {
    if (e.rotulo) map[e.rotulo] = e.identificadorCalendario;
  });
  return map;
}
