/**
 * Camada Data – Catálogo de salas (rótulo exibido em selects e na planilha).
 */

/**
 * @returns {{ rotulo: string }[]}
 */
function obterEntradasCatalogoRecursosSala() {
  const lista = Configuracoes.CATALOGO_RECURSOS_SALA || [];
  return lista.map((e) => ({
    rotulo: String(e && e.rotulo != null ? e.rotulo : "").trim()
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
 * Mapa rótulo → true (validação de sala reconhecida).
 * @returns {Object<string, boolean>}
 */
function montarConjuntoRotulosSalaValidos() {
  const set = {};
  listarRotulosCatalogoRecursosSala().forEach((r) => {
    set[r] = true;
  });
  return set;
}
