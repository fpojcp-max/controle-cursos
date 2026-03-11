/**
 * Camada Service – Regras de negócio, filtros, ordenação, conversão DTO.
 */

// ---- Helpers de data e número ----

function interpretarDataIso(str) {
  if (!str) return null;
  const s = String(str).trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return null;
  const [y, m, d] = s.split("-").map(Number);
  const dt = new Date(y, m - 1, d);
  return isNaN(dt.getTime()) ? null : dt;
}

function interpretarData(valor) {
  if (!valor) return null;
  if (valor instanceof Date) return isNaN(valor.getTime()) ? null : valor;
  const s = String(valor).trim();
  const iso = interpretarDataIso(s);
  if (iso) return iso;
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) {
    const dt = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
    return isNaN(dt.getTime()) ? null : dt;
  }
  const dt2 = new Date(s);
  return isNaN(dt2.getTime()) ? null : dt2;
}

function finalDoDia(data) {
  const dt = new Date(data.getTime());
  dt.setHours(23, 59, 59, 999);
  return dt;
}

function interpretarNumero(valor) {
  if (valor === null || valor === undefined) return null;
  const s = String(valor).trim().replace(/\./g, "").replace(",", ".");
  if (!s) return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

// ---- Colunas e filtros/ordenação ----

function normalizarColunasDoCabecalho(linhaCabecalho) {
  return (linhaCabecalho || []).map((h, idx) => {
    const label = (h && String(h).trim()) ? String(h).trim() : "Coluna " + (idx + 1);
    return { key: label, label: label };
  });
}

function indiceDaColuna(colunas, chave) {
  if (!chave) return -1;
  const k = String(chave).trim().toLowerCase();
  for (let i = 0; i < colunas.length; i++) {
    if (String(colunas[i].key).trim().toLowerCase() === k) return i;
  }
  const aliases = {
    "curso": "Curso", "turma": "Turma", "status": "Status", "sala": "Sala",
    "responsavel": "Responsável", "oferta": "Oferta", "datacadastro": "Data Cadastro", "id": "ID"
  };
  const aliased = aliases[k];
  if (aliased) return indiceDaColuna(colunas, aliased);
  return -1;
}

function aplicarFiltros(linhas, colunas, filtros) {
  const idxCurso = indiceDaColuna(colunas, "Curso");
  const idxTurma = indiceDaColuna(colunas, "Turma");
  const idxStatus = indiceDaColuna(colunas, "Status");
  const idxSala = indiceDaColuna(colunas, "Sala");
  const idxResp = indiceDaColuna(colunas, "Responsável");
  const idxOferta = indiceDaColuna(colunas, "Oferta");
  const idxInicio = indiceDaColuna(colunas, "Início");

  const curso = (filtros.curso || "").toString().trim();
  const turma = (filtros.turma || "").toString().trim();
  const status = (filtros.status || "").toString().trim();
  const sala = (filtros.sala || "").toString().trim();
  const responsavel = (filtros.responsavel || "").toString().trim();
  const oferta = (filtros.oferta || "").toString().trim();
  const inicioDe = interpretarDataIso(filtros.inicioDe);
  const inicioAte = interpretarDataIso(filtros.inicioAte);

  return (linhas || []).filter(r => {
    if (curso && idxCurso >= 0 && String(r[idxCurso] || "") !== curso) return false;
    if (turma && idxTurma >= 0 && String(r[idxTurma] || "") !== turma) return false;
    if (status && idxStatus >= 0 && String(r[idxStatus] || "") !== status) return false;
    if (sala && idxSala >= 0 && String(r[idxSala] || "") !== sala) return false;
    if (responsavel && idxResp >= 0 && String(r[idxResp] || "") !== responsavel) return false;
    if (oferta && idxOferta >= 0 && String(r[idxOferta] || "") !== oferta) return false;
    if ((inicioDe || inicioAte) && idxInicio >= 0) {
      const d = interpretarData(r[idxInicio]);
      if (inicioDe && (!d || d < inicioDe)) return false;
      if (inicioAte && (!d || d > finalDoDia(inicioAte))) return false;
    }
    return true;
  });
}

function aplicarOrdenacao(linhas, colunas, ordenacao) {
  const chave = ordenacao && ordenacao.key ? String(ordenacao.key) : "";
  const dir = (ordenacao && ordenacao.dir ? String(ordenacao.dir) : "asc").toLowerCase() === "desc" ? -1 : 1;
  const idx = indiceDaColuna(colunas, chave);
  if (idx < 0) return linhas || [];
  const copia = (linhas || []).slice();
  copia.sort((a, b) => {
    const av = a[idx];
    const bv = b[idx];
    const ad = interpretarData(av);
    const bd = interpretarData(bv);
    if (ad && bd) return (ad.getTime() - bd.getTime()) * dir;
    const an = interpretarNumero(av);
    const bn = interpretarNumero(bv);
    if (an !== null && bn !== null) return (an - bn) * dir;
    const as = (av === null || av === undefined) ? "" : String(av).toLowerCase();
    const bs = (bv === null || bv === undefined) ? "" : String(bv).toLowerCase();
    if (as < bs) return -1 * dir;
    if (as > bs) return 1 * dir;
    return 0;
  });
  return copia;
}

// ---- Conversão linha ↔ DTO ----

function linhaParaDados(linha) {
  return {
    curso: linha[0],
    turma: linha[1],
    inicio: linha[2],
    fim: linha[3],
    sala: linha[4],
    oferta: linha[5],
    responsavel: linha[6],
    prioridade: linha[7],
    fimInscricoes: linha[8],
    status: linha[9],
    linkFSA: linha[10],
    tr: linha[11],
    termo: linha[12],
    plataformas: linha[13],
    coffee: linha[14],
    agendarSala: linha[15],
    qrInscricao: linha[16],
    linkInscricao: linha[17],
    qrFrequencia: linha[18],
    frequencia: linha[19],
    codVerificacao: linha[20],
    qrAvaliacao: linha[21],
    avaliacaoReacao: linha[22],
    prepDivulgacao: linha[23],
    divulgacao: linha[24],
    linkPagina: linha[25]
  };
}

function dadosParaLinha(dados, metadados) {
  return [
    dados.curso,
    dados.turma,
    dados.inicio,
    dados.fim,
    dados.sala,
    dados.oferta,
    dados.responsavel,
    dados.prioridade,
    dados.fimInscricoes,
    dados.status,
    dados.linkFSA,
    dados.tr,
    dados.termo,
    dados.plataformas,
    dados.coffee,
    dados.agendarSala,
    dados.qrInscricao,
    dados.linkInscricao,
    dados.qrFrequencia,
    dados.frequencia,
    dados.codVerificacao,
    dados.qrAvaliacao,
    dados.avaliacaoReacao,
    dados.prepDivulgacao,
    dados.divulgacao,
    dados.linkPagina,
    metadados ? metadados.dataCadastro : null,
    metadados ? metadados.emailUsuario : null,
    metadados ? metadados.id : null
  ];
}

// ---- Orquestração ----

/**
 * Busca registros com filtros, ordenação e paginação.
 */
function buscarRegistrosComFiltros(filtros, ordenacao, paginacao) {
  const { valores, temCabecalho } = lerDadosPlanilha();
  if (!valores.length) {
    return { columns: obterColunasPadrao(), rows: [], total: 0, truncated: false };
  }
  const colunas = temCabecalho ? normalizarColunasDoCabecalho(valores[0]) : obterColunasPadrao();
  const linhas = valores.slice(temCabecalho ? 1 : 0);
  const filtradas = aplicarFiltros(linhas, colunas, filtros || {});
  const ordenadas = aplicarOrdenacao(filtradas, colunas, ordenacao || { key: "", dir: "asc" });
  const offset = Math.max(0, Number(paginacao && paginacao.offset) || 0);
  const limit = Math.min(100, Math.max(1, Number(paginacao && paginacao.limit) || 50));
  const paginadas = ordenadas.slice(offset, offset + limit);
  return {
    columns: colunas,
    rows: paginadas,
    total: ordenadas.length,
    truncated: ordenadas.length > paginadas.length
  };
}

/**
 * Cadastra um novo registro (insere linha).
 */
function cadastrarRegistro(dados) {
  const emailUsuario = Session.getActiveUser().getEmail();
  const dataAtual = new Date();
  const id = Utilities.getUuid();
  const metadados = { dataCadastro: dataAtual, emailUsuario: emailUsuario, id: id };
  const linha = dadosParaLinha(dados, metadados);
  inserirLinha(linha);
  return "Registro salvo.";
}

/**
 * Atualiza um registro existente pelo ID (atualiza linha, não insere).
 */
function atualizarRegistroPorId(id, dados) {
  const indice = buscarIndiceLinhaPorId(id);
  if (indice === -1) throw new Error("Registro não encontrado para atualização.");
  const linhaAtual = obterLinhaPorIndice(indice);
  linhaAtual[0] = dados.curso;
  linhaAtual[1] = dados.turma;
  linhaAtual[2] = dados.inicio;
  linhaAtual[3] = dados.fim;
  linhaAtual[4] = dados.sala;
  linhaAtual[5] = dados.oferta;
  linhaAtual[6] = dados.responsavel;
  linhaAtual[7] = dados.prioridade;
  linhaAtual[8] = dados.fimInscricoes;
  linhaAtual[9] = dados.status;
  linhaAtual[10] = dados.linkFSA;
  linhaAtual[11] = dados.tr;
  linhaAtual[12] = dados.termo;
  linhaAtual[13] = dados.plataformas;
  linhaAtual[14] = dados.coffee;
  linhaAtual[15] = dados.agendarSala;
  linhaAtual[16] = dados.qrInscricao;
  linhaAtual[17] = dados.linkInscricao;
  linhaAtual[18] = dados.qrFrequencia;
  linhaAtual[19] = dados.frequencia;
  linhaAtual[20] = dados.codVerificacao;
  linhaAtual[21] = dados.qrAvaliacao;
  linhaAtual[22] = dados.avaliacaoReacao;
  linhaAtual[23] = dados.prepDivulgacao;
  linhaAtual[24] = dados.divulgacao;
  linhaAtual[25] = dados.linkPagina;
  atualizarLinha(indice, linhaAtual);
  return "Registro atualizado.";
}

/**
 * Obtém um registro por ID (para preencher formulário de edição).
 */

/*
function obterRegistroPorIdServico(id) {
  const linha = buscarLinhaPorId(id);
  if (!linha) return null;
  return linhaParaDados(linha);
}
  */

function obterRegistroPorIdServico(id) {
  const linha = buscarLinhaPorId(id);
  if (!linha) return null;
  return linhaParaDados(linha);
}

/**
 * Exclui um registro por ID (exclusão física na planilha).
 */
function excluirRegistroPorId(id) {
  const indice = buscarIndiceLinhaPorId(id);
  if (indice === -1) throw new Error("Registro não encontrado para exclusão.");
  removerLinha(indice);
  return "Registro excluído.";
}
