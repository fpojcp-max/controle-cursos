/**
 * Camada Service – Regras de negócio, filtros, ordenação, conversão DTO.
 *
 * Encapsulado em `RegistroService` para reduzir funções globais expostas.
 * As funções globais no final são wrappers para compatibilidade.
 */

const RegistroService = (() => {
  const LIMITE_EXPORTACAO = 10000;

  // Schema: define ordem/índices (A..Z) dos campos editáveis no formulário
  const CAMPOS_EDITAVEIS = [
    "curso", "turma", "inicio", "fim", "sala", "oferta", "responsavel", "prioridade",
    "fimInscricoes", "status", "linkFSA", "tr", "termo", "plataformas", "coffee",
    "agendarSala", "qrInscricao", "linkInscricao", "qrFrequencia", "frequencia",
    "codVerificacao", "qrAvaliacao", "avaliacaoReacao", "prepDivulgacao", "divulgacao", "linkPagina"
  ];
  const IDX_DATA_CADASTRO = CAMPOS_EDITAVEIS.length;
  const IDX_EMAIL_USUARIO = CAMPOS_EDITAVEIS.length + 1;
  const IDX_ID = CAMPOS_EDITAVEIS.length + 2;

  // ---- Helpers de data e número ----
  function interpretarDataIso_(str) {
    if (!str) return null;
    const s = String(str).trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return null;
    const [y, m, d] = s.split("-").map(Number);
    const dt = new Date(y, m - 1, d);
    return isNaN(dt.getTime()) ? null : dt;
  }

  function interpretarData_(valor) {
    if (!valor) return null;
    if (valor instanceof Date) return isNaN(valor.getTime()) ? null : valor;
    const s = String(valor).trim();
    const iso = interpretarDataIso_(s);
    if (iso) return iso;
    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) {
      const dt = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
      return isNaN(dt.getTime()) ? null : dt;
    }
    const dt2 = new Date(s);
    return isNaN(dt2.getTime()) ? null : dt2;
  }

  function finalDoDia_(data) {
    const dt = new Date(data.getTime());
    dt.setHours(23, 59, 59, 999);
    return dt;
  }

  function interpretarNumero_(valor) {
    if (valor === null || valor === undefined) return null;
    const s = String(valor).trim().replace(/\./g, "").replace(",", ".");
    if (!s) return null;
    const n = Number(s);
    return Number.isFinite(n) ? n : null;
  }

  function normalizarColunasDoCabecalho_(linhaCabecalho) {
    return (linhaCabecalho || []).map((h, idx) => {
      const label = (h && String(h).trim()) ? String(h).trim() : "Coluna " + (idx + 1);
      return { key: label, label: label };
    });
  }

  function indiceDaColuna_(colunas, chave) {
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
    if (aliased) return indiceDaColuna_(colunas, aliased);
    return -1;
  }

  function aplicarFiltros_(linhas, colunas, filtros) {
    const idxCurso = indiceDaColuna_(colunas, "Curso");
    const idxTurma = indiceDaColuna_(colunas, "Turma");
    const idxStatus = indiceDaColuna_(colunas, "Status");
    const idxSala = indiceDaColuna_(colunas, "Sala");
    const idxResp = indiceDaColuna_(colunas, "Responsável");
    const idxOferta = indiceDaColuna_(colunas, "Oferta");
    const idxInicio = indiceDaColuna_(colunas, "Início");

    const curso = (filtros.curso || "").toString().trim();
    const turma = (filtros.turma || "").toString().trim();
    const status = (filtros.status || "").toString().trim();
    const sala = (filtros.sala || "").toString().trim();
    const responsavel = (filtros.responsavel || "").toString().trim();
    const oferta = (filtros.oferta || "").toString().trim();
    const inicioDe = interpretarDataIso_(filtros.inicioDe);
    const inicioAte = interpretarDataIso_(filtros.inicioAte);

    return (linhas || []).filter(r => {
      if (curso && idxCurso >= 0 && String(r[idxCurso] || "") !== curso) return false;
      if (turma && idxTurma >= 0 && String(r[idxTurma] || "") !== turma) return false;
      if (status && idxStatus >= 0 && String(r[idxStatus] || "") !== status) return false;
      if (sala && idxSala >= 0 && String(r[idxSala] || "") !== sala) return false;
      if (responsavel && idxResp >= 0 && String(r[idxResp] || "") !== responsavel) return false;
      if (oferta && idxOferta >= 0 && String(r[idxOferta] || "") !== oferta) return false;
      if ((inicioDe || inicioAte) && idxInicio >= 0) {
        const d = interpretarData_(r[idxInicio]);
        if (inicioDe && (!d || d < inicioDe)) return false;
        if (inicioAte && (!d || d > finalDoDia_(inicioAte))) return false;
      }
      return true;
    });
  }

  function aplicarOrdenacao_(linhas, colunas, ordenacao) {
    const chave = ordenacao && ordenacao.key ? String(ordenacao.key) : "";
    const dir = (ordenacao && ordenacao.dir ? String(ordenacao.dir) : "asc").toLowerCase() === "desc" ? -1 : 1;
    const idx = indiceDaColuna_(colunas, chave);
    if (idx < 0) return linhas || [];
    const copia = (linhas || []).slice();
    copia.sort((a, b) => {
      const av = a[idx];
      const bv = b[idx];
      const ad = interpretarData_(av);
      const bd = interpretarData_(bv);
      if (ad && bd) return (ad.getTime() - bd.getTime()) * dir;
      const an = interpretarNumero_(av);
      const bn = interpretarNumero_(bv);
      if (an !== null && bn !== null) return (an - bn) * dir;
      const as = (av === null || av === undefined) ? "" : String(av).toLowerCase();
      const bs = (bv === null || bv === undefined) ? "" : String(bv).toLowerCase();
      if (as < bs) return -1 * dir;
      if (as > bs) return 1 * dir;
      return 0;
    });
    return copia;
  }

  function linhaParaDados_(linha) {
    const o = {};
    CAMPOS_EDITAVEIS.forEach((k, idx) => { o[k] = linha[idx]; });
    return o;
  }

  function dadosParaLinha_(dados, metadados) {
    const linha = [];
    CAMPOS_EDITAVEIS.forEach((k) => { linha.push(dados ? dados[k] : null); });
    linha.push(metadados ? metadados.dataCadastro : null);
    linha.push(metadados ? metadados.emailUsuario : null);
    linha.push(metadados ? metadados.id : null);
    return linha;
  }

  function validarMinimo_(dados) {
    const obrigatorios = ["curso", "turma", "inicio", "fim", "sala", "oferta", "responsavel", "prioridade", "fimInscricoes", "status"];
    const faltando = obrigatorios.filter(k => !(dados && String(dados[k] || "").trim()));
    if (faltando.length) throw new Error("Campos obrigatórios ausentes: " + faltando.join(", "));
  }

  function buscarRegistrosComFiltros_(filtros, ordenacao, paginacao) {
    const { valores, temCabecalho } = RegistroRepo.lerDadosPlanilha();
    if (!valores.length) return { columns: RegistroRepo.obterColunasPadrao(), rows: [], total: 0, truncated: false };
    const colunas = temCabecalho ? normalizarColunasDoCabecalho_(valores[0]) : RegistroRepo.obterColunasPadrao();
    const linhas = valores.slice(temCabecalho ? 1 : 0);
    const filtradas = aplicarFiltros_(linhas, colunas, filtros || {});
    const ordenadas = aplicarOrdenacao_(filtradas, colunas, ordenacao || { key: "", dir: "asc" });
    const offset = Math.max(0, Number(paginacao && paginacao.offset) || 0);
    const limit = Math.min(100, Math.max(1, Number(paginacao && paginacao.limit) || 50));
    const paginadas = ordenadas.slice(offset, offset + limit);
    return { columns: colunas, rows: paginadas, total: ordenadas.length, truncated: ordenadas.length > paginadas.length };
  }

  function buscarRegistrosParaExportar_(filtros, ordenacao) {
    const { valores, temCabecalho } = RegistroRepo.lerDadosPlanilha();
    if (!valores.length) return { columns: RegistroRepo.obterColunasPadrao(), rows: [] };
    const colunas = temCabecalho ? normalizarColunasDoCabecalho_(valores[0]) : RegistroRepo.obterColunasPadrao();
    const linhas = valores.slice(temCabecalho ? 1 : 0);
    const filtradas = aplicarFiltros_(linhas, colunas, filtros || {});
    const ordenadas = aplicarOrdenacao_(filtradas, colunas, ordenacao || { key: "", dir: "asc" });
    return { columns: colunas, rows: ordenadas.slice(0, LIMITE_EXPORTACAO) };
  }

  function cadastrarRegistro_(dados) {
    validarMinimo_(dados);
    if (RegistroRepo.existeTurmaParaCurso(dados.curso, dados.turma, null))
      throw new Error("Já existe a turma " + "'" + (dados.turma || "").trim() + "'" + " para o curso " + "'" + (dados.curso || "").trim() + "'" + ".");
    const emailUsuario = Session.getActiveUser().getEmail();
    const dataAtual = new Date();
    const id = Utilities.getUuid();
    RegistroRepo.inserirLinha(dadosParaLinha_(dados, { dataCadastro: dataAtual, emailUsuario, id }));
    return "Registro salvo.";
  }

  function atualizarRegistroPorId_(id, dados) {
    if (!id) throw new Error("ID obrigatório para atualização.");
    validarMinimo_(dados);
    if (RegistroRepo.existeTurmaParaCurso(dados.curso, dados.turma, id))
      throw new Error("Já existe a turma " + "'" + (dados.turma || "").trim() + "'" + " para o curso " + "'" + (dados.curso || "").trim() + "'" + ".");
    const indice = RegistroRepo.buscarIndiceLinhaPorId(id);
    if (indice === -1) throw new Error("Registro não encontrado para atualização.");
    const linhaAtual = RegistroRepo.obterLinhaPorIndice(indice);
    const metadados = {
      dataCadastro: linhaAtual[IDX_DATA_CADASTRO] || null,
      emailUsuario: linhaAtual[IDX_EMAIL_USUARIO] || null,
      id: linhaAtual[IDX_ID] || id
    };
    RegistroRepo.atualizarLinha(indice, dadosParaLinha_(dados, metadados));
    return "Registro atualizado.";
  }

  function obterRegistroPorIdServico_(id) {
    if (!id) return null;
    const linha = RegistroRepo.buscarLinhaPorId(id);
    return linha ? linhaParaDados_(linha) : null;
  }

  function excluirRegistroPorId_(id) {
    if (!id) throw new Error("ID obrigatório para exclusão.");
    const indice = RegistroRepo.buscarIndiceLinhaPorId(id);
    if (indice === -1) throw new Error("Registro não encontrado para exclusão.");
    RegistroRepo.removerLinha(indice);
    return "Registro excluído.";
  }

  return {
    buscarRegistrosComFiltros: buscarRegistrosComFiltros_,
    buscarRegistrosParaExportar: buscarRegistrosParaExportar_,
    cadastrarRegistro: cadastrarRegistro_,
    atualizarRegistroPorId: atualizarRegistroPorId_,
    obterRegistroPorIdServico: obterRegistroPorIdServico_,
    excluirRegistroPorId: excluirRegistroPorId_
  };
})();

// ---- Wrappers globais (compatibilidade, Controller chama estes hoje) ----
function buscarRegistrosComFiltros(filtros, ordenacao, paginacao) { return RegistroService.buscarRegistrosComFiltros(filtros, ordenacao, paginacao); }
function buscarRegistrosParaExportar(filtros, ordenacao) { return RegistroService.buscarRegistrosParaExportar(filtros, ordenacao); }
function cadastrarRegistro(dados) { return RegistroService.cadastrarRegistro(dados); }
function atualizarRegistroPorId(id, dados) { return RegistroService.atualizarRegistroPorId(id, dados); }
function obterRegistroPorIdServico(id) { return RegistroService.obterRegistroPorIdServico(id); }
function excluirRegistroPorId(id) { return RegistroService.excluirRegistroPorId(id); }
