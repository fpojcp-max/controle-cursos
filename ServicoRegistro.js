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

  /** Delimita rótulos (turma, curso, etc.) em mensagens ao utilizador. */
  function citarRotuloMsg_(texto) {
    return "'" + String(texto != null ? texto : "").trim() + "'";
  }

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
      "responsavel": "Responsável", "oferta": "Oferta", "datacadastro": "Data Cadastro",
      "inicio": "Início", "id": "ID"
    };
    const aliased = aliases[k];
    if (aliased) {
      const proxK = String(aliased).trim().toLowerCase();
      // Evita recursão infinita quando o alias só muda caças (ex.: id ↔ ID): ambos normalizam para "id".
      if (proxK !== k) return indiceDaColuna_(colunas, aliased);
    }
    const nomeIdCfg = (typeof Configuracoes !== "undefined" && Configuracoes.NOME_COLUNA_ID)
      ? String(Configuracoes.NOME_COLUNA_ID).trim() : "";
    if (nomeIdCfg && k === "id") {
      const want = nomeIdCfg.toLowerCase();
      for (let i = 0; i < colunas.length; i++) {
        if (String(colunas[i].key).trim().toLowerCase() === want) return i;
      }
    }
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

  function compararCelulasOrdenacao_(av, bv, dir) {
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
  }

  function resolverSpecsOrdenacao_(colunas, ordenacao) {
    const specs = [];
    if (ordenacao && Array.isArray(ordenacao.keys) && ordenacao.keys.length) {
      ordenacao.keys.forEach((item) => {
        const chave = item && item.key ? String(item.key) : "";
        const dir = (item && item.dir ? String(item.dir) : "asc").toLowerCase() === "desc" ? -1 : 1;
        const idx = indiceDaColuna_(colunas, chave);
        if (idx >= 0) specs.push({ idx, dir });
      });
      return specs;
    }
    const chave = ordenacao && ordenacao.key ? String(ordenacao.key) : "";
    const dir = (ordenacao && ordenacao.dir ? String(ordenacao.dir) : "asc").toLowerCase() === "desc" ? -1 : 1;
    const idx = indiceDaColuna_(colunas, chave);
    if (idx >= 0) specs.push({ idx: idx, dir: dir });
    return specs;
  }

  function aplicarOrdenacao_(linhas, colunas, ordenacao) {
    const specs = resolverSpecsOrdenacao_(colunas, ordenacao);
    if (!specs.length) return linhas || [];
    const copia = (linhas || []).slice();
    copia.sort((a, b) => {
      for (let i = 0; i < specs.length; i++) {
        const s = specs[i];
        const c = compararCelulasOrdenacao_(a[s.idx], b[s.idx], s.dir);
        if (c !== 0) return c;
      }
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

  function obterColunasLinhasFiltradasOrdenadas_(filtros, ordenacao) {
    const { valores, temCabecalho } = RegistroRepo.lerDadosPlanilha();
    if (!valores.length) {
      return { colunas: RegistroRepo.obterColunasPadrao(), ordenadas: [] };
    }
    const colunas = temCabecalho ? normalizarColunasDoCabecalho_(valores[0]) : RegistroRepo.obterColunasPadrao();
    const linhas = valores.slice(temCabecalho ? 1 : 0);
    const filtradas = aplicarFiltros_(linhas, colunas, filtros || {});
    const ordenadas = aplicarOrdenacao_(filtradas, colunas, ordenacao || { key: "", dir: "asc" });
    return { colunas: colunas, ordenadas: ordenadas };
  }

  function buscarRegistrosComFiltros_(filtros, ordenacao, paginacao) {
    const { colunas, ordenadas } = obterColunasLinhasFiltradasOrdenadas_(filtros, ordenacao);
    const offset = Math.max(0, Number(paginacao && paginacao.offset) || 0);
    const limit = Math.min(100, Math.max(1, Number(paginacao && paginacao.limit) || 50));
    const paginadas = ordenadas.slice(offset, offset + limit);
    return { columns: colunas, rows: paginadas, total: ordenadas.length, truncated: ordenadas.length > paginadas.length };
  }

  function buscarRegistrosParaExportar_(filtros, ordenacao) {
    const { colunas, ordenadas } = obterColunasLinhasFiltradasOrdenadas_(filtros, ordenacao);
    return { columns: colunas, rows: ordenadas.slice(0, LIMITE_EXPORTACAO) };
  }

  /**
   * Pesquisa paginada + lista completa de IDs (filtro + ordenação) para "selecionar tudo" na exclusão em lote.
   */
  function buscarRegistrosExcluirTurmaTela_(filtros, ordenacao, paginacao) {
    const { colunas, ordenadas } = obterColunasLinhasFiltradasOrdenadas_(filtros, ordenacao);
    const offset = Math.max(0, Number(paginacao && paginacao.offset) || 0);
    const limit = Math.min(100, Math.max(1, Number(paginacao && paginacao.limit) || 50));
    const paginadas = ordenadas.slice(offset, offset + limit);
    const idIdx = indiceDaColuna_(colunas, "ID");
    const idsFiltradosOrdenados = [];
    if (idIdx >= 0) {
      for (let i = 0; i < ordenadas.length; i++) {
        const idVal = String(ordenadas[i][idIdx] || "").trim();
        if (idVal) idsFiltradosOrdenados.push(idVal);
      }
    }
    return {
      columns: colunas,
      rows: paginadas,
      total: ordenadas.length,
      truncated: ordenadas.length > paginadas.length,
      idsFiltradosOrdenados: idsFiltradosOrdenados,
      idColumnIndex: idIdx
    };
  }

  /**
   * Turma >> Editar: filtro obrigatório curso + turma; retorno completo (sem paginação) + índice da coluna ID.
   */
  function pesquisarRegistrosTurmaEditarTela_(curso, turma) {
    const filtros = {
      curso: String(curso || "").trim(),
      turma: String(turma || "").trim(),
      status: "",
      sala: "",
      oferta: "",
      responsavel: "",
      inicioDe: "",
      inicioAte: ""
    };
    const ordenacao = {
      keys: [
        { key: "Inicio", dir: "asc" },
        { key: "Curso", dir: "asc" },
        { key: "Turma", dir: "asc" }
      ]
    };
    const { colunas, ordenadas } = obterColunasLinhasFiltradasOrdenadas_(filtros, ordenacao);
    const idIdx = indiceDaColuna_(colunas, "ID");
    return {
      columns: colunas,
      rows: ordenadas,
      total: ordenadas.length,
      idColumnIndex: idIdx
    };
  }

  function obterLinhasRegistroTurmaPorIdsNaOrdem_(ids) {
    const lista = (ids || []).map((x) => String(x || "").trim()).filter(Boolean);
    if (!lista.length) {
      const { valores, temCabecalho } = RegistroRepo.lerDadosPlanilha();
      if (!valores.length) {
        const colunas0 = RegistroRepo.obterColunasPadrao();
        return { columns: colunas0, rows: [], idColumnIndex: indiceDaColuna_(colunas0, "ID") };
      }
      const colunas = temCabecalho ? normalizarColunasDoCabecalho_(valores[0]) : RegistroRepo.obterColunasPadrao();
      return { columns: colunas, rows: [], idColumnIndex: indiceDaColuna_(colunas, "ID") };
    }
    const { valores, temCabecalho } = RegistroRepo.lerDadosPlanilha();
    if (!valores.length) {
      const colunas0 = RegistroRepo.obterColunasPadrao();
      return { columns: colunas0, rows: [], idColumnIndex: indiceDaColuna_(colunas0, "ID") };
    }
    const colunas = temCabecalho ? normalizarColunasDoCabecalho_(valores[0]) : RegistroRepo.obterColunasPadrao();
    const linhas = valores.slice(temCabecalho ? 1 : 0);
    const idIdx = indiceDaColuna_(colunas, "ID");
    if (idIdx < 0) return { columns: colunas, rows: [], idColumnIndex: -1 };
    const mapa = {};
    for (let i = 0; i < linhas.length; i++) {
      const idVal = String(linhas[i][idIdx] || "").trim();
      if (idVal) mapa[idVal] = linhas[i];
    }
    const rows = [];
    for (let j = 0; j < lista.length; j++) {
      const row = mapa[lista[j]];
      if (row) rows.push(row);
    }
    return { columns: colunas, rows: rows, idColumnIndex: idIdx };
  }

  function formatarLinhaTurmaExcluidaMsg_(curso, turma) {
    return "Turma " + citarRotuloMsg_(turma) + " do curso " + citarRotuloMsg_(curso);
  }

  function excluirRegistroPorIdComDetalhe_(id) {
    if (!id) throw new Error("ID obrigatório para exclusão.");
    const indice = RegistroRepo.buscarIndiceLinhaPorId(id);
    if (indice === -1) throw new Error("Registro não encontrado para exclusão.");
    const linha = RegistroRepo.obterLinhaPorIndice(indice);
    const dados = linhaParaDados_(linha);
    const turma = String(dados.turma || "").trim();
    const curso = String(dados.curso || "").trim();
    AgendamentoService.excluirTodosAgendamentosPorIdTurmaAoExcluirRegistro(id);
    RegistroRepo.removerLinha(indice);
    return { curso: curso, turma: turma };
  }

  function excluirRegistrosTurmaLote_(ids) {
    const excluidos = [];
    const falhas = [];
    const vistos = {};
    const entrada = (ids || []).map((x) => String(x || "").trim()).filter(Boolean);
    for (let i = 0; i < entrada.length; i++) {
      const id = entrada[i];
      if (vistos[id]) continue;
      vistos[id] = true;
      try {
        const d = excluirRegistroPorIdComDetalhe_(id);
        excluidos.push({ id: id, curso: d.curso, turma: d.turma });
      } catch (e) {
        falhas.push({ id: id, mensagem: (e && e.message) ? e.message : String(e) });
      }
    }
    const linhasEx = excluidos.map((x) => formatarLinhaTurmaExcluidaMsg_(x.curso, x.turma));
    let tipo = "nenhuma";
    let mensagemTopo = "";
    if (excluidos.length && !falhas.length) {
      tipo = "total";
      mensagemTopo =
        excluidos.length === 1
          ? linhasEx[0] + " excluída com sucesso."
          : "As seguintes turmas foram excluídas:\n" + linhasEx.join("\n");
    } else if (excluidos.length && falhas.length) {
      tipo = "parcial";
      mensagemTopo =
        "Apenas as seguintes turmas foram excluídas:\n" +
        linhasEx.join("\n") +
        "\n\nTente excluir as demais turmas restantes na tabela de resultados mais tarde.";
    } else if (!excluidos.length && falhas.length) {
      tipo = "falha_total";
      mensagemTopo = "Não foi possível excluir as turmas selecionadas.";
    }
    return { excluidos: excluidos, falhas: falhas, tipo: tipo, mensagemTopo: mensagemTopo };
  }

  function cadastrarRegistro_(dados) {
    validarMinimo_(dados);
    if (RegistroRepo.existeTurmaParaCurso(dados.curso, dados.turma, null))
      throw new Error(
        "Já existe a turma " + citarRotuloMsg_(dados.turma) + " para o curso " + citarRotuloMsg_(dados.curso) + "."
      );
    const emailUsuario = Session.getActiveUser().getEmail();
    const dataAtual = new Date();
    const id = Utilities.getUuid();
    RegistroRepo.inserirLinha(dadosParaLinha_(dados, { dataCadastro: dataAtual, emailUsuario, id }));
    const turma = String(dados.turma || "").trim();
    const curso = String(dados.curso || "").trim();
    return "Turma " + citarRotuloMsg_(turma) + " do curso " + citarRotuloMsg_(curso) + " incluída com sucesso.";
  }

  function atualizarRegistroPorId_(id, dados) {
    if (!id) throw new Error("ID obrigatório para atualização.");
    validarMinimo_(dados);
    if (RegistroRepo.existeTurmaParaCurso(dados.curso, dados.turma, id))
      throw new Error(
        "Já existe a turma " + citarRotuloMsg_(dados.turma) + " para o curso " + citarRotuloMsg_(dados.curso) + "."
      );
    const indice = RegistroRepo.buscarIndiceLinhaPorId(id);
    if (indice === -1) throw new Error("Registro não encontrado para atualização.");
    const linhaAtual = RegistroRepo.obterLinhaPorIndice(indice);
    const metadados = {
      dataCadastro: linhaAtual[IDX_DATA_CADASTRO] || null,
      emailUsuario: linhaAtual[IDX_EMAIL_USUARIO] || null,
      id: linhaAtual[IDX_ID] || id
    };
    RegistroRepo.atualizarLinha(indice, dadosParaLinha_(dados, metadados));
    const turma = String(dados.turma || "").trim();
    const curso = String(dados.curso || "").trim();
    return (
      "O cadastro da turma " +
      citarRotuloMsg_(turma) +
      " do curso " +
      citarRotuloMsg_(curso) +
      " foi atualizado."
    );
  }

  function obterRegistroPorIdServico_(id) {
    if (!id) return null;
    const linha = RegistroRepo.buscarLinhaPorId(id);
    return linha ? linhaParaDados_(linha) : null;
  }

  function excluirRegistroPorId_(id) {
    const d = excluirRegistroPorIdComDetalhe_(id);
    return formatarLinhaTurmaExcluidaMsg_(d.curso, d.turma) + " excluída com sucesso.";
  }

  return {
    buscarRegistrosComFiltros: buscarRegistrosComFiltros_,
    buscarRegistrosParaExportar: buscarRegistrosParaExportar_,
    buscarRegistrosExcluirTurmaTela: buscarRegistrosExcluirTurmaTela_,
    pesquisarRegistrosTurmaEditarTela: pesquisarRegistrosTurmaEditarTela_,
    obterLinhasRegistroTurmaPorIdsNaOrdem: obterLinhasRegistroTurmaPorIdsNaOrdem_,
    excluirRegistrosTurmaLote: excluirRegistrosTurmaLote_,
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
function buscarRegistrosExcluirTurmaTela(filtros, ordenacao, paginacao) {
  return RegistroService.buscarRegistrosExcluirTurmaTela(filtros, ordenacao, paginacao);
}
function obterLinhasRegistroTurmaPorIdsNaOrdem(ids) {
  return RegistroService.obterLinhasRegistroTurmaPorIdsNaOrdem(ids);
}
