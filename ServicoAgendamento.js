/**
 * Camada Service – Agendamento (planilha interna, sem Google Calendar).
 */

const AgendamentoService = (() => {
  const HORA_REGEX = /^([01]?\d|2[0-3]):([0-5]\d)$/;
  const DIAS_JS = { seg: 1, ter: 2, qua: 3, qui: 4, sex: 5 };

  const MSG_EXCLUSAO_TURMA_AGENDAMENTOS_FALHOU =
    "Não foi possível excluir a turma. Tente novamente mais tarde.";
  const MSG_EXCLUSAO_AGENDAMENTOS_LOTE_FALHOU =
    "Não foi possível concluir a exclusão. Tente novamente mais tarde.";

  function citarRotuloMsg_(texto) {
    return "'" + String(texto != null ? texto : "").trim() + "'";
  }

  function tz_() {
    return Configuracoes.TIMEZONE_AGENDAMENTO || "America/Sao_Paulo";
  }

  function dataCivilHojeYmd_() {
    return Utilities.formatDate(new Date(), tz_(), "yyyy-MM-dd");
  }

  function validarHora_(h, nome) {
    const s = String(h || "").trim();
    if (!HORA_REGEX.test(s)) throw new Error("Hora inválida em " + nome + ".");
    return s;
  }

  function parseYmd_(s) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || "").trim());
    if (!m) throw new Error("Data inválida.");
    const y = parseInt(m[1], 10);
    const mo = parseInt(m[2], 10) - 1;
    const d = parseInt(m[3], 10);
    const dt = new Date(y, mo, d);
    if (isNaN(dt.getTime()) || dt.getFullYear() !== y || dt.getMonth() !== mo || dt.getDate() !== d) {
      throw new Error("Data inválida.");
    }
    return dt;
  }

  function formatarYmdParaMsgBr_(ymd) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(ymd || "").trim());
    if (!m) return String(ymd || "");
    return m[3] + "/" + m[2] + "/" + m[1];
  }

  function enumerarDatasInclusive_(inicioStr, fimStr) {
    const ini = parseYmd_(inicioStr);
    const fim = parseYmd_(fimStr);
    if (fim.getTime() < ini.getTime()) throw new Error("Data fim deve ser igual ou posterior à data início.");
    const out = [];
    const cur = new Date(ini.getTime());
    while (cur.getTime() <= fim.getTime()) {
      out.push(
        Utilities.formatDate(cur, tz_(), "yyyy-MM-dd")
      );
      cur.setDate(cur.getDate() + 1);
    }
    return out;
  }

  function parseDataHoraLocalMs_(dataYmd, hhmm) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(dataYmd || "").trim());
    const t = /^(\d{2}):(\d{2})$/.exec(String(hhmm || "").trim());
    if (!m || !t) return NaN;
    const y = parseInt(m[1], 10);
    const mo = parseInt(m[2], 10) - 1;
    const d = parseInt(m[3], 10);
    const hh = parseInt(t[1], 10);
    const mm = parseInt(t[2], 10);
    return new Date(y, mo, d, hh, mm, 0, 0).getTime();
  }

  function intervalosSobrepoem_(aStartMs, aEndMs, bStartMs, bEndMs) {
    return aStartMs < bEndMs && bStartMs < aEndMs;
  }

  function montarTitulo_(turma, curso) {
    return String(turma || "").trim() + " - " + String(curso || "").trim();
  }

  function formatarCriadoEm_() {
    return Utilities.formatDate(new Date(), tz_(), "yyyy-MM-dd HH:mm");
  }

  function normalizarCelulaVigenciaParaYmd_(valor, nomeCampo) {
    if (valor === null || valor === undefined || String(valor).trim() === "") {
      throw new Error("Data inválida em " + nomeCampo + ".");
    }
    if (Object.prototype.toString.call(valor) === "[object Date]" && !isNaN(valor.getTime())) {
      return Utilities.formatDate(valor, tz_(), "yyyy-MM-dd");
    }
    const s0 = String(valor).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s0)) {
      parseYmd_(s0);
      return s0;
    }
    const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s0);
    if (m) {
      const ymd =
        m[3] + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[1]).slice(-2);
      parseYmd_(ymd);
      return ymd;
    }
    throw new Error("Data inválida em " + nomeCampo + ".");
  }

  function obterVigenciaTurmaOuErro_(curso, turma) {
    const raw = RegistroRepo.buscarVigenciaInicioFimPorCursoTurma(curso, turma);
    if (!raw) throw new Error("Registro da turma não encontrado.");
    return {
      inicioYmd: normalizarCelulaVigenciaParaYmd_(raw.inicioVal, "Início"),
      fimYmd: normalizarCelulaVigenciaParaYmd_(raw.fimVal, "Fim")
    };
  }

  function expandirOcorrencias_(payload) {
    const tipo = String(payload.tipo || "").toLowerCase();
    const exdate = String(payload.exdate || "").trim();
    const rdate = String(payload.rdate || "").trim();

    if (tipo === "simples") {
      const d = String(payload.data || "").trim();
      parseYmd_(d);
      const dtSimples = parseYmd_(d);
      const dowS = dtSimples.getDay();
      if (dowS === 0 || dowS === 6) {
        throw new Error("Não são permitidos agendamentos para sábados e domingos.");
      }
      if (exdate || rdate) {
        throw new Error("Datas de exceção ou exceção positiva só se aplicam a evento recorrente.");
      }
      return [d];
    }

    if (tipo !== "recorrente") throw new Error("Tipo de evento inválido.");

    const dias = payload.diasSemana || {};
    const ativos = Object.keys(DIAS_JS).filter(
      (k) => dias[k] === true || dias[k] === "true" || dias[k] === 1
    );
    if (!ativos.length) throw new Error("Selecione pelo menos um dia da semana (segunda a sexta).");

    const permitidos = {};
    ativos.forEach((k) => {
      permitidos[DIAS_JS[k]] = true;
    });

    const todas = enumerarDatasInclusive_(payload.dataInicio, payload.dataFim);
    const setDatas = [];
    const pushUnique = (ymd) => {
      if (setDatas.indexOf(ymd) === -1) setDatas.push(ymd);
    };

    todas.forEach((ymd) => {
      const dt = parseYmd_(ymd);
      if (!permitidos[dt.getDay()]) return;
      if (exdate && ymd === exdate) return;
      pushUnique(ymd);
    });

    if (rdate) {
      const dtR = parseYmd_(rdate);
      if (dtR.getDay() === 0 || dtR.getDay() === 6) {
        throw new Error("Não são permitidos agendamentos para sábados e domingos.");
      }
      pushUnique(rdate);
    }

    setDatas.sort();
    return setDatas;
  }

  function validarConjuntoDatasAgendamento_(payload, datasYmd, curso, turma) {
    const hoje = dataCivilHojeYmd_();
    const vig = obterVigenciaTurmaOuErro_(curso, turma);
    for (let i = 0; i < datasYmd.length; i++) {
      const ymd = datasYmd[i];
      if (ymd < hoje) throw new Error("Datas passadas não são permitidas.");
      if (ymd < vig.inicioYmd || ymd > vig.fimYmd) {
        throw new Error(
          "A data " +
            formatarYmdParaMsgBr_(ymd) +
            " está fora do período de vigência da turma " +
            citarRotuloMsg_(turma) +
            " do curso " +
            citarRotuloMsg_(curso) +
            "."
        );
      }
    }
  }

  function validarSalaNome_(salaNome) {
    const nome = String(salaNome || "").trim();
    if (!nome) return "";
    const validos = montarConjuntoRotulosSalaValidos();
    if (!validos[nome]) throw new Error("Sala não reconhecida: " + nome);
    return nome;
  }

  /**
   * Conflito: mesma sala + mesma data + sobreposição de horário (todas as linhas da aba).
   * @param {string} salaNome
   * @param {{ ymd?: string, startMs: number, endMs: number }[]} periodos
   * @param {number} [excludeSheetRow]
   */
  function checarSalaLivrePeriodos_(salaNome, periodos, excludeSheetRow) {
    const sala = String(salaNome || "").trim();
    if (!sala || !periodos.length) return;

    const C = AgendamentoRepo.COL_AG;
    const todas = AgendamentoRepo.listarTodasLinhasAgendamento();
    const excl = parseInt(excludeSheetRow, 10);

    for (let p = 0; p < periodos.length; p++) {
      const ev = periodos[p];
      const ymdEv = ev.ymd || "";
      for (let i = 0; i < todas.length; i++) {
        const item = todas[i];
        if (!isNaN(excl) && excl >= 2 && item.sheetRow === excl) continue;
        const cells = item.cells;
        const salaCell = String(cells[C.NOME_SALA] || "").trim();
        if (salaCell !== sala) continue;
        const dataCell = String(cells[C.DATA] || "").trim();
        if (ymdEv && dataCell !== ymdEv) continue;
        const hi = String(cells[C.HORA_INI] || "").trim();
        const hf = String(cells[C.HORA_FIM] || "").trim();
        const bs = parseDataHoraLocalMs_(dataCell, hi);
        const be = parseDataHoraLocalMs_(dataCell, hf);
        if (isNaN(bs) || isNaN(be)) continue;
        if (intervalosSobrepoem_(ev.startMs, ev.endMs, bs, be)) {
          throw new Error(
            "A " + citarRotuloMsg_(sala) + " está ocupada. Tente outra sala ou outro horário."
          );
        }
      }
    }
  }

  function montarMensagemSucessoSimples_(turma, curso, horaInicio, horaFim, salaNome, dataYmd) {
    const evento = citarRotuloMsg_(turma) + " - " + citarRotuloMsg_(curso);
    const sala = String(salaNome || "").trim() || "—";
    return (
      "O seguinte agendamento foi gerado:\n" +
      "Evento: " + evento + "\n" +
      "Horário: " + horaInicio + " - " + horaFim + "\n" +
      "Sala: " + sala + "\n" +
      "Data: " + formatarYmdParaMsgBr_(dataYmd)
    );
  }

  const NOME_DIA_SEMANA_PT_ = {
    0: "domingo", 1: "segunda-feira", 2: "terça-feira", 3: "quarta-feira",
    4: "quinta-feira", 5: "sexta-feira", 6: "sábado"
  };
  const ORDEM_DIA_SEMANA_ = [1, 2, 3, 4, 5, 6, 0];

  function montarMensagemSucessoMultiplos_(turma, curso, horaInicio, horaFim, salaNome, periodos) {
    const evento = citarRotuloMsg_(turma) + " - " + citarRotuloMsg_(curso);
    const sala = String(salaNome || "").trim() || "—";
    const gruposPorDow = {};
    for (let i = 0; i < periodos.length; i++) {
      const ymd = periodos[i].ymd;
      const dow = parseYmd_(ymd).getDay();
      if (!gruposPorDow[dow]) gruposPorDow[dow] = [];
      gruposPorDow[dow].push(ymd);
    }
    let blocosDia = "";
    for (let k = 0; k < ORDEM_DIA_SEMANA_.length; k++) {
      const dow = ORDEM_DIA_SEMANA_[k];
      const lista = gruposPorDow[dow];
      if (!lista || !lista.length) continue;
      lista.sort();
      blocosDia += "Dia: " + NOME_DIA_SEMANA_PT_[dow] + "\n";
      blocosDia += "Data: " + lista.map(formatarYmdParaMsgBr_).join(", ") + "\n";
    }
    return (
      "Os seguintes agendamentos foram gerados:\n" +
      "Evento: " + evento + "\n" +
      "Horário: " + horaInicio + " - " + horaFim + "\n" +
      "Sala: " + sala + "\n" +
      blocosDia.replace(/\s+$/, "")
    );
  }

  function criarEventos_(request) {
    throw {
      code: "ENDPOINT_DISABLED",
      message: "API de agendamento descontinuada. Utilize a interface Web.",
      details: []
    };
  }

  function obterDadosIncluir_() {
    const salas = listarRotulosCatalogoRecursosSala().map((nome) => ({ nome: nome }));
    return {
      cursos: RegistroRepo.listarCursosDistintos(),
      salas: salas,
      timezone: tz_(),
      hojeYmd: dataCivilHojeYmd_()
    };
  }

  function listarTurmasPorCursoIncluir_(curso) {
    return RegistroRepo.listarTurmasDistintasPorCurso(curso);
  }

  function criarAgendamentos_(payload) {
    if (!payload || typeof payload !== "object") throw new Error("Dados inválidos.");

    const curso = String(payload.curso || "").trim();
    const turma = String(payload.turma || "").trim();
    const turmaIdCliente = String(payload.turmaId || "").trim();
    if (!curso || !turma) throw new Error("Selecione curso e turma.");

    const idTurma = RegistroRepo.buscarIdPorCursoTurma(curso, turma);
    if (!idTurma) {
      throw new Error(
        "Não existe registro na planilha de turmas para o curso " +
          citarRotuloMsg_(curso) + " e a turma " + citarRotuloMsg_(turma) + "."
      );
    }
    if (turmaIdCliente && turmaIdCliente !== idTurma) {
      throw new Error("Dados da turma não conferem com o registro selecionado. Recarregue a tela e tente novamente.");
    }

    PermissaoService.garantirPodeAgendarTurmaPorId(idTurma);

    const horaInicio = validarHora_(payload.horaInicio, "Hora início");
    const horaFim = validarHora_(payload.horaFim, "Hora fim");
    const minIni = parseInt(horaInicio.split(":")[0], 10) * 60 + parseInt(horaInicio.split(":")[1], 10);
    const minFim = parseInt(horaFim.split(":")[0], 10) * 60 + parseInt(horaFim.split(":")[1], 10);
    if (minFim <= minIni) throw new Error("Hora fim deve ser posterior à hora início.");

    const salaNome = validarSalaNome_(payload.salaNome || "");

    const datas = expandirOcorrencias_(payload);
    if (!datas.length) throw new Error("Nenhuma ocorrência no período e dias selecionados.");
    validarConjuntoDatasAgendamento_(payload, datas, curso, turma);

    const periodos = datas.map((ymd) => ({
      ymd: ymd,
      startMs: parseDataHoraLocalMs_(ymd, horaInicio),
      endMs: parseDataHoraLocalMs_(ymd, horaFim)
    }));

    if (salaNome) checarSalaLivrePeriodos_(salaNome, periodos, null);

    const criadoPor = Session.getActiveUser().getEmail() || "";
    const criadoEm = formatarCriadoEm_();
    const C = AgendamentoRepo.COL_AG;
    const linhas = periodos.map((p) => {
      const row = new Array(10).fill("");
      row[C.TURMA] = turma;
      row[C.CURSO] = curso;
      row[C.DATA] = p.ymd;
      row[C.NOME_SALA] = salaNome || "";
      row[C.HORA_INI] = horaInicio;
      row[C.HORA_FIM] = horaFim;
      row[C.CRIADO_EM] = criadoEm;
      row[C.CRIADO_POR] = criadoPor;
      row[C.ID_AGENDAMENTO] = Utilities.getUuid();
      row[C.ID_REGISTRO_TURMA] = idTurma;
      return row;
    });

    AgendamentoRepo.appendLinhas(linhas);

    const mensagemSucesso =
      periodos.length === 1
        ? montarMensagemSucessoSimples_(turma, curso, horaInicio, horaFim, salaNome, periodos[0].ymd)
        : montarMensagemSucessoMultiplos_(turma, curso, horaInicio, horaFim, salaNome, periodos);

    return { ocorrencias: periodos.length, mensagem: mensagemSucesso };
  }

  function limiteExclusaoAgendamentosLote_() {
    const n = Number(Configuracoes.LIMITE_EXCLUSAO_AGENDAMENTOS_LOTE);
    return n > 0 ? n : 100;
  }

  function montarMensagemExclusaoAgendamentos_(cellsList) {
    const C = AgendamentoRepo.COL_AG;
    if (!cellsList || !cellsList.length) return "Agendamentos excluídos.";
    const turma = cellsList[0][C.TURMA] || "";
    const curso = cellsList[0][C.CURSO] || "";
    const grupos = {};
    for (let i = 0; i < cellsList.length; i++) {
      const row = cellsList[i];
      const key = (row[C.HORA_INI] || "") + "\t" + (row[C.HORA_FIM] || "") + "\t" + (row[C.NOME_SALA] || "");
      if (!grupos[key]) grupos[key] = [];
      grupos[key].push(row[C.DATA] || "");
    }
    let out = "Os seguintes agendamentos foram excluídos:\n";
    out += "Evento: " + citarRotuloMsg_(turma) + " - " + citarRotuloMsg_(curso) + "\n";
    Object.keys(grupos).sort().forEach(function (key) {
      const parts = key.split("\t");
      out += "Horário: " + (parts[0] || "") + " - " + (parts[1] || "") + "\n";
      out += "Sala: " + (parts[2] || "—") + "\n";
      out += grupos[key].sort().map(formatarYmdParaMsgBr_).join(", ") + "\n";
    });
    return out;
  }

  function pesquisarAgendamentosExcluir_(curso, turma, offset, limit, sortCol, sortDir) {
    const c = String(curso || "").trim();
    const t = String(turma || "").trim();
    if (!c || !t) throw new Error("Selecione curso e turma.");
    const idTurma = RegistroRepo.buscarIdPorCursoTurma(c, t);
    if (!idTurma) {
      throw new Error(
        "Não existe registro na planilha de turmas para o curso " + citarRotuloMsg_(c) +
          " e a turma " + citarRotuloMsg_(t) + "."
      );
    }
    let sc = -1;
    if (sortCol !== undefined && sortCol !== null && String(sortCol).trim() !== "") {
      const n = parseInt(sortCol, 10);
      if (!isNaN(n)) sc = n;
    }
    const r = AgendamentoRepo.listarAgendamentosPaginadoPorIdTurma(
      idTurma, offset, limit, sc, String(sortDir == null ? "asc" : sortDir).toLowerCase() !== "desc"
    );
    return {
      success: true,
      idTurma: idTurma,
      curso: c,
      turma: t,
      cabecalho: r.cabecalho,
      total: r.total,
      allLinhas: r.allLinhas || [],
      itens: r.itens.map((item) => ({
        eventId: item.eventId,
        sheetRow: item.sheetRow,
        cells: item.cells
      }))
    };
  }

  const LIMITE_EXPORT_AGENDAMENTOS = 10000;

  function obterAgendamentosConsultaParaExportar_(curso, turma, sortCol, sortDir) {
    const c = String(curso || "").trim();
    const t = String(turma || "").trim();
    if (!c || !t) throw new Error("Selecione curso e turma.");
    const idTurma = RegistroRepo.buscarIdPorCursoTurma(c, t);
    if (!idTurma) {
      throw new Error(
        "Não existe registro na planilha de turmas para o curso " + citarRotuloMsg_(c) +
          " e a turma " + citarRotuloMsg_(t) + "."
      );
    }
    let sc = -1;
    if (sortCol !== undefined && sortCol !== null && String(sortCol).trim() !== "") {
      const n = parseInt(sortCol, 10);
      if (!isNaN(n)) sc = n;
    }
    const r = AgendamentoRepo.listarAgendamentosPaginadoPorIdTurma(
      idTurma, 0, LIMITE_EXPORT_AGENDAMENTOS, sc,
      String(sortDir == null ? "asc" : sortDir).toLowerCase() !== "desc"
    );
    const cab = r.cabecalho || [];
    return {
      columns: cab.map((h, idx) => ({ key: "c" + idx, label: String(h != null ? h : "") })),
      rows: (r.itens || []).map((item) =>
        (item.cells || []).map((cell) => (cell === null || cell === undefined ? "" : String(cell)))
      )
    };
  }

  function obterTodosEventIdsExcluir_(curso, turma) {
    const c = String(curso || "").trim();
    const t = String(turma || "").trim();
    if (!c || !t) throw new Error("Selecione curso e turma.");
    const idTurma = RegistroRepo.buscarIdPorCursoTurma(c, t);
    if (!idTurma) {
      throw new Error(
        "Não existe registro na planilha de turmas para o curso " + citarRotuloMsg_(c) +
          " e a turma " + citarRotuloMsg_(t) + "."
      );
    }
    return { success: true, eventIds: AgendamentoRepo.listarTodosEventIdsPorIdTurma(idTurma) };
  }

  function excluirTodosAgendamentosPorIdTurmaAoExcluirRegistro_(idTurma) {
    const idNorm = String(idTurma || "").trim();
    if (!idNorm) return;
    let linhas;
    try {
      linhas = AgendamentoRepo.listarLinhasAgendamentoPorIdTurmaCompleto(idNorm);
    } catch (e) {
      throw new Error(MSG_EXCLUSAO_TURMA_AGENDAMENTOS_FALHOU);
    }
    if (!linhas.length) return;
    const rowsDesc = [];
    const seenRow = {};
    for (let r = 0; r < linhas.length; r++) {
      const n = linhas[r].sheetRow;
      if (seenRow[n]) continue;
      seenRow[n] = 1;
      rowsDesc.push(n);
    }
    rowsDesc.sort((a, b) => b - a);
    try {
      AgendamentoRepo.excluirLinhasPorNumeros(rowsDesc);
    } catch (sheetErr) {
      throw new Error(MSG_EXCLUSAO_TURMA_AGENDAMENTOS_FALHOU);
    }
  }

  function excluirAgendamentosLote_(curso, turma, payload) {
    const LIM = limiteExclusaoAgendamentosLote_();
    const c = String(curso || "").trim();
    const t = String(turma || "").trim();
    if (!c || !t) throw new Error("Selecione curso e turma.");

    const idTurma = RegistroRepo.buscarIdPorCursoTurma(c, t);
    if (!idTurma) {
      throw new Error(
        "Não existe registro na planilha de turmas para o curso " + citarRotuloMsg_(c) +
          " e a turma " + citarRotuloMsg_(t) + "."
      );
    }

    PermissaoService.garantirPodeAgendarTurmaPorId(idTurma);

    const todas = AgendamentoRepo.listarLinhasAgendamentoPorIdTurma(idTurma);
    const byRow = {};
    for (let j = 0; j < todas.length; j++) byRow[todas[j].sheetRow] = todas[j];

    let selecionadas = [];
    const rowsIn = Array.isArray(payload && payload.sheetRows ? payload.sheetRows : []) ? payload.sheetRows : [];

    if (rowsIn.length) {
      const nums = [];
      const seenR = {};
      for (let i = 0; i < rowsIn.length; i++) {
        const n = parseInt(rowsIn[i], 10);
        if (isNaN(n) || n < 2 || seenR[n]) continue;
        seenR[n] = 1;
        nums.push(n);
      }
      if (!nums.length) throw new Error("Selecione ao menos um agendamento.");
      if (nums.length > LIM) throw new Error("Selecione no máximo " + LIM + " agendamentos para excluir por vez.");
      for (let k = 0; k < nums.length; k++) {
        const m = byRow[nums[k]];
        if (!m) throw new Error("Agendamento inválido ou não pertence à turma selecionada.");
        selecionadas.push(m);
      }
    } else {
      const idsIn = Array.isArray(payload && payload.eventIds ? payload.eventIds : [])
        ? payload.eventIds.map((x) => String(x || "").trim()).filter(Boolean) : [];
      const uniq = [];
      const seen = {};
      idsIn.forEach((id) => {
        if (!seen[id]) { seen[id] = 1; uniq.push(id); }
      });
      if (!uniq.length) throw new Error("Selecione ao menos um agendamento.");
      if (uniq.length > LIM) throw new Error("Selecione no máximo " + LIM + " agendamentos para excluir por vez.");
      const mapa = {};
      todas.forEach((m) => { mapa[m.eventId] = m; });
      uniq.forEach((ev) => {
        if (!mapa[ev]) throw new Error("Agendamento inválido ou não pertence à turma selecionada.");
        selecionadas.push(mapa[ev]);
      });
    }

    const rowsDesc = selecionadas.map((m) => m.sheetRow).sort((a, b) => b - a);
    try {
      AgendamentoRepo.excluirLinhasPorNumeros(rowsDesc);
    } catch (sheetErr) {
      throw new Error(MSG_EXCLUSAO_AGENDAMENTOS_LOTE_FALHOU);
    }

    return {
      success: true,
      mensagem: montarMensagemExclusaoAgendamentos_(selecionadas.map((s) => s.cells))
    };
  }

  function validarNovaDataEdicaoAgendamento_(ymd, curso, turma) {
    if (ymd < dataCivilHojeYmd_()) throw new Error("Datas passadas não são permitidas");
    const dt = parseYmd_(ymd);
    if (dt.getDay() === 0 || dt.getDay() === 6) {
      throw new Error("Não são permitidos agendamentos para sábados e domingos.");
    }
    const vig = obterVigenciaTurmaOuErro_(curso, turma);
    if (ymd < vig.inicioYmd || ymd > vig.fimYmd) {
      throw new Error("A nova data de agendamento está fora do período de vigência da turma");
    }
  }

  function obterAgendamentoParaEditar_(curso, turma, sheetRow) {
    const c = String(curso || "").trim();
    const t = String(turma || "").trim();
    const r = parseInt(sheetRow, 10);
    if (!c || !t) throw new Error("Selecione curso e turma.");
    if (isNaN(r) || r < 2) throw new Error("Linha do agendamento inválida.");

    const idTurma = RegistroRepo.buscarIdPorCursoTurma(c, t);
    if (!idTurma) {
      throw new Error(
        "Não existe registro na planilha de turmas para o curso " + citarRotuloMsg_(c) +
          " e a turma " + citarRotuloMsg_(t) + "."
      );
    }

    const linha = AgendamentoRepo.obterLinhaAgPorSheetRow(r);
    if (!linha) throw new Error("Agendamento não encontrado na planilha.");

    const cells = linha.cells;
    const C = AgendamentoRepo.COL_AG;
    if (String(cells[C.CURSO] || "").trim() !== c) throw new Error("O agendamento selecionado não pertence ao curso informado.");
    if (String(cells[C.TURMA] || "").trim() !== t) throw new Error("O agendamento selecionado não pertence à turma informada.");
    if (String(cells[C.ID_REGISTRO_TURMA] || "").trim() !== idTurma) {
      throw new Error("Dados da turma não conferem com o registro selecionado. Recarregue a tela e tente novamente.");
    }

    PermissaoService.garantirPodeAgendarTurmaPorId(idTurma);

    const vig = obterVigenciaTurmaOuErro_(c, t);
    const dadosIncluir = obterDadosIncluir_();
    return {
      curso: c,
      turma: t,
      idTurma: idTurma,
      sheetRow: r,
      eventId: String(linha.eventId || "").trim(),
      tituloEvento: montarTitulo_(t, c),
      data: String(cells[C.DATA] || "").trim(),
      horaInicio: String(cells[C.HORA_INI] || "").trim(),
      horaFim: String(cells[C.HORA_FIM] || "").trim(),
      salaNome: String(cells[C.NOME_SALA] || "").trim(),
      salas: dadosIncluir.salas,
      hojeYmd: dadosIncluir.hojeYmd,
      vigenciaInicioYmd: vig.inicioYmd,
      vigenciaFimYmd: vig.fimYmd,
      timezone: dadosIncluir.timezone
    };
  }

  function atualizarAgendamento_(payload) {
    if (!payload || typeof payload !== "object") throw new Error("Dados inválidos.");

    const curso = String(payload.curso || "").trim();
    const turma = String(payload.turma || "").trim();
    const turmaIdCliente = String(payload.turmaId || "").trim();
    const sheetRow = parseInt(payload.sheetRow, 10);
    if (!curso || !turma) throw new Error("Selecione curso e turma.");
    if (isNaN(sheetRow) || sheetRow < 2) throw new Error("Linha do agendamento inválida.");

    const idTurma = RegistroRepo.buscarIdPorCursoTurma(curso, turma);
    if (!idTurma) {
      throw new Error(
        "Não existe registro na planilha de turmas para o curso " + citarRotuloMsg_(curso) +
          " e a turma " + citarRotuloMsg_(turma) + "."
      );
    }
    if (turmaIdCliente && turmaIdCliente !== idTurma) {
      throw new Error("Dados da turma não conferem com o registro selecionado. Recarregue a tela e tente novamente.");
    }

    const linha = AgendamentoRepo.obterLinhaAgPorSheetRow(sheetRow);
    if (!linha) throw new Error("Agendamento não encontrado na planilha.");

    const cells = linha.cells;
    const C = AgendamentoRepo.COL_AG;
    if (String(cells[C.CURSO] || "").trim() !== curso) throw new Error("O agendamento selecionado não pertence ao curso informado.");
    if (String(cells[C.TURMA] || "").trim() !== turma) throw new Error("O agendamento selecionado não pertence à turma informada.");
    if (String(cells[C.ID_REGISTRO_TURMA] || "").trim() !== idTurma) {
      throw new Error("Dados da turma não conferem com o registro selecionado. Recarregue a tela e tente novamente.");
    }

    PermissaoService.garantirPodeAgendarTurmaPorId(idTurma);

    const dataYmd = String(payload.data || "").trim();
    parseYmd_(dataYmd);
    const horaInicio = validarHora_(payload.horaInicio, "Hora início");
    const horaFim = validarHora_(payload.horaFim, "Hora fim");
    const minIni = parseInt(horaInicio.split(":")[0], 10) * 60 + parseInt(horaInicio.split(":")[1], 10);
    const minFim = parseInt(horaFim.split(":")[0], 10) * 60 + parseInt(horaFim.split(":")[1], 10);
    if (minFim <= minIni) throw new Error("Hora fim deve ser posterior à hora início.");

    validarNovaDataEdicaoAgendamento_(dataYmd, curso, turma);
    const salaNome = validarSalaNome_(payload.salaNome || "");

    const periodos = [{
      ymd: dataYmd,
      startMs: parseDataHoraLocalMs_(dataYmd, horaInicio),
      endMs: parseDataHoraLocalMs_(dataYmd, horaFim)
    }];

    if (salaNome) checarSalaLivrePeriodos_(salaNome, periodos, sheetRow);

    const newRow = cells.slice();
    newRow[C.DATA] = dataYmd;
    newRow[C.NOME_SALA] = salaNome || "";
    newRow[C.HORA_INI] = horaInicio;
    newRow[C.HORA_FIM] = horaFim;

    AgendamentoRepo.atualizarLinhaCompletaAg(sheetRow, newRow);
  }

  return {
    criarEventos: criarEventos_,
    obterDadosIncluir: obterDadosIncluir_,
    listarTurmasPorCursoIncluir: listarTurmasPorCursoIncluir_,
    criarAgendamentos: criarAgendamentos_,
    pesquisarAgendamentosExcluir: pesquisarAgendamentosExcluir_,
    obterAgendamentosConsultaParaExportar: obterAgendamentosConsultaParaExportar_,
    obterTodosEventIdsExcluir: obterTodosEventIdsExcluir_,
    excluirAgendamentosLote: excluirAgendamentosLote_,
    excluirTodosAgendamentosPorIdTurmaAoExcluirRegistro: excluirTodosAgendamentosPorIdTurmaAoExcluirRegistro_,
    obterAgendamentoParaEditar: obterAgendamentoParaEditar_,
    atualizarAgendamento: atualizarAgendamento_
  };
})();

function criarEventosAgendamento(request) {
  return AgendamentoService.criarEventos(request);
}
