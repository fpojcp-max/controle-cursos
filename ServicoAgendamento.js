/**
 * Camada Service – Agendamento: API externa (criarEventos) e Web App (criarAgendamentos).
 */

const AgendamentoService = (() => {
  function isPlainObject_(v) {
    return v !== null && typeof v === "object" && !Array.isArray(v);
  }

  function isNonEmptyString_(v) {
    return v !== null && v !== undefined && String(v).trim() !== "";
  }

  /** Delimita rótulos (turma, curso, etc.) em mensagens ao utilizador. */
  function citarRotuloMsg_(texto) {
    return "'" + String(texto != null ? texto : "").trim() + "'";
  }

  function errorValidation_(message, details) {
    throw { code: "VALIDATION_ERROR", message: String(message || "Erro de validação"), details: details || [] };
  }

  function errorProcessing_(message, details) {
    throw { code: "PROCESSING_ERROR", message: String(message || "Erro ao processar"), details: details || [] };
  }

  function errorInternal_(message, details) {
    throw { code: "INTERNAL_ERROR", message: String(message || "Erro interno"), details: details || [] };
  }

  function criarEventos_(request) {
    if (!isPlainObject_(request)) {
      errorValidation_("Payload inválido", [{ field: "request", message: "JSON inválido" }]);
    }

    const details = [];

    const idReferencia = request.idReferencia;
    if (!isNonEmptyString_(idReferencia)) errorValidation_("idReferencia obrigatório", [{ field: "idReferencia", message: "Obrigatório" }]);

    const spreadsheetId = request.spreadsheetId;
    if (!isNonEmptyString_(spreadsheetId)) errorValidation_("spreadsheetId obrigatório", [{ field: "spreadsheetId", message: "Obrigatório" }]);

    const titulo = request.titulo;
    if (!isNonEmptyString_(titulo)) errorValidation_("titulo obrigatório", [{ field: "titulo", message: "Obrigatório" }]);

    const agendamentos = request.agendamentos;
    if (!Array.isArray(agendamentos)) {
      errorValidation_("agendamentos deve ser um array", [{ field: "agendamentos", message: "Array obrigatório" }]);
    }
    if (agendamentos.length < 1 || agendamentos.length > 100) {
      errorValidation_("agendamentos deve ter tamanho entre 1 e 100", [{ field: "agendamentos", message: "Tamanho inválido" }]);
    }

    const descricao = request.descricao || "";
    const criador = request.criador || "";

    // 1) Valida tudo antes de criar qualquer evento (all-or-nothing).
    const agendamentosNorm = [];
    for (let i = 0; i < agendamentos.length; i++) {
      const item = agendamentos[i];
      const baseField = `agendamentos[${i}]`;

      if (!isPlainObject_(item)) {
        errorValidation_("Agendamento inválido", [{ field: baseField, message: "Objeto obrigatório" }]);
      }

      const tipo = item.tipo;
      if (!isNonEmptyString_(tipo) || !["simples", "recorrente"].includes(String(tipo))) {
        errorValidation_("tipo inválido", [{ field: `${baseField}.tipo`, message: "Inválido" }]);
      }

      // data + horários + regra fim > inicio.
      const parsed = AgendamentoData.parseStartEnd_(item.data, item.horaInicio, item.horaFim, baseField, details);
      if (!parsed) {
        // parseStartEnd_ já populou detalhes[0] com field completo.
        const d = details.length ? details[0] : { field: `${baseField}.data`, message: "Inválido" };
        errorValidation_(d.message, [d]);
      }

      const salaId = item.salaId;
      if (!isNonEmptyString_(salaId)) errorValidation_("salaId obrigatório", [{ field: `${baseField}.salaId`, message: "Obrigatório" }]);

      const salaNome = item.salaNome;
      if (!isNonEmptyString_(salaNome)) errorValidation_("salaNome obrigatório", [{ field: `${baseField}.salaNome`, message: "Obrigatório" }]);

      let convidados = undefined;
      if (item.hasOwnProperty("convidados")) {
        if (item.convidados === null || item.convidados === undefined) {
          convidados = [];
        } else if (!Array.isArray(item.convidados)) {
          errorValidation_("convidados deve ser array", [{ field: `${baseField}.convidados`, message: "Array obrigatório" }]);
        } else {
          convidados = item.convidados.map(v => String(v).trim()).filter(v => v !== "");
        }
      }

      agendamentosNorm.push({
        tipo: String(tipo),
        data: String(item.data),
        horaInicio: String(item.horaInicio),
        horaFim: String(item.horaFim),
        salaId: String(salaId),
        salaNome: String(salaNome),
        convidados: convidados || undefined,
        dtInicio: parsed.dtInicio,
        dtFim: parsed.dtFim
      });

      // Limpa para garantir "primeiro erro" determinístico por item.
      details.length = 0;
    }

    const hojeApi = dataCivilHojeYmd_();
    for (let j = 0; j < agendamentosNorm.length; j++) {
      const it = agendamentosNorm[j];
      const dApi = String(it.data || "").trim();
      if (dApi < hojeApi) {
        errorValidation_("Datas passadas não são permitidas", [{ field: `agendamentos[${j}].data`, message: "Inválido" }]);
      }
      if (String(it.tipo).toLowerCase() === "simples") {
        const dtS = parseYmd_(dApi);
        const dowS = dtS.getDay();
        if (dowS === 0 || dowS === 6) {
          errorValidation_("Não são permitidos agendamentos para sábados e domingos.", [
            { field: `agendamentos[${j}].data`, message: "Inválido" }
          ]);
        }
      }
    }

    // 2) Cria eventos no Calendar (rollback em caso de falha).
    let createdEvents = [];
    let eventosResponse = [];
    try {
      const repoResp = criarEventosCalendarioInternal_(titulo, descricao, AgendamentoData.TIMEZONE, agendamentosNorm);
      createdEvents = repoResp.createdEvents;
      eventosResponse = repoResp.eventosResponse;
    } catch (err) {
      if (err && err.code && err.details) throw err;
      // Erro de processing por item.
      if (err && typeof err.itemIndex === "number") {
        const idx = err.itemIndex;
        errorProcessing_(err.message || "Erro ao processar agendamento", [{ field: `agendamentos[${idx}]`, message: "Erro ao criar evento no Calendar" }]);
      }
      errorProcessing_(
        err && err.message ? err.message : "Erro ao processar agendamentos",
        []
      );
    }

    // 3) Apêndice na planilha (append-only).
    try {
      const createdAt = new Date();
      const linhasA13 = agendamentosNorm.map((item, idx) => {
        const idUuid = Utilities.getUuid();
        const idGoogle = eventosResponse[idx] ? eventosResponse[idx].idGoogle : "";
        const convidadosVal = Array.isArray(item.convidados) ? item.convidados.join(",") : "";
        return [
          idUuid,
          idGoogle,
          idReferencia,
          item.data,
          titulo,
          descricao || "",
          item.horaInicio,
          item.horaFim,
          item.salaNome,
          item.salaId,
          convidadosVal,
          createdAt,
          criador || ""
        ];
      });

      appendEventosNaPlanilhaInternal_(spreadsheetId, linhasA13);
    } catch (err) {
      // Garantia: se persistir falhar após criar eventos, desfaz eventos.
      try {
        rollbackEventosCalendarioInternal_(createdEvents);
      } catch (_) {}
      errorProcessing_(err && err.message ? err.message : "Erro ao persistir na planilha", []);
    }

    return {
      status: "ok",
      total: eventosResponse.length,
      eventos: eventosResponse
    };
  }

  // Indireção para manter legibilidade (Service não deve conhecer detalhes internos do repo).
  function criarEventosCalendarioInternal_(titulo, descricao, timezone, agendamentosNorm) {
    return criarEventosCalendario(titulo, descricao, timezone, agendamentosNorm);
  }
  function appendEventosNaPlanilhaInternal_(spreadsheetId, linhasA13) {
    return appendEventosNaPlanilha(spreadsheetId, linhasA13);
  }
  function rollbackEventosCalendarioInternal_(createdEvents) {
    return rollbackEventosCalendario(createdEvents);
  }

  const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  const HORA_REGEX = /^([01]?\d|2[0-3]):([0-5]\d)$/;

  const DIAS_JS = { seg: 1, ter: 2, qua: 3, qui: 4, sex: 5 };

  function tz_() {
    return Configuracoes.TIMEZONE_AGENDAMENTO || "America/Sao_Paulo";
  }

  /** Data civil corrente no fuso de agendamento (yyyy-MM-dd), para comparar com ocorrências. */
  function dataCivilHojeYmd_() {
    return Utilities.formatDate(new Date(), tz_(), "yyyy-MM-dd");
  }

  function limiteConvidados_() {
    const n = Number(Configuracoes.LIMITE_CONVIDADOS_AGENDAMENTO);
    return n > 0 ? n : 50;
  }

  function mapaIdentificadorCalendarioPorRotulo_() {
    return montarMapaRotuloParaIdentificadorCalendario();
  }

  function validarHora_(h, nome) {
    if (!h || !HORA_REGEX.test(String(h).trim())) {
      throw new Error(nome + " inválida. Use HH:mm (00:00 a 23:59).");
    }
    return String(h).trim();
  }

  function parseYmd_(s) {
    const t = String(s || "").trim();
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(t);
    if (!m) throw new Error("Data inválida: use YYYY-MM-DD.");
    const y = parseInt(m[1], 10);
    const mo = parseInt(m[2], 10) - 1;
    const d = parseInt(m[3], 10);
    const dt = new Date(y, mo, d);
    if (dt.getFullYear() !== y || dt.getMonth() !== mo || dt.getDate() !== d) {
      throw new Error("Data inválida: " + t);
    }
    return dt;
  }

  function formatYmd_(d) {
    const y = d.getFullYear();
    const mo = ("0" + (d.getMonth() + 1)).slice(-2);
    const da = ("0" + d.getDate()).slice(-2);
    return y + "-" + mo + "-" + da;
  }

  function normalizarCelulaVigenciaParaYmd_(valor, nomeCampo) {
    if (valor === null || valor === undefined) return "";
    if (Object.prototype.toString.call(valor) === "[object Date]" && !isNaN(valor.getTime())) {
      return Utilities.formatDate(valor, tz_(), "yyyy-MM-dd");
    }
    const s = String(valor).trim();
    if (!s) return "";
    const mIso = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
    if (mIso) {
      parseYmd_(s);
      return s;
    }
    const mBr = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s);
    if (mBr) {
      const da = parseInt(mBr[1], 10);
      const mo = parseInt(mBr[2], 10);
      const y = parseInt(mBr[3], 10);
      const dt = new Date(y, mo - 1, da);
      if (dt.getFullYear() !== y || dt.getMonth() !== mo - 1 || dt.getDate() !== da) {
        throw new Error("Data de " + nomeCampo + " da turma na planilha é inválida.");
      }
      return formatYmd_(dt);
    }
    throw new Error("Data de " + nomeCampo + " da turma na planilha não está em formato reconhecido.");
  }

  function obterVigenciaTurmaOuErro_(curso, turma) {
    const raw = RegistroRepo.buscarVigenciaInicioFimPorCursoTurma(curso, turma);
    if (!raw) {
      throw new Error(
        "Não existe registro na planilha de turmas para o curso " +
          citarRotuloMsg_(curso) +
          " e a turma " +
          citarRotuloMsg_(turma) +
          "."
      );
    }
    const inicioYmd = normalizarCelulaVigenciaParaYmd_(raw.inicioVal, "Início");
    const fimYmd = normalizarCelulaVigenciaParaYmd_(raw.fimVal, "Fim");
    if (!inicioYmd || !fimYmd) {
      throw new Error(
        "A turma " +
          citarRotuloMsg_(turma) +
          " do curso " +
          citarRotuloMsg_(curso) +
          " não possui datas de Início e Fim de vigência definidas na planilha."
      );
    }
    if (inicioYmd > fimYmd) {
      throw new Error(
        "Período de vigência inválido (Início após Fim) para a turma " +
          citarRotuloMsg_(turma) +
          " do curso " +
          citarRotuloMsg_(curso) +
          "."
      );
    }
    return { inicioYmd: inicioYmd, fimYmd: fimYmd };
  }

  function lancarForaVigenciaTurma_(turma, curso) {
    throw new Error(
      "Agendamento não permitido. Possui data fora do período de vigência da turma " +
        citarRotuloMsg_(turma) +
        " do curso " +
        citarRotuloMsg_(curso) +
        "."
    );
  }

  function validarConjuntoDatasAgendamento_(payload, datasYmd, curso, turma) {
    const hoje = dataCivilHojeYmd_();
    const vig = obterVigenciaTurmaOuErro_(curso, turma);
    function checarYmd(ymd) {
      if (ymd < hoje) {
        throw new Error("Datas passadas não são permitidas");
      }
      if (ymd < vig.inicioYmd || ymd > vig.fimYmd) {
        lancarForaVigenciaTurma_(turma, curso);
      }
    }
    for (let i = 0; i < datasYmd.length; i++) {
      checarYmd(datasYmd[i]);
    }
    const exdate = String(payload.exdate || "").trim();
    const rdate = String(payload.rdate || "").trim();
    if (exdate) {
      parseYmd_(exdate);
      checarYmd(exdate);
    }
    if (rdate) {
      parseYmd_(rdate);
      checarYmd(rdate);
    }
  }

  /** Somente mensagens ao usuário: yyyy-mm-dd → dd/mm/aaaa (planilha/API seguem em ISO). */
  function formatarYmdParaMsgBr_(ymd) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(ymd || "").trim());
    if (!m) return String(ymd || "").trim();
    return m[3] + "/" + m[2] + "/" + m[1];
  }

  /** Datas do intervalo [inicio, fim] inclusivo (calendário local do runtime). */
  function enumerarDatasInclusive_(inicioStr, fimStr) {
    const ini = parseYmd_(inicioStr);
    const fim = parseYmd_(fimStr);
    if (ini.getTime() > fim.getTime()) {
      throw new Error("A data inicial não pode ser posterior à data final.");
    }
    const out = [];
    const cur = new Date(ini.getFullYear(), ini.getMonth(), ini.getDate());
    const end = new Date(fim.getFullYear(), fim.getMonth(), fim.getDate());
    while (cur.getTime() <= end.getTime()) {
      out.push(formatYmd_(cur));
      cur.setDate(cur.getDate() + 1);
    }
    return out;
  }

  function normalizarEmails_(texto) {
    const raw = String(texto || "")
      .split(/[;,]+/)
      .map((s) => s.trim())
      .filter(Boolean);
    const seen = {};
    const out = [];
    for (let i = 0; i < raw.length; i++) {
      const e = raw[i].toLowerCase();
      if (!EMAIL_REGEX.test(raw[i])) {
        throw new Error("E-mail inválido: " + raw[i]);
      }
      if (seen[e]) continue;
      seen[e] = true;
      out.push(raw[i]);
    }
    const lim = limiteConvidados_();
    if (out.length > lim) {
      throw new Error("Máximo de " + lim + " convidados.");
    }
    return out;
  }

  function montarTitulo_(turma, curso) {
    return String(turma || "").trim() + " - " + String(curso || "").trim();
  }

  function montarDateTimeIso_(dataYmd, horaHm) {
    const hi = validarHora_(horaHm, "Hora");
    return dataYmd + "T" + hi + ":00";
  }

  /** dateTime local para API (hh:mm já validado). */
  function isoLocalSemValidar_(ymd, hhmm) {
    return ymd + "T" + String(hhmm).trim() + ":00";
  }

  /** Instante em ms no fuso TIMEZONE_AGENDAMENTO (para FreeBusy); hh:mm já validado. */
  function parseDataHoraLocalMsLivre_(dataYmd, hhmm) {
    const s = dataYmd + " " + String(hhmm).trim() + ":00";
    return Utilities.parseDate(s, tz_(), "yyyy-MM-dd HH:mm:ss").getTime();
  }

  function parseRfc3339ParaMs_(s) {
    return new Date(s).getTime();
  }

  function intervalosSobrepoem_(aStartMs, aEndMs, bStartMs, bEndMs) {
    return aStartMs < bEndMs && bStartMs < aEndMs;
  }

  function expandirOcorrencias_(payload) {
    const tipo = String(payload.tipo || "").toLowerCase();
    const exdate = String(payload.exdate || "").trim();
    const rdate = String(payload.rdate || "").trim();

    if (tipo === "simples") {
      const d = String(payload.data || "").trim();
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

    if (tipo !== "recorrente") {
      throw new Error("Tipo de evento inválido.");
    }

    const dias = payload.diasSemana || {};
    const ativos = Object.keys(DIAS_JS).filter((k) => dias[k] === true || dias[k] === "true" || dias[k] === 1);
    if (!ativos.length) {
      throw new Error("Selecione pelo menos um dia da semana (segunda a sexta).");
    }
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
      const dow = dt.getDay();
      if (!permitidos[dow]) return;
      if (exdate && ymd === exdate) return;
      pushUnique(ymd);
    });

    if (rdate) {
      const dtR = parseYmd_(rdate);
      const dowR = dtR.getDay();
      if (dowR === 0 || dowR === 6) {
        throw new Error("Não são permitidos agendamentos para sábados e domingos.");
      }
      pushUnique(rdate);
    }

    setDatas.sort();
    return setDatas;
  }

  function checarSalaLivrePeriodos_(calendarIdSala, periodos) {
    if (!periodos.length) return;
    let minStart = periodos[0].startMs;
    let maxEnd = periodos[0].endMs;
    for (let i = 1; i < periodos.length; i++) {
      if (periodos[i].startMs < minStart) minStart = periodos[i].startMs;
      if (periodos[i].endMs > maxEnd) maxEnd = periodos[i].endMs;
    }
    const body = {
      timeMin: new Date(minStart - 60000).toISOString(),
      timeMax: new Date(maxEnd + 60000).toISOString(),
      items: [{ id: calendarIdSala }],
      timeZone: tz_()
    };
    const resp = CalendarAdapter.freeBusyQuery(body);
    const cal = resp.calendars && resp.calendars[calendarIdSala];
    if (!cal) {
      throw new Error(
        "Não foi possível consultar a disponibilidade da sala. Verifique o ID do recurso em Configuracoes."
      );
    }
    const busy = cal.busy || [];
    for (let p = 0; p < periodos.length; p++) {
      const ev = periodos[p];
      for (let b = 0; b < busy.length; b++) {
        const bs = parseRfc3339ParaMs_(busy[b].start);
        const be = parseRfc3339ParaMs_(busy[b].end);
        if (intervalosSobrepoem_(ev.startMs, ev.endMs, bs, be)) {
          throw new Error(
            "A sala está ocupada em pelo menos um dos horários. Nenhum evento foi criado."
          );
        }
      }
    }
  }

  function criarEventoNoCalendario_(titulo, startIso, endIso, emails, salaCalendarId) {
    const evento = {
      summary: titulo,
      start: { dateTime: startIso, timeZone: tz_() },
      end: { dateTime: endIso, timeZone: tz_() }
    };
    const atts = [];
    (emails || []).forEach((e) => atts.push({ email: e }));
    if (salaCalendarId) {
      atts.push({ email: salaCalendarId, resource: true });
    }
    if (atts.length) evento.attendees = atts;

    const criado = CalendarAdapter.eventsInsertPrimary(evento);
    return criado && criado.id ? String(criado.id) : null;
  }

  function removerEventosPrimario_(ids) {
    (ids || []).forEach((id) => {
      try {
        CalendarAdapter.eventsRemovePrimary(id);
      } catch (e) {
        // melhor esforço no rollback
      }
    });
  }

  function formatarCriadoEm_() {
    return Utilities.formatDate(new Date(), tz_(), "yyyy-MM-dd HH:mm");
  }

  /** Mensagem de sucesso — evento simples (uma ocorrência), padrão acordado. */
  function montarMensagemSucessoSimples_(turma, curso, horaInicio, horaFim, salaNome, dataYmd) {
    const evento = citarRotuloMsg_(turma) + " - " + citarRotuloMsg_(curso);
    const sala = String(salaNome || "").trim() || "—";
    return (
      "O seguinte agendamento foi gerado:\n" +
      "Evento: " +
      evento +
      "\n" +
      "Horário: " +
      horaInicio +
      " - " +
      horaFim +
      "\n" +
      "Sala: " +
      sala +
      "\n" +
      "Data: " +
      formatarYmdParaMsgBr_(dataYmd)
    );
  }

  const NOME_DIA_SEMANA_PT_ = {
    0: "domingo",
    1: "segunda-feira",
    2: "terça-feira",
    3: "quarta-feira",
    4: "quinta-feira",
    5: "sexta-feira",
    6: "sábado"
  };

  /** Ordem de exibição: seg → sex, depois sáb e dom (se existirem ocorrências). */
  const ORDEM_DIA_SEMANA_ = [1, 2, 3, 4, 5, 6, 0];

  /**
   * Várias ocorrências: cabeçalho fixo + agrupamento por dia da semana e datas em linha.
   */
  function montarMensagemSucessoMultiplos_(turma, curso, horaInicio, horaFim, salaNome, periodos) {
    const evento = citarRotuloMsg_(turma) + " - " + citarRotuloMsg_(curso);
    const sala = String(salaNome || "").trim() || "—";
    const gruposPorDow = {};
    for (let i = 0; i < periodos.length; i++) {
      const ymd = periodos[i].ymd;
      const dt = parseYmd_(ymd);
      const dow = dt.getDay();
      if (!gruposPorDow[dow]) gruposPorDow[dow] = [];
      gruposPorDow[dow].push(ymd);
    }
    for (let g = 0; g < ORDEM_DIA_SEMANA_.length; g++) {
      const d = ORDEM_DIA_SEMANA_[g];
      if (gruposPorDow[d]) gruposPorDow[d].sort();
    }
    let blocosDia = "";
    for (let k = 0; k < ORDEM_DIA_SEMANA_.length; k++) {
      const dow = ORDEM_DIA_SEMANA_[k];
      const lista = gruposPorDow[dow];
      if (!lista || !lista.length) continue;
      blocosDia += "Dia: " + NOME_DIA_SEMANA_PT_[dow] + "\n";
      blocosDia +=
        "Data: " +
        lista.map(function (ymd) {
          return formatarYmdParaMsgBr_(ymd);
        }).join(", ") +
        "\n";
    }
    return (
      "Os seguintes agendamentos foram gerados:\n" +
      "Evento: " +
      evento +
      "\n" +
      "Horário: " +
      horaInicio +
      " - " +
      horaFim +
      "\n" +
      "Sala: " +
      sala +
      "\n" +
      blocosDia.replace(/\s+$/, "")
    );
  }

  function obterDadosIncluir_() {
    const salas = obterEntradasCatalogoRecursosSala().map((s) => ({
      nome: s.rotulo,
      configurada: !!(s.identificadorCalendario && String(s.identificadorCalendario).trim())
    }));
    const cursos = RegistroRepo.listarCursosDistintos();
    return { cursos: cursos, salas: salas, timezone: tz_(), hojeYmd: dataCivilHojeYmd_() };
  }

  function listarTurmasPorCursoIncluir_(curso) {
    return RegistroRepo.listarTurmasDistintasPorCurso(curso);
  }

  function criarAgendamentos_(payload) {
    if (!payload || typeof payload !== "object") {
      throw new Error("Dados inválidos.");
    }

    const curso = String(payload.curso || "").trim();
    const turma = String(payload.turma || "").trim();
    const turmaIdCliente = String(payload.turmaId || "").trim();
    if (!curso || !turma) {
      throw new Error("Selecione curso e turma.");
    }

    const idTurma = RegistroRepo.buscarIdPorCursoTurma(curso, turma);
    if (!idTurma) {
      throw new Error(
        "Não existe registro na planilha de turmas para o curso " +
          citarRotuloMsg_(curso) +
          " e a turma " +
          citarRotuloMsg_(turma) +
          "."
      );
    }
    if (turmaIdCliente && turmaIdCliente !== idTurma) {
      throw new Error(
        "Dados da turma não conferem com o registro selecionado. Recarregue a tela e tente novamente."
      );
    }

    const horaInicio = validarHora_(payload.horaInicio, "Hora início");
    const horaFim = validarHora_(payload.horaFim, "Hora fim");
    const tIni = horaInicio.split(":");
    const tFim = horaFim.split(":");
    const minIni = parseInt(tIni[0], 10) * 60 + parseInt(tIni[1], 10);
    const minFim = parseInt(tFim[0], 10) * 60 + parseInt(tFim[1], 10);
    if (minFim <= minIni) {
      throw new Error("Hora fim deve ser posterior à hora início.");
    }

    const emails = normalizarEmails_(payload.convidados || "");

    const salaNome = String(payload.salaNome || "").trim();
    const mapa = mapaIdentificadorCalendarioPorRotulo_();
    let salaId = "";
    if (salaNome) {
      if (!Object.prototype.hasOwnProperty.call(mapa, salaNome)) {
        throw new Error("Sala não reconhecida: " + salaNome);
      }
      salaId = mapa[salaNome];
      if (!salaId) {
        throw new Error(
          "Sala \"" + salaNome + "\" sem identificador de calendário no catálogo (Configuracoes.CATALOGO_RECURSOS_SALA)."
        );
      }
    }

    const datas = expandirOcorrencias_(payload);
    if (!datas.length) {
      throw new Error("Nenhuma ocorrência no período e dias selecionados.");
    }
    validarConjuntoDatasAgendamento_(payload, datas, curso, turma);

    const titulo = montarTitulo_(turma, curso);
    const periodos = datas.map((ymd) => ({
      ymd: ymd,
      startIso: isoLocalSemValidar_(ymd, horaInicio),
      endIso: isoLocalSemValidar_(ymd, horaFim),
      startMs: parseDataHoraLocalMsLivre_(ymd, horaInicio),
      endMs: parseDataHoraLocalMsLivre_(ymd, horaFim)
    }));

    if (salaId) {
      checarSalaLivrePeriodos_(salaId, periodos);
    }

    const criadoPor = Session.getActiveUser().getEmail() || "";
    const criadoEm = formatarCriadoEm_();
    const convidadosPlanilha = emails.join("; ");

    const idsCriados = [];
    const linhas = [];

    try {
      for (let i = 0; i < periodos.length; i++) {
        const p = periodos[i];
        const eventId = criarEventoNoCalendario_(titulo, p.startIso, p.endIso, emails, salaId);
        if (!eventId) {
          throw new Error("Resposta sem ID do evento no Google Calendar.");
        }
        idsCriados.push(eventId);
        linhas.push([
          turma,
          curso,
          p.ymd,
          salaNome || "",
          horaInicio,
          horaFim,
          convidadosPlanilha,
          criadoEm,
          criadoPor,
          eventId,
          idTurma,
          salaId || ""
        ]);
      }
      AgendamentoRepo.appendLinhas(linhas);
    } catch (err) {
      removerEventosPrimario_(idsCriados);
      throw err;
    }

    let mensagemSucesso;
    if (periodos.length === 1) {
      mensagemSucesso = montarMensagemSucessoSimples_(
        turma,
        curso,
        horaInicio,
        horaFim,
        salaNome,
        periodos[0].ymd
      );
    } else {
      mensagemSucesso = montarMensagemSucessoMultiplos_(
        turma,
        curso,
        horaInicio,
        horaFim,
        salaNome,
        periodos
      );
    }
    return {
      ocorrencias: periodos.length,
      mensagem: mensagemSucesso
    };
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
      const hi = row[C.HORA_INI] || "";
      const hf = row[C.HORA_FIM] || "";
      const sala = row[C.NOME_SALA] || "";
      const dt = row[C.DATA] || "";
      const key = hi + "\t" + hf + "\t" + sala;
      if (!grupos[key]) grupos[key] = [];
      grupos[key].push(dt);
    }
    let out = "Os seguintes agendamentos foram excluídos:\n";
    out += "Evento: " + citarRotuloMsg_(turma) + " - " + citarRotuloMsg_(curso) + "\n";
    const keys = Object.keys(grupos).sort();
    for (let k = 0; k < keys.length; k++) {
      const parts = keys[k].split("\t");
      const hi = parts[0] || "";
      const hf = parts[1] || "";
      const sala = parts[2] || "";
      const datas = grupos[keys[k]]
        .slice()
        .sort()
        .map(function (ymd) {
          return formatarYmdParaMsgBr_(ymd);
        });
      out += "Horário: " + hi + " - " + hf + "\n";
      out += "Sala: " + (sala || "—") + "\n";
      out += datas.join(", ") + "\n";
    }
    return out;
  }

  function pesquisarAgendamentosExcluir_(curso, turma, offset, limit, sortCol, sortDir) {
    const c = String(curso || "").trim();
    const t = String(turma || "").trim();
    if (!c || !t) {
      throw new Error("Selecione curso e turma.");
    }
    const idTurma = RegistroRepo.buscarIdPorCursoTurma(c, t);
    if (!idTurma) {
      throw new Error(
        "Não existe registro na planilha de turmas para o curso " +
          citarRotuloMsg_(c) +
          " e a turma " +
          citarRotuloMsg_(t) +
          "."
      );
    }
    let sc = -1;
    if (sortCol !== undefined && sortCol !== null && String(sortCol).trim() !== "") {
      const n = parseInt(sortCol, 10);
      if (!isNaN(n)) sc = n;
    }
    const sortAsc = String(sortDir == null ? "asc" : sortDir).toLowerCase() !== "desc";
    const r = AgendamentoRepo.listarAgendamentosPaginadoPorIdTurma(
      idTurma,
      offset,
      limit,
      sc,
      sortAsc
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

  /**
   * Todos os agendamentos do filtro (curso+turma) com a mesma ordenação da pesquisa — para CSV.
   * @returns {{ columns: { key: string, label: string }[], rows: string[][] }}
   */
  function obterAgendamentosConsultaParaExportar_(curso, turma, sortCol, sortDir) {
    const c = String(curso || "").trim();
    const t = String(turma || "").trim();
    if (!c || !t) {
      throw new Error("Selecione curso e turma.");
    }
    const idTurma = RegistroRepo.buscarIdPorCursoTurma(c, t);
    if (!idTurma) {
      throw new Error(
        "Não existe registro na planilha de turmas para o curso " +
          citarRotuloMsg_(c) +
          " e a turma " +
          citarRotuloMsg_(t) +
          "."
      );
    }
    let sc = -1;
    if (sortCol !== undefined && sortCol !== null && String(sortCol).trim() !== "") {
      const n = parseInt(sortCol, 10);
      if (!isNaN(n)) sc = n;
    }
    const sortAsc = String(sortDir == null ? "asc" : sortDir).toLowerCase() !== "desc";
    const r = AgendamentoRepo.listarAgendamentosPaginadoPorIdTurma(
      idTurma,
      0,
      LIMITE_EXPORT_AGENDAMENTOS,
      sc,
      sortAsc
    );
    const cab = r.cabecalho || [];
    const columns = cab.map((h, idx) => ({
      key: "c" + idx,
      label: String(h != null ? h : "")
    }));
    const rows = (r.itens || []).map((item) =>
      (item.cells || []).map((cell) => (cell === null || cell === undefined ? "" : String(cell)))
    );
    return { columns: columns, rows: rows };
  }

  function obterTodosEventIdsExcluir_(curso, turma) {
    const c = String(curso || "").trim();
    const t = String(turma || "").trim();
    if (!c || !t) {
      throw new Error("Selecione curso e turma.");
    }
    const idTurma = RegistroRepo.buscarIdPorCursoTurma(c, t);
    if (!idTurma) {
      throw new Error(
        "Não existe registro na planilha de turmas para o curso " +
          citarRotuloMsg_(c) +
          " e a turma " +
          citarRotuloMsg_(t) +
          "."
      );
    }
    const ids = AgendamentoRepo.listarTodosEventIdsPorIdTurma(idTurma);
    return { success: true, eventIds: ids };
  }

  const MSG_EXCLUSAO_TURMA_AGENDAMENTOS_FALHOU =
    "Não foi possível excluir a turma. Tente novamente mais tarde.";

  const MSG_EXCLUSAO_AGENDAMENTOS_LOTE_FALHOU =
    "Não foi possível concluir a exclusão. Tente novamente mais tarde.";

  function uniqueEventIdsOrderedParaExclusaoTurma_(linhas) {
    const seen = {};
    const out = [];
    for (let i = 0; i < linhas.length; i++) {
      const e = String(linhas[i].eventId || "").trim();
      if (!e || seen[e]) continue;
      seen[e] = 1;
      out.push(e);
    }
    return out;
  }

  function recriarEventoAPartirLinhaAg_(cells) {
    const C = AgendamentoRepo.COL_AG;
    const turma = String(cells[C.TURMA] || "").trim();
    const curso = String(cells[C.CURSO] || "").trim();
    const dataYmd = String(cells[C.DATA] || "").trim();
    const hi = validarHora_(cells[C.HORA_INI], "Hora início");
    const hf = validarHora_(cells[C.HORA_FIM], "Hora fim");
    parseYmd_(dataYmd);
    const titulo = montarTitulo_(turma, curso);
    const startIso = isoLocalSemValidar_(dataYmd, hi);
    const endIso = isoLocalSemValidar_(dataYmd, hf);
    const conv = String(cells[C.CONVIDADOS] || "").trim();
    let emails = [];
    if (conv) {
      try {
        emails = normalizarEmails_(conv);
      } catch (_) {
        emails = [];
      }
    }
    const salaNome = String(cells[C.NOME_SALA] || "").trim();
    let salaCalendarId = String(cells[C.ID_SALA] || "").trim();
    if (!salaCalendarId && salaNome) {
      const mm = mapaIdentificadorCalendarioPorRotulo_();
      if (Object.prototype.hasOwnProperty.call(mm, salaNome)) {
        salaCalendarId = String(mm[salaNome] || "").trim();
      }
    }
    const novoId = criarEventoNoCalendario_(titulo, startIso, endIso, emails, salaCalendarId);
    if (!novoId) {
      throw new Error("Resposta sem ID do evento no Google Calendar.");
    }
    return novoId;
  }

  /**
   * Desfaz remoções já feitas no Calendar após falha no meio da sequência.
   * @param {{ sheetRow: number, cells: string[], eventId: string }[]} snapshotLinhas
   * @param {string[]} removedIdsInOrder
   */
  function rollbackRemocoesCalendarExclusaoTurma_(snapshotLinhas, removedIdsInOrder) {
    for (let i = removedIdsInOrder.length - 1; i >= 0; i--) {
      const oldEv = removedIdsInOrder[i];
      const comId = snapshotLinhas.filter(function (l) {
        return String(l.eventId || "").trim() === oldEv;
      });
      if (!comId.length) continue;
      const novoId = recriarEventoAPartirLinhaAg_(comId[0].cells);
      for (let j = 0; j < comId.length; j++) {
        AgendamentoRepo.atualizarIdGoogleNaLinha(comId[j].sheetRow, novoId);
      }
    }
  }

  /**
   * Remove todos os agendamentos ligados ao ID do registro da turma (Calendar + planilha).
   * Sem o limite da UX de lote; falha no Calendar antes de apagar planilha e com rollback das remoções já feitas no Calendar.
   * Falha ao apagar a planilha após Calendar ok: não recria eventos (planilha pode ficar defasada).
   */
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

    const idsCal = uniqueEventIdsOrderedParaExclusaoTurma_(linhas);
    const removedStack = [];
    for (let c = 0; c < idsCal.length; c++) {
      try {
        CalendarAdapter.eventsRemovePrimaryIdempotent(idsCal[c]);
        removedStack.push(idsCal[c]);
      } catch (calErr) {
        try {
          if (removedStack.length) {
            rollbackRemocoesCalendarExclusaoTurma_(linhas, removedStack);
          }
        } catch (_) {
          /* melhor esforço: estado pode ser inconsistente */
        }
        throw new Error(MSG_EXCLUSAO_TURMA_AGENDAMENTOS_FALHOU);
      }
    }

    const rowsDesc = [];
    const seenRow = {};
    for (let r = 0; r < linhas.length; r++) {
      const n = linhas[r].sheetRow;
      if (seenRow[n]) continue;
      seenRow[n] = 1;
      rowsDesc.push(n);
    }
    rowsDesc.sort(function (a, b) {
      return b - a;
    });
    const MAX_TRY = 3;
    let ultimoErro = null;
    for (let t = 0; t < MAX_TRY; t++) {
      try {
        AgendamentoRepo.excluirLinhasPorNumeros(rowsDesc);
        ultimoErro = null;
        break;
      } catch (sheetErr) {
        ultimoErro = sheetErr;
        if (t < MAX_TRY - 1) {
          Utilities.sleep(400 + t * 250);
        }
      }
    }
    if (ultimoErro) {
      throw new Error(MSG_EXCLUSAO_TURMA_AGENDAMENTOS_FALHOU);
    }
  }

  /**
   * @param {string} curso
   * @param {string} turma
   * @param {{ sheetRows?: number[], eventIds?: string[] }} payload — preferir sheetRows (uma linha por checkbox).
   */
  function excluirAgendamentosLote_(curso, turma, payload) {
    const LIM = limiteExclusaoAgendamentosLote_();
    const c = String(curso || "").trim();
    const t = String(turma || "").trim();
    if (!c || !t) {
      throw new Error("Selecione curso e turma.");
    }
    const pay = payload && typeof payload === "object" ? payload : {};
    const idTurma = RegistroRepo.buscarIdPorCursoTurma(c, t);
    if (!idTurma) {
      throw new Error(
        "Não existe registro na planilha de turmas para o curso " +
          citarRotuloMsg_(c) +
          " e a turma " +
          citarRotuloMsg_(t) +
          "."
      );
    }

    const todas = AgendamentoRepo.listarLinhasAgendamentoPorIdTurma(idTurma);
    const byRow = {};
    for (let j = 0; j < todas.length; j++) {
      byRow[todas[j].sheetRow] = todas[j];
    }

    let selecionadas = [];
    const rowsIn = Array.isArray(pay.sheetRows) ? pay.sheetRows : [];

    if (rowsIn.length) {
      const nums = [];
      const seenR = {};
      for (let i = 0; i < rowsIn.length; i++) {
        const n = parseInt(rowsIn[i], 10);
        if (isNaN(n) || n < 2) continue;
        if (seenR[n]) continue;
        seenR[n] = 1;
        nums.push(n);
      }
      if (!nums.length) {
        throw new Error("Selecione ao menos um agendamento.");
      }
      if (nums.length > LIM) {
        throw new Error("Selecione no máximo " + LIM + " agendamentos para excluir por vez.");
      }
      for (let k = 0; k < nums.length; k++) {
        const m = byRow[nums[k]];
        if (!m) {
          throw new Error("Agendamento inválido ou não pertence à turma selecionada.");
        }
        selecionadas.push(m);
      }
    } else {
      const idsIn = Array.isArray(pay.eventIds)
        ? pay.eventIds.map((x) => String(x || "").trim()).filter(Boolean)
        : [];
      const uniq = [];
      const seen = {};
      for (let i = 0; i < idsIn.length; i++) {
        const id = idsIn[i];
        if (!seen[id]) {
          seen[id] = 1;
          uniq.push(id);
        }
      }
      if (!uniq.length) {
        throw new Error("Selecione ao menos um agendamento.");
      }
      if (uniq.length > LIM) {
        throw new Error("Selecione no máximo " + LIM + " agendamentos para excluir por vez.");
      }
      const mapa = {};
      for (let j = 0; j < todas.length; j++) {
        mapa[todas[j].eventId] = todas[j];
      }
      for (let k = 0; k < uniq.length; k++) {
        const ev = uniq[k];
        if (!mapa[ev]) {
          throw new Error("Agendamento inválido ou não pertence à turma selecionada.");
        }
        selecionadas.push(mapa[ev]);
      }
    }

    const calSeen = {};
    for (let idx = 0; idx < selecionadas.length; idx++) {
      const ev = String(selecionadas[idx].eventId || "").trim();
      if (!ev || calSeen[ev]) continue;
      calSeen[ev] = 1;
      try {
        CalendarAdapter.eventsRemovePrimaryIdempotent(ev);
      } catch (calErr) {
        throw new Error(MSG_EXCLUSAO_AGENDAMENTOS_LOTE_FALHOU);
      }
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

  return {
    criarEventos: criarEventos_,
    obterDadosIncluir: obterDadosIncluir_,
    listarTurmasPorCursoIncluir: listarTurmasPorCursoIncluir_,
    criarAgendamentos: criarAgendamentos_,
    pesquisarAgendamentosExcluir: pesquisarAgendamentosExcluir_,
    obterAgendamentosConsultaParaExportar: obterAgendamentosConsultaParaExportar_,
    obterTodosEventIdsExcluir: obterTodosEventIdsExcluir_,
    excluirAgendamentosLote: excluirAgendamentosLote_,
    excluirTodosAgendamentosPorIdTurmaAoExcluirRegistro: excluirTodosAgendamentosPorIdTurmaAoExcluirRegistro_
  };
})();

function criarEventosAgendamento(request) {
  return AgendamentoService.criarEventos(request);
}
