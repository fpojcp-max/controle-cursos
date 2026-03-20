/**
 * Camada Service – Agendamento no Google Calendar + gravação na planilha de ocorrências.
 */

const AgendamentoService = (() => {
  const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  const HORA_REGEX = /^([01]?\d|2[0-3]):([0-5]\d)$/;

  const DIAS_JS = { seg: 1, ter: 2, qua: 3, qui: 4, sex: 5 };

  function tz_() {
    return Configuracoes.TIMEZONE_AGENDAMENTO || "America/Sao_Paulo";
  }

  function limiteConvidados_() {
    const n = Number(Configuracoes.LIMITE_CONVIDADOS_AGENDAMENTO);
    return n > 0 ? n : 50;
  }

  function mapaSalas_() {
    const list = Configuracoes.SALAS_RECURSOS_CALENDAR || [];
    const map = {};
    list.forEach((s) => {
      if (s && s.nome) map[String(s.nome).trim()] = String(s.calendarId || "").trim();
    });
    return map;
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
      parseYmd_(d);
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
      parseYmd_(rdate);
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
    const evento = String(turma || "").trim() + " - " + String(curso || "").trim();
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
    const evento = String(turma || "").trim() + " - " + String(curso || "").trim();
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
    const salas = (Configuracoes.SALAS_RECURSOS_CALENDAR || []).map((s) => ({
      nome: s.nome,
      configurada: !!(s.calendarId && String(s.calendarId).trim())
    }));
    const cursos = RegistroRepo.listarCursosDistintos();
    return { cursos: cursos, salas: salas, timezone: tz_() };
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
    if (!curso || !turma) {
      throw new Error("Selecione curso e turma.");
    }

    const idTurma = RegistroRepo.buscarIdPorCursoTurma(curso, turma);
    if (!idTurma) {
      throw new Error(
        "Não existe registro na planilha de turmas para a combinação Curso + Turma selecionada."
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
    const mapa = mapaSalas_();
    let salaId = "";
    if (salaNome) {
      if (!Object.prototype.hasOwnProperty.call(mapa, salaNome)) {
        throw new Error("Sala não reconhecida: " + salaNome);
      }
      salaId = mapa[salaNome];
      if (!salaId) {
        throw new Error(
          "Sala \"" + salaNome + "\" sem calendarId em Configuracoes.SALAS_RECURSOS_CALENDAR."
        );
      }
    }

    const datas = expandirOcorrencias_(payload);
    if (!datas.length) {
      throw new Error("Nenhuma ocorrência no período e dias selecionados.");
    }

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

  return {
    obterDadosIncluir: obterDadosIncluir_,
    listarTurmasPorCursoIncluir: listarTurmasPorCursoIncluir_,
    criarAgendamentos: criarAgendamentos_
  };
})();
