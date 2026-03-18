/**
 * Camada Service – Regras de negócio para agendamento (validação + all-or-nothing).
 */

const AgendamentoService = (() => {
  function isPlainObject_(v) {
    return v !== null && typeof v === "object" && !Array.isArray(v);
  }

  function isNonEmptyString_(v) {
    return v !== null && v !== undefined && String(v).trim() !== "";
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

  return {
    criarEventos: criarEventos_
  };
})();

// Wrappers globais
function criarEventosAgendamento(request) { return AgendamentoService.criarEventos(request); }

