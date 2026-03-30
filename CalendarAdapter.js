/**
 * Usa apenas o serviço avançado Google Calendar API (global Calendar).
 * Não usa UrlFetchApp — evita o escopo script.external_request na Web App,
 * onde esse escopo costuma não ser aplicado mesmo estando no manifest.
 *
 * Se Calendar não existir no contexto da execução, lança erro explicativo.
 */
const CalendarAdapter = (() => {
  function garantirCalendar_() {
    try {
      if (
        typeof Calendar !== "undefined" &&
        Calendar.Freebusy &&
        typeof Calendar.Freebusy.query === "function" &&
        Calendar.Events &&
        typeof Calendar.Events.insert === "function" &&
        typeof Calendar.Events.get === "function" &&
        typeof Calendar.Events.patch === "function"
      ) {
        return;
      }
    } catch (e) {
      /* continua para o throw */
    }
    throw new Error(
      "Google Calendar API (serviço avançado) não está disponível nesta execução. " +
        "No editor: Serviços (+) → Google Calendar API → Salvar. " +
        "Depois: Implantação da Web App → Nova versão → Implantar. " +
        "Se usar clasp push, abra o projeto no Google e confira se o serviço Calendar ainda aparece em Serviços."
    );
  }

  function usandoServicoAvancado() {
    try {
      return (
        typeof Calendar !== "undefined" &&
        Calendar.Events &&
        typeof Calendar.Events.insert === "function" &&
        typeof Calendar.Events.get === "function" &&
        typeof Calendar.Events.patch === "function"
      );
    } catch (e) {
      return false;
    }
  }

  function freeBusyQuery(body) {
    garantirCalendar_();
    return Calendar.Freebusy.query(body);
  }

  function eventsInsertPrimary(evento) {
    garantirCalendar_();
    return Calendar.Events.insert(evento, "primary", { sendUpdates: "all" });
  }

  function eventsRemovePrimary(eventId) {
    garantirCalendar_();
    try {
      Calendar.Events.remove("primary", eventId);
    } catch (e) {
      /* rollback: melhor esforço */
    }
  }

  /**
   * Evento no calendário primário (organizador).
   * @param {string} eventId
   * @returns {Object}
   */
  function eventsGetPrimary(eventId) {
    garantirCalendar_();
    const id = String(eventId || "").trim();
    if (!id) throw new Error("ID do evento inválido.");
    return Calendar.Events.get("primary", id);
  }

  /**
   * Atualiza evento no primário (campos parciais). sendUpdates padrão: all.
   * @param {string} eventId
   * @param {Object} resource
   * @param {{ sendUpdates?: string }=} optionalArgs
   */
  function eventsPatchPrimary(eventId, resource, optionalArgs) {
    garantirCalendar_();
    const id = String(eventId || "").trim();
    if (!id) throw new Error("ID do evento inválido.");
    const opts = optionalArgs && typeof optionalArgs === "object" ? Object.assign({}, optionalArgs) : {};
    if (!opts.sendUpdates) opts.sendUpdates = "all";
    return Calendar.Events.patch(resource, "primary", id, opts);
  }

  /**
   * Remove evento no calendário primário. 404 / já removido = sucesso (idempotente).
   * Outros erros propagam (exclusão em lote “tudo ou nada” após validação).
   * @param {string} eventId
   */
  function eventsRemovePrimaryIdempotent(eventId) {
    garantirCalendar_();
    const id = String(eventId || "").trim();
    if (!id) throw new Error("ID do evento inválido.");
    try {
      Calendar.Events.remove("primary", id);
    } catch (e) {
      if (isCalendarNotFoundError_(e)) return;
      const msg = e && e.message ? String(e.message) : String(e);
      throw new Error(msg);
    }
  }

  function isCalendarNotFoundError_(e) {
    if (!e) return false;
    try {
      const code = e.responseCode || e.statusCode || (e.details && e.details.status);
      if (code === 404 || code === "404") return true;
    } catch (_) {}
    const m = String(e.message || e.toString() || "");
    return /404|not\s*found|Not Found|Requested entity was not found/i.test(m);
  }

  return {
    usandoServicoAvancado: usandoServicoAvancado,
    freeBusyQuery: freeBusyQuery,
    eventsInsertPrimary: eventsInsertPrimary,
    eventsRemovePrimary: eventsRemovePrimary,
    eventsRemovePrimaryIdempotent: eventsRemovePrimaryIdempotent,
    eventsGetPrimary: eventsGetPrimary,
    eventsPatchPrimary: eventsPatchPrimary
  };
})();
