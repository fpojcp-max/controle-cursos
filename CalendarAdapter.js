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
        typeof Calendar.Events.insert === "function"
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
        typeof Calendar.Events.insert === "function"
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

  return {
    usandoServicoAvancado: usandoServicoAvancado,
    freeBusyQuery: freeBusyQuery,
    eventsInsertPrimary: eventsInsertPrimary,
    eventsRemovePrimary: eventsRemovePrimary
  };
})();
