/**
 * Camada Repository – Persistencia em Sheets e operações com Google Calendar.
 * (Dados reais entram/saiem deste nível; regras não devem estar aqui.)
 */

const AgendamentoRepo = (() => {
  function obterSheet_(spreadsheetId) {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const nomeAba = (Configuracoes.NOME_ABA_AGENDAMENTO || Configuracoes.NOME_ABA || "").toString();
    if (nomeAba) {
      const sheet = ss.getSheetByName(nomeAba);
      if (sheet) return sheet;
    }
    // Fallback: primeira aba.
    return ss.getSheets()[0];
  }

  function appendEventosNaPlanilha_(spreadsheetId, rowsA13) {
    if (!Array.isArray(rowsA13) || !rowsA13.length) return;
    const sheet = obterSheet_(spreadsheetId);
    const lastRow = sheet.getLastRow();
    const startRow = Math.max(1, lastRow + 1);
    sheet.getRange(startRow, 1, rowsA13.length, rowsA13[0].length).setValues(rowsA13);
  }

  function rollbackEventosCalendario_(createdEvents) {
    if (!Array.isArray(createdEvents) || !createdEvents.length) return;
    createdEvents.slice().reverse().forEach(meta => {
      try {
        const calendar = CalendarApp.getCalendarById(meta.calendarId);
        if (!calendar) return;
        const event = calendar.getEventById(meta.eventId);
        if (event) event.deleteEvent();
      } catch (_) {
        // Mantém rollback best-effort: não mascarar o erro original.
      }
    });
  }

  function criarEventosCalendario_(titulo, descricao, timezone, agendamentos) {
    const createdEvents = [];
    const eventosResponse = [];

    try {
      // IMPORTANTE: assume que agendamentos já vêm com dtInicio/dtFim válidos.
      for (let i = 0; i < agendamentos.length; i++) {
        const item = agendamentos[i];
        const calendar = CalendarApp.getCalendarById(item.salaId);
        if (!calendar) {
          throw new Error("CalendarId não encontrado: " + item.salaId);
        }

        const options = {
          description: descricao || ""
        };
        if (item.salaNome) options.location = item.salaNome;

        if (Array.isArray(item.convidados) && item.convidados.length) {
          // CalendarApp espera string (lista) para guests.
          options.guests = item.convidados.join(",");
          // sendInvites: intentamos convidar.
          options.sendInvites = true;
        }

        let event;
        try {
          event = calendar.createEvent(
            titulo,
            item.dtInicio,
            item.dtFim,
            options
          );
        } catch (err) {
          // Permite ao Service mapear o erro para detalhes do item na posição.
          throw { itemIndex: i, message: (err && err.message) ? err.message : String(err) };
        }

        const eventId = event.getId();
        createdEvents.push({ calendarId: item.salaId, eventId });
        eventosResponse.push({
          idGoogle: eventId,
          data: item.data,
          horaInicio: item.horaInicio,
          horaFim: item.horaFim
        });
      }
      return { createdEvents, eventosResponse };
    } catch (err) {
      // All-or-nothing para Calendar: desfaz o que já criou.
      rollbackEventosCalendario_(createdEvents);
      throw err;
    }
  }

  return {
    appendEventosNaPlanilha: appendEventosNaPlanilha_,
    rollbackEventosCalendario: rollbackEventosCalendario_,
    criarEventosCalendario: criarEventosCalendario_
  };
})();

// Wrappers globais (compatibilidade padrão do projeto)
function appendEventosNaPlanilha(spreadsheetId, rowsA13) { return AgendamentoRepo.appendEventosNaPlanilha(spreadsheetId, rowsA13); }
function rollbackEventosCalendario(createdEvents) { return AgendamentoRepo.rollbackEventosCalendario(createdEvents); }
function criarEventosCalendario(titulo, descricao, timezone, agendamentos) { return AgendamentoRepo.criarEventosCalendario(titulo, descricao, timezone, agendamentos); }

