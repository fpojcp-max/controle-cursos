/**
 * Camada Repository – Agendamento:
 * - API: planilha por spreadsheetId + CalendarApp nos recursos.
 * - Web App: planilha dedicada (ocorrências) + Calendar avançado via CalendarAdapter no service.
 */

const AgendamentoRepo = (() => {
  function obterSheet_(spreadsheetId) {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const nomeAba = (Configuracoes.NOME_ABA_AGENDAMENTO || Configuracoes.NOME_ABA || "").toString();
    if (nomeAba) {
      const sheet = ss.getSheetByName(nomeAba);
      if (sheet) return sheet;
    }
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
      } catch (_) {}
    });
  }

  function criarEventosCalendario_(titulo, descricao, timezone, agendamentos) {
    const createdEvents = [];
    const eventosResponse = [];

    try {
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
          options.guests = item.convidados.join(",");
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
          throw { itemIndex: i, message: err && err.message ? err.message : String(err) };
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
      rollbackEventosCalendario_(createdEvents);
      throw err;
    }
  }

  function obterPlanilha_() {
    const id = (Configuracoes.PLANILHA_AGENDAMENTOS_ID || "").toString().trim();
    if (!id) {
      throw new Error(
        "Defina Configuracoes.PLANILHA_AGENDAMENTOS_ID com o ID da planilha de agendamentos."
      );
    }
    return SpreadsheetApp.openById(id);
  }

  function obterAba_() {
    const nome = (Configuracoes.NOME_ABA_AGENDAMENTOS || "Agendamentos").toString().trim();
    const ss = obterPlanilha_();
    let aba = ss.getSheetByName(nome);
    if (!aba) {
      aba = ss.insertSheet(nome);
    }
    return aba;
  }

  function garantirCabecalho_(aba) {
    const esperado = Configuracoes.CABECALHOS_AGENDAMENTO;
    if (!esperado || !esperado.length) return;
    const lastCol = Math.max(aba.getLastColumn(), esperado.length);
    const primeira = lastCol > 0 ? aba.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    const vazia =
      !primeira.length ||
      primeira.every(c => c === "" || c === null || c === undefined);
    if (vazia) {
      aba.getRange(1, 1, 1, esperado.length).setValues([esperado.slice()]);
      return;
    }
    const h0 = String(primeira[0] || "").trim().toLowerCase();
    const exp0 = String(esperado[0] || "").trim().toLowerCase();
    if (h0 !== exp0) {
      throw new Error(
        "Aba de agendamentos: cabeçalho na linha 1 não confere com o esperado. Verifique a ordem dos títulos."
      );
    }
  }

  /**
   * @param {string[][]} linhas - Cada linha com o mesmo número de colunas (ex.: 12).
   */
  function appendLinhas_(linhas) {
    if (!linhas || !linhas.length) return;
    const aba = obterAba_();
    garantirCabecalho_(aba);
    const numCols = linhas[0].length;
    for (let i = 0; i < linhas.length; i++) {
      const r = linhas[i];
      if (!r || r.length !== numCols) {
        throw new Error(
          "Linha " + (i + 1) + " com número de colunas inválido (esperado " + numCols + ")."
        );
      }
      const targetRow = aba.getLastRow() + 1;
      aba.getRange(targetRow, 1, 1, numCols).setValues([r]);
    }
  }

  return {
    appendEventosNaPlanilha: appendEventosNaPlanilha_,
    rollbackEventosCalendario: rollbackEventosCalendario_,
    criarEventosCalendario: criarEventosCalendario_,
    appendLinhas: appendLinhas_,
    obterAba: obterAba_
  };
})();

function appendEventosNaPlanilha(spreadsheetId, rowsA13) {
  return AgendamentoRepo.appendEventosNaPlanilha(spreadsheetId, rowsA13);
}
function rollbackEventosCalendario(createdEvents) {
  return AgendamentoRepo.rollbackEventosCalendario(createdEvents);
}
function criarEventosCalendario(titulo, descricao, timezone, agendamentos) {
  return AgendamentoRepo.criarEventosCalendario(titulo, descricao, timezone, agendamentos);
}
