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

  const COL_AG_ = {
    TURMA: 0,
    CURSO: 1,
    DATA: 2,
    NOME_SALA: 3,
    HORA_INI: 4,
    HORA_FIM: 5,
    CONVIDADOS: 6,
    CRIADO_EM: 7,
    CRIADO_POR: 8,
    ID_GOOGLE: 9,
    ID_REGISTRO_TURMA: 10,
    ID_SALA: 11
  };

  function tzAg_() {
    return Configuracoes.TIMEZONE_AGENDAMENTO || "America/Sao_Paulo";
  }

  function normalizarCelulaAg_(cell, colIdx) {
    if (cell === null || cell === undefined) return "";
    if (Object.prototype.toString.call(cell) === "[object Date]" && !isNaN(cell.getTime())) {
      if (colIdx === COL_AG_.DATA) {
        return Utilities.formatDate(cell, tzAg_(), "yyyy-MM-dd");
      }
      if (colIdx === COL_AG_.HORA_INI || colIdx === COL_AG_.HORA_FIM) {
        return Utilities.formatDate(cell, tzAg_(), "HH:mm");
      }
      if (colIdx === COL_AG_.CRIADO_EM) {
        return Utilities.formatDate(cell, tzAg_(), "yyyy-MM-dd HH:mm");
      }
      return Utilities.formatDate(cell, tzAg_(), "yyyy-MM-dd");
    }
    return String(cell).trim();
  }

  function normalizarLinhaAg_(row) {
    const out = [];
    const n = Math.max(row.length, COL_AG_.ID_SALA + 1);
    for (let c = 0; c < n; c++) {
      out.push(normalizarCelulaAg_(row[c], c));
    }
    return out;
  }

  /**
   * Linhas da aba (filtro coluna ID Turma), sem ordenação estável explícita.
   * @param {string} idTurma
   * @returns {{ sheetRow: number, cells: string[], eventId: string }[]}
   */
  function coletarLinhasAgendamentoPorIdTurma_(idTurma) {
    const idNorm = String(idTurma || "").trim();
    if (!idNorm) return [];
    const aba = obterAba_();
    garantirCabecalho_(aba);
    const esperado = Configuracoes.CABECALHOS_AGENDAMENTO || [];
    const numCols = esperado.length;
    const lastRow = aba.getLastRow();
    if (lastRow < 2) return [];
    const raw = aba.getRange(2, 1, lastRow, numCols).getValues();
    const matches = [];
    for (let i = 0; i < raw.length; i++) {
      const row = raw[i];
      const linha = normalizarLinhaAg_(row);
      while (linha.length < numCols) linha.push("");
      const idCell = String(linha[COL_AG_.ID_REGISTRO_TURMA] || "").trim();
      if (idCell !== idNorm) continue;
      const eventId = String(linha[COL_AG_.ID_GOOGLE] || "").trim();
      if (!eventId) continue;
      matches.push({
        sheetRow: i + 2,
        cells: linha.slice(0, numCols),
        eventId: eventId
      });
    }
    return matches;
  }

  /**
   * Todas as linhas com o ID da turma, inclusive sem "ID Agendamento" (só planilha).
   * @param {string} idTurma
   * @returns {{ sheetRow: number, cells: string[], eventId: string }[]}
   */
  function coletarTodasLinhasAgendamentoPorIdTurma_(idTurma) {
    const idNorm = String(idTurma || "").trim();
    if (!idNorm) return [];
    const aba = obterAba_();
    garantirCabecalho_(aba);
    const esperado = Configuracoes.CABECALHOS_AGENDAMENTO || [];
    const numCols = esperado.length;
    const lastRow = aba.getLastRow();
    if (lastRow < 2) return [];
    const raw = aba.getRange(2, 1, lastRow, numCols).getValues();
    const matches = [];
    for (let i = 0; i < raw.length; i++) {
      const row = raw[i];
      const linha = normalizarLinhaAg_(row);
      while (linha.length < numCols) linha.push("");
      const idCell = String(linha[COL_AG_.ID_REGISTRO_TURMA] || "").trim();
      if (idCell !== idNorm) continue;
      const ev = String(linha[COL_AG_.ID_GOOGLE] || "").trim();
      matches.push({
        sheetRow: i + 2,
        cells: linha.slice(0, numCols),
        eventId: ev
      });
    }
    return matches;
  }

  function listarLinhasAgendamentoPorIdTurmaCompleto_(idTurma) {
    const m = coletarTodasLinhasAgendamentoPorIdTurma_(idTurma);
    ordenarLinhasAgendamento_(m, -1, true);
    return m;
  }

  /**
   * Atualiza só a coluna "ID Agendamento" (1-based row, cabeçalho na linha 1).
   * @param {number} sheetRow1Based
   * @param {string} novoEventId
   */
  function atualizarIdGoogleNaLinha_(sheetRow1Based, novoEventId) {
    const r = parseInt(sheetRow1Based, 10);
    if (isNaN(r) || r < 2) {
      throw new Error("Linha da planilha inválida.");
    }
    const id = String(novoEventId || "").trim();
    if (!id) throw new Error("ID do evento inválido.");
    const aba = obterAba_();
    garantirCabecalho_(aba);
    const col = COL_AG_.ID_GOOGLE + 1;
    aba.getRange(r, col).setValue(id);
  }

  /**
   * sortCol = -1 → padrão: Data asc, Hora início asc, eventId.
   * Caso contrário: coluna 0..n-1, string localeCompare com desempate por eventId.
   * @param {{ sheetRow: number, cells: string[], eventId: string }[]} matches
   * @param {number} sortCol
   * @param {boolean} sortAsc
   */
  function ordenarLinhasAgendamento_(matches, sortCol, sortAsc) {
    if (!matches || !matches.length) return;
    const numCols = (Configuracoes.CABECALHOS_AGENDAMENTO || []).length;
    const asc = sortAsc !== false;
    const dir = asc ? 1 : -1;
    if (sortCol === -1 || sortCol === null || sortCol === undefined) {
      matches.sort((a, b) => {
        const da = a.cells[COL_AG_.DATA] || "";
        const db = b.cells[COL_AG_.DATA] || "";
        if (da !== db) return (da < db ? -1 : da > db ? 1 : 0) * dir;
        const ha = a.cells[COL_AG_.HORA_INI] || "";
        const hb = b.cells[COL_AG_.HORA_INI] || "";
        if (ha !== hb) return (ha < hb ? -1 : ha > hb ? 1 : 0) * dir;
        const ie = String(a.eventId).localeCompare(String(b.eventId));
        return ie * dir;
      });
      return;
    }
    const col =
      typeof sortCol === "number" && sortCol >= 0 && sortCol < numCols
        ? sortCol
        : COL_AG_.DATA;
    matches.sort((a, b) => {
      const sa = String(a.cells[col] != null ? a.cells[col] : "").trim();
      const sb = String(b.cells[col] != null ? b.cells[col] : "").trim();
      let cmp = sa.localeCompare(sb, undefined, { numeric: true, sensitivity: "base" });
      if (cmp !== 0) return cmp * dir;
      cmp = String(a.eventId).localeCompare(String(b.eventId));
      return cmp * dir;
    });
  }

  /**
   * @param {string} idTurma
   * @returns {{ sheetRow: number, cells: string[], eventId: string }[]}
   */
  function listarLinhasAgendamentoPorIdTurma_(idTurma) {
    const m = coletarLinhasAgendamentoPorIdTurma_(idTurma);
    ordenarLinhasAgendamento_(m, -1, true);
    return m;
  }

  /**
   * @param {string} idTurma
   * @param {number} offset - base 0
   * @param {number} limit
   * @param {number} [sortCol] -1 = Data + Hora início
   * @param {boolean} [sortAsc]
   */
  function listarAgendamentosPaginadoPorIdTurma_(idTurma, offset, limit, sortCol, sortAsc) {
    const all = coletarLinhasAgendamentoPorIdTurma_(idTurma);
    const sc =
      sortCol === -1 || sortCol === null || sortCol === undefined
        ? -1
        : parseInt(sortCol, 10);
    const asc = sortAsc !== false;
    ordenarLinhasAgendamento_(all, isNaN(sc) ? -1 : sc, asc);
    const total = all.length;
    const off = Math.max(0, parseInt(offset, 10) || 0);
    const lim = Math.max(1, parseInt(limit, 10) || 50);
    const slice = all.slice(off, off + lim);
    return {
      cabecalho: (Configuracoes.CABECALHOS_AGENDAMENTO || []).slice(),
      total: total,
      itens: slice,
      /**
       * Uma entrada por linha da planilha (após ordenação). sheetRow é único; eventId pode repetir entre linhas.
       * O cliente deve selecionar por sheetRow, não só por eventId.
       */
      allLinhas: all.map((m) => ({
        sheetRow: m.sheetRow,
        eventId: String(m.eventId != null ? m.eventId : "").trim()
      }))
    };
  }

  function listarTodosEventIdsPorIdTurma_(idTurma) {
    return coletarLinhasAgendamentoPorIdTurma_(idTurma).map((m) => m.eventId);
  }

  /**
   * Remove linhas pela ordem de sheetRow (deve vir em ordem decrescente).
   * @param {number[]} sheetRowsDesc
   */
  function excluirLinhasPorNumeros_(sheetRowsDesc) {
    if (!sheetRowsDesc || !sheetRowsDesc.length) return;
    const aba = obterAba_();
    const sorted = sheetRowsDesc.slice().sort((a, b) => b - a);
    for (let i = 0; i < sorted.length; i++) {
      const r = sorted[i];
      if (typeof r !== "number" || r < 2) continue;
      aba.deleteRow(r);
    }
  }

  /**
   * Uma linha da aba (1-based, linha 1 = cabeçalho).
   * @param {number} sheetRow1Based
   * @returns {{ sheetRow: number, cells: string[], eventId: string } | null}
   */
  function obterLinhaAgPorSheetRow_(sheetRow1Based) {
    const r = parseInt(sheetRow1Based, 10);
    if (isNaN(r) || r < 2) return null;
    const aba = obterAba_();
    garantirCabecalho_(aba);
    const esperado = Configuracoes.CABECALHOS_AGENDAMENTO || [];
    const numCols = esperado.length;
    if (numCols < 1) return null;
    const lastRow = aba.getLastRow();
    if (r > lastRow) return null;
    const row = aba.getRange(r, 1, 1, numCols).getValues()[0];
    const linha = normalizarLinhaAg_(row);
    while (linha.length < numCols) linha.push("");
    const eventId = String(linha[COL_AG_.ID_GOOGLE] || "").trim();
    return {
      sheetRow: r,
      cells: linha.slice(0, numCols),
      eventId: eventId
    };
  }

  /**
   * Sobrescreve colunas da linha (mesma largura que cells).
   * @param {number} sheetRow1Based
   * @param {string[]} cells
   */
  function atualizarLinhaCompletaAg_(sheetRow1Based, cells) {
    const r = parseInt(sheetRow1Based, 10);
    if (isNaN(r) || r < 2) {
      throw new Error("Linha da planilha inválida.");
    }
    if (!cells || !cells.length) {
      throw new Error("Dados da linha inválidos.");
    }
    const aba = obterAba_();
    garantirCabecalho_(aba);
    const numCols = cells.length;
    aba.getRange(r, 1, 1, numCols).setValues([cells]);
  }

  return {
    appendEventosNaPlanilha: appendEventosNaPlanilha_,
    rollbackEventosCalendario: rollbackEventosCalendario_,
    criarEventosCalendario: criarEventosCalendario_,
    appendLinhas: appendLinhas_,
    obterAba: obterAba_,
    listarAgendamentosPaginadoPorIdTurma: listarAgendamentosPaginadoPorIdTurma_,
    listarLinhasAgendamentoPorIdTurma: listarLinhasAgendamentoPorIdTurma_,
    listarLinhasAgendamentoPorIdTurmaCompleto: listarLinhasAgendamentoPorIdTurmaCompleto_,
    listarTodosEventIdsPorIdTurma: listarTodosEventIdsPorIdTurma_,
    excluirLinhasPorNumeros: excluirLinhasPorNumeros_,
    atualizarIdGoogleNaLinha: atualizarIdGoogleNaLinha_,
    obterLinhaAgPorSheetRow: obterLinhaAgPorSheetRow_,
    atualizarLinhaCompletaAg: atualizarLinhaCompletaAg_,
    COL_AG: COL_AG_
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
