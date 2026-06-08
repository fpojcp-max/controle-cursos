/**
 * Camada Repository – Agendamento (aba Agendamentos na planilha associada ao script).
 */

const AgendamentoRepo = (() => {
  const COL_AG_ = {
    TURMA: 0,
    CURSO: 1,
    DATA: 2,
    NOME_SALA: 3,
    HORA_INI: 4,
    HORA_FIM: 5,
    CRIADO_EM: 6,
    CRIADO_POR: 7,
    ID_AGENDAMENTO: 8,
    ID_REGISTRO_TURMA: 9
  };

  function obterPlanilha_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error("Não há planilha associada a este projeto Apps Script.");
    }
    return ss;
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
      primeira.every((c) => c === "" || c === null || c === undefined);
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
    const n = Math.max(row.length, COL_AG_.ID_REGISTRO_TURMA + 1);
    for (let c = 0; c < n; c++) {
      out.push(normalizarCelulaAg_(row[c], c));
    }
    return out;
  }

  function lerTodasLinhasAg_(numCols) {
    const aba = obterAba_();
    garantirCabecalho_(aba);
    const lastRow = aba.getLastRow();
    if (lastRow < 2) return [];
    const raw = aba.getRange(2, 1, lastRow, numCols).getValues();
    const out = [];
    for (let i = 0; i < raw.length; i++) {
      const linha = normalizarLinhaAg_(raw[i]);
      while (linha.length < numCols) linha.push("");
      const idAg = String(linha[COL_AG_.ID_AGENDAMENTO] || "").trim();
      if (!idAg) continue;
      out.push({
        sheetRow: i + 2,
        cells: linha.slice(0, numCols),
        eventId: idAg
      });
    }
    return out;
  }

  function coletarLinhasAgendamentoPorIdTurma_(idTurma) {
    const idNorm = String(idTurma || "").trim();
    if (!idNorm) return [];
    const numCols = (Configuracoes.CABECALHOS_AGENDAMENTO || []).length;
    return lerTodasLinhasAg_(numCols).filter(function (m) {
      return String(m.cells[COL_AG_.ID_REGISTRO_TURMA] || "").trim() === idNorm;
    });
  }

  function coletarTodasLinhasAgendamentoPorIdTurma_(idTurma) {
    return coletarLinhasAgendamentoPorIdTurma_(idTurma);
  }

  function listarTodasLinhasAgendamento_() {
    const numCols = (Configuracoes.CABECALHOS_AGENDAMENTO || []).length;
    return lerTodasLinhasAg_(numCols);
  }

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
        return String(a.eventId).localeCompare(String(b.eventId)) * dir;
      });
      return;
    }
    const col =
      typeof sortCol === "number" && sortCol >= 0 && sortCol < numCols ? sortCol : COL_AG_.DATA;
    matches.sort((a, b) => {
      const sa = String(a.cells[col] != null ? a.cells[col] : "").trim();
      const sb = String(b.cells[col] != null ? b.cells[col] : "").trim();
      let cmp = sa.localeCompare(sb, undefined, { numeric: true, sensitivity: "base" });
      if (cmp !== 0) return cmp * dir;
      return String(a.eventId).localeCompare(String(b.eventId)) * dir;
    });
  }

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

  function listarLinhasAgendamentoPorIdTurma_(idTurma) {
    const m = coletarLinhasAgendamentoPorIdTurma_(idTurma);
    ordenarLinhasAgendamento_(m, -1, true);
    return m;
  }

  function listarLinhasAgendamentoPorIdTurmaCompleto_(idTurma) {
    const m = coletarTodasLinhasAgendamentoPorIdTurma_(idTurma);
    ordenarLinhasAgendamento_(m, -1, true);
    return m;
  }

  function listarAgendamentosPaginadoPorIdTurma_(idTurma, offset, limit, sortCol, sortAsc) {
    const all = coletarLinhasAgendamentoPorIdTurma_(idTurma);
    const sc =
      sortCol === -1 || sortCol === null || sortCol === undefined ? -1 : parseInt(sortCol, 10);
    ordenarLinhasAgendamento_(all, isNaN(sc) ? -1 : sc, sortAsc !== false);
    const total = all.length;
    const off = Math.max(0, parseInt(offset, 10) || 0);
    const lim = Math.max(1, parseInt(limit, 10) || 50);
    const slice = all.slice(off, off + lim);
    return {
      cabecalho: (Configuracoes.CABECALHOS_AGENDAMENTO || []).slice(),
      total: total,
      itens: slice,
      allLinhas: all.map(function (m) {
        return { sheetRow: m.sheetRow, eventId: String(m.eventId || "").trim() };
      })
    };
  }

  function listarTodosEventIdsPorIdTurma_(idTurma) {
    return coletarLinhasAgendamentoPorIdTurma_(idTurma).map((m) => m.eventId);
  }

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

  function obterLinhaAgPorSheetRow_(sheetRow1Based) {
    const r = parseInt(sheetRow1Based, 10);
    if (isNaN(r) || r < 2) return null;
    const aba = obterAba_();
    garantirCabecalho_(aba);
    const numCols = (Configuracoes.CABECALHOS_AGENDAMENTO || []).length;
    if (numCols < 1) return null;
    const lastRow = aba.getLastRow();
    if (r > lastRow) return null;
    const row = aba.getRange(r, 1, 1, numCols).getValues()[0];
    const linha = normalizarLinhaAg_(row);
    while (linha.length < numCols) linha.push("");
    const eventId = String(linha[COL_AG_.ID_AGENDAMENTO] || "").trim();
    return { sheetRow: r, cells: linha.slice(0, numCols), eventId: eventId };
  }

  function atualizarLinhaCompletaAg_(sheetRow1Based, cells) {
    const r = parseInt(sheetRow1Based, 10);
    if (isNaN(r) || r < 2) throw new Error("Linha da planilha inválida.");
    if (!cells || !cells.length) throw new Error("Dados da linha inválidos.");
    const aba = obterAba_();
    garantirCabecalho_(aba);
    aba.getRange(r, 1, 1, cells.length).setValues([cells]);
  }

  return {
    appendLinhas: appendLinhas_,
    obterAba: obterAba_,
    listarTodasLinhasAgendamento: listarTodasLinhasAgendamento_,
    listarAgendamentosPaginadoPorIdTurma: listarAgendamentosPaginadoPorIdTurma_,
    listarLinhasAgendamentoPorIdTurma: listarLinhasAgendamentoPorIdTurma_,
    listarLinhasAgendamentoPorIdTurmaCompleto: listarLinhasAgendamentoPorIdTurmaCompleto_,
    listarTodosEventIdsPorIdTurma: listarTodosEventIdsPorIdTurma_,
    excluirLinhasPorNumeros: excluirLinhasPorNumeros_,
    obterLinhaAgPorSheetRow: obterLinhaAgPorSheetRow_,
    atualizarLinhaCompletaAg: atualizarLinhaCompletaAg_,
    COL_AG: COL_AG_
  };
})();
