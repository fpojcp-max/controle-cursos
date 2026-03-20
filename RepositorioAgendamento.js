/**
 * Camada Repository – Planilha de ocorrências de agendamento (uma linha por evento).
 * Cabeçalhos fixos (linha 1): conforme contrato em Configuracoes.CABECALHOS_AGENDAMENTO.
 */

const AgendamentoRepo = (() => {
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
      // getRange(linha, coluna, numLinhas, numColunas) — o 3º parâmetro é QUANTIDADE de linhas, não a linha final.
      aba.getRange(targetRow, 1, 1, numCols).setValues([r]);
    }
  }

  return {
    appendLinhas: appendLinhas_,
    obterAba: obterAba_
  };
})();
