/**
 * Camada Repository – Acesso à planilha (leitura/escrita).
 */

/**
 * Retorna a aba configurada.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function _obterAba_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(Configuracoes.NOME_ABA);
  if (!aba) throw new Error("Aba não encontrada: " + Configuracoes.NOME_ABA);
  return aba;
}

/**
 * Lê todos os dados da planilha (cabeçalho + linhas).
 * @returns {{ valores: any[][], temCabecalho: boolean }}
 */
function lerDadosPlanilha() {
  const aba = _obterAba_();
  const lastRow = aba.getLastRow();
  const lastCol = aba.getLastColumn();
  if (lastRow === 0 || lastCol === 0) {
    return { valores: [], temCabecalho: false };
  }
  const valores = aba.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const primeiraLinha = valores[0] || [];
  const temCabecalho = primeiraLinha.some(v =>
    typeof v === "string" && /(curso|turma|status|respons[aá]vel|oferta|sala)/i.test(v)
  );
  return { valores, temCabecalho };
}

/**
 * Insere uma linha no fim da aba.
 * @param {any[]} valores - Array na ordem das colunas (A até AC).
 */
function inserirLinha(valores) {
  const aba = _obterAba_();
  aba.appendRow(valores);
}

/**
 * Retorna o índice (0-based) da coluna ID no cabeçalho ou assume última coluna.
 * @param {any[]} cabecalho - Linha de cabeçalho (array de valores).
 * @param {number} ultimaColuna - Última coluna preenchida (1-based).
 * @returns {number}
 */
function indiceColunaId(cabecalho, ultimaColuna) {
  if (cabecalho && cabecalho.length) {
    const nomeId = (Configuracoes.NOME_COLUNA_ID || "ID").toLowerCase();
    for (let i = 0; i < cabecalho.length; i++) {
      const label = (cabecalho[i] || "").toString().trim().toLowerCase();
      if (label === "id" || label === nomeId) return i;
    }
  }
  return Math.max(0, ultimaColuna - 1);
}

/**
 * Busca uma linha pelo ID. Retorna array de valores ou null.
 * @param {string} id - UUID do registro.
 * @returns {any[]|null}
 */
function buscarLinhaPorId(id) {
  if (!id) return null;
  const { valores, temCabecalho } = lerDadosPlanilha();
  if (!valores.length) return null;
  const cabecalho = valores[0];
  const idxId = indiceColunaId(cabecalho, cabecalho.length);
  const inicio = temCabecalho ? 1 : 0;
  for (let i = inicio; i < valores.length; i++) {
    if (String(valores[i][idxId]) === String(id)) return valores[i];
  }
  return null;
}

/**
 * Busca o índice 1-based da linha que contém o ID.
 * @param {string} id - UUID do registro.
 * @returns {number} Índice da linha (1-based) ou -1.
 */
function buscarIndiceLinhaPorId(id) {
  if (!id) return -1;
  const aba = _obterAba_();
  const lastRow = aba.getLastRow();
  const lastCol = aba.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return -1;
  const valores = aba.getRange(1, 1, lastRow, lastCol).getValues();
  const cabecalho = valores[0] || [];
  const temCabecalho = cabecalho.some(v =>
    typeof v === "string" && /(curso|turma|status|respons[aá]vel|oferta|sala)/i.test(v)
  );
  const idxId = indiceColunaId(cabecalho, lastCol);
  const inicio = temCabecalho ? 1 : 0;
  for (let i = inicio; i < valores.length; i++) {
    if (String(valores[i][idxId]) === String(id)) return i + 1;
  }
  return -1;
}

/**
 * Retorna os valores de uma linha pelo índice 1-based.
 * @param {number} indiceLinha - Número da linha (1-based).
 * @returns {any[]}
 */
function obterLinhaPorIndice(indiceLinha) {
  const aba = _obterAba_();
  const lastCol = aba.getLastColumn();
  return aba.getRange(indiceLinha, 1, 1, lastCol).getValues()[0];
}

/**
 * Atualiza uma linha da planilha pelo índice 1-based.
 * @param {number} indiceLinha - Número da linha (1-based).
 * @param {any[]} valores - Array de valores na ordem das colunas.
 */
function atualizarLinha(indiceLinha, valores) {
  const aba = _obterAba_();
  const lastCol = aba.getLastColumn();
  const range = aba.getRange(indiceLinha, 1, 1, lastCol);
  range.setValues([valores]);
}

/**
 * Retorna o schema padrão de colunas (quando não há cabeçalho na planilha).
 * @returns {{ key: string, label: string }[]}
 */
function obterColunasPadrao() {
  const nomeId = Configuracoes.NOME_COLUNA_ID || "ID";
  return [
    { key: "Curso", label: "Curso" },
    { key: "Turma", label: "Turma" },
    { key: "Inicio", label: "Início" },
    { key: "Fim", label: "Fim" },
    { key: "Sala", label: "Sala" },
    { key: "Oferta", label: "Oferta" },
    { key: "Responsavel", label: "Responsável" },
    { key: "Prioridade", label: "Prioridade" },
    { key: "FimInscricoes", label: "Fim das Inscrições" },
    { key: "Status", label: "Status" },
    { key: "LinkFSA", label: "Link FSA" },
    { key: "TR", label: "TR" },
    { key: "Termo", label: "Termo" },
    { key: "Plataformas", label: "Plataformas" },
    { key: "Coffee", label: "Coffee Break" },
    { key: "AgendarSala", label: "Agendar Sala" },
    { key: "QRInscricao", label: "QR Inscrição" },
    { key: "LinkInscricao", label: "Link Inscrição" },
    { key: "QRFrequencia", label: "QR Frequência" },
    { key: "Frequencia", label: "Frequência" },
    { key: "CodVerificacao", label: "Cód. Verificação" },
    { key: "QRAvaliacao", label: "QR Avaliação" },
    { key: "AvaliacaoReacao", label: "Avaliação Reação" },
    { key: "PrepDivulgacao", label: "Prep. Divulgação" },
    { key: "Divulgacao", label: "Divulgação" },
    { key: "LinkPagina", label: "Link Página" },
    { key: "DataCadastro", label: "Data Cadastro" },
    { key: "EmailUsuario", label: "E-mail Usuário" },
    { key: "ID", label: nomeId }
  ];
}

/**
 * Preenche a coluna ID para registros que ainda não têm (migração).
 * @returns {{ ok: boolean, message: string }}
 */
function preencherColunaId() {
  const aba = _obterAba_();
  const lastRow = aba.getLastRow();
  const lastCol = aba.getLastColumn();
  if (lastRow === 0) return { ok: true, message: "Planilha vazia; nada para migrar." };

  const firstRow = lastCol > 0 ? aba.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  const hasHeader = firstRow.some(v =>
    typeof v === "string" && /(curso|turma|status|respons[aá]vel|oferta|sala)/i.test(v)
  );

  const uuidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
  if (!hasHeader && lastCol > 0 && lastRow >= 5) {
    const sampleCount = Math.min(50, lastRow);
    const sample = aba.getRange(1, lastCol, sampleCount, 1).getValues().flat();
    const uuidCount = sample.filter(v => typeof v === "string" && uuidRegex.test(v)).length;
    if (uuidCount / sampleCount >= 0.8) {
      return { ok: true, message: "IDs parecem já existir (última coluna). Nada a fazer." };
    }
  }

  const idCol = lastCol + 1;
  let startRow = 1;
  let rowsToFill = lastRow;
  if (hasHeader) {
    aba.getRange(1, idCol).setValue(Configuracoes.NOME_COLUNA_ID || "ID");
    startRow = 2;
    rowsToFill = Math.max(0, lastRow - 1);
  }

  if (rowsToFill === 0) return { ok: true, message: "Nada para preencher." };

  const range = aba.getRange(startRow, idCol, rowsToFill, 1);
  const values = range.getValues();
  let filled = 0;
  for (let i = 0; i < values.length; i++) {
    if (!values[i][0]) {
      values[i][0] = Utilities.getUuid();
      filled++;
    }
  }
  range.setValues(values);
  return { ok: true, message: "Migração concluída. IDs preenchidos: " + filled + ". Coluna: " + idCol + "." };
}
