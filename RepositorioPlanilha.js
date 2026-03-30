/**
 * Camada Repository – Acesso à planilha (leitura/escrita).
 *
 * Observação: em Apps Script, toda função global fica chamável via `google.script.run`.
 * Para reduzir o "surface area", a implementação foi encapsulada em `RegistroRepo` e
 * as funções globais abaixo são apenas wrappers (compatibilidade / legibilidade).
 */

const RegistroRepo = (() => {
  const HEADER_HINT_REGEX = /(curso|turma|status|respons[aá]vel|oferta)/i;
  const UUID_REGEX = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

  function obterAba_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(Configuracoes.NOME_ABA);
    if (!aba) throw new Error("Aba não encontrada: " + Configuracoes.NOME_ABA);
    return aba;
  }

  function temCabecalho_(primeiraLinha) {
    const row = primeiraLinha || [];
    return row.some(v => typeof v === "string" && HEADER_HINT_REGEX.test(v));
  }

  function indiceColunaId_(cabecalho, ultimaColuna1Based) {
    if (cabecalho && cabecalho.length) {
      const nomeId = (Configuracoes.NOME_COLUNA_ID || "ID").toLowerCase();
      for (let i = 0; i < cabecalho.length; i++) {
        const label = (cabecalho[i] || "").toString().trim().toLowerCase();
        if (label === "id" || label === nomeId) return i;
      }
    }
    return Math.max(0, (ultimaColuna1Based || 1) - 1);
  }

  function lerDadosPlanilha_() {
    const aba = obterAba_();
    const lastRow = aba.getLastRow();
    const lastCol = aba.getLastColumn();
    if (lastRow === 0 || lastCol === 0) return { valores: [], temCabecalho: false };
    const valores = aba.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    return { valores, temCabecalho: temCabecalho_(valores[0]) };
  }

  function buscarIndiceLinhaPorId_(id) {
    if (!id) return -1;
    const aba = obterAba_();
    const lastRow = aba.getLastRow();
    const lastCol = aba.getLastColumn();
    if (lastRow < 1 || lastCol < 1) return -1;

    const valores = aba.getRange(1, 1, lastRow, lastCol).getValues();
    const cabecalho = valores[0] || [];
    const hasHeader = temCabecalho_(cabecalho);
    const idxId = indiceColunaId_(cabecalho, lastCol);
    const inicio = hasHeader ? 1 : 0;

    for (let i = inicio; i < valores.length; i++) {
      if (String(valores[i][idxId]) === String(id)) return i + 1; // 1-based
    }
    return -1;
  }

  function obterLinhaPorIndice_(indiceLinha1Based) {
    const aba = obterAba_();
    const lastCol = aba.getLastColumn();
    return aba.getRange(indiceLinha1Based, 1, 1, lastCol).getValues()[0];
  }

  function atualizarLinha_(indiceLinha1Based, valores) {
    const aba = obterAba_();
    const lastCol = aba.getLastColumn();
    aba.getRange(indiceLinha1Based, 1, 1, lastCol).setValues([valores]);
  }

  function removerLinha_(indiceLinha1Based) {
    const aba = obterAba_();
    aba.deleteRow(indiceLinha1Based);
  }

  function inserirLinha_(valores) {
    const aba = obterAba_();
    aba.appendRow(valores);
  }

  function obterColunasPadrao_() {
    const nomeId = Configuracoes.NOME_COLUNA_ID || "ID";
    return [
      { key: "Curso", label: "Curso" },
      { key: "Turma", label: "Turma" },
      { key: "Inicio", label: "Início" },
      { key: "Fim", label: "Fim" },
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
      { key: "QRInscricao", label: "QR Code Inscrição" },
      { key: "LinkInscricao", label: "Link Inscrição" },
      { key: "QRFrequencia", label: "QR Code Frequência" },
      { key: "Frequencia", label: "Frequência" },
      { key: "CodVerificacao", label: "Cód. Verificação" },
      { key: "QRAvaliacao", label: "QR Code Avaliação" },
      { key: "AvaliacaoReacao", label: "Avaliação Reação" },
      { key: "PrepDivulgacao", label: "Prep. Divulgação" },
      { key: "Divulgacao", label: "Divulgação" },
      { key: "LinkPagina", label: "Link Página" },
      { key: "DataCadastro", label: "Data Cadastro" },
      { key: "EmailUsuario", label: "E-mail Usuário" },
      { key: "ID", label: nomeId }
    ];
  }

  function indiceColunaPorLabel_(cabecalho, label) {
    if (!cabecalho || !cabecalho.length || !label) return -1;
    const want = String(label).trim().toLowerCase();
    for (let i = 0; i < cabecalho.length; i++) {
      const h = (cabecalho[i] || "").toString().trim().toLowerCase();
      if (h === want) return i;
    }
    return -1;
  }

  function existeTurmaParaCurso_(curso, turma, excluirId) {
    const { valores, temCabecalho } = lerDadosPlanilha_();
    if (!valores.length) return false;
    const cabecalho = valores[0] || [];
    const idxCurso = indiceColunaPorLabel_(cabecalho, "Curso");
    const idxTurma = indiceColunaPorLabel_(cabecalho, "Turma");
    const idxId = indiceColunaId_(cabecalho, cabecalho.length);
    if (idxCurso < 0 || idxTurma < 0) return false;
    const inicio = temCabecalho ? 1 : 0;
    const cursoNorm = String(curso || "").trim();
    const turmaNorm = String(turma || "").trim();
    for (let i = inicio; i < valores.length; i++) {
      const row = valores[i];
      if (excluirId && String(row[idxId] || "").trim() === String(excluirId || "").trim()) continue;
      if (String(row[idxCurso] || "").trim() === cursoNorm && String(row[idxTurma] || "").trim() === turmaNorm) return true;
    }
    return false;
  }

  function buscarLinhaPorId_(id) {
    if (!id) return null;
    const { valores, temCabecalho } = lerDadosPlanilha_();
    if (!valores.length) return null;
    const cabecalho = valores[0] || [];
    const idxId = indiceColunaId_(cabecalho, cabecalho.length);
    const inicio = temCabecalho ? 1 : 0;
    for (let i = inicio; i < valores.length; i++) {
      if (String(valores[i][idxId]) === String(id)) return valores[i];
    }
    return null;
  }

  function listarCursosDistintos_() {
    const { valores, temCabecalho } = lerDadosPlanilha_();
    if (!valores.length) return [];
    const cabecalho = valores[0] || [];
    const idxCurso = indiceColunaPorLabel_(cabecalho, "Curso");
    if (idxCurso < 0) return [];
    const inicio = temCabecalho ? 1 : 0;
    const set = {};
    for (let i = inicio; i < valores.length; i++) {
      const c = String(valores[i][idxCurso] || "").trim();
      if (c) set[c] = true;
    }
    return Object.keys(set).sort((a, b) => a.localeCompare(b, "pt-BR"));
  }

  function listarResponsaveisDistintos_() {
    const { valores, temCabecalho } = lerDadosPlanilha_();
    if (!valores.length) return [];
    const cabecalho = valores[0] || [];
    const idx = indiceColunaPorLabel_(cabecalho, "Responsável");
    if (idx < 0) return [];
    const inicio = temCabecalho ? 1 : 0;
    const set = {};
    for (let i = inicio; i < valores.length; i++) {
      const r = String(valores[i][idx] || "").trim();
      if (r) set[r] = true;
    }
    return Object.keys(set).sort((a, b) => a.localeCompare(b, "pt-BR"));
  }

  function listarTurmasDistintasPorCurso_(curso) {
    const cursoNorm = String(curso || "").trim();
    if (!cursoNorm) return [];
    const { valores, temCabecalho } = lerDadosPlanilha_();
    if (!valores.length) return [];
    const cabecalho = valores[0] || [];
    const idxCurso = indiceColunaPorLabel_(cabecalho, "Curso");
    const idxTurma = indiceColunaPorLabel_(cabecalho, "Turma");
    if (idxCurso < 0 || idxTurma < 0) return [];
    const inicio = temCabecalho ? 1 : 0;
    const set = {};
    for (let i = inicio; i < valores.length; i++) {
      const row = valores[i];
      if (String(row[idxCurso] || "").trim() !== cursoNorm) continue;
      const t = String(row[idxTurma] || "").trim();
      if (t) set[t] = true;
    }
    return Object.keys(set).sort((a, b) => a.localeCompare(b, "pt-BR"));
  }

  /**
   * ID (UUID) da linha cujo Curso+Turma coincide. Assume dupla única na planilha.
   * @returns {string|null}
   */
  function buscarIdPorCursoTurma_(curso, turma) {
    const cursoNorm = String(curso || "").trim();
    const turmaNorm = String(turma || "").trim();
    if (!cursoNorm || !turmaNorm) return null;
    const { valores, temCabecalho } = lerDadosPlanilha_();
    if (!valores.length) return null;
    const cabecalho = valores[0] || [];
    const idxCurso = indiceColunaPorLabel_(cabecalho, "Curso");
    const idxTurma = indiceColunaPorLabel_(cabecalho, "Turma");
    const idxId = indiceColunaId_(cabecalho, cabecalho.length);
    if (idxCurso < 0 || idxTurma < 0) return null;
    const inicio = temCabecalho ? 1 : 0;
    for (let i = inicio; i < valores.length; i++) {
      const row = valores[i];
      if (String(row[idxCurso] || "").trim() !== cursoNorm) continue;
      if (String(row[idxTurma] || "").trim() !== turmaNorm) continue;
      const idVal = String(row[idxId] || "").trim();
      return idVal || null;
    }
    return null;
  }

  /**
   * Valores brutos das células Início/Fim da linha Curso+Turma (getValues — Date ou texto).
   * @returns {{ inicioVal: *, fimVal: * }|null}
   */
  function buscarVigenciaInicioFimPorCursoTurma_(curso, turma) {
    const cursoNorm = String(curso || "").trim();
    const turmaNorm = String(turma || "").trim();
    if (!cursoNorm || !turmaNorm) return null;
    const aba = obterAba_();
    const lastRow = aba.getLastRow();
    const lastCol = aba.getLastColumn();
    if (lastRow < 1 || lastCol < 1) return null;
    const valores = aba.getRange(1, 1, lastRow, lastCol).getValues();
    const cabecalho = valores[0] || [];
    const temCab = temCabecalho_(cabecalho);
    const idxCurso = indiceColunaPorLabel_(cabecalho, "Curso");
    const idxTurma = indiceColunaPorLabel_(cabecalho, "Turma");
    let idxInicio = indiceColunaPorLabel_(cabecalho, "Início");
    if (idxInicio < 0) idxInicio = indiceColunaPorLabel_(cabecalho, "Inicio");
    const idxFim = indiceColunaPorLabel_(cabecalho, "Fim");
    if (idxCurso < 0 || idxTurma < 0 || idxInicio < 0 || idxFim < 0) return null;
    const inicio = temCab ? 1 : 0;
    for (let i = inicio; i < valores.length; i++) {
      const row = valores[i];
      if (String(row[idxCurso] || "").trim() !== cursoNorm) continue;
      if (String(row[idxTurma] || "").trim() !== turmaNorm) continue;
      return { inicioVal: row[idxInicio], fimVal: row[idxFim] };
    }
    return null;
  }

  function preencherColunaId_() {
    const aba = obterAba_();
    const lastRow = aba.getLastRow();
    const lastCol = aba.getLastColumn();
    if (lastRow === 0) return { ok: true, message: "Planilha vazia; nada para migrar." };

    const firstRow = lastCol > 0 ? aba.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    const hasHeader = temCabecalho_(firstRow);

    if (!hasHeader && lastCol > 0 && lastRow >= 5) {
      const sampleCount = Math.min(50, lastRow);
      const sample = aba.getRange(1, lastCol, sampleCount, 1).getValues().flat();
      const uuidCount = sample.filter(v => typeof v === "string" && UUID_REGEX.test(v)).length;
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

  return {
    lerDadosPlanilha: lerDadosPlanilha_,
    inserirLinha: inserirLinha_,
    obterColunasPadrao: obterColunasPadrao_,
    buscarLinhaPorId: buscarLinhaPorId_,
    buscarIndiceLinhaPorId: buscarIndiceLinhaPorId_,
    obterLinhaPorIndice: obterLinhaPorIndice_,
    atualizarLinha: atualizarLinha_,
    removerLinha: removerLinha_,
    preencherColunaId: preencherColunaId_,
    indiceColunaId: indiceColunaId_,
    existeTurmaParaCurso: existeTurmaParaCurso_,
    listarCursosDistintos: listarCursosDistintos_,
    listarResponsaveisDistintos: listarResponsaveisDistintos_,
    listarTurmasDistintasPorCurso: listarTurmasDistintasPorCurso_,
    buscarIdPorCursoTurma: buscarIdPorCursoTurma_,
    buscarVigenciaInicioFimPorCursoTurma: buscarVigenciaInicioFimPorCursoTurma_
  };
})();

// ---- Wrappers globais (compatibilidade) ----
function lerDadosPlanilha() { return RegistroRepo.lerDadosPlanilha(); }
function inserirLinha(valores) { return RegistroRepo.inserirLinha(valores); }
function obterColunasPadrao() { return RegistroRepo.obterColunasPadrao(); }
function buscarLinhaPorId(id) { return RegistroRepo.buscarLinhaPorId(id); }
function buscarIndiceLinhaPorId(id) { return RegistroRepo.buscarIndiceLinhaPorId(id); }
function obterLinhaPorIndice(indiceLinha) { return RegistroRepo.obterLinhaPorIndice(indiceLinha); }
function atualizarLinha(indiceLinha, valores) { return RegistroRepo.atualizarLinha(indiceLinha, valores); }
function removerLinha(indiceLinha) { return RegistroRepo.removerLinha(indiceLinha); }
function preencherColunaId() { return RegistroRepo.preencherColunaId(); }
function indiceColunaId(cabecalho, ultimaColuna) { return RegistroRepo.indiceColunaId(cabecalho, ultimaColuna); }
function existeTurmaParaCurso(curso, turma, excluirId) { return RegistroRepo.existeTurmaParaCurso(curso, turma, excluirId); }
function listarCursosDistintosPlanilha() { return RegistroRepo.listarCursosDistintos(); }
function listarResponsaveisDistintosPlanilha() { return RegistroRepo.listarResponsaveisDistintos(); }
function listarTurmasDistintasPorCursoPlanilha(curso) { return RegistroRepo.listarTurmasDistintasPorCurso(curso); }
function buscarIdRegistroPorCursoTurma(curso, turma) { return RegistroRepo.buscarIdPorCursoTurma(curso, turma); }
