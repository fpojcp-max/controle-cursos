function doGet(e) {
  const view = (e && e.parameter && e.parameter.view) ? String(e.parameter.view) : "consulta";
  const file = view === "cadastro" ? "Index" : "Consulta";

  return HtmlService.createTemplateFromFile(file)
    .evaluate()
    .setTitle('Sistema de Gestão de Cursos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getWebAppUrl(view) {
  const base = ScriptApp.getService().getUrl();
  if (!view) return base;
  return `${base}?view=${encodeURIComponent(String(view))}`;
}

function salvarChamado(dados) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(SETTINGS.NOME_ABA);
    const emailUsuario = Session.getActiveUser().getEmail();
    const dataAtual = new Date();
    const id = Utilities.getUuid();

    aba.appendRow([
      dados.curso,           // A
      dados.turma,           // B
      dados.inicio,          // C
      dados.fim,             // D
      dados.sala,            // E
      dados.oferta,          // F
      dados.responsavel,     // G
      dados.prioridade,      // H
      dados.fimInscricoes,   // I
      dados.status,          // J
      dados.linkFSA,         // K
      dados.tr,              // L
      dados.termo,           // M
      dados.plataformas,     // N
      dados.coffee,          // O
      dados.agendarSala,     // P
      dados.qrInscricao,     // Q
      dados.linkInscricao,   // R
      dados.qrFrequencia,    // S
      dados.frequencia,      // T
      dados.codVerificacao,  // U
      dados.qrAvaliacao,     // V
      dados.avaliacaoReacao, // W
      dados.prepDivulgacao,  // X
      dados.divulgacao,      // Y
      dados.linkPagina,      // Z
      dataAtual,             // AA
      emailUsuario,          // AB
      id                     // AC (ID)
    ]);
    
    return "Sucesso! Registro salvo.";
    
  } catch (e) {
    return "Erro ao salvar: " + e.toString();
  }
}

function migrarAdicionarId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(SETTINGS.NOME_ABA);
  if (!aba) throw new Error(`Aba não encontrada: ${SETTINGS.NOME_ABA}`);

  const lastRow = aba.getLastRow();
  const lastCol = aba.getLastColumn();
  if (lastRow === 0) return { ok: true, message: "Planilha vazia; nada para migrar." };

  const firstRow = lastCol > 0 ? aba.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  const hasHeader = firstRow.some(v =>
    typeof v === "string" && /(curso|turma|status|respons[aá]vel|oferta|sala)/i.test(v)
  );

  const uuidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

  // Heurística: se a última coluna já parece ser UUID na maioria das linhas, considera já migrado.
  if (!hasHeader && lastCol > 0 && lastRow >= 5) {
    const sampleCount = Math.min(50, lastRow);
    const sample = aba.getRange(1, lastCol, sampleCount, 1).getValues().flat();
    const uuidCount = sample.filter(v => typeof v === "string" && uuidRegex.test(v)).length;
    if (uuidCount / sampleCount >= 0.8) {
      return { ok: true, message: "IDs parecem já existir (última coluna). Nada a fazer." };
    }
  }

  // Define onde será a coluna de ID (sempre no fim, para não deslocar as colunas atuais).
  const idCol = lastCol + 1;

  let startRow = 1;
  let rowsToFill = lastRow;
  if (hasHeader) {
    const headerIdCell = aba.getRange(1, idCol);
    headerIdCell.setValue(SETTINGS.ID_COLUMN_NAME || "ID");
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
  return { ok: true, message: `Migração concluída. IDs preenchidos: ${filled}. Coluna: ${idCol}.` };
}

function pesquisarChamados(filtros, sort, page) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(SETTINGS.NOME_ABA);
  if (!aba) throw new Error(`Aba não encontrada: ${SETTINGS.NOME_ABA}`);

  const lastRow = aba.getLastRow();
  const lastCol = aba.getLastColumn();
  if (lastRow === 0 || lastCol === 0) {
    return { columns: _getDefaultColumns_(), rows: [], total: 0, truncated: false };
  }

  const values = aba.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const firstRow = values[0] || [];
  const hasHeader = firstRow.some(v =>
    typeof v === "string" && /(curso|turma|status|respons[aá]vel|oferta|sala)/i.test(v)
  );

  const columns = hasHeader ? _normalizeColumnsFromHeader_(firstRow) : _getDefaultColumns_();
  const data = values.slice(hasHeader ? 1 : 0);

  const filtered = _applyFilters_(data, columns, filtros || {});

  const sorted = _applySort_(filtered, columns, sort || { key: "", dir: "asc" });

  const offset = Math.max(0, Number(page && page.offset || 0));
  const limit = Math.min(500, Math.max(1, Number(page && page.limit || 200)));
  const paged = sorted.slice(offset, offset + limit);

  return {
    columns,
    rows: paged,
    total: sorted.length,
    truncated: sorted.length > paged.length
  };
}

function _getDefaultColumns_() {
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
    { key: "ID", label: (SETTINGS && SETTINGS.ID_COLUMN_NAME) ? SETTINGS.ID_COLUMN_NAME : "ID" }
  ];
}

function _normalizeColumnsFromHeader_(headerRow) {
  return (headerRow || []).map((h, idx) => {
    const label = (h && String(h).trim()) ? String(h).trim() : `Coluna ${idx + 1}`;
    return { key: label, label };
  });
}

function _colIndex_(columns, key) {
  if (!key) return -1;
  const k = String(key).trim().toLowerCase();
  for (let i = 0; i < columns.length; i++) {
    if (String(columns[i].key).trim().toLowerCase() === k) return i;
  }
  // fallback: aceitar nomes mais “tolerantes” para colunas padrão
  const aliases = {
    "curso": "Curso",
    "turma": "Turma",
    "status": "Status",
    "sala": "Sala",
    "responsavel": "Responsável",
    "oferta": "Oferta",
    "datacadastro": "Data Cadastro",
    "id": "ID"
  };
  const aliased = aliases[k];
  if (aliased) return _colIndex_(columns, aliased);
  return -1;
}

function _applyFilters_(rows, columns, filtros) {
  const idxCurso = _colIndex_(columns, "Curso");
  const idxTurma = _colIndex_(columns, "Turma");
  const idxStatus = _colIndex_(columns, "Status");
  const idxSala = _colIndex_(columns, "Sala");
  const idxResp = _colIndex_(columns, "Responsável");
  const idxOferta = _colIndex_(columns, "Oferta");
  const idxInicio = _colIndex_(columns, "Início");

  const curso = (filtros.curso || "").toString().trim();
  const turma = (filtros.turma || "").toString().trim();
  const status = (filtros.status || "").toString().trim();
  const sala = (filtros.sala || "").toString().trim();
  const responsavel = (filtros.responsavel || "").toString().trim();
  const oferta = (filtros.oferta || "").toString().trim();

  const inicioDe = _parseIsoDate_(filtros.inicioDe);
  const inicioAte = _parseIsoDate_(filtros.inicioAte);

  return (rows || []).filter(r => {
    if (curso && idxCurso >= 0 && String(r[idxCurso] || "") !== curso) return false;
    if (turma && idxTurma >= 0 && String(r[idxTurma] || "") !== turma) return false;
    if (status && idxStatus >= 0 && String(r[idxStatus] || "") !== status) return false;
    if (sala && idxSala >= 0 && String(r[idxSala] || "") !== sala) return false;
    if (responsavel && idxResp >= 0 && String(r[idxResp] || "") !== responsavel) return false;
    if (oferta && idxOferta >= 0 && String(r[idxOferta] || "") !== oferta) return false;

    if ((inicioDe || inicioAte) && idxInicio >= 0) {
      const d = _parseAnyDate_(r[idxInicio]);
      if (inicioDe && (!d || d < inicioDe)) return false;
      if (inicioAte && (!d || d > _endOfDay_(inicioAte))) return false;
    }
    return true;
  });
}

function _applySort_(rows, columns, sort) {
  const key = sort && sort.key ? String(sort.key) : "";
  const dir = (sort && sort.dir ? String(sort.dir) : "asc").toLowerCase() === "desc" ? -1 : 1;
  const idx = _colIndex_(columns, key);
  if (idx < 0) return rows || [];

  const copy = (rows || []).slice();
  copy.sort((a, b) => {
    const av = a[idx];
    const bv = b[idx];

    const ad = _parseAnyDate_(av);
    const bd = _parseAnyDate_(bv);
    if (ad && bd) return (ad.getTime() - bd.getTime()) * dir;

    const an = _parseNumber_(av);
    const bn = _parseNumber_(bv);
    if (an !== null && bn !== null) return (an - bn) * dir;

    const as = (av === null || av === undefined) ? "" : String(av).toLowerCase();
    const bs = (bv === null || bv === undefined) ? "" : String(bv).toLowerCase();
    if (as < bs) return -1 * dir;
    if (as > bs) return 1 * dir;
    return 0;
  });
  return copy;
}

function _parseIsoDate_(s) {
  if (!s) return null;
  const str = String(s).trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(str)) return null;
  const [y, m, d] = str.split("-").map(Number);
  const dt = new Date(y, m - 1, d);
  return isNaN(dt.getTime()) ? null : dt;
}

function _parseAnyDate_(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;

  const s = String(v).trim();
  // ISO yyyy-mm-dd
  const iso = _parseIsoDate_(s);
  if (iso) return iso;

  // dd/mm/yyyy (comum no BR)
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) {
    const dt = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
    return isNaN(dt.getTime()) ? null : dt;
  }

  const dt2 = new Date(s);
  return isNaN(dt2.getTime()) ? null : dt2;
}

function _endOfDay_(d) {
  const dt = new Date(d.getTime());
  dt.setHours(23, 59, 59, 999);
  return dt;
}

function _parseNumber_(v) {
  if (v === null || v === undefined) return null;
  const s = String(v).trim().replace(/\./g, "").replace(",", ".");
  if (!s) return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}