/**
 * Camada Repository – Aba Permissoes (planilha associada ao script).
 */

const PermissaoRepo = (() => {
  function chaveCabecalho_(valor) {
    let t = String(valor == null ? "" : valor)
      .trim()
      .toLowerCase();
    if (typeof t.normalize === "function") {
      t = t.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    }
    t = t.replace(/\s+/g, "_");
    if (t === "e-mail" || t === "e_mail") return "email";
    return t;
  }

  function obterAba_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const nome = String(Configuracoes.NOME_ABA_PERMISSOES || "Permissoes").trim();
    const aba = ss.getSheetByName(nome);
    if (!aba) {
      throw new Error("Aba não encontrada: " + nome);
    }
    return aba;
  }

  function indicesCabecalhos_(cabRow) {
    let idxNome = -1;
    let idxEmail = -1;
    let idxPerfil = -1;
    for (let c = 0; c < cabRow.length; c++) {
      const k = chaveCabecalho_(cabRow[c]);
      if (k === "nome") idxNome = c;
      if (k === "email") idxEmail = c;
      if (k === "perfil") idxPerfil = c;
    }
    return { idxNome: idxNome, idxEmail: idxEmail, idxPerfil: idxPerfil };
  }

  function obterPermissaoPorEmail_(emailNorm) {
    const email = String(emailNorm || "").trim().toLowerCase();
    if (!email) return null;
    const aba = obterAba_();
    const ultimaLinha = aba.getLastRow();
    const ultimaCol = aba.getLastColumn();
    if (ultimaLinha < 2 || ultimaCol < 1) return null;
    const cabRow = aba.getRange(1, 1, 1, ultimaCol).getValues()[0];
    const idx = indicesCabecalhos_(cabRow);
    if (idx.idxEmail < 0 || idx.idxPerfil < 0) return null;
    const dados = aba.getRange(2, 1, ultimaLinha, ultimaCol).getValues();
    for (let r = 0; r < dados.length; r++) {
      const row = dados[r];
      const em = String(row[idx.idxEmail] != null ? row[idx.idxEmail] : "").trim().toLowerCase();
      if (em !== email) continue;
      return {
        nome: idx.idxNome >= 0 ? String(row[idx.idxNome] != null ? row[idx.idxNome] : "").trim() : "",
        email: em,
        perfil: String(row[idx.idxPerfil] != null ? row[idx.idxPerfil] : "").trim()
      };
    }
    return null;
  }

  function listarDadosPermissoesPlanilha_() {
    let aba;
    try {
      aba = obterAba_();
    } catch (err) {
      return {
        idx: { idxNome: -1, idxEmail: -1, idxPerfil: -1 },
        linhas: [],
        numerosLinhaPlanilha: [],
        erro: err && err.message ? err.message : String(err)
      };
    }
    const ultimaLinha = aba.getLastRow();
    const ultimaCol = aba.getLastColumn();
    if (ultimaLinha < 2 || ultimaCol < 1) {
      return {
        idx: { idxNome: -1, idxEmail: -1, idxPerfil: -1 },
        linhas: [],
        numerosLinhaPlanilha: [],
        erro: null
      };
    }
    const cabRow = aba.getRange(1, 1, 1, ultimaCol).getValues()[0];
    const idx = indicesCabecalhos_(cabRow);
    if (idx.idxNome < 0 || idx.idxEmail < 0 || idx.idxPerfil < 0) {
      return {
        idx: idx,
        linhas: [],
        numerosLinhaPlanilha: [],
        erro: "Aba Permissoes: cabeçalhos obrigatórios Nome, E-mail e Perfil não encontrados."
      };
    }
    const dados = aba.getRange(2, 1, ultimaLinha, ultimaCol).getValues();
    const linhas = [];
    const numerosLinhaPlanilha = [];
    for (let r = 0; r < dados.length; r++) {
      linhas.push(dados[r]);
      numerosLinhaPlanilha.push(r + 2);
    }
    return { idx: idx, linhas: linhas, numerosLinhaPlanilha: numerosLinhaPlanilha, erro: null };
  }

  function inserirPermissao_(registro) {
    const aba = obterAba_();
    const ultimaCol = aba.getLastColumn();
    if (ultimaCol < 1) throw new Error("Aba Permissoes sem cabeçalhos na linha 1.");
    const cabRow = aba.getRange(1, 1, 1, ultimaCol).getValues()[0];
    const idx = indicesCabecalhos_(cabRow);
    if (idx.idxNome < 0 || idx.idxEmail < 0 || idx.idxPerfil < 0) {
      throw new Error("Aba Permissoes: cabeçalhos obrigatórios Nome, E-mail e Perfil não encontrados.");
    }
    const reg = registro && typeof registro === "object" ? registro : {};
    const row = new Array(ultimaCol).fill("");
    row[idx.idxNome] = String(reg.nome != null ? reg.nome : "").trim();
    row[idx.idxEmail] = String(reg.email != null ? reg.email : "").trim();
    row[idx.idxPerfil] = String(reg.perfil != null ? reg.perfil : "").trim();
    aba.appendRow(row);
  }

  function excluirLinhaPermissao_(sheetRow) {
    const aba = obterAba_();
    const sr = parseInt(String(sheetRow), 10);
    const ultimaLinha = aba.getLastRow();
    if (isNaN(sr) || sr < 2 || sr > ultimaLinha) {
      throw new Error("Linha da planilha inválida.");
    }
    aba.deleteRow(sr);
  }

  return {
    obterPermissaoPorEmail: obterPermissaoPorEmail_,
    listarDadosPermissoesPlanilha: listarDadosPermissoesPlanilha_,
    inserirPermissao: inserirPermissao_,
    excluirLinhaPermissao: excluirLinhaPermissao_
  };
})();
