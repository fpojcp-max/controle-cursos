/**
 * Camada Service – Permissões (aba Permissoes) e autorização por perfil.
 */

const PermissaoService = (() => {
  const MSG_SEM_CADASTRO =
    "Acesso negado. O seu e-mail não está cadastrado na aba Permissoes. Contacte um administrador.";

  const MSG_NEGADO_GERIR = "Acesso negado. Apenas administradores podem gerir permissões.";

  const MSG_NEGADO_CONSULTA = "Acesso negado. Apenas administradores podem consultar permissões.";

  const COLUNAS_CONSULTA_ = [
    { key: "nome", label: "Nome" },
    { key: "email", label: "E-mail" },
    { key: "perfil", label: "Perfil" }
  ];

  function perfilAdministrador_() {
    return String(Configuracoes.PERFIL_ADMINISTRADOR || "Administrador").trim();
  }

  function perfilResponsavel_() {
    return String(Configuracoes.PERFIL_RESPONSAVEL || "Responsável").trim();
  }

  function perfisValidos_() {
    return [perfilResponsavel_(), perfilAdministrador_()];
  }

  function ehAdministrador_(perfil) {
    return String(perfil || "").trim() === perfilAdministrador_();
  }

  function emailValidoBasico_(email) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || "").trim());
  }

  function obterPerfilVisitante_() {
    const email = SessaoWebApp.obterEmailAtivoNormalizado();
    const perm = PermissaoRepo.obterPermissaoPorEmail(email);
    if (!perm) {
      throw new Error(MSG_SEM_CADASTRO);
    }
    return {
      email: email,
      nome: String(perm.nome || "").trim(),
      perfil: String(perm.perfil || "").trim()
    };
  }

  function exigirAcessoSistema_() {
    return obterPerfilVisitante_();
  }

  function usuarioPodeGerirPermissoes_() {
    try {
      const v = obterPerfilVisitante_();
      return ehAdministrador_(v.perfil);
    } catch (e) {
      return false;
    }
  }

  function usuarioTemCadastroPermissao_(emailOpcional) {
    const email = String(emailOpcional || SessaoWebApp.obterEmailAtivoNormalizado() || "")
      .trim()
      .toLowerCase();
    if (!email) return false;
    return !!PermissaoRepo.obterPermissaoPorEmail(email);
  }

  function garantirPodeGerirPermissoes_() {
    if (!usuarioPodeGerirPermissoes_()) {
      throw new Error(MSG_NEGADO_GERIR);
    }
  }

  function garantirPodeAlterarTurmaLinha_(linha) {
    const visitante = obterPerfilVisitante_();
    if (ehAdministrador_(visitante.perfil)) return;
    const emailArm = RegistroRepo.extrairEmailUsuarioCriadorLinhaTurma(linha);
    SessaoWebApp.garantirMesmoUsuarioQueEmailArmazenadoOuErro(
      emailArm,
      SessaoWebApp.MSG_NAO_RESPONSAVEL,
      SessaoWebApp.MSG_SEM_EMAIL_CRIADOR_TURMA
    );
  }

  function garantirPodeAlterarTurmaPorId_(idTurma) {
    const id = String(idTurma || "").trim();
    if (!id) throw new Error("ID da turma inválido.");
    const linha = RegistroRepo.buscarLinhaPorId(id);
    if (!linha) throw new Error("Registro da turma não encontrado.");
    garantirPodeAlterarTurmaLinha_(linha);
  }

  function garantirPodeAgendarTurmaPorId_(idTurma) {
    garantirPodeAlterarTurmaPorId_(idTurma);
  }

  function celulaPerm_(row, map, chave) {
    const idx = map[chave];
    if (idx == null || idx < 0) return "";
    return String(row[idx] != null ? row[idx] : "").trim();
  }

  function mapaIndiceCelulas_(idxObj) {
    return {
      nome: idxObj.idxNome,
      email: idxObj.idxEmail,
      perfil: idxObj.idxPerfil
    };
  }

  function linhaSemDados_(row, map) {
    return (
      !celulaPerm_(row, map, "nome") &&
      !celulaPerm_(row, map, "email") &&
      !celulaPerm_(row, map, "perfil")
    );
  }

  function linhaPassaFiltros_(row, map, filtros) {
    const nomeF = String((filtros && filtros.nome) || "").trim().toLowerCase();
    const emailF = String((filtros && filtros.email) || "").trim().toLowerCase();
    const perfilF = String((filtros && filtros.perfil) || "").trim();
    if (nomeF) {
      const cell = celulaPerm_(row, map, "nome").toLowerCase();
      if (cell.indexOf(nomeF) < 0) return false;
    }
    if (emailF) {
      const cell = celulaPerm_(row, map, "email").toLowerCase();
      if (cell.indexOf(emailF) < 0) return false;
    }
    if (perfilF && celulaPerm_(row, map, "perfil") !== perfilF) return false;
    return true;
  }

  function ordenarPares_(pares, ordenacao, map) {
    const ord = ordenacao && typeof ordenacao === "object" ? ordenacao : {};
    const chave = String(ord.key || "nome").trim() || "nome";
    const dir = String(ord.dir || "asc").toLowerCase() === "desc" ? -1 : 1;
    pares.sort(function (a, b) {
      const va = celulaPerm_(a.row, map, chave).toLowerCase();
      const vb = celulaPerm_(b.row, map, chave).toLowerCase();
      const cmp = va.localeCompare(vb, undefined, { numeric: true, sensitivity: "base" });
      return cmp * dir;
    });
    return pares;
  }

  function dadosParaConsulta_() {
    const raw = PermissaoRepo.listarDadosPermissoesPlanilha();
    if (raw.erro) return { map: null, pares: [], erro: raw.erro };
    const map = mapaIndiceCelulas_(raw.idx || {});
    const pares = [];
    const linhas = raw.linhas || [];
    const nums = raw.numerosLinhaPlanilha || [];
    for (let i = 0; i < linhas.length; i++) {
      const row = linhas[i];
      if (linhaSemDados_(row, map)) continue;
      pares.push({ row: row, sheetRow: nums[i] });
    }
    return { map: map, pares: pares, erro: null };
  }

  function opcoesCadastroPermissoes_() {
    return {
      perfis: perfisValidos_(),
      labelPlaceholderSelect: "Selecione"
    };
  }

  function incluirPermissaoCadastro_(dados) {
    garantirPodeGerirPermissoes_();
    const p = dados && typeof dados === "object" ? dados : {};
    const nome = String(p.nome || "").trim();
    const email = String(p.email || "").trim().toLowerCase();
    const perfil = String(p.perfil || "").trim();
    if (!nome) return { success: false, message: "Informe o nome." };
    if (!emailValidoBasico_(email)) return { success: false, message: "Informe um e-mail válido." };
    if (perfisValidos_().indexOf(perfil) < 0) {
      return { success: false, message: "Perfil inválido." };
    }
    if (PermissaoRepo.obterPermissaoPorEmail(email)) {
      return { success: false, message: "E-mail já cadastrado." };
    }
    PermissaoRepo.inserirPermissao({ nome: nome, email: email, perfil: perfil });
    return { success: true, message: "Permissão cadastrada com sucesso." };
  }

  function obterOpcoesFiltroPermissoesConsulta_() {
    if (!usuarioPodeGerirPermissoes_()) {
      return { perfis: [], erro: MSG_NEGADO_CONSULTA };
    }
    return { perfis: perfisValidos_() };
  }

  function pesquisarPermissoesConsulta_(filtros, ordenacao, paginacao) {
    if (!usuarioPodeGerirPermissoes_()) {
      return { columns: [], rows: [], sheetRows: [], total: 0, erro: MSG_NEGADO_CONSULTA };
    }
    const base = dadosParaConsulta_();
    if (base.erro) {
      return { columns: [], rows: [], sheetRows: [], total: 0, erro: base.erro };
    }
    const map = base.map;
    let pares = (base.pares || []).filter(function (p) {
      return linhaPassaFiltros_(p.row, map, filtros);
    });
    pares = ordenarPares_(pares, ordenacao, map);
    const total = pares.length;
    const page = paginacao && typeof paginacao === "object" ? paginacao : {};
    const offset = Math.max(0, parseInt(String(page.offset || 0), 10) || 0);
    const limitRaw = parseInt(String(page.limit != null ? page.limit : 50), 10);
    const limit = limitRaw > 0 && limitRaw <= 500 ? limitRaw : 50;
    const slice = pares.slice(offset, offset + limit);
    return {
      columns: COLUNAS_CONSULTA_.slice(),
      rows: slice.map(function (p) {
        return [
          celulaPerm_(p.row, map, "nome"),
          celulaPerm_(p.row, map, "email"),
          celulaPerm_(p.row, map, "perfil")
        ];
      }),
      sheetRows: slice.map(function (p) {
        return p.sheetRow;
      }),
      total: total
    };
  }

  function obterPermissoesConsultaParaExportar_(filtros, ordenacao) {
    if (!usuarioPodeGerirPermissoes_()) {
      return { columns: [], rows: [], erro: MSG_NEGADO_CONSULTA };
    }
    const base = dadosParaConsulta_();
    if (base.erro) {
      return { columns: [], rows: [], erro: base.erro };
    }
    const map = base.map;
    let pares = (base.pares || []).filter(function (p) {
      return linhaPassaFiltros_(p.row, map, filtros);
    });
    pares = ordenarPares_(pares, ordenacao, map);
    return {
      columns: COLUNAS_CONSULTA_.slice(),
      rows: pares.map(function (p) {
        return [
          celulaPerm_(p.row, map, "nome"),
          celulaPerm_(p.row, map, "email"),
          celulaPerm_(p.row, map, "perfil")
        ];
      })
    };
  }

  function excluirPermissaoConsulta_(sheetRow) {
    garantirPodeGerirPermissoes_();
    const sr = parseInt(String(sheetRow), 10);
    if (isNaN(sr) || sr < 2) {
      return { success: false, message: "Número de linha inválido." };
    }
    const raw = PermissaoRepo.listarDadosPermissoesPlanilha();
    if (raw.erro) {
      return { success: false, message: raw.erro };
    }
    const map = mapaIndiceCelulas_(raw.idx || {});
    const nums = raw.numerosLinhaPlanilha || [];
    const linhas = raw.linhas || [];
    let targetRow = null;
    for (let i = 0; i < linhas.length; i++) {
      if (nums[i] === sr) {
        targetRow = linhas[i];
        break;
      }
    }
    if (!targetRow) {
      return { success: false, message: "Registo não encontrado." };
    }
    const adminLiteral = perfilAdministrador_();
    let adminCount = 0;
    for (let j = 0; j < linhas.length; j++) {
      if (linhaSemDados_(linhas[j], map)) continue;
      const p = celulaPerm_(linhas[j], map, "perfil");
      if (p === adminLiteral) adminCount++;
    }
    const perfilAlvo = celulaPerm_(targetRow, map, "perfil");
    if (perfilAlvo === adminLiteral && adminCount <= 1) {
      return {
        success: false,
        message: "É necessário manter ao menos um usuário administrador."
      };
    }
    try {
      PermissaoRepo.excluirLinhaPermissao(sr);
    } catch (err) {
      return { success: false, message: err && err.message ? err.message : String(err) };
    }
    return { success: true, message: "Permissão excluída com sucesso." };
  }

  return {
    MSG_SEM_CADASTRO: MSG_SEM_CADASTRO,
    MSG_NEGADO_GERIR: MSG_NEGADO_GERIR,
    MSG_NEGADO_CONSULTA: MSG_NEGADO_CONSULTA,
    exigirAcessoSistema: exigirAcessoSistema_,
    obterPerfilVisitante: obterPerfilVisitante_,
    usuarioPodeGerirPermissoes: usuarioPodeGerirPermissoes_,
    garantirPodeGerirPermissoes: garantirPodeGerirPermissoes_,
    usuarioTemCadastroPermissao: usuarioTemCadastroPermissao_,
    garantirPodeAlterarTurmaLinha: garantirPodeAlterarTurmaLinha_,
    garantirPodeAlterarTurmaPorId: garantirPodeAlterarTurmaPorId_,
    garantirPodeAgendarTurmaPorId: garantirPodeAgendarTurmaPorId_,
    opcoesCadastroPermissoes: opcoesCadastroPermissoes_,
    incluirPermissaoCadastro: incluirPermissaoCadastro_,
    obterOpcoesFiltroPermissoesConsulta: obterOpcoesFiltroPermissoesConsulta_,
    pesquisarPermissoesConsulta: pesquisarPermissoesConsulta_,
    obterPermissoesConsultaParaExportar: obterPermissoesConsultaParaExportar_,
    excluirPermissaoConsulta: excluirPermissaoConsulta_
  };
})();
