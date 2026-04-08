/**
 * Backup diário da planilha (script container bound): cópia completa no Drive + rotação.
 * Executar com trigger temporal (meia-noite) como o utilizador que cria o trigger (ex.: desenvolvedor).
 *
 * Após implantar: executar uma vez `instalarTriggerBackupDiarioMeiaNoite()` no editor (ou criar o trigger manualmente).
 */

const BackupService = (() => {
  const PROP_ULTIMO_SUCESSO_YMD = "backupDiarioUltimoSucessoYmd";
  const DRIVE_V3_BASE = "https://www.googleapis.com/drive/v3/files";
  const MIME_SHEETS = "application/vnd.google-apps.spreadsheet";

  function tz_() {
    return (Configuracoes.BACKUP_TIMEZONE || Configuracoes.TIMEZONE_AGENDAMENTO || "America/Sao_Paulo");
  }

  function prefixo_() {
    return String(Configuracoes.BACKUP_PREFIXO_NOME_ARQUIVO || "BACKUP_SGC_");
  }

  function maxCopias_() {
    const n = Number(Configuracoes.BACKUP_MAX_COPIAS);
    return n >= 1 ? n : 30;
  }

  function hojeYmd_() {
    return Utilities.formatDate(new Date(), tz_(), "yyyy-MM-dd");
  }

  function emailAlerta_() {
    const cfg = String(Configuracoes.BACKUP_EMAIL_ALERTA || "").trim();
    if (cfg) return cfg;
    try {
      return String(Session.getActiveUser().getEmail() || "").trim();
    } catch (e) {
      return "";
    }
  }

  function enviarEmailFalha_(assunto, corpo) {
    const para = emailAlerta_();
    if (!para) return;
    try {
      MailApp.sendEmail({
        to: para,
        subject: String(assunto || "Backup planilha"),
        body: String(corpo || "")
      });
    } catch (e) {
      // Evita falha em cadeia; o erro original já foi ou será registado.
    }
  }

  function nomeEhBackupNosso_(nomeArquivo) {
    return String(nomeArquivo || "").indexOf(prefixo_()) === 0;
  }

  function driveApiRequest_(method, url, payloadObj) {
    const token = ScriptApp.getOAuthToken();
    const opts = {
      method: method,
      muteHttpExceptions: true,
      headers: {
        Authorization: "Bearer " + token
      }
    };
    if (payloadObj !== undefined) {
      opts.contentType = "application/json";
      opts.payload = JSON.stringify(payloadObj);
    }
    const resp = UrlFetchApp.fetch(url, opts);
    const code = resp.getResponseCode();
    const body = resp.getContentText() || "";
    if (code >= 200 && code < 300) {
      return body ? JSON.parse(body) : {};
    }
    const msg = "Drive API " + method + " " + url + " -> HTTP " + code + ": " + body;
    throw new Error(msg);
  }

  function validarPastaAcessivel_(folderId) {
    const url = DRIVE_V3_BASE + "/" + encodeURIComponent(folderId) + "?fields=id,name,mimeType&supportsAllDrives=true";
    const meta = driveApiRequest_("get", url);
    if (!meta || meta.mimeType !== "application/vnd.google-apps.folder") {
      throw new Error("ID não corresponde a pasta do Drive: " + folderId);
    }
  }

  function listarBackupsNaPasta_(folderId) {
    const out = [];
    const q = "'" + String(folderId).replace(/'/g, "\\'") + "' in parents and trashed=false and mimeType='" + MIME_SHEETS + "'";
    let pageToken = "";
    do {
      const url =
        DRIVE_V3_BASE +
        "?supportsAllDrives=true&includeItemsFromAllDrives=true&corpora=allDrives" +
        "&q=" + encodeURIComponent(q) +
        "&fields=nextPageToken,files(id,name,createdTime,mimeType)" +
        "&pageSize=200" +
        (pageToken ? "&pageToken=" + encodeURIComponent(pageToken) : "");
      const data = driveApiRequest_("get", url);
      const files = (data && data.files) || [];
      files.forEach(function (f) {
        if (!f || f.mimeType !== MIME_SHEETS) return;
        if (!nomeEhBackupNosso_(f.name)) return;
        out.push({
          id: f.id,
          name: f.name,
          createdTime: f.createdTime
        });
      });
      pageToken = (data && data.nextPageToken) || "";
    } while (pageToken);
    return out;
  }

  function arquivoBackupDoDiaJaExiste_(folderId, ymd) {
    const pref = prefixo_() + ymd + "_";
    const files = listarBackupsNaPasta_(folderId);
    for (let i = 0; i < files.length; i++) {
      const n = String(files[i].name || "");
      if (n.indexOf(pref) === 0) return true;
    }
    return false;
  }

  function moverCopiaParaPasta_(fileId, folderId) {
    const getUrl = DRIVE_V3_BASE + "/" + encodeURIComponent(fileId) + "?fields=parents&supportsAllDrives=true";
    const meta = driveApiRequest_("get", getUrl);
    const parents = Array.isArray(meta.parents) ? meta.parents : [];
    const removeParents = parents.filter(function (p) { return p !== folderId; }).join(",");
    const patchUrl =
      DRIVE_V3_BASE +
      "/" +
      encodeURIComponent(fileId) +
      "?supportsAllDrives=true&addParents=" +
      encodeURIComponent(folderId) +
      (removeParents ? "&removeParents=" + encodeURIComponent(removeParents) : "");
    driveApiRequest_("patch", patchUrl, {});
  }

  function rotacionarSeNecessario_(folderId) {
    const lim = maxCopias_();
    const files = listarBackupsNaPasta_(folderId);
    if (files.length <= lim) return;
    files.sort(function (a, b) {
      return new Date(a.createdTime).getTime() - new Date(b.createdTime).getTime();
    });
    const excedente = files.length - lim;
    for (let i = 0; i < excedente; i++) {
      const trashUrl = DRIVE_V3_BASE + "/" + encodeURIComponent(files[i].id) + "?supportsAllDrives=true";
      driveApiRequest_("patch", trashUrl, { trashed: true });
    }
  }

  function sanitizarNomePlanilha_(nome) {
    return String(nome || "Planilha").replace(/[\\/:*?"<>|]+/g, "-").trim() || "Planilha";
  }

  /**
   * Cópia diária: uma vez por dia (fuso BACKUP_TIMEZONE); até N cópias, depois remove a mais antiga.
   * @returns {{ skipped: boolean, reason?: string, message?: string }}
   */
  function executarBackupDiario_() {
    const folderId = String(Configuracoes.BACKUP_DRIVE_FOLDER_ID || "").trim();
    if (!folderId) {
      enviarEmailFalha_(
        "[SGC] Backup: configuração em falta",
        "Defina Configuracoes.BACKUP_DRIVE_FOLDER_ID com o ID da pasta do Drive."
      );
      return { skipped: true, reason: "sem_pasta" };
    }

    const props = PropertiesService.getScriptProperties();
    const hoje = hojeYmd_();
    if (props.getProperty(PROP_ULTIMO_SUCESSO_YMD) === hoje) {
      return { skipped: true, reason: "ja_executado_hoje" };
    }

    try {
      validarPastaAcessivel_(folderId);
    } catch (e1) {
      const msg = e1 && e1.message ? e1.message : String(e1);
      enviarEmailFalha_(
        "[SGC] Backup: pasta inacessível",
        "Não foi possível abrir a pasta de backup (ID). Verifique partilha e ID.\n\n" + msg
      );
      throw e1;
    }

    if (arquivoBackupDoDiaJaExiste_(folderId, hoje)) {
      rotacionarSeNecessario_(folderId);
      props.setProperty(PROP_ULTIMO_SUCESSO_YMD, hoje);
      return { skipped: true, reason: "arquivo_ja_existia", message: "Marcado como concluído (cópia do dia já na pasta)." };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const nomeCopia = prefixo_() + hoje + "_" + sanitizarNomePlanilha_(ss.getName());

    let copia;
    try {
      copia = ss.copy(nomeCopia);
      moverCopiaParaPasta_(copia.getId(), folderId);
      rotacionarSeNecessario_(folderId);
      props.setProperty(PROP_ULTIMO_SUCESSO_YMD, hoje);
    } catch (e2) {
      const msg = e2 && e2.message ? e2.message : String(e2);
      enviarEmailFalha_("[SGC] Backup: falha ao copiar ou arquivar", msg);
      try {
        if (copia && copia.getId) {
          const trashUrl = DRIVE_V3_BASE + "/" + encodeURIComponent(copia.getId()) + "?supportsAllDrives=true";
          driveApiRequest_("patch", trashUrl, { trashed: true });
        }
      } catch (e3) {
        // ignora limpeza
      }
      throw e2;
    }

    return { skipped: false, message: "Backup criado: " + nomeCopia };
  }

  return {
    executarBackupDiario: executarBackupDiario_
  };
})();

/**
 * Função alvo do trigger diário (meia-noite, America/Sao_Paulo).
 */
function executarBackupDiarioAgendado() {
  BackupService.executarBackupDiario();
}

/**
 * Instala (ou substitui) o trigger diário à meia-noite no fuso configurado.
 * Executar uma vez no editor Apps Script com a conta que deve possuir os backups.
 */
function instalarTriggerBackupDiarioMeiaNoite() {
  const fn = "executarBackupDiarioAgendado";
  const remover = ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction() === fn;
  });
  remover.forEach(function (t) {
    ScriptApp.deleteTrigger(t);
  });
  const tz = Configuracoes.BACKUP_TIMEZONE || Configuracoes.TIMEZONE_AGENDAMENTO || "America/Sao_Paulo";
  ScriptApp.newTrigger(fn)
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .inTimezone(tz)
    .create();
}

/**
 * Backup manual (esporádico), p.ex. se o automático falhou no dia.
 * Mesma lógica do agendado; respeita idempotência (um ficheiro por dia no fuso configurado).
 * @returns {Object}
 */
function executarBackupManual() {
  return BackupService.executarBackupDiario();
}
