/**
 * Backup diário da planilha (script container bound): cópia completa no Drive + rotação.
 * Executar com trigger temporal (meia-noite) como o utilizador que cria o trigger (ex.: desenvolvedor).
 *
 * Após implantar: executar uma vez `instalarTriggerBackupDiarioMeiaNoite()` no editor (ou criar o trigger manualmente).
 */

const BackupService = (() => {
  const PROP_ULTIMO_SUCESSO_YMD = "backupDiarioUltimoSucessoYmd";

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

  function listarBackupsNaPasta_(pasta) {
    const out = [];
    const it = pasta.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      if (f.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;
      if (!nomeEhBackupNosso_(f.getName())) continue;
      out.push(f);
    }
    return out;
  }

  function arquivoBackupDoDiaJaExiste_(pasta, ymd) {
    const pref = prefixo_() + ymd + "_";
    const it = pasta.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      if (f.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;
      const n = String(f.getName());
      if (n.indexOf(pref) === 0) return true;
    }
    return false;
  }

  function moverCopiaParaPasta_(fileId, pastaDestino) {
    const file = DriveApp.getFileById(fileId);
    pastaDestino.addFile(file);
    const manterId = pastaDestino.getId();
    const parents = file.getParents();
    while (parents.hasNext()) {
      const p = parents.next();
      if (p.getId() !== manterId) {
        p.removeFile(file);
      }
    }
  }

  function rotacionarSeNecessario_(pasta) {
    const lim = maxCopias_();
    const files = listarBackupsNaPasta_(pasta);
    if (files.length <= lim) return;
    files.sort(function (a, b) {
      return a.getDateCreated().getTime() - b.getDateCreated().getTime();
    });
    const excedente = files.length - lim;
    for (let i = 0; i < excedente; i++) {
      files[i].setTrashed(true);
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

    let pasta;
    try {
      pasta = DriveApp.getFolderById(folderId);
    } catch (e1) {
      const msg = e1 && e1.message ? e1.message : String(e1);
      enviarEmailFalha_(
        "[SGC] Backup: pasta inacessível",
        "Não foi possível abrir a pasta de backup (ID). Verifique partilha e ID.\n\n" + msg
      );
      throw e1;
    }

    if (arquivoBackupDoDiaJaExiste_(pasta, hoje)) {
      rotacionarSeNecessario_(pasta);
      props.setProperty(PROP_ULTIMO_SUCESSO_YMD, hoje);
      return { skipped: true, reason: "arquivo_ja_existia", message: "Marcado como concluído (cópia do dia já na pasta)." };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const nomeCopia = prefixo_() + hoje + "_" + sanitizarNomePlanilha_(ss.getName());

    let copia;
    try {
      copia = ss.copy(nomeCopia);
      moverCopiaParaPasta_(copia.getId(), pasta);
      rotacionarSeNecessario_(pasta);
      props.setProperty(PROP_ULTIMO_SUCESSO_YMD, hoje);
    } catch (e2) {
      const msg = e2 && e2.message ? e2.message : String(e2);
      enviarEmailFalha_("[SGC] Backup: falha ao copiar ou arquivar", msg);
      try {
        if (copia && copia.getId) {
          DriveApp.getFileById(copia.getId()).setTrashed(true);
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
