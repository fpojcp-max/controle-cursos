/**
 * Camada Data – Normalização/parse de payloads (sem regra de negócio).
 */

const AgendamentoData = (() => {
  const TIMEZONE = "America/Sao_Paulo";

  function isNonEmptyString_(v) {
    return v !== null && v !== undefined && String(v).trim() !== "";
  }

  function validarEParsearDataISO_(dataISO, fieldLabel, details) {
    if (!isNonEmptyString_(dataISO) || !/^\d{4}-\d{2}-\d{2}$/.test(String(dataISO))) {
      details.push({ field: fieldLabel, message: "Data inválida" });
      return null;
    }

    // parseDate valida datas reais (ex: 2026-02-30).
    try {
      const dt = Utilities.parseDate(String(dataISO) + " 00:00", TIMEZONE, "yyyy-MM-dd HH:mm");
      if (!(dt instanceof Date) || isNaN(dt.getTime())) {
        details.push({ field: fieldLabel, message: "Data inválida" });
        return null;
      }
      return dt;
    } catch (_) {
      details.push({ field: fieldLabel, message: "Data inválida" });
      return null;
    }
  }

  function validarEParsearHoraHHMM_(horaHHMM, fieldLabel, details) {
    if (!isNonEmptyString_(horaHHMM) || !/^\d{2}:\d{2}$/.test(String(horaHHMM))) {
      details.push({ field: fieldLabel, message: "Hora inválida" });
      return null;
    }

    const m = String(horaHHMM).match(/^(\d{2}):(\d{2})$/);
    const hh = Number(m[1]);
    const mm = Number(m[2]);
    if (hh < 0 || hh > 23 || mm < 0 || mm > 59) {
      details.push({ field: fieldLabel, message: "Hora inválida" });
      return null;
    }
    return { hh, mm };
  }

  function parseStartEnd_(dataISO, horaInicio, horaFim, baseFieldPrefix, details) {
    const dataBase = validarEParsearDataISO_(dataISO, `${baseFieldPrefix}.data`, details);
    if (!dataBase) return null;

    const inicio = validarEParsearHoraHHMM_(horaInicio, `${baseFieldPrefix}.horaInicio`, details);
    if (!inicio) return null;

    const fim = validarEParsearHoraHHMM_(horaFim, `${baseFieldPrefix}.horaFim`, details);
    if (!fim) return null;

    // Constrói em TIMEZONE fixo para ser deterministico (DST incluído).
    try {
      const dtInicio = Utilities.parseDate(
        String(dataISO) + " " + String(horaInicio),
        TIMEZONE,
        "yyyy-MM-dd HH:mm"
      );
      const dtFim = Utilities.parseDate(
        String(dataISO) + " " + String(horaFim),
        TIMEZONE,
        "yyyy-MM-dd HH:mm"
      );
      if (dtFim.getTime() <= dtInicio.getTime()) {
        // Regra explicita: horaFim == horaInicio e horaFim < horaInicio são inválidos.
        details.push({ field: `${baseFieldPrefix}.horaFim`, message: "Hora final deve ser maior que a inicial" });
        return null;
      }
      return { dtInicio, dtFim };
    } catch (_) {
      details.push({ field: `${baseFieldPrefix}.data`, message: "Data inválida" });
      return null;
    }
  }

  return {
    TIMEZONE,
    parseStartEnd_: parseStartEnd_
  };
})();

