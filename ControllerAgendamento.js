/**
 * Camada Controller – Entrada/validação superficial para endpoint de Agendamento.
 * (Service contém as regras de negócio.)
 */

const AgendamentoController = (() => {
  function errorResponse_(code, message, details) {
    return {
      status: "erro",
      message: String(message || "Erro"),
      code: code || "INTERNAL_ERROR",
      details: Array.isArray(details) ? details : []
    };
  }

  function criarEventos_(request) {
    try {
      return AgendamentoService.criarEventos(request);
    } catch (err) {
      // Erros esperados lançados pela Service.
      if (err && err.code && typeof err.message === "string") {
        return errorResponse_(err.code, err.message, err.details);
      }

      // Erro inesperado.
      return errorResponse_("INTERNAL_ERROR", (err && err.message) ? err.message : "Erro interno", []);
    }
  }

  return {
    criarEventos: criarEventos_
  };
})();

// Wrapper global (padrão do projeto)
function criarEventosAgendamentoController(request) { return AgendamentoController.criarEventos(request); }

