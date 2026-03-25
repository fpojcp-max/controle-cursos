/**
 * Camada Data – Configurações e listas estáticas (sem lógica).
 */

const Configuracoes = {
  NOME_ABA: "Turmas",
  NOME_COLUNA_ID: "ID Turma",

  // API de Agendamento (payload com spreadsheetId): aba na planilha informada.
  // Se não existir, a repository faz fallback para a primeira aba.
  NOME_ABA_AGENDAMENTO: "Turmas",

  /**
   * ID do arquivo no Google Sheets (trecho da URL entre /d/ e /edit).
   * NÃO use o parâmetro gid= (esse é só o ID interno da aba).
   */
  PLANILHA_AGENDAMENTOS_ID: "1X64jQdAi39Efs0bq1fWbQlva9uqxuKOOzsCfGz7dRUo",

  /** Nome exato da aba na planilha de agendamentos (case-sensitive). */
  NOME_ABA_AGENDAMENTOS: "Agendamentos",

  /** Fuso usado na Calendar API e na expansão de datas (sem coluna de TZ na planilha). */
  TIMEZONE_AGENDAMENTO: "America/Sao_Paulo",

  /** Ordem das colunas A–L (contrato com o print acordado). */
  CABECALHOS_AGENDAMENTO: [
    "Turma",
    "Curso",
    "Data",
    "Sala",
    "Hora Início",
    "Hora Fim",
    "Convidados",
    "Criado em",
    "Criado por",
    "ID Agendamento",
    "ID Turma",
    "ID Sala"
  ],

  /** Limite de convidados após normalização (evita limites da API). */
  LIMITE_CONVIDADOS_AGENDAMENTO: 50,

  /** Máximo de agendamentos excluídos por operação (Web App). */
  LIMITE_EXCLUSAO_AGENDAMENTOS_LOTE: 100,

  /**
   * Catálogo único: `rotulo` = texto em selects e planilhas; `identificadorCalendario` = recurso no Google Calendar.
   */
  CATALOGO_RECURSOS_SALA: [
    { rotulo: "(EDUCORP)-Auditório", identificadorCalendario: "c_188fsj4m0hnfii6qgjn69muavm51g@resource.calendar.google.com" },
    { rotulo: "(EDUCORP)-Idiomas", identificadorCalendario: "c_188akjbiofvveiatinhe5dd2oclmi@resource.calendar.google.com" },
    { rotulo: "(EDUCORP)-Lab. 01", identificadorCalendario: "c_188evr21m28u4h53guerh2v826efe@resource.calendar.google.com" },
    { rotulo: "(EDUCORP)-Lab. 02", identificadorCalendario: "c_188djbpv265fqhjajsdtsmkagmgt4@resource.calendar.google.com" },
    { rotulo: "(EDUCORP)-M1", identificadorCalendario: "c_188dal6k74eoij0mm2hqqq53bucps@resource.calendar.google.com" },
    { rotulo: "(EDUCORP)-M2", identificadorCalendario: "c_1886ke3al9bb8g4ditrf4340nflb4@resource.calendar.google.com" },
    { rotulo: "Externo", identificadorCalendario: "c_188b5jvdenpk4igdlbv99lsd1i52k@resource.calendar.google.com" }
  ]
};

/**
 * Retorna as opções para os campos do formulário (cursos, turmas, salas, etc.).
 * @returns {Object} Objeto com CURSOS, TURMAS, SALAS, OFERTAS, RESPONSAVEIS, PRIORIDADES, STATUS, BOOLEANOS
 */
function obterOpcoesFormulario() {
  return {
    CURSOS: [
      "108 - Inglês Básico II", "119 - BLS", "124 - Inglês Intermediário I", "125 - Inglês Intermediário II",
      "126 - Inglês Intermediário III", "178 - Curso de Capacitação de Membros da CIPA", "181 - Inglês Básico III",
      "194 - Fundamentos e Práticas em Educação Ambiental", "213 - NR35", "267 - Espanhol Básico I", "346 - Python",
      "373 - Proteção Radiológica", "395 - Brigada de Incêndio", "404 - Saúde em Foco", "420 - Ferramentas Google",
      "426 - Secretaria de Colegiados", "433 - Plataforma Sucupira Avançado", "445 - Comunicação e Relacionamento Interpessoal",
      "575 - Boas Práticas em Laboratório", "580 - Plataforma Sucupira Básico", "582 - Riscos Ergonômicos em Laboratório",
      "607 - LGPD", "626 - Inglês Básico para Patrulheiros", "627 - Oficina de Redação para Patrulheiros",
      "628 - Gestão de Projetos de Infraestrutura de Redes Cabeadas", "631 - CNV - Comunicação Não-Violenta",
      "632 - ACLS", "640 - DGA - Fase Preparatória da Licitação, com base na Lei nº 14.133/2021", "641 - Workshop CNV",
      "642 - SIGAD Patrulheiros", "648 - Comunique, Potencializar na Busca de Resultados de Qualidade", "650 - Aprimorad 2.0",
      "662 - Espanhol Intermediário II", "670 - Protocolos e Eventos para Cerimonial Público", "685 - SIGAD - Doctos. Digitais",
      "688 - II Seminário De Gestão Pública - Desafios, Resultados e Conexões", "689 - DGA - Oficina Cadastro de Materiais e de Serviços",
      "690 - Descomplicando a Escrita 2", "691 - Acessos Vasculares I", "695 - English Conversation Session", "700 - Linguagem Simples",
      "720 - Nova Instrução de Compra por Adiantamento", "722 - Integração de Médicos Residentes Ingressantes", "734 - Looker Studio",
      "736 - Mestre de Cerimônias", "737 - DGA Treinamento SIAD - Solicitações", "738 - Vivência Ubuntu de Educação Antirracista",
      "741 - Workshop Gestão do Estresse e Resiliência no Ambiente de Trabalho", "742 - Ética e Filosofia Política",
      "744 - Workshop - O Ser Biopsicossocial - Uma Filosofia sobre a Existência Humana", "746 - Imantados pela Natureza",
      "747 - TEA: Transtorno do Espectro Autista", "748 - Vida em Ciclos - Cotidiano, Trabalho e Bem-estar", "749 - Formação de Multiplicadores em CNV",
      "750 - Formação de Acolhedores", "751 - Fluxo dos Processos de Compras da Unicamp", "752 - DGA - Licitações Internacionais",
      "753 - DGA - Consórcios Nova Lei e Contratos", "756 - Workshop Responsabilidade Social", "757 - Oficina Descomplicando a Prática da Meditação",
      "758 - DGA - Treinamento de Ingressantes na DGA", "760 - Google Sites", "761 - Oficina sobre Documentos Fase Preparatória - Aquisição de Materiais",
      "762 - Oficina sobre Documentos Fase Preparatória - Contratação Serviços Comuns", "763 - Oficina sobre Documentos Fase Preparatória - Contratação Serviços Contínuos",
      "764 - Oficina sobre Documentos Fase Preparatória - Contratação Materiais e Serviços TIC", "765 - Oficina sobre Documentos Fase Preparatória - Contratação Obras e Serviços Engenharia",
      "767 - Oficina - Retenções Tributárias nas Contratações cujos Fornecedores são Pessoas Jurídicas", "768 - DGA - Orientações Apuração de Multa de Mora",
      "769 - Kubernetes na Prática - Básico Ao Avançado", "770 - Acolhimento da Pessoa com Deficiência", "771 - Metodologias Ativas, Comunicação e Didática",
      "772 - Atendimento de Pacientes em Cuidados Paliativos e Fase Ativa de Óbito", "773 - Workshop de Acessibilidade",
      "777 - IA na Prática: Simplifique e Inove no Trabalho", "788 - Inglês Intermediário Superior I", "791 - Cuidados com as Pessoas em Sofrimento Psíquico Grave Hospitalizadas",
      "793 - Almoxarifado - Treinamento de Uso do CLIF (Sistema de Almoxarifado)", "794 - Responsabilidade Socioambiental: Potencializar Ações na Unicamp",
      "795 - Workshop GT's Educorp - Resultados e Impactos", "796 - Workshop Comitês LGPD", "797 - Oficina Comitês LGPD", "798 - Canva Descomplicado",
      "799 - Mindfulness para Gestão Emocional", "800 - Programa de Acolhimento e Atualização em Enfermagem no Cuidado ao Paciente Crítico do Hospital de Clínicas da Unicamp",
      "801 - Técnicas de Poda de Árvores conforme ABNT 16.246-1", "802 - Treinamento do Registro do Ponto em Formato Eletrônico - Gestores",
      "803 - Treinamento do Registro do Ponto em Formato Eletrônico – RH's Locais", "804 - Treinamento do Registro do Ponto em Formato Eletrônico - Servidores",
      "805 - Orçamentação de Obras", "807 - Oficina Liquidação de Despesas: SIAD / Liquidação de Despesas", "808 - Acessibilidade à Pessoa com Deficiência",
      "830 - Oficina de CNV", "902 - PDG - Desafios, Riscos e Crises - Estratégias para Gestão", "903 - PDG - Palestra - O Papel dos Gestores na Sustentação da Excelência da Unicamp"
    ],
    TURMAS: ["2601", "2602", "2603", "2604", "2605", "2606", "2607", "2608", "2609", "2610", "2611", "2612", "2613", "2614", "2615", "2616", "2617", "2618", "2619", "2620"],
    SALAS: listarRotulosCatalogoRecursosSala(),
    OFERTAS: ["Assíncrono", "Síncrono", "Híbrido", "Presencial"],
    RESPONSAVEIS: ["Alexandre Fagiani", "Carlos", "Cecília", "Cirlene", "Denilson", "Elson", "Kitaka", "Márcia", "Raquel", "Valéria"],
    PRIORIDADES: ["Baixa", "Média", "Alta", "Urgente"],
    STATUS: ["Ativo", "Cancelado", "Suspenso"],
    BOOLEANOS: ["Sim", "Não"]
  };
}
