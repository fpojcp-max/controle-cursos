# Mapeamento função por função – Arquitetura 5 camadas

Documento de referência: cada função da arquitetura alvo, sem código.  
Nomes em português; responsabilidades e fluxos definidos.

---

## Convenções

- **Assinatura**: parâmetros e tipo de retorno descritos em texto.
- **Chama**: outras funções (do backend) que esta invoca.
- **Chamada por**: quem invoca esta função (cliente = UI via `google.script.run`; ou outra função do backend).
- **Origem**: função atual que corresponde ou “Nova” se for criada na reestruturação.

---

# 1. Camada Controller (entrada do cliente)

Funções expostas ao cliente via `google.script.run`. Arquivos: **Main.gs**, **ControllerRegistro.gs**.

---

## 1.1 doGet

| Item | Valor |
|------|--------|
| **Nome** | `doGet` (manter – obrigatório GAS) |
| **Arquivo** | Main.gs |
| **Assinatura** | `doGet(e)` — `e` contém `parameter.view`. Retorna `HtmlService.HtmlOutput`. |
| **Responsabilidade** | Ponto de entrada do Web App: lê `view` (consulta ou cadastro), escolhe o HTML (Consulta ou Index) e devolve a página avaliada. |
| **Chama** | `HtmlService.createTemplateFromFile`, `.evaluate()`, etc. (apenas APIs do GAS). |
| **Chamada por** | GAS (runtime). |
| **Origem** | `doGet` em Código.js. |

---

## 1.2 obterUrlWebApp

| Item | Valor |
|------|--------|
| **Nome** | `obterUrlWebApp` |
| **Arquivo** | Main.gs |
| **Assinatura** | `obterUrlWebApp(visualizacao)` — `visualizacao`: string (ex.: `"consulta"`, `"cadastro"`) ou vazio. Retorna string (URL). |
| **Responsabilidade** | Montar a URL do Web App com query `?view=...` para links e redirecionamentos na UI. |
| **Chama** | `ScriptApp.getService().getUrl()`. |
| **Chamada por** | Cliente (templates Index.html e Consulta.html via `<?= obterUrlWebApp(...) ?>`) e eventualmente JS. |
| **Origem** | `getWebAppUrl` em Código.js. |

---

## 1.3 obterConfiguracaoFormulario

| Item | Valor |
|------|--------|
| **Nome** | `obterConfiguracaoFormulario` |
| **Arquivo** | ControllerRegistro.gs |
| **Assinatura** | `obterConfiguracaoFormulario()` — sem parâmetros. Retorna objeto com listas: CURSOS, TURMAS, SALAS, OFERTAS, RESPONSAVEIS, PRIORIDADES, STATUS, BOOLEANOS. |
| **Responsabilidade** | Expor ao cliente as opções dos formulários (cadastro e filtros da consulta); apenas repassa os dados da camada Data. |
| **Chama** | Função ou constante da camada Data que devolve as listas (ex.: `obterOpcoesFormulario()` ou leitura de constante). |
| **Chamada por** | Cliente (JavaScript.html e ConsultaJavaScript.html) via `google.script.run.obterConfiguracaoFormulario()`. |
| **Origem** | `getFormConfiguration` em Config.js. |

---

## 1.4 pesquisarRegistros

| Item | Valor |
|------|--------|
| **Nome** | `pesquisarRegistros` |
| **Arquivo** | ControllerRegistro.gs |
| **Assinatura** | `pesquisarRegistros(filtros, ordenacao, paginacao)` — filtros: objeto (curso, turma, status, sala, responsavel, oferta, inicioDe, inicioAte); ordenacao: { chave, direcao }; paginacao: { deslocamento, limite }. Retorna `{ colunas, linhas, total, truncado }`. |
| **Responsabilidade** | Entrada da tela de consulta: recebe filtros, ordenação e paginação, delega ao Service e devolve colunas e linhas para a tabela. |
| **Chama** | Serviço de busca (ex.: `buscarRegistrosComFiltros(filtros, ordenacao, paginacao)`). |
| **Chamada por** | Cliente (ConsultaJavaScript.html) via `google.script.run.pesquisarRegistros(...)`. |
| **Origem** | `pesquisarChamados` em Código.js. |

---

## 1.5 salvarRegistro

| Item | Valor |
|------|--------|
| **Nome** | `salvarRegistro` |
| **Arquivo** | ControllerRegistro.gs |
| **Assinatura** | `salvarRegistro(dados)` — dados: objeto com todos os campos do formulário de cadastro (curso, turma, inicio, fim, sala, oferta, etc.). Retorna string (mensagem de sucesso ou erro). |
| **Responsabilidade** | Entrada para cadastro de novo registro: valida/orquestra via Service e persiste via Repository; devolve mensagem ao cliente. |
| **Chama** | Service (ex.: `cadastrarRegistro(dados)`) ou diretamente Repository (ex.: `inserirLinha(...)`), conforme definido na implementação. |
| **Chamada por** | Cliente (JavaScript.html) via `google.script.run.salvarRegistro(dados)`. |
| **Origem** | `salvarChamado` em Código.js. |

---

## 1.6 atualizarRegistro

| Item | Valor |
|------|--------|
| **Nome** | `atualizarRegistro` |
| **Arquivo** | ControllerRegistro.gs |
| **Assinatura** | `atualizarRegistro(id, dados)` — id: string (UUID do registro); dados: mesmo objeto do formulário. Retorna string (mensagem de sucesso ou erro). |
| **Responsabilidade** | Entrada para edição: atualiza registro existente na planilha sem criar nova linha; delega ao Service/Repository. |
| **Chama** | Service (ex.: `atualizarRegistro(id, dados)`) que usa Repository para localizar linha e atualizar. |
| **Chamada por** | Cliente (JavaScript.html) via `google.script.run.atualizarRegistro(id, dados)`. |
| **Origem** | Nova (fluxo de edição; equivalente ao antigo `atualizarChamado` se existir). |

---

## 1.7 obterRegistroPorId

| Item | Valor |
|------|--------|
| **Nome** | `obterRegistroPorId` |
| **Arquivo** | ControllerRegistro.gs |
| **Assinatura** | `obterRegistroPorId(id)` — id: string (UUID). Retorna objeto com os campos do registro (curso, turma, inicio, fim, etc.) ou null se não encontrado. |
| **Responsabilidade** | Entrada para carregar a tela de edição: busca um registro pelo ID e devolve os dados para preencher o formulário. |
| **Chama** | Service (ex.: `obterRegistroPorId(id)`) ou Repository (ex.: `buscarLinhaPorId(id)` + conversão para DTO). |
| **Chamada por** | Cliente (JavaScript.html) via `google.script.run.obterRegistroPorId(id)`. |
| **Origem** | Nova (fluxo de edição; equivalente ao antigo `obterChamadoPorId` se existir). |

---

## 1.8 executarMigracaoIds

| Item | Valor |
|------|--------|
| **Nome** | `executarMigracaoIds` |
| **Arquivo** | ControllerRegistro.gs |
| **Assinatura** | `executarMigracaoIds()` — sem parâmetros. Retorna objeto `{ ok, message }`. |
| **Responsabilidade** | Executar uma única vez a migração que preenche a coluna ID nos registros que ainda não têm; usado manualmente pelo desenvolvedor. |
| **Chama** | Repository (ex.: `preencherColunaId()` ou funções equivalentes de leitura/escrita da planilha). |
| **Chamada por** | Desenvolvedor (execução manual no editor do GAS); não chamada pela UI. |
| **Origem** | `migrarAdicionarId` em Código.js. |

---

# 2. Camada Service (regras de negócio)

Orquestração, filtros, ordenação, paginação, conversão linha ↔ DTO. Arquivo: **ServicoRegistro.gs**.

---

## 2.1 buscarRegistrosComFiltros

| Item | Valor |
|------|--------|
| **Nome** | `buscarRegistrosComFiltros` |
| **Arquivo** | ServicoRegistro.gs |
| **Assinatura** | `buscarRegistrosComFiltros(filtros, ordenacao, paginacao)` — mesmos tipos do Controller. Retorna `{ colunas, linhas, total, truncado }`. |
| **Responsabilidade** | Obter dados brutos do Repository, aplicar filtros e ordenação (em memória), aplicar paginação e montar a resposta com colunas e linhas. |
| **Chama** | Repository: `lerDadosPlanilha()` (ou equivalente); internamente: `obterColunasPadrao` ou normalização de cabeçalho, `aplicarFiltros`, `aplicarOrdenacao`. |
| **Chamada por** | Controller `pesquisarRegistros`. |
| **Origem** | Lógica interna de `pesquisarChamados` + `_applyFilters_`, `_applySort_`, `_getDefaultColumns_`, `_normalizeColumnsFromHeader_`. |

---

## 2.2 aplicarFiltros

| Item | Valor |
|------|--------|
| **Nome** | `aplicarFiltros` |
| **Arquivo** | ServicoRegistro.gs |
| **Assinatura** | `aplicarFiltros(linhas, colunas, filtros)` — linhas: array de arrays (dados); colunas: array de { key, label }; filtros: objeto. Retorna array de linhas (filtradas). |
| **Responsabilidade** | Filtrar as linhas por curso, turma, status, sala, responsável, oferta e intervalo de data de início. |
| **Chama** | Helpers de data (ex.: `interpretarDataIso`, `interpretarData`, `finalDoDia`) e de índice (`indiceDaColuna`). |
| **Chamada por** | `buscarRegistrosComFiltros`. |
| **Origem** | `_applyFilters_` em Código.js. |

---

## 2.3 aplicarOrdenacao

| Item | Valor |
|------|--------|
| **Nome** | `aplicarOrdenacao` |
| **Arquivo** | ServicoRegistro.gs |
| **Assinatura** | `aplicarOrdenacao(linhas, colunas, ordenacao)` — ordenacao: { chave, direcao }. Retorna array de linhas (ordenadas). |
| **Responsabilidade** | Ordenar as linhas por uma coluna (chave) em ordem ascendente ou descendente; tratar data, número e texto. |
| **Chama** | `indiceDaColuna`, `interpretarData`, `interpretarNumero`. |
| **Chamada por** | `buscarRegistrosComFiltros`. |
| **Origem** | `_applySort_` em Código.js. |

---

## 2.4 indiceDaColuna

| Item | Valor |
|------|--------|
| **Nome** | `indiceDaColuna` |
| **Arquivo** | ServicoRegistro.gs |
| **Assinatura** | `indiceDaColuna(colunas, chave)` — chave: string (nome da coluna). Retorna número (índice) ou -1. |
| **Responsabilidade** | Dado o array de colunas (key/label), retornar o índice da coluna cujo key/label corresponde à chave (com aliases para nomes comuns). |
| **Chama** | Nenhuma (lógica pura). |
| **Chamada por** | `aplicarFiltros`, `aplicarOrdenacao`. |
| **Origem** | `_colIndex_` em Código.js. |

---

## 2.5 linhaParaDados

| Item | Valor |
|------|--------|
| **Nome** | `linhaParaDados` |
| **Arquivo** | ServicoRegistro.gs |
| **Assinatura** | `linhaParaDados(linha)` — linha: array de valores (uma linha da planilha, ordem A–Z). Retorna objeto com propriedades curso, turma, inicio, fim, sala, oferta, etc. |
| **Responsabilidade** | Converter uma linha lida da planilha (índices fixos A–Z) no objeto DTO usado pelo formulário de cadastro/edição. |
| **Chama** | Nenhuma. |
| **Chamada por** | Service ao montar resposta de `obterRegistroPorId`; ou Repository se a conversão ficar no Service. |
| **Origem** | `_rowToDadosObject_` (quando existia no fluxo de edição) ou equivalente. |

---

## 2.6 dadosParaLinha

| Item | Valor |
|------|--------|
| **Nome** | `dadosParaLinha` |
| **Arquivo** | ServicoRegistro.gs |
| **Assinatura** | `dadosParaLinha(dados, metadados)` — dados: objeto do formulário; metadados: { dataCadastro, emailUsuario, id } opcional. Retorna array de valores na ordem das colunas (A até AC). |
| **Responsabilidade** | Converter o objeto do formulário em array de valores para inserção ou atualização na planilha (incluindo data cadastro, e-mail e ID quando for novo registro). |
| **Chama** | Nenhuma. |
| **Chamada por** | Service de cadastro/atualização antes de chamar Repository. |
| **Origem** | Lógica hoje dentro de `salvarChamado` (montagem do array). |

---

## 2.7 cadastrarRegistro

| Item | Valor |
|------|--------|
| **Nome** | `cadastrarRegistro` |
| **Arquivo** | ServicoRegistro.gs |
| **Assinatura** | `cadastrarRegistro(dados)` — dados: objeto do formulário. Retorna string (mensagem de sucesso ou erro). |
| **Responsabilidade** | Gerar ID e metadados (data, e-mail), montar a linha com `dadosParaLinha`, chamar Repository para inserir e retornar mensagem. |
| **Chama** | `dadosParaLinha`; Repository `inserirLinha`. |
| **Chamada por** | Controller `salvarRegistro`. |
| **Origem** | Lógica de `salvarChamado` em Código.js. |

---

## 2.8 atualizarRegistro (Service)

| Item | Valor |
|------|--------|
| **Nome** | `atualizarRegistro` (camada Service) |
| **Arquivo** | ServicoRegistro.gs |
| **Assinatura** | `atualizarRegistro(id, dados)` — id: string; dados: objeto do formulário. Retorna string (mensagem). |
| **Responsabilidade** | Obter índice/linha do registro via Repository, montar array de valores (apenas colunas editáveis, mantendo data/e-mail/ID), chamar Repository para atualizar. |
| **Chama** | Repository `buscarIndiceLinhaPorId` (ou `buscarLinhaPorId`), `atualizarLinha`; `dadosParaLinha` (versão sem metadados de auditoria para atualização). |
| **Chamada por** | Controller `atualizarRegistro`. |
| **Origem** | Nova; equivalente ao antigo `atualizarChamado`. |

---

## 2.9 obterRegistroPorId (Service)

| Item | Valor |
|------|--------|
| **Nome** | `obterRegistroPorId` (camada Service) |
| **Arquivo** | ServicoRegistro.gs |
| **Assinatura** | `obterRegistroPorId(id)` — id: string. Retorna objeto DTO ou null. |
| **Responsabilidade** | Pedir ao Repository a linha do registro com esse ID; converter com `linhaParaDados` e retornar. |
| **Chama** | Repository `buscarLinhaPorId(id)`; `linhaParaDados`. |
| **Chamada por** | Controller `obterRegistroPorId`. |
| **Origem** | Nova; equivalente ao antigo `obterChamadoPorId`. |

---

## 2.10 interpretarDataIso

| Item | Valor |
|------|--------|
| **Nome** | `interpretarDataIso` |
| **Arquivo** | ServicoRegistro.gs |
| **Assinatura** | `interpretarDataIso(str)` — str: string no formato yyyy-mm-dd. Retorna Date ou null. |
| **Responsabilidade** | Interpretar string de data no formato ISO (yyyy-mm-dd) e retornar objeto Date. |
| **Chama** | Nenhuma. |
| **Chamada por** | `interpretarData`, `aplicarFiltros`. |
| **Origem** | `_parseIsoDate_` em Código.js. |

---

## 2.11 interpretarData

| Item | Valor |
|------|--------|
| **Nome** | `interpretarData` |
| **Arquivo** | ServicoRegistro.gs |
| **Assinatura** | `interpretarData(valor)` — valor: string, Date ou outro. Retorna Date ou null. |
| **Responsabilidade** | Interpretar valor como data (ISO, dd/mm/yyyy ou genérico) para uso em filtros e ordenação. |
| **Chama** | `interpretarDataIso`. |
| **Chamada por** | `aplicarFiltros`, `aplicarOrdenacao`. |
| **Origem** | `_parseAnyDate_` em Código.js. |

---

## 2.12 finalDoDia

| Item | Valor |
|------|--------|
| **Nome** | `finalDoDia` |
| **Arquivo** | ServicoRegistro.gs |
| **Assinatura** | `finalDoDia(data)` — data: Date. Retorna Date (mesmo dia às 23:59:59.999). |
| **Responsabilidade** | Ajustar uma data para o fim do dia (útil para filtro “início até”). |
| **Chama** | Nenhuma. |
| **Chamada por** | `aplicarFiltros`. |
| **Origem** | `_endOfDay_` em Código.js. |

---

## 2.13 interpretarNumero

| Item | Valor |
|------|--------|
| **Nome** | `interpretarNumero` |
| **Arquivo** | ServicoRegistro.gs |
| **Assinatura** | `interpretarNumero(valor)` — valor: qualquer. Retorna number ou null. |
| **Responsabilidade** | Tentar converter valor em número para ordenação numérica. |
| **Chama** | Nenhuma. |
| **Chamada por** | `aplicarOrdenacao`. |
| **Origem** | `_parseNumber_` em Código.js. |

---

## 2.14 normalizarColunasDoCabecalho

| Item | Valor |
|------|--------|
| **Nome** | `normalizarColunasDoCabecalho` |
| **Arquivo** | ServicoRegistro.gs (ou RepositorioPlanilha.gs se for considerado “estrutura da planilha”) |
| **Assinatura** | `normalizarColunasDoCabecalho(linhaCabecalho)` — array de strings. Retorna array de { key, label }. |
| **Responsabilidade** | Transformar a primeira linha lida da planilha em array de colunas com key e label (para exibição e filtro/ordenação). |
| **Chama** | Nenhuma. |
| **Chamada por** | `buscarRegistrosComFiltros` (após ler dados do Repository). |
| **Origem** | `_normalizeColumnsFromHeader_` em Código.js. |

---

# 3. Camada Repository (acesso à planilha)

Leitura e escrita na planilha; conhecimento de aba e estrutura. Arquivo: **RepositorioPlanilha.gs**.

---

## 3.1 lerDadosPlanilha

| Item | Valor |
|------|--------|
| **Nome** | `lerDadosPlanilha` |
| **Arquivo** | RepositorioPlanilha.gs |
| **Assinatura** | `lerDadosPlanilha()` — sem parâmetros. Retorna objeto `{ valores, temCabecalho }` ou `{ valores }`, onde valores é array 2D (getValues ou getDisplayValues). |
| **Responsabilidade** | Ler toda a aba configurada (Configuracoes.NOME_ABA) e devolver matriz de valores; opcionalmente indicar se a primeira linha é cabeçalho. |
| **Chama** | Configuracoes (leitura de NOME_ABA); SpreadsheetApp. |
| **Chamada por** | Service `buscarRegistrosComFiltros`; eventualmente Service de obterRegistroPorId/atualizarRegistro. |
| **Origem** | Lógica de leitura dentro de `pesquisarChamados` e de `obterChamadoPorId`/`atualizarChamado`. |

---

## 3.2 inserirLinha

| Item | Valor |
|------|--------|
| **Nome** | `inserirLinha` |
| **Arquivo** | RepositorioPlanilha.gs |
| **Assinatura** | `inserirLinha(valores)` — valores: array de valores na ordem das colunas (A até AC). Retorna void ou confirmação. |
| **Responsabilidade** | Inserir uma linha no fim da aba (appendRow ou setValues na próxima linha). |
| **Chama** | Configuracoes; SpreadsheetApp. |
| **Chamada por** | Service `cadastrarRegistro`. |
| **Origem** | Lógica de `salvarChamado` (appendRow). |

---

## 3.3 buscarLinhaPorId

| Item | Valor |
|------|--------|
| **Nome** | `buscarLinhaPorId` |
| **Arquivo** | RepositorioPlanilha.gs |
| **Assinatura** | `buscarLinhaPorId(id)` — id: string (UUID). Retorna array de valores (uma linha) ou null. |
| **Responsabilidade** | Localizar a linha cuja coluna ID contém o valor informado e retornar os valores dessa linha. |
| **Chama** | `lerDadosPlanilha` ou leitura direta; `indiceColunaId` (ou equivalente). |
| **Chamada por** | Service `obterRegistroPorId`, `atualizarRegistro`. |
| **Origem** | Lógica de `obterChamadoPorId` e de localização em `atualizarChamado`. |

---

## 3.4 buscarIndiceLinhaPorId

| Item | Valor |
|------|--------|
| **Nome** | `buscarIndiceLinhaPorId` |
| **Arquivo** | RepositorioPlanilha.gs |
| **Assinatura** | `buscarIndiceLinhaPorId(id)` — id: string. Retorna número (índice 1-based da linha na planilha) ou -1. |
| **Responsabilidade** | Encontrar o número da linha (para setValues) que contém o registro com o ID dado. |
| **Chama** | Leitura da planilha; identificação da coluna ID (Configuracoes ou helper). |
| **Chamada por** | Service `atualizarRegistro`; ou Repository `atualizarLinha` se receber o índice. |
| **Origem** | Parte da lógica de `atualizarChamado`. |

---

## 3.5 atualizarLinha

| Item | Valor |
|------|--------|
| **Nome** | `atualizarLinha` |
| **Arquivo** | RepositorioPlanilha.gs |
| **Assinatura** | `atualizarLinha(indiceLinha, valores)` — indiceLinha: número 1-based; valores: array na ordem das colunas. Retorna void. |
| **Responsabilidade** | Escrever os valores na linha indicada da aba (setValues em um range de uma linha). |
| **Chama** | Configuracoes; SpreadsheetApp. |
| **Chamada por** | Service `atualizarRegistro`. |
| **Origem** | Lógica de `atualizarChamado` (rowRange.setValues). |

---

## 3.6 obterColunasPadrao

| Item | Valor |
|------|--------|
| **Nome** | `obterColunasPadrao` |
| **Arquivo** | RepositorioPlanilha.gs (ou Configuracoes.gs se for só definição estática) |
| **Assinatura** | `obterColunasPadrao()` — sem parâmetros. Retorna array de { key, label } na ordem A–AC. |
| **Responsabilidade** | Definir o schema padrão da tabela quando a planilha não tem cabeçalho ou como fallback. |
| **Chama** | Configuracoes (para nome da coluna ID, se necessário). |
| **Chamada por** | Service `buscarRegistrosComFiltros` quando não houver cabeçalho. |
| **Origem** | `_getDefaultColumns_` em Código.js. |

---

## 3.7 indiceColunaId

| Item | Valor |
|------|--------|
| **Nome** | `indiceColunaId` |
| **Arquivo** | RepositorioPlanilha.gs |
| **Assinatura** | `indiceColunaId(cabecalhoOuUltimaColuna)` — cabecalho: linha de cabeçalho ou número da última coluna. Retorna número (índice 0-based da coluna ID). |
| **Responsabilidade** | Determinar em qual coluna está o ID (por nome no cabeçalho ou por convenção, ex.: última coluna). |
| **Chama** | Configuracoes (NOME_COLUNA_ID). |
| **Chamada por** | `buscarLinhaPorId`, `buscarIndiceLinhaPorId`, `preencherColunaId`. |
| **Origem** | `_getIdColumnIndexOnSheet_` (quando existia no fluxo de edição). |

---

## 3.8 preencherColunaId

| Item | Valor |
|------|--------|
| **Nome** | `preencherColunaId` |
| **Arquivo** | RepositorioPlanilha.gs |
| **Assinatura** | `preencherColunaId()` — sem parâmetros. Retorna objeto `{ ok, message }`. |
| **Responsabilidade** | Preencher a coluna ID para linhas que ainda não têm UUID; criar coluna de ID se necessário; escrever cabeçalho “ID” na primeira linha se houver cabeçalho. |
| **Chama** | Configuracoes; leitura/escrita da planilha; `indiceColunaId` ou lógica equivalente. |
| **Chamada por** | Controller `executarMigracaoIds`. |
| **Origem** | `migrarAdicionarId` em Código.js. |

---

# 4. Camada Data (configuração e listas)

Constantes e dados estáticos; sem lógica. Arquivo: **Configuracoes.gs** (ou Dados.gs).

---

## 4.1 Configuracoes (constante)

| Item | Valor |
|------|--------|
| **Nome** | `Configuracoes` (ou manter `SETTINGS`) |
| **Arquivo** | Configuracoes.gs |
| **Assinatura** | Objeto constante: `{ NOME_ABA, NOME_COLUNA_ID }` (ou ID_COLUMN_NAME). |
| **Responsabilidade** | Centralizar nome da aba da planilha e nome da coluna de ID para uso pelo Repository (e eventualmente Service). |
| **Chama** | Nenhuma. |
| **Chamada por** | RepositorioPlanilha.gs; eventualmente ServicoRegistro.gs. |
| **Origem** | `SETTINGS` em Config.js. |

---

## 4.2 obterOpcoesFormulario

| Item | Valor |
|------|--------|
| **Nome** | `obterOpcoesFormulario` |
| **Arquivo** | Configuracoes.gs |
| **Assinatura** | `obterOpcoesFormulario()` — sem parâmetros. Retorna objeto com CURSOS, TURMAS, SALAS, OFERTAS, RESPONSAVEIS, PRIORIDADES, STATUS, BOOLEANOS (arrays de strings). |
| **Responsabilidade** | Fonte única das listas estáticas usadas nos formulários e filtros; sem lógica, apenas retorno de constantes. |
| **Chama** | Nenhuma (retorna objeto literal ou constantes). |
| **Chamada por** | Controller `obterConfiguracaoFormulario`. |
| **Origem** | Conteúdo retornado por `getFormConfiguration` em Config.js. |

---

# 5. Resumo por arquivo

| Arquivo | Funções |
|---------|---------|
| **Main.gs** | doGet, obterUrlWebApp |
| **ControllerRegistro.gs** | obterConfiguracaoFormulario, pesquisarRegistros, salvarRegistro, atualizarRegistro, obterRegistroPorId, executarMigracaoIds |
| **ServicoRegistro.gs** | buscarRegistrosComFiltros, aplicarFiltros, aplicarOrdenacao, indiceDaColuna, linhaParaDados, dadosParaLinha, cadastrarRegistro, atualizarRegistro, obterRegistroPorId, interpretarDataIso, interpretarData, finalDoDia, interpretarNumero, normalizarColunasDoCabecalho |
| **RepositorioPlanilha.gs** | lerDadosPlanilha, inserirLinha, buscarLinhaPorId, buscarIndiceLinhaPorId, atualizarLinha, obterColunasPadrao, indiceColunaId, preencherColunaId |
| **Configuracoes.gs** | Configuracoes (constante), obterOpcoesFormulario |

---

# 6. Ajustes na UI (client-side)

Após a reestruturação, as chamadas `google.script.run` nos HTML/JS devem usar os novos nomes:

- `getWebAppUrl` → `obterUrlWebApp` (templates Index e Consulta).
- `getFormConfiguration` → `obterConfiguracaoFormulario` (JavaScript.html e ConsultaJavaScript.html).
- `pesquisarChamados` → `pesquisarRegistros` (ConsultaJavaScript.html).
- `salvarChamado` → `salvarRegistro` (JavaScript.html).
- Se existir: `obterChamadoPorId` → `obterRegistroPorId`; `atualizarChamado` → `atualizarRegistro`.

Nenhuma outra função do backend deve ser exposta diretamente ao cliente; apenas as listadas na seção 1 (Controller).
