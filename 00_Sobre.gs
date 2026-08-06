/**
 * ============================================================================
 * 📘 MANUAL TÉCNICO ANALÍTICO - ORQUESTRADOR GERAL (SISTEMA INCA)
 * ============================================================================
 * @project  Orquestrador de Estoque, Empenhos, Distribuição e Inteligência
 * @author   Flavio Garcia Diniz
 * @version  16.0 (Edição "Documentação Mestra")
 * @date     2026-02-06
 * * ESTE DOCUMENTO DESCREVE A LÓGICA, CÁLCULOS, ARQUIVOS E FLUXOS DO SISTEMA.
 * ============================================================================
 * * 1. 🏗️ ARQUITETURA DE DADOS (HUB & SPOKE)
 * ----------------------------------------------------------------------------
 * O sistema funciona como um "Hub" central (este script) que orquestra dados
 * de 6 fontes distintas para criar uma "Única Fonte de Verdade".
 * * [FONTES DE DADOS CONECTADAS - IDs em 01_Config.gs]
 * A. FONTE GERAL (EMS): 
 * - Origem: ERP Corporativo.
 * - Abas Leitura: "EntradaEmpenhos", "OutrasEntradas", "DadosEstoque".
 * - Função: Fornece a base financeira, logística e datas de validade.
 * * B. MOVIMENTO DE ESTOQUE (HISTÓRICO):
 * - ID Config: movimentoEstoque
 * - Aba Leitura: "Mov2022-2025" (Configurável).
 * - Função: Base para cálculo de Consumo Médio (CMA) e Previsão.
 * * C. ENTRADAS MANUAIS (LEGADO/APOIO):
 * - Materiais (ID: materiais): Inputs da equipe de Almoxarifado.
 * - Medicamentos (ID: medicamentos): Inputs da equipe de Farmácia.
 * - Função: Captura recebimentos provisórios e observações manuais.
 * * D. AUDITORIA E SAÍDA:
 * - Correção Externa (ID: correcaoExterna): Lista de empenhos/códigos para ajuste.
 * - Painel Equipe (ID: painelEquipe): Destino final dos dados para os planejadores.
 * * ============================================================================
 * * 2. 🧠 MOTORES DE CÁLCULO E LÓGICA
 * ----------------------------------------------------------------------------
 * * 2.1. ALGORITMO DE STATUS UNIFICADO (Status do Item)
 * O sistema classifica cada item seguindo uma hierarquia estrita de validação:
 * * 1. RECEBIDO A MAIOR (CRÍTICO):
 * - Lógica: (Qtd Empenho > 0) E (Saída Oficial > Qtd Empenho)
 * - Ação: Bloqueio de pagamento e alerta imediato.
 * * 2. CONCLUÍDO (SUCESSO):
 * - Lógica: (Qtd Empenho > 0) E (Saída Oficial == Qtd Empenho)
 * - Significado: Entrega finalizada perfeitamente.
 * * 3. ERRO DE CADASTRO (FANTASMA):
 * - Lógica: (Qtd Empenho == 0) E (Saída Oficial > 0)
 * - Significado: Item recebido sem empenho vinculado no sistema.
 * * 4. AGUARDANDO ENTREGA (PENDENTE):
 * - Lógica: (Qtd Empenho > 0) E (Saída Oficial == 0)
 * - Ação: Item monitorado para cobrança (Entra no cálculo de atraso).
 * * 5. ENTREGA PARCIAL (EM ANDAMENTO):
 * - Lógica: (Qtd Empenho > 0) E (Saída Oficial < Qtd Empenho)
 * - Ação: Calcula saldo residual e mantém monitoramento.
 * * 2.2. CÁLCULO DE CMA E COBERTURA (Gestão de Estoque)
 * Definido em: 08_Gestao_Estoque.gs
 * * - Período de Análise: 36 Meses (Configurável em 'MESES_ANALISE').
 * - Tipos de Movimento Considerados (Consumo Real):
 * ['REQ', 'RM', 'RFS', 'RDD', 'RCS', 'ACA', 'NFS']
 * - Tipos de Movimento Deduzidos (Devoluções):
 * ['DEV', 'DRM', 'RRQ']
 * - Fórmula Cobertura: (Saldo Atual / Consumo Médio Mensal) * 30 dias.
 * * 2.3. LÓGICA DE VALIDADE DE ATAS
 * Definido em: 09_Relatorios_Locais.gs
 * - Filtro: Compara data de hoje com data de vencimento da Ata.
 * - Saída: Relatório "Validade De Atas" com dias restantes.
 * * ============================================================================
 * * 3. 👥 GESTÃO DE EQUIPE E NOMES (HARDCODED)
 * ----------------------------------------------------------------------------
 * O sistema distribui as demandas automaticamente para os seguintes planejadores
 * definidos no arquivo '15_Distribuicao_Equipe.gs':
 * * - Bianca
 * - Katia
 * - Leonardo
 * - Moises
 * - Rafaelle
 * - Luciana
 * * * A interface (HTML) permite seleção múltipla via checkboxes.
 * * A distribuição cruza dados do Painel Equipe com as regras carregadas.
 * * ============================================================================
 * * 4. 📂 MAPA DE ARQUIVOS DO PROJETO (SCRIPTS)
 * ----------------------------------------------------------------------------
 * * [CONFIGURAÇÃO E NÚCLEO]
 * - 00_Sobre.gs ...............: Este manual técnico.
 * - 01_Config.gs ..............: Armazena IDs de planilhas e nomes de abas (Seguro).
 * - 02_Menu.gs ................: Cria o menu "Orquestrador" na planilha Google.
 * - 03_Helpers.gs .............: Funções auxiliares de formatação e normalização.
 * - 04_Ciclo_Completo.gs ......: Gatilho mestre. Exige senha de Admin para rodar tudo.
 * * [PROCESSAMENTO DE DADOS]
 * - 05_Materiais.gs ...........: Busca dados da planilha de Materiais (Almoxarifado).
 * - 06_Medicamentos.gs ........: Busca dados da planilha de Medicamentos (Farmácia).
 * - 07_Compilacao_Local.gs ....: O "Liquidificador". Cruza dados globais com locais.
 * - 08_Gestao_Estoque.gs ......: Calcula médias históricas e insere dados de estoque.
 * * [RELATÓRIOS E SAÍDAS]
 * - 09_Relatorios_Locais.gs ...: Gera relatórios específicos (ex: Validade Atas).
 * - 11_Sincronizacao_Externa.gs: Verifica divergências entre EMS e controles externos.
 * - 12_Dashboard.gs ...........: Gera gráficos e contagens de status (KPIs).
 * - 13_Relatorios_BI.gs .......: Prepara dados para exportação para Power BI/Looker.
 * * [EQUIPE E AUTOMAÇÃO]
 * - 15_Distribuicao_Equipe.gs .: Distribui tarefas entre os planejadores listados.
 * - 20_Buscar_Planejadores.gs .: Ferramenta de busca rápida por código no painel.
 * * ============================================================================
 * * 5. 🔒 SEGURANÇA
 * ----------------------------------------------------------------------------
 * - O "Ciclo Completo" é protegido por senha (verificado em ScriptProperties).
 * - IDs de arquivos sensíveis são chamados via referência, não expostos diretamente.
 * * ============================================================================
 */