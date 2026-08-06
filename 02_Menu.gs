// =================================================================
// --- BLOCO 2: MENU PERSONALIZADO (DEFINITIVO) ---
// =================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Orquestrador Geral')
  
    // --- BLOCO 1: PROCESSAMENTO PRINCIPAL ---
    .addItem('🔄 1. Executar CICLO COMPLETO (Tudo)', 'executarCicloCompleto')
    .addSeparator()
    .addItem('📦 2. Processar Materiais (Remoto)', 'processarMateriaisRemoto')
    .addItem('💊 3. Processar Medicamentos (Remoto)', 'processarMedicamentosRemoto')
    .addItem('📊 4. Apenas Compilar Dados (Local)', 'compilarDadosLocal')
    .addSeparator()

    // --- NOVO BLOCO: VERIFICAR SALDOS ---
    .addSubMenu(ui.createMenu('✅ Verificar Saldos')
        .addItem('📦 Verificar Saldos em Material', 'verificarSaldosMaterial')
        .addItem('💊 Verificar Saldos em Medicamentos', 'verificarSaldosMedicamentos')
        .addSeparator()
        .addItem('🔄 Verificar Saldos em Material e Medicamentos', 'verificarSaldosMaterialEMedicamentos'))
    .addSeparator()

    // --- BLOCO 2: DISTRIBUIÇÃO E EQUIPE ---
    .addSubMenu(ui.createMenu('👥 6. Distribuir para Equipe')
        .addItem('✅ Atualizar TODOS (Completo)', 'atualizarTodos')
        .addSeparator()
        .addItem('👤 Lorena', 'atualizarLorena')
        .addItem('👤 Katia', 'atualizarKatia')
        .addItem('👤 Leonardo', 'atualizarLeonardo')
        .addItem('👤 Moises', 'atualizarMoises')
        .addItem('👤 Rafaelle', 'atualizarRafaelle')
        .addItem('👤 Luciana', 'atualizarLuciana'))
        
    .addItem('🔍 7. Localizar Item (Qual Planejador?)', 'localizarItemNoPainelEquipe')
    .addSeparator()

    // --- BLOCO 3: INTELIGÊNCIA ---
    .addSubMenu(ui.createMenu('🧠 Inteligência & Automação')
        .addItem('📨 Gerar Rascunhos de Cobrança (Gmail)', 'gerarRascunhosCobranca')
        .addSeparator()
        .addItem('🎯 Preencher E-mails de Adiantamento', 'preencherEmailsFornecedores')
        .addItem('📝 Gerar Rascunhos de Adiantamento', 'gerarRascunhosAdiantamento'))
    .addSeparator()

    // --- BLOCO 4: RELATÓRIOS CHEFIA ---
    .addSubMenu(ui.createMenu('💼 Relatórios Gerenciais (Chefia)')
        .addItem('🏆 Ranking de Fornecedores (Performance)', 'gerarRelatorioPerformanceFornecedores'))
    .addSeparator()

    // --- BLOCO 5: GESTÃO DE ESTOQUE ---
    .addSubMenu(ui.createMenu('📈 Gestão de Estoque')
        .addItem('🔄 Sincronizar Controle de Estoque', 'sincronizarControleEstoque')
        .addItem('📊 Dashboard de Status', 'gerarDashboardStatus'))
    .addSeparator()

    // --- BLOCO 6: RELATÓRIOS OPERACIONAIS ---
    .addSubMenu(ui.createMenu('📑 Relatórios Operacionais')
        .addItem('📥 Importar Urgências e Gerar Status Report', 'fluxoImportacaoUrgencias')
        .addItem('📊 Apenas Processar Status Report (Atual)', 'processarMutirao')
        .addSeparator()
        .addItem('Rel. Validade de Atas', 'gerarRelatorioValidadeAtas')
        .addItem('Rel. Valor Resíduo 10%', 'gerarRelatorioResiduo10')
        .addItem('Rel. Itens em atraso >10', 'gerarRelatorioAtrasos'))
    .addSeparator()

    // --- BLOCO 7: SINCRONIZAÇÃO EXTERNA ---
    .addSubMenu(ui.createMenu('🔗 Sincronização Externa')
        .addItem('1. Analisar Divergências', 'buscarEmpenhosCodigosErrados')
        .addItem('2. Enviar Itens Faltantes', 'sincronizarEmpenhosNaExterna')
        .addItem('3. Reparar Dados Vazios', 'repararDadosFaltantesNaExterna'))
    .addSeparator()

    // --- BLOCO 8: FERRAMENTAS GERAIS ---
    .addItem('📄 Gerar PDF por Status', 'abrirMenuGerarPDF')
    .addSeparator()

    // --- BLOCO 9: RESTOS A PAGAR ---
    .addSubMenu(ui.createMenu('💰 Restos a Pagar e Empenhos a Liquidar')
        .addItem('⚙️ Processar Restos a Pagar', 'processarRestosAPagar')
        .addItem('⚙️ Processar Empenhos a Liquidar', 'processarEmpenhosALiquidar')
        .addSeparator()
        .addItem('📥 Importar respostas dos planejadores', 'importarRespostasPlanejadoresParaNaturezas')
        .addItem('📊 Verificar Preenchimento (Atualizar %)', 'verificarPreenchimentoRestos')
        .addItem('🔗 Verificar e Atualizar Links', 'verificarAtualizarLinksRestos'))
    
    .addToUi();
}