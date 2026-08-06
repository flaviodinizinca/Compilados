// =================================================================
// --- BLOCO 4: FUNÇÃO MESTRA - CICLO COMPLETO (COM SENHA) ---
// =================================================================
function executarCicloCompleto() {
  const ui = SpreadsheetApp.getUi();
  
  // --- TRAVA DE SEGURANÇA (SENHA) ---
  const respostaSenha = ui.prompt(
    '🔐 Autenticação Necessária', 
    'Digite a senha de administrador para iniciar o Ciclo Completo:', 
    ui.ButtonSet.OK_CANCEL
  );

  if (respostaSenha.getSelectedButton() !== ui.Button.OK) return;
  
  const senhaDigitada = respostaSenha.getResponseText();
  
  // BUSCA A SENHA SEGURA
  const senhaCorreta = PropertiesService.getScriptProperties().getProperty('SENHA_ADMIN');

  if (senhaDigitada !== senhaCorreta) {
    ui.alert("❌ Senha Incorreta. Acesso negado.");
    return;
  }
  // -----------------------------------

  const result = ui.alert(
    'Confirmar Execução (Modo Otimizado)',
    'Esta ação irá:\n1. Ler a Fonte de Dados (Filtro >= 2023)\n2. Atualizar MATERIAIS e MEDICAMENTOS diretamente\n3. Compilar Dados Locais\n\nDeseja continuar?',
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) return;

  const toast = (msg) => SpreadsheetApp.getActiveSpreadsheet().toast(msg, "Orquestrador", 10);

  try {
    toast("Lendo Fonte de Dados Global (Isso pode levar alguns segundos)...");
    const dadosGlobais = obterDadosEntradasGlobal();
    
    if (!dadosGlobais || dadosGlobais.length === 0) {
      throw new Error("A busca de dados retornou vazia. Verifique se a planilha fonte está conectada.");
    }
    
    // Log para confirmar que o filtro funcionou
    console.log(`Sucesso: ${dadosGlobais.length} registros carregados na memória.`);

    toast("Iniciando Materiais (Remoto)...");
    processarMateriaisRemoto(dadosGlobais); 
    
    toast("Iniciando Medicamentos (Remoto)...");
    processarMedicamentosRemoto(dadosGlobais); 
    
    toast("Executando rotina Local...");
    compilarDados(dadosGlobais); 

    ui.alert("Ciclo Completo Finalizado com Sucesso!");
  } catch (e) {
    console.error(e);
    ui.alert("Erro durante o ciclo: " + e.message);
  }
}