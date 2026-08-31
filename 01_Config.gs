// =================================================================
// --- BLOCO 1: CONFIGURAÇÃO GLOBAL (SEGURA) ---
// =================================================================

var CONFIG = {
  ids: {
    // Fonte Geral (DadosGlobais)
    fonteDadosGeral: PropertiesService.getScriptProperties().getProperty('ID_FONTE_GERAL'),
    
    // Nova planilha para arquivamento (Movimento_Estoque)
    movimentoEstoque: PropertiesService.getScriptProperties().getProperty('ID_MOVIMENTO_ESTOQUE'),
    
    // IDs de legado/apoio mantidos
    materiais: PropertiesService.getScriptProperties().getProperty('ID_MATERIAIS'),
    medicamentos: PropertiesService.getScriptProperties().getProperty('ID_MEDICAMENTOS'),
    correcaoExterna: PropertiesService.getScriptProperties().getProperty('ID_CORRECAO'),
    compiladosLocal: PropertiesService.getScriptProperties().getProperty('ID_COMPILADOS'),
    painelEquipe: PropertiesService.getScriptProperties().getProperty('ID_PAINEL_EQUIPE')
  },
  
  abas: {
    // --- Novos Nomes na Fonte Geral (DadosGlobais) ---
    entradaEmpenhos: "EntradaEmpenhos", // Antiga "dados"
    outrasEntradas: "OutrasEntradas",   // Antiga "Outras Entradas"
    dadosEstoque: "DadosEstoque",       // Veio da planilha Compilados
    gruposEstoque: "GruposEstoque",     // Nova guia de grupos (Col A: Cod, Col B: Nome)
    itensGeladeira: "ItensGeladeira",   // <--- NOVA GUIA ADICIONADA AQUI
    
    // --- Na planilha Movimento_Estoque ---
    mov2022_2025: "Mov2022-2025",
    
    // --- Mantidas ---
    configEquipe: "Config_Equipe",
    materiais: "Material",
    medicamentos: "Medicamentos",
    entradas: "Entradas",
    empenhos: "Empenhos Enviados",
    recProvisorio: "Rec.Provisorio",
    estoqueRemoto: "Cont.Estoque"
  },
  
  destino: {
    nomeAba: "Compilados",
    colunaControleCobranca: 22
  },
  
  emails: {
    para: PropertiesService.getScriptProperties().getProperty('EMAIL_GESTAO_PARA'),
    copia: PropertiesService.getScriptProperties().getProperty('EMAIL_GESTAO_COPIA'),
    mutiraoLista: PropertiesService.getScriptProperties().getProperty('EMAIL_MUTIRAO_LISTA'),
    mutiraoTeste: PropertiesService.getScriptProperties().getProperty('EMAIL_MUTIRAO_TESTE')
  },
  
  // Configurações para a Manutenção de Dados
  manutencao: {
    limiteLinhasAlerta: 900000,
    anosParaArquivar: [2022, 2023, 2024, 2025]
  },

  cores: {
    'CONCLUÍDO': '#d9ead3',                   
    'PENDENTE': '#f4cccc',                    
    'PENDENTE COM RESÍDUO': '#fff2cc',      
    'RESÍDUO 10%': '#cfe2f3',                 
    'SOLICITAR ASSOCIAÇÃO': '#ffe599', 
    'RECEBIDO. FALTA ASSOCIAR': '#b4a7d6',
    'ELIMINADA': '#999999',                   
    'RECEBIMENTO PROVISÓRIO': '#fce5cd',
    'REC. PROV. / COM RESIDUO': '#f9cb9c',
    'RECEBIDO A MAIOR': '#f6b26b',
    'SALDO CANCELADO': '#ead1dc',
    'ALERTA_CRITICO': '#ea9999',
    'ALERTA_ATENCAO': '#ffe599',
    'ALERTA_OK': '#b6d7a8'
  }
};

// Função de validação rápida para garantir que os IDs essenciais existem
function validarConfiguracao() {
  const ids = CONFIG.ids;
  const erros = [];
  
  if (!ids.fonteDadosGeral) erros.push("ID_FONTE_GERAL não configurado.");
  if (!ids.movimentoEstoque) erros.push("ID_MOVIMENTO_ESTOQUE não configurado.");
  
  if (erros.length > 0) {
    throw new Error("Erro de Configuração:\n" + erros.join("\n"));
  }
  Logger.log("Configuração validada. IDs carregados corretamente.");
}