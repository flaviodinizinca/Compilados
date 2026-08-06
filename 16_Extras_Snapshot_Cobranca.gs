// =================================================================
// --- BLOCO 16: NOVAS FUNCIONALIDADES (COBRANÇA) ---
// =================================================================

function gerarRascunhosCobranca() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaComp = ss.getSheetByName("Compilados");
  
  const lastRow = abaComp.getLastRow();
  if(lastRow < 2) return;

  const dados = abaComp.getRange(2, 1, lastRow-1, 22).getValues();
  
  let contadorDrafts = 0;
  const atualizacoes = []; 
  const hoje = new Date();
  
  console.log(`Iniciando verificação de ${dados.length} linhas para cobrança...`);

  dados.forEach((linha, index) => {
    const empenho = linha[0];
    const fornecedor = linha[4];
    
    const atrasoRaw = String(linha[17] || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, ""); 
    const status = String(linha[18] || "").toUpperCase(); 
    
    const ultCobranca = linha[21]; 
    
    let isAtrasado = false;
    
    if (atrasoRaw.includes("mes") || atrasoRaw.includes("ano")) {
      isAtrasado = true;
    } 
    else {
      const matchDias = atrasoRaw.match(/(\d+)\s*dia/);
      if (matchDias && parseInt(matchDias[1]) > 10) {
        isAtrasado = true;
      }
    }

    let podeCobrarNovamente = true;
    if (ultCobranca instanceof Date) {
      const diffDias = (hoje - ultCobranca) / (1000 * 60 * 60 * 24);
      if (diffDias < 15) podeCobrarNovamente = false; 
    }

    if (status.includes("PENDENTE") && isAtrasado && podeCobrarNovamente) {
      
      console.log(`GERAR: Linha ${index+2} | Emp: ${empenho} | Status: ${status} | Atraso: ${atrasoRaw}`);

      const assunto = `COBRANÇA: Empenho ${empenho} - ${fornecedor} (Atraso Detectado)`;
      const corpo = `
        Prezados ${fornecedor},
        
        Consta em nosso sistema o Empenho ${empenho} com status PENDENTE e entrega atrasada (${linha[17]}).
        
        Item: ${linha[5]} - ${linha[6]}
        Quantidade Pendente: ${linha[16]}
        
        Favor informar previsão de entrega urgente.
        
        Atenciosamente,
        Equipe de Gestão de Estoques
      `;
      
      GmailApp.createDraft(CONFIG.emails.para, assunto, corpo);
      
      contadorDrafts++;
      atualizacoes.push([index + 2, hoje]); 
    }
  });

  if (atualizacoes.length > 0) {
    atualizacoes.forEach(item => {
      abaComp.getRange(item[0], 22).setValue(item[1]).setNumberFormat("dd/mm/yyyy");
    });
    ui.alert(`Sucesso! ${contadorDrafts} rascunhos criados.\nDatas atualizadas na coluna V.`);
  } else {
    ui.alert("Nenhum item se enquadra nos critérios hoje.\n(Verifique se já foram cobrados há menos de 15 dias ou se o status não é Pendente)");
  }
}