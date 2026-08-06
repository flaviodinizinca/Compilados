// =================================================================
// --- BLOCO 9: RELATÓRIOS LOCAIS E COBRANÇA AUTOMÁTICA ---
// =================================================================

function gerarRelatorioAtrasos() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaComp = ss.getSheetByName("Compilados");
    let abaAtraso = ss.getSheetByName("Atraso>10");
    
    if (!abaComp) throw new Error("Aba Compilados não encontrada.");
    if (!abaAtraso) abaAtraso = ss.insertSheet("Atraso>10");
    
    abaAtraso.clear();

    const dados = abaComp.getRange("A1:U" + abaComp.getLastRow()).getValues();
    const body = dados.slice(1);
    const dadosFilt = [];
    
    // Cache para não buscar o mesmo fornecedor várias vezes no Gmail (ganho de performance)
    const cacheEmails = {};
    
    body.forEach(l => {
      const st = _norm(l[18]);
      let incluir = true;

      // Filtros de regra de negócio
      if (['RESÍDUO 10%', 'CONCLUÍDO', 'ELIMINADA'].includes(st)) incluir = false;
      const dias = parseAtrasoParaDias(l[17] ? String(l[17]) : "");
      if (dias < 11) incluir = false;
      
      const code = _norm(l[5]);
      if (['A3', 'A5', 'A7', 'C', 'D', 'P', 'S'].some(p => code.startsWith(p))) 
        incluir = false;
      if (_norm(l[1]) === 'MAI') incluir = false;
      if (['A+', 'NÃO COBRAR', 'CÓDIGO MAI'].includes(_norm(l[12]))) incluir = false;
      if (_norm(l[6]).match(/REAGENTE|COMPRESSIVA|^PRÓTESE/)) incluir = false;
      
      if (incluir) {
        const nomeFornecedor = String(l[4]).trim();
        let emailEncontrado = "";

        // Lógica de busca automática igual ao Adiantamento
        if (nomeFornecedor && nomeFornecedor !== "NÃO INFORMADO") {
          if (cacheEmails[nomeFornecedor]) {
            emailEncontrado = cacheEmails[nomeFornecedor];
          } else {
            // Busca no Gmail por mensagens enviadas para este fornecedor
            const threads = GmailApp.search('to:' + nomeFornecedor, 0, 1);
            if (threads.length > 0) {
              const msg = threads[0].getMessages();
              emailEncontrado = msg[msg.length - 1].getTo();
            }
            cacheEmails[nomeFornecedor] = emailEncontrado;
          }
        }

        dadosFilt.push([
          l[0],  // A: EMPENHO
          nomeFornecedor, // B: FORNECEDOR
          l[5],  // C: ITEM
          l[6],  // D: DESCRIÇÃO
          l[15], // E: QTD EMPENHO
          l[16], // F: QTD RECEBIDA
          (parseFloat(l[15]) - parseFloat(l[16])) || 0, // G: QTD RESIDUAL
          l[17], // H: TEMPO EM ATRASO
          l[19], // I: PROCESSO
          l[20], // J: MODALIDADE
          emailEncontrado // K: Email (Busca Automática)
        ]);
      }
    });

    const headerNovo = ["EMPENHO", "FORNECEDOR", "ITEM", "DESCRIÇÃO", "QTD EMPENHO", "QTD RECEBIDA", "QTD RESIDUAL", "TEMPO EM ATRASO", "PROCESSO", "MODALIDADE", "Email"];
    abaAtraso.getRange(1, 1, 1, 11).setValues([headerNovo]).setFontWeight('bold').setBackground("#cfe2f3");

    if (dadosFilt.length > 0) {
      abaAtraso.getRange(2, 1, dadosFilt.length, 11).setValues(dadosFilt);
      abaAtraso.setFrozenRows(1);
      abaAtraso.autoResizeColumns(1, 11);

      ui.alert("Relatório e E-mails Gerados", `${dadosFilt.length} itens encontrados. Os rascunhos de cobrança serão gerados agora.`, ui.ButtonSet.OK);
      // Chama a geração de rascunhos automaticamente ao finalizar o relatório
      gerarRascunhosCobrancaAtraso();
    } else {
      ui.alert("Info", "Nenhum item em atraso encontrado.", ui.ButtonSet.OK);
    }
  } catch (e) { 
    ui.alert("Erro", e.message, ui.ButtonSet.OK); 
  }
}

function gerarRascunhosCobrancaAtraso() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("Atraso>10");
  if (!aba || aba.getLastRow() < 2) return;

  const dados = aba.getRange(2, 1, aba.getLastRow() - 1, 11).getValues();
  const agrupado = {};
  
  // Agrupamento por Fornecedor (Coluna B) para evitar múltiplos e-mails
  dados.forEach(linha => {
    const fornecedor = linha[1];
    const email = linha[10];
    if (!fornecedor || !email) return;

    if (!agrupado[fornecedor]) {
      agrupado[fornecedor] = { email: email, processo: linha[8], itens: [] };
    }
    agrupado[fornecedor].itens.push({
      empenho: linha[0],
      descricao: linha[3],
      qtd: linha[6],
      atraso: linha[7]
    });
  });
  
  for (let forn in agrupado) {
    const info = agrupado[forn];
    let listaItens = info.itens.map(it => 
      `• Nota de Empenho nº ${it.empenho}: pendência no fornecimento de ${it.qtd} unidades do item "${it.descricao}" (Atraso: ${it.atraso}).`
    ).join("\n");
    
    const assunto = `Cobrança de Entrega - Processo Administrativo nº ${info.processo}`;
    const corpo = `Prezados Senhores,\n\nEm nome da chefia de suprimentos que nos lê em cópia, o Instituto Nacional de Câncer (INCA) entra em contato com a empresa ${forn}, para tratar de uma questão referente à execução dos contratos (Processo Administrativo nº ${info.processo}).\n\nEscrevemos para solicitar um posicionamento sobre a entrega da(s) seguinte(s) Nota(s) de Empenho:\n\n${listaItens}\n\nÉ de nosso dever ressaltar a importância crítica destes itens para a continuidade do tratamento de pacientes oncológicos.\n\nDiante do exposto, solicitamos que a empresa nos encaminhe, no prazo de 48 (quarenta e oito) horas, um esclarecimento formal e um cronograma para a regularização imediata.\n\nA ausência de um retorno conclusivo demandará a adoção das medidas administrativas cabíveis conforme a Lei nº 14.133/2021.\n\nAtenciosamente,\n\nSetor de Suprimentos - INCA`;
    
    GmailApp.createDraft(info.email, assunto, corpo, { cc: "suprimentos@inca.gov.br" }); // Exemplo de CC conforme seu padrão
  }
}