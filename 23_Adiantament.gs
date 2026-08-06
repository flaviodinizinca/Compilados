// =================================================================
// --- BLOCO 23: ADIANTAMENTO E BUSCA DE E-MAILS ---
// =================================================================

/**
 * Varre a aba "Adiantamento" para preencher E-mails (buscando no Gmail)
 * e preencher a Descrição do Item (buscando na aba Cont.Estoque).
 */
function preencherEmailsFornecedores() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaAdiantamento = planilha.getSheetByName("Adiantamento");
  const abaEstoque = planilha.getSheetByName("Cont.Estoque");
  
  if (!abaAdiantamento || !abaEstoque) {
    SpreadsheetApp.getUi().alert("Abas 'Adiantamento' ou 'Cont.Estoque' não encontradas.");
    return;
  }

  // 1. Mapear Descrições da aba Cont.Estoque 
  const dadosEstoque = abaEstoque.getDataRange().getValues();
  const mapaDescricoes = {};
  
  for (let i = 1; i < dadosEstoque.length; i++) {
    // ATENÇÃO: Na guia Cont.Estoque, o Código está na Coluna B (índice 1) e a Descrição na Coluna C (índice 2)
    let codRaw = dadosEstoque[i][1]; 
    let desc = dadosEstoque[i][2];   
    if (codRaw) {
      let codLimpo = codRaw.toString().trim().toUpperCase();
      mapaDescricoes[codLimpo] = desc;
    }
  }

  // 2. Preencher a aba Adiantamento
  const dados = abaAdiantamento.getDataRange().getValues();
  
  for (let i = 1; i < dados.length; i++) {
    let codigoRaw = dados[i][0];       // Coluna A
    let descricaoAtual = dados[i][1];  // Coluna B
    let nomeEmpresa = dados[i][2];     // Coluna C
    let emailAtual = dados[i][5];      // Coluna F
    
    // --- PREENCHER DESCRIÇÃO (COLUNA B) ---
    if (codigoRaw && (!descricaoAtual || descricaoAtual === "Não encontrado no Estoque")) {
      let codigoLimpo = codigoRaw.toString().trim().toUpperCase();
      let descEncontrada = mapaDescricoes[codigoLimpo] || "Não encontrado no Estoque";
      abaAdiantamento.getRange(i + 1, 2).setValue(descEncontrada);
    }

    // --- PREENCHER E-MAIL (COLUNA F) ---
    // Permite reprocessar se estiver vazio ou se falhou antes ("Não encontrado")
    if (nomeEmpresa && (!emailAtual || emailAtual === "Não encontrado")) {
      let emailsEncontrados = buscarNoGmail(nomeEmpresa);
      let celulaDestino = abaAdiantamento.getRange(i + 1, 6); 
      
      celulaDestino.clearDataValidations();
      celulaDestino.clearNote();
      
      if (emailsEncontrados.length > 0) {
        if (emailsEncontrados.length === 1) {
          celulaDestino.setValue(emailsEncontrados[0]);
        } else {
          let todosJuntos = emailsEncontrados.join(", ");
          let listaDropdown = [todosJuntos, ...emailsEncontrados]; 
          
          let regraValidacao = SpreadsheetApp.newDataValidation()
            .requireValueInList(listaDropdown, true)
            .build();
            
          celulaDestino.setDataValidation(regraValidacao);
          celulaDestino.setValue(todosJuntos); 
          celulaDestino.setNote("Vários e-mails encontrados! O padrão envia para todos. Clique na seta se quiser escolher um específico.");
        }
      } else {
        celulaDestino.setValue("Não encontrado");
        celulaDestino.setNote("Nenhum e-mail externo encontrado para este termo no seu Gmail.");
      }
    }
  }
}

/**
 * Função auxiliar: Vasculha as últimas 20 conversas no Gmail contendo o nome da empresa
 */
function buscarNoGmail(termoBusca) {
  const threads = GmailApp.search('"' + termoBusca + '"', 0, 20);
  const emailsUnicos = new Set();
  const regexEmail = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;

  for (let i = 0; i < threads.length; i++) {
    let mensagens = threads[i].getMessages();
    for (let j = 0; j < mensagens.length; j++) {
      let remetente = mensagens[j].getFrom();
      let destinatarios = mensagens[j].getTo();
      let copias = mensagens[j].getCc();
      
      let textoCompleto = remetente + " " + destinatarios + " " + copias;
      let ocorrencias = textoCompleto.match(regexEmail);
      
      if (ocorrencias) {
        ocorrencias.forEach(email => {
          let emailLimpo = email.toLowerCase();
          if (emailLimpo.indexOf('@inca.gov.br') === -1) {
            emailsUnicos.add(emailLimpo);
          }
        });
      }
    }
  }
  return Array.from(emailsUnicos);
}

/**
 * Gera os rascunhos no Gmail agrupados por fornecedor, 
 * buscando a Descrição e o CMM oficiais na aba Cont.Estoque.
 */
function gerarRascunhosAdiantamento() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaAdiantamento = planilha.getSheetByName("Adiantamento");
  const abaEstoque = planilha.getSheetByName("Cont.Estoque");
  
  if (!abaAdiantamento || !abaEstoque) {
    SpreadsheetApp.getUi().alert("Abas 'Adiantamento' ou 'Cont.Estoque' não encontradas.");
    return;
  }

  // Lista de e-mails em cópia conforme solicitado
  const CC_EMAILS = "cobranca.disup@inca.gov.br, elmoraes@inca.gov.br, andrea.veiga@inca.gov.br, jose.sousa@inca.gov.br, thiago.gomes@inca.gov.br, disupempenhos01@inca.gov.br, larissa.hilario@inca.gov.br, karine.teodoro@inca.gov.br, lmendonca@inca.gov.br, jssampaio@inca.gov.br";

  // 1. Mapear os dados da aba Cont.Estoque
  const dadosEstoque = abaEstoque.getDataRange().getValues();
  const mapaEstoque = {};
  
  for (let i = 1; i < dadosEstoque.length; i++) {
    let codRaw = dadosEstoque[i][1];        // Coluna B (índice 1) é o Código
    let descricaoItem = dadosEstoque[i][2]; // Coluna C (índice 2) é a Descrição
    let valorCMM = dadosEstoque[i][14];     // Coluna O (índice 14) é o CMM Antigo
    
    if (codRaw) {
      let codLimpo = codRaw.toString().trim().toUpperCase();
      mapaEstoque[codLimpo] = {
        descricao: descricaoItem || "Descrição não cadastrada",
        cmm: Math.ceil(Number(valorCMM)) || 0
      };
    }
  }

  // 2. Coletar os itens da aba Adiantamento
  const dadosAdiantamento = abaAdiantamento.getDataRange().getValues();
  let fornecedores = {};
  let rascunhosGerados = 0;

  for (let i = 1; i < dadosAdiantamento.length; i++) {
    let codigoRaw = dadosAdiantamento[i][0];   // Coluna A
    let empresa = dadosAdiantamento[i][2];     // Coluna C
    let processo = dadosAdiantamento[i][3];    // Coluna D
    let email = dadosAdiantamento[i][5];       // Coluna F

    if (!codigoRaw || !empresa) continue;

    let codigoLimpo = codigoRaw.toString().trim().toUpperCase();
    let infoOficial = mapaEstoque[codigoLimpo] || { descricao: "Item não encontrado no estoque", cmm: 0 };
    
    if (!fornecedores[empresa]) {
      fornecedores[empresa] = { email: email, processo: processo, itens: [] };
    }
    
    // Montando a linha do item exatamente como no seu modelo (em formato HTML para lista)
    fornecedores[empresa].itens.push(`<li><b>${codigoLimpo}</b> - ${infoOficial.descricao}. (Quantidade: ${infoOficial.cmm} un)</li>`);
  }

  // 3. Gerar os rascunhos no Gmail
  let erros = [];
  
  for (let nomeEmpresa in fornecedores) {
    let f = fornecedores[nomeEmpresa];
    let emailDestino = f.email && f.email !== "Não encontrado" ? f.email : "";
    
    // Monta o corpo do e-mail com formatação HTML para garantir a estética da lista e parágrafos
    let corpoHtml = `
      <p style="font-family: Arial, sans-serif; font-size: 14px;">Por orientação da chefia da Divisão de Suprimentos do INCA que nos lê em cópia.</p>
      <p style="font-family: Arial, sans-serif; font-size: 14px;">Considerando-se que a empresa <strong>${nomeEmpresa}</strong> sagrou-se vencedora para o processo de compras <strong>${f.processo}</strong>, solicito verificar a possibilidade de envio como adiantamento dos itens abaixo no almoxarifado:</p>
      <ul style="font-family: Arial, sans-serif; font-size: 14px;">
        ${f.itens.join("")}
      </ul>
      <br>
      <p style="font-family: Arial, sans-serif; font-size: 14px;">
        <strong>Endereço para entrega:</strong><br>
        Rua André Cavalcanti, 37 – Centro/RJ<br>
        CEP: 20231-050<br>
        Horário de atendimento: segunda a sexta-feira, das 9h às 15h<br>
        Almoxarifado: (21) 3207-6627 / 6628 / 6629 / 6630<br>
        Serviço de Patrimônio: (21) 3207-4635 / 4648 / 4638
      </p>
      <p style="font-family: Arial, sans-serif; font-size: 14px;">Desde já, agradecemos pela atenção e nos colocamos à disposição para quaisquer esclarecimentos adicionais.</p>
      <p style="font-family: Arial, sans-serif; font-size: 14px;">Atenciosamente,</p>
    `;

    // Texto sem HTML como redundância, que o Gmail exige por segurança
    let corpoTextoPuro = corpoHtml.replace(/<[^>]+>/g, '\n').replace(/\n\s*\n/g, '\n');

    try {
      GmailApp.createDraft(
        emailDestino, 
        `Solicitação de Adiantamento - ${nomeEmpresa} - ${f.processo}`, 
        corpoTextoPuro, 
        { 
          cc: CC_EMAILS,
          htmlBody: corpoHtml 
        }
      );
      rascunhosGerados++;
    } catch (e) {
      erros.push(`Falha ao gerar rascunho para ${nomeEmpresa}. Erro: ${e.message}`);
    }
  }

  // 4. Feedback final para o usuário
  if (erros.length > 0) {
    SpreadsheetApp.getUi().alert(`Gerados ${rascunhosGerados} rascunhos.\n\nAtenção, ocorreram erros:\n${erros.join("\n")}`);
  } else {
    SpreadsheetApp.getUi().alert(`Sucesso! ${rascunhosGerados} rascunho(s) foram gerados na sua caixa de Rascunhos do Gmail com o layout solicitado e a lista de cópias preenchida.`);
  }
}