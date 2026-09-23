// =================================================================
// --- BLOCO 14: STATUS REPORT (ESTRUTURA ATUALIZADA - 15 COLUNAS) ---
// =================================================================

// CONFIGURAÇÃO GERAL
const NOME_ABA_TRABALHO = "Status Report";
const CONFIG_EMAILS_MUTIRAO = {
  MODO_TESTE: false, 
  EMAIL_TESTE: PropertiesService.getScriptProperties().getProperty('EMAIL_MUTIRAO_TESTE'), 
  get LISTA_GERAL() {
    const listaTexto = PropertiesService.getScriptProperties().getProperty('EMAIL_MUTIRAO_LISTA');
    return listaTexto ? listaTexto.split(',').map(e => e.trim()) : [];
  }
};

function fluxoImportacaoUrgencias() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaStatus = ss.getSheetByName(NOME_ABA_TRABALHO);

  if (!abaStatus) {
    ui.alert("Erro", `Aba '${NOME_ABA_TRABALHO}' não encontrada.`, ui.ButtonSet.OK);
    return;
  }

  try {
    const idUrgencias = '1b8pSeKilEIkzyDgLjHwJpslXczIQ7Ki_4qDpnXbiEkM';
    const ssUrg = SpreadsheetApp.openById(idUrgencias);
    const abaUrg = ssUrg.getSheetByName('Urgências');
    
    if (!abaUrg) throw new Error("Aba 'Urgências' não encontrada na planilha de origem.");
    
    const lastRowUrg = abaUrg.getLastRow();
    if (lastRowUrg < 2) throw new Error("A planilha de Urgências está vazia.");
    
    const dadosUrg = abaUrg.getRange(2, 1, lastRowUrg - 1, 1).getValues();
    const colA = [];
    
    dadosUrg.forEach(linha => {
       const cod = linha[0];
       if (cod !== "" && cod != null) {
          colA.push([cod]);
       }
    });

    if (colA.length === 0) {
      throw new Error("Nenhum código válido encontrado na coluna A da aba de Urgências.");
    }
    
    // Limpa a aba e insere apenas os códigos na nova estrutura
    const maxRows = abaStatus.getMaxRows();
    if(maxRows > 1) {
        abaStatus.getRange(2, 1, maxRows - 1, abaStatus.getMaxColumns()).clearContent();
    }
    
    abaStatus.getRange(2, 1, colA.length, 1).setValues(colA);
    SpreadsheetApp.flush();
    processarMutirao();

  } catch(e) {
    ui.alert("Erro na Importação", e.message, ui.ButtonSet.OK);
  }
}

function processarMutirao() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: 'Segoe UI', sans-serif; padding: 20px; background-color: #f3f3f3; }
      h3 { margin-top: 0; color: #333; }
      .container { background: white; padding: 15px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
      label { display: block; margin-bottom: 12px; cursor: pointer; font-size: 15px; color: #444; }
      input[type="radio"] { transform: scale(1.3); margin-right: 10px; cursor: pointer; accent-color: #1c4587; }
      .btn {
          background-color: #1c4587; color: white; border: none;
          padding: 12px 0; width: 100%; border-radius: 5px;
          font-size: 16px; font-weight: bold; cursor: pointer; margin-top: 15px;
          transition: background 0.3s;
      }
      .btn:hover { background-color: #0f2e5e; }
      .loading { display: none; color: #666; font-size: 13px; text-align: center; margin-top: 10px; }
    </style>
    <div class="container">
      <h3>Definir Origem</h3>
      <p>Estes itens pertencem a qual grupo?</p>
      <form id="formOrigem">
        <label><input type="radio" name="opcao" value="MUTIRAO" checked> <b>Mutirão</b></label>
        <label><input type="radio" name="opcao" value="ACAO"> <b>Grupo Ação</b></label>
        <button type="button" class="btn" onclick="enviar()">Confirmar e Processar</button>
        <div id="msg" class="loading">Buscando dados na Nuvem (Isso pode levar alguns segundos)...</div>
      </form>
    </div>
    <script>
      function enviar() {
        var radios = document.getElementsByName('opcao');
        var selecionado = 'MUTIRAO';
        for (var i = 0; i < radios.length; i++) {
          if (radios[i].checked) { selecionado = radios[i].value; break; }
        }
        document.getElementById('msg').style.display = 'block';
        google.script.run
          .withSuccessHandler(function() { google.script.host.close(); })
          .withFailureHandler(function(e) { alert("Erro: " + e); google.script.host.close(); })
          .executarMutiraoComContexto(selecionado);
      }
    </script>
  `).setWidth(320).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Orquestrador de Estoque');
}

function executarMutiraoComContexto(tipoOrigem) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const abaMutirao = ss.getSheetByName(NOME_ABA_TRABALHO);        
    const idFonte = CONFIG.ids.fonteDadosGeral;
    
    if (!idFonte) throw new Error("ID_FONTE_GERAL não configurado nas propriedades.");
    const ssFonte = SpreadsheetApp.openById(idFonte);
    
    const nomeAbaDados = CONFIG.abas.dadosEstoque || "DadosEstoque";
    const abaDados = ssFonte.getSheetByName(nomeAbaDados);
    
    if (!abaMutirao) throw new Error(`A guia '${NOME_ABA_TRABALHO}' não foi encontrada.`);
    if (!abaDados) throw new Error(`A guia '${nomeAbaDados}' não encontrada na externa.`);
    
    const NOME_CONTEXTO = (tipoOrigem === "MUTIRAO") ? "MUTIRÃO" : "GRUPO AÇÃO";
    const COR_TITULO = (tipoOrigem === "MUTIRAO") ? "#1c4587" : "#cc0000";

    const lastRowMult = abaMutirao.getLastRow();
    if (lastRowMult < 2) { ui.alert(`A guia '${NOME_ABA_TRABALHO}' parece vazia.`); return; }

    const dadosAtuais = abaMutirao.getRange(2, 1, lastRowMult - 1, 1).getValues();
    const setCodigos = new Set();
    
    dadosAtuais.forEach(linha => {
      const cod = _norm(linha[0]);
      if (cod) setCodigos.add(cod);
    });

    if (setCodigos.size === 0) { ui.alert("Nenhum código válido na Coluna A."); return; }

    // Busca valor unitário da EntradaEmpenhos
    const mapaEntradas = _buscarDadosEntradaEmpenhos(setCodigos);

    const lastRowDados = abaDados.getLastRow();
    const mapaDados = new Map();
    
    const linhaInicialDados = 3; 
    if (lastRowDados >= linhaInicialDados) {
      // Leitura da base de dados com cobertura até a coluna AH (índice 33)
      const vDados = abaDados.getRange(linhaInicialDados, 1, lastRowDados - (linhaInicialDados-1), 43).getValues();
      
      for (let i = 0; i < vDados.length; i++) {
        const linha = vDados[i];
        const cod = _norm(linha[1]); // Col B
        
        if (cod && setCodigos.has(cod)) {
           if (!mapaDados.has(cod)) {
             mapaDados.set(cod, {
                estoqueTotal: 0,
                cmm: 0,
                descricaoSet: new Set(),
                aeSet: new Set(),
                notesSet: new Set(),
                procAndamentoSet: new Set(),
                qtdAndamentoSet: new Set(),
                modAndamentoSet: new Set(),
                procAtaSet: new Set(),
                modAtaSet: new Set(),
                saldoAtaSet: new Set(),
                empenhoSet: new Set(),
                qtdSet: new Set()
             });
           }
           const d = mapaDados.get(cod);
           
           d.estoqueTotal = Math.max(d.estoqueTotal, parseFloat(linha[7]) || 0); // Col H (Saldo Atual)
           d.cmm = Math.max(d.cmm, parseFloat(linha[8]) || 0); // Col I (CMM12)
           if (linha[2]) d.descricaoSet.add(String(linha[2]).trim()); // Col C (Descrição)
           
           // Extração Estrita por Índices Fixos
           const empenho = linha[16] !== undefined ? String(linha[16]).trim() : ""; // Col Q (Num.Empenho)
           const qtdReceber = linha[17] !== undefined ? String(linha[17]).trim() : ""; // Col R (Qtde a Receber)
           const procAta = linha[22] !== undefined ? String(linha[22]).trim() : ""; // Col W (Processo Ata/Direta)
           const modAta = linha[23] !== undefined ? String(linha[23]).trim() : ""; // Col X (Modalidade Homologado)
           const saldoAta = linha[26] !== undefined ? String(linha[26]).trim() : ""; // Col AA (Saldo da Ata)
           const notes = linha[29] !== undefined ? String(linha[29]).trim() : ""; // Col AD (Solicitação)
           const procAndamento = linha[30] !== undefined ? String(linha[30]).trim() : ""; // Col AE (Processo Em Andamento)
           const modAndamento = linha[31] !== undefined ? String(linha[31]).trim() : ""; // Col AF (Modalidade Em Andamento)
           const qtdAndamento = linha[32] !== undefined ? String(linha[32]).trim() : ""; // Col AG (Qtd Em Andamento)
           const ae = linha[33] !== undefined ? String(linha[33]).trim() : ""; // Col AH (AE)

           if (ae) d.aeSet.add(ae);
           
           // Regra: Ignora "notes" que começam com "5"
           if (notes && !notes.startsWith("5")) d.notesSet.add(notes);
           
           if (procAndamento) d.procAndamentoSet.add(procAndamento);
           if (qtdAndamento) d.qtdAndamentoSet.add(qtdAndamento);
           if (modAndamento) d.modAndamentoSet.add(modAndamento);
           if (procAta) d.procAtaSet.add(procAta);
           if (modAta) d.modAtaSet.add(modAta);
           
           // Regra: Ocultar Saldo da Ata se estiver zerado
           if (saldoAta) {
             const saldoAtaNum = parseFloat(saldoAta.replace(",", "."));
             if (isNaN(saldoAtaNum) || saldoAtaNum !== 0) {
               d.saldoAtaSet.add(saldoAta);
             }
           }
           
           if (empenho) d.empenhoSet.add(empenho);
           if (qtdReceber) d.qtdSet.add(qtdReceber);
        }
      }
    }

    const outputDados = [];
    const listaCompletaEmail = [];

    dadosAtuais.forEach((linhaInput) => {
      const codAtual = _norm(linhaInput[0]);
      if (!codAtual) {
         outputDados.push(Array(15).fill(""));
         return; 
      }
      
      const infoEstoque = mapaDados.get(codAtual) || {
        estoqueTotal: 0, cmm: 0, descricaoSet: new Set(), aeSet: new Set(), notesSet: new Set(), 
        procAndamentoSet: new Set(), qtdAndamentoSet: new Set(), modAndamentoSet: new Set(), procAtaSet: new Set(), 
        modAtaSet: new Set(), saldoAtaSet: new Set(), empenhoSet: new Set(), qtdSet: new Set()
      };
      
      const infoEntradas = mapaEntradas.get(codAtual) || { valor: 0, descricao: "" };
      
      const descricaoFinal = Array.from(infoEstoque.descricaoSet)[0] || infoEntradas.descricao || "Descrição não encontrada";
      const precoUnit = infoEntradas.valor;
      const joinSet = (setObj) => Array.from(setObj).join("\n");

      // Montagem Exata da Estrutura Nova (15 Colunas)
      const linhaFinal = [
        codAtual,                                // 1. (A) Código
        descricaoFinal,                          // 2. (B) Descrição
        precoUnit,                               // 3. (C) Valor Unitário
        infoEstoque.estoqueTotal,                // 4. (D) Estoque
        infoEstoque.cmm,                         // 5. (E) CMM
        joinSet(infoEstoque.aeSet),              // 6. (F) AE
        joinSet(infoEstoque.notesSet),           // 7. (G) Notes(Em Adamento)
        joinSet(infoEstoque.procAndamentoSet),   // 8. (H) Processos Em Andamento
        joinSet(infoEstoque.qtdAndamentoSet),    // 9. (I) Qtd Em Andamento
        joinSet(infoEstoque.modAndamentoSet),    // 10. (J) Modalidade Em Andamento
        joinSet(infoEstoque.procAtaSet),         // 11. (K) Processos Com Ata/CompraDireta
        joinSet(infoEstoque.modAtaSet),          // 12. (L) Modalidade Homologado
        joinSet(infoEstoque.saldoAtaSet),        // 13. (M) Saldo Ata
        joinSet(infoEstoque.empenhoSet),         // 14. (N) Empenho
        joinSet(infoEstoque.qtdSet)              // 15. (O) Quantidade
      ];

      outputDados.push(linhaFinal);

      listaCompletaEmail.push({
        codigo: codAtual,
        descricao: descricaoFinal,
        preco: precoUnit,
        estoque: infoEstoque.estoqueTotal,
        cmm: infoEstoque.cmm,
        ae: joinSet(infoEstoque.aeSet),
        notes: joinSet(infoEstoque.notesSet),
        procAndamento: joinSet(infoEstoque.procAndamentoSet),
        qtdAndamento: joinSet(infoEstoque.qtdAndamentoSet),
        modAndamento: joinSet(infoEstoque.modAndamentoSet),
        procAta: joinSet(infoEstoque.procAtaSet),
        modAta: joinSet(infoEstoque.modAtaSet),
        saldoAta: joinSet(infoEstoque.saldoAtaSet),
        empenho: joinSet(infoEstoque.empenhoSet),
        quantidade: joinSet(infoEstoque.qtdSet)
      });
    });

    if (outputDados.length > 0) {
      const numLinhas = outputDados.length;
      
      abaMutirao.getRange(2, 1, abaMutirao.getMaxRows() - 1, abaMutirao.getMaxColumns()).clearContent();
      abaMutirao.getRange(2, 1, numLinhas, 15).setValues(outputDados);
      
      // Formatações
      abaMutirao.getRange(2, 3, numLinhas, 1).setNumberFormat("R$ #,##0.00");
      abaMutirao.getRange(2, 4, numLinhas, 1).setNumberFormat("#,##0");
      abaMutirao.getRange(2, 5, numLinhas, 1).setNumberFormat("#,##0.##");
      abaMutirao.getRange(2, 13, numLinhas, 1).setNumberFormat("#,##0"); // Saldo Ata formatado como número se aplicável
      
      abaMutirao.getRange(2, 6, numLinhas, 10).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      abaMutirao.getRange(2, 6, numLinhas, 10).setVerticalAlignment("middle");
      
      const headers = [
        "Código", "Descrição", "Valor Unitário", "Estoque", "CMM", "AE", "Notes(Em Adamento)", 
        "Processos Em Andamento", "Qtd Em Andamento", "Modalidade Em Andamento", "Processos Com Ata/CompraDireta", 
        "Modalidade Homologado", "Saldo Ata", "Empenho", "Quantidade"
      ];
      
      abaMutirao.getRange(1, 1, 1, 15).setValues([headers])
        .setFontWeight("bold").setBackground(COR_TITULO).setFontColor("white").setHorizontalAlignment("center");
               
      abaMutirao.setColumnWidth(2, 300);

      if (listaCompletaEmail.length > 0) {
        const resp = ui.alert("Concluído", "Status Report Gerado!\nDeseja enviar o E-MAIL (PDF + Excel segregados)?", ui.ButtonSet.YES_NO);
        if (resp === ui.Button.YES) enviarEmailComAnexos(listaCompletaEmail, NOME_CONTEXTO, COR_TITULO);
      }
    } else {
      ui.alert("Aviso", "Nenhum dado encontrado para os códigos informados.", ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert("Erro ao processar: " + e.message);
  }
}

function _buscarDadosEntradaEmpenhos(setCodigos) {
  const mapa = new Map();
  try {
    if (!CONFIG.ids.fonteDadosGeral) throw new Error("ID Fonte Geral indefinido.");
    const ssFonte = SpreadsheetApp.openById(CONFIG.ids.fonteDadosGeral);
    
    const nomeAbaEntradas = CONFIG.abas.entradaEmpenhos || "EntradaEmpenhos";
    const aba = ssFonte.getSheetByName(nomeAbaEntradas);
    
    if (!aba) return mapa;      
    const lastRow = aba.getLastRow();
    if (lastRow < 2) return mapa;
    
    const dados = aba.getRange(2, 1, lastRow - 1, 17).getValues();
    
    dados.forEach(linha => {
      const cod = _norm(linha[2]);               
      const desc = String(linha[3]).trim();  
      let valor = linha[16];                                     
      
      if (cod && setCodigos.has(cod)) {
        if (typeof valor === 'string') valor = parseFloat(valor.replace("R$", "").replace(".", "").replace(",", ".").trim());
        else valor = parseFloat(valor) || 0;
        
        if (!mapa.has(cod)) {
          mapa.set(cod, { valor: 0, descricao: desc });
        }
        
        const obj = mapa.get(cod);
        if (valor > 0) obj.valor = valor; 
        if (desc && !obj.descricao) obj.descricao = desc;
      }
    });
  } catch (e) { console.error("Erro busca EntradaEmpenhos: " + e.message); }
  return mapa;
}

function enviarEmailComAnexos(listaItens, nomeContexto, corTitulo) {
  const isTeste = CONFIG_EMAILS_MUTIRAO.MODO_TESTE;
  const destinatarios = isTeste ? CONFIG_EMAILS_MUTIRAO.EMAIL_TESTE : CONFIG_EMAILS_MUTIRAO.LISTA_GERAL.join(",");
  const assunto = `[INFORMATIVO: ${nomeContexto} - Relatório de Status (${new Date().toLocaleDateString()})]`;

  let htmlEmail = `<div style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
      <h2 style="color: ${corTitulo}; border-bottom: 2px solid ${corTitulo}; padding-bottom: 5px;">
        ${nomeContexto} - Relatório de Status
      </h2>
      <p>Abaixo encontra-se a consolidação dos dados mapeados.</p>
      <p>Os relatórios detalhados contendo todas as 15 colunas encontram-se em anexo.</p><br>`;

  let htmlPdf = `<html><head><style>
       @page { size: landscape; margin: 10mm; } 
       body { font-family: Arial, sans-serif; font-size: 8px; } 
       h2 { background: ${corTitulo}; color: white; padding: 6px; text-align: center; margin-bottom: 10px; border-radius: 4px; } 
       table { width: 100%; border-collapse: collapse; margin-bottom: 15px; } 
       th, td { border: 1px solid #ccc; padding: 4px; vertical-align: top; } 
       th { background: #f2f2f2; }
  </style></head><body>
  <h2>Relatório Detalhado: ${nomeContexto} (${new Date().toLocaleDateString()})</h2>`;

  const pdfHeader = `<thead><tr>
    <th width="5%">Cód.</th>
    <th width="14%">Descrição</th>
    <th width="5%">Unitário</th>
    <th width="3%">Estoque</th>
    <th width="3%">CMM</th>
    <th width="5%">AE</th>
    <th width="7%">Notes</th>
    <th width="7%">Proc. Andam.</th>
    <th width="4%">Qtd. Andam.</th>
    <th width="7%">Mod. Andam.</th>
    <th width="7%">Proc. Ata</th>
    <th width="7%">Mod. Homol.</th>
    <th width="6%">Saldo Ata</th>
    <th width="14%">Empenho</th>
    <th width="6%">Qtd.</th>
  </tr></thead>`;

  const exHeaders = [
    "Código", "Descrição", "Valor Unitário", "Estoque", "CMM", "AE", "Notes(Em Adamento)", 
    "Processos Em Andamento", "Qtd Em Andamento", "Modalidade Em Andamento", "Processos Com Ata/CompraDireta", 
    "Modalidade Homologado", "Saldo Ata", "Empenho", "Quantidade"
  ];
  
  const dadosExcel = [exHeaders];

  htmlEmail += `<table style="width: 100%; border-collapse: collapse; font-size: 13px;">
                  <thead>
                    <tr style="background-color: #f2f2f2;">
                      <th style="border: 1px solid #ddd; padding: 6px; width: 15%;">Código</th>
                      <th style="border: 1px solid #ddd; padding: 6px; width: 70%;">Descrição</th>
                      <th style="border: 1px solid #ddd; padding: 6px; width: 15%;">Estoque</th>
                    </tr>
                  </thead><tbody>`;
                  
  htmlPdf += `<table>${pdfHeader}<tbody>`;

  listaItens.forEach(item => {
      const precoFmt = item.preco ? item.preco.toLocaleString('pt-BR', {style: 'currency', currency: 'BRL'}) : '-';
      
      htmlEmail += `<tr>
        <td style="border: 1px solid #ddd; padding: 6px; text-align: center;"><strong>${item.codigo}</strong></td>
        <td style="border: 1px solid #ddd; padding: 6px;">${item.descricao.substring(0, 90)}...</td>
        <td style="border: 1px solid #ddd; padding: 6px; text-align: center;">${item.estoque}</td>
      </tr>`;
      
      htmlPdf += `<tr>
        <td><strong>${item.codigo}</strong></td>
        <td>${item.descricao.substring(0, 85)}</td>
        <td align="right">${precoFmt}</td>
        <td align="center">${item.estoque}</td>
        <td align="center">${item.cmm || '-'}</td>
        <td>${item.ae.replace(/\n/g, "<br>")}</td>
        <td>${item.notes.replace(/\n/g, "<br>")}</td>
        <td>${item.procAndamento.replace(/\n/g, "<br>")}</td>
        <td align="center">${item.qtdAndamento}</td>
        <td>${item.modAndamento.replace(/\n/g, "<br>")}</td>
        <td>${item.procAta.replace(/\n/g, "<br>")}</td>
        <td>${item.modAta.replace(/\n/g, "<br>")}</td>
        <td align="center">${item.saldoAta.replace(/\n/g, "<br>")}</td>
        <td>${item.empenho.replace(/\n/g, "<br>")}</td>
        <td align="center">${item.quantidade}</td>
      </tr>`;
      
      dadosExcel.push([
        item.codigo, item.descricao, item.preco, item.estoque, item.cmm, 
        item.ae, item.notes, item.procAndamento, item.qtdAndamento, item.modAndamento, 
        item.procAta, item.modAta, item.saldoAta, item.empenho, item.quantidade
      ]);
  });

  htmlEmail += `</tbody></table>`;
  htmlPdf += `</tbody></table>`;
  
  htmlEmail += `<br><p>Atenciosamente,<br><strong>Orquestrador de Estoque INCA</strong></p></div>`;
  htmlPdf += `</body></html>`;

  const pdfBlob = HtmlService.createHtmlOutput(htmlPdf).getAs(MimeType.PDF).setName(`Relatorio_Status_${nomeContexto}.pdf`);
  let excelBlob = null;

  try {
    const tempSS = SpreadsheetApp.create("Temp_Export");
    const sheet = tempSS.getSheets()[0];
    const rangeTemp = sheet.getRange(1, 1, dadosExcel.length, exHeaders.length);
    rangeTemp.setValues(dadosExcel);
    
    sheet.getRange(1, 1, 1, exHeaders.length).setBackground("#d9ead3").setFontWeight("bold");
    SpreadsheetApp.flush();
    
    const url = "https://docs.google.com/spreadsheets/d/" + tempSS.getId() + "/export?format=xlsx";
    excelBlob = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob().setName(`Relatorio_Status_${nomeContexto}.xlsx`);
    DriveApp.getFileById(tempSS.getId()).setTrashed(true);
  } catch (ex) { console.error("Erro no Excel: " + ex.message); }

  const anexos = [pdfBlob];
  if (excelBlob) anexos.push(excelBlob);

  MailApp.sendEmail({
    to: Session.getActiveUser().getEmail(), 
    cc: destinatarios, 
    subject: assunto, 
    htmlBody: htmlEmail, 
    attachments: anexos
  });
  
  SpreadsheetApp.getUi().alert(`E-mail com PDF e Excel enviado com sucesso em Cópia Oculta (BCC)!`);
}

function _norm(t) { return t ? String(t).trim().toUpperCase() : ""; }