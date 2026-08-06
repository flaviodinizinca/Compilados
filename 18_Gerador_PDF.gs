// =================================================================
// --- BLOCO 18: GERADOR DE PDF POR STATUS (SEM LOGO / MULTI-SELEÇÃO) ---
// =================================================================

/**
 * Abre a janela (Dialog) com Checkboxes para seleção múltipla.
 */
function abrirMenuGerarPDF() {
  const ui = SpreadsheetApp.getUi();
  const htmlTemplate = HtmlService.createTemplate(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: 'Segoe UI', sans-serif; padding: 15px; background-color: #f9f9f9; text-align: center; }
          h3 { color: #274e13; margin-top: 0; }
          
          /* Estilo dos Checkboxes em Grid */
          .checkbox-container {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 10px;
            text-align: left;
            max-height: 250px;
            overflow-y: auto;
            background: white;
            padding: 10px;
            border: 1px solid #ccc;
            border-radius: 5px;
            margin-bottom: 15px;
          }
          .checkbox-item { font-size: 14px; color: #333; cursor: pointer; }
          .checkbox-item input { margin-right: 8px; transform: scale(1.2); }

          /* Botões de Ação */
          .btn-group { display: flex; gap: 10px; justify-content: center; margin-bottom: 15px; }
          .btn-small { background: #eee; border: none; padding: 5px 10px; cursor: pointer; border-radius: 3px; font-size: 12px; }
          .btn-small:hover { background: #ddd; }

          button.main { background-color: #274e13; color: white; padding: 12px 0; border: none; border-radius: 5px; cursor: pointer; font-size: 16px; width: 100%; font-weight: bold; transition: background 0.3s; }
          button.main:hover { background-color: #183209; }
          button:disabled { background-color: #ccc; cursor: not-allowed; }

          .loader { display: none; color: #555; margin-top: 15px; font-size: 14px; }
          a { color: #274e13; font-weight: bold; text-decoration: none; }
        </style>
      </head>
      <body>
        <h3>📄 Gerar Relatório (PDF)</h3>
        
        <div class="btn-group">
          <button type="button" class="btn-small" onclick="toggle(true)">Marcar Todos</button>
          <button type="button" class="btn-small" onclick="toggle(false)">Desmarcar</button>
        </div>

        <div class="checkbox-container">
          <? for (var i = 0; i < listaStatus.length; i++) { ?>
            <label class="checkbox-item">
              <input type="checkbox" name="status" value="<?= listaStatus[i] ?>">
              <?= listaStatus[i] ?>
            </label>
          <? } ?>
        </div>
        
        <button onclick="gerar()" id="btnGerar" class="main">Gerar PDF</button>
        
        <div id="loading" class="loader">
          🔄 Processando...<br>Compilando dados e gerando arquivo...
        </div>

        <script>
          function toggle(source) {
            var checkboxes = document.getElementsByName('status');
            for(var i=0, n=checkboxes.length;i<n;i++) {
              checkboxes[i].checked = source;
            }
          }

          function gerar() {
            var checkboxes = document.getElementsByName('status');
            var selecionados = [];
            for (var i=0; i<checkboxes.length; i++) {
              if (checkboxes[i].checked) {
                selecionados.push(checkboxes[i].value);
              }
            }

            if (selecionados.length === 0) {
              alert("Por favor, selecione pelo menos um status.");
              return;
            }

            document.getElementById('loading').style.display = 'block';
            document.getElementById('btnGerar').disabled = true;
            
            google.script.run
              .withSuccessHandler(function(url) {
                 document.getElementById('loading').innerHTML = '✅ <b>Sucesso!</b><br><br><a href="' + url + '" target="_blank" style="font-size:16px; border: 1px solid #274e13; padding: 10px; display:block; border-radius:4px; background: #edf7ed;">📥 BAIXAR PDF</a>';
              })
              .withFailureHandler(function(err) {
                 document.getElementById('loading').innerHTML = '<span style="color:red">❌ Erro: ' + err + '</span>';
                 document.getElementById('btnGerar').disabled = false;
              })
              .processarGeracaoPDF(selecionados);
          }
        </script>
      </body>
    </html>
  `);

  htmlTemplate.listaStatus = _obterStatusUnicos();
  ui.showModalDialog(htmlTemplate.evaluate().setWidth(450).setHeight(500), 'Gerador de PDF INCA');
}

/**
 * Processa a geração aceitando MÚLTIPLOS status (Array).
 */
function processarGeracaoPDF(listaStatusEscolhidos) {
  const ssAtual = SpreadsheetApp.getActiveSpreadsheet();
  let tempSS = null;
  
  try {
    const abaComp = ssAtual.getSheetByName(CONFIG.destino.nomeAba); 
    if (!abaComp) throw new Error("Aba Compilados não encontrada.");

    // 1. Coleta e Filtra os Dados
    const dados = abaComp.getRange(2, 1, abaComp.getLastRow() - 1, 21).getValues();
    
    // Verifica se o status da linha está INCLUÍDO na lista de escolhidos
    const dadosFiltrados = dados
      .filter(r => {
        const st = r[18] ? String(r[18]).trim() : "";
        return listaStatusEscolhidos.includes(st);
      })
      .map(r => [
        r[0], // Nº Empenho
        r[4], // Fornecedor
        r[5], // Item
        r[15], // Qtd Recebida
        r[16], // Qtd Residual
        r[19], // Processo
        r[20]  // Modalidade
      ]);

    if (dadosFiltrados.length === 0) throw new Error("Nenhum item encontrado para os status selecionados.");

    // 2. Define Nome do Arquivo Inteligente
    let sufixoNome = "";
    if (listaStatusEscolhidos.length === 1) {
      sufixoNome = listaStatusEscolhidos[0].replace(/[^a-zA-Z0-9]/g, '');
    } else {
      sufixoNome = "Multiplos_Status";
    }
    const nomeArquivo = `Relatorio_INCA_${sufixoNome}_${new Date().getTime()}`;

    // 3. Cria Planilha Temporária
    tempSS = SpreadsheetApp.create(nomeArquivo);
    const sheet = tempSS.getSheets()[0];

    // 4. Cabeçalho e Título (CENTRALIZADO, SEM LOGO)
    // Mescla de A1 até G2 para o título principal
    sheet.getRange("A1:G2").merge()
         .setValue("RELATÓRIO DE EMPENHOS")
         .setFontWeight("bold")
         .setFontSize(18)
         .setHorizontalAlignment("center")
         .setVerticalAlignment("middle")
         .setBackground("white"); // Fundo limpo
         
    // Subtítulo com filtros usados
    const textoFiltro = listaStatusEscolhidos.length > 3 ? 
        "Vários Status Selecionados" : 
        listaStatusEscolhidos.join(" | ");

    sheet.getRange("A3:G3").merge()
         .setValue(`Filtro: ${textoFiltro} | Gerado em: ${new Date().toLocaleString('pt-BR')}`)
         .setFontSize(10)
         .setHorizontalAlignment("center")
         .setFontStyle("italic");

    const headers = ["Nº Empenho", "FORNECEDOR", "ITEM", "QTD REC.", "QTD RESIDUAL", "PROCESSO", "MODALIDADE"];
    
    // Cabeçalho da tabela começa na linha 5
    const headerRow = 5;
    const headerRange = sheet.getRange(headerRow, 1, 1, headers.length);
    headerRange.setValues([headers])
               .setFontWeight("bold")
               .setBackground("#d9ead3") // Verde claro padrão INCA/Excel
               .setBorder(true, true, true, true, true, true)
               .setHorizontalAlignment("center");

    // 5. Escreve os Dados
    const dataRange = sheet.getRange(headerRow + 1, 1, dadosFiltrados.length, headers.length);
    dataRange.setValues(dadosFiltrados);
    
    // Formatação da Tabela
    dataRange.setBorder(true, true, true, true, true, true);
    dataRange.setVerticalAlignment("middle");
    
    // Formatação de Números (Colunas D e E -> Índices 4 e 5)
    sheet.getRange(headerRow + 1, 4, dadosFiltrados.length, 2).setNumberFormat("#,##0");
    
    // Alinhamentos
    sheet.getRange(headerRow + 1, 1, dadosFiltrados.length, 1).setHorizontalAlignment("center"); // Empenho
    sheet.getRange(headerRow + 1, 3, dadosFiltrados.length, 1).setHorizontalAlignment("center"); // Item
    sheet.getRange(headerRow + 1, 4, dadosFiltrados.length, 2).setHorizontalAlignment("center"); // Qtds

    // Ajuste de Largura das Colunas
    sheet.autoResizeColumns(1, 7);
    
    // Ajustes finos para colunas de texto longo
    if (sheet.getColumnWidth(2) > 250) sheet.setColumnWidth(2, 250); // Fornecedor
    if (sheet.getColumnWidth(7) > 200) sheet.setColumnWidth(7, 200); // Modalidade
    
    dataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

    SpreadsheetApp.flush();

    // 6. Exportação PDF (Configurações)
    const urlExport = 'https://docs.google.com/spreadsheets/d/' + tempSS.getId() + '/export?' +
      'exportFormat=pdf&format=pdf' +
      '&size=A4' +
      '&portrait=false' +   // Paisagem
      '&fitw=true' +        // Ajustar à largura
      '&sheetnames=false&printtitle=false&pagenumbers=true' +
      '&gridlines=true' +
      '&fzr=false' +
      '&gid=' + sheet.getSheetId();

    const params = {
      method: "GET",
      headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(urlExport, params);
    
    if (response.getResponseCode() !== 200) {
      throw new Error("Falha na conversão PDF: " + response.getContentText());
    }

    const blobPDF = response.getBlob();
    blobPDF.setName(nomeArquivo + ".pdf");

    const arquivoFinal = DriveApp.createFile(blobPDF);
    
    // Lixeira o arquivo temporário
    DriveApp.getFileById(tempSS.getId()).setTrashed(true);

    return arquivoFinal.getUrl();

  } catch (e) {
    // Tenta limpar o arquivo temporário em caso de erro
    if (tempSS) {
      try { DriveApp.getFileById(tempSS.getId()).setTrashed(true); } catch(x){}
    }
    throw e.message;
  }
}

function _obterStatusUnicos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(CONFIG.destino.nomeAba);
  if (!aba) return [];

  const lastRow = aba.getLastRow();
  if (lastRow < 2) return [];

  const valores = aba.getRange(2, 19, lastRow - 1, 1).getValues().flat();
  
  const unicos = [...new Set(valores)]
    .filter(v => v !== "")
    .map(v => String(v).trim())
    .sort();
    
  return unicos;
}