function diagnosticarListaItens() {
  // Lista de itens solicitada
  const listaItens = ["1757"];
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  let relatorio = "=== RELATÓRIO DE DIAGNÓSTICO (LISTA) ===\n\n";

  // --- 1. PREPARAÇÃO DOS DADOS DA GUIA 'Temp' (LOCAL) ---
  const sheetTemp = ss.getSheetByName("Temp");
  let dadosTemp = [];
  let cabTemp = [];
  let idxItemTemp = -1;

  if (sheetTemp) {
    dadosTemp = sheetTemp.getDataRange().getValues();
    cabTemp = dadosTemp[0];
    // Tenta achar pelo nome, senão assume Coluna A (índice 0)
    idxItemTemp = buscarCol(cabTemp, ["item", "código", "codigo"]);
    if (idxItemTemp === -1) idxItemTemp = 0; // Fallback Coluna A
    
    relatorio += `[GUIA TEMP]: Lendo coluna de índice ${idxItemTemp} (${String.fromCharCode(65 + idxItemTemp)})\n`;
  } else {
    relatorio += "[GUIA TEMP]: Guia não encontrada!\n";
  }

  // --- 2. PREPARAÇÃO DOS DADOS DA 'EntradaEmpenhos' (EXTERNA) ---
  const ID_EXTERNA = "1s44YD2ozLAbBdGQbBE5iW7HcUzvQULZqd4ynYlV_HXA";
  let dadosExt = [];
  let idxItemExt = -1;
  let acessouExt = false;

  try {
    const ssExt = SpreadsheetApp.openById(ID_EXTERNA);
    const sheetExt = ssExt.getSheetByName("EntradaEmpenhos");
    if (sheetExt) {
      dadosExt = sheetExt.getDataRange().getValues();
      const cabExt = dadosExt[0];
      // Tenta achar pelo nome, senão assume Coluna C (índice 2) conforme mencionado
      idxItemExt = buscarCol(cabExt, ["Item", "Material", "Código"]);
      if (idxItemExt === -1) idxItemExt = 2; // Fallback Coluna C
      
      relatorio += `[ENTRADA EMPENHOS]: Lendo coluna de índice ${idxItemExt} (${String.fromCharCode(65 + idxItemExt)})\n`;
      acessouExt = true;
    } else {
      relatorio += "[ENTRADA EMPENHOS]: Aba não encontrada na planilha externa.\n";
    }
  } catch (e) {
    relatorio += "[ENTRADA EMPENHOS]: Erro de conexão (" + e.message + ")\n";
  }

  relatorio += "\n--------------------------------------------------\n";

  // --- 3. VERIFICAÇÃO ITEM A ITEM ---
  listaItens.forEach(itemAlvo => {
    itemAlvo = String(itemAlvo).trim(); // Garante que é texto sem espaços extras
    relatorio += `\n>>> ITEM: ${itemAlvo}\n`;

    // A. Busca na Temp
    let achouTemp = false;
    if (sheetTemp && idxItemTemp > -1) {
      for (let i = 1; i < dadosTemp.length; i++) {
        if (String(dadosTemp[i][idxItemTemp]).trim() === itemAlvo) {
          achouTemp = true;
          // Mostra detalhes (Qtd e Valor Total se achar as colunas, ou a linha inteira resumida)
          relatorio += `    [TEMP] Encontrado na Linha ${i + 1}\n`;
        }
      }
    }
    if (!achouTemp) relatorio += "    [TEMP] NÃO encontrado.\n";

    // B. Busca na Externa
    let achouExt = false;
    if (acessouExt && idxItemExt > -1) {
      for (let i = 1; i < dadosExt.length; i++) {
        if (String(dadosExt[i][idxItemExt]).trim() === itemAlvo) {
          achouExt = true;
          // Pega Coluna L (Qtd - idx 11) e O (Data - idx 14) para confirmar
          const qtd = dadosExt[i][11]; 
          const data = dadosExt[i][14];
          relatorio += `    [EXTERNA] Encontrado na Linha ${i + 1} | Data: ${formatarData(data)} | Qtd: ${qtd}\n`;
        }
      }
    }
    if (!achouExt) relatorio += "    [EXTERNA] NÃO encontrado.\n";
  });

  // --- 4. EXIBIR ---
  Logger.log(relatorio);
  // O alerta tem limite de caracteres, se for muito grande pode cortar, mas para 8 itens deve dar.
  ui.alert(relatorio);
}

// Auxiliares
function buscarCol(cabecalho, nomes) {
  for (let i = 0; i < cabecalho.length; i++) {
    const h = String(cabecalho[i]).toUpperCase().trim();
    for (let n of nomes) {
      if (h === n.toUpperCase().trim() || h.includes(n.toUpperCase().trim())) return i;
    }
  }
  return -1;
}

function formatarData(d) {
  if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
  return String(d);
}

//ATUALIZA GUIA MONITOR

function atualizarMonitor() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var monitorSheet = ss.getSheetByName("Monitor");
  
  if (!monitorSheet) {
    ss.insertSheet("Monitor");
    monitorSheet = ss.getSheetByName("Monitor");
  }

  var totalCelulasCriadas = 0;
  var limiteGoogle = 10000000; // Limite de 10 milhões

  // Percorre todas as abas para somar a grade total
  sheets.forEach(function(sheet) {
    var colunas = sheet.getMaxColumns();
    var linhas = sheet.getMaxRows();
    totalCelulasCriadas += (colunas * linhas);
  });

  var disponivel = limiteGoogle - totalCelulasCriadas;
  var percentualOcupado = (totalCelulasCriadas / limiteGoogle);

  // Escrevendo os dados na aba Monitor
  monitorSheet.clear();
  monitorSheet.getRange("A1:B1").setValues([["Métrica", "Valor"]]).setFontWeight("bold");
  monitorSheet.getRange("A2:B5").setValues([
    ["Células em Uso (Grade Total)", totalCelulasCriadas],
    ["Limite Total do Google", limiteGoogle],
    ["Células Restantes", disponivel],
    ["Percentual de Uso", percentualOcupado]
  ]);

  // Formatação rápida
  monitorSheet.getRange("B5").setNumberFormat("0.00%");
  monitorSheet.getRange("B2:B4").setNumberFormat("#,##0");
  monitorSheet.autoResizeColumns(1, 2);
  
  SpreadsheetApp.getUi().alert("Monitor atualizado com sucesso!");
}