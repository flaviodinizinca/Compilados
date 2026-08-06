// =================================================================
// --- BLOCO 12: DASHBOARD DE STATUS ---
// =================================================================

function gerarDashboardStatus() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    const nomeAbaOrigem = CONFIG.destino.nomeAba; 
    const nomeAbaDash = "Dashboard";
    
    const abaDados = ss.getSheetByName(nomeAbaOrigem);
    if (!abaDados) throw new Error(`Aba '${nomeAbaOrigem}' não encontrada.`);
    
    const lastRow = abaDados.getLastRow();
    if (lastRow < 2) throw new Error("Não há dados suficientes para gerar o dashboard.");
    
    const valoresStatus = abaDados.getRange(2, 19, lastRow - 1, 1).getValues().flat();
    
    const contagem = {};
    let totalItens = 0;
    
    valoresStatus.forEach(status => {
      const chave = status ? String(status).trim().toUpperCase() : "(VAZIO)";
      contagem[chave] = (contagem[chave] || 0) + 1;
      totalItens++;
    });
    
    let dadosRelatorio = Object.keys(contagem).map(chave => {
      const qtd = contagem[chave];
      const pct = qtd / totalItens;
      return [chave, qtd, pct];
    });
    
    dadosRelatorio.sort((a, b) => b[1] - a[1]);
    
    dadosRelatorio.unshift(["STATUS", "QTD", "%"]);
    
    let abaDash = ss.getSheetByName(nomeAbaDash);
    if (!abaDash) {
      abaDash = ss.insertSheet(nomeAbaDash);
    } else {
      abaDash.clear(); 
    }
    
    const numLinhas = dadosRelatorio.length;
    abaDash.getRange(2, 2, numLinhas, 3).setValues(dadosRelatorio); 
    
    const rangeHeader = abaDash.getRange(2, 2, 1, 3);
    rangeHeader.setFontWeight("bold").setBackground("#4a86e8").setFontColor("white").setHorizontalAlignment("center");
    
    abaDash.getRange(3, 4, numLinhas - 1, 1).setNumberFormat("0.00%"); 
    
    abaDash.getRange(2, 2, numLinhas, 3).setBorder(true, true, true, true, true, true);
    
    abaDash.setColumnWidth(1, 20); 
    abaDash.setColumnWidth(2, 250); 
    
    const charts = abaDash.getCharts();
    charts.forEach(c => abaDash.removeChart(c));
    
    const chart = abaDash.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(abaDash.getRange(2, 2, numLinhas, 2)) 
      .setPosition(2, 6, 0, 0) 
      .setOption('title', 'Distribuição dos Status dos Empenhos')
      .setOption('pieSliceText', 'percentage')
      .setOption('width', 600)
      .setOption('height', 400)
      .setOption('is3D', true) 
      .build();
      
    abaDash.insertChart(chart);
    
    abaDash.activate();
    ss.toast("Dashboard atualizado com sucesso!", "Concluído");
    
  } catch (e) {
    ui.alert("Erro ao gerar Dashboard: " + e.message);
  }
}