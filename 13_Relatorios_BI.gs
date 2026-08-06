// =================================================================
// --- BLOCO 13: RELATÓRIOS DE INTELIGÊNCIA (BI) ---
// =================================================================

function gerarRelatorioPerformanceFornecedores() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    const abaComp = ss.getSheetByName(CONFIG.destino.nomeAba); 
    if (!abaComp) throw new Error("Aba Compilados não encontrada.");

    const dados = abaComp.getRange(2, 1, abaComp.getLastRow() - 1, 19).getValues();
    const stats = new Map();

    dados.forEach(r => {
      const fornecedor = r[4] ? String(r[4]).trim().toUpperCase() : "NÃO INFORMADO";
      const status = r[18] ? String(r[18]).trim().toUpperCase() : "";

      if (status === "ELIMINADA" || status === "SOLICITAR ASSOCIAÇÃO") return;

      if (!stats.has(fornecedor)) {
        stats.set(fornecedor, { total: 0, pendentes: 0, concluidos: 0 });
      }

      const s = stats.get(fornecedor);
      s.total++;

      if (status.includes("PENDENTE")) s.pendentes++;
      else if (status === "CONCLUÍDO") s.concluidos++;
    });

    const relatorio = [];
    stats.forEach((val, key) => {
      if (val.total > 1) {
        const taxaProblema = val.total > 0 ? (val.pendentes / val.total) : 0;
        relatorio.push([
          key, 
          val.total, 
          val.concluidos, 
          val.pendentes, 
          taxaProblema
        ]);
      }
    });

    relatorio.sort((a, b) => b[3] - a[3] || b[4] - a[4]);

    let abaRel = ss.getSheetByName("BI_Fornecedores");
    if (!abaRel) abaRel = ss.insertSheet("BI_Fornecedores");
    abaRel.clear();

    const cabecalho = ["Fornecedor", "Total de Itens", "Entregues", "Pendentes (Atraso)", "% de Pendência"];
    abaRel.getRange(1, 1, 1, 5).setValues([cabecalho])
      .setFontWeight("bold")
      .setBackground("#4c1130") 
      .setFontColor("white");

    if (relatorio.length > 0) {
      abaRel.getRange(2, 1, relatorio.length, 5).setValues(relatorio);
      abaRel.getRange(2, 5, relatorio.length, 1).setNumberFormat("0.0%");
      abaRel.getRange(2, 1, relatorio.length, 5).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
    }

    abaRel.autoResizeColumns(1, 5);
    abaRel.activate();
    ui.alert("Relatório de Performance Gerado com Sucesso!");

  } catch (e) {
    ui.alert("Erro BI Fornecedores: " + e.message);
  }
}