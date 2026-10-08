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

// =================================================================
// --- RELATÓRIO: VALOR RESÍDUO 10% ---
// CORRIGIDO: Colunas I (QTD EMPENHO), P (QTD RECEBIDA), Q (QTD RESIDUAL)
// Filtra itens onde RESIDUAL ≤ 10% do Empenho
// =================================================================

function gerarRelatorioResiduo10() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    const abaComp = ss.getSheetByName("Compilados");
    if (!abaComp) throw new Error("Aba Compilados não encontrada.");
    
    // Buscar dados de EntradaEmpenhos para valor unitário
    const idFonteGeral = CONFIG.ids.fonteDadosGeral;
    if (!idFonteGeral) throw new Error("ID_FONTE_GERAL não configurado.");
    
    const ssFonte = SpreadsheetApp.openById(idFonteGeral);
    const nomeAbaEntradas = CONFIG.abas.entradaEmpenhos || "EntradaEmpenhos";
    const abaEntradas = ssFonte.getSheetByName(nomeAbaEntradas);
    
    if (!abaEntradas) throw new Error(`Aba '${nomeAbaEntradas}' não encontrada na fonte geral.`);
    
    // Mapa de valores unitários: chave = "EMPENHO||ITEM", valor = VALOR_UNITÁRIO
    const mapaValoresUnitarios = new Map();
    const lastRowEntradas = abaEntradas.getLastRow();
    
    if (lastRowEntradas >= 2) {
      // A (0) = Empenho, C (2) = Item, Q (16) = VALOR_UNITÁRIO
      const dadosEntradas = abaEntradas.getRange(2, 1, lastRowEntradas - 1, 17).getValues();
      
      dadosEntradas.forEach(linha => {
        const empenho = _normalizarEmpenhoBusca(String(linha[0] || '').trim());
        const item = String(linha[2] || '').trim();
        const valorUnitario = parseFloat(linha[16]) || 0;
        
        if (empenho && item) {
          const chave = `${empenho}||${item}`;
          if (!mapaValoresUnitarios.has(chave)) {
            mapaValoresUnitarios.set(chave, valorUnitario);
          }
        }
      });
    }
    
    let abaRel = ss.getSheetByName("Resíduo 10%");
    if (!abaRel) abaRel = ss.insertSheet("Resíduo 10%");
    abaRel.clear();
    
    // Ler aba Compilados com todas as colunas necessárias
    const lastRowComp = abaComp.getLastRow();
    const dados = abaComp.getRange(2, 1, lastRowComp - 1, 21).getValues();
    const dadosFilt = [];
    
    // Filtrar itens onde QTD RESIDUAL ≤ 10% do QTD EMPENHO
    dados.forEach(l => {
      const empenho = String(l[0] || '').trim();
      const item = String(l[5] || '').trim();
      
      // Colunas corretas:
      // I (índice 8) = QTD EMPENHO
      // P (índice 15) = QTD RECEBIDA
      // Q (índice 16) = QTD RESIDUAL
      
      const qtdEmpenho = parseFloat(l[8]) || 0;
      const qtdRecebida = parseFloat(l[15]) || 0;
      const qtdResidual = parseFloat(l[16]) || 0;
      
      // Validar se há empenho
      if (!empenho || qtdEmpenho <= 0) return;
      
      // Calcular percentual do residual
      const percentualResidual = (qtdResidual / qtdEmpenho) * 100;
      
      // Filtrar: apenas itens onde residual ≤ 10% do empenho
      if (percentualResidual > 10) return;
      
      // Buscar valor unitário em EntradaEmpenhos
      const empenhoNorm = _normalizarEmpenhoBusca(empenho);
      const chave = `${empenhoNorm}||${item}`;
      const valorUnitario = mapaValoresUnitarios.get(chave) || 0;
      
      // Calcular valor total = QTD RESIDUAL × VALOR UNITÁRIO
      const valorTotal = qtdResidual * valorUnitario;
      
      dadosFilt.push([
        empenho,                           // A: EMPENHO
        l[4],                              // B: FORNECEDOR
        item,                              // C: ITEM
        l[6],                              // D: DESCRIÇÃO
        qtdEmpenho,                        // E: QTD EMPENHO
        qtdRecebida,                       // F: QTD RECEBIDA
        qtdResidual,                       // G: QTD RESIDUAL (10%)
        percentualResidual.toFixed(2),     // H: % RESIDUAL
        l[19],                             // I: PROCESSO
        l[20],                             // J: MODALIDADE
        valorUnitario,                     // K: VALOR UNITÁRIO
        valorTotal                         // L: VALOR TOTAL (RESIDUAL × VLR UNITÁRIO)
      ]);
    });
    
    const headerNovo = [
      "EMPENHO", 
      "FORNECEDOR", 
      "ITEM", 
      "DESCRIÇÃO", 
      "QTD EMPENHO", 
      "QTD RECEBIDA", 
      "QTD RESIDUAL", 
      "% RESIDUAL",
      "PROCESSO", 
      "MODALIDADE",
      "VALOR UNITÁRIO",
      "VALOR TOTAL"
    ];
    
    abaRel.getRange(1, 1, 1, headerNovo.length)
      .setValues([headerNovo])
      .setFontWeight('bold')
      .setBackground("#cfe2f3")
      .setHorizontalAlignment('center');
    
    if (dadosFilt.length > 0) {
      abaRel.getRange(2, 1, dadosFilt.length, headerNovo.length).setValues(dadosFilt);
      
      // Formatação de valores monetários (colunas K e L)
      abaRel.getRange(2, 11, dadosFilt.length, 2).setNumberFormat('R$ #,##0.00');
      
      // Formatação de quantidades (colunas E, F, G)
      abaRel.getRange(2, 5, dadosFilt.length, 3).setNumberFormat('#,##0');
      
      // Formatação de percentual (coluna H)
      abaRel.getRange(2, 8, dadosFilt.length, 1).setNumberFormat('0.00%');
      
      abaRel.setFrozenRows(1);
      abaRel.autoResizeColumns(1, headerNovo.length);
      
      ui.alert(
        "Relatório Gerado", 
        `${dadosFilt.length} itens com saldo ≤ 10% encontrados.\nValor total em resíduos: R$ ${dadosFilt.reduce((sum, row) => sum + parseFloat(row[11] || 0), 0).toLocaleString('pt-BR')}`, 
        ui.ButtonSet.OK
      );
    } else {
      ui.alert("Info", "Nenhum item com saldo residual ≤ 10% encontrado.", ui.ButtonSet.OK);
    }
    
  } catch (e) { 
    ui.alert("Erro", e.message, ui.ButtonSet.OK); 
  }
}

// =================================================================
// --- RELATÓRIO: VALIDADE DE ATAS ---
// =================================================================

function gerarRelatorioValidadeAtas() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    const idFonteGeral = CONFIG.ids.fonteDadosGeral;
    if (!idFonteGeral) throw new Error("ID_FONTE_GERAL não configurado.");
    
    const ssFonte = SpreadsheetApp.openById(idFonteGeral);
    const nomeAbaDados = CONFIG.abas.dadosEstoque || "DadosEstoque";
    const abaDados = ssFonte.getSheetByName(nomeAbaDados);
    
    if (!abaDados) throw new Error(`Aba '${nomeAbaDados}' não encontrada na fonte geral.`);
    
    let abaRel = ss.getSheetByName("Validade Atas");
    if (!abaRel) abaRel = ss.insertSheet("Validade Atas");
    abaRel.clear();
    
    const linhaInicialDados = 3;
    const lastRowDados = abaDados.getLastRow();
    
    if (lastRowDados < linhaInicialDados) {
      ui.alert("Info", "Nenhum dado encontrado na fonte de dados.", ui.ButtonSet.OK);
      return;
    }
    
    const dados = abaDados.getRange(linhaInicialDados, 1, lastRowDados - (linhaInicialDados - 1), 43).getValues();
    const dadosFilt = [];
    const hoje = new Date();
    
    dados.forEach(l => {
      // Coluna W (22) = Processo Ata/Direta
      // Coluna AA (26) = Data Vigência Ata (ou similar)
      // Coluna X (23) = Modalidade Homologado
      
      const procAta = l[22] ? String(l[22]).trim() : "";
      let dataVigencia = l[25]; // Ajustar conforme sua estrutura
      
      if (!procAta) return; // Pula se não houver processo de ata
      
      // Tenta converter data
      let dataObj = null;
      if (dataVigencia) {
        if (Object.prototype.toString.call(dataVigencia) === '[object Date]') {
          dataObj = dataVigencia;
        } else if (typeof dataVigencia === 'string') {
          dataObj = new Date(dataVigencia);
        }
      }
      
      // Calcula dias até vencimento
      let diasRestantes = "";
      let status = "Ativa";
      let cor = "#d4edda";
      
      if (dataObj && !isNaN(dataObj)) {
        const diffTime = dataObj - hoje;
        const diasAteVencimento = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        diasRestantes = diasAteVencimento;
        
        if (diasAteVencimento < 0) {
          status = "🔴 VENCIDA";
          cor = "#f8d7da";
        } else if (diasAteVencimento < 30) {
          status = "⚠️ VENCE EM BREVE";
          cor = "#fff3cd";
        } else if (diasAteVencimento < 90) {
          status = "🟡 ATENÇÃO";
          cor = "#fff3cd";
        }
      }
      
      dadosFilt.push([
        l[1] ? String(l[1]).trim() : "",  // Item
        l[2] ? String(l[2]).trim() : "",  // Descrição
        procAta,                           // Processo Ata
        l[23] ? String(l[23]).trim() : "", // Modalidade
        dataVigencia ? dataVigencia : "",  // Data de Vigência
        diasRestantes,                     // Dias Restantes
        status                             // Status
      ]);
    });
    
    const headerNovo = [
      "ITEM",
      "DESCRIÇÃO",
      "PROCESSO ATA/DIRETA",
      "MODALIDADE",
      "DATA DE VIGÊNCIA",
      "DIAS RESTANTES",
      "STATUS"
    ];
    
    abaRel.getRange(1, 1, 1, headerNovo.length)
      .setValues([headerNovo])
      .setFontWeight('bold')
      .setBackground("#1f4e78")
      .setFontColor("white")
      .setHorizontalAlignment('center');
    
    if (dadosFilt.length > 0) {
      abaRel.getRange(2, 1, dadosFilt.length, headerNovo.length).setValues(dadosFilt);
      
      // Formatação de data
      abaRel.getRange(2, 5, dadosFilt.length, 1).setNumberFormat('dd/mm/yyyy');
      
      // Cor condicional no status
      for (let i = 2; i <= dadosFilt.length + 1; i++) {
        const statusCell = abaRel.getRange(i, 7);
        const statusValue = String(statusCell.getValue());
        
        if (statusValue.includes("VENCIDA")) {
          statusCell.setBackground("#f8d7da");
        } else if (statusValue.includes("BREVE") || statusValue.includes("ATENÇÃO")) {
          statusCell.setBackground("#fff3cd");
        } else {
          statusCell.setBackground("#d4edda");
        }
      }
      
      abaRel.setFrozenRows(1);
      abaRel.autoResizeColumns(1, headerNovo.length);
      
      ui.alert(
        "Relatório Gerado",
        `${dadosFilt.length} atas encontradas. Verifique status de vigência.`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert("Info", "Nenhuma ata encontrada.", ui.ButtonSet.OK);
    }
    
  } catch (e) {
    ui.alert("Erro", e.message, ui.ButtonSet.OK);
  }
}

// =================================================================
// --- FUNÇÕES AUXILIARES ---
// =================================================================

function _norm(t) { 
  return t ? String(t).trim().toUpperCase() : ""; 
}

function _normalizarEmpenhoBusca(valor) {
  const texto = String(valor || '').trim();
  
  if (!texto) return '';

  const somenteDigitos = texto.replace(/\D/g, '');
  if (!somenteDigitos) return '';
  
  if (somenteDigitos.length === 8) {
    return somenteDigitos;
  }

  if (somenteDigitos.length > 8) {
    const ano = somenteDigitos.slice(0, 4);
    const numero = somenteDigitos.slice(-4);
    
    if (/^\d{4}$/.test(ano) && /^\d{4}$/.test(numero)) {
      return ano + numero;
    }
  }

  return somenteDigitos;
}
