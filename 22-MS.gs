/**
 * ATUALIZAÇÃO GUIA MS - CORREÇÃO DE COLUNAS
 * * Estrutura de Saída (Guia MS):
 * - Coluna E (Índice 4): Consumo Médio Mensal (Soma Qtd 2025 / 12)
 * - Coluna F (Índice 5): Quantidade Total 2025 (Soma Qtd 2025)
 * - Coluna G (Índice 6): Preço Unitário (Do lançamento mais recente)
 * * * Fontes de Dados:
 * 1. EntradaEmpenhos (Externa)
 * 2. Temp (Local)
 */

function atualizarGuiaMS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // --- 1. CONFIGURAÇÕES E IDs ---
  const ID_ENTRADA_EMPENHOS = "1s44YD2ozLAbBdGQbBE5iW7HcUzvQULZqd4ynYlV_HXA";
  
  // --- 2. ACESSAR GUIA ALVO (MS) ---
  const sheetMS = ss.getSheetByName("MS");
  if (!sheetMS) { 
    ui.alert("ERRO: Guia 'MS' não encontrada!");
    return; 
  }
  
  const lastRow = sheetMS.getLastRow();
  if (lastRow < 2) { 
    ui.alert("Guia 'MS' vazia."); 
    return;
  }
  
  // Mapeamos da Coluna A (1) até a G (7)
  // Índices no Array: A=0 ... E=4, F=5, G=6
  const rangeMS = sheetMS.getRange(2, 1, lastRow - 1, 7);
  const valoresMS = rangeMS.getValues();

  // --- 3. COLETAR DADOS (CONSOLIDAÇÃO) ---
  const mapaDados = {}; 
  // Estrutura: { 'CODIGO': [ { data: Date, qtd: Number, preco: Number, linhaOriginal: Number }, ... ] }

  // === 3.a. Ler EntradaEmpenhos (EXTERNA) ===
  try {
    const ssEnt = SpreadsheetApp.openById(ID_ENTRADA_EMPENHOS);
    let sheetEntradas = ssEnt.getSheetByName("EntradaEmpenhos");
    if (!sheetEntradas) sheetEntradas = ssEnt.getSheets().find(s => s.getName().includes("Entrada") || s.getName().includes("Empenho"));

    if (sheetEntradas) {
      const dadosEnt = sheetEntradas.getDataRange().getValues();
      const cabecalhoEnt = dadosEnt[0];

      // Mapeamento
      const idxItemEnt = buscarColuna(cabecalhoEnt, ["Item", "Material", "Código"]);
      const idxPrecoEnt = 8;      // Col I
      const idxQtdEnt = 11;       // Col L
      const idxDataEnt = 14;      // Col O

      if (idxItemEnt > -1) {
        for (let i = 1; i < dadosEnt.length; i++) {
          const row = dadosEnt[i];
          const item = row[idxItemEnt];
          if (!item) continue;

          if (!mapaDados[item]) mapaDados[item] = [];
          
          let dataValida = converterData(row[idxDataEnt]);
          let qtd = tratarNumero(row[idxQtdEnt]);
          let preco = tratarNumero(row[idxPrecoEnt]);

          mapaDados[item].push({
            data: dataValida,
            qtd: qtd,
            preco: preco,
            linhaOriginal: i, // Desempate
            origem: "EXTERNA"
          });
        }
      }
    }
  } catch(e) {
    ui.alert("Aviso: Erro ao ler EntradaEmpenhos Externa: " + e.message);
  }

  // === 3.b. Ler Guia Temp (LOCAL) ===
  try {
    const sheetTemp = ss.getSheetByName("Temp");
    if (sheetTemp) {
      const dadosTemp = sheetTemp.getDataRange().getValues();
      const cabecalhoTemp = dadosTemp[0];

      const idxItemT = buscarColuna(cabecalhoTemp, ["item", "código", "codigo"]);
      const idxQtdT = buscarColuna(cabecalhoTemp, ["quantidade", "qtd"]);
      const idxValorTotalT = buscarColuna(cabecalhoTemp, ["valor", "valor total"]);
      const idxDataT = buscarColuna(cabecalhoTemp, ["dt-trans", "data"]);

      if (idxItemT > -1) {
        for (let i = 1; i < dadosTemp.length; i++) {
          const row = dadosTemp[i];
          const item = row[idxItemT];
          if (!item) continue;

          if (!mapaDados[item]) mapaDados[item] = [];

          let dataValida = (idxDataT > -1) ? converterData(row[idxDataT]) : new Date();
          let qtd = tratarNumero(row[idxQtdT]);
          let valorTotal = tratarNumero(row[idxValorTotalT]);
          
          let precoUnitario = 0;
          if (qtd > 0) precoUnitario = valorTotal / qtd;

          // Peso extra na linhaOriginal para priorizar Temp em caso de empate exato de data
          mapaDados[item].push({
            data: dataValida,
            qtd: qtd,
            preco: precoUnitario,
            linhaOriginal: i + 1000000, 
            origem: "TEMP"
          });
        }
      }
    }
  } catch (e) {
    console.warn("Guia Temp não processada: " + e.message);
  }

  // --- 4. PROCESSAMENTO DOS DADOS NA MS ---
  const resultados = valoresMS.map(linha => {
    const itemCodigo = linha[0]; // Coluna A
    
    if (!itemCodigo) return linha;

    const historico = mapaDados[itemCodigo];

    if (historico && historico.length > 0) {
      
      // Soma Qtd 2025
      let somaQtd2025 = 0;
      historico.forEach(h => {
        if (h.data.getFullYear() === 2025) {
          somaQtd2025 += h.qtd;
        }
      });
      
      // Coluna E: Consumo Médio Mensal
      linha[4] = somaQtd2025 / 12; 

      // Coluna F: Quantidade Total 2025
      linha[5] = somaQtd2025;

      // Coluna G: Preço Unitário Mais Recente
      historico.sort((a, b) => {
        const timeDiff = b.data.getTime() - a.data.getTime(); 
        if (timeDiff !== 0) return timeDiff;
        return b.linhaOriginal - a.linhaOriginal; 
      });

      // Pega o melhor preço (mais recente > 0)
      let melhorPreco = historico[0].preco;
      if (melhorPreco === 0 && historico.length > 1) {
        for (let k = 0; k < historico.length; k++) {
           if (historico[k].preco > 0) {
             melhorPreco = historico[k].preco;
             break;
           }
        }
      }
      linha[6] = melhorPreco; // Col G

    } else {
      linha[4] = 0; // CMM
      linha[5] = 0; // Qtd Total
      // Coluna G mantém o valor original se não achar nada? Ou zera? 
      // Se quiser zerar, descomente abaixo:
      // linha[6] = 0; 
    }

    return linha;
  });

  // --- 5. SALVAR ---
  rangeMS.setValues(resultados);
  ui.alert("Atualização Concluída! Colunas E (CMM), F (Qtd Total) e G (Preço) atualizadas.");
}

// --- FUNÇÕES AUXILIARES ---

function tratarNumero(valor) {
  if (typeof valor === 'number') return valor;
  if (!valor) return 0;
  
  let str = String(valor).trim();
  str = str.replace(/[R$\s]/g, "");
  
  if (str.includes(",") && str.includes(".")) {
    str = str.replace(/\./g, "").replace(",", ".");
  } else if (str.includes(",")) {
    str = str.replace(",", ".");
  }
  
  return parseFloat(str) || 0;
}

function converterData(valor) {
  if (valor instanceof Date) return valor;
  if (!valor) return new Date(2020, 0, 1);
  
  if (String(valor).trim() === "2025") return new Date(2025, 11, 31);

  let data = new Date(valor);
  if (isNaN(data.getTime())) return new Date(2020, 0, 1);
  return data;
}

function buscarColuna(cabecalho, nomesPossiveis) {
  for (let i = 0; i < cabecalho.length; i++) {
    const header = String(cabecalho[i]).toUpperCase().trim();
    for (let nome of nomesPossiveis) {
      if (header === nome.toUpperCase().trim() || header.includes(nome.toUpperCase().trim())) return i;
    }
  }
  return -1;
}