// =================================================================
// --- BLOCO 7: COMPILAÇÃO LOCAL (COM VALIDAÇÃO OFICIAL E CORREÇÃO DE LIMITE) ---
// =================================================================

function compilarDadosLocal() { 
  try {
    const dados = obterDadosEntradasGlobal();
    compilarDados(dados); 
  } catch (e) {
    SpreadsheetApp.getUi().alert("Erro ao buscar dados para compilação: " + e.message);
  }
}

/**
 * Extrai o ano (AAAA) de um empenho no formato AAAANNNN.
 * Retorna null se não conseguir extrair um ano válido.
 */
function _extrairAnoEmpenho(valorEmpenho) {
  if (valorEmpenho === null || valorEmpenho === undefined) return null;

  // Mantém apenas dígitos (caso venha como texto com espaços, etc.)
  const digits = String(valorEmpenho).replace(/\D/g, "");
  if (digits.length < 4) return null;

  const ano = parseInt(digits.slice(0, 4), 10);
  if (!ano || ano < 2000 || ano > 2100) return null;
  return ano;
}

function compilarDados(dadosGlobais) {
  var ui = SpreadsheetApp.getUi();
  try {
    if (!dadosGlobais) {
      dadosGlobais = obterDadosEntradasGlobal();
    }

    var ssDestino = SpreadsheetApp.getActiveSpreadsheet();
    var abaDestino = ssDestino.getSheetByName(CONFIG.destino.nomeAba);

    // Compilados é obrigatório (é onde vamos escrever)
    if (!abaDestino) {
      throw new Error(`Aba de destino '${CONFIG.destino.nomeAba}' não encontrada.`);
    }

    // =================================================================
    // Rec.Provisorio
    // =================================================================
    let recProvisorioMap = new Map();
    try {
      if (CONFIG && CONFIG.ids && CONFIG.ids.materiais) {
        const ssMat = SpreadsheetApp.openById(CONFIG.ids.materiais);
        const nomeAbaRec = (CONFIG.abas && CONFIG.abas.recProvisorio) ? CONFIG.abas.recProvisorio : "Rec.Provisorio";
        const abaRecProvisorioRemota = ssMat.getSheetByName(nomeAbaRec);

        if (abaRecProvisorioRemota) {
          const lastRowRec = abaRecProvisorioRemota.getLastRow();
          if (lastRowRec >= 2) {
            const recData = abaRecProvisorioRemota.getRange(2, 1, lastRowRec - 1, 6).getValues();
            recData.forEach(row => {
              const itemCode = _norm(row[0]);
              const qtd = Math.round(parseFloat(row[2]) || 0);
              const empenhoStr = String(row[5]).trim();
              if (itemCode && empenhoStr.includes('/')) {
                const parts = empenhoStr.split('/');
                if (parts.length === 2) {
                  const numero = parts[0].padStart(4, '0');
                  const ano = (parts[1].length === 2) ? '20' + parts[1] : parts[1];
                  recProvisorioMap.set(`${ano}${numero}|${itemCode}`, qtd);
                }
              }
            });
          }
        } else {
          console.warn(`Aviso: Aba '${nomeAbaRec}' não encontrada.`);
        }
      }
    } catch (e) {
      console.warn("Aviso: Falha ao ler Rec.Provisorio remoto. Detalhes: " + e.message);
      recProvisorioMap = new Map();
    }

    // 1. Mapa de Qtd Oficial (Dados Globais)
    const mapaQtdOficial = new Map();
    
    // --- NOVO: Mapa para armazenar a última data de entrada POR ITEM ---
    const mapaUltimaDataPorItem = new Map(); 

    dadosGlobais.forEach(r => {
       const emp = String(r[0]).trim();
       const cod = _norm(r[2]);
       const qS = Math.round(parseFloat(r[19]) || 0); // Índice 19 = QTDE_RECEBIDA
       const dataStr = r[23]; // --- NOVO: Índice 23 = Coluna X (DATA_ULT_ENTRADA) ---

       if (emp && cod) {
          const k = `${emp}|${cod}`;
          const atual = mapaQtdOficial.get(k) || 0;
          mapaQtdOficial.set(k, atual + qS);
       }
       
       // --- NOVO: Lógica para salvar sempre a maior data baseada no código do item ---
       if (cod && dataStr) {
           const dataAtual = new Date(dataStr).getTime();
           if (!isNaN(dataAtual)) {
               const dataSalva = mapaUltimaDataPorItem.get(cod);
               if (!dataSalva || dataAtual > dataSalva.getTime()) {
                   mapaUltimaDataPorItem.set(cod, new Date(dataStr));
               }
           }
       }
    });

    // Mapeamento Eliminadas
    const mapaStatusEliminados = new Map();
    dadosGlobais.forEach(linha => {
        if (linha.some(celula => _norm(celula) === "ELIMINADA")) {
            const empenho = parseInt(linha[0], 10);
            const codigo = _norm(linha[2]);
            if (empenho && codigo) {
              mapaStatusEliminados.set(`${empenho}|${codigo}`, "Eliminada");
            }
        }
    });

    var dadosCompilados = [];
    const fontes = [
        {id: CONFIG.ids.materiais, nomeAba: CONFIG.abas.materiais},
        {id: CONFIG.ids.medicamentos, nomeAba: CONFIG.abas.medicamentos}
    ];

    fontes.forEach(fonte => {
      try {
        var ssFonte = SpreadsheetApp.openById(fonte.id);
        var abaFonte = ssFonte.getSheetByName(fonte.nomeAba);
        if (abaFonte) {
          var dadosFonte = abaFonte.getDataRange().getValues();
          if (dadosFonte.length > 1) dadosCompilados.push(...dadosFonte.slice(1));
        }
      } catch (e) {
        console.warn(`Aviso: Não foi possível ler a fonte ${fonte.id} (${e.message})`);
      }
    });

    // --- FILTRO POR ANO DO EMPENHO ---
    const anoAtual = new Date().getFullYear();
    const anoMinimo = anoAtual - 1;
    dadosCompilados = dadosCompilados.filter(linha => {
      const ano = _extrairAnoEmpenho(linha[0]);
      if (ano === null) return false;
      if (ano < anoMinimo) return false;
      return true;
    });

    // --- PROCESSAMENTO FINAL ---
    dadosCompilados.forEach(linha => {
      // --- NOVO: Garante que a linha tenha pelo menos 22 posições (A a V) ---
      while(linha.length < 22) {
          linha.push("");
      }

      const empenho = linha[0] ? parseInt(linha[0], 10) : null;
      const codigo = _norm(linha[5]);
      const chave = (empenho && codigo) ? `${empenho}|${codigo}` : null;
      const isProvisorio = (chave && recProvisorioMap.has(chave));

      if (chave && mapaStatusEliminados.has(chave)) {
          linha[18] = "Eliminada";
      } else {
          var vI = Math.round(parseFloat(linha[8]) || 0);  
          var vP_Visual = Math.round(parseFloat(linha[15]) || 0);
          var qS_Real_Oficial = 0;
          
          if (chave) qS_Real_Oficial = mapaQtdOficial.get(chave) || 0;

          var vQ = vI - vP_Visual;

          linha[8] = vI;
          linha[15] = vP_Visual;
          linha[16] = vQ;

          var statusOrig = linha[18] ? linha[18].toString().trim() : '';
          let statusFinal = _calcularStatusUnificado(vI, qS_Real_Oficial, vQ, isProvisorio, statusOrig);

          if (vQ === 0 && (statusFinal === "Recebido a Maior" || statusFinal === "Pendente")) {
              statusFinal = isProvisorio ? "Recebimento Provisório" : "Concluído";
          }
          linha[18] = statusFinal;
      }

      // --- NOVO: Preencher a coluna V (índice 21) com a última data do item ---
      if (codigo && mapaUltimaDataPorItem.has(codigo)) {
          linha[21] = mapaUltimaDataPorItem.get(codigo);
      }
    });

    // ======================================================================
    // === ESCRITA OTIMIZADA ===
    // ======================================================================
    if (dadosCompilados.length > 0) {
      // --- NOVO: Alterado de "A:U" para "A:V" para considerar a nova coluna ---
      abaDestino.unhideColumn(abaDestino.getRange("A:V")); 

      const numLinhasDados = dadosCompilados.length;
      const numColunasDados = dadosCompilados[0].length; // Agora será 22 devido ao "while" acima
      const maxLinhasAtual = abaDestino.getMaxRows();
      const maxColunasAtual = abaDestino.getMaxColumns();

      if (maxLinhasAtual > 1) {
        abaDestino.getRange(2, 1, maxLinhasAtual - 1, maxColunasAtual).clearContent().clearFormat();
      }

      const margemSeguranca = 50;
      const linhasNecessariasTotal = numLinhasDados + 1;
      const linhasAlvo = linhasNecessariasTotal + margemSeguranca;

      if (maxLinhasAtual > linhasAlvo) {
         const linhasParaDeletar = maxLinhasAtual - linhasAlvo;
         if (linhasParaDeletar > 0) {
            try {
               abaDestino.deleteRows(linhasAlvo + 1, linhasParaDeletar);
            } catch (errDel) {}
         }
      } else if (maxLinhasAtual < linhasNecessariasTotal) {
         const linhasParaAdicionar = linhasNecessariasTotal - maxLinhasAtual;
         abaDestino.insertRowsAfter(maxLinhasAtual, linhasParaAdicionar);
      }

      // Escreve as matrizes
      abaDestino.getRange(2, 1, numLinhasDados, numColunasDados).setValues(dadosCompilados);

      var numRows = dadosCompilados.length;
      abaDestino.getRange(2, 9, numRows, 1).setNumberFormat("0");
      abaDestino.getRange(2, 16, numRows, 1).setNumberFormat("0");
      abaDestino.getRange(2, 17, numRows, 1).setNumberFormat("0");

      // --- NOVO: Alterado a leitura e mapeamento de fundo de 21 para 22 colunas ---
      var valores = abaDestino.getRange(2, 1, numRows, 22).getValues();

      var backgrounds = valores.map(l => {
        return new Array(22).fill(CONFIG.cores[_norm(l[18])] || null);
      });

      abaDestino.getRange(2, 1, numRows, 22).setBackgrounds(backgrounds);

      abaDestino.hideColumns(2, 3);
      abaDestino.hideColumns(8, 1);
      abaDestino.hideColumns(10, 6);
      abaDestino.hideColumns(18, 1);
    }
  } catch (e) { 
    ui.alert("Erro Compilados Local", "Detalhes: " + e.message, ui.ButtonSet.OK); 
  }
}