// =================================================================
// --- BLOCO 08: GESTÃO DE ESTOQUE ---
// =================================================================

function sincronizarControleEstoque() {
  const planilhaLocal = SpreadsheetApp.getActiveSpreadsheet();
  const interfaceUsuario = SpreadsheetApp.getUi();

  // --- CONFIGURAÇÃO ---
  const ID_PLANILHA_HISTORICO = CONFIG.ids.movimentoEstoque; 
  const NOME_ABA_HISTORICO = CONFIG.abas.mov2022_2025 || "Mov2022-2025";
  const MESES_ANALISE = 36; // Período base para média

  // --- DEFINIÇÃO DE REGRAS (LÓGICA BLINDADA) ---
  const TIPOS_CONSUMO_REAL = [
    "REQ", "RM", "RFS", "RDD", "RCS", "ACA", "NFS"
  ];
  const TIPOS_DEVOLUCAO = [
    "DEV", "DRM", "RRQ"
  ];

  try {
    if (!ID_PLANILHA_HISTORICO) throw new Error("ID_MOVIMENTO_ESTOQUE não encontrado no Config.");
    
    // 1. MAPEAMENTO DE EMPENHOS
    const destinosConsulta = [
      { id: CONFIG.ids.materiais, aba: CONFIG.abas.materiais },
      { id: CONFIG.ids.medicamentos, aba: CONFIG.abas.medicamentos }
    ];
    const mapaEmpenhos = new Map();
    
    destinosConsulta.forEach(d => {
      try {
        const ss = SpreadsheetApp.openById(d.id);
        const ws = ss.getSheetByName(d.aba);
        if (ws && ws.getLastRow() > 1) {
          ws.getRange(2, 1, ws.getLastRow()-1, 6).getValues().forEach(r => {
            const cod = _norm(r[5]);
            if(cod) mapaEmpenhos.set(cod, { empenho: r[0], fornecedor: r[4] });
          });
        }
      } catch (e) { console.log("Erro leitura remota: " + e.message); }
    });
    
    // 2. PROCESSAMENTO DO HISTÓRICO REAL
    const mapaConsumoReal = new Map();
    try {
      const ssHist = SpreadsheetApp.openById(ID_PLANILHA_HISTORICO);
      const wsHist = ssHist.getSheetByName(NOME_ABA_HISTORICO);
      if (!wsHist) {
        throw new Error(`Aba '${NOME_ABA_HISTORICO}' não encontrada na planilha de Movimento.`);
      }
      const dadosHist = wsHist.getDataRange().getValues();
      for (let i = 1; i < dadosHist.length; i++) {
        const row = dadosHist[i];
        
        const codigoItem = _norm(row[0]);
        const dataMov = row[2];
        const qtd = parseFloat(row[3]) || 0;
        
        let rawDoc = String(row[1]);
        if (rawDoc.includes("-")) {
          rawDoc = rawDoc.split("-").pop();
        }
        const codigoDoc = rawDoc.trim().toUpperCase();

        let ano = 0;
        if (Object.prototype.toString.call(dataMov) === '[object Date]') ano = dataMov.getFullYear();
        else if (String(dataMov).includes("/")) ano = parseInt(String(dataMov).split("/")[2]);
        
        if (ano === 2024) continue;
        
        // LÓGICA DE SOMA/SUBTRAÇÃO
        if (TIPOS_CONSUMO_REAL.includes(codigoDoc)) {
          mapaConsumoReal.set(codigoItem, (mapaConsumoReal.get(codigoItem) || 0) + qtd);
        } 
        else if (TIPOS_DEVOLUCAO.includes(codigoDoc)) {
          mapaConsumoReal.set(codigoItem, (mapaConsumoReal.get(codigoItem) || 0) - qtd);
        }
      }
    } catch (e) { throw new Error("Erro no Histórico: " + e.message); }

    // 3. CRUZAMENTO E CÁLCULOS FINAIS
    const nomeAbaDados = CONFIG.abas.dadosEstoque || "DadosEstoque";
    if (!CONFIG.ids.fonteDadosGeral) throw new Error("ID_FONTE_GERAL não configurado.");
    
    const ssDadosGeral = SpreadsheetApp.openById(CONFIG.ids.fonteDadosGeral);
    const abaDados = ssDadosGeral.getSheetByName(nomeAbaDados);
    if (!abaDados) throw new Error(`Aba '${nomeAbaDados}' não encontrada.`);

    // --- Lendo a partir da linha 3 - Nova Estrutura (40 Colunas) ---
    const dadosSistema = abaDados.getRange(3, 1, abaDados.getLastRow() - 2, 40).getValues();

    const output = [];
    dadosSistema.forEach(r => {
      const cod = _norm(r[1]); // Col B (Item)
      if (!cod) return;
      
      const tipo = _definirTipo(r[0], cod); // r[0] = Local
      const desc = r[2]; // Col C (Descricao)
      const empenhoInfo = mapaEmpenhos.get(cod) || { empenho: "---", fornecedor: "---" };
      
      // AJUSTE: r[19] e r[26] foram deslocados para r[20] e r[27]
      const processo = [r[20], r[27]].filter(Boolean).join(" / ");
      
      const estoque = parseFloat(r[7]) || 0; // Col H (SaldoAtual)
      const cmmSistemaAntigo = parseFloat(r[8]) || 0; // Col I (Cmm12)
      
      // Cálculos
      const totalSaidasLiquidas = Math.max(0, mapaConsumoReal.get(cod) || 0);
      const cmmNovo = totalSaidasLiquidas / MESES_ANALISE; 
      
      // Cobertura
      let diasCob = 0;
      let dataEsgotamento = "";

      if (cmmNovo > 0) {
        diasCob = Math.floor(estoque / (cmmNovo / 30));
        const dt = new Date(); dt.setDate(dt.getDate() + diasCob);
        dataEsgotamento = diasCob > 365 ? "Estável (>1 ano)" : dt;
      } else {
        diasCob = 9999;
        dataEsgotamento = estoque > 0 ? "S/ Consumo" : "Zerado";
      }
      
      // Status
      let status = "Suprimento Ok";
      let corStatus = CONFIG.cores.ALERTA_OK;
      const consumo60dias = cmmNovo * 2;
      
      if (estoque <= 0) {
        status = "Crítico"; corStatus = CONFIG.cores.ALERTA_CRITICO;
      } else if (estoque < consumo60dias) {
        status = "Atenção (60d)"; corStatus = "#ff9900"; 
      } else if (diasCob < 90) {
        status = "Atenção"; corStatus = CONFIG.cores.ALERTA_ATENCAO;
      }

      // Sugestão
      let sugestao = 0;
      let prevSugestao = "";

      if (cmmNovo > 0) {
        const meta = cmmNovo * 6;
        if (estoque < meta) sugestao = Math.round(meta - estoque);
        
        const estoqueFuturo = estoque + sugestao;
        const diasFuturos = Math.floor(estoqueFuturo / (cmmNovo / 30));
        const dtf = new Date(); dtf.setDate(dtf.getDate() + diasFuturos);
        prevSugestao = dtf;
      } else {
        prevSugestao = "S/ Consumo";
      }

      output.push([
        tipo, cod, desc, empenhoInfo.fornecedor, empenhoInfo.empenho, 
        estoque, cmmNovo, status, diasCob === 9999 ? "-" : diasCob, 
        dataEsgotamento, processo, sugestao, prevSugestao, 
        "|", cmmSistemaAntigo, (cmmNovo - cmmSistemaAntigo)
      ]);
      output[output.length-1].cor = corStatus;
    });

    // 4. ESCRITA NA PLANILHA
    let aba = planilhaLocal.getSheetByName(CONFIG.abas.estoqueRemoto);
    if (!aba) aba = planilhaLocal.insertSheet(CONFIG.abas.estoqueRemoto);

    aba.clear();
    
    const header = [
      "Tipo", "Item", "Descrição", "Fornecedor", "Empenho", 
      "Estoque", "CMM (Real Higienizado)", "Status", "Cobertura (Dias)", 
      "Data Esgot.", "Processo SEI", "Sugestão (6 Meses)", "Prev. Sugestão",
      "|", "CMM (Sistema Antigo)", "Diferença"
    ];
    
    aba.getRange(1, 1, 1, header.length).setValues([header])
       .setFontWeight("bold").setBackground("#0c343d").setFontColor("white");
    
    if (output.length > 0) {
      const dados = output.map(L => L.slice(0, 16));
      const cores = output.map(L => [L.cor]);
      
      aba.getRange(2, 1, dados.length, 16).setValues(dados);
      aba.getRange(2, 6, dados.length, 2).setNumberFormat("#,##0.00"); 
      aba.getRange(2, 12, dados.length, 1).setNumberFormat("#,##0");
      aba.getRange(2, 10, dados.length, 1).setNumberFormat("dd/mm/yyyy");
      aba.getRange(2, 13, dados.length, 1).setNumberFormat("dd/mm/yyyy"); 
      aba.getRange(2, 15, dados.length, 2).setNumberFormat("#,##0.00"); 
      aba.getRange(2, 8, dados.length, 1).setBackgrounds(cores).setFontWeight("bold");
      
      aba.setFrozenRows(1);
      if (aba.getFilter()) aba.getFilter().remove();
      aba.getDataRange().createFilter();
    }
    
    interfaceUsuario.alert("Sucesso", "Estoque Recalculado!", interfaceUsuario.ButtonSet.OK);

  } catch (e) {
    console.error(e);
    interfaceUsuario.alert("Erro", e.message, interfaceUsuario.ButtonSet.OK);
  }
}

function _norm(t) { return t ? String(t).trim().toUpperCase() : ""; }
function _definirTipo(ref, cod) {
  if (ref === "FAR" || /^\d/.test(cod)) return "MEDICAMENTO";
  return "MATERIAL";
}