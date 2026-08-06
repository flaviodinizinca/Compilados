// =================================================================
// --- BLOCO 15: DISTRIBUIÇÃO POR EQUIPE (DEFINITIVO COM 31 COLUNAS) ---
// =================================================================

function abrirSeletorPlanejadores() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: 'Segoe UI', sans-serif; padding: 15px; background-color: #f3f3f3; }
      h3 { margin-top: 0; color: #333; font-size: 16px; }
      .container { background: white; padding: 15px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
      .checkbox-group { display: flex; flex-direction: column; gap: 10px; margin-bottom: 15px; }
      label { cursor: pointer; font-size: 14px; color: #444; display: flex; align-items: center; }
      input[type="checkbox"] { transform: scale(1.2); margin-right: 10px; accent-color: #1c4587; }
      
      .actions { display: flex; gap: 10px; margin-bottom: 15px; }
      .small-btn { background: #ddd; border: none; padding: 5px 10px; border-radius: 4px; cursor: pointer; font-size: 12px; }
      .small-btn:hover { background: #ccc; }

      .btn { 
        background-color: #1c4587; color: white; border: none; 
        padding: 12px 0; width: 100%; border-radius: 5px; 
        font-size: 16px; font-weight: bold; cursor: pointer; 
        transition: background 0.3s; }
      .btn:hover { background-color: #0f2e5e; }
      .loading { display: none; color: #666; font-size: 13px; text-align: center; margin-top: 10px; }
    </style>

    <div class="container">
      <h3>Selecionar Planejadores</h3>
      
      <div class="actions">
        <button type="button" class="small-btn" onclick="toggle(true)">Marcar Todos</button>
        <button type="button" class="small-btn" onclick="toggle(false)">Desmarcar</button>
      </div>

      <div class="checkbox-group">
        <label><input type="checkbox" name="planner" value="Lorena" checked> 👤 Lorena</label>
        <label><input type="checkbox" name="planner" value="Katia" checked> 👤 Katia</label>
        <label><input type="checkbox" name="planner" value="Leonardo" checked> 👤 Leonardo</label>
        <label><input type="checkbox" name="planner" value="Moises" checked> 👤 Moises</label>
        <label><input type="checkbox" name="planner" value="Rafaelle" checked> 👤 Rafaelle</label>
        <label><input type="checkbox" name="planner" value="Luciana" checked> 👤 Luciana</label>
      </div>
      
      <button type="button" class="btn" onclick="enviar()">Atualizar Selecionados</button>
      <div id="msg" class="loading">🔄 Processando... aguarde a conclusão.</div>
    </div>

    <script>
      function toggle(estado) {
        const boxes = document.getElementsByName('planner');
        for(let box of boxes) box.checked = estado;
      }

      function enviar() {
        const selecionados = [];
        const boxes = document.getElementsByName('planner');
        for (let box of boxes) {
          if (box.checked) selecionados.push(box.value);
        }

        if (selecionados.length === 0) {
          alert("Selecione pelo menos um planejador.");
          return;
        }

        document.getElementById('msg').style.display = 'block';
        google.script.run
          .withSuccessHandler(function() { google.script.host.close(); })
          .withFailureHandler(function(e) { alert("Erro: " + e); google.script.host.close(); })
          .processarSelecaoPlanejadores(selecionados);
      }
    </script>
  `).setWidth(300).setHeight(420);

  SpreadsheetApp.getUi().showModalDialog(html, 'Distribuição de Equipe');
}

function processarSelecaoPlanejadores(listaNomes) {
  distribuirDadosPorEquipe(listaNomes);
}

function testarDistribuicaoNovoLayout() {
  distribuirDadosPorEquipe(null, true);
}

function distribuirDadosPorEquipe(alvos = null, isTest = false) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    if (!CONFIG.ids.fonteDadosGeral) throw new Error("ID_FONTE_GERAL não configurado.");
    if (!CONFIG.ids.painelEquipe) throw new Error("ID_PAINEL_EQUIPE não configurado.");
    
    const ssFonteGeral = SpreadsheetApp.openById(CONFIG.ids.fonteDadosGeral);
    const ssDestino = SpreadsheetApp.openById(CONFIG.ids.painelEquipe);
    const regras = carregarRegrasDistribuicao(ssFonteGeral);

    const nomeAbaDados = CONFIG.abas.dadosEstoque || "DadosEstoque";
    const abaDadosAtuais = ssFonteGeral.getSheetByName(nomeAbaDados);
    if (!abaDadosAtuais) throw new Error(`Aba '${nomeAbaDados}' não encontrada na Fonte Geral.`);

    const primeiraLinhaDados = 3;
    const lastRowDados = abaDadosAtuais.getLastRow();
    if (lastRowDados < primeiraLinhaDados) throw new Error(`Aba '${nomeAbaDados}' não possui dados a partir da linha ${primeiraLinhaDados}.`);
    const numLinhasDados = lastRowDados - (primeiraLinhaDados - 1);

    const rangeEstoque = abaDadosAtuais.getRange(primeiraLinhaDados, 1, numLinhasDados, 37);
    const vDadosRaw = rangeEstoque.getValues();
    const vDadosDisplay = rangeEstoque.getDisplayValues();

    // Mapeamento da Coluna AJ original
    const mapCodParaAJ = new Map();
    for (let i = 0; i < vDadosDisplay.length; i++) {
      const cod = _norm(vDadosDisplay[i][1]);
      const valAJ = vDadosDisplay[i][35] || ""; // Coluna AJ
      if (cod && !mapCodParaAJ.has(cod)) mapCodParaAJ.set(cod, valAJ);
    }
    
    const setGeladeira = new Set();
    const abaGeladeira = ssFonteGeral.getSheetByName(CONFIG.abas.itensGeladeira || "ItensGeladeira");
    if (abaGeladeira && abaGeladeira.getLastRow() > 0) {
      abaGeladeira.getRange(1, 1, abaGeladeira.getLastRow(), 1).getValues().forEach(r => {
        if (r[0]) setGeladeira.add(_norm(r[0]));
      });
    }

    const mapItemParaPDM = new Map();
    const abaPDM = ssFonteGeral.getSheetByName("PDM");
    if (abaPDM) {
      const dadosPDM = abaPDM.getDataRange().getValues();
      for (let i = 1; i < dadosPDM.length; i++) {
        const item = _norm(dadosPDM[i][0]);
        const codPDM = _norm(dadosPDM[i][1]);
        if (item && codPDM) mapItemParaPDM.set(item, codPDM);
      }
    }

    const abaEntradas = ssFonteGeral.getSheetByName(CONFIG.abas.entradaEmpenhos || "EntradaEmpenhos");
    const mapEmpenhosPendentes = new Map();
    const mapSomaPDM = new Map();
    const mapSomaItem = new Map();
    
    if (abaEntradas && abaEntradas.getLastRow() >= 2) {
      const dadosEntradas = abaEntradas.getRange(2, 1, abaEntradas.getLastRow() - 1, 23).getValues();
      dadosEntradas.forEach(r => {
        const emp = String(r[0]).trim();        
        const dataEmissao = r[1];               
        const cod = _norm(r[2]);                
        const forn = String(r[8]).trim();       
        const modalidade = String(r[10]).trim().toUpperCase(); 
        const qtdEmp = parseFloat(r[17]) || 0;  
        const qtdRec = parseFloat(r[19]) || 0;  
        const saldoEmp = parseFloat(r[21]) || 0;

        let valorS = r[18]; 
        if (typeof valorS === 'string') valorS = parseFloat(valorS.replace(/\./g, '').replace(',', '.'));
        valorS = parseFloat(valorS) || 0;

        if (cod && qtdEmp !== qtdRec && saldoEmp > 0) {
          if (!mapEmpenhosPendentes.has(cod)) mapEmpenhosPendentes.set(cod, []);
          mapEmpenhosPendentes.get(cod).push({
            empenho: emp, fornecedor: forn, status: "Pendente",
            qtdEmpenho: qtdEmp, saldoEmpenho: saldoEmp
          });
        }

        let ano = "";
        if (dataEmissao instanceof Date) ano = dataEmissao.getFullYear().toString();
        else if (typeof dataEmissao === 'string' && dataEmissao.includes("2026")) ano = "2026";

        if (ano === "2026" && modalidade === "DISPENSA 75-II") {
          mapSomaItem.set(cod, (mapSomaItem.get(cod) || 0) + valorS);
          const codPDM = mapItemParaPDM.get(cod);
          if (codPDM) mapSomaPDM.set(codPDM, (mapSomaPDM.get(codPDM) || 0) + valorS);
        }
      });
    }

    const cabecalho = [
      "Item", "Descrição", "Estoque", "CMM Atual", "Saldo em Dias",
      "Observações (Equipe)", "Qtd Empenho", "Quantidade AE", "Saldo Virtual + Saldo Físico",
      "Notes", "Quantidade Original", "Saldo Ata", "Preço Unitário", "Vencimento Ata",
      "Fornecedor", "Processo SEI", "Empenho", "Status Empenho", "Saldo Empenho",
      "AE em andamento", "CMA Histórico (15 Meses)", "Status Estoque", "Classificação",
      "Família", "Previsão Esgotamento", "Sugestão Pedido (4 Meses)", "Prev. Esgotamento (Sugestão)",
      "Código PDM", "Valor Grupo PDM", "Soma PDM", "Soma Item"
    ];

    const listaFinal = [];
    const assinaturasUnicas = new Set(); 
    const hoje = new Date();
    hoje.setHours(0,0,0,0);
    
    for (let i = 0; i < vDadosDisplay.length; i++) {
      const rStr = vDadosDisplay[i];
      const rRaw = vDadosRaw[i];
      const cod = _norm(rStr[1]);
      if (!cod) continue;
      
      const desc = rStr[2];
      const familiaLimpa = String(rStr[4]).replace(',', '.').trim().toUpperCase();
      const isNumeric = /^\d/.test(cod);
      
      const destinos = new Set();
      if (regras[cod]) destinos.add(regras[cod]);
      if (regras[familiaLimpa]) destinos.add(regras[familiaLimpa]);
      if (destinos.size === 0) destinos.add(isNumeric ? "Rafaelle" : "Triagem");
      
      const arrayDestinos = Array.from(destinos);
      const isConflito = (arrayDestinos.length > 1);
      
      const estoque = parseFloat(rRaw[7]) || 0; 
      const cmmAtual = parseFloat(rRaw[8]) || 0; 
      const cmaHistorico = cmmAtual * 15;
      const saldoDias = parseFloat(rRaw[10]) || 0; 
      const classificacao = String(rStr[11]).trim(); 
      
      const notesAta = String(rStr[17]).trim(); 
      const notesAnalise = String(rStr[24]).trim();
      const notes = [notesAta, notesAnalise].filter(Boolean).join(" / ");
      
      const procS = String(rStr[18]).trim(); 
      const procY = String(rStr[25]).trim(); 
      let processos = [];
      if (procS) processos.push(procS + " - Ata");
      if (procY) processos.push(procY + " - Novo");
      const processoFinal = processos.join("\n");
      
      const saldoAta = parseFloat(rRaw[21]) || 0; 
      let vencAta = rStr[23]; 
      const dataVencObj = _parseDataSegura(vencAta);
      if (dataVencObj && dataVencObj < hoje) vencAta = ""; 

      const aeAndamento = String(rStr[27]).trim(); 
      const qtdAE = rRaw[28];
      const valorAJ = mapCodParaAJ.get(cod) || "";
      
      const qtdOriginal = rRaw[20] !== undefined && rRaw[20] !== "" ? parseFloat(rRaw[20]) : 0;
      const precoUnitario = rRaw[22] !== undefined && rRaw[22] !== "" ? parseFloat(rRaw[22]) : 0;

      const pdmCode = mapItemParaPDM.get(cod) || "";
      const pdmSum = mapSomaPDM.get(pdmCode) || 0;
      const itemSum = mapSomaItem.get(cod) || 0;

      let statusEstoque = "Suprimento Ok";
      if (estoque <= 0) statusEstoque = "Crítico";
      else if (estoque < (cmmAtual * 2)) statusEstoque = "Atenção (60d)";
      else if (saldoDias > 0 && saldoDias < 90) statusEstoque = "Atenção";
      
      let sugestao = Math.round((cmmAtual * 4) - estoque);
      if (sugestao < 0) sugestao = 0;
      
      let prevEsgot = "Sem Consumo";
      let prevEsgotSug = "Sem Consumo";
      if (cmmAtual > 0) {
        let d1 = new Date(); d1.setDate(d1.getDate() + Math.floor(estoque / (cmmAtual / 30))); prevEsgot = d1;
        let d2 = new Date(); d2.setDate(d2.getDate() + Math.floor((estoque + sugestao) / (cmmAtual / 30))); prevEsgotSug = d2;
      }
      
      const empenhosDesteItem = mapEmpenhosPendentes.get(cod) || [];
      const gerarEInserirLinha = (forn, emp, stat, qEmp, sEmp) => {
        const linhaCompleta = [
            cod, desc, estoque, cmmAtual, saldoDias,
            "", 
            qEmp, qtdAE, valorAJ, notes,
            qtdOriginal, saldoAta, precoUnitario, vencAta,
            forn, processoFinal, emp, stat, sEmp,
            aeAndamento, cmaHistorico, statusEstoque, classificacao,
            familiaLimpa, prevEsgot, sugestao, prevEsgotSug
        ];

        arrayDestinos.forEach(dono => {
            let notaConflito = "";
            let corLinhaVermelha = false;
            if (isConflito) {
                corLinhaVermelha = true;
                const outros = arrayDestinos.filter(d => d !== dono).join(", ");
                notaConflito = `⚠️ COMPARTILHADO:\nItem também gerido por: ${outros}`;
            }

            // RECUPERADO: Regra de deduplicação original. Linhas idênticas são fundidas.
            const assinatura = dono + "|||" + linhaCompleta.join("|||"); 
            if (!assinaturasUnicas.has(assinatura)) {
                assinaturasUnicas.add(assinatura);
                listaFinal.push({
                   responsavel: dono, codItem: cod, linhaDados: linhaCompleta,
                   isConflito: corLinhaVermelha, notaItem: notaConflito,
                   pdmCode: pdmCode, pdmSum: pdmSum, itemSum: itemSum 
                });
            }
        });
      };
      
      if (empenhosDesteItem.length === 0) {
          gerarEInserirLinha("---", "---", "---", "---", "---");
      } else {
          empenhosDesteItem.forEach(e => {
              gerarEInserirLinha(e.fornecedor, e.empenho, e.status, e.qtdEmpenho, e.saldoEmpenho);
          });
      }
    }

    let nomesAbas = [];
    let isFullUpdate = false;

    if (isTest) {
      nomesAbas = ["Teste_Layout"];
      listaFinal.forEach(d => d.responsavel = "Teste_Layout");
    } else {
      if (alvos === null) {
        nomesAbas = Array.from(new Set(listaFinal.map(d => d.responsavel)));
        isFullUpdate = true;
      } else if (Array.isArray(alvos)) {
        nomesAbas = alvos;
      } else {
        nomesAbas = [alvos];
      }
    }

    const listaDadosCOAGE = [];
    
    nomesAbas.forEach(nomeAba => {
      let aba = ssDestino.getSheetByName(nomeAba);
      if (!aba) aba = ssDestino.insertSheet(nomeAba);

      const mapObservacoes = new Map();
      const ultimaLinhaAba = aba.getLastRow();
      const ultimaColunaAba = aba.getLastColumn();
     
      if (ultimaLinhaAba >= 2 && ultimaColunaAba >= 2) { 
        const headersAntigos = aba.getRange(1, 1, 1, ultimaColunaAba).getValues()[0];
        let idxObsAntigo = headersAntigos.findIndex(h => String(h).toUpperCase().includes("OBSERVA"));
        
        if (idxObsAntigo === -1 && ultimaColunaAba >= 25) idxObsAntigo = 24; 

        if (idxObsAntigo !== -1) {
            const rangeCodigos = aba.getRange(2, 1, ultimaLinhaAba - 1, 1).getValues();
            const rangeObs = aba.getRange(2, idxObsAntigo + 1, ultimaLinhaAba - 1, 1).getValues();
            for (let r = 0; r < rangeCodigos.length; r++) {
               const codExistente = _norm(rangeCodigos[r][0]);
               const obsExistente = String(rangeObs[r][0] || "").trim();
               if (codExistente && obsExistente) {
                   const obsSalvaAnterior = mapObservacoes.get(codExistente) || "";
                   if (!obsSalvaAnterior.includes(obsExistente)) {
                       mapObservacoes.set(codExistente, obsSalvaAnterior ? obsSalvaAnterior + " | " + obsExistente : obsExistente);
                   }
               }
            }
        }
      }

      const dadosParaEscrever = [];
      const metadadosLinha = []; 

      listaFinal.filter(d => d.responsavel === nomeAba).forEach(d => {
          const obsSalva = mapObservacoes.get(d.codItem) || "";
          const linhaPronta = [...d.linhaDados];
          linhaPronta[5] = obsSalva; 
          linhaPronta.push(d.pdmCode, d.pdmSum, "", d.itemSum);
          
          dadosParaEscrever.push(linhaPronta);
          metadadosLinha.push({ conflito: d.isConflito, nota: d.notaItem, isGeladeira: setGeladeira.has(d.codItem) });
      });
      
      if (isFullUpdate && !isTest) {
         listaFinal.filter(d => d.responsavel === nomeAba).forEach(d => {
             const obsSalva = mapObservacoes.get(d.codItem) || "";
             const linhaCoage = [...d.linhaDados];
             linhaCoage.splice(5, 0, nomeAba); 
             linhaCoage[6] = obsSalva;
             linhaCoage.push(d.pdmCode, d.pdmSum, "", d.itemSum);
             listaDadosCOAGE.push(linhaCoage);
         });
      }

      if (aba.getFilter()) aba.getFilter().remove();
      aba.clear();
      aba.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]).setFontWeight("bold").setBackground("#1c4587").setFontColor("#ffffff");
         
      if (dadosParaEscrever.length > 0) _formatarAbaDestino(aba, dadosParaEscrever, cabecalho, false, metadadosLinha, nomeAba);
    });

    if (isFullUpdate && !isTest) {
      let abaCoage = ssDestino.getSheetByName("COAGE");
      if (!abaCoage) abaCoage = ssDestino.insertSheet("COAGE");
      if (abaCoage.getFilter()) abaCoage.getFilter().remove();

      abaCoage.clear();
      const cabecalhoCOAGE = [...cabecalho];
      cabecalhoCOAGE.splice(5, 0, "Planejador"); 
      abaCoage.getRange(1, 1, 1, cabecalhoCOAGE.length).setValues([cabecalhoCOAGE]).setFontWeight("bold").setBackground("#4c1130").setFontColor("#ffffff");
      if (listaDadosCOAGE.length > 0) _formatarAbaDestino(abaCoage, listaDadosCOAGE, cabecalhoCOAGE, true);
    }

    const msgSucesso = isTest ? "Aba de Teste criada/atualizada com sucesso!"
      : isFullUpdate ? "Painel de TODOS atualizado com sucesso!" : `Painéis atualizados: ${nomesAbas.join(", ")}`;
    ui.alert("Sucesso", msgSucesso, ui.ButtonSet.OK);
  } catch (e) { SpreadsheetApp.getUi().alert("Erro na Distribuição", e.message, SpreadsheetApp.getUi().ButtonSet.OK); }
}

function _formatarAbaDestino(aba, dados, headers, isCoage = false, metadados = [], nomeAba = "") {
    const numLinhas = dados.length;
    if (numLinhas === 0) return;

    aba.getRange(2, 1, numLinhas, headers.length).setValues(dados);

    const mapaFormatos = {
        "Estoque": "#,##0", "CMM Atual": "#,##0.00", "Saldo em Dias": "#,##0",
        "Qtd Empenho": "#,##0", "Quantidade AE": "#,##0", "Saldo Virtual + Saldo Físico": "#,##0",
        "Quantidade Original": "#,##0", "Saldo Ata": "#,##0", "Preço Unitário": "R$ #,##0.00",
        "Vencimento Ata": "dd/mm/yyyy", "Empenho": "@", "Saldo Empenho": "#,##0",
        "CMA Histórico (15 Meses)": "#,##0", "Previsão Esgotamento": "dd/mm/yyyy",
        "Sugestão Pedido (4 Meses)": "#,##0", "Prev. Esgotamento (Sugestão)": "dd/mm/yyyy",
        "Valor Grupo PDM": "R$ #,##0.00", "Soma Item": "R$ #,##0.00"
    };

    const formatosPorColuna = headers.map(h => mapaFormatos[h] || "General");
    const matrizFormatos = Array(numLinhas).fill(formatosPorColuna);
    aba.getRange(2, 1, numLinhas, headers.length).setNumberFormats(matrizFormatos);

    const matrizCores = Array(numLinhas).fill(null).map(() => Array(headers.length).fill(null));
    const matrizNotes = Array(numLinhas).fill(null).map(() => Array(1).fill(null));

    const idxStatus = headers.indexOf("Status Estoque");
    const idxObs = headers.indexOf("Observações (Equipe)");

    if (idxObs !== -1) {
        aba.getRange(2, idxObs + 1, numLinhas, 1).setHorizontalAlignment("center").setFontWeight("bold").setFontColor("#cc0000");
    }

    for (let i = 0; i < numLinhas; i++) {
        if (idxStatus !== -1) {
            const st = dados[i][idxStatus];
            let corSt = null;
            
            // RECUPERADO: Uso das configurações de cores originais do seu painel
            if (st === "Crítico") corSt = (typeof CONFIG !== 'undefined' && CONFIG.cores && CONFIG.cores.ALERTA_CRITICO) ? CONFIG.cores.ALERTA_CRITICO : "#ff0000";
            else if (st === "Atenção (60d)") corSt = "#ff9900"; 
            else if (st === "Atenção") corSt = (typeof CONFIG !== 'undefined' && CONFIG.cores && CONFIG.cores.ALERTA_ATENCAO) ? CONFIG.cores.ALERTA_ATENCAO : "#ffff00";
            else if (st === "Suprimento Ok") corSt = (typeof CONFIG !== 'undefined' && CONFIG.cores && CONFIG.cores.ALERTA_OK) ? CONFIG.cores.ALERTA_OK : "#00ff00";
            
            matrizCores[i][idxStatus] = corSt;
        }

        if (!isCoage && metadados[i]) {
            let corLinha = null;
            if (nomeAba === "Luciana" && metadados[i].isGeladeira) corLinha = '#cfe2f3';
            else if (metadados[i].conflito) corLinha = '#ea9999';

            if (corLinha) {
                for (let j = 0; j < headers.length; j++) {
                    if (j !== idxStatus || !matrizCores[i][j]) {
                        matrizCores[i][j] = corLinha;
                    }
                }
            }
            if (metadados[i].conflito && metadados[i].nota) matrizNotes[i][0] = metadados[i].nota;
        }
    }

    aba.getRange(2, 1, numLinhas, headers.length).setBackgrounds(matrizCores);
    if (!isCoage && metadados.length > 0) aba.getRange(2, 2, numLinhas, 1).setNotes(matrizNotes);

    // RECUPERADO: O seu sistema matemático de cálculo de colunas 
    for (let c = 0; c < headers.length; c++) {
       let largura = (headers[c].length * 10) + 35;
       if (largura < 70) largura = 70;
       aba.setColumnWidth(c + 1, largura);
    }
    
    if (aba.getFilter()) aba.getFilter().remove();
    aba.getDataRange().createFilter();
}

function _obterAnoSeguro(valor) {
  if (!valor) return null;
  if (valor instanceof Date) return valor.getFullYear();
  const str = String(valor).trim();
  const partes = str.split('/');
  if (partes.length === 3) {
    const ano = parseInt(partes[2]);
    if (!isNaN(ano) && ano > 2000 && ano < 2100) return ano;
  }
  return null;
}

function carregarRegrasDistribuicao(ssFonte) {
  const nomeAbaConfig = CONFIG.abas.configEquipe || "Config_Equipe";
  let abaConfig = ssFonte.getSheetByName(nomeAbaConfig);
  if (!abaConfig) return {};
  
  const lastRow = abaConfig.getLastRow();
  if (lastRow < 2) return {};
  
  // RECUPERADO: Voltou a usar getDisplayValues() para ler as regras exatamente como na tela
  const dados = abaConfig.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
  const regrasDinamicas = {};
  dados.forEach(linha => {
    const chave = String(linha[0]).replace(',', '.').trim().toUpperCase(); 
    const responsavel = String(linha[1]).trim();
    if (chave && responsavel) regrasDinamicas[chave] = responsavel;
  });
  return regrasDinamicas;
}

function _parseDataSegura(valorData) {
  if (!valorData) return null;
  if (valorData instanceof Date) return valorData;
  const str = String(valorData).trim();
  const ps = str.split("/");
  if (ps.length === 3) {
      const d = parseInt(ps[0], 10), m = parseInt(ps[1], 10), y = parseInt(ps[2], 10);
      if (d > 0 && m > 0 && y > 2000) return new Date(y, m - 1, d);
  }
  return null;
}

function _norm(t) { return t ? String(t).trim().toUpperCase() : ""; }

function atualizarTodos() { distribuirDadosPorEquipe(null); }
function atualizarLorena() { distribuirDadosPorEquipe("Lorena"); }
function atualizarKatia() { distribuirDadosPorEquipe("Katia"); }
function atualizarLeonardo() { distribuirDadosPorEquipe("Leonardo"); }
function atualizarMoises() { distribuirDadosPorEquipe("Moises"); }
function atualizarRafaelle() { distribuirDadosPorEquipe("Rafaelle"); }
function atualizarLuciana() { distribuirDadosPorEquipe("Luciana"); }