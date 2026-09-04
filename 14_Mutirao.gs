// =================================================================
// --- BLOCO 14: STATUS REPORT (ATUALIZADO COM FLUXO E PRIORIDADES E CMM) ---
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
    
    const dadosUrg = abaUrg.getRange(2, 1, lastRowUrg - 1, 9).getValues();
    const colA = []; 
    const colPrio = [];
    const colPrazo = [];
    
    dadosUrg.forEach(linha => {
       const cod = linha[0];
       if (cod !== "" && cod != null) {
          colA.push([cod]);
          
          let valPrio = linha[7];
          let valPrazo = linha[8];
          valPrio = (valPrio === 0 || valPrio === "0") ? "0" : (valPrio ? String(valPrio).trim() : "");
          valPrazo = (valPrazo === 0 || valPrazo === "0") ? "0" : (valPrazo ? String(valPrazo).trim() : "");
          
          colPrio.push([valPrio]);
          colPrazo.push([valPrazo]);
        }
    });
    if (colA.length === 0) {
      throw new Error("Nenhum código válido encontrado na coluna A da aba de Urgências.");
    }
    
    if (abaStatus.getMaxColumns() < 14) {
      abaStatus.insertColumnsAfter(abaStatus.getMaxColumns(), 14 - abaStatus.getMaxColumns());
    }
    const maxRows = abaStatus.getMaxRows();
    if(maxRows > 1) {
        abaStatus.getRange(2, 1, maxRows - 1, 14).clearContent();
    }
    
    abaStatus.getRange(2, 13, colPrio.length, 2).setNumberFormat("@");
    abaStatus.getRange(2, 1, colA.length, 1).setValues(colA);
    abaStatus.getRange(2, 13, colPrio.length, 1).setValues(colPrio);
    abaStatus.getRange(2, 14, colPrazo.length, 1).setValues(colPrazo);
    
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
        <div id="msg" class="loading">⏳ Buscando dados na Nuvem (Isso pode levar alguns segundos)...</div>
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

    const mapObservacoesSalvas = new Map();
    const mapQtdSolicitada = new Map();
    const dadosAtuais = abaMutirao.getRange(2, 1, lastRowMult - 1, 14).getValues();
    const setCodigos = new Set();
    
    dadosAtuais.forEach(linha => {
      const cod = _norm(linha[0]);
      const qtd = linha[4]; 
      const obs = String(linha[11]).trim(); 
        
      if (cod) {
        setCodigos.add(cod);
        if (obs) mapObservacoesSalvas.set(cod, obs);
        if (qtd !== "" && qtd != null) mapQtdSolicitada.set(cod, qtd);
      }
    });

    if (setCodigos.size === 0) { ui.alert("Nenhum código válido na Coluna A."); return; }

    const mapaEntradas = _buscarDadosEntradaEmpenhos(setCodigos);
    const mapaEmpenhos = _buscarEmpenhosCompilados(setCodigos, ssFonte);

    const lastRowDados = abaDados.getLastRow();
    const mapaDados = new Map();
    
    const linhaInicialDados = 3; 
    if (lastRowDados >= linhaInicialDados) {
      // AJUSTE: Leitura da nova estrutura com 43 colunas (+3 colunas novas adicionadas)
      const vDados = abaDados.getRange(linhaInicialDados, 1, lastRowDados - (linhaInicialDados-1), 43).getValues();
      
      for (let i = 0; i < vDados.length; i++) {
        const linha = vDados[i];
        const cod = _norm(linha[1]); // Col B
        
        if (cod && setCodigos.has(cod)) {
           if (!mapaDados.has(cod)) {
             mapaDados.set(cod, {
                estoqueTotal: 0,
                aeSet: new Set(),
                notesSet: new Set(),
                procAtaSet: new Set(),
                procAnaliseSet: new Set(),
                descricaoSet: new Set(),
                cmm: 0
             });
           }
           const d = mapaDados.get(cod);
           
           const estoqueLinha = parseFloat(linha[7]) || 0; // Col H (SaldoAtual)
           d.estoqueTotal = Math.max(d.estoqueTotal, estoqueLinha);

           const cmmLinha = parseFloat(linha[8]) || 0; // Col I (Cmm12)
           d.cmm = Math.max(d.cmm, cmmLinha);

           if (linha[2]) d.descricaoSet.add(String(linha[2]).trim()); // Col C (Descricao)
           
           // === EXTRAÇÃO DAS COLUNAS RELEVANTES (USANDO OS ÍNDICES DA MATRIZ) ===
           // Ajustados para acomodar o pulo de 3 colunas extras adicionadas a planilha
           const valorSolicitacao = linha[22] !== undefined ? String(linha[22]).trim() : ""; // Deslocado de 19 para 22
           const valorProcAta = linha[24] !== undefined ? String(linha[24]).trim() : "";      // Deslocado de 21 para 24
           const valorProcAnalise = linha[30] !== undefined ? String(linha[30]).trim() : "";  // Deslocado de 27 para 30
           const valorColunaAE = linha[33] !== undefined ? String(linha[33]).trim() : "";     // Deslocado de 30 para 33

           if (valorSolicitacao) {
             const primeiroChar = String(valorSolicitacao).charAt(0);
             if (/\d/.test(primeiroChar)) {
               const digito = parseInt(primeiroChar, 10);
               if (!isNaN(digito) && digito >= 6) {
                 d.notesSet.add(valorSolicitacao);
               }
             }
           }

           if (valorColunaAE) {
             d.aeSet.add(valorColunaAE);
           }

           if (valorProcAta) d.procAtaSet.add(valorProcAta);
           if (valorProcAnalise) d.procAnaliseSet.add(valorProcAnalise);
        }
      }
    }

    const outputDados = [];
    const outputDescricoes = [];
    const outputPrecos = [];
    const outputCmm = [];
    const outputQtdSolicitada = [];
    const outputFormulasTotal = [];
    const outputObservacoes = [];
    
    const listaCompletaEmail = [];

    dadosAtuais.forEach((linhaInput, index) => {
      const codAtual = _norm(linhaInput[0]);
      if (!codAtual) { 
        outputDescricoes.push([""]); outputPrecos.push([""]); outputCmm.push([""]); outputQtdSolicitada.push([""]);
        outputFormulasTotal.push([""]); outputDados.push(["", "", "", "", ""]); outputObservacoes.push([""]);
        return; 
      }
      
      const prioridade = linhaInput[12] != null ? String(linhaInput[12]).trim() : ""; 
      const prazo = linhaInput[13] != null ? String(linhaInput[13]).trim() : ""; 

      const infoEstoque = mapaDados.get(codAtual) || {
        estoqueTotal: 0, aeSet: new Set(), notesSet: new Set(), procAtaSet: new Set(), procAnaliseSet: new Set(), descricaoSet: new Set(), cmm: 0
      };
      
      const infoEntradas = mapaEntradas.get(codAtual) || {
        valor: 0, planejadores: new Set(), descricao: ""
      };
      
      let descricaoFinal = Array.from(infoEstoque.descricaoSet)[0] || infoEntradas.descricao || "Descrição não encontrada";
      const precoUnit = infoEntradas.valor;
      const cmmItem = infoEstoque.cmm;
      
      const qtdSol = mapQtdSolicitada.has(codAtual) ? mapQtdSolicitada.get(codAtual) : "";
      
      const linhaPlanilha = index + 2;
      
      const formulaTotal = `=IF(ISNUMBER(E${linhaPlanilha}); D${linhaPlanilha}*E${linhaPlanilha}; 0)`;

      const txtNotesArray = Array.from(infoEstoque.notesSet);
      const txtAEArray = Array.from(infoEstoque.aeSet);
      const txtNotesFinal = txtNotesArray.length ? txtNotesArray.join("\n") : "";
      const txtAEFinal = txtAEArray.length ? txtAEArray.join("\n") : "";

      let blocoAENotes = "";
      if (txtNotesFinal) blocoAENotes += `Notes ata: ${txtNotesFinal}`;
      if (txtAEFinal) blocoAENotes += (blocoAENotes ? "\n" : "") + `AE: ${txtAEFinal}`;

      const procAtaTxt = Array.from(infoEstoque.procAtaSet).join("\n");
      const procAnaliseTxt = Array.from(infoEstoque.procAnaliseSet).join("\n");
      let blocoProcessos = "";
      if (procAtaTxt) blocoProcessos += `Processo Ata: ${procAtaTxt}`;
      if (procAnaliseTxt) blocoProcessos += (blocoProcessos ? "\n" : "") + `Processo em Andamento: ${procAnaliseTxt}`;

      const txtPlan = infoEntradas.planejadores.size > 0 ? Array.from(infoEntradas.planejadores).join(" / ") : "Não Encontrado";
      const txtEmp = Array.from(mapaEmpenhos.get(codAtual) || new Set()).join("\n");
      const estoqueTotal = infoEstoque.estoqueTotal;

      outputDescricoes.push([descricaoFinal]); 
      outputPrecos.push([precoUnit]);
      outputCmm.push([cmmItem]);
      outputQtdSolicitada.push([qtdSol]);
      outputFormulasTotal.push([formulaTotal]);
      
      outputDados.push([ blocoAENotes, txtEmp, estoqueTotal, blocoProcessos, txtPlan ]); 
      outputObservacoes.push([ mapObservacoesSalvas.get(codAtual) || "" ]);

      listaCompletaEmail.push({
        codigo: codAtual, descricao: descricaoFinal, preco: precoUnit, cmm: cmmItem, qtdSol: qtdSol,
        estoque: estoqueTotal, empenho: txtEmp, aeNotes: blocoAENotes,
        processos: blocoProcessos, planejador: txtPlan, observacoes: mapObservacoesSalvas.get(codAtual) || "",
        prioridade: prioridade, prazo: prazo
      });
    });

    if (outputDados.length > 0) {
      const numLinhas = outputDados.length;
      abaMutirao.getRange(2, 2, abaMutirao.getMaxRows() - 1, 11).clearContent();
      abaMutirao.getRange(2, 2, numLinhas, 1).setValues(outputDescricoes);
      abaMutirao.getRange(2, 3, numLinhas, 1).setValues(outputPrecos).setNumberFormat("R$ #,##0.00");
      abaMutirao.getRange(2, 4, numLinhas, 1).setValues(outputCmm).setNumberFormat("#,##0.##");
      abaMutirao.getRange(2, 5, numLinhas, 1).setValues(outputQtdSolicitada).setNumberFormat("#,##0");
      abaMutirao.getRange(2, 6, numLinhas, 1).setFormulas(outputFormulasTotal).setNumberFormat("R$ #,##0.00").setFontWeight("bold");
      
      abaMutirao.getRange(2, 7, numLinhas, 5).setValues(outputDados);
      abaMutirao.getRange(2, 7, numLinhas, 5).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      abaMutirao.getRange(2, 7, numLinhas, 5).setVerticalAlignment("middle");
      
      abaMutirao.getRange(2, 9, numLinhas, 1).setNumberFormat("#,##0");
      abaMutirao.getRange(2, 12, numLinhas, 1).setValues(outputObservacoes);
      
      const headers = ["Código", "Descrição", "Valor Unitário", "CMM", "Qtd. Solicitada", "Valor Total", "AE / Notes", "Empenho", "Estoque", "Processos", "Planejador", "Observações", "Prioridade", "Prazo"];
      abaMutirao.getRange(1, 1, 1, 14).setValues([headers])
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

function _buscarEmpenhosCompilados(setCodigos, ssFonte) {
  const mapa = new Map();
  try {
    const nomeAbaCompilados = CONFIG.abas.compilados || "Compilados";
    let aba = ssFonte.getSheetByName(nomeAbaCompilados);
    if (!aba) aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeAbaCompilados);
    if (!aba) return mapa;

    const lastRow = aba.getLastRow();
    if (lastRow < 2) return mapa;

    const dados = aba.getRange(2, 1, lastRow - 1, 19).getValues();

    dados.forEach(linha => {
      const emp = String(linha[0]).trim();        
      const cod = _norm(linha[5]);                
      const status = String(linha[18]).trim().toUpperCase(); 

      if (cod && setCodigos.has(cod) && status.includes("PENDENTE")) {
        if (!mapa.has(cod)) mapa.set(cod, new Set());
        if (emp) mapa.get(cod).add(emp);
      }
    });
  } catch(e) { console.error("Erro na busca de empenhos Pendentes: " + e.message); }
  return mapa;
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
      const plan = String(linha[14]).trim();   
      let valor = linha[16];                   
                   
      if (cod && setCodigos.has(cod)) {
        if (typeof valor === 'string') valor = parseFloat(valor.replace("R$", "").replace(".", "").replace(",", ".").trim());
        else valor = parseFloat(valor) || 0;

        if (!mapa.has(cod)) {
          mapa.set(cod, { valor: 0, planejadores: new Set(), descricao: desc });
        }
        
        const obj = mapa.get(cod);
        
        if (plan) obj.planejadores.add(plan);
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

  const assunto = `[INFORMATIVO: ${nomeContexto} - Urgências e Prazos (${new Date().toLocaleDateString()})]`;
  
  const blocos = {
      "0": { titulo: "[PRIORIDADE: IMEDIATA]", cor: "#CC0000", itens: [] },
      "1": { titulo: "[PRIORIDADE: 3 a 5 DIAS]", cor: "#E69138", itens: [] },
      "2": { titulo: "[PRIORIDADE: 7 A 10 DIAS]", cor: "#F1C232", itens: [] },
      "OUTROS": { titulo: "[SEM PRIORIDADE DEFINIDA]", cor: "#666666", itens: [] }
  };

  listaItens.forEach(item => {
      const p = String(item.prioridade).trim().toLowerCase();
      if (p === "0" || p === "imediata" || p === "imediato") blocos["0"].itens.push(item);
      else if (p === "1") blocos["1"].itens.push(item);
      else if (p === "2") blocos["2"].itens.push(item);
      else blocos["OUTROS"].itens.push(item);
  });

  let htmlEmail = `<div style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
      <h2 style="color: ${corTitulo}; border-bottom: 2px solid ${corTitulo}; padding-bottom: 5px;">
        ${nomeContexto} - Relatório de Urgências
      </h2>
      <p>Abaixo estão os códigos categorizados por seu respectivo nível de criticidade e prazo de resolução.</p>
      <p>Os relatórios completos detalhados encontram-se em anexo.</p><br>`;

  let htmlPdf = `<html><head><style>
       @page { size: landscape; margin: 10mm; } 
       body { font-family: Arial, sans-serif; font-size: 10px; } 
       h2 { background: ${corTitulo}; color: white; padding: 6px; text-align: center; margin-bottom: 10px; border-radius: 4px; } 
       .bloco-titulo { color: white; padding: 5px; margin-top: 20px; border-radius: 3px; font-size: 12px; font-weight: bold; }
       table { width: 100%; border-collapse: collapse; margin-bottom: 15px; } 
       th, td { border: 1px solid #ccc; padding: 4px; vertical-align: top; } 
       th { background: #f2f2f2; }
  </style></head><body>
  <h2>Relatório Detalhado: ${nomeContexto} (${new Date().toLocaleDateString()})</h2>`;

  const pdfHeader = `<thead><tr>
    <th width="6%">Cód.</th>
    <th width="15%">Descrição</th>
    <th width="6%">Unitário</th>
    <th width="4%">CMM</th>
    <th width="4%">Qtd</th>
    <th width="5%">Total</th>
    <th width="10%">AE/Notes</th>
    <th width="10%">Empenho</th>
    <th width="4%">Estoque</th>
    <th width="10%">Processos</th>
    <th width="14%">Planejador</th>
    <th width="12%">Prazo</th>
  </tr></thead>`;

  const exHeaders = ["Código", "Descrição", "Valor Unitário", "CMM", "Qtd. Solicitada", "Valor Total", "AE / Notes", "Empenho", "Estoque", "Processos", "Planejador", "Observações", "Prioridade", "Prazo"];
  const dadosExcel = [exHeaders];
  const linhasBlocoExcel = [];

  Object.keys(blocos).forEach(chave => {
      const bloco = blocos[chave];
      if (bloco.itens.length > 0) {
          htmlEmail += `<h3 style="background-color: ${bloco.cor}; color: white; padding: 8px; border-radius: 4px; margin-top: 25px; margin-bottom: 5px;">
                        ${bloco.titulo}
                        </h3>
                        <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
                          <thead>
                            <tr style="background-color: #f2f2f2;">
                              <th style="border: 1px solid #ddd; padding: 6px; width: 15%;">Código</th>
                              <th style="border: 1px solid #ddd; padding: 6px; width: 65%;">Descrição</th>
                              <th style="border: 1px solid #ddd; padding: 6px; width: 20%;">Prazo</th>
                            </tr>
                          </thead><tbody>`;
                          
          htmlPdf += `<div class="bloco-titulo" style="background-color: ${bloco.cor};">${bloco.titulo}</div>
                      <table>${pdfHeader}<tbody>`;
                      
          linhasBlocoExcel.push(dadosExcel.length);
          dadosExcel.push([bloco.titulo, "", "", "", "", "", "", "", "", "", "", "", "", ""]);

          bloco.itens.forEach(item => {
              const precoFmt = item.preco ? item.preco.toLocaleString('pt-BR', {style: 'currency', currency: 'BRL'}) : '-';
              const prazoTxt = item.prazo || '-';
              const qtd = parseFloat(item.qtdSol) || 0;
              const cmm = parseFloat(item.cmm) || 0;
              const valorTotalNum = qtd * cmm;
              const valorTotalFmt = valorTotalNum > 0 ? valorTotalNum.toLocaleString('pt-BR', {style: 'currency', currency: 'BRL'}) : '-';
              
              htmlEmail += `<tr>
                <td style="border: 1px solid #ddd; padding: 6px; text-align: center;"><strong>${item.codigo}</strong></td>
                <td style="border: 1px solid #ddd; padding: 6px;">${item.descricao.substring(0, 90)}...</td>
                <td style="border: 1px solid #ddd; padding: 6px; text-align: center;">${prazoTxt}</td>
              </tr>`;
              
              htmlPdf += `<tr>
                <td><strong>${item.codigo}</strong></td>
                <td>${item.descricao.substring(0, 85)}</td>
                <td align="right">${precoFmt}</td>
                <td align="center">${item.cmm || '-'}</td>
                <td align="center">${item.qtdSol || '-'}</td>
                <td align="right">${valorTotalFmt}</td>
                <td>${item.aeNotes.replace(/\n/g, "<br>")}</td>
                <td>${item.empenho.replace(/\n/g, "<br>")}</td>
                <td align="center">${item.estoque}</td>
                <td>${item.processos.replace(/\n/g, "<br>")}</td>
                <td>${item.planejador}</td>
                <td align="center"><strong>${prazoTxt}</strong></td>
              </tr>`;
              
              dadosExcel.push([
                item.codigo, item.descricao, item.preco, item.cmm, item.qtdSol, valorTotalNum, 
                item.aeNotes, item.empenho, item.estoque, item.processos, item.planejador,
                item.observações, bloco.titulo, prazoTxt
              ]);
          });
          htmlEmail += `</tbody></table>`;
          htmlPdf += `</tbody></table>`;
      }
  });

  htmlEmail += `<br><p>Atenciosamente,<br><strong>Orquestrador de Estoque INCA</strong></p></div>`;
  htmlPdf += `</body></html>`;

  const pdfBlob = HtmlService.createHtmlOutput(htmlPdf).getAs(MimeType.PDF).setName(`Relatorio_Prioridades_${nomeContexto}.pdf`);
  
  let excelBlob = null;
  try {
    const tempSS = SpreadsheetApp.create("Temp_Export");
    const sheet = tempSS.getSheets()[0];
    const rangeTemp = sheet.getRange(1, 1, dadosExcel.length, exHeaders.length);
    rangeTemp.setValues(dadosExcel);
    
    sheet.getRange(1, 1, 1, exHeaders.length).setBackground("#d9ead3").setFontWeight("bold");
    linhasBlocoExcel.forEach(idx => {
       sheet.getRange(idx + 1, 1, 1, exHeaders.length).setBackground("#444444").setFontColor("white").setFontWeight("bold");
    });
    
    SpreadsheetApp.flush();
    const url = "https://docs.google.com/spreadsheets/d/" + tempSS.getId() + "/export?format=xlsx";
    excelBlob = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob().setName(`Relatorio_Prioridades_${nomeContexto}.xlsx`);
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