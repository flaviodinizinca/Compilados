// =================================================================
// --- BLOCO 11: SINCRONIZAÇÃO EXTERNA (ATUALIZADO: FONTE GERAL + DEDUPLICAÇÃO) ---
// =================================================================

/**
 * Busca divergências entre a Fonte Geral (DadosGlobais) e a Planilha Externa de Empenhos.
 * Verifica itens que estão em DadosGlobais mas faltam na Externa (Empenhos Enviados), e vice-versa.
 */
function buscarEmpenhosCodigosErrados() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // 1. Conecta à Fonte Geral e Busca Dados
    if (!CONFIG.ids.fonteDadosGeral) throw new Error("ID_FONTE_GERAL não configurado.");
    const ssFonteGeral = SpreadsheetApp.openById(CONFIG.ids.fonteDadosGeral);
    const dadosGlobais = _fetchDadosGlobais(ssFonteGeral);

    // 2. Prepara aba local de Relatório
    let abaCorrigir = ss.getSheetByName("Empenhos a Corrigir");
    if (!abaCorrigir) {
      abaCorrigir = ss.insertSheet("Empenhos a Corrigir");
    } else {
      abaCorrigir.clear();
    }

    // 3. Lê dados da Planilha Externa
    if (!CONFIG.ids.correcaoExterna) throw new Error("ID_CORRECAO_EXTERNA não configurado.");
    const ssExterna = SpreadsheetApp.openById(CONFIG.ids.correcaoExterna);
    const abaExterna = ssExterna.getSheetByName("Empenhos");
    
    const lastRowExt = abaExterna.getLastRow();
    const mapExt = new Set(); 
    const listExt = [];

    if(lastRowExt > 1) {
       // Lê colunas A a F da externa
       abaExterna.getRange(2,1,lastRowExt-1,6).getValues().forEach(r => {
          const e = _norm(r[3]); // Coluna D: Empenho
          const c = _norm(r[4]); // Coluna E: Código
          if(e) { 
            mapExt.add(e); 
            if(c) listExt.push({e,c}); 
          }
       });
    }
    
    // 4. Processa Dados da Fonte Geral
    const mapEnt = new Map(); 
    const mapCodEmp = new Map();
    
    dadosGlobais.forEach(r => {
       // Estrutura esperada do _fetchDadosGlobais: [Empenho, Data, Código, ..., Status, ..., Item]
       const e = _norm(r[0]); 
       const c = _norm(r[2]);
       const status = r[13] ? _norm(r[13]) : "";

       if(e && c && status !== 'ELIMINADA' && status !== 'ANULADO') {
          if(!mapEnt.has(e)) mapEnt.set(e, new Set()); 
          mapEnt.get(e).add(c);
          
          if(!mapCodEmp.has(c)) mapCodEmp.set(c, new Set()); 
          mapCodEmp.get(c).add(e);
       }
    });
    
    // 5. Gera Relatório de Divergências
    const rel = []; 
    const chk = new Set();

    // A) Itens que estão na Geral mas FALTAM na Externa
    mapEnt.forEach((v,k) => {
       if(mapExt.has(k)) { // Só verifica se o empenho já existe na externa (para não pedir cadastro de empenhos novos inteiros aqui, talvez?)
          v.forEach(c => {
             // Verifica se o par Empenho/Codigo existe na lista externa
             if(!listExt.some(x => x.e === k && x.c === c) && !chk.has(`${k}|${c}`)) {
                rel.push([k, c, "-", "FALTANTE", "Item no EMS mas falta na Planilha de Empenhos Enviados"]);
                chk.add(`${k}|${c}`);
             }
          });
       }
    });

    // B) Itens na Externa que NÃO batem com a Geral (Código Errado ou Empenho Errado)
    listExt.forEach(x => {
       const v = mapEnt.get(x.e);
       if(v && !v.has(x.c)) {
          const assoc = mapCodEmp.has(x.c) ? Array.from(mapCodEmp.get(x.c)).join(", ") : "Cod Inexistente";
          rel.push([x.e, x.c, assoc, "CÓDIGO ERRADO", "Cod dos enviados não bate com EMS"]);
       }
    });

    // 6. Saída
    if(rel.length > 0) {
       abaCorrigir.getRange(1,1,1,5).setValues([["Empenho", "Código", "Empenho", "Erro", "Descrição"]]);
       abaCorrigir.getRange(2,1,rel.length,5).setValues(rel);
       abaCorrigir.getRange(2,1,rel.length,1).setNumberFormat("@");
       abaCorrigir.getRange(1,1,1,5).setFontWeight("bold").setBackground("#fce5cd");

       const resposta = ui.alert(
         "Divergências Encontradas", 
         `Foram encontrados ${rel.length} erros.\nDeseja enviar este relatório por e-mail para a equipe agora?`, 
         ui.ButtonSet.YES_NO
       );

       if (resposta == ui.Button.YES) {
         if (typeof enviarEmailDivergencias === 'function') {
            enviarEmailDivergencias(rel);
            ui.alert("E-mail enviado com sucesso!");
         } else {
            ui.alert("Função de envio de email não encontrada.");
         }
       }

    } else {
       ui.alert("Análise concluída: Tudo OK! Nenhuma divergência encontrada.");
    }

  } catch(e) { 
    ui.alert("Erro na análise: " + e.message);
    console.error(e);
  }
}

/**
 * Sincroniza os itens marcados como FALTANTE na planilha externa.
 * Também remove duplicatas da planilha externa para garantir limpeza.
 */
function sincronizarEmpenhosNaExterna() {
   const ui = SpreadsheetApp.getUi();
   const ss = SpreadsheetApp.getActiveSpreadsheet();

   try {
     const abaCorr = ss.getSheetByName("Empenhos a Corrigir");
     if(!abaCorr || abaCorr.getLastRow()<2) return ui.alert("Nada a corrigir (Aba 'Empenhos a Corrigir' vazia).");
     
     // 1. Busca Dados Frescos da Fonte Geral
     const ssFonteGeral = SpreadsheetApp.openById(CONFIG.ids.fonteDadosGeral);
     const dadosEnt = _fetchDadosGlobais(ssFonteGeral);

     const list = abaCorr.getRange(2,1,abaCorr.getLastRow()-1,4).getValues();
     
     // Mapeia dados da Fonte Geral para preenchimento: Chave = Empenho|Codigo
     const mapEnt = new Map();
     dadosEnt.forEach(r => {
        // r[0]=Emp, r[1]=Data, r[2]=Cod, r[6]=IG(aprox), r[19]=Item(aprox)
        // Ajuste conforme layout real da EntradaEmpenhos. 
        // Assumindo r[6]=IG e r[19]=DescricaoItem baseada no script anterior
        if(r[0]&&r[2]) mapEnt.set(`${_norm(r[0])}|${_norm(r[2])}`, {d:r[1], ig:r[6], it:r[19]});
     });

     const ssExt = SpreadsheetApp.openById(CONFIG.ids.correcaoExterna);
     const abaExt = ssExt.getSheetByName("Empenhos");
     
     // 2. Passo Prévio: Remover Duplicatas Existentes na Externa
     // Isso resolve o problema de "manter duplicatas"
     _removerDuplicatasExterna(abaExt);

     // 3. Adicionar Novos Itens
     const novos = [];
     const adicionadosAgora = new Set(); // Evita duplicar na própria inserção

     list.forEach(r => {
        if(r[3]==='FALTANTE') {
           const k = `${_norm(r[0])}|${_norm(r[1])}`; // Empenho (r[0]) | Código (r[1] na aba corrigir)
           
           if(mapEnt.has(k) && !adicionadosAgora.has(k)) {
              const d = mapEnt.get(k);
              // Fórmula para Status (Coluna F)
              const rowNum = abaExt.getLastRow() + novos.length + 2; // +2 offset aproximado
              const form = `=IF(A${rowNum}="";"";IFERROR(VLOOKUP(E${rowNum};ITENS!A:B;2;0);"Cadastrar"))`;
              
              // Ordem Colunas Externa: [Data, IG, Item, Empenho, Código, Status]
              novos.push([d.d, d.ig, d.it, r[0], r[1], form]);
              adicionadosAgora.add(k);
           }
        }
     });

     if(novos.length) {
        abaExt.getRange(abaExt.getLastRow()+1, 1, novos.length, 6).setValues(novos);
        abaCorr.clearContents();
        ui.alert(`Sincronização concluída!\n\nLimpeza de duplicatas executada.\n${novos.length} novos itens adicionados.`);
     } else {
        ui.alert("Limpeza executada, mas nenhum item 'FALTANTE' válido encontrado para adição.");
     }

   } catch(e) { 
       ui.alert("Erro na sincronização: " + e.message);
       console.error(e);
   }
}

/**
 * Função Auxiliar: Remove linhas duplicadas na planilha externa, mantendo a primeira ocorrência.
 * Baseado na chave composta: Empenho (Col D) + Código (Col E).
 */
function _removerDuplicatasExterna(aba) {
  const lastRow = aba.getLastRow();
  if (lastRow < 2) return;

  const dados = aba.getRange(2, 1, lastRow - 1, 6).getValues();
  const linhasParaManter = [];
  const chavesVistas = new Set();
  const linhasDuplicadasIndices = []; // Para log ou exclusão em massa (complexo no AppScript direto)

  // Abordagem segura: Ler tudo, filtrar em memória, reescrever se houver duplicatas.
  // Como deletar linhas uma a uma é lento, reescrever o range é mais rápido se não houver fórmulas complexas em outras colunas.
  // Assumindo que os dados são estáticos nas colunas A-E e fórmula na F.
  
  let temDuplicata = false;

  dados.forEach((linha) => {
    const emp = _norm(linha[3]);
    const cod = _norm(linha[4]);
    const chave = `${emp}|${cod}`;

    if (emp && cod) {
        if (chavesVistas.has(chave)) {
            temDuplicata = true;
            // É duplicata, não adiciona ao array final
        } else {
            chavesVistas.add(chave);
            linhasParaManter.push(linha);
        }
    } else {
        // Linhas vazias ou incompletas, mantemos? Geralmente sim, ou descartamos. 
        // Vamos manter se tiver algum conteúdo relevante.
        if(linha.join("").length > 0) linhasParaManter.push(linha);
    }
  });

  if (temDuplicata) {
      // Limpa e Reescreve
      aba.getRange(2, 1, lastRow - 1, 6).clearContent();
      if (linhasParaManter.length > 0) {
          aba.getRange(2, 1, linhasParaManter.length, 6).setValues(linhasParaManter);
      }
      console.log("Duplicatas removidas da planilha externa.");
  }
}

/**
 * Função Helper para buscar dados unificados da ID_FONTE_GERAL
 */
function _fetchDadosGlobais(ss) {
  const dados = [];
  
  // 1. EntradaEmpenhos
  const abaEnt = ss.getSheetByName(CONFIG.abas.entradaEmpenhos || "EntradaEmpenhos");
  if (abaEnt && abaEnt.getLastRow() > 1) {
     // Assume colunas padrão do EMS/Entradas:
     // A(0)=Empenho, B(1)=Processo, C(2)=Cod, G(6)=Data, N(13)=Status, T(19)=Descricao
     // Ajuste o range conforme a largura real da sua planilha
     abaEnt.getRange(2, 1, abaEnt.getLastRow() - 1, 25).getValues().forEach(r => dados.push(r));
  }

  // 2. OutrasEntradas
  const abaOut = ss.getSheetByName(CONFIG.abas.outrasEntradas || "OutrasEntradas");
  if (abaOut && abaOut.getLastRow() > 1) {
      // Mapeamento de Outras Entradas para estrutura comum se necessário
      // Geralmente OutrasEntradas tem estrutura diferente. 
      // O script original tratava tudo junto. Se a estrutura for compatível (ex: Empenho na col A), ok.
      // Caso contrário, precisaria de um adaptador. 
      // Por segurança, vou carregar, mas o filtro principal é pelo status/código.
      abaOut.getRange(2, 1, abaOut.getLastRow() - 1, 20).getValues().forEach(r => dados.push(r));
  }

  return dados;
}

/**
 * Repara informações faltantes (Data, IG, Descrição) na planilha externa
 * usando os dados da Fonte Geral.
 */
function repararDadosFaltantesNaExterna() {
   const ui = SpreadsheetApp.getUi();
   try {
      const ssFonteGeral = SpreadsheetApp.openById(CONFIG.ids.fonteDadosGeral);
      const dadosEnt = _fetchDadosGlobais(ssFonteGeral);
      
      const map = new Map();
      dadosEnt.forEach(r => { 
          if(r[0]) map.set(_norm(r[0]), {d:r[1], g:r[6], t:r[19]}); 
      });

      const ssExt = SpreadsheetApp.openById(CONFIG.ids.correcaoExterna);
      const abaExt = ssExt.getSheetByName("Empenhos");
      const lastRow = abaExt.getLastRow();
      if(lastRow < 2) return;

      const range = abaExt.getRange("A2:F"+lastRow); // Aumentei para F para cobrir colunas usadas
      const data = range.getValues();
      
      let cnt = 0;
      data.forEach((r, i) => {
         const e = _norm(r[3]); // Empenho (Col D)
         if(e && map.has(e)) {
            const inf = map.get(e);
            let m = false;
            // Se Data (Col A) vazia
            if(!r[0] && inf.d){ r[0]=inf.d; m=true; }
            // Se IG (Col B) vazia
            if(!r[1] && inf.g){ r[1]=inf.g; m=true; }
            // Se Descrição (Col C) vazia
            if(!r[2] && inf.t){ r[2]=inf.t; m=true; }
            
            if(m) {
               // Atualiza apenas colunas A, B, C
               abaExt.getRange(i+2, 1, 1, 3).setValues([[r[0], r[1], r[2]]]);
               cnt++;
            }
         }
      });
      ui.alert(`${cnt} linhas reparadas com informações da Fonte Geral.`);
   } catch(e) { ui.alert("Erro ao reparar: " + e.message); }
}

// Pequeno Helper local caso _norm não esteja global
function _norm(val) {
  return val ? String(val).trim().toUpperCase() : "";
}