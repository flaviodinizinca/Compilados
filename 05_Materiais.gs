// =================================================================
// --- BLOCO 5: PROCESSAMENTO REMOTO - MATERIAIS (LÓGICA HÍBRIDA) ---
// =================================================================

function processarMateriaisRemoto(dadosGlobais) {
  try {
    if (!dadosGlobais) {
      dadosGlobais = obterDadosEntradasGlobal();
    }

    const ssMat = SpreadsheetApp.openById(CONFIG.ids.materiais);
    const wsEmpenhos = ssMat.getSheetByName(CONFIG.abas.empenhos);
    const wsRecProv = ssMat.getSheetByName(CONFIG.abas.recProvisorio);
    const wsDestino = ssMat.getSheetByName(CONFIG.abas.materiais);

    if (!wsEmpenhos || !wsDestino || !wsRecProv) {
      throw new Error("Abas 'Empenhos Enviados', 'Rec.Provisorio' ou 'Material' não encontradas em Materiais.");
    }

    const _cleanStr = (s) => String(s || '').trim().replace(/\.0$/, '').toUpperCase();

    const recProvisorioMap = new Map();
    const lastRowRec = wsRecProv.getLastRow();
    
    if (lastRowRec >= 2) {
      wsRecProv.getRange(2, 1, lastRowRec - 1, 6).getValues().forEach(row => {
        const itemCode = _cleanStr(_norm(row[0]));
        const qtd = Math.round(parseFloat(row[2]) || 0);
        const empRaw = String(row[5]).trim();
        if (itemCode && empRaw.includes('/')) {
          const parts = empRaw.split('/');
          if (parts.length === 2) {
            const numero = parts[0].padStart(4, '0');
            const ano = (parts[1].length === 2) ? '20' + parts[1] : parts[1];
            const empMontado = _cleanStr(`${ano}${numero}`);
            recProvisorioMap.set(`${empMontado}-${itemCode}`, qtd);
          }
        }
      });
    }

    const anotacoesMap = new Map();
    const lastRowDest = wsDestino.getLastRow();
    
    if (lastRowDest >= 2) {
      wsDestino.getRange(2, 1, lastRowDest - 1, 14).getValues().forEach(r => {
        const empDest = _cleanStr(r[0]);
        const itemDest = _cleanStr(_norm(r[5]));
        const k = `${empDest}-${itemDest}`;
        if (empDest) anotacoesMap.set(k, r.slice(9, 14));
      });
    }

    const entradasMap = new Map();
    const modalidadeMap = new Map();

    dadosGlobais.forEach(r => {
      const emp = _cleanStr(r[0]);
      const item = _cleanStr(_norm(r[2]));
      
      // --- NOVO FILTRO: IGNORAR ITENS COMEÇADOS COM C, A50, D, A70 ---
      if (/^(C|A50|D|A70)/i.test(item)) return;

      const key = `${emp}-${item}`;

      if (emp && item) {
        const qE = Math.round(parseFloat(r[17]) || 0);
        const qS = Math.round(parseFloat(r[19]) || 0);
        const saldoAEmpenhar = Math.round(parseFloat(r[21]) || 0);
        
        // Captura a data da coluna X (Índice 23)
        const dataUltimaEntrada = r[23] instanceof Date ? r[23] : null;

        if (entradasMap.has(key)) {
          let e = entradasMap.get(key);
          e.qE = Math.max(e.qE, qE);
          e.qS += qS;
          e.saldoAEmpenhar = Math.min(e.saldoAEmpenhar, saldoAEmpenhar);
          // Garante que mantém a data mais recente se houver duplicidade de chaves
          if (dataUltimaEntrada && (!e.dataEntrada || dataUltimaEntrada > e.dataEntrada)) {
            e.dataEntrada = dataUltimaEntrada;
          }
        } else {
          entradasMap.set(key, {
            proc: r[9],
            forn: r[8],
            desc: r[3],
            unit: r[14],
            qE: qE,
            val: r[23], // Mantido caso seja usado como valor em outra regra
            dataEntrada: dataUltimaEntrada, // Nova propriedade dedicada à data
            qS: qS,
            saldoAEmpenhar: saldoAEmpenhar,
            local: r[6] || ''
          });
        }
      }

      if (r[9]) modalidadeMap.set(String(r[9]).trim(), r[10]);
    });

    const output = [];
    const lastRowEmp = wsEmpenhos.getLastRow();

    // --- CÁLCULO DE LIMITE DE ANO PARA FILTRO ---
    const anoAtual = new Date().getFullYear();
    const anoMinimo = anoAtual - 1;

    if (lastRowEmp >= 2) {
      wsEmpenhos.getRange(2, 1, lastRowEmp - 1, 6).getValues().forEach(r => {
        const emp = _cleanStr(r[3]);
        const cod = _cleanStr(_norm(r[4]));

        if (!emp || !cod) return;

        // --- NOVO FILTRO: IGNORAR ITENS COMEÇADOS COM C, A50, D, A70 ---
        if (/^(C|A50|D|A70)/i.test(cod)) return;

        // --- APLICAÇÃO DO FILTRO DE ANO ---
        const anoEmp = _extrairAnoEmpenho(emp);
        if (anoEmp === null || anoEmp < anoMinimo) return;

        const key = `${emp}-${cod}`;
        
        let ent = entradasMap.get(key);
        
        if (!ent) {
          ent = { qE: 0, qS: 0, saldoAEmpenhar: 0, local: '', dataEntrada: null, val: '' };
        }

        const localLimpo = String(ent.local).replace(/\s+/g, '').toUpperCase();
        
        if (localLimpo !== 'ALM' && localLimpo !== '') return;

        const localFinal = 'ALM';
        const forn = ent.forn || r[1] || "Não informado";
        const desc = ent.desc || r[5] || "";
        const unit = ent.unit || "";
        const val = ent.val || "";
        const proc = ent.proc || "";
        const dataRecebimento = ent.dataEntrada || null;

        const qE = ent.qE;
        let qS_Oficial = ent.qS;
        const saldoAEmpenhar = Math.round(parseFloat(ent.saldoAEmpenhar) || 0);

        let status = (qE === 0) ? "Solicitar Associação" : "";

        let qS_Fisico = qS_Oficial;
        const isProvisorio = recProvisorioMap.has(key);

        if (isProvisorio) {
          const qtdProv = recProvisorioMap.get(key);
          if (qtdProv > qS_Fisico) {
            qS_Fisico = qtdProv;
          }
        }

        const saldoFisico = qE - qS_Fisico;

        // Se ainda não for Solicitar Associação, calcula normalmente.
        if (status !== "Solicitar Associação") {
           status = _calcularStatusUnificado(qE, qS_Oficial, saldoFisico, isProvisorio, status);
        }

        if (saldoAEmpenhar === 0 && qE > 0) {
          status = "Concluído";
        }

        if (saldoFisico === 0 && (status === "Recebido a Maior" || status === "Pendente" || status === "Recebimento Provisório")) {
          status = "Concluído";
        }

        let obsAtraso = "";
        const dVenc = (r[0] instanceof Date ? _addDays(r[0], 10) : null);

        // ---> MODIFICAÇÃO AQUI: Implementação da observação na Coluna R <---
        if (status === 'Saldo Cancelado') {
           obsAtraso = 'Saldo Cancelado';
        } else if (status === 'Concluído' || (saldoAEmpenhar === 0 && qE > 0)) {
          obsAtraso = 'Entregue';
        } else if (status === 'Solicitar Associação') {
          obsAtraso = '';
        } else if (status.includes('Pendente') && dVenc) {
          const diff = new Date().setHours(0, 0, 0, 0) - dVenc.getTime();
          obsAtraso = diff > 0 ? _diasParaTexto(Math.floor(diff / 86400000)) : "No prazo";
        }

        const empOriginal = String(r[3]).trim();
        const codOriginal = _norm(r[4]);

        const linha = [
          empOriginal,                                      // A
          localFinal,                                       // B
          (r[0] instanceof Date ? r[0] : null),             // C
          dVenc,                                            // D
          forn,                                             // E
          codOriginal,                                      // F
          desc,                                             // G
          unit,                                             // H
          qE,                                               // I
          ...(anotacoesMap.get(key) || Array(5).fill('')),  // J, K, L, M, N (Mantém notas manuais)
          dataRecebimento,                                  // O - RECEBIMENTO ENTRADA (Data da coluna X)
          qS_Fisico,                                        // P
          saldoFisico,                                      // Q
          obsAtraso,                                        // R
          status,                                           // S
          proc,                                             // T
          modalidadeMap.get(proc) || null                   // U
        ];
        output.push(linha);
      });
    }

    if (output.length > 0) {
      const maxLinhasAtual = wsDestino.getMaxRows();
      const maxColunasAtual = wsDestino.getMaxColumns();

      if (maxLinhasAtual > 1) {
        wsDestino.getRange(2, 1, maxLinhasAtual - 1, maxColunasAtual).clearContent().clearFormat();
      }

      const numLinhasDados = output.length;
      const numColunasDados = output[0].length;
      const margemSeguranca = 50;
      const linhasNecessariasTotal = numLinhasDados + 1;
      const linhasAlvo = linhasNecessariasTotal + margemSeguranca;

      if (maxLinhasAtual > linhasAlvo) {
        const linhasParaDeletar = maxLinhasAtual - linhasAlvo;
        if (linhasParaDeletar > 0) {
          try {
            wsDestino.deleteRows(linhasAlvo + 1, linhasParaDeletar);
          } catch (errDel) {
            Logger.log("Aviso: Não foi possível deletar linhas excedentes.");
          }
        }
      } else if (maxLinhasAtual < linhasNecessariasTotal) {
        const linhasParaAdicionar = linhasNecessariasTotal - maxLinhasAtual;
        wsDestino.insertRowsAfter(maxLinhasAtual, linhasParaAdicionar);
      }

      wsDestino.getRange(2, 1, numLinhasDados, numColunasDados).setValues(output);
      
      // Formata as colunas de data: C, D e agora a coluna O (coluna 15)
      wsDestino.getRange(2, 3, numLinhasDados, 2).setNumberFormat("dd/mm/yyyy");
      wsDestino.getRange(2, 15, numLinhasDados, 1).setNumberFormat("dd/mm/yyyy"); // Formato para RECEBIMENTO ENTRADA
      
      const coresGrid = output.map(r => {
        const statusCell = String(r[18]).trim().toUpperCase();
        const dataVenc = r[3];
        let cor = CONFIG.cores[statusCell] || null;

        if (!cor && statusCell === 'PENDENTE' && dataVenc instanceof Date) {
          const diff = new Date().setHours(0, 0, 0, 0) - dataVenc.getTime();
          if (diff > 0) cor = '#FFCDD2';
          else cor = '#FFFFE0';
        } else if (!cor && statusCell === 'PENDENTE') {
          cor = CONFIG.cores['PENDENTE'];
        }
        return new Array(21).fill(cor);
      });

      wsDestino.getRange(2, 1, numLinhasDados, 21).setBackgrounds(coresGrid);
      wsDestino.hideColumns(10);

      const range = wsDestino.getDataRange();
      if (range.getFilter()) range.getFilter().remove();
      range.createFilter();

    } else {
      Logger.log("Nenhum dado processado para inserir na aba de Materiais.");
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert("Erro Materiais: " + e.message);
  }
}