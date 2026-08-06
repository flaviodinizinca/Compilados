// =================================================================
// --- BLOCO 10: FUNÇÕES AUXILIARES DE RELATÓRIOS ---
// =================================================================

function processarRelatorioItens(config) {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaRel = ss.getSheetByName(config.nomeAbaRelatorio);
    const abaComp = ss.getSheetByName("Compilados");
    if (!abaRel || !abaComp) throw new Error("Aba não encontrada.");

    const dados = abaComp.getRange("A2:T" + abaComp.getLastRow()).getValues();
    const mapa = new Map();

    for (const l of dados) {
      const cod = _norm(l[5]); if (!cod) continue;
      const st = _norm(l[18]);
      if (!mapa.has(cod)) mapa.set(cod, { pen: [], proc: [], conc: [] });
      const info = mapa.get(cod);
      if (st.includes('PENDENTE')) { if(l[0]) info.pen.push(l[0]); if (l[19]) info.proc.push(l[19]); }
      else if (st === 'CONCLUÍDO' || st.includes('PROVISÓRIO')) { if (l[0]) info.conc.push(l[0]); }
    }

    const lrRel = abaRel.getLastRow();
    if (lrRel < config.linhaInicialDados) { ui.alert("Vazio", ui.ButtonSet.OK); return; }
    const codigos = abaRel.getRange(config.linhaInicialDados, config.colunaCodigoItem, lrRel - config.linhaInicialDados + 1, 1).getValues();
    
    const resPen=[], resProc=[], resConc=[];
    let count = 0;
    
    codigos.forEach(r => {
        const c = _norm(r[0]);
        let p='', pr='', co='';
        if(c && mapa.has(c)) {
            const m = mapa.get(c);
            p = m.pen.join('\n'); pr = m.proc.join('\n'); co = m.conc.join('\n');
            if(m.pen.length || m.conc.length) count++;
        }
        resPen.push([p]); resProc.push([pr]); resConc.push([co]);
    });

    if (resPen.length) {
       abaRel.getRange(config.linhaInicialDados, config.colunaSaidaPendente, resPen.length, 1).setValues(resPen).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
       abaRel.getRange(config.linhaInicialDados, config.colunaSaidaProcesso, resPen.length, 1).setValues(resProc).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
       abaRel.getRange(config.linhaInicialDados, config.colunaSaidaConcluido, resPen.length, 1).setValues(resConc).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    }
    ui.alert("Concluído", `${count} correspondências.`, ui.ButtonSet.OK);
  } catch (e) { ui.alert("Erro", e.message, ui.ButtonSet.OK); }
}

function parseAtrasoParaDias(txt) {
  if (!txt) return 0;
  let t = 0; const s = txt.toLowerCase();
  const m = s.match(/(\d+)\s+m(ê|e)s/); if (m) t += parseInt(m[1])*30;
  const d = s.match(/(\d+)\s+dia/); if (d) t += parseInt(d[1]);
  return t;
}

function gerarBlocoDeRelatorio(sh, col, h, d, s, tit) {
  sh.getRange(1, col, 1, h.length).setValues([h]).setFontWeight("bold").setBackground("#1155cc").setFontColor("white");
  if(d.length) {
    sh.getRange(2, col, d.length, h.length).setValues(d);
    sh.getRange(2, col+3, d.length, 1).setNumberFormat('R$ #,##0.00');
    sh.getRange(2, col+5, d.length, 1).setNumberFormat('R$ #,##0.00');
    var r = 2+d.length;
    sh.getRange(r, col, 1, 4).merge().setValue(tit).setHorizontalAlignment("right").setFontWeight("bold");
    sh.getRange(r, col+4, 1, 2).merge().setValue(s).setNumberFormat('R$ #,##0.00').setFontWeight("bold");
  }
}