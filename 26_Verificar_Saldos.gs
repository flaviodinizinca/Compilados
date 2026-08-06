// =================================================================
// --- BLOCO 26: VERIFICAÇÃO DE SALDOS POR EMPENHO + ITEM ---
// =================================================================

function verificarSaldosMaterial() {
  verificarSaldosPorAba_("Material");
}

function verificarSaldosMedicamentos() {
  verificarSaldosPorAba_("Medicamentos");
}

function verificarSaldosMaterialEMedicamentos() {
  const ui = SpreadsheetApp.getUi();

  try {
    const atualizadosMaterial = verificarSaldosPorAba_("Material", true);
    const atualizadosMedicamentos = verificarSaldosPorAba_("Medicamentos", true);

    ui.alert(
      "Verificação concluída",
      "Atualizações realizadas:\n\n" +
      "Material: " + atualizadosMaterial + "\n" +
      "Medicamentos: " + atualizadosMedicamentos,
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert("Erro ao verificar saldos: " + e.message);
  }
}

function verificarSaldosPorAba_(nomeAbaDestino, executarSilencioso) {
  const ssFonteGeral = SpreadsheetApp.openById(CONFIG.ids.fonteDadosGeral);
  const abaEntradaEmpenhos = ssFonteGeral.getSheetByName(CONFIG.abas.entradaEmpenhos || "EntradaEmpenhos");

  if (!abaEntradaEmpenhos) {
    throw new Error('A guia "EntradaEmpenhos" não foi encontrada na Fonte Geral.');
  }

  let ssDestino;
  if (nomeAbaDestino === "Material") {
    ssDestino = SpreadsheetApp.openById(CONFIG.ids.materiais);
  } else if (nomeAbaDestino === "Medicamentos") {
    ssDestino = SpreadsheetApp.openById(CONFIG.ids.medicamentos);
  } else {
    throw new Error('A aba de destino deve ser "Material" ou "Medicamentos".');
  }

  const abaDestino = ssDestino.getSheetByName(nomeAbaDestino);
  if (!abaDestino) {
    throw new Error('A guia "' + nomeAbaDestino + '" não foi encontrada.');
  }

  const ultimaLinhaEntradaEmpenhos = abaEntradaEmpenhos.getLastRow();
  if (ultimaLinhaEntradaEmpenhos < 2) {
    if (!executarSilencioso) {
      SpreadsheetApp.getUi().alert('A guia "EntradaEmpenhos" não possui dados para comparação.');
    }
    return 0;
  }

  const dadosEntradaEmpenhos = abaEntradaEmpenhos.getRange(2, 1, ultimaLinhaEntradaEmpenhos - 1, 22).getValues();

  const mapaSaldosZerados = new Map();

  dadosEntradaEmpenhos.forEach(function(linha) {
    const empenho = String(linha[0]).trim();
    const codigoItem = _norm(linha[2]);
    const saldoAEmpenhar = parseFloat(linha[21]) || 0;

    if (!empenho || !codigoItem) return;

    if (saldoAEmpenhar <= 0 || saldoAEmpenhar < 1) {
      const chave = empenho + '|' + codigoItem;
      mapaSaldosZerados.set(chave, true);
    }
  });

  const ultimaLinhaDestino = abaDestino.getLastRow();
  if (ultimaLinhaDestino < 2) {
    if (!executarSilencioso) {
      SpreadsheetApp.getUi().alert('A guia "' + nomeAbaDestino + '" não possui dados para atualização.');
    }
    return 0;
  }

  const larguraLeitura = Math.max(abaDestino.getLastColumn(), 19);
  const dadosDestino = abaDestino.getRange(2, 1, ultimaLinhaDestino - 1, larguraLeitura).getValues();

  let totalAtualizados = 0;

  for (let i = 0; i < dadosDestino.length; i++) {
    const linha = dadosDestino[i];

    const empenho = String(linha[0]).trim();
    const codigoItem = _norm(linha[5]);

    if (!empenho || !codigoItem) continue;

    const chave = empenho + '|' + codigoItem;

    if (mapaSaldosZerados.has(chave)) {
      if (String(linha[18]).trim() !== "Concluído") {
        dadosDestino[i][18] = "Concluído";
        totalAtualizados++;
      }
    }
  }

  abaDestino.getRange(2, 1, dadosDestino.length, larguraLeitura).setValues(dadosDestino);

  if (!executarSilencioso) {
    SpreadsheetApp.getUi().alert(
      'Verificação concluída na guia "' + nomeAbaDestino + '".\n\n' +
      'Total de linhas atualizadas para "Concluído": ' + totalAtualizados
    );
  }

  return totalAtualizados;
}