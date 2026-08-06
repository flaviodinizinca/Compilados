// =================================================================
// --- BLOCO 24: PROCESSAMENTO DE RESTOS A PAGAR ---
// =================================================================

const ID_PLANILHA_PLANEJADORES_RESTOS = '1as5-pkFFEZgXFpvO7kVID3JDdGgHNYhDEhQn7VQd2Dc';
const NOME_GUIA_MENU_PLANEJADORES = 'Guia Menu';
const NOME_GUIA_DASHBOARD_RP = 'Dashboard RP';

function processarRestosAPagar() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const promptAno = ui.prompt(
    'Processar Restos a Pagar',
    'Informe o ano que está sendo processado:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (promptAno.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const anoInformado = String(promptAno.getResponseText() || '').trim();
  if (!/^\d{4}$/.test(anoInformado)) {
    ui.alert('Processamento cancelado: informe um ano válido com 4 dígitos.');
    return;
  }

  const abaRestos = ss.getSheetByName('Restos_a_pagar');
  const abaCompilados = ss.getSheetByName('Compilados');
  
  if (!abaRestos) {
    ui.alert('Erro', "A guia 'Restos_a_pagar' não foi encontrada.", ui.ButtonSet.OK);
    return;
  }

  if (!abaCompilados) {
    ui.alert('Erro', "A guia 'Compilados' não foi encontrada.", ui.ButtonSet.OK);
    return;
  }

  const idFonteGeral = PropertiesService.getScriptProperties().getProperty('ID_FONTE_GERAL');
  if (!idFonteGeral) {
    ui.alert('Erro', "A propriedade de script 'ID_FONTE_GERAL' não está configurada.", ui.ButtonSet.OK);
    return;
  }

  const ssFonteGeral = SpreadsheetApp.openById(idFonteGeral);
  const abaEntradaEmpenhos = ssFonteGeral.getSheetByName('EntradaEmpenhos');
  
  if (!abaEntradaEmpenhos) {
    ui.alert('Erro', "A guia 'EntradaEmpenhos' não foi encontrada na planilha de ID_FONTE_GERAL.", ui.ButtonSet.OK);
    return;
  }

  const ultimaLinhaCompilados = abaCompilados.getLastRow();
  const ultimaLinhaEntrada = abaEntradaEmpenhos.getLastRow();
  const ultimaLinhaRestos = abaRestos.getLastRow();
  
  if (ultimaLinhaRestos < 2) {
    ui.alert('Aviso', "A guia 'Restos_a_pagar' está vazia.", ui.ButtonSet.OK);
    return;
  }

  const dadosCompilados = ultimaLinhaCompilados >= 2
    ? abaCompilados.getRange(1, 1, ultimaLinhaCompilados, 19).getValues() // A:S
    : [];
    
  const dadosEntradaEmpenhos = ultimaLinhaEntrada >= 2
    ? abaEntradaEmpenhos.getRange(1, 1, ultimaLinhaEntrada, 21).getValues() // A:U
    : [];

  const dadosRestos = abaRestos.getRange(1, 1, ultimaLinhaRestos, 8).getValues(); // A:H

  // ================================================================
  // MAPAS
  // ================================================================
  const mapaCompilados = new Map();
  const mapaPlanejadorPorEmpenho = new Map();
  const mapaPlanejadorEntradaPorEmpenho = new Map();
  
  for (let i = 1; i < dadosCompilados.length; i++) {
    const empenho = _normalizarEmpenhoBusca(dadosCompilados[i][0]); // A
    const codigoItem = _normalizarCodigoItem(dadosCompilados[i][5]); // F
    const planejador = String(dadosCompilados[i][7] || '').trim(); // H

    if (empenho && planejador && !mapaPlanejadorPorEmpenho.has(empenho)) {
      mapaPlanejadorPorEmpenho.set(empenho, planejador);
    }

    if (!empenho || !codigoItem) continue;

    const chave = _montarChaveEmpenhoItem(empenho, codigoItem);
    
    if (!mapaCompilados.has(chave)) {
      mapaCompilados.set(chave, {
        codigoItem: dadosCompilados[i][5],
        descricaoItem: dadosCompilados[i][6],
        planejador: planejador,
        qtdEmpenho: dadosCompilados[i][8],
        qtdRecebida: dadosCompilados[i][15],
        qtdResidual: dadosCompilados[i][16],
        status: dadosCompilados[i][18]
      });
    }
  }

  const mapaEntradaEmpenhos = new Map();

  for (let i = 1; i < dadosEntradaEmpenhos.length; i++) {
    const empenho = _normalizarEmpenhoBusca(dadosEntradaEmpenhos[i][0]); // A
    const codigoItem = _normalizarCodigoItem(dadosEntradaEmpenhos[i][2]); // C
    const planejadorEntrada = String(dadosEntradaEmpenhos[i][14] || '').trim(); // O

    if (empenho && planejadorEntrada && !mapaPlanejadorEntradaPorEmpenho.has(empenho)) {
      mapaPlanejadorEntradaPorEmpenho.set(empenho, planejadorEntrada);
    }

    if (!empenho || !codigoItem) continue;

    const chave = _montarChaveEmpenhoItem(empenho, codigoItem);
    
    if (!mapaEntradaEmpenhos.has(chave)) {
      mapaEntradaEmpenhos.set(chave, []);
    }

    mapaEntradaEmpenhos.get(chave).push({
      codigoItem: dadosEntradaEmpenhos[i][2],
      qtdEmpenhoEntrada: dadosEntradaEmpenhos[i][17],  // R
      valorEmpenhado: dadosEntradaEmpenhos[i][18],     // S
      qtdRecebidaEntrada: dadosEntradaEmpenhos[i][19], // T
      valorRecebido: dadosEntradaEmpenhos[i][20]       // U
    });
  }

  // ================================================================
  // PROCESSAMENTO
  // ================================================================
  const resultados = [];
  const resumoNaturezas = [];
  const resumoNaturezasAbaixo150 = [];
  const resumoPlanejadores = [];

  resultados.push([
    'Natureza Despesa',
    'NE CCor',
    'Empenho Editado',
    'Código Item',
    'Descrição Item',
    'Status',
    'Qtd Empenho',
    'Qtd Recebida',
    'Qtd Residual',
    'Valor Empenhado',
    'Valor Recebido',
    'Aviso / Diferença'
  ]);
  
  for (let i = 1; i < dadosRestos.length; i++) {
    const natureza = String(dadosRestos[i][1] || '').trim(); // B
    const tipoDespesa = String(dadosRestos[i][2] || '').trim(); // C
    const favorecido = String(dadosRestos[i][4] || '').trim(); // E
    const empenhoOriginal = String(dadosRestos[i][6] || '').trim(); // G
    const saldoMoedaOrigem = _toNumber(dadosRestos[i][7]); // H
    
    const empenhoEditado = _converterNeCcorParaEmpenho(empenhoOriginal);
    const planejador =
      mapaPlanejadorPorEmpenho.get(empenhoEditado) ||
      mapaPlanejadorEntradaPorEmpenho.get(empenhoEditado) ||
      '';

    // resumos por planejador usam todas as naturezas
    if (saldoMoedaOrigem !== 0) {
      resumoPlanejadores.push([
        natureza,
        tipoDespesa,
        favorecido,
        empenhoEditado || empenhoOriginal,
        saldoMoedaOrigem,
        planejador,
        '',
        '',
        ''
      ]);
    }

    // resumos por natureza agora usam todas as naturezas
    if (saldoMoedaOrigem !== 0) {
      const linhaResumo = [
        natureza,
        tipoDespesa,
        favorecido,
        empenhoEditado || empenhoOriginal,
        saldoMoedaOrigem
      ];

      resumoNaturezas.push(linhaResumo);
      
      if (saldoMoedaOrigem < 150) {
        resumoNaturezasAbaixo150.push(linhaResumo);
      }
    }

    // guia principal continua com o comportamento que você já vinha usando
    if (!natureza.startsWith('339030')) {
      continue;
    }

    if (!empenhoEditado) {
      resultados.push([
        natureza,
        empenhoOriginal,
        '',
        '',
        '',
        'Empenho inválido',
        '',
        '',
        '',
        '',
        '',
        ''
      ]);
      continue;
    }

    const itensEncontrados = _listarItensDoEmpenho(
      empenhoEditado,
      mapaCompilados,
      mapaEntradaEmpenhos
    );
    
    if (itensEncontrados.length === 0) {
      resultados.push([
        natureza,
        empenhoOriginal,
        empenhoEditado,
        '',
        '',
        'Não localizado',
        '',
        '',
        '',
        '',
        '',
        ''
      ]);
      continue;
    }

    itensEncontrados.forEach(item => {
      resultados.push([
        natureza,
        empenhoOriginal,
        empenhoEditado,
        item.codigoItem,
        item.descricaoItem,
        item.status,
        item.qtdEmpenho,
        item.qtdRecebida,
        item.qtdResidual,
        item.valorEmpenhadoEntrada,
        item.valorRecebidoEntrada,
        item.aviso
      ]);
    });
  }

  // ================================================================
  // GUIA PRINCIPAL
  // ================================================================
  const nomeGuiaDestino = anoInformado;
  let guiaDestino = ss.getSheetByName(nomeGuiaDestino);
  
  if (!guiaDestino) {
    guiaDestino = ss.insertSheet(nomeGuiaDestino);
  } else {
    guiaDestino.clearContents();
    guiaDestino.clearFormats();
  }

  guiaDestino.getRange(1, 1, resultados.length, resultados[0].length).setValues(resultados);

  guiaDestino.getRange(1, 1, 1, resultados[0].length)
    .setFontWeight('bold')
    .setBackground('#cfe2f3')
    .setBorder(true, true, true, true, true, true);
    
  if (resultados.length > 1) {
    guiaDestino.getRange(2, 7, resultados.length - 1, 3).setNumberFormat('#,##0.00');
    guiaDestino.getRange(2, 10, resultados.length - 1, 2).setNumberFormat('R$ #,##0.00');
    guiaDestino.getRange(2, 12, resultados.length - 1, 1).setNumberFormat('R$ #,##0.00');
  }

  guiaDestino.setFrozenRows(1);
  guiaDestino.autoResizeColumns(1, resultados[0].length);
  
  // ================================================================
  // GUIAS RESUMO
  // ================================================================
  const nomeGuiaResumo = `${anoInformado}_Naturezas`;
  let guiaResumo = ss.getSheetByName(nomeGuiaResumo);
  
  if (!guiaResumo) {
    guiaResumo = ss.insertSheet(nomeGuiaResumo);
  } else {
    guiaResumo.clearContents();
    guiaResumo.clearFormats();
  }

  _montarGuiaResumoNaturezas(
    guiaResumo,
    resumoNaturezas,
    `Resumo Executivo por Natureza de Despesa - ${anoInformado}`
  );
  
  const nomeGuiaResumo150 = `${anoInformado}_Naturezas_<150`;
  let guiaResumo150 = ss.getSheetByName(nomeGuiaResumo150);

  if (!guiaResumo150) {
    guiaResumo150 = ss.insertSheet(nomeGuiaResumo150);
  } else {
    guiaResumo150.clearContents();
    guiaResumo150.clearFormats();
  }

  _montarGuiaResumoNaturezas(
    guiaResumo150,
    resumoNaturezasAbaixo150,
    `Resumo Executivo por Natureza de Despesa (< 150) - ${anoInformado}`
  );
  
  // ================================================================
  // EXPORTAÇÃO POR PLANEJADOR
  // ================================================================
  _exportarGuiasPorPlanejador(resumoPlanejadores, anoInformado);

  ss.setActiveSheet(guiaDestino);
  
  if (resultados.length === 1) {
    ui.alert(
      'Aviso',
      "Nenhum registro com natureza de despesa iniciada em '339030' foi encontrado na guia 'Restos_a_pagar'.",
      ui.ButtonSet.OK
    );
    return;
  }

  ui.alert(
    'Processamento Concluído',
    `Foram geradas ${resultados.length - 1} linhas na aba '${nomeGuiaDestino}', criadas as guias '${nomeGuiaResumo}' e '${nomeGuiaResumo150}', e exportadas as guias por planejador na planilha externa.`,
    ui.ButtonSet.OK
  );
}

function importarRespostasPlanejadoresParaNaturezas() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const promptAno = ui.prompt(
    'Importar respostas dos planejadores',
    'Informe o ano da guia AAAA_Naturezas que deve ser atualizada:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (promptAno.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const anoInformado = String(promptAno.getResponseText() || '').trim();
  if (!/^\d{4}$/.test(anoInformado)) {
    ui.alert('Importação cancelada: informe um ano válido com 4 dígitos.');
    return;
  }

  const nomeGuiaResumo = `${anoInformado}_Naturezas`;
  const guiaResumo = ss.getSheetByName(nomeGuiaResumo);
  
  if (!guiaResumo) {
    ui.alert('Erro', `A guia '${nomeGuiaResumo}' não foi encontrada.`, ui.ButtonSet.OK);
    return;
  }

  const ssPlanejadores = SpreadsheetApp.openById(ID_PLANILHA_PLANEJADORES_RESTOS);
  const abas = ssPlanejadores.getSheets().filter(aba => aba.getName() !== NOME_GUIA_MENU_PLANEJADORES);
  const mapaRespostas = new Map();
  
  abas.forEach(aba => {
    const ultimaLinha = aba.getLastRow();
    if (ultimaLinha < 3) return;

    const ultimaColuna = Math.max(aba.getLastColumn(), 9);
    const dados = aba.getRange(1, 1, ultimaLinha, ultimaColuna).getValues();

    for (let i = 2; i < dados.length; i++) {
      const natureza = String(dados[i][0] || '').trim();
      const tipoDespesa = String(dados[i][1] || '').trim();
      const favorecido = String(dados[i][2] || '').trim();
      const empenho = String(dados[i][3] || '').trim();
      const saldo = _toNumber(dados[i][4]);
      const planejador = String(dados[i][5] || '').trim();
      const manter = String(dados[i][6] || '').trim();

      if (!natureza && !empenho && !planejador) continue;

      const chave = _montarChaveResumoNatureza(natureza, tipoDespesa, favorecido, empenho, saldo);
      mapaRespostas.set(chave, {
        planejador: planejador,
        manter: manter
      });
    }
  });
  
  const ultimaLinhaResumo = guiaResumo.getLastRow();
  if (ultimaLinhaResumo < 3) {
    ui.alert('Aviso', `A guia '${nomeGuiaResumo}' não possui dados para atualização.`, ui.ButtonSet.OK);
    return;
  }

  const dadosResumo = guiaResumo.getRange(1, 1, ultimaLinhaResumo, 7).getValues();
  
  for (let i = 0; i < dadosResumo.length; i++) {
    const linha = dadosResumo[i];
    const colA = String(linha[0] || '').trim();
    const colB = String(linha[1] || '').trim();
    const colC = String(linha[2] || '').trim();
    const colD = String(linha[3] || '').trim();
    const colE = _toNumber(linha[4]);
    
    if (
      colA === 'Natureza Despesa Detalhada' ||
      colA.indexOf('Natureza de Despesa:') === 0 ||
      colA === 'Total' ||
      colA.indexOf('Resumo Executivo') === 0 ||
      !colA
    ) {
      continue;
    }

    const chave = _montarChaveResumoNatureza(colA, colB, colC, colD, colE);
    
    if (mapaRespostas.has(chave)) {
      const resposta = mapaRespostas.get(chave);
      dadosResumo[i][5] = resposta.planejador; // F
      dadosResumo[i][6] = resposta.manter;     // G
    }
  }

  guiaResumo.getRange(1, 1, dadosResumo.length, 7).setValues(dadosResumo);
  guiaResumo.autoResizeColumns(1, 7);
  _criarOuAtualizarDashboardRP(ss, anoInformado, _calcularIndicadoresDashboardRP(dadosResumo));
  
  ui.alert(
    'Importação concluída',
    `As respostas dos planejadores foram importadas para a guia '${nomeGuiaResumo}' e o '${NOME_GUIA_DASHBOARD_RP}' foi atualizado.`,
    ui.ButtonSet.OK
  );
}

// =================================================================
// VERIFICAR PREENCHIMENTO (ATUALIZADO)
// =================================================================

function verificarPreenchimentoRestos() {
  const ui = SpreadsheetApp.getUi();
  const promptAno = ui.prompt(
    'Verificar Preenchimento',
    'Informe o ano para calcular o percentual de respostas:',
    ui.ButtonSet.OK_CANCEL
  );

  if (promptAno.getSelectedButton() !== ui.Button.OK) return;
  const anoInformado = String(promptAno.getResponseText() || '').trim();
  
  if (!/^\d{4}$/.test(anoInformado)) {
    ui.alert('Operação cancelada: informe um ano válido com 4 dígitos.');
    return;
  }

  const ssPlanejadores = SpreadsheetApp.openById(ID_PLANILHA_PLANEJADORES_RESTOS);
  const menuSheet = ssPlanejadores.getSheetByName(NOME_GUIA_MENU_PLANEJADORES);
  
  if (!menuSheet) {
    ui.alert('Erro', `A guia '${NOME_GUIA_MENU_PLANEJADORES}' não foi encontrada.`, ui.ButtonSet.OK);
    return;
  }

  const ultimaLinha = menuSheet.getLastRow();
  if (ultimaLinha < 4) {
    ui.alert('Aviso', 'Não há planejadores listados na Guia Menu.', ui.ButtonSet.OK);
    return;
  }

  const dadosMenu = menuSheet.getRange(4, 1, ultimaLinha - 3, 1).getValues();
  const estatisticas = [];

  for (let i = 0; i < dadosMenu.length; i++) {
    const planejador = String(dadosMenu[i][0] || '').trim();
    if (!planejador) continue;

    const nomeAba = _sanitizarNomeAba(`${anoInformado}_${planejador}`);
    const aba = ssPlanejadores.getSheetByName(nomeAba);

    if (aba) {
      const ultimaLinhaAba = aba.getLastRow();
      if (ultimaLinhaAba >= 3) {
        const dadosG = aba.getRange(3, 7, ultimaLinhaAba - 2, 1).getValues(); // Coluna G: "Manter?"
        const totalItens = dadosG.length;
        const respondidos = dadosG.filter(linha => {
          const valor = String(linha[0] || '').trim().toLowerCase();
          return valor === 'sim' || valor === 'não' || valor === 'nao';
        }).length;
        
        const percentual = totalItens > 0 ? respondidos / totalItens : 0;
        estatisticas.push({ planejador: planejador, percentual: percentual });
      } else {
        estatisticas.push({ planejador: planejador, percentual: 0 });
      }
    } else {
      estatisticas.push({ planejador: planejador, percentual: 0 });
    }
  }

  _atualizarPercentuaisGuiaMenu(menuSheet, estatisticas);
  
  ui.alert(
    'Concluído',
    `Os percentuais de preenchimento para o ano de ${anoInformado} foram atualizados com sucesso na Guia Menu.`,
    ui.ButtonSet.OK
  );
}

// =================================================================
// VERIFICAR E ATUALIZAR LINKS (PRESERVA COMPORTAMENTO INDEPENDENTE)
// =================================================================

function verificarAtualizarLinksRestos() {
  const ui = SpreadsheetApp.getUi();
  const promptAno = ui.prompt(
    'Verificar e Atualizar Links',
    'Informe o ano para atualizar os links das abas dos planejadores:',
    ui.ButtonSet.OK_CANCEL
  );

  if (promptAno.getSelectedButton() !== ui.Button.OK) return;
  const anoInformado = String(promptAno.getResponseText() || '').trim();
  
  if (!/^\d{4}$/.test(anoInformado)) {
    ui.alert('Operação cancelada: informe um ano válido com 4 dígitos.');
    return;
  }

  const ssPlanejadores = SpreadsheetApp.openById(ID_PLANILHA_PLANEJADORES_RESTOS);
  const menuSheet = ssPlanejadores.getSheetByName(NOME_GUIA_MENU_PLANEJADORES);
  
  if (!menuSheet) {
    ui.alert('Erro', `A guia '${NOME_GUIA_MENU_PLANEJADORES}' não foi encontrada.`, ui.ButtonSet.OK);
    return;
  }

  const ultimaLinha = menuSheet.getLastRow();
  if (ultimaLinha < 4) {
    ui.alert('Aviso', 'Não há planejadores listados na Guia Menu para atualizar os links.', ui.ButtonSet.OK);
    return;
  }

  const dadosMenu = menuSheet.getRange(4, 1, ultimaLinha - 3, 1).getValues();

  for (let i = 0; i < dadosMenu.length; i++) {
    const planejador = String(dadosMenu[i][0] || '').trim();
    if (!planejador) continue;

    const nomeAba = _sanitizarNomeAba(`${anoInformado}_${planejador}`);
    const aba = ssPlanejadores.getSheetByName(nomeAba);
    const linhaLink = i + 4;

    if (aba) {
      const urlPlanilha = ssPlanejadores.getUrl();
      const formula = `=HYPERLINK("${urlPlanilha}#gid=${aba.getSheetId()}"; "Abrir guia")`;
      menuSheet.getRange(linhaLink, 3).setFormula(formula);
    } else {
      menuSheet.getRange(linhaLink, 3).setValue('Aba não encontrada');
    }
  }

  ui.alert(
    'Concluído',
    `Os links para o ano de ${anoInformado} foram atualizados com sucesso na Guia Menu.`,
    ui.ButtonSet.OK
  );
}

// =================================================================
// FUNÇÕES AUXILIARES DE SUPORTE
// =================================================================

function _exportarGuiasPorPlanejador(linhas, anoInformado) {
  const ssPlanejadores = SpreadsheetApp.openById(ID_PLANILHA_PLANEJADORES_RESTOS);
  const grupos = new Map();
  
  linhas.forEach(linha => {
    const planejador = String(linha[5] || '').trim() || 'SEM_PLANEJADOR';
    if (!grupos.has(planejador)) {
      grupos.set(planejador, []);
    }
    grupos.get(planejador).push(linha);
  });
  
  const estatisticas = [];
  const menuSheet = _criarOuAtualizarGuiaMenuPlanejadores(ssPlanejadores, anoInformado, grupos);
  
  grupos.forEach((itens, planejador) => {
    const nomeAba = _sanitizarNomeAba(`${anoInformado}_${planejador}`);
    let aba = ssPlanejadores.getSheetByName(nomeAba);
    const respostasExistentes = aba ? _mapearRespostasExistentesPlanejador(aba) : new Map();

    if (!aba) {
      aba = ssPlanejadores.insertSheet(nomeAba);
    } else {
      aba.clearContents();
      aba.clearFormats();
      aba.getDataRange().clearDataValidations();
    }

    const itensComRespostasPreservadas = itens.map(linha => {
      const chave = _montarChaveResumoNatureza(linha[0], inline = linha[1], linha[2], linha[3], linha[4]);
      const dadosExistentes = respostasExistentes.get(chave) || { manter: '', observacao1: '', observacao2: '' };
      return [
        linha[0],
        linha[1],
        linha[2],
        linha[3],
        linha[4],
        linha[5],
        dadosExistentes.manter || '',
        dadosExistentes.observacao1 || '',
        dadosExistentes.observacao2 || ''
      ];
    });

    const dados = [[
      'Voltar para Guia Menu', '', '', '', '', '', '', '', ''
    ], [
      'Natureza Despesa Detalhada',
      'Tipo de despesa',
      'Favorecido',
      'NE CCor',
      'Saldo - Moeda Origem',
      'Planejador',
      'Manter?',
      'Observação 1',
      'Observação 2'
    ]].concat(itensComRespostasPreservadas);

    aba.getRange(1, 1, dados.length, 9).setValues(dados);

    aba.getRange(2, 1, 1, 9)
      .setFontWeight('bold')
      .setBackground('#cfe2f3')
      .setBorder(true, true, true, true, true, true);
      
    aba.getRange('A1')
      .setFormula(`=HYPERLINK("${ssPlanejadores.getUrl()}#gid=${menuSheet.getSheetId()}"; "⬅ Voltar para Guia Menu")`)
      .setFontWeight('bold')
      .setFontColor('#1155cc');
      
    aba.getRange(1, 1, 1, 9).setBackground('#f3f3f3');

    if (dados.length > 2) {
      aba.getRange(3, 5, dados.length - 2, 1).setNumberFormat('R$ #,##0.00');
      
      const regra = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Sim', 'Não'], true)
        .setAllowInvalid(false)
        .build();
        
      aba.getRange(3, 7, dados.length - 2, 1).setDataValidation(regra);
    }

    aba.setFrozenRows(2);
    aba.autoResizeColumns(1, 9);

    const totalItens = itensComRespostasPreservadas.length;
    
    const respondidos = itensComRespostasPreservadas.filter(linha => {
      const valor = String(linha[6] || '').trim().toLowerCase();
      return valor === 'sim' || valor === 'não' || valor === 'nao';
    }).length;
    
    const percentual = totalItens > 0 ? respondidos / totalItens : 0;
    
    estatisticas.push({
      planejador: planejador,
      percentual: percentual
    });
  });
  
  _atualizarPercentuaisGuiaMenu(menuSheet, estatisticas);
}

function _criarOuAtualizarGuiaMenuPlanejadores(ssPlanejadores, anoInformado, grupos) {
  let menuSheet = ssPlanejadores.getSheetByName(NOME_GUIA_MENU_PLANEJADORES);
  
  if (!menuSheet) {
    menuSheet = ssPlanejadores.insertSheet(NOME_GUIA_MENU_PLANEJADORES, 0);
  } else {
    menuSheet.clearContents();
    menuSheet.clearFormats();
  }

  menuSheet.getRange(1, 1, 1, 4).merge();
  menuSheet.getRange(1, 1)
    .setValue(`Planejadores - Restos a Pagar (${anoInformado})`)
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
    
  menuSheet.getRange(3, 1, 1, 4).setValues([['Planejador', 'Aba', 'Acesso', '% Respondido']]);
  menuSheet.getRange(3, 1, 1, 4)
    .setFontWeight('bold')
    .setBackground('#cfe2f3')
    .setBorder(true, true, true, true, true, true);
    
  const planejadoresOrdenados = Array.from(grupos.keys())
    .sort((a, b) => a.localeCompare(b, 'pt-BR', { numeric: true }));
    
  const linhas = planejadoresOrdenados.map(nome => {
    const nomeAba = _sanitizarNomeAba(`${anoInformado}_${nome}`);
    return [nome, nomeAba, '', ''];
  });
  
  if (linhas.length) {
    menuSheet.getRange(4, 1, linhas.length, 4).setValues(linhas);
  }

  planejadoresOrdenados.forEach((nome, index) => {
    const linha = 4 + index;
    const nomeAba = _sanitizarNomeAba(`${anoInformado}_${nome}`);
    let aba = ssPlanejadores.getSheetByName(nomeAba);
    if (!aba) {
      aba = ssPlanejadores.insertSheet(nomeAba);
    }
    menuSheet.getRange(linha, 3)
      .setFormula(`=HYPERLINK("${ssPlanejadores.getUrl()}#gid=${aba.getSheetId()}"; "Abrir guia")`);
  });
  
  menuSheet.setFrozenRows(3);
  menuSheet.autoResizeColumns(1, 4);
  return menuSheet;
}

// =================================================================
// FUNÇÃO FIXED: ATUALIZA EXCLUSIVAMENTE A COLUNA D (PRESERVA A COLUNA C)
// =================================================================
function _atualizarPercentuaisGuiaMenu(menuSheet, estatisticas) {
  if (!estatisticas.length) return;

  const mapaPercentuais = new Map();
  
  estatisticas.forEach(item => {
    mapaPercentuais.set(item.planejador, item.percentual);
  });

  const ultimaLinha = menuSheet.getLastRow();
  if (ultimaLinha < 4) return;
  
  // Seleciona de forma isolada a Coluna A (Planejadores) e a Coluna D (% Respondido)
  const rangePlanejadores = menuSheet.getRange(4, 1, ultimaLinha - 3, 1);
  const rangePercentuais = menuSheet.getRange(4, 4, ultimaLinha - 3, 1);
  
  const planejadores = rangePlanejadores.getValues();
  const percentuaisAtuais = rangePercentuais.getValues();

  for (let i = 0; i < planejadores.length; i++) {
    const planejador = String(planejadores[i][0] || '').trim();
    
    if (mapaPercentuais.has(planejador)) {
      percentuaisAtuais[i][0] = mapaPercentuais.get(planejador);
    }
  }

  // Grava de volta UNICAMENTE na coluna D, sem tocar na coluna C
  rangePercentuais.setValues(percentuaisAtuais);
  rangePercentuais.setNumberFormat('0%');
  menuSheet.autoResizeColumns(4, 1);
}

function _mapearRespostasExistentesPlanejador(aba) {
  const mapa = new Map();
  const ultimaLinha = aba.getLastRow();
  
  if (ultimaLinha < 3) return mapa;

  const ultimaColuna = Math.max(aba.getLastColumn(), 9);
  const dados = aba.getRange(3, 1, ultimaLinha - 2, ultimaColuna).getValues();
  
  dados.forEach(linha => {
    const natureza = String(linha[0] || '').trim();
    const tipoDespesa = String(linha[1] || '').trim();
    const favorecido = String(linha[2] || '').trim();
    const empenho = String(linha[3] || '').trim();
    const saldo = _toNumber(linha[4]);
    const manter = String(linha[6] || '').trim();
    const observacao1 = String(linha[7] || '').trim();
    const observacao2 = String(linha[8] || '').trim();

    if (!natureza || natureza === 'Natureza Despesa Detalhada') return;

    const chave = _montarChaveResumoNatureza(natureza, tipoDespesa, favorecido, empenho, saldo);
    mapa.set(chave, {
      manter: manter,
      observacao1: observacao1,
      observacao2: observacao2
    });
  });
  
  return mapa;
}

function _montarGuiaResumoNaturezas(sheet, linhas, titulo) {
  let linhaAtual = 1;

  sheet.getRange(linhaAtual, 1, 1, 7).merge();
  
  sheet.getRange(linhaAtual, 1)
    .setValue(titulo)
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
    
  linhaAtual += 2;

  if (!linhas.length) {
    sheet.getRange(linhaAtual, 1).setValue('Nenhum registro encontrado para este resumo.');
    sheet.autoResizeColumns(1, 7);
    return;
  }

  const grupos = new Map();

  linhas.forEach(linha => {
    const natureza = linha[0] || 'Sem Natureza';
    if (!grupos.has(natureza)) {
      grupos.set(natureza, []);
    }
    grupos.get(natureza).push(linha);
  });
  
  const naturezasOrdenadas = Array.from(grupos.keys())
    .sort((a, b) => a.localeCompare(b, 'pt-BR', { numeric: true }));
    
  naturezasOrdenadas.forEach(natureza => {
    const itens = grupos.get(natureza);
    const total = itens.reduce((acc, item) => acc + _toNumber(item[4]), 0);

    sheet.getRange(linhaAtual, 1, 1, 7).merge();
    sheet.getRange(linhaAtual, 1)
      .setValue(`Natureza de Despesa: ${natureza}`)
      .setFontWeight('bold')
      .setBackground('#d9eaf7');
    linhaAtual++;

    sheet.getRange(linhaAtual, 1, 1, 7).setValues([[
      'Natureza Despesa Detalhada',
      'Tipo de despesa',
      'Favorecido',
      'NE CCor',
      'Saldo - Moeda Origem',
      'Planejador',
      'Manter?'
    ]]);

    sheet.getRange(linhaAtual, 1, 1, 7)
      .setFontWeight('bold')
      .setBackground('#cfe2f3')
      .setBorder(true, true, true, true, true, true);
    linhaAtual++;

    const itensComRespostas = itens.map(linha => [
      linha[0],
      linha[1],
      linha[2],
      linha[3],
      linha[4],
      '',
      ''
    ]);

    sheet.getRange(linhaAtual, 1, itensComRespostas.length, 7).setValues(itensComRespostas);
    sheet.getRange(linhaAtual, 5, itensComRespostas.length, 1).setNumberFormat('R$ #,##0.00');
    linhaAtual += itensComRespostas.length;

    sheet.getRange(linhaAtual, 1, 1, 4).merge();
    
    sheet.getRange(linhaAtual, 1)
      .setValue('Total')
      .setFontWeight('bold')
      .setHorizontalAlignment('right')
      .setBackground('#e2f0d9');
      
    sheet.getRange(linhaAtual, 5)
      .setValue(total)
      .setFontWeight('bold')
      .setBackground('#e2f0d9')
      .setNumberFormat('R$ #,##0.00');
      
    linhaAtual += 2;
  });

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 7);
}

function _listarItensDoEmpenho(empenhoEditado, mapaCompilados, mapaEntradaEmpenhos) {
  const itensMap = new Map();
  
  mapaCompilados.forEach((infoCompilado, chave) => {
    const partes = chave.split('||');
    const empenho = partes[0];
    const codigoItem = partes[1];

    if (empenho !== empenhoEditado) return;

    if (!itensMap.has(codigoItem)) {
      itensMap.set(codigoItem, {
        codigoItem: infoCompilado.codigoItem || codigoItem,
        descricaoItem: infoCompilado.descricaoItem || '',
        status: infoCompilado.status || 'Não localizado',
        qtdEmpenho: '',
        qtdRecebida: '',
        qtdResidual: '',
        valorEmpenhadoEntrada: '',
        valorRecebidoEntrada: '',
        aviso: ''
      });
    }

    const chaveEntrada = _montarChaveEmpenhoItem(empenhoEditado, codigoItem);
    const entradas = mapaEntradaEmpenhos.get(chaveEntrada) || [];

    if (entradas.length > 0) {
      const totalQtdEmpenho = entradas.reduce((acc, item) => acc + _toNumber(item.qtdEmpenhoEntrada), 0);
      const totalQtdRecebida = entradas.reduce((acc, item) => acc + _toNumber(item.qtdRecebidaEntrada), 0);
      const totalEmpenhado = entradas.reduce((acc, item) => acc + _toNumber(item.valorEmpenhado), 0);
      const totalRecebido = entradas.reduce((acc, item) => acc + _toNumber(item.valorRecebido), 0);
      
      const itemAtual = itensMap.get(codigoItem);
      itemAtual.qtdEmpenho = totalQtdEmpenho;
      itemAtual.qtdRecebida = totalQtdRecebida;
      itemAtual.qtdResidual = totalQtdEmpenho - totalQtdRecebida;
      itemAtual.valorEmpenhadoEntrada = totalEmpenhado;
      itemAtual.valorRecebidoEntrada = totalRecebido;
      itemAtual.aviso = _compararValoresEQuantidades(totalQtdEmpenho, totalQtdRecebida, totalEmpenhado, totalRecebido);
    } else {
      const qtdEmpenhoComp = _toNumber(infoCompilado.qtdEmpenho);
      const qtdRecebidaComp = _toNumber(infoCompilado.qtdRecebida);
      const qtdResidualComp = _toNumber(infoCompilado.qtdResidual);

      const itemAtual = itensMap.get(codigoItem);
      itemAtual.qtdEmpenho = qtdEmpenhoComp || infoCompilado.qtdEmpenho || '';
      itemAtual.qtdRecebida = qtdRecebidaComp || infoCompilado.qtdRecebida || '';
      itemAtual.qtdResidual = (qtdEmpenhoComp || qtdRecebidaComp || qtdResidualComp)
        ? (qtdEmpenhoComp - qtdRecebidaComp)
        : (infoCompilado.qtdResidual || '');
    }
  });
  
  mapaEntradaEmpenhos.forEach((entradas, chave) => {
    const partes = chave.split('||');
    const empenho = partes[0];
    const codigoItem = partes[1];

    if (empenho !== empenhoEditado) return;
    if (itensMap.has(codigoItem)) return;

    const totalQtdEmpenho = entradas.reduce((acc, item) => acc + _toNumber(item.qtdEmpenhoEntrada), 0);
    const totalQtdRecebida = entradas.reduce((acc, item) => acc + _toNumber(item.qtdRecebidaEntrada), 0);
    const totalEmpenhado = entradas.reduce((acc, item) => acc + _toNumber(item.valorEmpenhado), 0);
    const totalRecebido = entradas.reduce((acc, item) => acc + _toNumber(item.valorRecebido), 0);

    itensMap.set(codigoItem, {
      codigoItem: entradas[0].codigoItem || codigoItem,
      descricaoItem: '',
      status: 'Localizado apenas em EntradaEmpenhos',
      qtdEmpenho: totalQtdEmpenho,
      qtdRecebida: totalQtdRecebida,
      qtdResidual: totalQtdEmpenho - totalQtdRecebida,
      valorEmpenhadoEntrada: totalEmpenhado,
      valorRecebidoEntrada: totalRecebido,
      aviso: _compararValoresEQuantidades(totalQtdEmpenho, totalQtdRecebida, totalEmpenhado, totalRecebido)
    });
  });

  return Array.from(itensMap.values()).sort((a, b) =>
    String(a.codigoItem).localeCompare(String(b.codigoItem), 'pt-BR', { numeric: true })
  );
}

function _compararValoresEQuantidades(qtdEmpenho, qtdRecebida, valorEmpenhado, valorRecebido) {
  if (qtdEmpenho > 0 && qtdEmpenho === qtdRecebida) {
    return 'Recebido';
  }

  const diferenca = _toNumber(valorEmpenhado) - _toNumber(valorRecebido);
  if (diferenca !== 0) {
    return diferenca;
  }

  return '';
}

function _calcularIndicadoresDashboardRP(dadosResumo) {
  const COLUNA_VALOR = 4;
  const COLUNA_TIPO_DESPESA = 1;
  const COLUNA_PLANEJADOR = 5;
  const COLUNA_MANTER = 6;
  const porNatureza = new Map();
  const porPlanejador = new Map();
  let valorVoltaBrasilia = 0;
  let valorFicaInca = 0;
  let valorTotalGeral = 0;

  for (const linha of dadosResumo) {
    const colA = String(linha[0] || '').trim();
    const colB = String(linha[COLUNA_TIPO_DESPESA] || '').trim();
    const colE = _toNumber(linha[COLUNA_VALOR]);
    const planejador = String(linha[COLUNA_PLANEJADOR] || '').trim() || 'Sem planejador';
    const manter = String(linha[COLUNA_MANTER] || '').trim();

    if (!_ehLinhaValidaDashboardRP(colA)) continue;

    // Acumula o valor do item para o Total Geral
    valorTotalGeral += colE;

    const manterNormalizado = _normalizarRespostaManterDashboardRP(manter);
    if (manterNormalizado !== 'sim' && manterNormalizado !== 'nao') continue;

    const dadosNatureza = _obterDadosNaturezaDashboardRP(colA, colB);
    const chaveNatureza = JSON.stringify([dadosNatureza.codigo, dadosNatureza.nomeNovo]);

    if (!porNatureza.has(chaveNatureza)) {
      porNatureza.set(chaveNatureza, {
        codigo: dadosNatureza.codigo,
        nomeNovo: dadosNatureza.nomeNovo,
        valorVoltaBrasilia: 0,
        valorFicaInca: 0
      });
    }

    if (!porPlanejador.has(planejador)) {
      porPlanejador.set(planejador, {
        planejador: planejador,
        quantidade: 0,
        valorVoltaBrasilia: 0,
        valorFicaInca: 0
      });
    }

    const linhaNatureza = porNatureza.get(chaveNatureza);
    const linhaPlanejador = porPlanejador.get(planejador);
    linhaPlanejador.quantidade += 1;

    if (manterNormalizado === 'sim') {
      valorFicaInca += colE;
      linhaNatureza.valorFicaInca += colE;
      linhaPlanejador.valorFicaInca += colE;
    } else if (manterNormalizado === 'nao') {
      valorVoltaBrasilia += colE;
      linhaNatureza.valorVoltaBrasilia += colE;
      linhaPlanejador.valorVoltaBrasilia += colE;
    }
  }

  // O pulo do gato para a prova real: o que sobra é o pendente
  const valorPendente = valorTotalGeral - valorFicaInca - valorVoltaBrasilia;

  const listaNaturezas = Array.from(porNatureza.values())
    .map(item => ({
      codigo: item.codigo,
      nomeNovo: item.nomeNovo,
      valorVoltaBrasilia: item.valorVoltaBrasilia,
      valorFicaInca: item.valorFicaInca,
      diferenca: item.valorFicaInca - item.valorVoltaBrasilia,
      total: item.valorFicaInca + item.valorVoltaBrasilia
    }))
    .sort((a, b) => {
      const porCodigo = String(a.codigo).localeCompare(String(b.codigo), 'pt-BR', { numeric: true });
      if (porCodigo !== 0) return porCodigo;
      return String(a.nomeNovo).localeCompare(String(b.nomeNovo), 'pt-BR');
    });

  const listaPlanejadores = Array.from(porPlanejador.values())
    .map(item => ({
      planejador: item.planejador,
      quantidade: item.quantidade,
      valorVoltaBrasilia: item.valorVoltaBrasilia,
      valorFicaInca: item.valorFicaInca,
      diferenca: item.valorFicaInca - item.valorVoltaBrasilia,
      total: item.valorFicaInca + item.valorVoltaBrasilia
    }))
    .sort((a, b) => String(a.planejador).localeCompare(String(b.planejador), 'pt-BR'));

  return {
    valorTotalGeral: valorTotalGeral,
    valorPendente: valorPendente, // Enviando o valor pendente para o Dashboard
    valorVoltaBrasilia: valorVoltaBrasilia,
    valorFicaInca: valorFicaInca,
    diferencaTotal: valorFicaInca - valorVoltaBrasilia,
    porNatureza: listaNaturezas,
    porPlanejador: listaPlanejadores
  };
}

function _criarOuAtualizarDashboardRP(ss, anoInformado, indicadores) {
  let dashboardSheet = ss.getSheetByName(NOME_GUIA_DASHBOARD_RP);

  if (!dashboardSheet) {
    dashboardSheet = ss.insertSheet(NOME_GUIA_DASHBOARD_RP);
  } else {
    dashboardSheet.clear();
  }
  dashboardSheet.getCharts().forEach(chart => dashboardSheet.removeChart(chart));

  dashboardSheet.getRange(1, 1, 1, 6).merge();
  dashboardSheet.getRange(1, 1)
    .setValue(`Dashboard RP (${anoInformado})`)
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  dashboardSheet.getRange(3, 1, 1, 2).setValues([['Resumo geral (Prova Real)', 'Valor']]);
  dashboardSheet.getRange(3, 1, 1, 2)
    .setFontWeight('bold')
    .setBackground('#cfe2f3')
    .setBorder(true, true, true, true, true, true);

  // Quadro de Resumo Geral reestruturado para mostrar a soma exata
  const dadosResumoGeral = [
    ['Total Geral (Soma de todos os itens na base)', indicadores.valorTotalGeral],
    ['1. Valor que fica no INCA (Respostas "Sim")', indicadores.valorFicaInca],
    ['2. Valor que volta para Brasília (Respostas "Não")', indicadores.valorVoltaBrasilia],
    ['3. Valor Pendente (Sem resposta ou inválido)', indicadores.valorPendente],
    ['Diferença de Decisão (Fica INCA - Volta Brasília)', indicadores.diferencaTotal]
  ];

  dashboardSheet.getRange(4, 1, dadosResumoGeral.length, 2).setValues(dadosResumoGeral);
  dashboardSheet.getRange(4, 2, dadosResumoGeral.length, 1).setNumberFormat('R$ #,##0.00');
  dashboardSheet.getRange(4, 1, dadosResumoGeral.length, 2).setBorder(true, true, true, true, true, true);

  // Destaca a linha do Total Geral para evidenciar a base de cálculo
  dashboardSheet.getRange(4, 1, 1, 2).setBackground('#e2f0d9').setFontWeight('bold');

  const cabecalhoNatureza = [['Código ND', 'Nome novo / descrição', 'Valor que volta para Brasília', 'Valor que fica no INCA', 'Diferença', 'Total analisado']];
  
  // Desceu para a linha 11 para acomodar o novo tamanho do quadro de resumo superior
  const linhaCabecalhoNatureza = 11;
  dashboardSheet.getRange(linhaCabecalhoNatureza, 1, 1, cabecalhoNatureza[0].length).setValues(cabecalhoNatureza);
  dashboardSheet.getRange(linhaCabecalhoNatureza, 1, 1, cabecalhoNatureza[0].length)
    .setFontWeight('bold')
    .setBackground('#cfe2f3')
    .setBorder(true, true, true, true, true, true);

  const linhasNatureza = indicadores.porNatureza.length
    ? indicadores.porNatureza.map(item => [
      item.codigo,
      item.nomeNovo,
      item.valorVoltaBrasilia,
      item.valorFicaInca,
      item.diferenca,
      item.total
    ])
    : [['Sem dados', '', 0, 0, 0, 0]];

  dashboardSheet.getRange(linhaCabecalhoNatureza + 1, 1, linhasNatureza.length, 6).setValues(linhasNatureza);
  dashboardSheet.getRange(linhaCabecalhoNatureza + 1, 3, linhasNatureza.length, 4).setNumberFormat('R$ #,##0.00');
  dashboardSheet.getRange(linhaCabecalhoNatureza + 1, 1, linhasNatureza.length, 6).setBorder(true, true, true, true, true, true);

  const linhaCabecalhoPlanejador = linhaCabecalhoNatureza + linhasNatureza.length + 3;

  const cabecalhoPlanejador = [['Planejador', 'Quantidade de itens', 'Valor que volta para Brasília', 'Valor que fica no INCA', 'Diferença', 'Total analisado']];

  dashboardSheet.getRange(linhaCabecalhoPlanejador, 1, 1, cabecalhoPlanejador[0].length).setValues(cabecalhoPlanejador);
  dashboardSheet.getRange(linhaCabecalhoPlanejador, 1, 1, cabecalhoPlanejador[0].length)
    .setFontWeight('bold')
    .setBackground('#cfe2f3')
    .setBorder(true, true, true, true, true, true);

  const linhasPlanejador = indicadores.porPlanejador.length
    ? indicadores.porPlanejador.map(item => [
      item.planejador,
      item.quantidade,
      item.valorVoltaBrasilia,
      item.valorFicaInca,
      item.diferenca,
      item.total
    ])
    : [['Sem dados', 0, 0, 0, 0, 0]];

  dashboardSheet.getRange(linhaCabecalhoPlanejador + 1, 1, linhasPlanejador.length, 6).setValues(linhasPlanejador);
  dashboardSheet.getRange(linhaCabecalhoPlanejador + 1, 3, linhasPlanejador.length, 4).setNumberFormat('R$ #,##0.00');
  dashboardSheet.getRange(linhaCabecalhoPlanejador + 1, 1, linhasPlanejador.length, 6).setBorder(true, true, true, true, true, true);

  _criarGraficosDashboardRP(dashboardSheet, indicadores.porNatureza, indicadores.porPlanejador);

  dashboardSheet.setFrozenRows(3);
  dashboardSheet.autoResizeColumns(1, 6);
}

function _ehLinhaValidaDashboardRP(colA) {
  return !(
    colA === 'Natureza Despesa Detalhada' ||
    colA.indexOf('Natureza de Despesa:') === 0 ||
    colA === 'Total' ||
    colA.indexOf('Resumo Executivo') === 0 ||
    !colA
  );
}

function _normalizarRespostaManterDashboardRP(valor) {
  return String(valor || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function _obterDadosNaturezaDashboardRP(naturezaDetalhada, tipoDespesa) {
  const naturezaTexto = String(naturezaDetalhada || '').trim();
  const codigoMatch = naturezaTexto.match(/^(\d+)/);
  const codigo = codigoMatch ? codigoMatch[1] : 'Sem código';
  const nomeNovo = String(tipoDespesa || '').trim() ||
    naturezaTexto.replace(/^\d+\s*[-–:]?\s*/, '').trim() ||
    naturezaTexto ||
    'Sem descrição';

  return { codigo: codigo, nomeNovo: nomeNovo };
}

function _criarGraficosDashboardRP(sheet, listaNaturezas, listaPlanejadores) {
  const COLUNA_GRAFICO_NATUREZA = 9;  // I
  const COLUNA_GRAFICO_PLANEJADOR = 13; // M

  if (listaNaturezas.length > 0) {
    const dadosNaturezaGrafico = [['Natureza de Despesa', 'Volta para Brasília', 'Fica no INCA']]
      .concat(listaNaturezas.map(item => [
        `${item.codigo} - ${item.nomeNovo}`,
        item.valorVoltaBrasilia,
        item.valorFicaInca
      ]));
    sheet.getRange(2, COLUNA_GRAFICO_NATUREZA, dadosNaturezaGrafico.length, 3).setValues(dadosNaturezaGrafico);

    const graficoNatureza = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange(2, COLUNA_GRAFICO_NATUREZA, dadosNaturezaGrafico.length, 3))
      .setPosition(3, 8, 0, 0)
      .setOption('title', 'Distribuição por Natureza de Despesa')
      .setOption('legend', { position: 'top' })
      .setOption('hAxis', { title: 'Natureza de Despesa' })
      .setOption('vAxis', { title: 'Valor (R$)' })
      .setOption('height', 320)
      .setOption('width', 780)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.SHOW_BOTH) // Mantém os dados visíveis no gráfico
      .build();
    sheet.insertChart(graficoNatureza);
  }

  if (listaPlanejadores.length > 0) {
    const dadosPlanejadorGrafico = [['Planejador', 'Volta para Brasília', 'Fica no INCA']]
      .concat(listaPlanejadores.map(item => [
        item.planejador,
        item.valorVoltaBrasilia,
        item.valorFicaInca
      ]));
    sheet.getRange(2, COLUNA_GRAFICO_PLANEJADOR, dadosPlanejadorGrafico.length, 3).setValues(dadosPlanejadorGrafico);

    const graficoPlanejador = sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(sheet.getRange(2, COLUNA_GRAFICO_PLANEJADOR, dadosPlanejadorGrafico.length, 3))
      .setPosition(22, 8, 0, 0)
      .setOption('title', 'Distribuição por planejador')
      .setOption('legend', { position: 'top' })
      .setOption('hAxis', { title: 'Valor (R$)' })
      .setOption('vAxis', { title: 'Planejador' })
      .setOption('height', 320)
      .setOption('width', 780)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.SHOW_BOTH) // Mantém os dados visíveis no gráfico
      .build();
    sheet.insertChart(graficoPlanejador);
  }

  sheet.hideColumns(COLUNA_GRAFICO_NATUREZA, 3);
  sheet.hideColumns(COLUNA_GRAFICO_PLANEJADOR, 3);
}

function _montarChaveResumoNatureza(natureza, tipoDespesa, favorecido, empenho, saldo) {
  return [natureza, tipoDespesa, favorecido, empenho, _toNumber(saldo).toFixed(2)].join('||');
}

function _sanitizarNomeAba(nome) {
  return String(nome || 'SEM_PLANEJADOR')
    .replace(/[\\\/\?\*\[\]\:]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 99);
}

function _converterNeCcorParaEmpenho(valor) {
  const texto = String(valor || '').trim().toUpperCase();
  if (!texto || texto.indexOf('NE') === -1) return '';
  
  const partes = texto.split('NE');
  if (partes.length !== 2) return '';

  const prefixo = partes[0].replace(/\D/g, '');
  const sufixo = partes[1].replace(/\D/g, '');
  
  if (prefixo.length < 4 || sufixo.length < 4) return '';

  const ano = prefixo.slice(-4);
  const ultimosQuatro = sufixo.slice(-4);
  
  if (!/^\d{4}$/.test(ano) || !/^\d{4}$/.test(ultimosQuatro)) return '';

  return ano + ultimosQuatro;
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

function _normalizarCodigoItem(valor) {
  return String(valor || '').trim();
}

function _montarChaveEmpenhoItem(empenho, codigoItem) {
  return `${empenho}||${codigoItem}`;
}

function _toNumber(valor) {
  if (typeof valor === 'number') return valor;
  
  if (valor === null || valor === undefined || valor === '') return 0;

  let texto = String(valor).trim();
  
  if (!texto) return 0;

  texto = texto.replace(/\./g, '').replace(',', '.');
  const numero = Number(texto);
  return isNaN(numero) ? 0 : numero;
}