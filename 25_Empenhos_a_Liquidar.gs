// =================================================================
// --- BLOCO 25: PROCESSAMENTO DE EMPENHOS A LIQUIDAR (ATUALIZADO) ---
// =================================================================

function processarEmpenhosALiquidar() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const promptAno = ui.prompt(
    'Processar Empenhos a Liquidar',
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

  const abaEmpenhos = ss.getSheetByName('Empenhos_a_Liquidar');
  const abaCompilados = ss.getSheetByName('Compilados');
  
  if (!abaEmpenhos) {
    ui.alert('Erro', "A guia 'Empenhos_a_Liquidar' não foi encontrada.", ui.ButtonSet.OK);
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

  // ================================================================
  // 1. APAGAR LINHAS COM VALOR ZERO NA COLUNA H (Índice 7)
  // ================================================================
  let ultimaLinhaEmp = abaEmpenhos.getLastRow();
  if (ultimaLinhaEmp >= 2) {
    const dadosParaApagar = abaEmpenhos.getRange(1, 1, ultimaLinhaEmp, 8).getValues();
    for (let i = ultimaLinhaEmp - 1; i >= 1; i--) {
      const valorH = _toNumber(dadosParaApagar[i][7]);
      if (valorH === 0) {
        abaEmpenhos.deleteRow(i + 1);
      }
    }
  }

  ultimaLinhaEmp = abaEmpenhos.getLastRow();
  if (ultimaLinhaEmp < 2) {
    ui.alert('Aviso', "A guia 'Empenhos_a_Liquidar' está vazia ou todos os valores eram zero.", ui.ButtonSet.OK);
    return;
  }
  const dadosEmpenhos = abaEmpenhos.getRange(1, 1, ultimaLinhaEmp, 8).getValues();

  const ultimaLinhaCompilados = abaCompilados.getLastRow();
  const ultimaLinhaEntrada = abaEntradaEmpenhos.getLastRow();

  const dadosCompilados = ultimaLinhaCompilados >= 2
    ? abaCompilados.getRange(1, 1, ultimaLinhaCompilados, 19).getValues() 
    : [];
    
  const dadosEntradaEmpenhos = ultimaLinhaEntrada >= 2
    ? abaEntradaEmpenhos.getRange(1, 1, ultimaLinhaEntrada, 21).getValues() 
    : [];

  // ================================================================
  // MAPAS DE ITENS E DATAS
  // ================================================================
  const mapaCompilados = new Map();
  const mapaDataEnvio = new Map(); 

  for (let i = 1; i < dadosCompilados.length; i++) {
    const empenho = _normalizarEmpenhoBusca(dadosCompilados[i][0]);
    const dataEnvioBruta = dadosCompilados[i][2]; 
    const codigoItem = _normalizarCodigoItem(dadosCompilados[i][5]); 
    
    if (empenho && dataEnvioBruta && !mapaDataEnvio.has(empenho)) {
      mapaDataEnvio.set(empenho, _parseDate(dataEnvioBruta));
    }
    
    if (!empenho || !codigoItem) continue;

    const chave = _montarChaveEmpenhoItem(empenho, codigoItem);
    if (!mapaCompilados.has(chave)) {
      mapaCompilados.set(chave, {
        codigoItem: dadosCompilados[i][5],
        descricaoItem: dadosCompilados[i][6],
        qtdEmpenho: dadosCompilados[i][8],
        qtdRecebida: dadosCompilados[i][15],
        qtdResidual: dadosCompilados[i][16],
        status: dadosCompilados[i][18]
      });
    }
  }

  const mapaEntradaEmpenhos = new Map();
  for (let i = 1; i < dadosEntradaEmpenhos.length; i++) {
    const empenho = _normalizarEmpenhoBusca(dadosEntradaEmpenhos[i][0]);
    const codigoItem = _normalizarCodigoItem(dadosEntradaEmpenhos[i][2]);

    if (!empenho || !codigoItem) continue;

    const chave = _montarChaveEmpenhoItem(empenho, codigoItem);
    if (!mapaEntradaEmpenhos.has(chave)) {
      mapaEntradaEmpenhos.set(chave, []);
    }

    mapaEntradaEmpenhos.get(chave).push({
      codigoItem: dadosEntradaEmpenhos[i][2],
      qtdEmpenhoEntrada: dadosEntradaEmpenhos[i][17],  
      valorEmpenhado: dadosEntradaEmpenhos[i][18],     
      qtdRecebidaEntrada: dadosEntradaEmpenhos[i][19], 
      valorRecebido: dadosEntradaEmpenhos[i][20]       
    });
  }

  // ================================================================
  // PROCESSAMENTO E ALOCAÇÃO DE DADOS
  // ================================================================
  const resultados = [];
  const resumoNaturezas = [];
  const resumoNaturezasAbaixo150 = [];
  const mapaMesesAntigos = new Map(); 
  const mapaMesesAntigos339030 = new Map(); 

  resultados.push([
    'Natureza Despesa', 'NE CCor', 'Empenho Editado', 'Código Item', 'Descrição Item',
    'Status', 'Qtd Empenho', 'Qtd Recebida', 'Qtd Residual', 'Valor Empenhado', 'Valor Recebido', 'Aviso / Diferença'
  ]);
  
  const dataAtual = new Date();
  const mesAtual = dataAtual.getMonth(); 
  const limiteMes = mesAtual - 2;

  for (let i = 1; i < dadosEmpenhos.length; i++) {
    const natureza = String(dadosEmpenhos[i][1] || '').trim();
    const nomeNatureza = String(dadosEmpenhos[i][2] || '').trim();
    const nomeFavorecido = String(dadosEmpenhos[i][4] || '').trim();
    const empenhoOriginal = String(dadosEmpenhos[i][6] || '').trim();
    const valorEmpenho = _toNumber(dadosEmpenhos[i][7]);
    
    const empenhoEditado = _converterNeCcorParaEmpenho(empenhoOriginal);

    const dataEmissaoObj = _parseDate(dadosEmpenhos[i][5]); 
    const dataEnvioObj = empenhoEditado ? mapaDataEnvio.get(empenhoEditado) : null;
    
    const dataEmissaoStr = _formatDate(dataEmissaoObj);
    const dataEnvioStr = _formatDate(dataEnvioObj);
    const difDias = _diffDays(dataEmissaoObj, dataEnvioObj);

    const linhaResumo = [
      natureza,
      nomeNatureza,
      nomeFavorecido,
      empenhoEditado || empenhoOriginal,
      valorEmpenho,
      dataEmissaoStr,
      dataEnvioStr,
      difDias
    ];

    resumoNaturezas.push(linhaResumo);
    
    if (valorEmpenho <= 150) {
      resumoNaturezasAbaixo150.push(linhaResumo);
    }

    // ================================================================
    // LÓGICA DA ABA ANTIGOS (CORRIGIDA)
    // ================================================================
    let mesEmissao = -1;
    if (dataEmissaoObj) {
      mesEmissao = dataEmissaoObj.getMonth();
    }

    // Se o mês estiver entre Janeiro (0) e (Mês Atual - 2), ele entra na regra
    if (mesEmissao >= 0 && mesEmissao <= limiteMes) {
      const chaveMes = `${dataEmissaoObj.getFullYear()}-${String(dataEmissaoObj.getMonth() + 1).padStart(2, '0')}`;
      const nomeMesExtenso = `${['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'][dataEmissaoObj.getMonth()]} de ${dataEmissaoObj.getFullYear()}`;

      // Popula o Mapa GERAL
      if (!mapaMesesAntigos.has(chaveMes)) {
        mapaMesesAntigos.set(chaveMes, { titulo: nomeMesExtenso, itens: [], total: 0 });
      }
      mapaMesesAntigos.get(chaveMes).itens.push([natureza, nomeFavorecido, empenhoEditado || empenhoOriginal, dataEmissaoStr, valorEmpenho]);
      mapaMesesAntigos.get(chaveMes).total += valorEmpenho;

      // Popula o Mapa ESPECÍFICO (Somente 339030)
      if (natureza.startsWith('339030')) {
        if (!mapaMesesAntigos339030.has(chaveMes)) {
          mapaMesesAntigos339030.set(chaveMes, { titulo: nomeMesExtenso, itens: [], total: 0 });
        }
        mapaMesesAntigos339030.get(chaveMes).itens.push([natureza, nomeFavorecido, empenhoEditado || empenhoOriginal, dataEmissaoStr, valorEmpenho]);
        mapaMesesAntigos339030.get(chaveMes).total += valorEmpenho;
      }
    }

    if (!natureza.startsWith('339030')) {
      continue;
    }

    if (!empenhoEditado) {
      resultados.push([natureza, empenhoOriginal, '', '', '', 'Empenho inválido', '', '', '', '', '', '']);
      continue;
    }

    const itensEncontrados = _listarItensDoEmpenho(empenhoEditado, mapaCompilados, mapaEntradaEmpenhos);
    
    if (itensEncontrados.length === 0) {
      resultados.push([natureza, empenhoOriginal, empenhoEditado, '', '', 'Não localizado', '', '', '', '', '', '']);
      continue;
    }

    itensEncontrados.forEach(item => {
      resultados.push([
        natureza, empenhoOriginal, empenhoEditado, item.codigoItem, item.descricaoItem,
        item.status, item.qtdEmpenho, item.qtdRecebida, item.qtdResidual,
        item.valorEmpenhadoEntrada, item.valorRecebidoEntrada, item.aviso
      ]);
    });
  }

  // ================================================================
  // CRIAÇÃO: GUIA PRINCIPAL (AAAA)
  // ================================================================
  const nomeGuiaDestino = anoInformado;
  let guiaDestino = ss.getSheetByName(nomeGuiaDestino);
  
  if (!guiaDestino) {
    guiaDestino = ss.insertSheet(nomeGuiaDestino);
  } else {
    guiaDestino.clearContents();
    guiaDestino.clearFormats();
  }

  if (resultados.length > guiaDestino.getMaxRows()) {
    guiaDestino.insertRowsAfter(guiaDestino.getMaxRows(), (resultados.length - guiaDestino.getMaxRows()) + 20);
  }

  guiaDestino.getRange(1, 1, resultados.length, resultados[0].length).setValues(resultados);
  guiaDestino.getRange(1, 1, 1, resultados[0].length).setFontWeight('bold').setBackground('#cfe2f3').setBorder(true, true, true, true, true, true);
    
  if (resultados.length > 1) {
    guiaDestino.getRange(2, 7, resultados.length - 1, 3).setNumberFormat('#,##0.00');
    guiaDestino.getRange(2, 10, resultados.length - 1, 2).setNumberFormat('R$ #,##0.00');
    guiaDestino.getRange(2, 12, resultados.length - 1, 1).setNumberFormat('R$ #,##0.00');
  }

  guiaDestino.setFrozenRows(1);
  guiaDestino.autoResizeColumns(1, resultados[0].length);
  
  // ================================================================
  // CRIAÇÃO: GUIAS RESUMO ORDENADAS
  // ================================================================
  resumoNaturezas.sort((a, b) => _toNumber(b[4]) - _toNumber(a[4]));
  resumoNaturezasAbaixo150.sort((a, b) => _toNumber(b[4]) - _toNumber(a[4]));

  const nomeGuiaResumo = `${anoInformado}_Naturezas`;
  let guiaResumo = ss.getSheetByName(nomeGuiaResumo);
  
  if (!guiaResumo) {
    guiaResumo = ss.insertSheet(nomeGuiaResumo);
  } else {
    guiaResumo.clearContents();
    guiaResumo.clearFormats();
  }

  _montarGuiaResumoNaturezasOrdenado(guiaResumo, resumoNaturezas, `Resumo Empenhos a Liquidar por Natureza - ${anoInformado}`);
  
  const nomeGuiaResumo150 = `${anoInformado}_Naturezas_<150`;
  let guiaResumo150 = ss.getSheetByName(nomeGuiaResumo150);

  if (!guiaResumo150) {
    guiaResumo150 = ss.insertSheet(nomeGuiaResumo150);
  } else {
    guiaResumo150.clearContents();
    guiaResumo150.clearFormats();
  }

  _montarGuiaResumoNaturezasOrdenado(guiaResumo150, resumoNaturezasAbaixo150, `Resumo Empenhos a Liquidar (< 150) - ${anoInformado}`);
  
  // ================================================================
  // CRIAÇÃO: GUIAS DE ANTIGOS (GERAL E FILTRADA)
  // ================================================================
  const nomeGuiaAntigos = `${anoInformado}_Antigos`;
  _gerarAbaAntigos(ss, nomeGuiaAntigos, mapaMesesAntigos, `Empenhos por Mês de Emissão (Geral) - ${anoInformado}`);

  const nomeGuiaAntigos339030 = `${anoInformado}_Antigos_339030`;
  _gerarAbaAntigos(ss, nomeGuiaAntigos339030, mapaMesesAntigos339030, `Empenhos por Mês de Emissão (Apenas 339030) - ${anoInformado}`);

  ss.setActiveSheet(guiaDestino);
  
  ui.alert(
    'Processamento Concluído',
    `Abas criadas/atualizadas com sucesso:\n- '${nomeGuiaDestino}'\n- '${nomeGuiaResumo}'\n- '${nomeGuiaResumo150}'\n- '${nomeGuiaAntigos}'\n- '${nomeGuiaAntigos339030}'`,
    ui.ButtonSet.OK
  );
}

// =================================================================
// FUNÇÕES AUXILIARES EXCLUSIVAS PARA EMPENHOS A LIQUIDAR
// =================================================================

function _parseDate(valor) {
  if (!valor) return null;
  if (valor instanceof Date) return valor;
  const str = String(valor).trim();
  const matchBR = str.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (matchBR) {
    return new Date(matchBR[3], parseInt(matchBR[2], 10) - 1, matchBR[1]);
  }
  const d = new Date(str);
  if (!isNaN(d.getTime())) return d;
  return null;
}

function _formatDate(date) {
  if (!date || !(date instanceof Date) || isNaN(date.getTime())) return '';
  const d = String(date.getDate()).padStart(2, '0');
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const y = date.getFullYear();
  return `${d}/${m}/${y}`;
}

function _diffDays(d1, d2) {
  if (!d1 || !d2 || isNaN(d1.getTime()) || isNaN(d2.getTime())) return '';
  const t1 = d1.getTime();
  const t2 = d2.getTime();
  return Math.floor((t2 - t1) / (1000 * 3600 * 24));
}

function _gerarAbaAntigos(ss, nomeGuia, mapaMeses, tituloPrincipal) {
  let guia = ss.getSheetByName(nomeGuia);
  if (!guia) {
    guia = ss.insertSheet(nomeGuia);
  } else {
    guia.clearContents();
    guia.clearFormats();
  }

  const chavesOrdenadas = Array.from(mapaMeses.keys()).sort(); 

  if (chavesOrdenadas.length === 0) {
    guia.getRange(1, 1).setValue("Nenhum registro antigo encontrado para este período.");
    guia.autoResizeColumns(1, 1);
    return;
  }

  let linhasNecessarias = 3;
  chavesOrdenadas.forEach(chave => {
    linhasNecessarias += mapaMeses.get(chave).itens.length + 5; 
  });

  if (linhasNecessarias > guia.getMaxRows()) {
    guia.insertRowsAfter(guia.getMaxRows(), (linhasNecessarias - guia.getMaxRows()) + 20);
  }

  let linhaAtual = 1;

  guia.getRange(linhaAtual, 1, 1, 5).merge();
  guia.getRange(linhaAtual, 1).setValue(tituloPrincipal)
    .setFontSize(14).setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff').setHorizontalAlignment('center');
  linhaAtual += 2;

  chavesOrdenadas.forEach(chave => {
    const bloco = mapaMeses.get(chave);
    
    bloco.itens.sort((a, b) => _toNumber(b[4]) - _toNumber(a[4]));

    guia.getRange(linhaAtual, 1, 1, 5).merge().setValue(`Mês: ${bloco.titulo}`)
      .setFontWeight('bold').setBackground('#d9eaf7');
    linhaAtual++;

    guia.getRange(linhaAtual, 1, 1, 5).setValues([['Natureza', 'Nome do Favorecido', 'NE CCor', 'Data Emissão', 'Valor']]);
    guia.getRange(linhaAtual, 1, 1, 5).setFontWeight('bold').setBackground('#cfe2f3').setBorder(true, true, true, true, true, true);
    linhaAtual++;

    if (bloco.itens.length > 0) {
      guia.getRange(linhaAtual, 1, bloco.itens.length, 5).setValues(bloco.itens);
      guia.getRange(linhaAtual, 5, bloco.itens.length, 1).setNumberFormat('R$ #,##0.00');
      linhaAtual += bloco.itens.length;
    }

    guia.getRange(linhaAtual, 1, 1, 4).merge().setValue('Total do Mês').setFontWeight('bold').setHorizontalAlignment('right').setBackground('#e2f0d9');
    guia.getRange(linhaAtual, 5).setValue(bloco.total).setFontWeight('bold').setBackground('#e2f0d9').setNumberFormat('R$ #,##0.00');

    linhaAtual += 2;
  });

  guia.setFrozenRows(1);
  guia.autoResizeColumns(1, 5);
}

function _montarGuiaResumoNaturezasOrdenado(sheet, linhas, titulo) {
  if (!linhas.length) {
    sheet.getRange(1, 1, 1, 8).merge();
    sheet.getRange(1, 1).setValue(titulo).setFontSize(14).setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff').setHorizontalAlignment('center');
    sheet.getRange(3, 1).setValue('Nenhum registro encontrado para este resumo.');
    sheet.autoResizeColumns(1, 8);
    return;
  }

  const grupos = new Map();
  const totaisGrupos = new Map();

  linhas.forEach(linha => {
    const natureza = linha[0] || 'Sem Natureza';
    if (!grupos.has(natureza)) {
      grupos.set(natureza, []);
      totaisGrupos.set(natureza, 0);
    }
    grupos.get(natureza).push(linha);
    totaisGrupos.set(natureza, totaisGrupos.get(natureza) + _toNumber(linha[4])); 
  });
  
  const naturezasOrdenadas = Array.from(grupos.keys())
    .sort((a, b) => totaisGrupos.get(b) - totaisGrupos.get(a));

  let linhasNecessarias = 3; 
  naturezasOrdenadas.forEach(natureza => {
    linhasNecessarias += grupos.get(natureza).length + 5; 
  });

  const maxRows = sheet.getMaxRows();
  if (linhasNecessarias > maxRows) {
    sheet.insertRowsAfter(maxRows, (linhasNecessarias - maxRows) + 20);
  }

  let linhaAtual = 1;

  sheet.getRange(linhaAtual, 1, 1, 8).merge();
  sheet.getRange(linhaAtual, 1)
    .setValue(titulo).setFontSize(14).setFontWeight('bold')
    .setBackground('#1f4e78').setFontColor('#ffffff').setHorizontalAlignment('center');
    
  linhaAtual += 2;

  naturezasOrdenadas.forEach(natureza => {
    const itens = grupos.get(natureza);
    const total = totaisGrupos.get(natureza);

    sheet.getRange(linhaAtual, 1, 1, 8).merge();
    sheet.getRange(linhaAtual, 1)
      .setValue(`Natureza de Despesa: ${natureza}`)
      .setFontWeight('bold').setBackground('#d9eaf7');
    linhaAtual++;

    sheet.getRange(linhaAtual, 1, 1, 8).setValues([[
      'Natureza Despesa Detalhada', 'Nome Despesa Detalhada', 'Nome do Favorecido', 'NE CCor', 'Valor', 'Data Emissão', 'Data Envio', 'Dif. Dias'
    ]]);

    sheet.getRange(linhaAtual, 1, 1, 8)
      .setFontWeight('bold').setBackground('#cfe2f3').setBorder(true, true, true, true, true, true);
    linhaAtual++;

    const itensFormatados = itens.map(linha => [
      linha[0], linha[1], linha[2], linha[3], linha[4], linha[5], linha[6], linha[7]
    ]);

    sheet.getRange(linhaAtual, 1, itensFormatados.length, 8).setValues(itensFormatados);
    sheet.getRange(linhaAtual, 5, itensFormatados.length, 1).setNumberFormat('R$ #,##0.00');
    sheet.getRange(linhaAtual, 8, itensFormatados.length, 1).setHorizontalAlignment('center'); 
    linhaAtual += itensFormatados.length;

    sheet.getRange(linhaAtual, 1, 1, 4).merge().setValue('Total').setFontWeight('bold').setHorizontalAlignment('right').setBackground('#e2f0d9');
    sheet.getRange(linhaAtual, 5).setValue(total).setFontWeight('bold').setBackground('#e2f0d9').setNumberFormat('R$ #,##0.00');
    sheet.getRange(linhaAtual, 6, 1, 3).setBackground('#e2f0d9'); 
      
    linhaAtual += 2;
  });

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 8);
}