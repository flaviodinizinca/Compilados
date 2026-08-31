// =================================================================
// --- BLOCO 3: FUNÇÕES AUXILIARES (HELPERS) - ATUALIZADO ---
// =================================================================

function obterDadosEntradasGlobal() {
  try {
    // 1. TENTATIVA DE RECUPERAÇÃO DE ID (BLINDAGEM CONTRA ERRO 'CONFIG NOT DEFINED')
    let idFonte;
    try {
      if (typeof CONFIG !== 'undefined' && CONFIG.ids) {
        idFonte = CONFIG.ids.fonteDadosGeral;
      }
    } catch (ignore) {}

    // Se o CONFIG falhou, busca direto nas Propriedades do Script (Fallback)
    if (!idFonte) {
      idFonte = PropertiesService.getScriptProperties().getProperty('ID_FONTE_GERAL');
    }

    if (!idFonte) {
      throw new Error("ID da Fonte de Dados Geral não encontrado (CONFIG ou ScriptProperties).");
    }

    // 2. ABERTURA DA PLANILHA
    const ssOrigem = SpreadsheetApp.openById(idFonte);
    
    // --- CORREÇÃO DO ERRO 'UNDEFINED' ---
    // Antes buscava 'fonteDadosNome', agora busca 'entradaEmpenhos'
    let nomeAba = "EntradaEmpenhos"; 
    
    if (typeof CONFIG !== 'undefined' && CONFIG.abas) {
      // Prioriza a configuração se ela existir
      if (CONFIG.abas.entradaEmpenhos) {
        nomeAba = CONFIG.abas.entradaEmpenhos;
      } else if (CONFIG.abas.fonteDadosNome) {
        nomeAba = CONFIG.abas.fonteDadosNome; 
      }
    }

    const abaOrigem = ssOrigem.getSheetByName(nomeAba);
    
    if (!abaOrigem) {
      throw new Error(`Aba '${nomeAba}' não encontrada na planilha fonte.`);
    }

    const dados = abaOrigem.getDataRange().getValues();
    if (dados.length > 0) {
      dados.shift(); // Remove cabeçalho
    }
    
    // 3. OTIMIZAÇÃO DE MEMÓRIA (FILTRO >= 2023)
    const anoCorte = 2023;
    const dadosFiltrados = dados.filter(linha => {
      const colA = linha[0]; // Pode ser Data ou Empenho
      if (!colA) return false; 
      
      if (colA instanceof Date) {
        return colA.getFullYear() >= anoCorte;
      }
      return true; // Se não for data (ex: número do empenho), mantém a linha.
    });
    
    console.log(`Dados carregados: ${dados.length} linhas totais. Mantidas: ${dadosFiltrados.length}.`);
    return dadosFiltrados;

  } catch (e) {
    throw new Error("Erro Crítico ao ler Fonte de Dados Global: " + e.message);
  }
}

// =================================================================
// --- FUNÇÕES UTILITÁRIAS GERAIS ---
// =================================================================

function _norm(v) { 
  return v ? String(v).trim().toUpperCase() : ""; 
}

/**
 * Extrai e valida o ano do empenho a partir do número no formato AAAANNNN.
 * Considera os 4 primeiros dígitos como o ano (ex.: 20254569 → 2025).
 * @param {*} valor - Valor da coluna A (número do empenho).
 * @returns {number|null} Ano extraído (4 dígitos entre 2000–2100) ou null se inválido.
 */
function _extrairAnoEmpenho(valor) {
  if (!valor) return null;
  const str = String(valor).trim();
  if (!/^\d{8,}$/.test(str)) return null;
  const ano = parseInt(str.substring(0, 4), 10);
  if (isNaN(ano) || ano < 2000 || ano > 2100) return null;
  return ano;
}

function _addDays(date, days) {
  let result = new Date(date);
  let added = 0;
  while (added < days) {
    result.setDate(result.getDate() + 1);
    added++;
  }
  return result;
}

function _diasParaTexto(total) {
  if (total < 0) total = 0;
  const m = Math.floor(total / 30);
  const d = total % 30;
  let txt = '';
  if (m > 0) txt += `${m} ${m > 1 ? 'meses' : 'mês'}`;
  if (d > 1) txt += (txt ? ' e ' : '') + `${d} dias`;
  else if (d === 1) txt += (txt ? ' e ' : '') + `${d} dia`;
  return txt || '0 dias';
}

function _parseDataSegura(valor) {
  if (!valor) return null;
  if (valor instanceof Date) return valor;
  if (typeof valor === 'string') {
    const partes = valor.trim().split('/');
    if (partes.length === 3) {
      return new Date(partes[2], partes[1] - 1, partes[0]);
    }
  }
  return null; 
}

function _calcularStatusUnificado(qEmpenhada, qSaidaOficial, saldoFisico, isRecProvisorio, statusAtual) {
  if (qEmpenhada > 0 && qSaidaOficial > qEmpenhada) return 'Recebido a Maior';
  if (qEmpenhada > 0 && qSaidaOficial === qEmpenhada) return 'Concluído'; 
  if (qEmpenhada === 0 && qSaidaOficial > 0) return 'Recebido. Falta associar';
  
  // ---> MODIFICAÇÃO AQUI: Implementação da regra do Saldo Cancelado <---
  if (qEmpenhada === 0) return 'Saldo Cancelado';

  if (isRecProvisorio && qSaidaOficial === 0) {
      if (saldoFisico > 0 && saldoFisico <= (qEmpenhada * 0.10)) return 'Resíduo 10%';
      if (saldoFisico > 0) return 'Rec. Prov. / Com Residuo';
      return 'Recebimento Provisório';
  }

  if (statusAtual === 'Empenho não está na guia "Entradas"') return statusAtual;
  
  if (qSaidaOficial === 0 && qEmpenhada > 0 && saldoFisico === qEmpenhada) return 'Pendente';
  if (saldoFisico > (qEmpenhada * 0.10)) return 'Pendente com Resíduo';
  if (saldoFisico > 0 && saldoFisico <= (qEmpenhada * 0.10)) return 'Resíduo 10%';
  if (saldoFisico <= 0) return 'Resíduo 10%'; 

  return statusAtual || 'Pendente'; 
}

// =================================================================
// --- CLASSIFICAÇÃO DE MOVIMENTO (NOVO 2026) ---
// =================================================================

// Listas Globais de Códigos de Documento
const TIPOS_CONSUMO_REAL = ["REQ", "RM", "RFS", "RDD", "RCS", "ACA", "NFS"];
const TIPOS_DEVOLUCAO = ["DEV", "DRM", "RRQ"];

/**
 * Classifica um código de documento para determinar seu papel no estoque.
 * @param {string} doc - O código do documento (ex: "REQ", "DEV", "TRA")
 * @returns {string} - "CONSUMO", "DEVOLUCAO" ou "NEUTRO"
 */
function verificarCategoriaMovimento(doc) {
  const d = _norm(doc);
  
  if (TIPOS_CONSUMO_REAL.includes(d)) {
    return "CONSUMO"; // Deve ser SOMADO ao uso
  }
  
  if (TIPOS_DEVOLUCAO.includes(d)) {
    return "DEVOLUCAO"; // Deve ser SUBTRAÍDO do uso
  }
  
  return "NEUTRO"; // Deve ser IGNORADO no cálculo de média
}

function _extrairAnoEmpenho(valorEmpenho) {
  if (valorEmpenho === null || valorEmpenho === undefined) return null;
  const digits = String(valorEmpenho).replace(/\D/g, "");
  if (digits.length < 4) return null;
  const ano = parseInt(digits.slice(0, 4), 10);
  if (!ano || ano < 2000 || ano > 2100) return null;
  return ano;
}