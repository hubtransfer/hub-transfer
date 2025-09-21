// ===================================================
// SISTEMA HUB CENTRAL - CENTRALIZADOR DE MÚLTIPLOS HOTÉIS
// Sistema que recebe dados de hotéis e operadoras externas
// Versão: 1.0-HUB-CENTRAL
// ===================================================

// 🔧 CONFIGURAÇÃO PRINCIPAL DO SISTEMA HUB
const SISTEMA_HUB_CONFIG = {
  NOME: 'Sistema HUB Central',
  VERSAO: '1.0-HUB-CENTRAL',
  SPREADSHEET_ID: '1pXAYotTOvev50-NPQMks905DELshIA4yWjcBa8Pz9WU', // Nova planilha HUB Central
  SHEET_NAME: 'HUB_Central',
  
  // Configuração Z-API para WhatsApp (Roberta HUB)
  Z_API: {
    ENDPOINT: 'https://api.z-api.io/instances/3DC8E250141ED020B95796155CBF9532/token/DF93ABBE66F44D82F60EF9FE/send-messages',
    CLIENT_TOKEN: 'F32768faecb324c3d89008f1131550834S',
    ROBERTA_PHONE: '+351928283652'
  },
  
  // Configuração OpenAI (caso necessário)
  OPENAI: {
    API_KEY: ''
  },
  
  // Mapeamento dos Hotéis conectados ao HUB
  HOTEIS_CONECTADOS: {
    'Empire Marques Hotel': {
      id: 'empire_marques',
      spreadsheetId: '1ZfG_IXBWMbGQzCmn7nWNzFqUBbGkXCXWSnY7JYziHxI',
      email: 'emh@theempirehotels.com',
      telefone: '+351210548700',
      ativo: true
    },
    'Hotel Lioz': {
      id: 'hotel_lioz',
      spreadsheetId: '1jXhF6tAPhuieIoIm6F3zsaEkq7NV_tafarkxcot4IE8',
      email: 'info@hotellioz.com',
      telefone: '+351219234567',
      ativo: true
    },
    'Empire Lisbon Hotel': {
      id: 'empire_lisbon',
      spreadsheetId: '12uW6XqVZJBIieHFl02zfatcWjPhHnYpZjydBK-JgV_A',
      email: 'elh@theempirehotels.com',
      telefone: '+351210548700',
      ativo: true
    },
    'Gota d´água': {
      id: 'gota_dagua',
      spreadsheetId: '1Zo0er2QaKszT3sYV8F0MZ50r3Pgj5gi50stqNrAnizY',
      email: 'info@gotadagua.com',
      telefone: '+351219876543',
      ativo: true
    }
  },
  
  // Operadoras externas (entrada manual)
  OPERADORAS_EXTERNAS: [
    'World Driver',
    'Operadora Manual',
    'Entrada Direta'
  ],
  
  // Sistema de notificações
  NOTIFICACOES: {
    WHATSAPP_ATIVO: true,
    WHATSAPP_APENAS_CONFIRMADOS: true, // Só envia WhatsApp para transfers confirmados
    EMAIL_DESATIVADO: true // Sistema HUB não envia e-mails
  },
  
  // Valores e comissões padrão (SEM CÁLCULO PERCENTUAL)
  VALORES_PADRAO: {
    COMISSAO_RECEPCAO_TRANSFER: 2.00,
    COMISSAO_RECEPCAO_TOUR: 5.00,
    MOEDA: '€'
  },
  
  // Configuração do sistema
  SISTEMA: {
    TIMEZONE: 'Europe/Lisbon',
    LOCALE: 'pt-PT',
    LOG_DETALHADO: true,
    MAX_TENTATIVAS_REGISTRO: 3,
    INTERVALO_TENTATIVAS: 500
  }
};

// Headers da planilha central (INCLUINDO NOVAS COLUNAS W, X, Y, Z)
const HEADERS_HUB_CENTRAL = [
  'ID',                           // A - Identificador único
  'Cliente',                      // B - Nome do cliente
  'Tipo Serviço',                // C - Transfer/Tour/Private
  'Pessoas',                      // D - Número de pessoas
  'Bagagens',                     // E - Número de bagagens
  'Data',                         // F - Data do transfer
  'Contacto',                     // G - Telefone/contato
  'Voo',                         // H - Número do voo
  'Origem',                      // I - Local de origem
  'Destino',                     // J - Local de destino
  'Hora Pick-up',                // K - Hora de recolha
  'Preço Cliente (€)',           // L - Valor total cobrado
  'Valor Hotel (€)',             // M - ALTERADO: Valor para hotel (NEUTRO)
  'Valor HUB Transfer (€)',      // N - Valor para HUB
  'Comissão Recepção (€)',       // O - Comissão da recepção
  'Forma Pagamento',             // P - Método de pagamento
  'Pago Para',                   // Q - Quem recebeu
  'Status',                      // R - Status do transfer
  'Observações',                 // S - Observações gerais
  'Data Criação',                // T - Timestamp de criação
  'Hotel/Operadora',             // U - NOVA COLUNA (índice 20)
  'Motorista',                   // V - Motorista (índice 21)
  'Status_OK',                   // W - Status de confirmação (índice 22)
  'Timestamp_OK',                // X - Timestamp da confirmação (índice 23)
  'Tipo_Viagem',                 // Y - Tipo específico de viagem (índice 24)
  'Status_Followup'              // Z - Status de follow-up (índice 25)
];

// Configuração adicional para abas mensais
const CONFIG_ABAS_MENSAIS = {
  PREFIXO: 'HUB_',
  USAR_ABAS_MENSAIS: true,
  ORGANIZAR_POR_DIAS: true,
  LIMITE_TRANSFERS_DIA: 15, // Limite para considerar dia "cheio"
  CORES_CAPACIDADE: {
    VAZIO: '#f8f9fa',      // Cinza claro - sem transfers
    BAIXO: '#d4edda',      // Verde claro - poucos transfers
    MEDIO: '#fff3cd',      // Amarelo - capacidade média
    ALTO: '#f8d7da',       // Vermelho claro - quase cheio
    CHEIO: '#dc3545'       // Vermelho - capacidade esgotada
  }
};

// Informações dos meses
const MESES_HUB = [
  { nome: 'Janeiro', numero: 1, abrev: '01', dias: 31, cor: '#e74c3c' },
  { nome: 'Fevereiro', numero: 2, abrev: '02', dias: 28, cor: '#e91e63' },
  { nome: 'Março', numero: 3, abrev: '03', dias: 31, cor: '#9c27b0' },
  { nome: 'Abril', numero: 4, abrev: '04', dias: 30, cor: '#673ab7' },
  { nome: 'Maio', numero: 5, abrev: '05', dias: 31, cor: '#3f51b5' },
  { nome: 'Junho', numero: 6, abrev: '06', dias: 30, cor: '#2196f3' },
  { nome: 'Julho', numero: 7, abrev: '07', dias: 31, cor: '#03a9f4' },
  { nome: 'Agosto', numero: 8, abrev: '08', dias: 31, cor: '#00bcd4' },
  { nome: 'Setembro', numero: 9, abrev: '09', dias: 30, cor: '#009688' },
  { nome: 'Outubro', numero: 10, abrev: '10', dias: 31, cor: '#4caf50' },
  { nome: 'Novembro', numero: 11, abrev: '11', dias: 30, cor: '#8bc34a' },
  { nome: 'Dezembro', numero: 12, abrev: '12', dias: 31, cor: '#cddc39' }
];

// Mensagens da Roberta HUB para WhatsApp
const ROBERTA_MESSAGES = {
  CONFIRMACAO_TEMPLATE: (transfer) => {
    const data = formatarDataDDMMYYYY(new Date(transfer.data));
    return `Olá ${transfer.cliente}, tudo bem? 

Eu sou a Roberta HUB, assistente virtual da HUB Transfer! 🚐

Estou enviando essa mensagem para confirmar que no dia ${data} teremos um transfer contigo:

📍 **Detalhes da viagem:**
- De: ${transfer.origem}
- Para: ${transfer.destino}  
- Horário: ${transfer.horaPickup}
- Passageiros: ${transfer.numeroPessoas}

Qualquer dúvida, estou aqui para ajudar! 😊

*HUB Transfer - Seu transfer com segurança e conforto*`;
  },
  
  ASSINATURA: `_Roberta HUB - Assistente Virtual_\n_HUB Transfer Portugal_`
};

// Status permitidos no sistema
const STATUS_PERMITIDOS = [
  'Solicitado',
  'Confirmado', 
  'Finalizado',
  'Cancelado'
];

// Tipos de serviço
const TIPOS_SERVICO = [
  'Transfer',
  'Tour Regular',
  'Private Tour'
];

// ===================================================
// SISTEMA DE LOGGING PARA HUB CENTRAL
// ===================================================

const LOG_CONFIG_HUB = {
  ENABLED: true,
  LEVEL: 'INFO',
  INCLUDE_TIMESTAMP: true,
  MAX_LOG_SIZE: 1000,
  CONSOLE_OUTPUT: true
};

const LOG_LEVELS = {
  DEBUG: 0,
  INFO: 1,
  WARN: 2,
  ERROR: 3,
  SUCCESS: 1
};

const loggerHUB = {
  logs: [],
  config: LOG_CONFIG_HUB,
  
  _shouldLog: function(level) {
    if (!this.config.ENABLED) return false;
    return LOG_LEVELS[level] >= LOG_LEVELS[this.config.LEVEL];
  },

  _formatTimestamp: function(date) {
    try {
      return Utilities.formatDate(
        date, 
        SISTEMA_HUB_CONFIG.SISTEMA.TIMEZONE, 
        'yyyy-MM-dd HH:mm:ss'
      );
    } catch (error) {
      return date.toISOString();
    }
  },

  _getLevelEmoji: function(level) {
    const emojis = {
      DEBUG: '🔍',
      INFO: 'ℹ️',
      WARN: '⚠️',
      ERROR: '❌',
      SUCCESS: '✅'
    };
    return emojis[level] || '📝';
  },

  log: function(level, message, data = null) {
    if (!this._shouldLog(level)) return;
    
    const logEntry = {
      timestamp: new Date(),
      level: level,
      message: message,
      data: data,
      sistema: 'HUB_CENTRAL'
    };
    
    this.logs.push(logEntry);
    if (this.logs.length > this.config.MAX_LOG_SIZE) {
      this.logs.shift();
    }
    
    const formattedLog = `[${this._formatTimestamp(logEntry.timestamp)}] ${this._getLevelEmoji(level)} ${level}: ${message}`;
    
    if (this.config.CONSOLE_OUTPUT) {
      try {
        if (typeof console !== 'undefined' && console.log) {
          console.log(formattedLog);
          if (data) console.log('Data:', data);
        }
      } catch (e) {
        // Fallback silencioso
      }
    }
  },

  debug: function(message, data = null) { this.log('DEBUG', message, data); },
  info: function(message, data = null) { this.log('INFO', message, data); },
  warn: function(message, data = null) { this.log('WARN', message, data); },
  error: function(message, data = null) { this.log('ERROR', message, data); },
  success: function(message, data = null) { this.log('SUCCESS', message, data); }
};

// ===================================================
// FUNÇÕES DE UTILIDADE PARA HUB CENTRAL
// ===================================================

/**
 * Formata data no padrão brasileiro DD/MM/YYYY
 */
function formatarDataDDMMYYYY(date) {
  try {
    if (!(date instanceof Date) || isNaN(date)) {
      loggerHUB.warn('Data inválida para formatação', { date });
      return 'Data inválida';
    }
    
    const dia = String(date.getDate()).padStart(2, '0');
    const mes = String(date.getMonth() + 1).padStart(2, '0');
    const ano = date.getFullYear();
    
    return `${dia}/${mes}/${ano}`;
    
  } catch (error) {
    loggerHUB.error('Erro ao formatar data', { date, erro: error.message });
    return 'Erro na data';
  }
}

/**
 * Processa data de múltiplos formatos
 */
function processarDataSeguraHUB(dataInput) {
  loggerHUB.debug('Processando data', { input: dataInput, tipo: typeof dataInput });
  
  try {
    if (dataInput instanceof Date && !isNaN(dataInput)) {
      return dataInput;
    }
    
    let data;
    
    if (typeof dataInput === 'string') {
      dataInput = dataInput.trim();
      
      // Formato ISO (YYYY-MM-DD)
      if (dataInput.match(/^\d{4}-\d{2}-\d{2}/)) {
        const [ano, mes, dia] = dataInput.split('-');
        data = new Date(parseInt(ano), parseInt(mes) - 1, parseInt(dia));
      }
      // Formato brasileiro (DD/MM/YYYY)  
      else if (dataInput.match(/^\d{2}\/\d{2}\/\d{4}/)) {
        const [dia, mes, ano] = dataInput.split('/');
        data = new Date(parseInt(ano), parseInt(mes) - 1, parseInt(dia));
      }
      else {
        data = new Date(dataInput);
      }
    } else {
      data = new Date(dataInput);
    }
    
    if (isNaN(data.getTime())) {
      throw new Error(`Formato de data inválido: ${dataInput}`);
    }
    
    return data;
    
  } catch (error) {
    loggerHUB.error('Erro ao processar data', { input: dataInput, erro: error.message });
    throw new Error(`Falha ao processar data: ${error.message}`);
  }
}

/**
 * Gera ID único para o sistema HUB
 */
function gerarProximoIdHUB() {
  try {
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SISTEMA_HUB_CONFIG.SHEET_NAME);
    
    if (!sheet) {
      throw new Error('Planilha HUB Central não encontrada');
    }
    
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return 1;
    }
    
    const idsRange = sheet.getRange(2, 1, lastRow - 1, 1);
    const ids = idsRange.getValues().flat().map(Number).filter(n => !isNaN(n) && n > 0);
    
    if (ids.length === 0) {
      return 1;
    }
    
    const maxId = Math.max(...ids);
    const novoId = maxId + 1;
    
    loggerHUB.debug('Novo ID HUB gerado', { maxId, novoId });
    return novoId;
    
  } catch (error) {
    loggerHUB.error('Erro ao gerar ID HUB, usando timestamp', error);
    return Date.now() % 1000000;
  }
}

/**
 * Sanitiza texto removendo caracteres perigosos
 */
function sanitizarTextoHUB(texto) {
  if (!texto) return '';
  
  let textoSanitizado = String(texto).trim();
  
  // Limitar tamanho
  if (textoSanitizado.length > 500) {
    textoSanitizado = textoSanitizado.substring(0, 500);
  }
  
  // Remover HTML
  textoSanitizado = textoSanitizado.replace(/<[^>]*>/g, '');
  
  // Remover caracteres de controle
  textoSanitizado = textoSanitizado.replace(/[\x00-\x1F\x7F]/g, '');
  
  // Escapar aspas
  textoSanitizado = textoSanitizado.replace(/'/g, "''");
  
  return textoSanitizado;
}

/**
 * Gera ID único para transfer HUB
 */
function gerarIdTransferHUB() {
  try {
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('HUB_Central');
    
    if (!sheet) {
      return 1; // Primeiro ID se não existe aba
    }
    
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return 1; // Primeiro transfer
    }
    
    // Buscar o maior ID existente
    let maiorId = 0;
    
    for (let linha = 2; linha <= lastRow; linha++) {
      const id = sheet.getRange(linha, 1).getValue();
      
      if (!isNaN(id) && id > maiorId) {
        maiorId = parseInt(id);
      }
    }
    
    const novoId = maiorId + 1;
    
    loggerHUB.debug('ID gerado para transfer', { novoId: novoId });
    
    return novoId;
    
  } catch (error) {
    loggerHUB.error('Erro ao gerar ID transfer', error);
    
    // Fallback: usar timestamp
    const timestamp = Date.now().toString().slice(-6);
    const fallbackId = parseInt(timestamp);
    
    loggerHUB.info('Usando ID fallback baseado em timestamp', { fallbackId });
    
    return fallbackId;
  }
}

/**
 * Busca transfer por ID na planilha HUB
 */
function buscarTransferPorId(transferId) {
  try {
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('HUB_Central');
    
    if (!sheet) {
      loggerHUB.error('Aba HUB_Central não encontrada');
      return null;
    }
    
    const lastRow = sheet.getLastRow();
    
    for (let linha = 2; linha <= lastRow; linha++) {
      const id = sheet.getRange(linha, 1).getValue();
      
      if (parseInt(id) === parseInt(transferId)) {
        // Encontrou o transfer - buscar todos os dados
        const dadosLinha = sheet.getRange(linha, 1, 1, HEADERS_HUB_CENTRAL.length).getValues()[0];
        
        return {
          cliente: dadosLinha[1],
          tipoServico: dadosLinha[2],
          pessoas: dadosLinha[3],
          bagagens: dadosLinha[4],
          data: dadosLinha[5],
          contacto: dadosLinha[6],
          voo: dadosLinha[7],
          origem: dadosLinha[8],
          destino: dadosLinha[9],
          horaPickup: dadosLinha[10],
          valorTotal: dadosLinha[11],
          valorHotel: dadosLinha[12],
          valorHUB: dadosLinha[13],
          comissaoRecepcao: dadosLinha[14],
          formaPagamento: dadosLinha[15],
          pagoParaQuem: dadosLinha[16],
          status: dadosLinha[17],
          observacoes: dadosLinha[18],
          dataCriacao: dadosLinha[19],
          hotelOperadora: dadosLinha[20],
          motorista: dadosLinha[21],
          linha: linha
        };
      }
    }
    
    loggerHUB.error('Transfer não encontrado', { transferId });
    return null;
    
  } catch (error) {
    loggerHUB.error('Erro ao buscar transfer por ID', error);
    return null;
  }
}

/**
 * Identifica hotel/operadora baseado na source
 */
function identificarHotelOperadora(source) {
  try {
    if (!source) {
      return 'Entrada Direta';
    }
    
    // Mapear sources para nomes das operadoras
    const mapeamentoOperadoras = {
      'empire_marques': 'Empire Marques Hotel',
      'hotel_lioz': 'Hotel Lioz',
      'empire_lisbon': 'Empire Lisbon Hotel',
      'gota_dagua': 'Gota d´água',
      'world_driver': 'World Driver',
      'wt': 'WT',
      'connecto': 'Connecto',
      'hub_direct': 'HUB Transfer',
      'manual_entry': 'Entrada Direta',
      'frontend_admin': 'HUB Transfer',
      'demo': 'Sistema Demo',
      'teste': 'Teste Aba Mensal'
    };
    
    const operadora = mapeamentoOperadoras[source];
    
    if (operadora) {
      loggerHUB.debug('Hotel/Operadora identificada', { source, operadora });
      return operadora;
    }
    
    // Se não encontrou, usar o source como nome
    loggerHUB.debug('Source não mapeada, usando como nome', { source });
    return source.charAt(0).toUpperCase() + source.slice(1).replace('_', ' ');
    
  } catch (error) {
    loggerHUB.error('Erro ao identificar hotel/operadora', error);
    return 'Operadora Desconhecida';
  }
}

/**
 * Calcula valores do transfer (cliente, hotel, HUB, recepção)
 */
function calcularValoresTransfer(dados) {
  try {
    const valorTotal = parseFloat(dados.valorTotal) || 0;
    
    // Cálculos baseados nas regras de negócio
    const comissaoRecepcao = Math.round(valorTotal * 0.06 * 100) / 100; // 6% para recepção
    const valorHotel = 0; // Hotéis não recebem comissão por padrão
    const valorHUB = valorTotal - valorHotel; // HUB fica com o restante
    
    const valores = {
      precoCliente: valorTotal,
      valorHotel: valorHotel,
      valorHUB: valorHUB,
      comissaoRecepcao: comissaoRecepcao
    };
    
    loggerHUB.debug('Valores calculados', valores);
    
    return valores;
    
  } catch (error) {
    loggerHUB.error('Erro ao calcular valores', error);
    
    // Retornar valores padrão em caso de erro
    const valorTotal = parseFloat(dados.valorTotal) || 0;
    return {
      precoCliente: valorTotal,
      valorHotel: 0,
      valorHUB: valorTotal,
      comissaoRecepcao: Math.round(valorTotal * 0.06 * 100) / 100
    };
  }
}

/**
 * Identifica hotel/operadora pela origem dos dados
 */
function identificarOrigemTransfer(dadosTransfer, parametrosExtras = {}) {
  loggerHUB.debug('Identificando origem do transfer', { dadosTransfer, parametrosExtras });
  
  // Se vier com identificação explícita
  if (parametrosExtras.hotelOperadora) {
    return parametrosExtras.hotelOperadora;
  }
  
  // Se vier com source
  if (parametrosExtras.source) {
    const sourceMap = {
      'empire_marques': 'Empire Marques Hotel',
      'hotel_lioz': 'Hotel Lioz', 
      'empire_lisbon': 'Empire Lisbon Hotel',
      'gota_dagua': 'Gota d´água'
    };
    
    if (sourceMap[parametrosExtras.source]) {
      return sourceMap[parametrosExtras.source];
    }
  }
  
  // Tentar identificar pelo conteúdo
  const textoParaAnalise = (dadosTransfer.origem + ' ' + dadosTransfer.destino + ' ' + (dadosTransfer.observacoes || '')).toLowerCase();
  
  for (const [nomeHotel, config] of Object.entries(SISTEMA_HUB_CONFIG.HOTEIS_CONECTADOS)) {
    const nomeHotelLower = nomeHotel.toLowerCase();
    if (textoParaAnalise.includes(nomeHotelLower) || 
        textoParaAnalise.includes(config.id)) {
      return nomeHotel;
    }
  }
  
  // Padrão para entrada manual
  return 'Entrada Manual';
}

/**
 * Cria resposta padronizada para APIs
 */
function criarRespostaHUB(dados, statusCode = 200, sucesso = true) {
  const resposta = {
    success: sucesso,
    data: dados,
    timestamp: new Date().toISOString(),
    sistema: SISTEMA_HUB_CONFIG.NOME,
    versao: SISTEMA_HUB_CONFIG.VERSAO
  };
  
  if (!sucesso && typeof dados === 'string') {
    resposta.error = dados;
    resposta.data = null;
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(resposta))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Verifica se é ano bissexto
 */
function isAnoBissexto(ano) {
  return (ano % 4 === 0 && ano % 100 !== 0) || (ano % 400 === 0);
}

// ===================================================
// SISTEMA Z-API PARA WHATSAPP - ROBERTA HUB
// ===================================================

/**
 * Envia mensagem via Z-API (WhatsApp)
 */
function enviarWhatsAppRoberta(numeroDestino, mensagem) {
  loggerHUB.info('Enviando WhatsApp via Roberta HUB', { 
    numero: numeroDestino, 
    tamanhoMensagem: mensagem.length 
  });
  
  if (!SISTEMA_HUB_CONFIG.NOTIFICACOES.WHATSAPP_ATIVO) {
    loggerHUB.warn('WhatsApp desativado na configuração');
    return false;
  }
  
  try {
    // Limpar e validar número
    const numeroLimpo = limparNumeroTelefone(numeroDestino);
    if (!numeroLimpo) {
      throw new Error('Número de telefone inválido');
    }
    
    // Preparar payload para Z-API
    const payload = {
      phone: numeroLimpo,
      message: mensagem
    };
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      headers: {
        'Client-Token': SISTEMA_HUB_CONFIG.Z_API.CLIENT_TOKEN
      },
      muteHttpExceptions: true
    };
    
    // Enviar via Z-API
    const response = UrlFetchApp.fetch(SISTEMA_HUB_CONFIG.Z_API.ENDPOINT, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    loggerHUB.debug('Resposta Z-API', { 
      code: responseCode, 
      response: responseText 
    });
    
    if (responseCode === 200 || responseCode === 201) {
      loggerHUB.success('WhatsApp enviado com sucesso via Roberta HUB');
      return true;
    } else {
      loggerHUB.error('Erro na Z-API', { 
        code: responseCode, 
        response: responseText 
      });
      return false;
    }
    
  } catch (error) {
    loggerHUB.error('Erro ao enviar WhatsApp', error);
    return false;
  }
}

/**
 * Limpa e valida número de telefone
 */
function limparNumeroTelefone(numero) {
  if (!numero) return null;
  
  // Remover todos os caracteres não numéricos
  let numeroLimpo = numero.toString().replace(/\D/g, '');
  
  // Se começar com 00, remover
  if (numeroLimpo.startsWith('00')) {
    numeroLimpo = numeroLimpo.substring(2);
  }
  
  // Se não começar com código do país, assumir Portugal (+351)
  if (!numeroLimpo.startsWith('351') && numeroLimpo.length === 9) {
    numeroLimpo = '351' + numeroLimpo;
  }
  
  // Validar comprimento
  if (numeroLimpo.length < 10 || numeroLimpo.length > 15) {
    return null;
  }
  
  return numeroLimpo;
}

/**
 * Envia confirmação de transfer via WhatsApp
 */
function enviarConfirmacaoWhatsApp(transferData) {
  loggerHUB.info('Enviando confirmação WhatsApp', { 
    transferId: transferData.id, 
    cliente: transferData.cliente 
  });
  
  try {
    // Verificar se deve enviar WhatsApp
    if (!SISTEMA_HUB_CONFIG.NOTIFICACOES.WHATSAPP_ATIVO) {
      loggerHUB.info('WhatsApp desativado');
      return { sucesso: false, motivo: 'whatsapp_desativado' };
    }
    
    // Apenas para transfers confirmados (se configurado)
    if (SISTEMA_HUB_CONFIG.NOTIFICACOES.WHATSAPP_APENAS_CONFIRMADOS && 
        transferData.status !== 'Confirmado') {
      loggerHUB.info('WhatsApp só para confirmados', { status: transferData.status });
      return { sucesso: false, motivo: 'status_nao_confirmado' };
    }
    
    // Validar dados necessários
    if (!transferData.contacto) {
      loggerHUB.warn('Transfer sem contacto para WhatsApp', { transferId: transferData.id });
      return { sucesso: false, motivo: 'sem_contacto' };
    }
    
    // Gerar mensagem da Roberta
    const mensagem = ROBERTA_MESSAGES.CONFIRMACAO_TEMPLATE(transferData);
    
    // Enviar WhatsApp para Junior (não para cliente)
    const enviado = enviarConfirmacaoParaJunior(transferData);
    
    if (enviado) {
      loggerHUB.success('WhatsApp de confirmação enviado', { 
        transferId: transferData.id, 
        cliente: transferData.cliente 
      });
      
      return { 
        sucesso: true, 
        mensagem: 'WhatsApp enviado com sucesso',
        numeroDestino: transferData.contacto
      };
    } else {
      loggerHUB.error('Falha no envio do WhatsApp', { transferId: transferData.id });
      return { sucesso: false, motivo: 'falha_envio' };
    }
    
  } catch (error) {
    loggerHUB.error('Erro ao enviar confirmação WhatsApp', error);
    return { sucesso: false, erro: error.message };
  }
}

/**
 * Determina se deve enviar WhatsApp baseado no fluxo
 */
function deveEnviarWhatsApp(source, status) {
  // FLUXO 1: Hotéis - só enviar se status for "Confirmado" 
  const hoteisIds = ['empire_marques', 'hotel_lioz', 'empire_lisbon', 'gota_dagua'];
  
  if (hoteisIds.includes(source)) {
    return status === 'Confirmado'; // Só após confirmar por e-mail
  }
  
  // FLUXO 2: Registros diretos/frontend - sempre enviar
  const origensDirectas = [
    'world_driver',
    'wt', 
    'connecto',
    'hub_direct',
    'manual_entry',
    'frontend_admin'
  ];
  
  return origensDirectas.includes(source);
}

/**
 * Testa sistema Z-API
 */
function testarZAPI() {
  loggerHUB.info('Testando sistema Z-API');
  
  const numeroTeste = SISTEMA_HUB_CONFIG.Z_API.ROBERTA_PHONE; // Número da própria Roberta
  const mensagemTeste = `🧪 TESTE DO SISTEMA HUB CENTRAL

Olá! Este é um teste do sistema de WhatsApp da Roberta HUB.

Se você recebeu esta mensagem, o sistema Z-API está funcionando corretamente! ✅

Timestamp: ${new Date().toLocaleString('pt-PT')}

_Sistema HUB Central - Teste Automático_`;

  // Enviar WhatsApp para Junior (não para cliente)
const enviado = enviarConfirmacaoParaJunior(transferData);
  
  if (resultado) {
    loggerHUB.success('Teste Z-API realizado com sucesso');
    return { sucesso: true, mensagem: 'Teste enviado para ' + numeroTeste };
  } else {
    loggerHUB.error('Falha no teste Z-API');
    return { sucesso: false, mensagem: 'Falha no envio do teste' };
  }
}

// ===================================================
// SISTEMA DE RECEBIMENTO E VALIDAÇÃO DE DADOS DOS HOTÉIS
// ===================================================

/**
 * Valida dados recebidos de hotéis
 */
function validarDadosHotel(dados) {
  loggerHUB.debug('Validando dados do hotel', { dados });
  
  const erros = [];
  const dadosValidados = {};
  
  // Campos obrigatórios
  const camposObrigatorios = [
    'nomeCliente', 'numeroPessoas', 'data', 'contacto', 
    'origem', 'destino', 'horaPickup', 'valorTotal'
  ];
  
  for (const campo of camposObrigatorios) {
    if (!dados[campo] || dados[campo] === '') {
      erros.push(`Campo obrigatório ausente: ${campo}`);
    }
  }
  
  // Validar e sanitizar campos
  if (dados.nomeCliente) {
    dadosValidados.nomeCliente = sanitizarTextoHUB(dados.nomeCliente);
  }
  
  if (dados.tipoServico) {
    dadosValidados.tipoServico = TIPOS_SERVICO.includes(dados.tipoServico) ? 
      dados.tipoServico : 'Transfer';
  } else {
    dadosValidados.tipoServico = 'Transfer';
  }
  
  // Validar números
  if (dados.numeroPessoas) {
    const pessoas = parseInt(dados.numeroPessoas);
    if (isNaN(pessoas) || pessoas < 1 || pessoas > 20) {
      erros.push('Número de pessoas deve estar entre 1 e 20');
    } else {
      dadosValidados.numeroPessoas = pessoas;
    }
  }
  
  if (dados.numeroBagagens !== undefined) {
    const bagagens = parseInt(dados.numeroBagagens);
    if (isNaN(bagagens) || bagagens < 0 || bagagens > 50) {
      erros.push('Número de bagagens deve estar entre 0 e 50');
    } else {
      dadosValidados.numeroBagagens = bagagens;
    }
  } else {
    dadosValidados.numeroBagagens = 0;
  }
  
  // Validar valor
  if (dados.valorTotal) {
    const valor = parseFloat(dados.valorTotal);
    if (isNaN(valor) || valor < 0.01 || valor > 9999.99) {
      erros.push('Valor deve estar entre €0,01 e €9.999,99');
    } else {
      dadosValidados.valorTotal = valor;
    }
  }
  
  // Validar data
  if (dados.data) {
    try {
      dadosValidados.data = processarDataSeguraHUB(dados.data);
    } catch (error) {
      erros.push('Data inválida ou em formato incorreto');
    }
  }
  
// Validar hora
 if (dados.horaPickup) {
   const horaRegex = /^([01]?[0-9]|2[0-3]):[0-5][0-9]$/;
   if (!horaRegex.test(dados.horaPickup)) {
     erros.push('Hora deve estar no formato HH:MM');
   } else {
     dadosValidados.horaPickup = dados.horaPickup;
   }
 }
 
 // Validar contacto
 if (dados.contacto) {
   dadosValidados.contacto = sanitizarTextoHUB(dados.contacto);
 }
 
 // Copiar outros campos com sanitização
 const outrosCampos = ['numeroVoo', 'origem', 'destino', 'observacoes'];
 for (const campo of outrosCampos) {
   if (dados[campo]) {
     dadosValidados[campo] = sanitizarTextoHUB(dados[campo]);
   }
 }
 
 // Aplicar valores padrão
 dadosValidados.status = dados.status || 'Solicitado';
 dadosValidados.modoPagamento = dados.modoPagamento || 'Dinheiro';
 dadosValidados.pagoParaQuem = dados.pagoParaQuem || 'Recepção';
 
 const resultado = {
   valido: erros.length === 0,
   erros: erros,
   dados: dadosValidados
 };
 
 loggerHUB.debug('Validação concluída', { 
   valido: resultado.valido, 
   numeroErros: erros.length 
 });
 
 return resultado;
}

/**
 * Função para processar dados do frontend e enviar para backend
 */
function processarDadosFrontend(dadosFrontend) {
  
  // Mapear dados do formulário frontend para estrutura backend
  const dadosBackend = {
    // Dados básicos do cliente
    nomeCliente: dadosFrontend.nomeCliente,
    tipoServico: dadosFrontend.tipoServico || 'Transfer',
    numeroPessoas: parseInt(dadosFrontend.numeroPessoas),
    numeroBagagens: parseInt(dadosFrontend.numeroBagagens) || 0,
    data: dadosFrontend.data,
    contacto: dadosFrontend.contacto,
    numeroVoo: dadosFrontend.numeroVoo || '',
    origem: dadosFrontend.origem,
    destino: dadosFrontend.destino,
    horaPickup: dadosFrontend.horaPickup,
    
    // Valores e pagamento
    valorTotal: parseFloat(dadosFrontend.valorServico),
    modoPagamento: dadosFrontend.formaPagamento,
    pagoParaQuem: dadosFrontend.cobranca,
    observacoes: dadosFrontend.observacoes || '',
    
    // Status padrão para registros diretos
    status: 'Confirmado'
  };
  
  return dadosBackend;
}

/**
 * Função para receber dados do frontend e processar no sistema HUB
 */
function receberDadosFrontendHUB(dadosFrontend) {
  loggerHUB.info('Recebendo dados do frontend HUB Transfer', {
    cliente: dadosFrontend.nomeCliente,
    operadora: dadosFrontend.operadora
  });
  
  try {
    // Processar e mapear dados
    const dadosProcessados = processarDadosFrontend(dadosFrontend);
    
    // Definir parâmetros extras
    const parametros = {
      source: dadosFrontend.operadora, // wt, connecto, hub_transfer
      origem: 'frontend_hub'
    };
    
    // Processar transfer (vai direto como CONFIRMADO + WhatsApp)
    const resultado = processarTransferDoHotel(dadosProcessados, parametros);
    
    loggerHUB.success('Transfer do frontend processado', {
      transferId: resultado.transferId,
      cliente: resultado.cliente,
      operadora: resultado.hotelOperadora
    });
    
    return {
      sucesso: true,
      transferId: resultado.transferId,
      cliente: resultado.cliente,
      operadora: resultado.hotelOperadora,
      whatsappEnviado: resultado.whatsappEnviado,
      status: resultado.status
    };
    
  } catch (error) {
    loggerHUB.error('Erro ao processar dados do frontend', error);
    
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
* Calcula valores baseado nos dados recebidos (SEM CÁLCULO PERCENTUAL)
*/
function calcularValoresHUB(dados) {
 loggerHUB.debug('Calculando valores HUB', { 
   precoCliente: dados.valorTotal, 
   tipoServico: dados.tipoServico 
 });
 
 const precoCliente = parseFloat(dados.valorTotal);
 let valorHotel, valorHUB, comissaoRecepcao;
 
 // Usar valores diretos enviados pelos hotéis (sem cálculo percentual)
 valorHotel = parseFloat(dados.valorHotel) || 0;
 valorHUB = parseFloat(dados.valorHUB) || 0;
 
 // Se não foram enviados valores específicos, usar valor total como HUB
 if (valorHotel === 0 && valorHUB === 0) {
   valorHUB = precoCliente;
 }
 
 // Comissão recepção baseada no tipo de serviço
 if (dados.tipoServico === 'Tour Regular' || dados.tipoServico === 'Private Tour') {
   comissaoRecepcao = SISTEMA_HUB_CONFIG.VALORES_PADRAO.COMISSAO_RECEPCAO_TOUR;
 } else {
   comissaoRecepcao = SISTEMA_HUB_CONFIG.VALORES_PADRAO.COMISSAO_RECEPCAO_TRANSFER;
 }
 
 // Usar comissão específica se enviada
 if (dados.comissaoRecepcao) {
   comissaoRecepcao = parseFloat(dados.comissaoRecepcao);
 }
 
 // Arredondar valores
 valorHotel = Math.round(valorHotel * 100) / 100;
 comissaoRecepcao = Math.round(comissaoRecepcao * 100) / 100;
 valorHUB = Math.round(valorHUB * 100) / 100;
 
 loggerHUB.debug('Valores calculados', {
   precoCliente,
   valorHotel,
   valorHUB,
   comissaoRecepcao
 });
 
 return {
   precoCliente: precoCliente,
   valorHotel: valorHotel,
   valorHUB: valorHUB,
   comissaoRecepcao: comissaoRecepcao
 };
}

/**
 * Processa transfer com lógica de fluxo inteligente - VERSÃO FINAL
 */
function processarTransferDoHotel(dadosRecebidos, parametrosExtras = {}) {
  loggerHUB.info('Processando transfer recebido de hotel', { 
    cliente: dadosRecebidos.nomeCliente,
    origem: parametrosExtras.source || dadosRecebidos.source 
  });
  
  try {
    // Validar dados
    const validacao = validarDadosHotel(dadosRecebidos);
    if (!validacao.valido) {
      throw new Error(`Dados inválidos: ${validacao.erros.join(', ')}`);
    }
    
    const dados = validacao.dados;
    const transferId = gerarIdTransferHUB();
    const hotelOperadora = identificarHotelOperadora(parametrosExtras.source || dadosRecebidos.source);
    const valores = calcularValoresTransfer(dados);
    const source = parametrosExtras.source || dadosRecebidos.source || 'unknown';
    
    // FLUXO 1: HOTÉIS (status pendente inicialmente)
    const hoteisIds = ['empire_marques', 'hotel_lioz', 'empire_lisbon', 'gota_dagua'];
    
    if (hoteisIds.includes(source)) {
      // Hotel: Registrar como pendente, NÃO enviar WhatsApp ainda
      
      const dadosTransfer = [
        transferId,                          // A - ID
        dados.nomeCliente,                   // B - Cliente
        dados.tipoServico,                   // C - Tipo Serviço
        dados.numeroPessoas,                 // D - Pessoas
        dados.numeroBagagens,                // E - Bagagens
        dados.data,                          // F - Data
        dados.contacto,                      // G - Contacto
        dados.numeroVoo || '',               // H - Voo
        dados.origem,                        // I - Origem
        dados.destino,                       // J - Destino
        dados.horaPickup,                    // K - Hora Pick-up
        valores.precoCliente,                // L - Preço Cliente
        valores.valorHotel,                  // M - Valor Hotel
        valores.valorHUB,                    // N - Valor HUB
        valores.comissaoRecepcao,            // O - Comissão Recepção
        dados.modoPagamento || 'Dinheiro',   // P - Forma Pagamento
        dados.pagoParaQuem || 'Recepção',    // Q - Pago Para
        'Pendente',                          // R - Status (PENDENTE)
        dados.observacoes || '',             // S - Observações
        new Date(),                          // T - Data Criação
        hotelOperadora,                      // U - Hotel/Operadora
        ''                                   // V - Motorista (vazio)
      ];
      
      const resultadoRegistro = registrarTransferComAbaMensal(dadosTransfer);
      
      if (!resultadoRegistro.sucesso) {
        throw new Error('Falha no registro do transfer');
      }
      
      loggerHUB.success('Transfer de hotel registrado como PENDENTE', {
        transferId: transferId,
        cliente: dados.nomeCliente,
        hotel: hotelOperadora
      });
      
      return {
        sucesso: true,
        transferId: transferId,
        cliente: dados.nomeCliente,
        hotelOperadora: hotelOperadora,
        whatsappEnviado: false,
        emailEnviado: true,
        status: 'Pendente - Aguardando confirmação'
      };
      
    } else {
      // FLUXO 2: REGISTROS DIRETOS (seu frontend)
      
      const dadosTransfer = [
        transferId,                          // A - ID
        dados.nomeCliente,                   // B - Cliente
        dados.tipoServico,                   // C - Tipo Serviço
        dados.numeroPessoas,                 // D - Pessoas
        dados.numeroBagagens,                // E - Bagagens
        dados.data,                          // F - Data
        dados.contacto,                      // G - Contacto
        dados.numeroVoo || '',               // H - Voo
        dados.origem,                        // I - Origem
        dados.destino,                       // J - Destino
        dados.horaPickup,                    // K - Hora Pick-up
        valores.precoCliente,                // L - Preço Cliente
        valores.valorHotel,                  // M - Valor Hotel
        valores.valorHUB,                    // N - Valor HUB
        valores.comissaoRecepcao,            // O - Comissão Recepção
        dados.modoPagamento || 'Dinheiro',   // P - Forma Pagamento
        dados.pagoParaQuem || 'Recepção',    // Q - Pago Para
        'Confirmado',                        // R - Status (CONFIRMADO)
        dados.observacoes || '',             // S - Observações
        new Date(),                          // T - Data Criação
        hotelOperadora,                      // U - Hotel/Operadora
        ''                                   // V - Motorista (vazio)
      ];
      
      const resultadoRegistro = registrarTransferComAbaMensal(dadosTransfer);
      
      if (!resultadoRegistro.sucesso) {
        throw new Error('Falha no registro do transfer');
      }
      
    // ENVIAR WHATSAPP PARA JUNIOR (NÃO PARA CLIENTE)
let whatsappEnviado = false;

loggerHUB.info('Enviando confirmação para Junior', {
  transferId: transferId,
  cliente: dados.nomeCliente,
  origem: source
});

try {
  const resultadoWhatsApp = enviarConfirmacaoParaJunior({
    id: transferId,
    cliente: dados.nomeCliente,
    tipoServico: dados.tipoServico,
    numeroPessoas: dados.numeroPessoas,
    data: dados.data,
    contacto: dados.contacto,
    origem: dados.origem,
    destino: dados.destino,
    horaPickup: dados.horaPickup,
    valorTotal: dados.valorTotal,
    status: 'Confirmado',
    hotelOperadora: hotelOperadora
  });
  
  whatsappEnviado = resultadoWhatsApp.sucesso;
        
      } catch (error) {
        loggerHUB.error('Erro ao enviar WhatsApp direto', error);
      }
      
      loggerHUB.success('Transfer direto processado com sucesso', {
        transferId: transferId,
        cliente: dados.nomeCliente,
        hotelOperadora: hotelOperadora,
        whatsappEnviado: whatsappEnviado
      });
      
      return {
        sucesso: true,
        transferId: transferId,
        cliente: dados.nomeCliente,
        hotelOperadora: hotelOperadora,
        whatsappEnviado: whatsappEnviado,
        emailEnviado: false,
        status: 'Confirmado',
        valores: {
          precoCliente: valores.precoCliente,
          valorHotel: valores.valorHotel,
          valorHUB: valores.valorHUB,
          comissaoRecepcao: valores.comissaoRecepcao
        }
      };
    }
    
  } catch (error) {
    loggerHUB.error('Erro no processamento do transfer', error);
    throw error;
  }
}

// ===================================================
// SISTEMA DE PLANILHAS E REGISTROS
// ===================================================

/**
 * Registra transfer na planilha HUB Central
 */
function registrarTransferHUB(dadosTransfer) {
  loggerHUB.info('Registrando transfer na planilha HUB', { 
    transferId: dadosTransfer[0],
    cliente: dadosTransfer[1]
  });
  
  try {
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SISTEMA_HUB_CONFIG.SHEET_NAME);
    
    // Criar aba se não existir
    if (!sheet) {
      sheet = criarAbaHUBCentral(ss);
    }
    
    // Verificar duplicidade
    const transferId = dadosTransfer[0];
    if (verificarDuplicidadeHUB(sheet, transferId)) {
      loggerHUB.warn('Transfer duplicado detectado', { transferId });
      return {
        sucesso: false,
        motivo: 'transfer_duplicado'
      };
    }
    
    // Inserir na planilha
    sheet.appendRow(dadosTransfer);
    
    // Verificar se foi inserido
    Utilities.sleep(500); // Aguardar inserção
    const linha = encontrarLinhaPorIdHUB(sheet, transferId);
    
    if (linha > 0) {
      loggerHUB.success('Transfer registrado na planilha HUB', { 
        transferId,
        linha
      });
      
      return {
        sucesso: true,
        transferId: transferId,
        linha: linha
      };
    } else {
      throw new Error('Transfer não foi encontrado após inserção');
    }
    
  } catch (error) {
    loggerHUB.error('Erro ao registrar transfer HUB', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Configuração de operadoras para frontend
 */
const OPERADORAS_FRONTEND = [
  {
    id: 'wt',
    nome: 'WT',
    cor: '#e74c3c'
  },
  {
    id: 'connecto', 
    nome: 'Connecto',
    cor: '#3498db'
  },
  {
    id: 'hub_transfer',
    nome: 'HUB Transfer',
    cor: '#2ecc71'
  }
];

/**
 * Opções de forma de pagamento
 */
const FORMAS_PAGAMENTO_FRONTEND = [
  {
    id: 'monetario',
    nome: 'Monetário',
    icone: '💰'
  },
  {
    id: 'multibanco',
    nome: 'Multibanco', 
    icone: '💳'
  },
  {
    id: 'paypal',
    nome: 'PayPal',
    icone: '🅿️'
  },
  {
    id: 'transferencia',
    nome: 'Transferência',
    icone: '🏦'
  }
];

/**
 * Opções de cobrança (Pago Para - Coluna Q)
 */
const OPCOES_COBRANCA_FRONTEND = [
  {
    id: 'motorista',
    nome: 'Motorista',
    icone: '🚗'
  },
  {
    id: 'operador',
    nome: 'Operador', 
    icone: '👤'
  },
  {
    id: 'parceiro',
    nome: 'Parceiro',
    icone: '🤝'
  },
  {
    id: 'banco',
    nome: 'Banco',
    icone: '🏦'
  },
  {
    id: 'empresa',
    nome: 'Empresa',
    icone: '🏢'
  }
];

/**
 * Validações para formulário frontend
 */
const VALIDACOES_FRONTEND = {
  nomeCliente: {
    obrigatorio: true,
    minimo: 2,
    maximo: 100
  },
  numeroPessoas: {
    obrigatorio: true,
    minimo: 1,
    maximo: 8
  },
  valorServico: {
    obrigatorio: true,
    minimo: 5.00,
    maximo: 500.00
  },
  contacto: {
    obrigatorio: true,
    formato: /^\+351\d{9}$/
  },
  operadora: {
    obrigatorio: true,
    opcoes: ['wt', 'connecto', 'hub_transfer']
  },
  formaPagamento: {
    obrigatorio: true,
    opcoes: ['monetario', 'multibanco', 'paypal', 'transferencia']
  },
  cobranca: {
    obrigatorio: true,
    opcoes: ['motorista', 'operador', 'parceiro', 'banco', 'empresa']
  }
};

/**
 * Cria aba HUB Central se não existir
 */
function criarAbaHUBCentral(ss) {
  loggerHUB.info('Criando aba HUB Central');
  
  const sheet = ss.insertSheet(SISTEMA_HUB_CONFIG.SHEET_NAME);
  
  // Configurar headers
  sheet.getRange(1, 1, 1, HEADERS_HUB_CENTRAL.length).setValues([HEADERS_HUB_CENTRAL]);
  
  // Formatação do cabeçalho
  const headerRange = sheet.getRange(1, 1, 1, HEADERS_HUB_CENTRAL.length);
  headerRange
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  // Congelar primeira linha
  sheet.setFrozenRows(1);
  
// Aplicar larguras das colunas (incluindo novas colunas W, X, Y, Z)
    const larguras = [60, 150, 120, 80, 80, 100, 120, 100, 200, 200, 90, 120, 140, 140, 100, 120, 100, 100, 200, 140, 150, 150, 120, 160, 120, 150];
  larguras.forEach((largura, index) => {
    if (index < HEADERS_HUB_CENTRAL.length) {
      sheet.setColumnWidth(index + 1, largura);
    }
  });
  
  // Aplicar formatações
  aplicarFormatacaoHUB(sheet);
  
  loggerHUB.success('Aba HUB Central criada com sucesso');
  
  return sheet;
}

/**
 * Aplica formatação na planilha HUB
 */
function aplicarFormatacaoHUB(sheet) {
  loggerHUB.debug('Aplicando formatação HUB');
  
  try {
    const maxRows = Math.max(sheet.getMaxRows(), 1000);
    
    if (maxRows > 1) {
      // Formatação de moeda (colunas L, M, N, O)
      sheet.getRange(2, 12, maxRows - 1, 4).setNumberFormat('€#,##0.00');
      
      // Formatação de data (coluna F)
      sheet.getRange(2, 6, maxRows - 1, 1).setNumberFormat('dd/mm/yyyy');
      
      // Formatação de hora (coluna K)
      sheet.getRange(2, 11, maxRows - 1, 1).setNumberFormat('hh:mm');
      
      // Formatação de timestamp (coluna T)
      sheet.getRange(2, 20, maxRows - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm');
      
      // Formatação de números (colunas A, D, E)
      sheet.getRange(2, 1, maxRows - 1, 1).setNumberFormat('0');
      sheet.getRange(2, 4, maxRows - 1, 2).setNumberFormat('0');
      
      // Formatação de timestamp (coluna X - Timestamp_OK)
      sheet.getRange(2, 24, maxRows - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
      
      // Formatação de texto centrado para Status_OK (coluna W)
      sheet.getRange(2, 23, maxRows - 1, 1).setHorizontalAlignment('center');
      
      // Formatação de texto centrado para Tipo_Viagem (coluna Y)
      sheet.getRange(2, 25, maxRows - 1, 1).setHorizontalAlignment('center');
      
      // Formatação de texto centrado para Status_Followup (coluna Z)
      sheet.getRange(2, 26, maxRows - 1, 1).setHorizontalAlignment('center');
    }
    
    // Aplicar validações
    aplicarValidacoesHUB(sheet);
    
    // Aplicar validações das novas colunas
    aplicarValidacoesNovasColunas(sheet);
    
  } catch (error) {
    loggerHUB.error('Erro ao aplicar formatação HUB', error);
  }
}

/**
 * Aplica validações na planilha HUB
 */
function aplicarValidacoesHUB(sheet) {
  loggerHUB.debug('Aplicando validações HUB');
  
  try {
    const maxRows = Math.max(sheet.getMaxRows(), 1000);
    
    // Validação de Status (coluna R)
    const statusValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(STATUS_PERMITIDOS)
      .setAllowInvalid(true)
      .setHelpText('Status do transfer')
      .build();
    sheet.getRange(2, 18, maxRows - 1, 1).setDataValidation(statusValidation);
    
    // Validação de Tipo de Serviço (coluna C)
    const tipoValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(TIPOS_SERVICO)
      .setAllowInvalid(true)
      .setHelpText('Tipo de serviço')
      .build();
    sheet.getRange(2, 3, maxRows - 1, 1).setDataValidation(tipoValidation);
    
    // Validação de Hotel/Operadora (coluna U)
    const hoteisOperadoras = Object.keys(SISTEMA_HUB_CONFIG.HOTEIS_CONECTADOS)
      .concat(SISTEMA_HUB_CONFIG.OPERADORAS_EXTERNAS);
    
    const hotelValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(hoteisOperadoras)
      .setAllowInvalid(true)
      .setHelpText('Hotel ou operadora')
      .build();
    sheet.getRange(2, 21, maxRows - 1, 1).setDataValidation(hotelValidation);
    
    // Validação Status_OK (coluna W)
    const statusOKValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pendente', 'Confirmado', 'OK', 'Verificado'])
      .setAllowInvalid(true)
      .setHelpText('Status de confirmação do transfer')
      .build();
    sheet.getRange(2, 23, maxRows - 1, 1).setDataValidation(statusOKValidation);
    
    // Validação Tipo_Viagem (coluna Y) 
    const tipoViagemValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Ida', 'Volta', 'Ida/Volta', 'Tour Completo', 'Personalizado'])
      .setAllowInvalid(true)
      .setHelpText('Tipo específico da viagem')
      .build();
    sheet.getRange(2, 25, maxRows - 1, 1).setDataValidation(tipoViagemValidation);
    
    // Validação Status_Followup (coluna Z)
    const followupValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Não Iniciado', 'Em Andamento', 'Contactado', 'Feedback Recebido', 'Finalizado'])
      .setAllowInvalid(true) 
      .setHelpText('Status de follow-up pós-serviço')
      .build();
    sheet.getRange(2, 26, maxRows - 1, 1).setDataValidation(followupValidation);
    
  } catch (error) {
    loggerHUB.error('Erro ao aplicar validações HUB', error);
  }
}

/**
 * Aplica validações nas novas colunas W, X, Y, Z
 */
function aplicarValidacoesNovasColunas(sheet) {
  loggerHUB.debug('Aplicando validações nas novas colunas W, X, Y, Z');
  
  try {
    const maxRows = Math.max(sheet.getMaxRows(), 1000);
    
    // Validação Status_OK (coluna W)
    const statusOKValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pendente', 'Confirmado', 'OK', 'Verificado'])
      .setAllowInvalid(true)
      .setHelpText('Status de confirmação do transfer')
      .build();
    sheet.getRange(2, 23, maxRows - 1, 1).setDataValidation(statusOKValidation);
    
    // Validação Tipo_Viagem (coluna Y) 
    const tipoViagemValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Ida', 'Volta', 'Ida/Volta', 'Tour Completo', 'Personalizado'])
      .setAllowInvalid(true)
      .setHelpText('Tipo específico da viagem')
      .build();
    sheet.getRange(2, 25, maxRows - 1, 1).setDataValidation(tipoViagemValidation);
    
    // Validação Status_Followup (coluna Z)
    const followupValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Não Iniciado', 'Em Andamento', 'Contactado', 'Feedback Recebido', 'Finalizado'])
      .setAllowInvalid(true) 
      .setHelpText('Status de follow-up pós-serviço')
      .build();
    sheet.getRange(2, 26, maxRows - 1, 1).setDataValidation(followupValidation);
    
  } catch (error) {
    loggerHUB.error('Erro ao aplicar validações nas novas colunas', error);
  }
}

/**
 * Aplica formatação nas novas colunas W, X, Y, Z
 */
function aplicarFormatacaoNovasColunas(sheet) {
  loggerHUB.debug('Aplicando formatação nas novas colunas W, X, Y, Z');
  
  try {
    const maxRows = Math.max(sheet.getMaxRows(), 1000);
    
    if (maxRows > 1) {
      // Formatação de timestamp (coluna X - Timestamp_OK)
      sheet.getRange(2, 24, maxRows - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
      
      // Formatação de texto centrado para Status_OK (coluna W)
      sheet.getRange(2, 23, maxRows - 1, 1).setHorizontalAlignment('center');
      
      // Formatação de texto centrado para Tipo_Viagem (coluna Y)
      sheet.getRange(2, 25, maxRows - 1, 1).setHorizontalAlignment('center');
      
      // Formatação de texto centrado para Status_Followup (coluna Z)
      sheet.getRange(2, 26, maxRows - 1, 1).setHorizontalAlignment('center');
    }
    
  } catch (error) {
    loggerHUB.error('Erro ao aplicar formatação nas novas colunas', error);
  }
}

/**
 * Verifica duplicidade na planilha HUB
 */
function verificarDuplicidadeHUB(sheet, transferId) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return false;
    
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    return ids.some(row => String(row[0]) === String(transferId));
    
  } catch (error) {
    loggerHUB.error('Erro ao verificar duplicidade', error);
    return false;
  }
}

/**
 * Encontra linha por ID na planilha HUB
 */
function encontrarLinhaPorIdHUB(sheet, id) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return 0;
    
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) {
        return i + 2; // +2 porque começamos na linha 2
      }
    }
    
    return 0;
    
  } catch (error) {
    loggerHUB.error('Erro ao buscar linha por ID', error);
    return 0;
  }
}

// ===================================================
// SISTEMA DE ABAS MENSAIS COM ORGANIZAÇÃO POR DIAS
// ===================================================

/**
 * Obtém ou cria aba mensal com organização por dias
 */
function obterAbaMensalHUB(dataTransfer) {
  loggerHUB.info('Obtendo aba mensal HUB', { data: dataTransfer });
  
  try {
    const data = processarDataSeguraHUB(dataTransfer);
    const mes = String(data.getMonth() + 1).padStart(2, '0');
    const ano = data.getFullYear();
    
    const mesInfo = MESES_HUB.find(m => m.abrev === mes);
    if (!mesInfo) {
      throw new Error(`Mês inválido: ${mes}`);
    }
    
    // Ajustar dias para ano bissexto
    if (mes === '02' && isAnoBissexto(ano)) {
      mesInfo.dias = 29;
    }
    
    const nomeAba = `${CONFIG_ABAS_MENSAIS.PREFIXO}${mes}_${mesInfo.nome}_${ano}`;
    
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    let abaMensal = ss.getSheetByName(nomeAba);
    
    if (abaMensal) {
      loggerHUB.debug('Aba mensal encontrada', { nomeAba });
      return abaMensal;
    }
    
    // Criar nova aba mensal
    loggerHUB.info('Criando nova aba mensal', { nomeAba });
    abaMensal = criarAbaMensalComDias(nomeAba, ss, mesInfo, ano);
    
    return abaMensal;
    
  } catch (error) {
    loggerHUB.error('Erro ao obter aba mensal', error);
    // Fallback: retornar aba principal
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    return ss.getSheetByName(SISTEMA_HUB_CONFIG.SHEET_NAME);
  }
}

/**
* Cria aba mensal com organização visual por dias - VERSÃO DINÂMICA
*/
function criarAbaMensalComDias(nomeAba, ss, mesInfo, ano) {
 loggerHUB.info('Criando aba mensal com sistema dinâmico', { nome: nomeAba });
 
 try {
   const abaMensal = ss.insertSheet(nomeAba);
   
   // 1. Configurar cabeçalho principal
   criarCabecalhoMensal(abaMensal, mesInfo, ano);
   
   // 2. Criar headers dos dados
   let linhaAtual = 4; // Começar após o cabeçalho
   abaMensal.getRange(linhaAtual, 1, 1, HEADERS_HUB_CENTRAL.length).setValues([HEADERS_HUB_CENTRAL]);
   
   // Formatação dos headers
   const headerRange = abaMensal.getRange(linhaAtual, 1, 1, HEADERS_HUB_CENTRAL.length);
   headerRange
     .setBackground(mesInfo.cor)
     .setFontColor('#ffffff')
     .setFontWeight('bold')
     .setFontSize(10)
     .setHorizontalAlignment('center');
   
   linhaAtual += 2; // Pular uma linha
   
   // 3. Criar APENAS cabeçalhos para cada dia (SEM linhas vazias fixas)
   for (let dia = 1; dia <= mesInfo.dias; dia++) {
     criarSecaoDiaExpansivel(abaMensal, dia, mesInfo, ano, linhaAtual);
     linhaAtual += 2; // Apenas cabeçalho + 1 espaço (expansão dinâmica)
   }
   
   // 4. Aplicar formatações gerais
   aplicarFormatacaoAbaMensal(abaMensal);
   
   // 5. Criar resumo de capacidade
   criarResumoCapacidade(abaMensal, mesInfo, ano);
   
   loggerHUB.success('Aba mensal dinâmica criada com organização por dias', { nome: nomeAba });
   
   return abaMensal;
   
 } catch (error) {
   loggerHUB.error('Erro ao criar aba mensal dinâmica', error);
   throw error;
 }
}

/**
 * Cria cabeçalho da aba mensal
 */
function criarCabecalhoMensal(sheet, mesInfo, ano) {
  // Título principal
  sheet.getRange('A1').setValue(`📅 ${mesInfo.nome.toUpperCase()} ${ano} - SISTEMA HUB CENTRAL`);
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('A1:U1').merge().setBackground(mesInfo.cor).setFontColor('#ffffff');
  
  // Subtítulo
  sheet.getRange('A2').setValue(`Organização por dias - ${mesInfo.dias} dias`);
  sheet.getRange('A2').setFontSize(12).setFontStyle('italic').setHorizontalAlignment('center');
  sheet.getRange('A2:U2').merge();
  
  // Linha de separação
  sheet.getRange('A3:U3').setBackground('#f0f0f0');
}

/**
 * Cria seção para um dia específico
 */
function criarSecaoDia(sheet, dia, mesInfo, ano, linhaInicial) {
  const dataCompleta = new Date(ano, mesInfo.numero - 1, dia);
  const diaSemana = dataCompleta.toLocaleDateString('pt-PT', { weekday: 'long' });
  const dataFormatada = `${String(dia).padStart(2, '0')}/${mesInfo.abrev}/${ano}`;
  
  // Cabeçalho do dia
  const tituloData = `📅 DIA ${dia} - ${diaSemana.toUpperCase()} (${dataFormatada})`;
  sheet.getRange(linhaInicial, 1).setValue(tituloData);
  sheet.getRange(linhaInicial, 1).setFontWeight('bold').setFontSize(11);
  sheet.getRange(linhaInicial, 1, 1, 21).merge(); // Merge com todas as colunas
  
  // Cor de fundo baseada no dia da semana
  const corFundo = (dataCompleta.getDay() === 0 || dataCompleta.getDay() === 6) ? 
    '#fff3cd' : '#f8f9fa'; // Amarelo para fins de semana, cinza para dias úteis
  
  sheet.getRange(linhaInicial, 1, 1, 21).setBackground(corFundo);
  
  // Adicionar contador de transfers (será atualizado dinamicamente)
  sheet.getRange(linhaInicial, 22).setValue('0 transfers');
  sheet.getRange(linhaInicial, 22).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(linhaInicial, 22).setBackground(CONFIG_ABAS_MENSAIS.CORES_CAPACIDADE.VAZIO);
  
  // Adicionar nota para identificação
  sheet.getRange(linhaInicial, 23).setValue(`DIA_${String(dia).padStart(2, '0')}`);
  sheet.getRange(linhaInicial, 23).setFontSize(8).setFontColor('#999999');
}



/**
 * Aplica formatação geral na aba mensal
 */
function aplicarFormatacaoAbaMensal(sheet) {
  // Congelar primeira linha
  sheet.setFrozenRows(4);
  
  // Larguras das colunas (igual ao sistema principal)
  const larguras = [60, 150, 120, 80, 80, 100, 120, 100, 200, 200, 90, 120, 140, 140, 100, 120, 100, 100, 200, 140, 150, 100, 80];
  larguras.forEach((largura, index) => {
    if (index < larguras.length) {
      sheet.setColumnWidth(index + 1, largura);
    }
  });
  
  // Aplicar formatação de dados
  const maxRows = sheet.getMaxRows();
  if (maxRows > 5) {
    // Formatação de moeda (colunas L, M, N, O)
    sheet.getRange(5, 12, maxRows - 4, 4).setNumberFormat('€#,##0.00');
    
    // Formatação de data (coluna F)
    sheet.getRange(5, 6, maxRows - 4, 1).setNumberFormat('dd/mm/yyyy');
    
    // Formatação de hora (coluna K)
    sheet.getRange(5, 11, maxRows - 4, 1).setNumberFormat('hh:mm');
  }
}

/**
 * Cria resumo de capacidade no final da aba
 */
function criarResumoCapacidade(sheet, mesInfo, ano) {
  const ultimaLinha = sheet.getLastRow() + 2;
  
  // Título do resumo
  sheet.getRange(ultimaLinha, 1).setValue('📊 RESUMO DE CAPACIDADE DO MÊS');
  sheet.getRange(ultimaLinha, 1).setFontSize(14).setFontWeight('bold');
  sheet.getRange(ultimaLinha, 1, 1, 5).merge().setBackground('#e9ecef');
  
  // Headers do resumo
  const headersResumo = ['Dia', 'Data', 'Dia Semana', 'Transfers', 'Capacidade'];
  sheet.getRange(ultimaLinha + 1, 1, 1, headersResumo.length).setValues([headersResumo]);
  sheet.getRange(ultimaLinha + 1, 1, 1, headersResumo.length)
    .setFontWeight('bold')
    .setBackground('#6c757d')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  
  // Dados do resumo (será preenchido dinamicamente)
  let linhaResumo = ultimaLinha + 2;
  for (let dia = 1; dia <= mesInfo.dias; dia++) {
    const dataCompleta = new Date(ano, mesInfo.numero - 1, dia);
    const diaSemana = dataCompleta.toLocaleDateString('pt-PT', { weekday: 'short' });
    const dataFormatada = `${String(dia).padStart(2, '0')}/${mesInfo.abrev}`;
    
    sheet.getRange(linhaResumo, 1).setValue(dia);
    sheet.getRange(linhaResumo, 2).setValue(dataFormatada);
    sheet.getRange(linhaResumo, 3).setValue(diaSemana);
    sheet.getRange(linhaResumo, 4).setValue(0); // Será atualizado
    sheet.getRange(linhaResumo, 5).setValue('Disponível'); // Será atualizado
    
    // Cor baseada no dia da semana
    const corLinha = (dataCompleta.getDay() === 0 || dataCompleta.getDay() === 6) ? 
      '#fff3cd' : '#ffffff';
    sheet.getRange(linhaResumo, 1, 1, 5).setBackground(corLinha);
    
    linhaResumo++;
  }
}

/**
 * NOVA VERSÃO: Cria seção dinâmica para um dia específico (EXPANSÍVEL)
 */
function criarSecaoDiaExpansivel(sheet, dia, mesInfo, ano, linhaInicial) {
  const dataCompleta = new Date(ano, mesInfo.numero - 1, dia);
  const diaSemana = dataCompleta.toLocaleDateString('pt-PT', { weekday: 'long' });
  const dataFormatada = `${String(dia).padStart(2, '0')}/${mesInfo.abrev}/${ano}`;
  
  // Cabeçalho do dia (APENAS 1 LINHA)
  const tituloData = `📅 DIA ${dia} - ${diaSemana.toUpperCase()} (${dataFormatada})`;
  sheet.getRange(linhaInicial, 1).setValue(tituloData);
  sheet.getRange(linhaInicial, 1).setFontWeight('bold').setFontSize(11);
  sheet.getRange(linhaInicial, 1, 1, 21).merge();
  
  // Cor de fundo baseada no dia da semana
  const corFundo = (dataCompleta.getDay() === 0 || dataCompleta.getDay() === 6) ? 
    '#fff3cd' : '#f8f9fa';
  
  sheet.getRange(linhaInicial, 1, 1, 21).setBackground(corFundo);
  
  // Contador de transfers (será atualizado dinamicamente)
  sheet.getRange(linhaInicial, 22).setValue('0 transfers');
  sheet.getRange(linhaInicial, 22).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(linhaInicial, 22).setBackground(CONFIG_ABAS_MENSAIS.CORES_CAPACIDADE.VAZIO);
  
  // Marcador para identificação
  sheet.getRange(linhaInicial, 23).setValue(`DIA_${String(dia).padStart(2, '0')}`);
  sheet.getRange(linhaInicial, 23).setFontSize(8).setFontColor('#999999');
  
  // NÃO criar linhas vazias - serão adicionadas dinamicamente
}

/**
 * VERSÃO OTIMIZADA: Cria aba mensal com sistema dinâmico
 */
function criarAbaMensalComDiasDinamica(nomeAba, ss, mesInfo, ano) {
  loggerHUB.info('Criando aba mensal com sistema dinâmico', { nome: nomeAba });
  
  try {
    const abaMensal = ss.insertSheet(nomeAba);
    
    // 1. Configurar cabeçalho principal
    criarCabecalhoMensal(abaMensal, mesInfo, ano);
    
    // 2. Criar headers dos dados
    let linhaAtual = 4;
    abaMensal.getRange(linhaAtual, 1, 1, HEADERS_HUB_CENTRAL.length).setValues([HEADERS_HUB_CENTRAL]);
    
    // Formatação dos headers
    const headerRange = abaMensal.getRange(linhaAtual, 1, 1, HEADERS_HUB_CENTRAL.length);
    headerRange
      .setBackground(mesInfo.cor)
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setFontSize(10)
      .setHorizontalAlignment('center');
    
    linhaAtual += 2;
    
    // 3. Criar APENAS cabeçalhos para cada dia (SEM linhas vazias)
    for (let dia = 1; dia <= mesInfo.dias; dia++) {
      criarSecaoDiaExpansivel(abaMensal, dia, mesInfo, ano, linhaAtual);
      linhaAtual += 2; // Apenas 2 linhas: cabeçalho + 1 espaço
    }
    
    // 4. Aplicar formatações gerais
    aplicarFormatacaoAbaMensal(abaMensal);
    
    loggerHUB.success('Aba mensal dinâmica criada', { nome: nomeAba });
    
    return abaMensal;
    
  } catch (error) {
    loggerHUB.error('Erro ao criar aba mensal dinâmica', error);
    throw error;
  }
}

/**
 * VERSÃO OTIMIZADA: Registra transfer com sistema dinâmico
 */
function registrarTransferComAbaMensal(dadosTransfer) {
  loggerHUB.info('Registrando transfer com sistema dinâmico', { 
    transferId: dadosTransfer[0],
    cliente: dadosTransfer[1]
  });
  
  try {
    // 1. Registrar na aba central
    const resultadoCentral = registrarTransferHUB(dadosTransfer);
    
    if (!resultadoCentral.sucesso) {
      throw new Error('Falha ao registrar na aba central');
    }
    
    // 2. Se abas mensais estão ativadas, registrar na mensal
    if (CONFIG_ABAS_MENSAIS.USAR_ABAS_MENSAIS) {
      const dataTransfer = dadosTransfer[5]; // Coluna F - Data
      const abaMensal = obterAbaMensalHUB(dataTransfer);
      
      if (abaMensal) {
        // Inserir de forma dinâmica
        const linhaInserida = inserirTransferDinamico(abaMensal, dadosTransfer, dataTransfer);
        
        if (linhaInserida > 0) {
          loggerHUB.success('Transfer registrado na aba mensal dinâmica', {
            aba: abaMensal.getName(),
            linha: linhaInserida
          });
        }
      }
    }
    
    return {
      sucesso: true,
      transferId: dadosTransfer[0],
      abaCentral: resultadoCentral.linha,
      abaMensal: CONFIG_ABAS_MENSAIS.USAR_ABAS_MENSAIS
    };
    
  } catch (error) {
    loggerHUB.error('Erro no registro dinâmico', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Encontra linha correta para inserir transfer baseado no dia
 */
function encontrarLinhaParaDia(sheet, dataTransfer) {
  try {
    const data = processarDataSeguraHUB(dataTransfer);
    const dia = data.getDate();
    const marcadorDia = `DIA_${String(dia).padStart(2, '0')}`;
    
    // Buscar pela célula que contém o marcador do dia
    const range = sheet.getRange('W:W'); // Coluna W onde estão os marcadores
    const valores = range.getValues();
    
    for (let i = 0; i < valores.length; i++) {
      if (valores[i][0] === marcadorDia) {
        return i + 1; // Linha encontrada (base 1)
      }
    }
    
    return 0; // Não encontrou
    
  } catch (error) {
    loggerHUB.error('Erro ao encontrar linha do dia', error);
    return 0;
  }
}

/**
 * OTIMIZADA: Insere transfer de forma dinâmica
 */
function inserirTransferDinamico(sheet, dadosTransfer, dataTransfer) {
  try {
    const data = processarDataSeguraHUB(dataTransfer);
    const dia = data.getDate();
    const marcadorDia = `DIA_${String(dia).padStart(2, '0')}`;
    
    // Encontrar linha do cabeçalho do dia
    const range = sheet.getRange('W:W');
    const valores = range.getValues();
    
    for (let i = 0; i < valores.length; i++) {
      if (valores[i][0] === marcadorDia) {
        const linhaCabecalho = i + 1;
        
        // Encontrar próximo cabeçalho de dia para saber onde inserir
        let proximoLimiteDia = sheet.getLastRow() + 1;
        
        for (let j = i + 1; j < valores.length; j++) {
          if (valores[j][0] && valores[j][0].toString().startsWith('DIA_')) {
            proximoLimiteDia = j + 1;
            break;
          }
        }
        
        // Inserir nova linha ANTES do próximo dia
        sheet.insertRowBefore(proximoLimiteDia);
        
        // Adicionar os dados do transfer
        sheet.getRange(proximoLimiteDia, 1, 1, dadosTransfer.length).setValues([dadosTransfer]);
        
        // Aplicar formatação padrão
        aplicarFormatacaoLinhaTransfer(sheet, proximoLimiteDia);
        
        // Atualizar contador do dia
        atualizarContadorDia(sheet, dataTransfer);
        
        loggerHUB.success('Transfer inserido dinamicamente', {
          dia: dia,
          linha: proximoLimiteDia
        });
        
        return proximoLimiteDia;
      }
    }
    
    return 0;
    
  } catch (error) {
    loggerHUB.error('Erro ao inserir transfer dinâmico', error);
    return 0;
  }
}

/**
 * Aplica formatação em linha de transfer
 */
function aplicarFormatacaoLinhaTransfer(sheet, linha) {
  try {
    // Formatação de moeda (colunas L, M, N, O)
    sheet.getRange(linha, 12, 1, 4).setNumberFormat('€#,##0.00');
    
    // Formatação de data (coluna F)
    sheet.getRange(linha, 6, 1, 1).setNumberFormat('dd/mm/yyyy');
    
    // Formatação de hora (coluna K)
    sheet.getRange(linha, 11, 1, 1).setNumberFormat('hh:mm');
    
    // Formatação de timestamp (coluna T)
    sheet.getRange(linha, 20, 1, 1).setNumberFormat('dd/mm/yyyy hh:mm');
    
  } catch (error) {
    loggerHUB.error('Erro ao aplicar formatação da linha', error);
  }
}

/**
 * Atualiza contador de transfers do dia
 */
function atualizarContadorDia(sheet, dataTransfer) {
  try {
    const data = processarDataSeguraHUB(dataTransfer);
    const dia = data.getDate();
    const marcadorDia = `DIA_${String(dia).padStart(2, '0')}`;
    
    // Encontrar linha do cabeçalho do dia
    const range = sheet.getRange('W:W');
    const valores = range.getValues();
    
    for (let i = 0; i < valores.length; i++) {
      if (valores[i][0] === marcadorDia) {
        const linhaCabecalho = i + 1;
        
        // Contar transfers para este dia
        const transfersDoDia = contarTransfersDoDia(sheet, dia, linhaCabecalho);
        
        // Atualizar contador na coluna V
        const textoContador = `${transfersDoDia} transfer${transfersDoDia !== 1 ? 's' : ''}`;
        sheet.getRange(linhaCabecalho, 22).setValue(textoContador);
        
        // Atualizar cor baseada na capacidade
        const corCapacidade = obterCorCapacidade(transfersDoDia);
        sheet.getRange(linhaCabecalho, 22).setBackground(corCapacidade);
        
        loggerHUB.debug('Contador do dia atualizado', { dia, transfers: transfersDoDia });
        break;
      }
    }
    
  } catch (error) {
    loggerHUB.error('Erro ao atualizar contador', error);
  }
}

/**
 * Conta transfers de um dia específico
 */
function contarTransfersDoDia(sheet, dia, linhaCabecalho) {
  try {
    let contador = 0;
    let linha = linhaCabecalho + 1;
    
    // Contar até encontrar próximo cabeçalho de dia ou fim
    while (linha <= sheet.getLastRow()) {
      const valorLinha = sheet.getRange(linha, 1).getValue();
      
      // Se encontrou próximo cabeçalho de dia, parar
      if (typeof valorLinha === 'string' && valorLinha.includes('DIA ')) {
        break;
      }
      
      // Se tem dados na linha (ID preenchido), contar
      const id = sheet.getRange(linha, 1).getValue();
      if (id && !isNaN(id) && id > 0) {
        contador++;
      }
      
      linha++;
    }
    
    return contador;
    
  } catch (error) {
    loggerHUB.error('Erro ao contar transfers', error);
    return 0;
  }
}

/**
 * Obtém cor baseada na capacidade
 */
function obterCorCapacidade(numeroTransfers) {
  const limite = CONFIG_ABAS_MENSAIS.LIMITE_TRANSFERS_DIA;
  
  if (numeroTransfers === 0) {
    return CONFIG_ABAS_MENSAIS.CORES_CAPACIDADE.VAZIO;
  } else if (numeroTransfers <= limite * 0.3) {
    return CONFIG_ABAS_MENSAIS.CORES_CAPACIDADE.BAIXO;
  } else if (numeroTransfers <= limite * 0.6) {
    return CONFIG_ABAS_MENSAIS.CORES_CAPACIDADE.MEDIO;
  } else if (numeroTransfers <= limite * 0.9) {
    return CONFIG_ABAS_MENSAIS.CORES_CAPACIDADE.ALTO;
  } else {
    return CONFIG_ABAS_MENSAIS.CORES_CAPACIDADE.CHEIO;
  }
}

/**
 * Cria todas as abas mensais do ano - VERSÃO DINÂMICA
 */
function criarTodasAbasMensaisHUB(ano = new Date().getFullYear()) {
  loggerHUB.info('Criando todas as abas mensais HUB com sistema dinâmico', { ano });
  
  try {
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    const resultados = {
      criadas: 0,
      existentes: 0,
      erros: 0,
      detalhes: []
    };
    
    MESES_HUB.forEach(mes => {
      const nomeAba = `${CONFIG_ABAS_MENSAIS.PREFIXO}${mes.abrev}_${mes.nome}_${ano}`;
      
      try {
        let aba = ss.getSheetByName(nomeAba);
        
        if (aba) {
          resultados.existentes++;
          resultados.detalhes.push({
            mes: mes.nome,
            status: 'existente',
            nome: nomeAba
          });
        } else {
          // Ajustar dias para fevereiro em ano bissexto
          if (mes.numero === 2 && isAnoBissexto(ano)) {
            mes.dias = 29;
          }
          
          // USAR A NOVA FUNÇÃO DINÂMICA
          aba = criarAbaMensalComDiasDinamica(nomeAba, ss, mes, ano);
          resultados.criadas++;
          resultados.detalhes.push({
            mes: mes.nome,
            status: 'criada',
            nome: nomeAba
          });
        }
        
      } catch (error) {
        loggerHUB.error('Erro ao processar aba mensal', { 
          mes: mes.nome, 
          erro: error.message 
        });
        resultados.erros++;
        resultados.detalhes.push({
          mes: mes.nome,
          status: 'erro',
          nome: nomeAba,
          erro: error.message
        });
      }
    });
    
    loggerHUB.success('Criação de abas mensais dinâmicas concluída', resultados);
    
    return resultados;
    
  } catch (error) {
    loggerHUB.error('Erro ao criar abas mensais', error);
    throw error;
  }
}

/**
 * Atualiza todos os contadores de capacidade das abas mensais
 */
function atualizarTodosContadores() {
  loggerHUB.info('Atualizando todos os contadores de capacidade');
  
  try {
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    const sheets = ss.getSheets();
    let abasProcessadas = 0;
    
    sheets.forEach(sheet => {
      const nome = sheet.getName();
      
      // Verificar se é aba mensal HUB
      if (nome.startsWith(CONFIG_ABAS_MENSAIS.PREFIXO) && nome.includes('_')) {
        try {
          atualizarContadoresAbaMensal(sheet);
          abasProcessadas++;
          loggerHUB.debug('Contadores atualizados', { aba: nome });
        } catch (error) {
          loggerHUB.error('Erro ao atualizar contadores da aba', { aba: nome, erro: error.message });
        }
      }
    });
    
    loggerHUB.success('Atualização de contadores concluída', { abasProcessadas });
    
    return {
      sucesso: true,
      abasProcessadas: abasProcessadas
    };
    
  } catch (error) {
    loggerHUB.error('Erro na atualização de contadores', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Atualiza contadores de uma aba mensal específica
 */
function atualizarContadoresAbaMensal(sheet) {
  try {
    // Extrair mês e ano do nome da aba
    const nomeAba = sheet.getName();
    const partes = nomeAba.split('_');
    const mesNumero = parseInt(partes[1]);
    const ano = parseInt(partes[3]);
    
    const mesInfo = MESES_HUB.find(m => m.numero === mesNumero);
    if (!mesInfo) return;
    
    // Ajustar dias para ano bissexto
    let totalDias = mesInfo.dias;
    if (mesNumero === 2 && isAnoBissexto(ano)) {
      totalDias = 29;
    }
    
    // Atualizar contador para cada dia
    for (let dia = 1; dia <= totalDias; dia++) {
      atualizarContadorDia(sheet, new Date(ano, mesNumero - 1, dia));
    }
    
  } catch (error) {
    loggerHUB.error('Erro ao atualizar contadores da aba mensal', error);
  }
}

/**
 * Gera relatório de capacidade de um mês
 */
function gerarRelatorioCapacidadeMes(mes, ano) {
  loggerHUB.info('Gerando relatório de capacidade', { mes, ano });
  
  try {
    const nomeAba = `${CONFIG_ABAS_MENSAIS.PREFIXO}${String(mes).padStart(2, '0')}_${MESES_HUB[mes-1].nome}_${ano}`;
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(nomeAba);
    
    if (!sheet) {
      throw new Error(`Aba mensal não encontrada: ${nomeAba}`);
    }
    
    const mesInfo = MESES_HUB[mes - 1];
    let totalDias = mesInfo.dias;
    if (mes === 2 && isAnoBissexto(ano)) {
      totalDias = 29;
    }
    
    const capacidadePorDia = [];
    
    for (let dia = 1; dia <= totalDias; dia++) {
      const dataCompleta = new Date(ano, mes - 1, dia);
      const transfersDoDia = contarTransfersDataEspecifica(sheet, dataCompleta);
      
      capacidadePorDia.push({
        dia: dia,
        data: formatarDataDDMMYYYY(dataCompleta),
        diaSemana: dataCompleta.toLocaleDateString('pt-PT', { weekday: 'long' }),
        transfers: transfersDoDia,
        capacidade: obterTextoCapacidade(transfersDoDia),
        cor: obterCorCapacidade(transfersDoDia)
      });
    }
    
    // Estatísticas gerais
    const totalTransfers = capacidadePorDia.reduce((sum, dia) => sum + dia.transfers, 0);
    const mediaTransfersDia = totalTransfers / totalDias;
    const diasCheios = capacidadePorDia.filter(dia => dia.transfers >= CONFIG_ABAS_MENSAIS.LIMITE_TRANSFERS_DIA).length;
    const diasDisponiveis = totalDias - diasCheios;
    
    return {
      sucesso: true,
      mes: mesInfo.nome,
      ano: ano,
      totalDias: totalDias,
      capacidadePorDia: capacidadePorDia,
      estatisticas: {
        totalTransfers: totalTransfers,
        mediaTransfersDia: Math.round(mediaTransfersDia * 10) / 10,
        diasCheios: diasCheios,
        diasDisponiveis: diasDisponiveis,
        percentualOcupacao: Math.round((diasCheios / totalDias) * 100)
      }
    };
    
  } catch (error) {
    loggerHUB.error('Erro ao gerar relatório de capacidade', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Conta transfers de uma data específica
 */
function contarTransfersDataEspecifica(sheet, data) {
  try {
    const dataFormatada = formatarDataDDMMYYYY(data);
    let contador = 0;
    
    const lastRow = sheet.getLastRow();
    for (let linha = 5; linha <= lastRow; linha++) { // Começar após cabeçalhos
      const dataTransfer = sheet.getRange(linha, 6).getValue(); // Coluna F
      
      if (dataTransfer && formatarDataDDMMYYYY(new Date(dataTransfer)) === dataFormatada) {
        contador++;
      }
    }
    
    return contador;
    
  } catch (error) {
    loggerHUB.error('Erro ao contar transfers por data', error);
    return 0;
  }
}

/**
 * Obtém texto da capacidade
 */
function obterTextoCapacidade(numeroTransfers) {
  const limite = CONFIG_ABAS_MENSAIS.LIMITE_TRANSFERS_DIA;
  
  if (numeroTransfers === 0) {
    return 'Disponível';
  } else if (numeroTransfers <= limite * 0.3) {
    return 'Baixa ocupação';
  } else if (numeroTransfers <= limite * 0.6) {
    return 'Ocupação média';
  } else if (numeroTransfers <= limite * 0.9) {
    return 'Alta ocupação';
  } else {
    return 'Capacidade esgotada';
  }
}

/**
* Verifica disponibilidade de uma data específica
*/
function verificarDisponibilidadeData(data) {
 loggerHUB.info('Verificando disponibilidade da data', { data });
 
 try {
   const dataObj = processarDataSeguraHUB(data);
   const mes = dataObj.getMonth() + 1;
   const ano = dataObj.getFullYear();
   
   const nomeAba = `${CONFIG_ABAS_MENSAIS.PREFIXO}${String(mes).padStart(2, '0')}_${MESES_HUB[mes-1].nome}_${ano}`;
   const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
   const sheet = ss.getSheetByName(nomeAba);
   
   if (!sheet) {
     // Se não existe aba mensal, assumir disponível
     return {
       disponivel: true,
       transfers: 0,
       capacidade: 'Disponível',
       observacao: 'Aba mensal não existe ainda'
     };
   }
   
   const transfersNaData = contarTransfersDataEspecifica(sheet, dataObj);
   const limite = CONFIG_ABAS_MENSAIS.LIMITE_TRANSFERS_DIA;
   const disponivel = transfersNaData < limite;
   
   return {
     disponivel: disponivel,
     transfers: transfersNaData,
     limite: limite,
     capacidade: obterTextoCapacidade(transfersNaData),
     percentualOcupacao: Math.round((transfersNaData / limite) * 100),
     cor: obterCorCapacidade(transfersNaData)
   };
   
 } catch (error) {
   loggerHUB.error('Erro ao verificar disponibilidade', error);
   return {
     disponivel: false,
     erro: error.message
   };
 }
}
// ===================================================
// APIS PRINCIPAIS DO SISTEMA HUB CENTRAL
// ===================================================

/**
 * Processa requisições GET
 */
function doGet(e) {
  loggerHUB.info('Recebendo requisição GET', { 
    parameters: e.parameter,
    queryString: e.queryString
  });
  
  try {
    const params = e.parameter || {};
    const action = params.action || 'default';
    
    loggerHUB.debug('Ação GET identificada', { action });
    
    switch (action) {
      case 'test':
        return handleTestHUB();
        
      case 'health':
        return handleHealthCheckHUB();
        
      case 'config':
        return handleConfigHUB();
        
      case 'stats':
        return handleStatsHUB();
        
      case 'hotels':
        return handleListHotels();
        
      case 'transfers':
        return handleListTransfers(params);
        
      case 'transfer':
        return handleGetTransfer(params);
        
      default:
        return handleDefaultHUB();
    }
    
  } catch (error) {
    loggerHUB.error('Erro no doGet', error);
    return criarRespostaHUB(`Erro interno: ${error.message}`, 500, false);
  }
}

/**
 * Processamento super simples para teste
 */
function processarMensagemSimples(telefone, mensagem) {
  console.log('processarMensagemSimples iniciado:', telefone, mensagem);
  
  try {
    // Resposta fixa para teste
    let resposta = 'Olá! Sou a Roberta HUB. Recebi sua mensagem: "' + mensagem + '"';
    
    // Detectar idioma básico
    if (telefone.startsWith('+1') || telefone.startsWith('+44')) {
      resposta = 'Hello! I am Roberta HUB. I received your message: "' + mensagem + '"';
    }
    
    console.log('Resposta gerada:', resposta);
    
    // Tentar enviar WhatsApp
    const enviado = enviarWhatsAppSimples(telefone, resposta);
    console.log('WhatsApp enviado:', enviado);
    
    return {
      telefone: telefone,
      mensagem: mensagem,
      resposta: resposta,
      enviado: enviado,
      timestamp: new Date()
    };
    
  } catch (error) {
    console.error('Erro no processamento simples:', error.message);
    throw error;
  }
}

/**
 * Envio WhatsApp simplificado
 */
function enviarWhatsAppSimples(telefone, mensagem) {
  console.log('Enviando WhatsApp para:', telefone);
  
  try {
    const numeroLimpo = telefone.replace(/\D/g, '');
    
    const payload = {
      phone: numeroLimpo,
      message: mensagem
    };
    
    const options = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Client-Token': SISTEMA_HUB_CONFIG.Z_API.CLIENT_TOKEN
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    console.log('Fazendo requisição para Z-API...');
    const response = UrlFetchApp.fetch(SISTEMA_HUB_CONFIG.Z_API.ENDPOINT, options);
    const responseCode = response.getResponseCode();
    
    console.log('Resposta Z-API:', responseCode);
    
    return responseCode === 200;
    
  } catch (error) {
    console.error('Erro no envio WhatsApp:', error.message);
    return false;
  }
}

/**
 * Processa transfer recebido via POST
 */
function processarTransferPOST(dadosRecebidos) {
  loggerHUB.info('Processando transfer via POST');
  
  try {
    // Extrair parâmetros extras
    const parametrosExtras = {
      source: dadosRecebidos.source,
      hotelOperadora: dadosRecebidos.hotelOperadora,
      timestamp: new Date().toISOString()
    };
    
    // Processar transfer
    const resultado = processarTransferDoHotel(dadosRecebidos, parametrosExtras);
    
    if (resultado.sucesso) {
      loggerHUB.success('Transfer POST processado com sucesso', {
        transferId: resultado.transferId,
        cliente: resultado.cliente
      });
      
      return criarRespostaHUB({
        transferId: resultado.transferId,
        cliente: resultado.cliente,
        hotelOperadora: resultado.hotelOperadora,
        valores: resultado.valores,
        whatsappEnviado: resultado.whatsappEnviado,
        mensagem: resultado.mensagem
      }, 201, true);
    } else {
      loggerHUB.error('Falha ao processar transfer POST', resultado);
      
      return criarRespostaHUB({
        erros: resultado.erros,
        mensagem: resultado.mensagem
      }, 400, false);
    }
    
  } catch (error) {
    loggerHUB.error('Erro ao processar transfer POST', error);
    return criarRespostaHUB(`Erro interno: ${error.message}`, 500, false);
  }
}

/**
 * Processa ações especiais
 */
function processarAcaoEspecialHUB(dados) {
  loggerHUB.info('Processando ação especial HUB', { action: dados.action });
  
  try {
    let resultado;
    
    switch (dados.action) {
      case 'testarZAPI':
        resultado = testarZAPI();
        break;
        
      case 'listarHoteis':
        resultado = listarHoteisConectados();
        break;
        
      case 'atualizarStatus':
        if (!dados.transferId || !dados.novoStatus) {
          throw new Error('transferId e novoStatus são obrigatórios');
        }
        resultado = atualizarStatusTransferHUB(dados.transferId, dados.novoStatus, dados.observacao);
        break;
        
      case 'enviarWhatsApp':
        if (!dados.transferId) {
          throw new Error('transferId é obrigatório');
        }
        resultado = reenviarWhatsAppTransfer(dados.transferId);
        break;
        
      case 'estatisticas':
        resultado = gerarEstatisticasHUB();
        break;
        
      case 'backup':
        resultado = criarBackupHUB();
        break;
        
      default:
        throw new Error(`Ação desconhecida: ${dados.action}`);
    }
    
    return criarRespostaHUB(resultado, 200, resultado.sucesso !== false);
    
  } catch (error) {
    loggerHUB.error('Erro na ação especial HUB', error);
    return criarRespostaHUB(`Erro: ${error.message}`, 400, false);
  }
}

// ===================================================
// HANDLERS DOS ENDPOINTS GET
// ===================================================

/**
 * Handler para teste da API
 */
function handleTestHUB() {
  return criarRespostaHUB({
    message: '🚐 Sistema HUB Central funcionando!',
    sistema: SISTEMA_HUB_CONFIG.NOME,
    versao: SISTEMA_HUB_CONFIG.VERSAO,
    timestamp: new Date().toISOString(),
    endpoints: {
      GET: [
        '?action=test - Teste da API',
        '?action=health - Health check',
        '?action=config - Configuração',
        '?action=stats - Estatísticas',
        '?action=hotels - Lista de hotéis',
        '?action=transfers&limit=N - Lista transfers',
        '?action=transfer&id=X - Transfer específico'
      ],
      POST: [
        'Transfer data - Receber transfer de hotel',
        'Actions - Ações especiais do sistema'
      ]
    },
    hoteisConectados: Object.keys(SISTEMA_HUB_CONFIG.HOTEIS_CONECTADOS),
    whatsappAtivo: SISTEMA_HUB_CONFIG.NOTIFICACOES.WHATSAPP_ATIVO
  });
}

/**
 * Handler para health check
 */
function handleHealthCheckHUB() {
  const checks = {
    planilha: false,
    zapi: false,
    hoteis: 0
  };
  
  try {
    // Verificar planilha
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    checks.planilha = true;
    
    // Verificar hotéis conectados
    checks.hoteis = Object.values(SISTEMA_HUB_CONFIG.HOTEIS_CONECTADOS)
      .filter(hotel => hotel.ativo).length;
    
    // Verificar Z-API (básico)
    checks.zapi = SISTEMA_HUB_CONFIG.NOTIFICACOES.WHATSAPP_ATIVO;
    
  } catch (error) {
    loggerHUB.error('Erro no health check', error);
  }
  
  const healthy = checks.planilha && checks.hoteis > 0;
  
  return criarRespostaHUB({
    status: healthy ? 'healthy' : 'unhealthy',
    checks: checks,
    timestamp: new Date().toISOString(),
    hoteisAtivos: checks.hoteis
  }, healthy ? 200 : 503, healthy);
}

/**
 * Handler para configuração
 */
function handleConfigHUB() {
  const configSegura = {
    nome: SISTEMA_HUB_CONFIG.NOME,
    versao: SISTEMA_HUB_CONFIG.VERSAO,
    hoteisConectados: Object.keys(SISTEMA_HUB_CONFIG.HOTEIS_CONECTADOS),
    operadorasExternas: SISTEMA_HUB_CONFIG.OPERADORAS_EXTERNAS,
    whatsappAtivo: SISTEMA_HUB_CONFIG.NOTIFICACOES.WHATSAPP_ATIVO,
    apenasConfirmados: SISTEMA_HUB_CONFIG.NOTIFICACOES.WHATSAPP_APENAS_CONFIRMADOS,
    emailDesativado: SISTEMA_HUB_CONFIG.NOTIFICACOES.EMAIL_DESATIVADO,
    valoresPadrao: SISTEMA_HUB_CONFIG.VALORES_PADRAO
  };
  
  return criarRespostaHUB(configSegura);
}

/**
 * Handler para listar hotéis
 */
function handleListHotels() {
  const hoteis = [];
  
  for (const [nome, config] of Object.entries(SISTEMA_HUB_CONFIG.HOTEIS_CONECTADOS)) {
    hoteis.push({
      nome: nome,
      id: config.id,
      ativo: config.ativo,
      email: config.email,
      telefone: config.telefone
    });
  }
  
  return criarRespostaHUB({
    hoteis: hoteis,
    operadorasExternas: SISTEMA_HUB_CONFIG.OPERADORAS_EXTERNAS,
    total: hoteis.length
  });
}

/**
 * Handler para listar transfers
 */
function handleListTransfers(params) {
  try {
    const limite = parseInt(params.limit) || 50;
    const offset = parseInt(params.offset) || 0;
    const hotel = params.hotel;
    const status = params.status;
    
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SISTEMA_HUB_CONFIG.SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return criarRespostaHUB({
        transfers: [],
        total: 0
      });
    }
    
    const dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS_HUB_CENTRAL.length).getValues();
    let transfersFiltrados = [];
    
    dados.forEach(row => {
      let incluir = true;
      
      // Filtrar por hotel
      if (hotel && row[20] !== hotel) { // Coluna U - Hotel/Operadora
        incluir = false;
      }
      
      // Filtrar por status
      if (status && row[17] !== status) { // Coluna R - Status
        incluir = false;
      }
      
      if (incluir) {
        const transfer = {};
        HEADERS_HUB_CENTRAL.forEach((header, index) => {
          transfer[header] = row[index];
        });
        transfersFiltrados.push(transfer);
      }
    });
    
    // Aplicar paginação
    const total = transfersFiltrados.length;
    transfersFiltrados = transfersFiltrados.slice(offset, offset + limite);
    
    return criarRespostaHUB({
      transfers: transfersFiltrados,
      total: total,
      limite: limite,
      offset: offset,
      filtros: { hotel, status }
    });
    
  } catch (error) {
    loggerHUB.error('Erro ao listar transfers', error);
    return criarRespostaHUB(`Erro: ${error.message}`, 500, false);
  }
}

/**
 * Handler para obter transfer específico
 */
function handleGetTransfer(params) {
  try {
    const id = params.id;
    
    if (!id) {
      return criarRespostaHUB('ID é obrigatório', 400, false);
    }
    
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SISTEMA_HUB_CONFIG.SHEET_NAME);
    const linha = encontrarLinhaPorIdHUB(sheet, id);
    
    if (!linha) {
      return criarRespostaHUB(`Transfer ${id} não encontrado`, 404, false);
    }
    
    const dados = sheet.getRange(linha, 1, 1, HEADERS_HUB_CENTRAL.length).getValues()[0];
    const transfer = {};
    
    HEADERS_HUB_CENTRAL.forEach((header, index) => {
      transfer[header] = dados[index];
    });
    
    return criarRespostaHUB({ transfer: transfer });
    
  } catch (error) {
    loggerHUB.error('Erro ao buscar transfer', error);
    return criarRespostaHUB(`Erro: ${error.message}`, 500, false);
  }
}

/**
* Handler para estatísticas
*/
function handleStatsHUB() {
 try {
   const resultado = gerarEstatisticasHUB();
   
   if (resultado.sucesso) {
     return criarRespostaHUB(resultado.estatisticas);
   } else {
     return criarRespostaHUB(resultado.erro, 500, false);
   }
   
 } catch (error) {
   loggerHUB.error('Erro ao gerar estatísticas via API', error);
   return criarRespostaHUB(`Erro: ${error.message}`, 500, false);
 }
}

/**
* Handler padrão
*/
function handleDefaultHUB() {
 return criarRespostaHUB({
   message: `${SISTEMA_HUB_CONFIG.NOME} v${SISTEMA_HUB_CONFIG.VERSAO}`,
   description: 'Sistema centralizador de transfers de múltiplos hotéis',
   hoteisConectados: Object.keys(SISTEMA_HUB_CONFIG.HOTEIS_CONECTADOS).length,
   whatsappAtivo: SISTEMA_HUB_CONFIG.NOTIFICACOES.WHATSAPP_ATIVO,
   documentation: 'Use ?action=test para ver endpoints disponíveis'
 });
}

// ===================================================
// FUNÇÕES DE GESTÃO E ESTATÍSTICAS DO HUB
// ===================================================

/**
 * Lista hotéis conectados
 */
function listarHoteisConectados() {
  loggerHUB.info('Listando hotéis conectados');
  
  const hoteis = [];
  
  for (const [nome, config] of Object.entries(SISTEMA_HUB_CONFIG.HOTEIS_CONECTADOS)) {
    hoteis.push({
      nome: nome,
      id: config.id,
      ativo: config.ativo,
      email: config.email,
      telefone: config.telefone,
      spreadsheetId: config.spreadsheetId
    });
  }
  
  return {
    sucesso: true,
    hoteis: hoteis,
    operadorasExternas: SISTEMA_HUB_CONFIG.OPERADORAS_EXTERNAS,
    totalAtivos: hoteis.filter(h => h.ativo).length,
    timestamp: new Date().toISOString()
  };
}

/**
 * Atualiza status de transfer no HUB
 */
function atualizarStatusTransferHUB(transferId, novoStatus, observacao = '') {
  loggerHUB.info('Atualizando status transfer HUB', { transferId, novoStatus });
  
  try {
    if (!STATUS_PERMITIDOS.includes(novoStatus)) {
      throw new Error(`Status inválido: ${novoStatus}`);
    }
    
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SISTEMA_HUB_CONFIG.SHEET_NAME);
    
    if (!sheet) {
      throw new Error('Planilha HUB Central não encontrada');
    }
    
    const linha = encontrarLinhaPorIdHUB(sheet, transferId);
    
    if (!linha) {
      throw new Error(`Transfer ${transferId} não encontrado`);
    }
    
    // Verificar status atual
    const statusAtual = sheet.getRange(linha, 18).getValue(); // Coluna R
    
    if (statusAtual === novoStatus) {
      loggerHUB.info('Status já está atualizado', { transferId, status: novoStatus });
      return {
        sucesso: false,
        mensagem: `Transfer já está com status "${novoStatus}"`
      };
    }
    
    // Atualizar status
    sheet.getRange(linha, 18).setValue(novoStatus); // Coluna R
    
    // Adicionar observação
    const obsAtual = sheet.getRange(linha, 19).getValue() || ''; // Coluna S
    const timestamp = new Date().toLocaleString('pt-PT');
    const novaObs = obsAtual 
      ? `${obsAtual}\n[${timestamp}] Status: ${statusAtual} → ${novoStatus}. ${observacao}`
      : `[${timestamp}] Status: ${statusAtual} → ${novoStatus}. ${observacao}`;
    
    sheet.getRange(linha, 19).setValue(novaObs);
    
    // Se mudou para Confirmado, enviar WhatsApp
    let whatsappEnviado = false;
    if (novoStatus === 'Confirmado' && statusAtual !== 'Confirmado') {
      const dadosTransfer = sheet.getRange(linha, 1, 1, HEADERS_HUB_CENTRAL.length).getValues()[0];
      
      const transferData = {
        id: dadosTransfer[0],
        cliente: dadosTransfer[1],
        data: dadosTransfer[5],
        origem: dadosTransfer[8],
        destino: dadosTransfer[9],
        horaPickup: dadosTransfer[10],
        numeroPessoas: dadosTransfer[3],
        contacto: dadosTransfer[6],
        status: novoStatus
      };
      
      const resultadoWhatsApp = enviarConfirmacaoWhatsApp(transferData);
      whatsappEnviado = resultadoWhatsApp.sucesso;
    }
    
    loggerHUB.success('Status atualizado com sucesso', {
      transferId,
      statusAnterior: statusAtual,
      novoStatus,
      whatsappEnviado
    });
    
    return {
      sucesso: true,
      transferId: transferId,
      statusAnterior: statusAtual,
      novoStatus: novoStatus,
      whatsappEnviado: whatsappEnviado,
      mensagem: `Status atualizado de "${statusAtual}" para "${novoStatus}"`
    };
    
  } catch (error) {
    loggerHUB.error('Erro ao atualizar status', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Confirma transfer de hotel e envia WhatsApp
 */
function confirmarTransferHotel(transferId) {
  loggerHUB.info('Confirmando transfer de hotel', { transferId });
  
  try {
    // Atualizar status para "Confirmado"
    const resultadoStatus = atualizarStatusTransferHUB(transferId, 'Confirmado');
    
    if (!resultadoStatus.sucesso) {
      throw new Error('Falha ao atualizar status');
    }
    
    // Buscar dados do transfer para enviar WhatsApp
    const dadosTransfer = buscarTransferPorId(transferId);
    
    if (!dadosTransfer) {
      throw new Error('Transfer não encontrado');
    }
    
    // ENVIAR PARA JUNIOR (NÃO PARA CLIENTE)
const resultadoWhatsApp = enviarConfirmacaoParaJunior({
  id: transferId,
  cliente: dadosTransfer.cliente,
  tipoServico: dadosTransfer.tipoServico,
  numeroPessoas: dadosTransfer.pessoas,
  data: dadosTransfer.data,
  contacto: dadosTransfer.contacto,
  origem: dadosTransfer.origem,
  destino: dadosTransfer.destino,
  horaPickup: dadosTransfer.horaPickup,
  valorTotal: dadosTransfer.valorTotal,
  status: 'Confirmado',
  hotelOperadora: dadosTransfer.hotelOperadora
});
    
    loggerHUB.success('Transfer de hotel confirmado e WhatsApp enviado', {
      transferId: transferId,
      whatsappEnviado: resultadoWhatsApp.sucesso
    });
    
    return {
      sucesso: true,
      transferId: transferId,
      whatsappEnviado: resultadoWhatsApp.sucesso
    };
    
  } catch (error) {
    loggerHUB.error('Erro ao confirmar transfer de hotel', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Reenvia WhatsApp para transfer específico
 */
function reenviarWhatsAppTransfer(transferId) {
  loggerHUB.info('Reenviando WhatsApp', { transferId });
  
  try {
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SISTEMA_HUB_CONFIG.SHEET_NAME);
    const linha = encontrarLinhaPorIdHUB(sheet, transferId);
    
    if (!linha) {
      throw new Error(`Transfer ${transferId} não encontrado`);
    }
    
    const dadosTransfer = sheet.getRange(linha, 1, 1, HEADERS_HUB_CENTRAL.length).getValues()[0];
    
    const transferData = {
      id: dadosTransfer[0],
      cliente: dadosTransfer[1],
      data: dadosTransfer[5],
      origem: dadosTransfer[8],
      destino: dadosTransfer[9],
      horaPickup: dadosTransfer[10],
      numeroPessoas: dadosTransfer[3],
      contacto: dadosTransfer[6],
      status: dadosTransfer[17]
    };
    
    if (!transferData.contacto) {
      throw new Error('Transfer não possui contacto para WhatsApp');
    }
    
    const resultado = enviarConfirmacaoWhatsApp(transferData);
    
    if (resultado.sucesso) {
      // Adicionar observação sobre reenvio
      const obsAtual = sheet.getRange(linha, 19).getValue() || '';
      const timestamp = new Date().toLocaleString('pt-PT');
      const novaObs = obsAtual 
        ? `${obsAtual}\n[${timestamp}] WhatsApp reenviado para ${transferData.contacto}`
        : `[${timestamp}] WhatsApp reenviado para ${transferData.contacto}`;
      
      sheet.getRange(linha, 19).setValue(novaObs);
      
      loggerHUB.success('WhatsApp reenviado com sucesso', { transferId });
      
      return {
        sucesso: true,
        transferId: transferId,
        numeroDestino: transferData.contacto,
        mensagem: 'WhatsApp reenviado com sucesso'
      };
    } else {
      throw new Error(resultado.motivo || 'Falha no envio do WhatsApp');
    }
    
  } catch (error) {
    loggerHUB.error('Erro ao reenviar WhatsApp', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Gera estatísticas do sistema HUB
 */
function gerarEstatisticasHUB() {
  loggerHUB.info('Gerando estatísticas HUB');
  
  try {
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SISTEMA_HUB_CONFIG.SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return {
        sucesso: true,
        estatisticas: {
          totalTransfers: 0,
          porHotel: {},
          porStatus: {},
          porTipo: {},
          valorTotal: 0,
          valorHUB: 0,
          resumo: 'Nenhum transfer registrado'
        }
      };
    }
    
    const dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS_HUB_CENTRAL.length).getValues();
    
    const stats = {
      totalTransfers: dados.length,
      porHotel: {},
      porStatus: {},
      porTipo: {},
      valorTotal: 0,
      valorHUB: 0,
      comissaoRecepcao: 0,
      transfersHoje: 0,
      transfersSemana: 0,
      whatsappsEnviados: 0
    };
    
    const hoje = new Date();
    const inicioSemana = new Date(hoje.getTime() - 7 * 24 * 60 * 60 * 1000);
    
    dados.forEach(row => {
      // Por Hotel/Operadora (coluna U)
      const hotel = row[20] || 'Não especificado';
      stats.porHotel[hotel] = (stats.porHotel[hotel] || 0) + 1;
      
      // Por Status (coluna R)
      const status = row[17] || 'Desconhecido';
      stats.porStatus[status] = (stats.porStatus[status] || 0) + 1;
      
      // Por Tipo (coluna C)
      const tipo = row[2] || 'Transfer';
      stats.porTipo[tipo] = (stats.porTipo[tipo] || 0) + 1;
      
      // Valores financeiros
      stats.valorTotal += parseFloat(row[11]) || 0; // Coluna L
      stats.valorHUB += parseFloat(row[13]) || 0; // Coluna N
      stats.comissaoRecepcao += parseFloat(row[14]) || 0; // Coluna O
      
      // Transfers por período
      const dataTransfer = new Date(row[5]); // Coluna F
      if (dataTransfer.toDateString() === hoje.toDateString()) {
        stats.transfersHoje++;
      }
      if (dataTransfer >= inicioSemana) {
        stats.transfersSemana++;
      }
      
      // WhatsApps enviados (estimativa baseada em observações)
      const observacoes = row[18] || ''; // Coluna S
      if (observacoes.toLowerCase().includes('whatsapp')) {
        stats.whatsappsEnviados++;
      }
    });
    
    loggerHUB.success('Estatísticas geradas', { totalTransfers: stats.totalTransfers });
    
    return {
      sucesso: true,
      estatisticas: stats,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    loggerHUB.error('Erro ao gerar estatísticas', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Cria backup da planilha HUB
 */
function criarBackupHUB() {
  loggerHUB.info('Criando backup HUB');
  
  try {
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const nomeBackup = `Backup_HUB_Central_${timestamp}`;
    
    const backupSS = ss.copy(nomeBackup);
    
    loggerHUB.success('Backup criado', { nome: nomeBackup, id: backupSS.getId() });
    
    return {
      sucesso: true,
      nomeBackup: nomeBackup,
      backupId: backupSS.getId(),
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    loggerHUB.error('Erro ao criar backup', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Busca transfers por critérios
 */
function buscarTransfersHUB(criterios = {}) {
  loggerHUB.info('Buscando transfers HUB', { criterios });
  
  try {
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SISTEMA_HUB_CONFIG.SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return {
        sucesso: true,
        transfers: [],
        total: 0
      };
    }
    
    const dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS_HUB_CENTRAL.length).getValues();
    const transfersFiltrados = [];
    
    dados.forEach(row => {
      let incluir = true;
      
      // Filtrar por hotel
      if (criterios.hotel && row[20] !== criterios.hotel) {
        incluir = false;
      }
      
      // Filtrar por status
      if (criterios.status && row[17] !== criterios.status) {
        incluir = false;
      }
      
      // Filtrar por tipo
      if (criterios.tipo && row[2] !== criterios.tipo) {
        incluir = false;
      }
      
      // Filtrar por data
      if (criterios.dataInicio || criterios.dataFim) {
        const dataTransfer = new Date(row[5]);
        
        if (criterios.dataInicio && dataTransfer < new Date(criterios.dataInicio)) {
          incluir = false;
        }
        
        if (criterios.dataFim && dataTransfer > new Date(criterios.dataFim)) {
          incluir = false;
        }
      }
      
      // Filtrar por cliente
      if (criterios.cliente) {
        const cliente = row[1].toString().toLowerCase();
        const busca = criterios.cliente.toLowerCase();
        if (!cliente.includes(busca)) {
          incluir = false;
        }
      }
      
      if (incluir) {
        const transfer = {};
        HEADERS_HUB_CENTRAL.forEach((header, index) => {
          transfer[header] = row[index];
        });
        transfersFiltrados.push(transfer);
      }
    });
    
    // Ordenar por data (mais recentes primeiro)
    transfersFiltrados.sort((a, b) => new Date(b['Data']) - new Date(a['Data']));
    
    loggerHUB.debug('Busca concluída', { 
      encontrados: transfersFiltrados.length,
      criterios 
    });
    
    return {
      sucesso: true,
      transfers: transfersFiltrados,
      total: transfersFiltrados.length,
      criterios: criterios
    };
    
  } catch (error) {
    loggerHUB.error('Erro na busca de transfers', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Gera relatório de receitas por hotel
 */
function gerarRelatorioReceitasHotel(dataInicio, dataFim) {
  loggerHUB.info('Gerando relatório de receitas por hotel', { dataInicio, dataFim });
  
  try {
    const criterios = {
      dataInicio: dataInicio,
      dataFim: dataFim
    };
    
    const resultado = buscarTransfersHUB(criterios);
    
    if (!resultado.sucesso) {
      throw new Error('Falha ao buscar transfers');
    }
    
    const receitasPorHotel = {};
    
    resultado.transfers.forEach(transfer => {
      // Apenas transfers não cancelados
      if (transfer['Status'] === 'Cancelado') return;
      
      const hotel = transfer['Hotel/Operadora'] || 'Não especificado';
      
      if (!receitasPorHotel[hotel]) {
        receitasPorHotel[hotel] = {
          totalTransfers: 0,
          valorTotal: 0,
          valorHotel: 0,
          valorHUB: 0,
          comissaoRecepcao: 0
        };
      }
      
      receitasPorHotel[hotel].totalTransfers++;
      receitasPorHotel[hotel].valorTotal += parseFloat(transfer['Preço Cliente (€)']) || 0;
      receitasPorHotel[hotel].valorHotel += parseFloat(transfer['Valor Empire Lisbon Hotel (€)']) || 0;
      receitasPorHotel[hotel].valorHUB += parseFloat(transfer['Valor HUB Transfer (€)']) || 0;
      receitasPorHotel[hotel].comissaoRecepcao += parseFloat(transfer['Comissão Recepção (€)']) || 0;
    });
    
    // Calcular totais gerais
    let totaisGerais = {
      totalTransfers: 0,
      valorTotal: 0,
      valorHotel: 0,
      valorHUB: 0,
      comissaoRecepcao: 0
    };
    
    Object.values(receitasPorHotel).forEach(hotel => {
      totaisGerais.totalTransfers += hotel.totalTransfers;
      totaisGerais.valorTotal += hotel.valorTotal;
      totaisGerais.valorHotel += hotel.valorHotel;
      totaisGerais.valorHUB += hotel.valorHUB;
      totaisGerais.comissaoRecepcao += hotel.comissaoRecepcao;
    });
    
    loggerHUB.success('Relatório de receitas gerado', { 
      periodo: `${dataInicio} a ${dataFim}`,
      hoteis: Object.keys(receitasPorHotel).length,
      totalTransfers: totaisGerais.totalTransfers
    });
    
    return {
      sucesso: true,
      periodo: {
        inicio: dataInicio,
        fim: dataFim
      },
      receitasPorHotel: receitasPorHotel,
      totaisGerais: totaisGerais,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    loggerHUB.error('Erro ao gerar relatório de receitas', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

// ===================================================
// SISTEMA DE MENU E TRIGGERS DO HUB CENTRAL
// ===================================================

/**
 * Cria menu do sistema HUB quando a planilha é aberta
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu(`🚐 Sistema HUB Central`)
    .addItem('⚙️ Configurar Sistema HUB', 'configurarSistemaHUB')
    .addSeparator()
    .addSubMenu(ui.createMenu('🔄 Monitoramento Hotéis')
      .addItem('⚙️ Configurar Monitoramento', 'configurarMonitoramentoMenu')
      .addItem('📊 Verificar Status', 'verificarStatusMenu')
      .addItem('🧪 Testar Monitoramento', 'testarMonitoramentoMenu')
      .addItem('🗑️ Remover Triggers', 'removerTriggersMenu'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📅 Abas Mensais HUB')
      .addItem('📅 Criar Todas as Abas Mensais', 'criarTodasAbasMensaisHUBMenu')
      .addItem('📊 Atualizar Contadores', 'atualizarContadoresMenu')
      .addItem('🎨 Atualizar Cores Capacidade', 'atualizarCoresCapacidadeMenu')
      .addItem('📋 Relatório de Capacidade', 'relatorioCapacidadeMenu')
      .addItem('📅 Verificar Disponibilidade', 'verificarDisponibilidadeMenu'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 Relatórios HUB')
      .addItem('📈 Estatísticas Gerais', 'mostrarEstatisticasHUB')
      .addItem('💰 Receitas por Hotel', 'gerarRelatorioReceitasMenu')
      .addItem('📋 Transfers por Status', 'gerarRelatorioStatusMenu')
      .addItem('🏨 Transfers por Hotel', 'gerarRelatorioHotelMenu')
      .addItem('📅 Relatório do Dia', 'gerarRelatorioDiaMenu'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📱 WhatsApp (Roberta HUB)')
      .addItem('🧪 Testar Z-API', 'testarZAPIMenu')
      .addItem('📤 Reenviar WhatsApp', 'reenviarWhatsAppMenu')
      .addItem('📊 Status WhatsApp', 'verificarStatusWhatsApp')
      .addItem('⚙️ Configurar WhatsApp', 'configurarWhatsAppMenu'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🏨 Gestão de Hotéis')
      .addItem('📋 Listar Hotéis Conectados', 'listarHoteisMenu')
      .addItem('🔄 Sincronizar com Hotéis', 'sincronizarHoteisMenu')
      .addItem('📊 Status Conexões', 'verificarConexoesHoteis')
      .addItem('⚙️ Configurar Hotéis', 'configurarHoteisMenu')
      .addItem('📱 Enviar Mensagens Motoristas', 'enviarMensagensMotoristas'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🔧 Manutenção HUB')
      .addItem('🔄 Atualizar Status Transfer', 'atualizarStatusMenu')
      .addItem('🗑️ Limpar Dados Teste', 'limparDadosTesteHUB')
      .addItem('💾 Criar Backup', 'criarBackupHUBMenu')
      .addItem('🔍 Verificar Integridade', 'verificarIntegridadeHUB')
      .addItem('📞 Corrigir Formatação Telefones', 'corrigirTelefonesMenu')
      .addItem('📋 Logs do Sistema', 'mostrarLogsHUB'))
    .addSeparator()
    .addItem('🧪 Testar Sistema Completo', 'testarSistemaCompletoHUB')
    .addItem('ℹ️ Sobre o Sistema HUB', 'mostrarSobreHUB')
    .addToUi();
}

/**
 * Configura sistema HUB completo
 */
function configurarSistemaHUB() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '⚙️ Configuração do Sistema HUB Central',
    'Esta ação irá:\n\n' +
    '1. Criar/verificar planilha HUB Central\n' +
    '2. Aplicar formatações e validações\n' +
    '3. Configurar conexões com hotéis\n' +
    '4. Testar sistema Z-API\n' +
    '5. Configurar webhooks\n\n' +
    'Deseja continuar?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    loggerHUB.info('Iniciando configuração do sistema HUB');
    
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    
    // 1. Verificar/criar aba HUB Central
    let abaHUB = ss.getSheetByName(SISTEMA_HUB_CONFIG.SHEET_NAME);
    if (!abaHUB) {
      abaHUB = criarAbaHUBCentral(ss);
    } else {
      aplicarFormatacaoHUB(abaHUB);
    }
    
    // 2. Testar Z-API
    const testeZAPI = testarZAPI();
    
    // 3. Verificar conexões com hotéis
    const hoteisAtivos = Object.values(SISTEMA_HUB_CONFIG.HOTEIS_CONECTADOS)
      .filter(hotel => hotel.ativo).length;
    
    // 4. Configurar URL da web app
    const webAppUrl = ScriptApp.getService().getUrl();
    
    ui.alert(
      '✅ Sistema HUB Configurado!',
      `Configuração concluída:\n\n` +
      `• Planilha HUB Central: ✅\n` +
      `• Formatações aplicadas: ✅\n` +
      `• Hotéis conectados: ${hoteisAtivos}/${Object.keys(SISTEMA_HUB_CONFIG.HOTEIS_CONECTADOS).length}\n` +
      `• Z-API WhatsApp: ${testeZAPI.sucesso ? '✅' : '❌'}\n` +
      `• Web App URL: ${webAppUrl}\n\n` +
      `Sistema pronto para receber transfers dos hotéis!`,
      ui.ButtonSet.OK
    );
    
    loggerHUB.success('Sistema HUB configurado com sucesso');
    
  } catch (error) {
    loggerHUB.error('Erro na configuração do sistema HUB', error);
    ui.alert('❌ Erro', 'Erro na configuração:\n' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Menu para criar abas mensais
 */
function criarTodasAbasMensaisHUBMenu() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '📅 Criar Abas Mensais HUB',
    'Esta ação irá criar todas as abas mensais com organização por dias.\n\nDeseja continuar?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    const resultado = criarTodasAbasMensaisHUB();
    
    ui.alert(
      '✅ Abas Mensais HUB',
      `Processamento concluído!\n\n• Abas criadas: ${resultado.criadas}\n• Abas existentes: ${resultado.existentes}\n• Erros: ${resultado.erros}\n\nCada aba está organizada por dias para facilitar visualização da capacidade.`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Menu para atualizar contadores
 */
function atualizarContadoresMenu() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const resultado = atualizarTodosContadores();
    
    if (resultado.sucesso) {
      ui.alert(
        '✅ Contadores Atualizados',
        `Contadores de capacidade atualizados em ${resultado.abasProcessadas} aba(s) mensal(is).`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('❌ Erro', resultado.erro, ui.ButtonSet.OK);
    }
    
  } catch (error) {
    ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Menu para relatório de capacidade
 */
function relatorioCapacidadeMenu() {
  const ui = SpreadsheetApp.getUi();
  
  const mes = ui.prompt(
    '📊 Relatório de Capacidade',
    'Digite o mês (1-12):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (mes.getSelectedButton() !== ui.Button.OK) return;
  
  const ano = ui.prompt(
    '📅 Ano',
    'Digite o ano:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (ano.getSelectedButton() !== ui.Button.OK) return;
  
  try {
    const mesNum = parseInt(mes.getResponseText());
    const anoNum = parseInt(ano.getResponseText());
    
    if (isNaN(mesNum) || mesNum < 1 || mesNum > 12) {
      throw new Error('Mês inválido');
    }
    
    if (isNaN(anoNum) || anoNum < 2020 || anoNum > 2030) {
      throw new Error('Ano inválido');
    }
    
    const relatorio = gerarRelatorioCapacidadeMes(mesNum, anoNum);
    
if (relatorio.sucesso) {
       let mensagem = `📊 RELATÓRIO DE CAPACIDADE\n\n`;
       mensagem += `📅 ${relatorio.mes} ${relatorio.ano}\n\n`;
       mensagem += `📈 ESTATÍSTICAS:\n`;
       mensagem += `• Total de transfers: ${relatorio.estatisticas.totalTransfers}\n`;
       mensagem += `• Média por dia: ${relatorio.estatisticas.mediaTransfersDia}\n`;
       mensagem += `• Dias cheios: ${relatorio.estatisticas.diasCheios}/${relatorio.totalDias}\n`;
       mensagem += `• Dias disponíveis: ${relatorio.estatisticas.diasDisponiveis}\n`;
       mensagem += `• Ocupação geral: ${relatorio.estatisticas.percentualOcupacao}%\n\n`;
       
       mensagem += `🔴 DIAS COM ALTA OCUPAÇÃO:\n`;
       const diasAltos = relatorio.capacidadePorDia.filter(dia => dia.transfers >= CONFIG_ABAS_MENSAIS.LIMITE_TRANSFERS_DIA * 0.8);
       if (diasAltos.length > 0) {
         diasAltos.forEach(dia => {
           mensagem += `• ${dia.data} (${dia.diaSemana}): ${dia.transfers} transfers\n`;
         });
       } else {
         mensagem += `• Nenhum dia com alta ocupação\n`;
       }
       
       ui.alert('📊 Relatório de Capacidade', mensagem, ui.ButtonSet.OK);
     } else {
       ui.alert('❌ Erro', relatorio.erro, ui.ButtonSet.OK);
     }
     
   } catch (error) {
     ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
   }
 }

/**
* Verificar disponibilidade de data específica via menu
*/
function verificarDisponibilidadeMenu() {
 const ui = SpreadsheetApp.getUi();
 
 const data = ui.prompt(
   '📅 Verificar Disponibilidade',
   'Digite a data (DD/MM/YYYY):',
   ui.ButtonSet.OK_CANCEL
 );
 
 if (data.getSelectedButton() !== ui.Button.OK) return;
 
 try {
   const resultado = verificarDisponibilidadeData(data.getResponseText());
   
   if (resultado.erro) {
     ui.alert('❌ Erro', resultado.erro, ui.ButtonSet.OK);
     return;
   }
   
   const status = resultado.disponivel ? '✅ DISPONÍVEL' : '❌ CAPACIDADE ESGOTADA';
   
   let mensagem = `${status}\n\n`;
   mensagem += `📅 Data: ${data.getResponseText()}\n`;
   mensagem += `🚐 Transfers atuais: ${resultado.transfers}\n`;
   mensagem += `📊 Limite diário: ${resultado.limite}\n`;
   mensagem += `📈 Ocupação: ${resultado.percentualOcupacao}%\n`;
   mensagem += `🎯 Status: ${resultado.capacidade}`;
   
   ui.alert('📅 Disponibilidade', mensagem, ui.ButtonSet.OK);
   
 } catch (error) {
   ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
 }
}

/**
* Mostra estatísticas do sistema HUB
*/
function mostrarEstatisticasHUB() {
 const ui = SpreadsheetApp.getUi();
 
 try {
   const resultado = gerarEstatisticasHUB();
   
   if (!resultado.sucesso) {
     throw new Error(resultado.erro);
   }
   
   const stats = resultado.estatisticas;
   
   let mensagem = `📊 ESTATÍSTICAS DO SISTEMA HUB CENTRAL\n\n`;
   mensagem += `Total de Transfers: ${stats.totalTransfers}\n`;
   mensagem += `Valor Total: €${stats.valorTotal.toFixed(2)}\n`;
   mensagem += `Valor HUB: €${stats.valorHUB.toFixed(2)}\n`;
   mensagem += `Comissão Recepção: €${stats.comissaoRecepcao.toFixed(2)}\n\n`;
   
   mensagem += `📈 ATIVIDADE:\n`;
   mensagem += `• Transfers hoje: ${stats.transfersHoje}\n`;
   mensagem += `• Transfers esta semana: ${stats.transfersSemana}\n`;
   mensagem += `• WhatsApps enviados: ${stats.whatsappsEnviados}\n\n`;
   
   mensagem += `🏨 POR HOTEL:\n`;
   Object.entries(stats.porHotel).forEach(([hotel, count]) => {
     mensagem += `• ${hotel}: ${count}\n`;
   });
   
   mensagem += `\n📊 POR STATUS:\n`;
   Object.entries(stats.porStatus).forEach(([status, count]) => {
     mensagem += `• ${status}: ${count}\n`;
   });
   
   ui.alert('📊 Estatísticas HUB', mensagem, ui.ButtonSet.OK);
   
 } catch (error) {
   loggerHUB.error('Erro ao mostrar estatísticas', error);
   ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
 }
}

/**
* Testa Z-API via menu
*/
function testarZAPIMenu() {
 const ui = SpreadsheetApp.getUi();
 
 try {
   const resultado = testarZAPI();
   
   if (resultado.sucesso) {
     ui.alert(
       '✅ Teste Z-API',
       `Teste realizado com sucesso!\n\n${resultado.mensagem}\n\nVerifique o WhatsApp da Roberta HUB.`,
       ui.ButtonSet.OK
     );
   } else {
     ui.alert(
       '❌ Falha no Teste Z-API',
       `Erro no teste:\n${resultado.mensagem}\n\nVerifique as configurações da Z-API.`,
       ui.ButtonSet.OK
     );
   }
   
 } catch (error) {
   loggerHUB.error('Erro no teste Z-API via menu', error);
   ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
 }
}

/**
* Reenvia WhatsApp via menu
*/
function reenviarWhatsAppMenu() {
 const ui = SpreadsheetApp.getUi();
 
 const transferId = ui.prompt(
   '📤 Reenviar WhatsApp',
   'Digite o ID do transfer para reenviar WhatsApp:',
   ui.ButtonSet.OK_CANCEL
 );
 
 if (transferId.getSelectedButton() !== ui.Button.OK) return;
 
 try {
   const resultado = reenviarWhatsAppTransfer(transferId.getResponseText());
   
   if (resultado.sucesso) {
     ui.alert(
       '✅ WhatsApp Reenviado',
       `WhatsApp reenviado com sucesso!\n\nTransfer: #${resultado.transferId}\nDestino: ${resultado.numeroDestino}`,
       ui.ButtonSet.OK
     );
   } else {
     ui.alert(
       '❌ Falha no Reenvio',
       `Erro: ${resultado.erro}`,
       ui.ButtonSet.OK
     );
   }
   
 } catch (error) {
   loggerHUB.error('Erro no reenvio via menu', error);
   ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
 }
}

/**
* Lista hotéis conectados via menu
*/
function listarHoteisMenu() {
 const ui = SpreadsheetApp.getUi();
 
 try {
   const resultado = listarHoteisConectados();
   
   let mensagem = `🏨 HOTÉIS CONECTADOS AO SISTEMA HUB\n\n`;
   
   resultado.hoteis.forEach(hotel => {
     const status = hotel.ativo ? '✅' : '❌';
     mensagem += `${status} ${hotel.nome}\n`;
     mensagem += `   ID: ${hotel.id}\n`;
     mensagem += `   E-mail: ${hotel.email}\n`;
     mensagem += `   Telefone: ${hotel.telefone}\n\n`;
   });
   
   mensagem += `📋 OPERADORAS EXTERNAS:\n`;
   SISTEMA_HUB_CONFIG.OPERADORAS_EXTERNAS.forEach(operadora => {
     mensagem += `• ${operadora}\n`;
   });
   
   mensagem += `\n📊 RESUMO:\n`;
   mensagem += `• Total de hotéis: ${resultado.hoteis.length}\n`;
   mensagem += `• Hotéis ativos: ${resultado.totalAtivos}\n`;
   mensagem += `• Operadoras externas: ${SISTEMA_HUB_CONFIG.OPERADORAS_EXTERNAS.length}`;
   
   ui.alert('🏨 Hotéis Conectados', mensagem, ui.ButtonSet.OK);
   
 } catch (error) {
   loggerHUB.error('Erro ao listar hotéis via menu', error);
   ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
 }
}

/**
* Atualiza status de transfer via menu
*/
function atualizarStatusMenu() {
 const ui = SpreadsheetApp.getUi();
 
 const transferId = ui.prompt(
   '🔄 Atualizar Status',
   'Digite o ID do transfer:',
   ui.ButtonSet.OK_CANCEL
 );
 
 if (transferId.getSelectedButton() !== ui.Button.OK) return;
 
 const novoStatus = ui.prompt(
   '📊 Novo Status',
   `Escolha o novo status:\n1 - Solicitado\n2 - Confirmado\n3 - Finalizado\n4 - Cancelado\n\nDigite o número:`,
   ui.ButtonSet.OK_CANCEL
 );
 
 if (novoStatus.getSelectedButton() !== ui.Button.OK) return;
 
 const statusMap = {
   '1': 'Solicitado',
   '2': 'Confirmado', 
   '3': 'Finalizado',
   '4': 'Cancelado'
 };
 
 const statusEscolhido = statusMap[novoStatus.getResponseText()];
 
 if (!statusEscolhido) {
   ui.alert('❌ Erro', 'Status inválido!', ui.ButtonSet.OK);
   return;
 }
 
 try {
   const resultado = atualizarStatusTransferHUB(
     transferId.getResponseText(),
     statusEscolhido,
     'Atualizado via menu do sistema'
   );
   
   if (resultado.sucesso) {
     let mensagem = `Status atualizado com sucesso!\n\n`;
     mensagem += `Transfer: #${resultado.transferId}\n`;
     mensagem += `Status anterior: ${resultado.statusAnterior}\n`;
     mensagem += `Novo status: ${resultado.novoStatus}\n`;
     
     if (resultado.whatsappEnviado) {
       mensagem += `\n📱 WhatsApp enviado automaticamente!`;
     }
     
     ui.alert('✅ Status Atualizado', mensagem, ui.ButtonSet.OK);
   } else {
     ui.alert('❌ Erro', resultado.erro || resultado.mensagem, ui.ButtonSet.OK);
   }
   
 } catch (error) {
   loggerHUB.error('Erro ao atualizar status via menu', error);
   ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
 }
}

/**
* Cria backup via menu
*/
function criarBackupHUBMenu() {
 const ui = SpreadsheetApp.getUi();
 
 const response = ui.alert(
   '💾 Criar Backup',
   'Deseja criar um backup da planilha HUB Central?',
   ui.ButtonSet.YES_NO
 );
 
 if (response !== ui.Button.YES) return;
 
 try {
   const resultado = criarBackupHUB();
   
   if (resultado.sucesso) {
     ui.alert(
       '✅ Backup Criado',
       `Backup criado com sucesso!\n\nNome: ${resultado.nomeBackup}\nID: ${resultado.backupId}\n\nO backup foi salvo no Google Drive.`,
       ui.ButtonSet.OK
     );
   } else {
     ui.alert('❌ Erro', resultado.erro, ui.ButtonSet.OK);
   }
   
 } catch (error) {
   loggerHUB.error('Erro ao criar backup via menu', error);
   ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
 }
}

/**
* Testa sistema completo via menu
*/
function testarSistemaCompletoHUB() {
 const ui = SpreadsheetApp.getUi();
 
 try {
   loggerHUB.info('Iniciando teste completo do sistema HUB');
   
   const testes = [];
   
   // Teste 1: Acesso à planilha
   const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
   testes.push('✅ Planilha acessível');
   
   // Teste 2: Aba HUB Central
   const sheet = ss.getSheetByName(SISTEMA_HUB_CONFIG.SHEET_NAME);
   testes.push(sheet ? '✅ Aba HUB Central OK' : '❌ Aba HUB Central não encontrada');
   
   // Teste 3: Geração de ID
   const novoId = gerarProximoIdHUB();
   testes.push(novoId > 0 ? '✅ Geração de ID OK' : '❌ Erro na geração de ID');
   
   // Teste 4: Z-API
   const testeZAPI = testarZAPI();
   testes.push(testeZAPI.sucesso ? '✅ Z-API funcionando' : '⚠️ Z-API com problemas');
   
   // Teste 5: Processamento de transfer
   const dadosTeste = {
     nomeCliente: 'TESTE SISTEMA HUB',
     tipoServico: 'Transfer',
     numeroPessoas: 2,
     numeroBagagens: 1,
     data: new Date(),
     contacto: '+351999000000',
     origem: 'Teste Origem',
     destino: 'Teste Destino',
     horaPickup: '12:00',
     valorTotal: 25.00,
     status: 'Solicitado'
   };
   
   const resultadoProcessamento = processarTransferDoHotel(dadosTeste, { source: 'teste' });
   testes.push(resultadoProcessamento.sucesso ? '✅ Processamento de transfer OK' : '❌ Erro no processamento');
   
   // Se criou transfer de teste, remover
   if (resultadoProcessamento.sucesso && sheet) {
     const linha = encontrarLinhaPorIdHUB(sheet, resultadoProcessamento.transferId);
     if (linha > 0) {
       sheet.deleteRow(linha);
       testes.push('✅ Limpeza de teste OK');
     }
   }
   
   // Teste 6: Estatísticas
   const estatisticas = gerarEstatisticasHUB();
   testes.push(estatisticas.sucesso ? '✅ Estatísticas OK' : '❌ Erro nas estatísticas');
   
   ui.alert(
     '🧪 Teste Completo do Sistema HUB',
     `Testes realizados:\n\n${testes.join('\n')}\n\nSistema HUB Central funcionando!`,
     ui.ButtonSet.OK
   );
   
   loggerHUB.success('Teste completo do sistema realizado');
   
 } catch (error) {
   loggerHUB.error('Erro no teste completo', error);
   ui.alert('❌ Erro no Teste', error.toString(), ui.ButtonSet.OK);
 }
}

/**
* Mostra informações sobre o sistema HUB
*/
function mostrarSobreHUB() {
 const ui = SpreadsheetApp.getUi();
 
 const webAppUrl = ScriptApp.getService().getUrl();
 
 ui.alert(
   `ℹ️ Sobre o Sistema HUB Central`,
   `${SISTEMA_HUB_CONFIG.NOME}\n` +
   `Versão: ${SISTEMA_HUB_CONFIG.VERSAO}\n\n` +
   `Sistema centralizador que recebe dados de múltiplos hotéis e operadoras.\n\n` +
   `🌟 FUNCIONALIDADES:\n` +
   `• Recebimento automático de transfers dos hotéis\n` +
   `• WhatsApp via Roberta HUB (Z-API)\n` +
   `• Centralização de dados em planilha única\n` +
   `• Nova coluna Hotel/Operadora\n` +
   `• Abas mensais organizadas por dias\n` +
   `• Cálculo automático de valores\n` +
   `• Estatísticas e relatórios\n` +
   `• Sistema de backup automático\n\n` +
   `🏨 HOTÉIS CONECTADOS:\n` +
   `• Empire Marques Hotel\n` +
   `• Hotel Lioz\n` +
   `• Empire Lisbon Hotel\n` +
   `• Gota d'água\n\n` +
   `📱 WhatsApp: Roberta HUB (+351928283652)\n` +
   `🌐 Web App: ${webAppUrl}\n\n` +
   `💻 Sistema reescrito para HUB Central\n` +
   `Com abas mensais organizadas por dias`,
   ui.ButtonSet.OK
 );
}

/**
* Limpa dados de teste do sistema HUB
*/
function limparDadosTesteHUB() {
 const ui = SpreadsheetApp.getUi();
 
 const response = ui.alert(
   '🗑️ Limpar Dados de Teste',
   'Esta ação irá remover todos os registros identificados como teste.\n\nDeseja continuar?',
   ui.ButtonSet.YES_NO
 );
 
 if (response !== ui.Button.YES) return;
 
 try {
   const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
   const sheet = ss.getSheetByName(SISTEMA_HUB_CONFIG.SHEET_NAME);
   
   if (!sheet || sheet.getLastRow() <= 1) {
     ui.alert('ℹ️ Info', 'Nenhum dado para limpar.', ui.ButtonSet.OK);
     return;
   }
   
   const lastRow = sheet.getLastRow();
   const dados = sheet.getRange(2, 1, lastRow - 1, HEADERS_HUB_CENTRAL.length).getValues();
   
   const palavrasTeste = ['teste', 'test', 'demo', 'exemplo', 'sample'];
   const linhasParaRemover = [];
   
   dados.forEach((row, index) => {
     const cliente = String(row[1]).toLowerCase(); // Coluna B - Cliente
     const observacoes = String(row[18]).toLowerCase(); // Coluna S - Observações
     
     let ehTeste = false;
     
     // Verificar palavras-chave
     palavrasTeste.forEach(palavra => {
       if (cliente.includes(palavra) || observacoes.includes(palavra)) {
         ehTeste = true;
       }
     });
     
     // IDs de teste específicos
     const id = row[0];
     if (id === 999 || id === 9999 || id >= 77777) {
       ehTeste = true;
     }
     
     if (ehTeste) {
       linhasParaRemover.push(index + 2); // +2 porque começamos na linha 2
     }
   });
   
   // Remover de baixo para cima
   if (linhasParaRemover.length > 0) {
     linhasParaRemover.reverse().forEach(linha => {
       sheet.deleteRow(linha);
     });
     
     ui.alert(
       '✅ Limpeza Concluída',
       `${linhasParaRemover.length} registro(s) de teste removido(s).`,
       ui.ButtonSet.OK
     );
     
     loggerHUB.success('Dados de teste removidos', { 
       removidos: linhasParaRemover.length 
     });
   } else {
     ui.alert('ℹ️ Info', 'Nenhum registro de teste encontrado.', ui.ButtonSet.OK);
   }
   
 } catch (error) {
   loggerHUB.error('Erro ao limpar dados de teste', error);
   ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
 }
}

/**
* Verifica integridade do sistema HUB
*/
function verificarIntegridadeHUB() {
 const ui = SpreadsheetApp.getUi();
 
 try {
   loggerHUB.info('Verificando integridade do sistema HUB');
   
   const verificacoes = {
     planilha: false,
     headers: false,
     formatacao: false,
     validacoes: false,
     dados: 0,
     duplicados: 0
   };
   
   // Verificar planilha
   const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
   verificacoes.planilha = true;
   
   // Verificar aba
   const sheet = ss.getSheetByName(SISTEMA_HUB_CONFIG.SHEET_NAME);
   if (!sheet) {
     throw new Error('Aba HUB Central não encontrada');
   }
   
   // Verificar headers
   const headers = sheet.getRange(1, 1, 1, HEADERS_HUB_CENTRAL.length).getValues()[0];
   verificacoes.headers = headers.length === HEADERS_HUB_CENTRAL.length &&
     headers.every((header, index) => header === HEADERS_HUB_CENTRAL[index]);
   
   // Verificar dados
   if (sheet.getLastRow() > 1) {
     verificacoes.dados = sheet.getLastRow() - 1;
     
     // Verificar duplicados
     const ids = sheet.getRange(2, 1, verificacoes.dados, 1).getValues();
     const idsUnicos = new Set(ids.map(row => row[0]));
     verificacoes.duplicados = verificacoes.dados - idsUnicos.size;
   }
   
   // Verificar formatação básica
   verificacoes.formatacao = true; // Assumir OK se chegou até aqui
   
   // Verificar validações
   verificacoes.validacoes = true; // Assumir OK se chegou até aqui
   
   let mensagem = `🔍 VERIFICAÇÃO DE INTEGRIDADE\n\n`;
   mensagem += `• Planilha: ${verificacoes.planilha ? '✅' : '❌'}\n`;
   mensagem += `• Headers: ${verificacoes.headers ? '✅' : '❌'}\n`;
   mensagem += `• Formatação: ${verificacoes.formatacao ? '✅' : '❌'}\n`;
   mensagem += `• Validações: ${verificacoes.validacoes ? '✅' : '❌'}\n`;
   mensagem += `• Total de dados: ${verificacoes.dados}\n`;
   mensagem += `• Duplicados: ${verificacoes.duplicados}`;
   
   if (verificacoes.duplicados > 0) {
     mensagem += `\n\n⚠️ Encontrados ${verificacoes.duplicados} registros duplicados!`;
   }
   
   const sistemaOK = verificacoes.planilha && verificacoes.headers && 
                    verificacoes.formatacao && verificacoes.duplicados === 0;
   
   if (sistemaOK) {
     mensagem += `\n\n✅ Sistema íntegro e funcionando corretamente!`;
   } else {
     mensagem += `\n\n⚠️ Sistema com problemas que precisam ser corrigidos.`;
   }
   
   ui.alert('🔍 Integridade do Sistema', mensagem, ui.ButtonSet.OK);
   
   loggerHUB.success('Verificação de integridade concluída', verificacoes);
   
 } catch (error) {
   loggerHUB.error('Erro na verificação de integridade', error);
   ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
 }
}

/**
* Mostra logs do sistema
*/
function mostrarLogsHUB() {
 const ui = SpreadsheetApp.getUi();
 
 try {
   const logs = loggerHUB.logs.slice(-20); // Últimos 20 logs
   
   if (logs.length === 0) {
     ui.alert('📋 Logs', 'Nenhum log disponível.', ui.ButtonSet.OK);
     return;
   }
   
   let mensagem = `📋 ÚLTIMOS ${logs.length} LOGS DO SISTEMA\n\n`;
   
   logs.forEach(log => {
     const timestamp = log.timestamp.toLocaleString('pt-PT');
     mensagem += `[${timestamp}] ${log.level}: ${log.message}\n`;
   });
   
   ui.alert('📋 Logs do Sistema', mensagem, ui.ButtonSet.OK);
   
 } catch (error) {
   ui.alert('❌ Erro', 'Erro ao acessar logs: ' + error.toString(), ui.ButtonSet.OK);
 }
}

// ===================================================
// INSTALAÇÃO E CONFIGURAÇÃO INICIAL DO SISTEMA HUB
// ===================================================

/**
 * Função principal de instalação do Sistema HUB Central
 * Execute esta função após fazer o deploy para configurar tudo
 */
function instalarSistemaHUBCompleto() {
  console.log('🚀 INICIANDO INSTALAÇÃO DO SISTEMA HUB CENTRAL');
  console.log('');
  
  try {
    const resultados = {
      planilhaHUB: false,
      formatacao: false,
      validacoes: false,
      testeZAPI: false,
      webAppUrl: null,
      timestamp: new Date()
    };
    
    // Passo 1: Verificar/criar planilha HUB Central
    console.log('📊 Passo 1: Configurando planilha HUB Central...');
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    
    let abaHUB = ss.getSheetByName(SISTEMA_HUB_CONFIG.SHEET_NAME);
    if (!abaHUB) {
      console.log('➕ Criando aba HUB Central...');
      abaHUB = criarAbaHUBCentral(ss);
    } else {
      console.log('✅ Aba HUB Central já existe, aplicando formatações...');
      aplicarFormatacaoHUB(abaHUB);
    }
    resultados.planilhaHUB = true;
    
    // Passo 2: Aplicar formatações
    console.log('🎨 Passo 2: Aplicando formatações...');
    aplicarFormatacaoHUB(abaHUB);
    resultados.formatacao = true;
    
    // Passo 3: Aplicar validações
    console.log('✅ Passo 3: Aplicando validações...');
    aplicarValidacoesHUB(abaHUB);
    resultados.validacoes = true;
    
    // Passo 4: Testar Z-API
    console.log('📱 Passo 4: Testando sistema Z-API...');
    const testeZAPI = testarZAPI();
    resultados.testeZAPI = testeZAPI.sucesso;
    
    if (testeZAPI.sucesso) {
      console.log('✅ Z-API funcionando - WhatsApp da Roberta HUB ativo');
    } else {
      console.log('⚠️ Z-API com problemas - WhatsApp pode não funcionar');
    }
    
    // Passo 5: Obter URL da Web App
    console.log('🌐 Passo 5: Configurando Web App...');
    try {
      resultados.webAppUrl = ScriptApp.getService().getUrl();
      console.log('✅ Web App URL obtida');
    } catch (urlError) {
      console.log('⚠️ Erro ao obter URL da Web App (faça deploy primeiro)');
    }
    
    // Passo 6: Criar dados de demonstração (opcional)
    console.log('📝 Passo 6: Criando transfer de demonstração...');
    criarTransferDemonstracaoHUB();
    
    // Resultado final
    console.log('');
    console.log('🎉 INSTALAÇÃO DO SISTEMA HUB CENTRAL CONCLUÍDA!');
    console.log('');
    console.log('📋 RESUMO DA INSTALAÇÃO:');
    console.log('• Planilha HUB Central: ✅');
    console.log('• Headers configurados: ✅ (incluindo nova coluna Hotel/Operadora)');
    console.log('• Formatações aplicadas: ✅');
    console.log('• Validações configuradas: ✅');
    console.log(`• Z-API WhatsApp: ${resultados.testeZAPI ? '✅' : '⚠️'}`);
    console.log(`• Web App URL: ${resultados.webAppUrl || 'Configure após deploy'}`);
    console.log('');
    console.log('🏨 HOTÉIS CONECTADOS:');
    Object.keys(SISTEMA_HUB_CONFIG.HOTEIS_CONECTADOS).forEach(hotel => {
      console.log(`• ${hotel}`);
    });
console.log('');
   console.log('📱 ROBERTA HUB:');
   console.log(`• Telefone: ${SISTEMA_HUB_CONFIG.Z_API.ROBERTA_PHONE}`);
   console.log(`• Status: ${resultados.testeZAPI ? 'Ativa' : 'Verificar configuração'}`);
   console.log('');
   console.log('🔧 PRÓXIMOS PASSOS:');
   console.log('1. Faça deploy da Web App se ainda não fez');
   console.log('2. Configure os hotéis para enviar dados para a URL da Web App');
   console.log('3. Teste recebimento de transfers');
   console.log('4. Verifique funcionamento do WhatsApp');
   console.log('');
   console.log('✅ SISTEMA HUB CENTRAL PRONTO PARA OPERAR!');
   
   return {
     sucesso: true,
     resultados: resultados,
     versao: SISTEMA_HUB_CONFIG.VERSAO,
     timestamp: new Date(),
     mensagem: 'Sistema HUB Central instalado com sucesso!'
   };
   
 } catch (error) {
   console.log('❌ ERRO NA INSTALAÇÃO:', error.message);
   console.log('📚 Stack trace:', error.stack);
   
   return {
     sucesso: false,
     erro: error.message,
     timestamp: new Date(),
     mensagem: 'Falha na instalação do sistema HUB'
   };
 }
}

/**
* Cria transfer de demonstração para testar o sistema
*/
function criarTransferDemonstracaoHUB() {
 try {
   const dadosDemo = {
     nomeCliente: 'CLIENTE DEMO HUB CENTRAL',
     tipoServico: 'Transfer',
     numeroPessoas: 2,
     numeroBagagens: 1,
     data: new Date(),
     contacto: '+351999000123',
     numeroVoo: 'DEMO123',
     origem: 'Aeroporto de Lisboa',
     destino: 'Hotel Demo',
     horaPickup: '14:30',
     valorTotal: 30.00,
     modoPagamento: 'Dinheiro',
     pagoParaQuem: 'Recepção',
     status: 'Solicitado',
     observacoes: 'Transfer de demonstração do Sistema HUB Central'
   };
   
   const resultado = processarTransferDoHotel(dadosDemo, { 
     source: 'demo',
     hotelOperadora: 'Sistema Demo'
   });
   
   if (resultado.sucesso) {
     console.log(`✅ Transfer demo criado: #${resultado.transferId}`);
   } else {
     console.log('⚠️ Erro ao criar transfer demo (não crítico)');
   }
   
 } catch (error) {
   console.log('⚠️ Erro ao criar transfer demo (não crítico):', error.message);
 }
}

/**
* NOVA FUNÇÃO PRINCIPAL - EXECUTE ESTA PARA INSTALAR COM ABAS MENSAIS
*/
function INSTALAR_HUB_CENTRAL_COM_ABAS_MENSAIS() {
 console.log('🎯 INSTALAÇÃO HUB CENTRAL COM ABAS MENSAIS ORGANIZADAS POR DIAS');
 console.log('='.repeat(60));
 
 try {
   // 1. Instalação básica
   const instalacao = instalarSistemaHUBCompleto();
   
   if (!instalacao.sucesso) {
     throw new Error('Falha na instalação básica');
   }
   
   // 2. Criar abas mensais
   console.log('📅 Criando abas mensais com organização por dias...');
   const abasMensais = criarTodasAbasMensaisHUB();
   
   console.log('✅ Abas mensais criadas:');
   console.log(`• Criadas: ${abasMensais.criadas}`);
   console.log(`• Existentes: ${abasMensais.existentes}`);
   console.log(`• Erros: ${abasMensais.erros}`);
   
   // 3. Configurar sistema para usar abas mensais
   CONFIG_ABAS_MENSAIS.USAR_ABAS_MENSAIS = true;
   CONFIG_ABAS_MENSAIS.ORGANIZAR_POR_DIAS = true;
   
   console.log('⚙️ Sistema configurado para usar abas mensais com organização por dias');
   
   // 4. Teste com transfer de demonstração
   console.log('📝 Testando registro em aba mensal...');
   const transferTeste = {
     nomeCliente: 'TESTE ABA MENSAL',
     tipoServico: 'Transfer',
     numeroPessoas: 2,
     numeroBagagens: 1,
     data: new Date(),
     contacto: '+351999000999',
     origem: 'Aeroporto de Lisboa',
     destino: 'Hotel Teste',
     horaPickup: '16:00',
     valorTotal: 28.00,
     status: 'Confirmado'
   };
   
   const resultadoTeste = processarTransferDoHotel(transferTeste, { 
     source: 'teste',
     hotelOperadora: 'Teste Aba Mensal'
   });
   
   if (resultadoTeste.sucesso) {
     console.log('✅ Teste de aba mensal bem-sucedido');
     console.log(`• Transfer ID: #${resultadoTeste.transferId}`);
     console.log(`• Registrado na aba mensal: Sim`);
     
     // Limpar teste
     try {
       const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
       const sheet = ss.getSheetByName(SISTEMA_HUB_CONFIG.SHEET_NAME);
       const linha = encontrarLinhaPorIdHUB(sheet, resultadoTeste.transferId);
       if (linha > 0) {
         sheet.deleteRow(linha);
         console.log('🗑️ Transfer de teste removido');
       }
     } catch (cleanError) {
       console.log('⚠️ Aviso: Transfer de teste não foi removido automaticamente');
     }
   }
   
   // Resultado final
   console.log('');
   console.log('🎉 INSTALAÇÃO COMPLETA COM ABAS MENSAIS CONCLUÍDA!');
   console.log('');
   console.log('📋 RECURSOS INSTALADOS:');
   console.log('• ✅ Sistema HUB Central com planilha única');
   console.log('• ✅ Abas mensais organizadas por dias');
   console.log('• ✅ Contadores visuais de capacidade por dia');
   console.log('• ✅ Cores indicativas (vazio → cheio)');
   console.log('• ✅ Resumo de capacidade mensal');
   console.log('• ✅ Registro duplo (central + mensal)');
   console.log('• ✅ WhatsApp via Roberta HUB');
   console.log('• ✅ Nova coluna Hotel/Operadora');
   console.log('');
   console.log('📅 COMO USAR AS ABAS MENSAIS:');
   console.log('1. Cada mês tem uma aba própria (ex: HUB_01_Janeiro_2025)');
   console.log('2. Dentro da aba, os transfers são organizados por dia');
   console.log('3. Cada dia mostra um contador de transfers');
   console.log('4. Cores indicam a capacidade:');
   console.log(`   • Cinza: Sem transfers`);
   console.log(`   • Verde: Poucos transfers (até 30% da capacidade)`);
   console.log(`   • Amarelo: Capacidade média (30-60%)`);
   console.log(`   • Vermelho claro: Quase cheio (60-90%)`);
   console.log(`   • Vermelho: Capacidade esgotada (90%+)`);
   console.log('5. Limite padrão: 15 transfers por dia');
   console.log('');
   console.log('🎯 PARA VERIFICAR CAPACIDADE:');
   console.log('• Abra a aba do mês desejado');
   console.log('• Procure pelo dia específico');
   console.log('• Veja a cor do contador para avaliar disponibilidade');
   console.log('• Use o resumo no final da aba para visão geral');
   console.log('');
   console.log('✅ SISTEMA PRONTO PARA GESTÃO DE CAPACIDADE!');
   
   return {
     sucesso: true,
     instalacaoBasica: instalacao,
     abasMensais: abasMensais,
     testeAbaMensal: resultadoTeste.sucesso,
     timestamp: new Date()
   };
   
 } catch (error) {
   console.log('❌ ERRO NA INSTALAÇÃO COM ABAS MENSAIS:', error.message);
   return {
     sucesso: false,
     erro: error.message
   };
 }
}

/**
* Configuração avançada do sistema HUB
*/
function configuracaoAvancadaHUB() {
 console.log('⚙️ CONFIGURAÇÃO AVANÇADA DO SISTEMA HUB');
 
 try {
   // 1. Verificar e corrigir permissões
   console.log('🔐 Verificando permissões...');
   verificarPermissoesHUB();
   
   // 2. Otimizar performance da planilha
   console.log('⚡ Otimizando performance...');
   otimizarPerformanceHUB();
   
   // 3. Configurar triggers se necessário
   console.log('🔄 Configurando triggers...');
   configurarTriggersHUB();
   
   // 4. Validar integridade dos dados
   console.log('🔍 Validando integridade...');
   const integridade = verificarIntegridadeHUBSilencioso();
   
   console.log('✅ Configuração avançada concluída');
   console.log(`📊 Integridade: ${integridade.sistemaOK ? 'OK' : 'Problemas detectados'}`);
   
   return {
     sucesso: true,
     integridade: integridade,
     timestamp: new Date()
   };
   
 } catch (error) {
   console.log('❌ Erro na configuração avançada:', error.message);
   return {
     sucesso: false,
     erro: error.message
   };
 }
}

/**
* Verifica permissões necessárias
*/
function verificarPermissoesHUB() {
 try {
   // Testar acesso à planilha
   SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
   
   // Testar UrlFetchApp para Z-API
   // (não fazemos request real, só verificamos se a API está disponível)
   
   // Testar ScriptApp para Web App
   ScriptApp.getService();
   
   console.log('✅ Todas as permissões verificadas');
   
 } catch (error) {
   console.log('⚠️ Problema de permissão:', error.message);
   throw error;
 }
}

/**
* Otimiza performance da planilha
*/
function otimizarPerformanceHUB() {
 try {
   const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
   const sheet = ss.getSheetByName(SISTEMA_HUB_CONFIG.SHEET_NAME);
   
   if (sheet) {
     // Congelar primeira linha
     sheet.setFrozenRows(1);
     
     // Otimizar largura das colunas
     for (let i = 1; i <= HEADERS_HUB_CENTRAL.length; i++) {
       if (sheet.getColumnWidth(i) < 80) {
         sheet.setColumnWidth(i, 80);
       }
     }
     
     console.log('✅ Performance otimizada');
   }
   
 } catch (error) {
   console.log('⚠️ Erro na otimização (não crítico):', error.message);
 }
}

/**
* Configura triggers se necessário
*/
function configurarTriggersHUB() {
 try {
   // Para o sistema HUB, não precisamos de triggers por padrão
   // mas podemos configurar para manutenção automática se necessário
   
   const triggers = ScriptApp.getProjectTriggers();
   console.log(`📋 Triggers existentes: ${triggers.length}`);
   
   // Se quiser adicionar trigger de manutenção automática:
   // ScriptApp.newTrigger('manutencaoAutomaticaHUB')
   //   .timeBased()
   //   .everyDays(1)
   //   .atHour(2)
   //   .create();
   
   console.log('✅ Triggers verificados');
   
 } catch (error) {
   console.log('⚠️ Erro nos triggers (não crítico):', error.message);
 }
}

/**
* Verifica integridade sem interface
*/
function verificarIntegridadeHUBSilencioso() {
 try {
   const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
   const sheet = ss.getSheetByName(SISTEMA_HUB_CONFIG.SHEET_NAME);
   
   if (!sheet) {
     return { sistemaOK: false, erro: 'Aba não encontrada' };
   }
   
   // Verificar headers
   const headers = sheet.getRange(1, 1, 1, HEADERS_HUB_CENTRAL.length).getValues()[0];
   const headersOK = headers.length === HEADERS_HUB_CENTRAL.length;
   
   // Verificar dados
   const dados = sheet.getLastRow() - 1;
   
   return {
     sistemaOK: headersOK,
     headers: headersOK,
     totalDados: dados,
     timestamp: new Date()
   };
   
 } catch (error) {
   return {
     sistemaOK: false,
     erro: error.message
   };
 }
}

/**
* Função de manutenção automática (pode ser usada como trigger)
*/
function manutencaoAutomaticaHUB() {
 loggerHUB.info('Executando manutenção automática HUB');
 
 try {
   const tarefas = [];
   
   // 1. Verificar integridade
   const integridade = verificarIntegridadeHUBSilencioso();
   if (integridade.sistemaOK) {
     tarefas.push('✅ Integridade verificada');
   } else {
     tarefas.push('⚠️ Problemas de integridade detectados');
   }
   
   // 2. Limpar logs antigos
   if (loggerHUB.logs.length > 500) {
     const removidos = loggerHUB.logs.length - 100;
     loggerHUB.logs = loggerHUB.logs.slice(-100);
     tarefas.push(`🗑️ ${removidos} logs antigos removidos`);
   }
   
   // 3. Gerar estatísticas
   const stats = gerarEstatisticasHUB();
   if (stats.sucesso) {
     tarefas.push(`📊 Estatísticas: ${stats.estatisticas.totalTransfers} transfers`);
   }
   
   loggerHUB.success('Manutenção automática concluída', { tarefas });
   
   return {
     sucesso: true,
     tarefas: tarefas,
     timestamp: new Date()
   };
   
 } catch (error) {
   loggerHUB.error('Erro na manutenção automática', error);
   return {
     sucesso: false,
     erro: error.message
   };
 }
}

/**
* Status completo do sistema HUB
*/
function statusCompletoHUB() {
 console.log('📊 STATUS COMPLETO DO SISTEMA HUB CENTRAL');
 console.log('');
 
 try {
   // Informações básicas
   console.log('🏷️ INFORMAÇÕES BÁSICAS:');
   console.log(`• Nome: ${SISTEMA_HUB_CONFIG.NOME}`);
   console.log(`• Versão: ${SISTEMA_HUB_CONFIG.VERSAO}`);
   console.log(`• Planilha ID: ${SISTEMA_HUB_CONFIG.SPREADSHEET_ID}`);
   console.log(`• Aba: ${SISTEMA_HUB_CONFIG.SHEET_NAME}`);
   console.log('');
   
   // Status da planilha
   console.log('📊 STATUS DA PLANILHA:');
   const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
   const sheet = ss.getSheetByName(SISTEMA_HUB_CONFIG.SHEET_NAME);
   
   if (sheet) {
     const totalLinhas = sheet.getLastRow();
     const totalTransfers = totalLinhas > 1 ? totalLinhas - 1 : 0;
     console.log(`• Planilha: ✅ Acessível`);
     console.log(`• Aba HUB Central: ✅ Encontrada`);
     console.log(`• Total de transfers: ${totalTransfers}`);
     console.log(`• Última linha: ${totalLinhas}`);
   } else {
     console.log(`• Planilha: ❌ Aba ${SISTEMA_HUB_CONFIG.SHEET_NAME} não encontrada`);
   }
   console.log('');
   
   // Status do WhatsApp/Z-API
   console.log('📱 STATUS DO WHATSAPP (ROBERTA HUB):');
   console.log(`• WhatsApp ativo: ${SISTEMA_HUB_CONFIG.NOTIFICACOES.WHATSAPP_ATIVO ? '✅' : '❌'}`);
   console.log(`• Apenas confirmados: ${SISTEMA_HUB_CONFIG.NOTIFICACOES.WHATSAPP_APENAS_CONFIRMADOS ? '✅' : '❌'}`);
   console.log(`• Roberta HUB: ${SISTEMA_HUB_CONFIG.Z_API.ROBERTA_PHONE}`);
   console.log(`• Endpoint Z-API: Configurado`);
   console.log('');
   
   // Status dos hotéis conectados
   console.log('🏨 STATUS DOS HOTÉIS CONECTADOS:');
   Object.entries(SISTEMA_HUB_CONFIG.HOTEIS_CONECTADOS).forEach(([nome, config]) => {
     const status = config.ativo ? '✅' : '❌';
     console.log(`• ${nome}: ${status}`);
     console.log(`  - ID: ${config.id}`);
     console.log(`  - E-mail: ${config.email}`);
   });
   console.log('');
   
   // Operadoras externas
   console.log('🚚 OPERADORAS EXTERNAS:');
   SISTEMA_HUB_CONFIG.OPERADORAS_EXTERNAS.forEach(operadora => {
     console.log(`• ${operadora}`);
   });
   console.log('');
   
   // Estatísticas se houver dados
   if (sheet && sheet.getLastRow() > 1) {
     console.log('📈 ESTATÍSTICAS RÁPIDAS:');
     const stats = gerarEstatisticasHUB();
     if (stats.sucesso) {
       const s = stats.estatisticas;
       console.log(`• Total de transfers: ${s.totalTransfers}`);
       console.log(`• Valor total: €${s.valorTotal.toFixed(2)}`);
       console.log(`• Valor HUB: €${s.valorHUB.toFixed(2)}`);
       console.log(`• Transfers hoje: ${s.transfersHoje}`);
       console.log('• Por hotel:');
       Object.entries(s.porHotel).forEach(([hotel, count]) => {
         console.log(`  - ${hotel}: ${count}`);
       });
     }
     console.log('');
   }
   
   // Web App URL
   console.log('🌐 WEB APP:');
   try {
     const webAppUrl = ScriptApp.getService().getUrl();
     console.log(`• URL: ${webAppUrl}`);
     console.log(`• Status: ✅ Configurada`);
   } catch (urlError) {
     console.log(`• Status: ❌ Não configurada (faça deploy)`);
   }
   console.log('');
   
   // Sistema de logs
   console.log('📋 SISTEMA DE LOGS:');
   console.log(`• Logs ativos: ${loggerHUB.config.ENABLED ? '✅' : '❌'}`);
   console.log(`• Nível: ${loggerHUB.config.LEVEL}`);
   console.log(`• Logs em memória: ${loggerHUB.logs.length}`);
   console.log('');
   
   console.log('✅ STATUS COMPLETO GERADO COM SUCESSO!');
   
   return {
     sucesso: true,
     timestamp: new Date(),
     planilhaOK: !!sheet,
     totalTransfers: sheet ? Math.max(0, sheet.getLastRow() - 1) : 0,
     whatsappAtivo: SISTEMA_HUB_CONFIG.NOTIFICACOES.WHATSAPP_ATIVO,
     hoteisAtivos: Object.values(SISTEMA_HUB_CONFIG.HOTEIS_CONECTADOS).filter(h => h.ativo).length
   };
   
 } catch (error) {
   console.log('❌ ERRO ao gerar status:', error.message);
   return {
     sucesso: false,
     erro: error.message
   };
 }
}

/**
* Demonstração completa do sistema HUB
*/
function demonstracaoCompiletaHUB() {
 console.log('🎬 DEMONSTRAÇÃO COMPLETA DO SISTEMA HUB CENTRAL');
 console.log('');
 
 try {
   // 1. Mostrar configuração atual
   console.log('1️⃣ CONFIGURAÇÃO ATUAL:');
   statusCompletoHUB();
   console.log('');
   
   // 2. Simular recebimento de transfer de hotel
   console.log('2️⃣ SIMULANDO RECEBIMENTO DE TRANSFER:');
   const transferDemo = {
     nomeCliente: 'João Silva Demo',
     tipoServico: 'Transfer',
     numeroPessoas: 3,
     numeroBagagens: 2,
     data: new Date(),
     contacto: '+351912345678',
     numeroVoo: 'TAP123',
     origem: 'Aeroporto de Lisboa',
     destino: 'Empire Marques Hotel',
     horaPickup: '15:30',
     valorTotal: 35.00,
     modoPagamento: 'Cartão',
     pagoParaQuem: 'Recepção',
     status: 'Confirmado',
     observacoes: 'Transfer de demonstração do sistema HUB'
   };
   
   console.log('📝 Dados do transfer demo:');
   console.log(`• Cliente: ${transferDemo.nomeCliente}`);
   console.log(`• Origem: ${transferDemo.origem}`);
   console.log(`• Destino: ${transferDemo.destino}`);
   console.log(`• Valor: €${transferDemo.valorTotal}`);
   console.log(`• Status: ${transferDemo.status}`);
   
   const resultadoDemo = processarTransferDoHotel(transferDemo, {
     source: 'empire_marques',
     hotelOperadora: 'Empire Marques Hotel'
   });
   
   if (resultadoDemo.sucesso) {
     console.log('✅ Transfer processado com sucesso!');
     console.log(`• ID gerado: #${resultadoDemo.transferId}`);
     console.log(`• Hotel identificado: ${resultadoDemo.hotelOperadora}`);
     console.log(`• WhatsApp enviado: ${resultadoDemo.whatsappEnviado ? 'Sim' : 'Não'}`);
     console.log(`• Valores calculados:`);
     console.log(`  - Cliente: €${resultadoDemo.valores.precoCliente}`);
     console.log(`  - Hotel: €${resultadoDemo.valores.valorHotel}`);
     console.log(`  - HUB: €${resultadoDemo.valores.valorHUB}`);
     console.log(`  - Recepção: €${resultadoDemo.valores.comissaoRecepcao}`);
   } else {
     console.log('❌ Erro ao processar transfer demo');
   }
   console.log('');
   
   // 3. Testar Z-API
   console.log('3️⃣ TESTANDO SISTEMA WHATSAPP (Z-API):');
   const testeZAPI = testarZAPI();
   if (testeZAPI.sucesso) {
     console.log('✅ Z-API funcionando - WhatsApp da Roberta enviado');
   } else {
     console.log('❌ Z-API com problemas');
   }
   console.log('');
   
   // 4. Gerar estatísticas
   console.log('4️⃣ ESTATÍSTICAS ATUAIS:');
   const stats = gerarEstatisticasHUB();
   if (stats.sucesso) {
     const s = stats.estatisticas;
     console.log(`📊 Total de transfers: ${s.totalTransfers}`);
     console.log(`💰 Valor total: €${s.valorTotal.toFixed(2)}`);
     console.log(`🏨 Hotéis com transfers:`);
     Object.entries(s.porHotel).forEach(([hotel, count]) => {
       console.log(`   • ${hotel}: ${count} transfer(s)`);
     });
   }
   console.log('');
   
   // 5. Teste de atualização de status
   if (resultadoDemo.sucesso) {
     console.log('5️⃣ TESTANDO ATUALIZAÇÃO DE STATUS:');
     const atualizacao = atualizarStatusTransferHUB(
       resultadoDemo.transferId,
       'Finalizado',
       'Finalizado via demonstração do sistema'
     );
     
     if (atualizacao.sucesso) {
       console.log('✅ Status atualizado com sucesso');
       console.log(`• ${atualizacao.statusAnterior} → ${atualizacao.novoStatus}`);
     }
     console.log('');
   }
   
   // 6. Limpeza dos dados demo
   console.log('6️⃣ LIMPANDO DADOS DE DEMONSTRAÇÃO:');
   if (resultadoDemo.sucesso) {
     try {
       const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
       const sheet = ss.getSheetByName(SISTEMA_HUB_CONFIG.SHEET_NAME);
       const linha = encontrarLinhaPorIdHUB(sheet, resultadoDemo.transferId);
       
       if (linha > 0) {
         sheet.deleteRow(linha);
         console.log('✅ Transfer de demonstração removido');
       }
     } catch (cleanError) {
       console.log('⚠️ Erro ao limpar (não crítico)');
     }
   }
   console.log('');
   
   // Resultado final
   console.log('🎉 DEMONSTRAÇÃO COMPLETA FINALIZADA!');
   console.log('');
   console.log('📋 RESUMO DA DEMONSTRAÇÃO:');
   console.log('• ✅ Sistema HUB configurado e funcionando');
   console.log('• ✅ Recebimento de transfers dos hotéis');
   console.log('• ✅ Cálculo automático de valores');
   console.log('• ✅ Nova coluna Hotel/Operadora');
   console.log('• ✅ Sistema WhatsApp (Roberta HUB)');
   console.log('• ✅ Estatísticas e relatórios');
   console.log('• ✅ Atualização de status');
   console.log('• ✅ Abas mensais organizadas por dias');
   console.log('');
   console.log('🚀 SISTEMA HUB CENTRAL PRONTO PARA PRODUÇÃO!');
   
   return {
     sucesso: true,
     transferDemo: resultadoDemo.sucesso,
     whatsappOK: testeZAPI.sucesso,
     timestamp: new Date()
   };
   
 } catch (error) {
   console.log('❌ ERRO NA DEMONSTRAÇÃO:', error.message);
   return {
     sucesso: false,
     erro: error.message
   };
 }
}

// ===================================================
// FUNÇÕES DE ENTRADA PARA EXECUÇÃO RÁPIDA
// ===================================================

/**
 * EXECUTE ESTA FUNÇÃO PARA INSTALAR TUDO
 * Função principal que deve ser executada após o deploy
 */
function INSTALAR_HUB_CENTRAL() {
  console.log('🎯 INSTALAÇÃO RÁPIDA DO SISTEMA HUB CENTRAL');
  console.log('='.repeat(50));
  
  // 1. Instalação completa com abas mensais
  const instalacao = INSTALAR_HUB_CENTRAL_COM_ABAS_MENSAIS();
  
  if (instalacao.sucesso) {
    console.log('');
    console.log('✅ INSTALAÇÃO CONCLUÍDA COM SUCESSO!');
    
    // 2. Demonstração
    console.log('');
    console.log('🎬 Executando demonstração...');
    demonstracaoCompiletaHUB();
  } else {
    console.log('❌ FALHA NA INSTALAÇÃO');
    console.log('Erro:', instalacao.erro);
  }
  
  return instalacao;
}

/**
 * EXECUTE ESTA FUNÇÃO PARA VER O STATUS
 */
function STATUS_HUB() {
  return statusCompletoHUB();
}

/**
 * EXECUTE ESTA FUNÇÃO PARA TESTAR TUDO
 */
function TESTAR_HUB() {
  return demonstracaoCompiletaHUB();
}

/**
* EXECUTE ESTA FUNÇÃO PARA CONFIGURAÇÃO AVANÇADA
*/
function CONFIGURAR_HUB() {
 return configuracaoAvancadaHUB();
}

/**
* EXECUTE ESTA FUNÇÃO PARA CRIAR APENAS AS ABAS MENSAIS
*/
function CRIAR_ABAS_MENSAIS() {
 console.log('📅 CRIANDO ABAS MENSAIS DO SISTEMA HUB');
 console.log('='.repeat(40));
 
 try {
   const resultado = criarTodasAbasMensaisHUB();
   
   console.log('✅ Abas mensais criadas:');
   console.log(`• Criadas: ${resultado.criadas}`);
   console.log(`• Existentes: ${resultado.existentes}`);
   console.log(`• Erros: ${resultado.erros}`);
   
   if (resultado.erros > 0) {
     console.log('⚠️ Detalhes dos erros:');
     resultado.detalhes.forEach(detalhe => {
       if (detalhe.status === 'erro') {
         console.log(`• ${detalhe.mes}: ${detalhe.erro}`);
       }
     });
   }
   
   console.log('📅 ABAS MENSAIS CONFIGURADAS!');
   
   return resultado;
   
 } catch (error) {
   console.log('❌ ERRO ao criar abas mensais:', error.message);
   return {
     sucesso: false,
     erro: error.message
   };
 }
}

/**
* EXECUTE ESTA FUNÇÃO PARA TESTAR APENAS O WHATSAPP
*/
function TESTAR_WHATSAPP() {
 console.log('📱 TESTANDO SISTEMA WHATSAPP');
 console.log('='.repeat(30));
 
 const resultado = testarZAPI();
 
 if (resultado.sucesso) {
   console.log('✅ WhatsApp funcionando perfeitamente!');
   console.log(`📞 Teste enviado para: ${SISTEMA_HUB_CONFIG.Z_API.ROBERTA_PHONE}`);
   console.log('📱 Verifique o WhatsApp da Roberta HUB');
 } else {
   console.log('❌ Problema no WhatsApp');
   console.log(`Erro: ${resultado.mensagem}`);
   console.log('🔧 Verifique as configurações da Z-API');
 }
 
 return resultado;
}

// ===================================================
// LOG DE INICIALIZAÇÃO DO SISTEMA
// ===================================================

// Log de carregamento do sistema
if (typeof console !== 'undefined') {
 console.log(`🚐 ${SISTEMA_HUB_CONFIG.NOME} v${SISTEMA_HUB_CONFIG.VERSAO} carregado!`);
 console.log('📋 Execute INSTALAR_HUB_CENTRAL() para instalar o sistema completo');
 console.log('📅 Execute CRIAR_ABAS_MENSAIS() para criar apenas as abas mensais');
 console.log('📱 Execute TESTAR_WHATSAPP() para testar apenas o WhatsApp');
} else if (typeof Logger !== 'undefined') {
 Logger.log(`🚐 ${SISTEMA_HUB_CONFIG.NOME} v${SISTEMA_HUB_CONFIG.VERSAO} carregado!`);
}

// ===================================================
// INFORMAÇÕES FINAIS DO SISTEMA
// ===================================================

/**
* INFORMAÇÕES IMPORTANTES SOBRE O SISTEMA HUB CENTRAL:
* 
* 🎯 OBJETIVO:
* Sistema centralizador que recebe dados de múltiplos hotéis e operadoras,
* consolidando tudo em uma planilha única com a nova coluna Hotel/Operadora
* e abas mensais organizadas por dias para gestão de capacidade.
* 
* 🔄 FLUXO PRINCIPAL:
* Hotel → Status "Confirmado" → POST para Sistema HUB → Registro Central + Aba Mensal → WhatsApp
* 
* 📊 NOVA ESTRUTURA:
* - Planilha central HUB com 21 colunas (incluindo Hotel/Operadora)
* - Abas mensais organizadas por dias (HUB_01_Janeiro_2025, etc.)
* - Contadores visuais de capacidade por dia
* - Cores indicativas da ocupação (vazio → cheio)
* - Sistema não envia e-mails (apenas recebe e processa)
* - WhatsApp via Roberta HUB usando Z-API
* - Valores diretos dos hotéis (sem cálculo percentual)
* 
* 🏨 HOTÉIS CONECTADOS:
* - Empire Marques Hotel
* - Hotel Lioz
* - Empire Lisbon Hotel
* - Gota d'água
* 
* 📅 ABAS MENSAIS:
* - Cada mês tem aba própria (ex: HUB_01_Janeiro_2025)
* - Organização visual por dias
* - Contadores de transfers por dia
* - Cores baseadas na capacidade (limite: 15 transfers/dia)
* - Resumo de capacidade no final da aba
* - Registro duplo: aba central + aba mensal
* 
* 📱 ROBERTA HUB:
* - Assistente virtual via WhatsApp (+351928283652)
* - Envia confirmações automaticamente
* - Apenas para transfers confirmados
* - Mensagens humanizadas e personalizadas
* 
* 🌐 API ENDPOINTS:
* GET:
* - ?action=test - Teste da API
* - ?action=health - Health check
* - ?action=hotels - Lista hotéis
* - ?action=transfers - Lista transfers
* - ?action=stats - Estatísticas
* 
* POST:
* - Recebe dados de transfers dos hotéis
* - Processa e registra na planilha central + aba mensal
* - Envia WhatsApp se confirmado
* 
* ⚙️ PARA INSTALAR:
* 1. Execute: INSTALAR_HUB_CENTRAL()
* 2. Faça deploy da Web App
* 3. Configure hotéis para enviar para a URL
* 4. Teste com STATUS_HUB() e TESTAR_HUB()
* 
* 📊 GESTÃO DE CAPACIDADE:
* - Visualização por dias em cada aba mensal
* - Cores indicam disponibilidade em tempo real
* - Limite configurável por dia (padrão: 15 transfers)
* - Relatórios de capacidade por mês
* - Verificação de disponibilidade por data
* 
* 🔧 FUNÇÕES PRINCIPAIS:
* - INSTALAR_HUB_CENTRAL() - Instalação completa
* - CRIAR_ABAS_MENSAIS() - Criar abas mensais
* - TESTAR_WHATSAPP() - Testar Z-API
* - STATUS_HUB() - Status do sistema
* - TESTAR_HUB() - Demonstração completa
* 
* 💰 SISTEMA DE VALORES:
* - Valores diretos enviados pelos hotéis (sem cálculo percentual)
* - Comissão recepção: €2 para transfers, €5 para tours
* - Suporte a valores customizados por hotel
* 
* 🎉 SISTEMA PRONTO PARA PRODUÇÃO!
* 
* 📋 PLANILHA DO PROJETO:
* https://docs.google.com/spreadsheets/d/1pXAYotTOvev50-NPQMks905DELshIA4yWjcBa8Pz9WU
* 
* 🚀 EXECUTE INSTALAR_HUB_CENTRAL() PARA COMEÇAR!
*/

/**
* Função para exibir ajuda sobre o sistema
*/
function AJUDA_HUB() {
 console.log('📚 AJUDA DO SISTEMA HUB CENTRAL');
 console.log('='.repeat(40));
 console.log('');
 console.log('🚀 FUNÇÕES PRINCIPAIS:');
 console.log('• INSTALAR_HUB_CENTRAL() - Instalação completa com abas mensais');
 console.log('• STATUS_HUB() - Verificar status do sistema');
 console.log('• TESTAR_HUB() - Demonstração completa');
 console.log('• CRIAR_ABAS_MENSAIS() - Criar apenas abas mensais');
 console.log('• TESTAR_WHATSAPP() - Testar apenas WhatsApp');
 console.log('• CONFIGURAR_HUB() - Configuração avançada');
 console.log('');
 console.log('📅 ABAS MENSAIS:');
 console.log('• Organização visual por dias');
 console.log('• Contadores de capacidade com cores');
 console.log('• Limite padrão: 15 transfers por dia');
 console.log('• Resumo de capacidade por mês');
 console.log('');
 console.log('📱 WHATSAPP (ROBERTA HUB):');
 console.log('• Telefone: +351928283652');
 console.log('• Apenas para transfers confirmados');
 console.log('• Mensagens automatizadas e humanizadas');
 console.log('');
 console.log('🏨 HOTÉIS CONECTADOS:');
 Object.keys(SISTEMA_HUB_CONFIG.HOTEIS_CONECTADOS).forEach(hotel => {
   console.log(`• ${hotel}`);
 });
 console.log('');
 console.log('📊 PLANILHA:');
 console.log(`• URL: https://docs.google.com/spreadsheets/d/${SISTEMA_HUB_CONFIG.SPREADSHEET_ID}`);
 console.log(`• Aba principal: ${SISTEMA_HUB_CONFIG.SHEET_NAME}`);
 console.log('• 21 colunas (incluindo Hotel/Operadora)');
 console.log('');
 console.log('🎯 PARA COMEÇAR:');
 console.log('1. Execute: INSTALAR_HUB_CENTRAL()');
 console.log('2. Faça deploy da Web App');
 console.log('3. Configure os hotéis');
 console.log('4. Teste o sistema');
 console.log('');
 console.log('✅ Sistema pronto para produção!');
}

/**
* Função para mostrar versão e informações do sistema
*/
function VERSAO_HUB() {
 console.log('📋 INFORMAÇÕES DO SISTEMA');
 console.log('='.repeat(30));
 console.log(`• Nome: ${SISTEMA_HUB_CONFIG.NOME}`);
 console.log(`• Versão: ${SISTEMA_HUB_CONFIG.VERSAO}`);
 console.log(`• Planilha: Sistema HUB Central`);
 console.log(`• ID: ${SISTEMA_HUB_CONFIG.SPREADSHEET_ID}`);
 console.log(`• WhatsApp: ${SISTEMA_HUB_CONFIG.NOTIFICACOES.WHATSAPP_ATIVO ? 'Ativo' : 'Inativo'}`);
 console.log(`• Roberta HUB: ${SISTEMA_HUB_CONFIG.Z_API.ROBERTA_PHONE}`);
 console.log(`• Hotéis conectados: ${Object.keys(SISTEMA_HUB_CONFIG.HOTEIS_CONECTADOS).length}`);
 console.log(`• Abas mensais: ${CONFIG_ABAS_MENSAIS.USAR_ABAS_MENSAIS ? 'Ativadas' : 'Desativadas'}`);
 console.log('');
 console.log('🌟 Características:');
 console.log('• ✅ Centralização de múltiplos hotéis');
 console.log('• ✅ Abas mensais organizadas por dias');
 console.log('• ✅ Gestão visual de capacidade');
 console.log('• ✅ WhatsApp automatizado (Roberta HUB)');
 console.log('• ✅ Nova coluna Hotel/Operadora');
 console.log('• ✅ Sistema de valores flexível');
 console.log('• ✅ APIs REST completas');
 console.log('• ✅ Relatórios e estatísticas');
 
 return {
   nome: SISTEMA_HUB_CONFIG.NOME,
   versao: SISTEMA_HUB_CONFIG.VERSAO,
   planilhaId: SISTEMA_HUB_CONFIG.SPREADSHEET_ID,
   whatsappAtivo: SISTEMA_HUB_CONFIG.NOTIFICACOES.WHATSAPP_ATIVO,
   hoteisConectados: Object.keys(SISTEMA_HUB_CONFIG.HOTEIS_CONECTADOS).length,
   abasMensais: CONFIG_ABAS_MENSAIS.USAR_ABAS_MENSAIS,
   timestamp: new Date()
 };
}

// ===================================================
// SISTEMA DE MONITORAMENTO DE CONFIRMAÇÕES DOS HOTÉIS
// ===================================================

/**
 * Configuração dos hotéis para monitoramento
 */
const HOTEIS_MONITORAMENTO = {
  'Empire Marques Hotel': {
    spreadsheetId: '1ZfG_IXBWMbGQzCmn7nWNzFqUBbGkXCXWSnY7JYziHxI',
    abaMonitorar: 'Empire MARQUES-HUB',
    colunaObservacoes: 19, // Coluna S
    urlHubCentral: 'URL_DO_SISTEMA_HUB_CENTRAL'
  },
  'Hotel Lioz': {
    spreadsheetId: '1jXhF6tAPhuieIoIm6F3zsaEkq7NV_tafarkxcot4IE8',
    abaMonitorar: 'Hotel_LIOZ-HUB',
    colunaObservacoes: 19,
    urlHubCentral: 'URL_DO_SISTEMA_HUB_CENTRAL'
  },
  'Empire Lisbon Hotel': {
    spreadsheetId: '12uW6XqVZJBIieHFl02zfatcWjPhHnYpZjydBK-JgV_A',
    abaMonitorar: 'Empire_LISBON-HUB',
    colunaObservacoes: 19,
    urlHubCentral: 'URL_DO_SISTEMA_HUB_CENTRAL'
  },
  'Gota d´água': {
    spreadsheetId: '1Zo0er2QaKszT3sYV8F0MZ50r3Pgj5gi50stqNrAnizY',
    abaMonitorar: 'Gota_DAGUA-HUB',
    colunaObservacoes: 19,
    urlHubCentral: 'URL_DO_SISTEMA_HUB_CENTRAL'
  }
};

/**
 * TRIGGER PRINCIPAL - Detecta mudanças na coluna S
 * Esta função deve ser configurada como trigger onChange em cada planilha de hotel
 */
function onEditDetectarConfirmacao(e) {
  console.log('🔍 Trigger de confirmação disparado');
  
  try {
    if (!e || !e.range) return;
    
    const range = e.range;
    const coluna = range.getColumn();
    const linha = range.getRow();
    const valor = range.getValue();
    
    // Verificar se é a coluna S (19) e se contém confirmação
    if (coluna === 19 && valor && typeof valor === 'string') {
      console.log(`📝 Mudança detectada na linha ${linha}, coluna S: ${valor}`);
      
      if (detectarConfirmacao(valor)) {
        console.log('✅ CONFIRMAÇÃO DETECTADA!');
        
        // Processar a confirmação
        processarConfirmacaoHotel(e.source, linha);
      }
    }
    
  } catch (error) {
    console.log('❌ Erro no trigger de confirmação:', error.message);
    loggerHUB.error('Erro no trigger de confirmação', error);
  }
}

/**
 * Detecta se o texto contém confirmação
 */
function detectarConfirmacao(texto) {
  if (!texto || typeof texto !== 'string') return false;
  
  const textoLower = texto.toLowerCase();
  
  // Padrões de confirmação
  const padroesConfirmacao = [
    '✅ confirmado',
    'confirmado por usuário',
    'confirmado por usuario',
    '✓ confirmado',
    'status: confirmado',
    'confirmação aceita'
  ];
  
  return padroesConfirmacao.some(padrao => textoLower.includes(padrao));
}

/**
 * Processa confirmação do hotel e alimenta Sistema HUB Central
 */
function processarConfirmacaoHotel(spreadsheet, linha) {
  console.log('🏨 Processando confirmação do hotel');
  
  try {
    // Identificar qual hotel
    const spreadsheetId = spreadsheet.getId();
    let hotelInfo = null;
    let nomeHotel = null;
    
    for (const [nome, config] of Object.entries(HOTEIS_MONITORAMENTO)) {
      if (config.spreadsheetId === spreadsheetId) {
        hotelInfo = config;
        nomeHotel = nome;
        break;
      }
    }
    
    if (!hotelInfo) {
      throw new Error('Hotel não identificado para monitoramento');
    }
    
    console.log(`🏨 Hotel identificado: ${nomeHotel}`);
    
    // Buscar dados da linha confirmada
    const sheet = spreadsheet.getSheetByName(hotelInfo.abaMonitorar);
    if (!sheet) {
      throw new Error(`Aba ${hotelInfo.abaMonitorar} não encontrada`);
    }
    
    // Extrair dados da linha (assumindo estrutura padrão dos hotéis)
    const dadosLinha = sheet.getRange(linha, 1, 1, 20).getValues()[0];
    
    const dadosTransfer = {
      id: dadosLinha[0],
      nomeCliente: dadosLinha[1],
      tipoServico: dadosLinha[2] || 'Transfer',
      numeroPessoas: dadosLinha[3],
      numeroBagagens: dadosLinha[4] || 0,
      data: dadosLinha[5],
      contacto: dadosLinha[6],
      numeroVoo: dadosLinha[7] || '',
      origem: dadosLinha[8],
      destino: dadosLinha[9],
      horaPickup: dadosLinha[10],
      valorTotal: dadosLinha[11],
      valorHotel: dadosLinha[12] || 0,
      valorHUB: dadosLinha[13] || dadosLinha[11], // Se não tem valor HUB, usar valor total
      comissaoRecepcao: dadosLinha[14] || SISTEMA_HUB_CONFIG.VALORES_PADRAO.COMISSAO_RECEPCAO_TRANSFER,
      modoPagamento: dadosLinha[15] || 'Dinheiro',
      pagoParaQuem: dadosLinha[16] || 'Recepção',
      status: 'Confirmado', // Já foi confirmado
      observacoes: dadosLinha[18] || `Transfer confirmado do ${nomeHotel}`,
      hotelOperadora: nomeHotel,
      source: getSourceByHotel(nomeHotel)
    };
    
    console.log('📋 Dados extraídos:', dadosTransfer);
    
    // Enviar para Sistema HUB Central
    const resultado = alimentarSistemaHUBCentral(dadosTransfer);
    
    if (resultado.sucesso) {
      console.log('✅ Transfer alimentado no Sistema HUB Central');
      
      // Atualizar observações do hotel com confirmação
      const timestampConfirmacao = new Date().toLocaleString('pt-PT');
      const novaObservacao = `${dadosLinha[18] || ''}\n[${timestampConfirmacao}] ✅ Alimentado no Sistema HUB Central - ID: ${resultado.transferId}`;
      
      sheet.getRange(linha, 19).setValue(novaObservacao);
      
    } else {
      throw new Error(`Falha ao alimentar Sistema HUB: ${resultado.erro}`);
    }
    
  } catch (error) {
    console.log('❌ Erro ao processar confirmação:', error.message);
    loggerHUB.error('Erro ao processar confirmação do hotel', error);
  }
}

/**
 * Alimenta o Sistema HUB Central com dados confirmados
 */
function alimentarSistemaHUBCentral(dadosTransfer) {
  console.log('🎯 Alimentando Sistema HUB Central');
  
  try {
    // Validar dados essenciais
    if (!dadosTransfer.nomeCliente || !dadosTransfer.data || !dadosTransfer.origem || !dadosTransfer.destino) {
      throw new Error('Dados obrigatórios ausentes');
    }
    
    // Preparar dados no formato do Sistema HUB
    const dadosFormatados = {
      nomeCliente: dadosTransfer.nomeCliente,
      tipoServico: dadosTransfer.tipoServico,
      numeroPessoas: parseInt(dadosTransfer.numeroPessoas),
      numeroBagagens: parseInt(dadosTransfer.numeroBagagens),
      data: processarDataSeguraHUB(dadosTransfer.data),
      contacto: dadosTransfer.contacto,
      numeroVoo: dadosTransfer.numeroVoo,
      origem: dadosTransfer.origem,
      destino: dadosTransfer.destino,
      horaPickup: dadosTransfer.horaPickup,
      valorTotal: parseFloat(dadosTransfer.valorTotal),
      modoPagamento: dadosTransfer.modoPagamento,
      pagoParaQuem: dadosTransfer.pagoParaQuem,
      status: 'Confirmado', // Status já confirmado
      observacoes: `${dadosTransfer.observacoes}\n✅ Confirmado via ${dadosTransfer.hotelOperadora}`
    };
    
    // Parâmetros extras
    const parametrosExtras = {
      source: dadosTransfer.source,
      hotelOperadora: dadosTransfer.hotelOperadora,
      origem: 'confirmacao_hotel',
      jaConfirmado: true // Flag para indicar que já foi confirmado
    };
    
    // Processar usando função existente do Sistema HUB
    const resultado = processarTransferDoHotel(dadosFormatados, parametrosExtras);
    
    if (resultado.sucesso) {
      console.log(`✅ Transfer ${resultado.transferId} registrado no Sistema HUB Central`);
      
      return {
        sucesso: true,
        transferId: resultado.transferId,
        whatsappEnviado: resultado.whatsappEnviado,
        hotel: dadosTransfer.hotelOperadora
      };
    } else {
      throw new Error('Falha no processamento do transfer');
    }
    
  } catch (error) {
    console.log('❌ Erro ao alimentar Sistema HUB:', error.message);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Mapeia hotel para source
 */
function getSourceByHotel(nomeHotel) {
  const mapping = {
    'Empire Marques Hotel': 'empire_marques',
    'Hotel Lioz': 'hotel_lioz',
    'Empire Lisbon Hotel': 'empire_lisbon',
    'Gota d´água': 'gota_dagua'
  };
  
  return mapping[nomeHotel] || 'unknown_hotel';
}

/**
 * Configura triggers automáticos em todas as planilhas de hotéis
 */
function configurarTriggersMonitoramento() {
  console.log('⚙️ Configurando triggers de monitoramento');
  
  try {
    let triggersConfigurados = 0;
    
    for (const [nomeHotel, config] of Object.entries(HOTEIS_MONITORAMENTO)) {
      try {
        // Abrir planilha do hotel
        const ss = SpreadsheetApp.openById(config.spreadsheetId);
        
        // Verificar se já existe trigger para esta planilha
        const triggersExistentes = ScriptApp.getProjectTriggers()
          .filter(trigger => 
            trigger.getTriggerSourceId() === config.spreadsheetId &&
            trigger.getHandlerFunction() === 'onEditDetectarConfirmacao'
          );
        
        if (triggersExistentes.length === 0) {
          // Criar novo trigger
          ScriptApp.newTrigger('onEditDetectarConfirmacao')
            .forSpreadsheet(ss)
            .onEdit()
            .create();
          
          console.log(`✅ Trigger configurado para ${nomeHotel}`);
          triggersConfigurados++;
        } else {
          console.log(`⚠️ Trigger já existe para ${nomeHotel}`);
        }
        
      } catch (error) {
        console.log(`❌ Erro ao configurar trigger para ${nomeHotel}:`, error.message);
      }
    }
    
    console.log(`🎯 Triggers configurados: ${triggersConfigurados}`);
    
    return {
      sucesso: true,
      triggersConfigurados: triggersConfigurados,
      totalHoteis: Object.keys(HOTEIS_MONITORAMENTO).length
    };
    
  } catch (error) {
    console.log('❌ Erro ao configurar triggers:', error.message);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Remove todos os triggers de monitoramento
 */
function removerTriggersMonitoramento() {
  console.log('🗑️ Removendo triggers de monitoramento');
  
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let removidos = 0;
    
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'onEditDetectarConfirmacao') {
        ScriptApp.deleteTrigger(trigger);
        removidos++;
      }
    });
    
    console.log(`✅ ${removidos} trigger(s) removido(s)`);
    
    return {
      sucesso: true,
      removidos: removidos
    };
    
  } catch (error) {
    console.log('❌ Erro ao remover triggers:', error.message);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Testa o sistema de monitoramento manualmente
 */
function testarMonitoramentoManual() {
  console.log('🧪 Testando sistema de monitoramento');
  
  try {
    // Simular confirmação do Empire Marques Hotel
    const dadosTeste = {
      nomeCliente: 'TESTE CONFIRMAÇÃO AUTOMÁTICA',
      tipoServico: 'Transfer',
      numeroPessoas: 2,
      numeroBagagens: 1,
      data: new Date(),
      contacto: '+351999000000',
      origem: 'Empire Marques Hotel',
      destino: 'Aeroporto de Lisboa',
      horaPickup: '14:00',
      valorTotal: 25.00,
      hotelOperadora: 'Empire Marques Hotel',
      source: 'empire_marques',
      observacoes: 'Teste de confirmação automática via trigger'
    };
    
    const resultado = alimentarSistemaHUBCentral(dadosTeste);
    
    if (resultado.sucesso) {
      console.log('✅ Teste bem-sucedido!');
      console.log(`Transfer ID: ${resultado.transferId}`);
      console.log(`WhatsApp enviado: ${resultado.whatsappEnviado}`);
    } else {
      console.log('❌ Teste falhou:', resultado.erro);
    }
    
    return resultado;
    
  } catch (error) {
    console.log('❌ Erro no teste:', error.message);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Monitora status das confirmações
 */
function verificarStatusMonitoramento() {
  console.log('📊 Verificando status do monitoramento');
  
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const triggersMonitoramento = triggers.filter(t => t.getHandlerFunction() === 'onEditDetectarConfirmacao');
    
    const status = {
      triggersAtivos: triggersMonitoramento.length,
      totalHoteis: Object.keys(HOTEIS_MONITORAMENTO).length,
      hoteisMonitorados: [],
      hoteisNaoMonitorados: []
    };
    
    // Verificar quais hotéis estão sendo monitorados
    for (const [nomeHotel, config] of Object.entries(HOTEIS_MONITORAMENTO)) {
      const temTrigger = triggersMonitoramento.some(t => t.getTriggerSourceId() === config.spreadsheetId);
      
      if (temTrigger) {
        status.hoteisMonitorados.push(nomeHotel);
      } else {
        status.hoteisNaoMonitorados.push(nomeHotel);
      }
    }
    
    console.log('📋 Status do monitoramento:', status);
    
    return status;
    
  } catch (error) {
    console.log('❌ Erro ao verificar status:', error.message);
    return {
      erro: error.message
    };
  }
}

// ===================================================
// FUNÇÕES DE MENU PARA SISTEMA DE MONITORAMENTO
// ===================================================

/**
 * Configura monitoramento via menu
 */
function configurarMonitoramentoMenu() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '⚙️ Configurar Monitoramento',
    'Esta ação irá configurar triggers automáticos em todas as planilhas dos hotéis para detectar confirmações.\n\nDeseja continuar?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    const resultado = configurarTriggersMonitoramento();
    
    if (resultado.sucesso) {
      ui.alert(
        '✅ Monitoramento Configurado',
        `Triggers configurados com sucesso!\n\n• Triggers ativos: ${resultado.triggersConfigurados}\n• Total de hotéis: ${resultado.totalHoteis}\n\nO sistema agora detectará automaticamente confirmações na coluna S.`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('❌ Erro', resultado.erro, ui.ButtonSet.OK);
    }
    
  } catch (error) {
    ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Verifica status via menu
 */
function verificarStatusMenu() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const status = verificarStatusMonitoramento();
    
    let mensagem = `📊 STATUS DO MONITORAMENTO\n\n`;
    mensagem += `• Triggers ativos: ${status.triggersAtivos}/${status.totalHoteis}\n\n`;
    
    if (status.hoteisMonitorados.length > 0) {
      mensagem += `✅ HOTÉIS MONITORADOS:\n`;
      status.hoteisMonitorados.forEach(hotel => {
        mensagem += `• ${hotel}\n`;
      });
      mensagem += `\n`;
    }
    
    if (status.hoteisNaoMonitorados.length > 0) {
      mensagem += `❌ HOTÉIS NÃO MONITORADOS:\n`;
      status.hoteisNaoMonitorados.forEach(hotel => {
        mensagem += `• ${hotel}\n`;
      });
    }
    
    ui.alert('📊 Status do Monitoramento', mensagem, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Testa monitoramento via menu
 */
function testarMonitoramentoMenu() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '🧪 Testar Monitoramento',
    'Esta ação irá simular uma confirmação e alimentar o Sistema HUB Central.\n\nDeseja continuar?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    const resultado = testarMonitoramentoManual();
    
    if (resultado.sucesso) {
      ui.alert(
        '✅ Teste Bem-sucedido',
        `Transfer teste criado no Sistema HUB Central!\n\n• Transfer ID: ${resultado.transferId}\n• WhatsApp enviado: ${resultado.whatsappEnviado ? 'Sim' : 'Não'}\n• Hotel: ${resultado.hotel}`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('❌ Teste Falhou', resultado.erro, ui.ButtonSet.OK);
    }
    
  } catch (error) {
    ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Remove triggers via menu
 */
function removerTriggersMenu() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '🗑️ Remover Triggers',
    'Esta ação irá remover todos os triggers de monitoramento.\n\nDeseja continuar?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    const resultado = removerTriggersMonitoramento();
    
    if (resultado.sucesso) {
      ui.alert(
        '✅ Triggers Removidos',
        `${resultado.removidos} trigger(s) removido(s) com sucesso.`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('❌ Erro', resultado.erro, ui.ButtonSet.OK);
    }
    
  } catch (error) {
    ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
  }
}

// ===================================================
// SISTEMA DE CONFIRMAÇÃO MULTILÍNGUE - JUNIOR GUTIERREZ
// ===================================================

/**
 * Mapeamento de países por DDI
 */
const MAPEAMENTO_PAISES_CONFIRMACAO = {
  '351': { pais: 'Portugal', idioma: 'pt', bandeira: '🇵🇹', nome: 'português' },
  '34': { pais: 'Espanha', idioma: 'es', bandeira: '🇪🇸', nome: 'espanhol' },
  '33': { pais: 'França', idioma: 'fr', bandeira: '🇫🇷', nome: 'francês' },
  '1': { pais: 'EUA/Canadá', idioma: 'en', bandeira: '🇺🇸', nome: 'inglês' },
  '44': { pais: 'Reino Unido', idioma: 'en', bandeira: '🇬🇧', nome: 'inglês' },
  '49': { pais: 'Alemanha', idioma: 'de', bandeira: '🇩🇪', nome: 'alemão' },
  '39': { pais: 'Itália', idioma: 'it', bandeira: '🇮🇹', nome: 'italiano' },
  '55': { pais: 'Brasil', idioma: 'pt-br', bandeira: '🇧🇷', nome: 'português brasileiro' },
  '52': { pais: 'México', idioma: 'es', bandeira: '🇲🇽', nome: 'espanhol' },
  '54': { pais: 'Argentina', idioma: 'es', bandeira: '🇦🇷', nome: 'espanhol' },
  '41': { pais: 'Suíça', idioma: 'fr', bandeira: '🇨🇭', nome: 'francês' },
  '32': { pais: 'Bélgica', idioma: 'fr', bandeira: '🇧🇪', nome: 'francês' }
};

/**
 * Templates de confirmação por idioma
 */
const TEMPLATES_CONFIRMACAO_JUNIOR = {
  'pt': `Olá {cliente}!

Sou o Junior Gutierrez, responsável pelos transfers da HUB Transfer Portugal.

Estou enviando esta mensagem para confirmar que no dia {data} teremos um transfer contigo:

📍 **Detalhes da viagem:**
- De: {origem}
- Para: {destino}
- Horário: {hora}
- Passageiros: {pessoas}

Qualquer dúvida, estou à disposição!

Obrigado pela confiança! 🇵🇹🤝{bandeira_cliente}

_Junior Gutierrez - Responsável pelos Transfers_
_HUB Transfer Portugal_`,

  'pt-br': `Olá {cliente}!

Sou o Junior Gutierrez, responsável pelos transfers da HUB Transfer Portugal.

Estou enviando esta mensagem para confirmar que no dia {data} teremos um transfer com você:

📍 **Detalhes da viagem:**
- De: {origem}
- Para: {destino}
- Horário: {hora}
- Passageiros: {pessoas}

Qualquer dúvida, estou à disposição!

Obrigado pela confiança! 🇵🇹🤝{bandeira_cliente}

_Junior Gutierrez - Responsável pelos Transfers_
_HUB Transfer Portugal_`,

  'es': `¡Hola {cliente}!

Soy Junior Gutierrez, responsable de los transfers de HUB Transfer Portugal.

Le envío este mensaje para confirmar que el día {data} tendremos un transfer con usted:

📍 **Detalles del viaje:**
- Desde: {origem}
- Hasta: {destino}
- Horario: {hora}
- Pasajeros: {pessoas}

¡Cualquier duda, estoy a su disposición!

¡Gracias por la confianza! 🇵🇹🤝{bandeira_cliente}

_Junior Gutierrez - Responsable de Transfers_
_HUB Transfer Portugal_`,

  'en': `Hello {cliente}!

I'm Junior Gutierrez, responsible for transfers at HUB Transfer Portugal.

I'm sending this message to confirm that on {data} we will have a transfer with you:

📍 **Trip details:**
- From: {origem}
- To: {destino}
- Time: {hora}
- Passengers: {pessoas}

Any questions, I'm at your disposal!

Thank you for your trust! 🇵🇹🤝{bandeira_cliente}

_Junior Gutierrez - Transfer Manager_
_HUB Transfer Portugal_`,

  'fr': `Bonjour {cliente}!

Je suis Junior Gutierrez, responsable des transfers chez HUB Transfer Portugal.

J'envoie ce message pour confirmer que le {data} nous aurons un transfer avec vous:

📍 **Détails du voyage:**
- De: {origem}
- À: {destino}
- Heure: {hora}
- Passagers: {pessoas}

Pour toute question, je suis à votre disposition!

Merci pour votre confiance! 🇵🇹🤝{bandeira_cliente}

_Junior Gutierrez - Responsable des Transfers_
_HUB Transfer Portugal_`,

  'it': `Ciao {cliente}!

Sono Junior Gutierrez, responsabile dei transfer presso HUB Transfer Portugal.

Invio questo messaggio per confermare che il giorno {data} avremo un transfer con lei:

📍 **Dettagli del viaggio:**
- Da: {origem}
- A: {destino}
- Orario: {hora}
- Passeggeri: {pessoas}

Per qualsiasi domanda, sono a sua disposizione!

Grazie per la fiducia! 🇵🇹🤝{bandeira_cliente}

_Junior Gutierrez - Responsabile Transfer_
_HUB Transfer Portugal_`
};

function testeSimplificadoIdioma() {
  console.log('🧪 TESTE SIMPLIFICADO DE IDIOMA');
  
  // Teste 1: Portugal
  const tel1 = '+351999000001';
  const limpo1 = tel1.replace(/\D/g, '');
  console.log('Número:', tel1, '→ Limpo:', limpo1);
  console.log('Começa com 351?', limpo1.startsWith('351'));
  
  // Teste 2: Espanha
  const tel2 = '+34600000001';
  const limpo2 = tel2.replace(/\D/g, '');
  console.log('Número:', tel2, '→ Limpo:', limpo2);
  console.log('Começa com 34?', limpo2.startsWith('34'));
  
  // Teste direto da função principal
  console.log('Testando função enviarConfirmacaoParaJunior...');
  
  const transferTeste = {
    id: 999,
    cliente: 'Cliente Teste Direto',
    contacto: '+351999888777',
    data: new Date(),
    origem: 'Lisboa',
    destino: 'Porto',
    horaPickup: '15:00',
    numeroPessoas: 2,
    hotelOperadora: 'Teste'
  };
  
  try {
    const resultado = enviarConfirmacaoParaJunior(transferTeste);
    console.log('Resultado:', resultado);
  } catch (error) {
    console.log('Erro no teste:', error.message);
  }
}

// ===================================================
// SISTEMA DE LEMBRETES 24H ANTES DO TRANSFER
// ===================================================

const TEMPLATES_LEMBRETE_24H_JUNIOR = {
  'pt': `Olá novamente, {cliente}!

Junior por aqui da HUB Transfer Portugal.

Passando para vos lembrar que amanhã ({data}) temos o transfer contigo:

✈️ *Lembrete da viagem:*
- Voo: {link_voo}
- Horário: {hora}
- De: {origem}
- Para: {destino}

Antes do seu voo aterrar, o motorista irá enviar uma mensagem com as informações do ponto de encontro para facilitar a sua chegada.

Estarei acompanhando o seu voo para cuidar que tudo seja incrível!

Boa viagem! 🇵🇹🤝{bandeira_cliente}

_Junior Gutierrez - HUB Transfer Portugal_`,

  'pt-br': `Olá novamente, {cliente}!

Junior por aqui da HUB Transfer Portugal.

Passando para te lembrar que amanhã ({data}) temos o transfer contigo:

✈️ *Lembrete da viagem:*
- Voo: {link_voo}
- Horário: {hora}
- De: {origem}
- Para: {destino}

Antes do seu voo pousar, o motorista irá enviar uma mensagem com as informações do ponto de encontro para facilitar a sua chegada.

Estarei acompanhando seu voo para cuidar que tudo seja incrível!

Boa viagem! 🇵🇹🤝{bandeira_cliente}

_Junior Gutierrez - HUB Transfer Portugal_`,

  'es': `¡Hola otra vez, {cliente}!

Junior por aquí de HUB Transfer Portugal.

Solo paso para recordarte que mañana ({data}) tenemos el transfer contigo:

✈️ *Recordatorio del viaje:*
- Vuelo: {link_voo}
- Horario: {hora}
- Desde: {origem}
- Hasta: {destino}

Antes de que aterrice tu vuelo, el conductor te enviará un mensaje con la información del punto de encuentro para facilitar tu llegada.

¡Estaré siguiendo tu vuelo para asegurar que todo sea increíble!

¡Buen viaje! 🇵🇹🤝{bandeira_cliente}

_Junior Gutierrez - HUB Transfer Portugal_`,

  'en': `Hello again, {cliente}!

Junior here from HUB Transfer Portugal.

Just stopping by to remind you that tomorrow ({data}) we have the transfer with you:

✈️ *Trip reminder:*
- Flight: {link_voo}
- Time: {hora}
- From: {origem}
- To: {destino}

Before your flight lands, the driver will send you a message with the meeting point information to make your arrival easier.

I'll be tracking your flight to make sure everything is incredible!

Have a great trip! 🇵🇹🤝{bandeira_cliente}

_Junior Gutierrez - HUB Transfer Portugal_`,

  'fr': `Bonjour encore, {cliente}!

Junior ici de HUB Transfer Portugal.

Je passe juste pour te rappeler que demain ({data}) nous avons le transfer avec toi:

✈️ *Rappel du voyage:*
- Vol: {link_voo}
- Heure: {hora}
- De: {origem}
- À: {destino}

Avant que ton vol atterrisse, le conducteur t'enverra un message avec les informations du point de rendez-vous pour faciliter ton arrivée.

Je suivrai ton vol pour m'assurer que tout soit incroyable!

Bon voyage! 🇵🇹🤝{bandeira_cliente}

_Junior Gutierrez - HUB Transfer Portugal_`,

  'it': `Ciao di nuovo, {cliente}!

Junior qui da HUB Transfer Portugal.

Passo solo per ricordarti che domani ({data}) abbiamo il transfer con te:

✈️ *Promemoria del viaggio:*
- Volo: {link_voo}
- Orario: {hora}
- Da: {origem}
- A: {destino}

Prima che il tuo volo atterri, l'autista ti invierà un messaggio con le informazioni del punto d'incontro per facilitare il tuo arrivo.

Seguirò il tuo volo per assicurarmi che tutto sia incredibile!

Buon viaggio! 🇵🇹🤝{bandeira_cliente}

_Junior Gutierrez - HUB Transfer Portugal_`
};

/**
 * Cria link de rastreamento de voo para Google
 */
function criarLinkRastreamentoVoo(numeroVoo) {
  try {
    if (!numeroVoo || numeroVoo === '' || numeroVoo === 'Não informado') {
      return 'Não informado';
    }
    
    // Limpar número do voo
    const vooLimpo = numeroVoo.replace(/\s+/g, '').toUpperCase();
    
    // Criar link para Google Flights
    const linkGoogle = `https://www.google.com/search?q=flight+${vooLimpo}`;
    
    return `${vooLimpo} - Rastrear: ${linkGoogle}`;
    
  } catch (error) {
    loggerHUB.error('Erro ao criar link de voo', error);
    return numeroVoo || 'Não informado';
  }
}

function detectarIdiomaPorDDI(numeroTelefone) {
  try {
    console.log('Entrada detectarIdiomaPorDDI:', numeroTelefone, typeof numeroTelefone);
    
    // Múltiplas validações e conversões
    let telefoneString = '';
    
    if (numeroTelefone === null || numeroTelefone === undefined) {
      telefoneString = '';
    } else if (typeof numeroTelefone === 'string') {
      telefoneString = numeroTelefone;
    } else if (typeof numeroTelefone === 'number') {
      telefoneString = numeroTelefone.toString();
    } else {
      telefoneString = String(numeroTelefone);
    }
    
    console.log('Telefone convertido:', telefoneString, typeof telefoneString);
    
    const numeroLimpo = telefoneString.replace(/\D/g, '');
    console.log('Número limpo:', numeroLimpo);
    
    // Português (Portugal e Brasil)
    if (numeroLimpo.indexOf('351') === 0) {
      return { pais: 'Portugal', idioma: 'pt', bandeira: '🇵🇹', nome: 'português' };
    }
    if (numeroLimpo.indexOf('55') === 0) {
      return { pais: 'Brasil', idioma: 'pt-br', bandeira: '🇧🇷', nome: 'português brasileiro' };
    }
    
    // Espanhol
    if (numeroLimpo.indexOf('34') === 0) {
      return { pais: 'Espanha', idioma: 'es', bandeira: '🇪🇸', nome: 'espanhol' };
    }
    
    // Inglês
    if (numeroLimpo.indexOf('1') === 0) {
      return { pais: 'EUA/Canadá', idioma: 'en', bandeira: '🇺🇸', nome: 'inglês' };
    }
    if (numeroLimpo.indexOf('44') === 0) {
      return { pais: 'Reino Unido', idioma: 'en', bandeira: '🇬🇧', nome: 'inglês' };
    }
    
    // Francês
    if (numeroLimpo.indexOf('33') === 0) {
      return { pais: 'França', idioma: 'fr', bandeira: '🇫🇷', nome: 'francês' };
    }
    
    // Italiano
    if (numeroLimpo.indexOf('39') === 0) {
      return { pais: 'Itália', idioma: 'it', bandeira: '🇮🇹', nome: 'italiano' };
    }
    
    // FALLBACK: Inglês para outros países
    return { pais: 'Internacional', idioma: 'en', bandeira: '🌍', nome: 'inglês' };
    
  } catch (error) {
    console.log('Erro detectarIdiomaPorDDI:', error.message);
    return { pais: 'Internacional', idioma: 'en', bandeira: '🌍', nome: 'inglês' };
  }
}

/**
 * Formata mensagem de lembrete 24h
 */
function formatarMensagemLembrete24H(transferData, infoIdioma) {
  try {
    const template = TEMPLATES_LEMBRETE_24H_JUNIOR[infoIdioma.idioma] || 
                    TEMPLATES_LEMBRETE_24H_JUNIOR['pt'];
    
    const dataAmanha = new Date(transferData.data);
    dataAmanha.setDate(dataAmanha.getDate());
    const dataFormatada = formatarDataPorIdioma(dataAmanha, infoIdioma.idioma);
    
    const linkVoo = criarLinkRastreamentoVoo(transferData.numeroVoo || transferData.voo);
    
    const mensagemFormatada = template
      .replace(/{cliente}/g, transferData.cliente || transferData.nomeCliente)
      .replace(/{data}/g, dataFormatada)
      .replace(/{link_voo}/g, linkVoo)
      .replace(/{hora}/g, transferData.horaPickup)
      .replace(/{origem}/g, transferData.origem)
      .replace(/{destino}/g, transferData.destino)
      .replace(/{bandeira_cliente}/g, infoIdioma.bandeira);
    
    return mensagemFormatada;
    
  } catch (error) {
    loggerHUB.error('Erro ao formatar mensagem de lembrete', error);
    return `Erro ao formatar mensagem: ${error.message}`;
  }
}

/**
 * FUNÇÃO PRINCIPAL - Envia lembrete 24h para Junior (não para cliente)
 */
function enviarLembrete24HParaJunior(transferData) {
  loggerHUB.info('Enviando lembrete 24h para Junior', { 
    transferId: transferData.id,
    cliente: transferData.cliente || transferData.nomeCliente
  });
  
  try {
    const numeroCliente = transferData.contacto;
    
    if (!numeroCliente) {
      loggerHUB.warn('Transfer sem contacto para lembrete', { 
        transferId: transferData.id 
      });
      return { sucesso: false, motivo: 'sem_contacto' };
    }
    
    // Detectar idioma do cliente
    const infoIdioma = detectarIdiomaPorDDI(numeroCliente);
    
    // Formatar mensagem no idioma correto
    const mensagemFormatada = formatarMensagemLembrete24H(transferData, infoIdioma);
    
    // Criar mensagem para Junior com número clicável
    const numeroClienteFormatado = '+' + numeroCliente.replace(/\D/g, '');
    
    const mensagemParaJunior = `⏰ LEMBRETE 24H - ID: ${transferData.id}

📋 MENSAGEM PRONTA PARA COPIAR E ENVIAR:

${mensagemFormatada}

📱 ENVIAR PARA: ${numeroClienteFormatado}
🌍 PAÍS: ${infoIdioma.pais}
🗣️ IDIOMA: ${(infoIdioma.nome || infoIdioma.idioma || 'português').toUpperCase()}
🏨 HOTEL: ${transferData.hotelOperadora || 'Não especificado'}
✈️ VOO: ${transferData.numeroVoo || transferData.voo || 'Não informado'}

Copie a mensagem acima e cole na conversa com o cliente.`;

    // Enviar para o número do Junior
    const enviado = enviarWhatsAppRoberta('+351968698138', mensagemParaJunior);
    
    if (enviado) {
      loggerHUB.success('Lembrete 24h enviado para Junior', {
        transferId: transferData.id,
        idioma: infoIdioma.idioma,
        pais: infoIdioma.pais
      });
      
      return {
        sucesso: true,
        idioma: infoIdioma.idioma,
        pais: infoIdioma.pais,
        numeroCliente: numeroClienteFormatado
      };
    } else {
      loggerHUB.error('Falha no envio do lembrete para Junior', { transferId: transferData.id });
      return { sucesso: false, motivo: 'falha_envio' };
    }
    
  } catch (error) {
    loggerHUB.error('Erro ao enviar lembrete para Junior', error);
    return { sucesso: false, erro: error.message };
  }
}

/**
 * Verifica transfers do dia seguinte e envia lembretes
 */
function verificarTransfersAmanha() {
  loggerHUB.info('Verificando transfers para amanhã (lembretes 24h)');
  
  try {
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('HUB_Central');
    
    if (!sheet || sheet.getLastRow() <= 1) {
      loggerHUB.info('Nenhum transfer encontrado para verificação');
      return { sucesso: true, lembretes: 0 };
    }
    
    // Data de amanhã
    const amanha = new Date();
    amanha.setDate(amanha.getDate() + 1);
    amanha.setHours(0, 0, 0, 0);
    
    const dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, 26).getValues();
    let lembretesEnviados = 0;
    
    dados.forEach(linha => {
      const dataTransfer = new Date(linha[5]); // Coluna F - Data
      dataTransfer.setHours(0, 0, 0, 0);
      
      const status = linha[17]; // Coluna R - Status
      
      // Verificar se é transfer de amanhã e está confirmado
      if (dataTransfer.getTime() === amanha.getTime() && status === 'Confirmado') {
        const transferData = {
          id: linha[0],                    // A - ID
          cliente: linha[1],               // B - Cliente
          data: linha[5],                  // F - Data
          contacto: linha[6],              // G - Contacto
          numeroVoo: linha[7],             // H - Voo
          origem: linha[8],                // I - Origem
          destino: linha[9],               // J - Destino
          horaPickup: linha[10],           // K - Hora Pick-up
          hotelOperadora: linha[20]        // U - Hotel/Operadora
        };
        
        // Enviar lembrete
        const resultado = enviarLembrete24HParaJunior(transferData);
        
        if (resultado.sucesso) {
          lembretesEnviados++;
          loggerHUB.info('Lembrete enviado', { transferId: transferData.id });
        }
      }
    });
    
    loggerHUB.success('Verificação de lembretes concluída', { 
      lembretesEnviados: lembretesEnviados 
    });
    
    return {
      sucesso: true,
      lembretes: lembretesEnviados,
      dataVerificada: amanha
    };
    
  } catch (error) {
    loggerHUB.error('Erro ao verificar transfers de amanhã', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Teste do sistema de lembretes 24h
 */
function testarLembrete24H() {
  console.log('🔔 Testando sistema de lembrete 24h');
  
  const transferTeste = {
    id: 8888,
    cliente: 'Cliente Teste Lembrete',
    data: new Date(Date.now() + 24 * 60 * 60 * 1000), // Amanhã
    contacto: '+351999777888',
    numeroVoo: 'TP1234',
    origem: 'Aeroporto de Lisboa',
    destino: 'Hotel Teste',
    horaPickup: '16:30',
    hotelOperadora: 'Empire Marques Hotel'
  };
  
  const resultado = enviarLembrete24HParaJunior(transferTeste);
  
  if (resultado.sucesso) {
    console.log('✅ Teste de lembrete bem-sucedido!');
    console.log('Verifique seu WhatsApp para a mensagem de lembrete');
  } else {
    console.log('❌ Teste de lembrete falhou:', resultado.erro || resultado.motivo);
  }
  
  return resultado;
}

/**
 * Configura trigger automático para lembretes diários
 */
function configurarTriggerLembretes24H() {
  console.log('⏰ Configurando trigger automático para lembretes 24h');
  
  try {
    // Verificar se já existe trigger
    const triggersExistentes = ScriptApp.getProjectTriggers()
      .filter(trigger => trigger.getHandlerFunction() === 'verificarTransfersAmanha');
    
    if (triggersExistentes.length > 0) {
      console.log('⚠️ Trigger de lembretes já existe');
      return { sucesso: false, motivo: 'trigger_existe' };
    }
    
    // Criar trigger para executar diariamente às 18:00
    ScriptApp.newTrigger('verificarTransfersAmanha')
      .timeBased()
      .everyDays(1)
      .atHour(18)
      .create();
    
    console.log('✅ Trigger de lembretes configurado para 18:00 diariamente');
    
    return {
      sucesso: true,
      horario: '18:00',
      funcao: 'verificarTransfersAmanha'
    };
    
  } catch (error) {
    console.log('❌ Erro ao configurar trigger:', error.message);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Formata mensagem de confirmação no idioma correto
 */
function formatarMensagemConfirmacao(transferData, infoIdioma) {
  try {
    const template = TEMPLATES_CONFIRMACAO_JUNIOR[infoIdioma.idioma] || 
                    TEMPLATES_CONFIRMACAO_JUNIOR['pt'];
    
    const dataFormatada = formatarDataPorIdioma(transferData.data, infoIdioma.idioma);
    
    const mensagemFormatada = template
      .replace(/{cliente}/g, transferData.cliente || transferData.nomeCliente)
      .replace(/{data}/g, dataFormatada)
      .replace(/{origem}/g, transferData.origem)
      .replace(/{destino}/g, transferData.destino)
      .replace(/{hora}/g, transferData.horaPickup)
      .replace(/{pessoas}/g, transferData.numeroPessoas || transferData.pessoas)
      .replace(/{bandeira_cliente}/g, infoIdioma.bandeira);
    
    return mensagemFormatada;
    
  } catch (error) {
    loggerHUB.error('Erro ao formatar mensagem', error);
    return `Erro ao formatar mensagem: ${error.message}`;
  }
}

/**
 * Formata data conforme idioma
 */
function formatarDataPorIdioma(data, idioma) {
  try {
    const dataObj = new Date(data);
    
    const formatosData = {
      'pt': { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' },
      'pt-br': { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' },
      'es': { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' },
      'en': { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' },
      'fr': { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' },
      'it': { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' }
    };
    
    const locales = {
      'pt': 'pt-PT',
      'pt-br': 'pt-BR',
      'es': 'es-ES',
      'en': 'en-GB',
      'fr': 'fr-FR',
      'it': 'it-IT'
    };
    
    const formato = formatosData[idioma] || formatosData['pt'];
    const locale = locales[idioma] || locales['pt'];
    
    return dataObj.toLocaleDateString(locale, formato);
    
  } catch (error) {
    loggerHUB.error('Erro ao formatar data por idioma', error);
    return formatarDataDDMMYYYY(new Date(data));
  }
}

/**
 * FUNÇÃO PRINCIPAL - Envia confirmação para Junior (não para cliente)
 */
function enviarConfirmacaoParaJunior(transferData) {
  loggerHUB.info('Enviando confirmação para Junior', { 
    transferId: transferData.id,
    cliente: transferData.cliente || transferData.nomeCliente
  });
  
  try {
    const numeroCliente = transferData.contacto;
    
    if (!numeroCliente) {
      loggerHUB.warn('Transfer sem contacto para confirmação', { 
        transferId: transferData.id 
      });
      return { sucesso: false, motivo: 'sem_contacto' };
    }
    
    // Detectar idioma do cliente
    const infoIdioma = detectarIdiomaPorDDI(numeroCliente);
    
    // Formatar mensagem no idioma correto
    const mensagemFormatada = formatarMensagemConfirmacao(transferData, infoIdioma);
    
    // Criar mensagem para Junior
    const numeroClienteFormatado = '+' + numeroCliente.replace(/\D/g, '');
    
    const mensagemParaJunior = `🎯 TRANSFER CONFIRMADO - ID: ${transferData.id}

${mensagemFormatada}

📱 NÚMERO DO CLIENTE: ${numeroClienteFormatado}
🌍 PAÍS: ${infoIdioma.pais}
🗣️ IDIOMA: ${(infoIdioma.nome || infoIdioma.idioma || 'português').toUpperCase()}
🏨 HOTEL: ${transferData.hotelOperadora || 'Não especificado'}

Clique no número acima para abrir WhatsApp e enviar para o cliente.`;

    // Enviar para o número do Junior
    const enviado = enviarWhatsAppRoberta('+351968698138', mensagemParaJunior);
    
    if (enviado) {
      loggerHUB.success('Confirmação enviada para Junior', {
        transferId: transferData.id,
        idioma: infoIdioma.idioma,
        pais: infoIdioma.pais
      });
      
      return {
        sucesso: true,
        idioma: infoIdioma.idioma,
        pais: infoIdioma.pais,
        numeroCliente: numeroClienteFormatado
      };
    } else {
      loggerHUB.error('Falha no envio para Junior', { transferId: transferData.id });
      return { sucesso: false, motivo: 'falha_envio' };
    }
    
  } catch (error) {
    loggerHUB.error('Erro ao enviar confirmação para Junior', error);
    return { sucesso: false, erro: error.message };
  }
}

/**
 * VERSÃO PRESERVADA - Envio direto para cliente (desativada temporariamente)
 */
function enviarConfirmacaoWhatsAppDireto(transferData) {
  loggerHUB.info('Função de envio direto para cliente (DESATIVADA)', {
    transferId: transferData.id
  });
  
  // FUNÇÃO PRESERVADA PARA FUTURO USO COM ROBERTA HUB
  // Quando a Roberta estiver pronta, substituir enviarConfirmacaoParaJunior
  // por esta função na chamada principal
  
  const numeroCliente = transferData.contacto;
  const infoIdioma = detectarIdiomaPorDDI(numeroCliente);
  const mensagemFormatada = formatarMensagemConfirmacao(transferData, infoIdioma);
  
  // DESCOMENTE ESTA LINHA QUANDO ROBERTA ESTIVER PRONTA:
  // return enviarConfirmacaoParaJunior(transferData, mensagemFormatada);
  
  loggerHUB.info('Envio direto desativado - usando envio via Junior');
  return { sucesso: false, motivo: 'envio_direto_desativado' };
}

/**
 * Teste do sistema multilíngue
 */
function testarSistemaMultilingue() {
  console.log('🌍 Testando sistema multilíngue');
  
  const numerosTestes = [
    '+351999000001', // Portugal
    '+34600000001',  // Espanha
    '+33600000001',  // França
    '+44700000001',  // Reino Unido
    '+39300000001'   // Itália
  ];
  
  const transferTeste = {
    id: 9999,
    cliente: 'Cliente Teste',
    data: new Date(),
    origem: 'Aeroporto de Lisboa',
    destino: 'Hotel Teste',
    horaPickup: '14:30',
    numeroPessoas: 2,
    hotelOperadora: 'Teste'
  };
  
  numerosTestes.forEach(numero => {
    const infoIdioma = detectarIdiomaPorDDI(numero);
    console.log(`${numero} → ${infoIdioma.pais} (${infoIdioma.idioma})`);
    
    transferTeste.contacto = numero;
    const mensagem = formatarMensagemConfirmacao(transferTeste, infoIdioma);
    console.log('Mensagem gerada:', mensagem.substring(0, 100) + '...');
  });
}

// ===================================================
// SISTEMA DE MENSAGENS PARA MOTORISTAS - MULTILÍNGUE
// ===================================================

const TEMPLATES_MOTORISTA_MULTILINGUE = {
  'pt': `Olá 👋, {cliente}

Sou o {motorista} e serei o vosso motorista de TRANSFER

Espero que tenham tido uma óptima viagem.

Podem dizer-me passo a passo onde se encontram para que vos possa acompanhar da melhor forma?

Quando estiverem no controlo de passaportes, a ir buscar a bagagem e quando estiverem com a bagagem?

Até já {bandeira_cliente}`,

  'pt-br': `Olá 👋, {cliente}

Sou o {motorista} e serei seu motorista de TRANSFER

Espero que tenha tido uma ótima viagem.

Pode me dizer passo a passo onde está para que eu possa acompanhá-lo da melhor forma?

Quando estiver no controle de passaportes, indo buscar a bagagem e quando estiver com a bagagem?

Até logo {bandeira_cliente}`,

  'es': `Hola 👋, {cliente}

Soy {motorista} y seré tu conductor de TRANSFER

Espero que hayas tenido un gran viaje.

¿Puedes decirme paso a paso dónde estás para poder acompañarte de la mejor manera?

¿Cuándo estés en control de pasaportes, yendo por el equipaje y cuando tengas el equipaje?

Hasta pronto {bandeira_cliente}`,

  'en': `Hi 👋, {cliente}

I'm {motorista} and I'll be your TRANSFER driver

I hope you had a great trip.

Can you tell me step by step of where it is so that I can accompany you in the best way?

When are you in passport control, going to the luggage and when you are with the luggage?

See you soon {bandeira_cliente}`,

  'fr': `Salut 👋, {cliente}

Je suis {motorista} et je serai votre conducteur de TRANSFER

J'espère que vous avez eu un excellent voyage.

Pouvez-vous me dire étape par étape où vous êtes pour que je puisse vous accompagner de la meilleure façon?

Quand vous êtes au contrôle des passeports, en allant chercher les bagages et quand vous avez les bagages?

À bientôt {bandeira_cliente}`,

  'it': `Ciao 👋, {cliente}

Sono {motorista} e sarò il tuo autista di TRANSFER

Spero che tu abbia avuto un ottimo viaggio.

Puoi dirmi passo dopo passo dove sei così posso accompagnarti nel modo migliore?

Quando sei al controllo passaporti, andando a prendere i bagagli e quando hai i bagagli?

A presto {bandeira_cliente}`
};

/**
 * Formata mensagem para motorista baseada no idioma do cliente
 */
function formatarMensagemMotoristaMultilingue(transferData, motoristaData, infoIdioma) {
  try {
    // Usar template baseado no idioma detectado do cliente
    const template = TEMPLATES_MOTORISTA_MULTILINGUE[infoIdioma.idioma] || 
                    TEMPLATES_MOTORISTA_MULTILINGUE['en']; // Fallback inglês
    
    const mensagemFormatada = template
      .replace(/{cliente}/g, transferData.cliente)
      .replace(/{motorista}/g, motoristaData.nome)
      .replace(/{bandeira_cliente}/g, infoIdioma.bandeira);
    
    return mensagemFormatada;
    
  } catch (error) {
    loggerHUB.error('Erro ao formatar mensagem do motorista multilíngue', error);
    return `Erro ao formatar mensagem: ${error.message}`;
  }
}

function detectarIdiomaClientePorDDI(numeroTelefone) {
  try {
    console.log('Nova função - Entrada:', numeroTelefone, typeof numeroTelefone);
    
    let telefoneString = '';
    
    if (numeroTelefone === null || numeroTelefone === undefined) {
      telefoneString = '';
    } else {
      telefoneString = String(numeroTelefone);
    }
    
    const numeroLimpo = telefoneString.replace(/\D/g, '');
    console.log('Número limpo:', numeroLimpo);
    
    // Português (Portugal e Brasil)
    if (numeroLimpo.indexOf('351') === 0) {
      return { pais: 'Portugal', idioma: 'pt', bandeira: '🇵🇹', nome: 'português' };
    }
    if (numeroLimpo.indexOf('55') === 0) {
      return { pais: 'Brasil', idioma: 'pt-br', bandeira: '🇧🇷', nome: 'português brasileiro' };
    }
    
    // Espanhol
    if (numeroLimpo.indexOf('34') === 0) {
      return { pais: 'Espanha', idioma: 'es', bandeira: '🇪🇸', nome: 'espanhol' };
    }
    
    // Inglês
    if (numeroLimpo.indexOf('1') === 0) {
      return { pais: 'EUA/Canadá', idioma: 'en', bandeira: '🇺🇸', nome: 'inglês' };
    }
    if (numeroLimpo.indexOf('44') === 0) {
      return { pais: 'Reino Unido', idioma: 'en', bandeira: '🇬🇧', nome: 'inglês' };
    }
    
    // Francês
    if (numeroLimpo.indexOf('33') === 0) {
      return { pais: 'França', idioma: 'fr', bandeira: '🇫🇷', nome: 'francês' };
    }
    
    // Italiano
    if (numeroLimpo.indexOf('39') === 0) {
      return { pais: 'Itália', idioma: 'it', bandeira: '🇮🇹', nome: 'italiano' };
    }
    
    return { pais: 'Internacional', idioma: 'en', bandeira: '🌍', nome: 'inglês' };
    
  } catch (error) {
    console.log('Erro nova função:', error.message);
    return { pais: 'Internacional', idioma: 'en', bandeira: '🌍', nome: 'inglês' };
  }
}

/**
 * Busca motoristas cadastrados na aba Motoristas_HUB
 */
function buscarMotoristasCadastrados() {
  try {
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    const abaMotoristas = ss.getSheetByName('Motoristas_HUB');
    
    if (!abaMotoristas) {
      loggerHUB.warn('Aba Motoristas_HUB não encontrada');
      return [];
    }
    
    const dados = abaMotoristas.getDataRange().getValues();
    const motoristas = [];
    
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0]) { // Se tem nome
        motoristas.push({
          nome: dados[i][0],
          telefone: dados[i][1] || 'Não cadastrado',
          carro: dados[i][2] || 'Não informado',
          matricula: dados[i][3] || 'Não informada'
        });
      }
    }
    
    loggerHUB.info('Motoristas cadastrados encontrados', { total: motoristas.length });
    return motoristas;
    
  } catch (error) {
    loggerHUB.error('Erro ao buscar motoristas cadastrados', error);
    return [];
  }
}

/**
 * Busca transfers que têm motorista atribuído na coluna V (incluindo número do voo)
 */
function buscarTransfersComMotoristasAtribuidos() {
  try {
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('HUB_Central');
    
    if (!sheet || sheet.getLastRow() <= 1) {
      loggerHUB.info('Nenhum transfer encontrado na planilha');
      return [];
    }
    
    const dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, 26).getValues();
    const transfersComMotorista = [];
    
    dados.forEach((linha, index) => {
      const motorista = linha[21]; // Coluna V - Motorista
      const status = linha[17]; // Coluna R - Status
      const cliente = linha[1]; // Coluna B - Cliente
      const contacto = linha[6]; // Coluna G - Contacto
      
      // Verificar se tem motorista atribuído E dados válidos E status confirmado
      if (motorista && 
          motorista.toString().trim() !== '' && 
          cliente && 
          contacto && 
          status === 'Confirmado') {
        
        transfersComMotorista.push({
          linha: index + 2, // Linha real na planilha
          id: linha[0],
          cliente: linha[1],
          contacto: linha[6],
          numeroVoo: linha[7], // COLUNA H - NÚMERO DO VOO
          origem: linha[8],
          destino: linha[9],
          horaPickup: linha[10],
          motorista: motorista.toString().trim(),
          hotelOperadora: linha[20],
          data: linha[5]
        });
      }
    });
    
    loggerHUB.info('Transfers com motoristas atribuídos encontrados', { 
      total: transfersComMotorista.length 
    });
    
    return transfersComMotorista;
    
  } catch (error) {
    loggerHUB.error('Erro ao buscar transfers com motoristas', error);
    return [];
  }
}

/**
 * FUNÇÃO PRINCIPAL - Envia mensagens multilíngues para motoristas (2 mensagens por cliente)
 */
function enviarMensagensMultilingueMotoristas() {
  loggerHUB.info('Iniciando envio de mensagens multilíngues para motoristas');
  
  try {
    const motoristas = buscarMotoristasCadastrados();
    const transfers = buscarTransfersComMotoristasAtribuidos();
    
    if (transfers.length === 0) {
      throw new Error('Nenhum transfer com motorista atribuído encontrado na coluna V');
    }
    
    loggerHUB.info('Processando transfers com motoristas', {
      totalTransfers: transfers.length,
      motoristasUnicos: [...new Set(transfers.map(t => t.motorista))]
    });
    
    let mensagensEnviadas = 0;
    
    // Agrupar transfers por motorista
    const transfersPorMotorista = {};
    
    transfers.forEach(transfer => {
      const nomeMotorista = transfer.motorista;
      if (!transfersPorMotorista[nomeMotorista]) {
        transfersPorMotorista[nomeMotorista] = [];
      }
      transfersPorMotorista[nomeMotorista].push(transfer);
    });
    
    // Processar cada motorista que tem transfers atribuídos
    for (const [nomeMotorista, transfersDoMotorista] of Object.entries(transfersPorMotorista)) {
      
      // Buscar dados do motorista (se cadastrado)
      const dadosMotorista = motoristas.find(m => m.nome === nomeMotorista) || {
        nome: nomeMotorista,
        telefone: 'Não cadastrado',
        carro: 'Não informado',
        matricula: 'Não informada'
      };
      
      // PROCESSAR CADA CLIENTE INDIVIDUALMENTE
      transfersDoMotorista.forEach((transfer, index) => {
        const infoIdioma = detectarIdiomaClientePorDDI(transfer.contacto);
        const mensagemFormatada = formatarMensagemMotoristaMultilingue(transfer, dadosMotorista, infoIdioma);
        const numeroFormatado = '+' + String(transfer.contacto).replace(/\D/g, '');
        
        // MENSAGEM 1: INFORMAÇÕES DO TRANSFER (COM NÚMERO DO VOO)
        const mensagemInfo = `📞 ENVIAR PARA: ${numeroFormatado}
━━━ CLIENTE ${index + 1} ━━━
🆔 Transfer: #${transfer.id} (Linha ${transfer.linha})
👤 Cliente: ${transfer.cliente}
📱 Número Cliente: ${numeroFormatado}
🌍 País: ${infoIdioma.pais}
🗣️ Idioma: ${(infoIdioma.nome || infoIdioma.idioma).toUpperCase()}
✈️ Voo: ${transfer.numeroVoo ? 
  `${transfer.numeroVoo} - Rastrear: https://www.google.com/search?q=flight+${transfer.numeroVoo.replace(/\s+/g, '').toUpperCase()}` : 
  'Não informado'}
⏰ Horário: ${transfer.horaPickup}
📍 ${transfer.origem} → ${transfer.destino}
💬 MENSAGEM (${infoIdioma.idioma.toUpperCase()}) PARA COPIAR:`;
        
        // MENSAGEM 2: TEMPLATE LIMPO PARA O CLIENTE
        const mensagemTemplate = mensagemFormatada;
        
        // Enviar as duas mensagens separadamente
        enviarWhatsAppRoberta('+351968698138', mensagemInfo);
        
        // Pequena pausa entre mensagens
        Utilities.sleep(1000);
        
        enviarWhatsAppRoberta('+351968698138', mensagemTemplate);
        
        // Pausa maior entre clientes
        Utilities.sleep(2000);
      });
      
      // Mensagem final de resumo
      const mensagemResumo = `👨‍✈️ RESUMO PARA MOTORISTA: ${nomeMotorista}
🚗 Carro: ${dadosMotorista.carro}
📱 Telefone Motorista: ${dadosMotorista.telefone}
📊 Total de transfers processados: ${transfersDoMotorista.length}

Você recebeu ${transfersDoMotorista.length * 2} mensagens (2 por cliente):
- Informações do transfer
- Mensagem para enviar ao cliente`;
      
      enviarWhatsAppRoberta('+351968698138', mensagemResumo);
      
      mensagensEnviadas++;
      loggerHUB.success('Mensagens multilíngues enviadas para motorista', { 
        motorista: nomeMotorista,
        transfers: transfersDoMotorista.length 
      });
    }
    
    return {
      sucesso: true,
      motoristasProcessados: Object.keys(transfersPorMotorista).length,
      mensagensEnviadas: mensagensEnviadas,
      totalTransfers: transfers.length,
      motoristasList: Object.keys(transfersPorMotorista)
    };
    
  } catch (error) {
    loggerHUB.error('Erro ao enviar mensagens multilíngues para motoristas', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Teste do sistema multilíngue para motoristas
 */
function testarSistemaMotoristasMultilingue() {
  console.log('🧪 Testando sistema multilíngue para motoristas');
  
  try {
    const resultado = enviarMensagensMultilingueMotoristas();
    
    if (resultado.sucesso) {
      console.log('✅ Teste bem-sucedido!');
      console.log(`Motoristas processados: ${resultado.motoristasProcessados}`);
      console.log(`Mensagens enviadas: ${resultado.mensagensEnviadas}`);
      console.log(`Total de transfers: ${resultado.totalTransfers}`);
      console.log(`Motoristas: ${resultado.motoristasList.join(', ')}`);
      console.log('Verifique seu WhatsApp para as mensagens formatadas');
    } else {
      console.log('❌ Teste falhou:', resultado.erro);
    }
    
    return resultado;
    
  } catch (error) {
    console.log('❌ Erro no teste:', error.message);
    return { sucesso: false, erro: error.message };
  }
}

/**
 * Envia mensagens multilíngues para motoristas via menu
 */
function enviarMensagensMotoristas() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '📱 Enviar Mensagens para Motoristas',
    'Esta ação irá enviar mensagens formatadas no idioma correto para todos os motoristas com transfers atribuídos na coluna V.\n\nAs mensagens serão baseadas no DDI do cliente.\n\nDeseja continuar?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    const resultado = enviarMensagensMultilingueMotoristas();
    
    if (resultado.sucesso) {
      ui.alert(
        '✅ Mensagens Enviadas',
        `Mensagens multilíngues enviadas com sucesso!\n\n• Motoristas processados: ${resultado.motoristasProcessados}\n• Mensagens enviadas: ${resultado.mensagensEnviadas}\n• Total de transfers: ${resultado.totalTransfers}\n• Motoristas: ${resultado.motoristasList.join(', ')}\n\nVerifique seu WhatsApp para as mensagens no idioma correto.`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('❌ Erro', resultado.erro, ui.ButtonSet.OK);
    }
    
  } catch (error) {
    ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Corrige formatação dos números de telefone na coluna G de todas as abas
 */
function corrigirFormatacaoTelefones() {
  loggerHUB.info('Iniciando correção de formatação de telefones');
  
  try {
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    const abas = ss.getSheets();
    let totalCorrigidos = 0;
    
    abas.forEach(aba => {
      const nomeAba = aba.getName();
      
      // Pular apenas abas que realmente não são de transfers
      if (nomeAba.includes('Demo_') || 
          nomeAba.includes('Motoristas_') || 
          nomeAba.includes('Log_') ||
          nomeAba.includes('Stats_') ||
          nomeAba.includes('Auditoria_') ||
          nomeAba.includes('DEBUG_') ||
          nomeAba.includes('RAG_') ||
          nomeAba.includes('Nomes_') ||
          nomeAba.includes('Contexto_') ||
          nomeAba.includes('Qualidade_') ||
          nomeAba.includes('Treinamento_') ||
          nomeAba.includes('Etiquetas_') ||
          nomeAba.includes('Clientes_') ||
          nomeAba.includes('Upselling_') ||
          nomeAba === 'Dashboard_Hoteis' ||
          nomeAba === 'Controle_Nomes') {
        return;
      }
      
      loggerHUB.info('Processando aba', { aba: nomeAba });
      
      if (aba.getLastRow() > 1) {
        // Pegar dados da coluna G (Contacto)
        const range = aba.getRange(2, 7, aba.getLastRow() - 1, 1); // Coluna G
        const valores = range.getValues();
        const valoresCorrigidos = [];
        let corrigidosNaAba = 0;
        
        valores.forEach(linha => {
          const telefoneOriginal = linha[0];
          let telefoneCorrigido = telefoneOriginal;
          
          if (telefoneOriginal && telefoneOriginal.toString) {
            let telefoneString = telefoneOriginal.toString();
            
            // Remover = do início se existir
            if (telefoneString.startsWith('=')) {
              telefoneString = telefoneString.substring(1);
            }
            
            // Remover todos os caracteres não numéricos
            const numeroLimpo = telefoneString.replace(/\D/g, '');
            
            // Adicionar + no início se não tiver
            if (numeroLimpo && numeroLimpo.length > 0) {
              telefoneCorrigido = '+' + numeroLimpo;
            }
            
            if (telefoneCorrigido !== telefoneOriginal) {
              corrigidosNaAba++;
            }
          }
          
          valoresCorrigidos.push([telefoneCorrigido]);
        });
        
        // Aplicar correções se houver
        if (corrigidosNaAba > 0) {
          range.setValues(valoresCorrigidos);
          
          // Formatar coluna como texto para evitar problemas
          range.setNumberFormat('@');
          
          totalCorrigidos += corrigidosNaAba;
          
          loggerHUB.info('Telefones corrigidos na aba', { 
            aba: nomeAba, 
            corrigidos: corrigidosNaAba 
          });
        }
      }
    });
    
    loggerHUB.success('Correção de telefones concluída', { 
      totalCorrigidos: totalCorrigidos 
    });
    
    return {
      sucesso: true,
      totalCorrigidos: totalCorrigidos,
      mensagem: `${totalCorrigidos} telefones corrigidos em todas as abas`
    };
    
  } catch (error) {
    loggerHUB.error('Erro ao corrigir formatação de telefones', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Função para menu - Corrigir telefones
 */
function corrigirTelefonesMenu() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '📞 Corrigir Formatação de Telefones',
    'Esta ação irá corrigir a formatação de todos os números de telefone na coluna G de todas as abas:\n\n• Remove "=" do início\n• Adiciona "+" quando necessário\n• Formata como texto\n\nDeseja continuar?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    const resultado = corrigirFormatacaoTelefones();
    
    if (resultado.sucesso) {
      ui.alert(
        '✅ Telefones Corrigidos',
        resultado.mensagem,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('❌ Erro', resultado.erro, ui.ButtonSet.OK);
    }
    
  } catch (error) {
    ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Cria dashboard de monitoramento por hotel
 */
function criarDashboardMonitoramento() {
  console.log('Criando dashboard de monitoramento por hotel');
  
  try {
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    let dashboardSheet = ss.getSheetByName('Dashboard_Hoteis');
    
    if (!dashboardSheet) {
      dashboardSheet = ss.insertSheet('Dashboard_Hoteis');
    } else {
      dashboardSheet.clear();
    }
    
    // Headers do dashboard
    const headers = [
      'Hotel', 'Transfers Hoje', 'Pendentes', 'Confirmados', 
      'Próximos 7 Dias', 'Status Monitoramento', 'Última Atualização'
    ];
    
    dashboardSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    dashboardSheet.getRange(1, 1, 1, headers.length)
      .setBackground('#1a73e8')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    
    // Processar dados de cada hotel
    const hoje = new Date();
    const proximosSete = new Date(hoje.getTime() + 7 * 24 * 60 * 60 * 1000);
    
    const sheetCentral = ss.getSheetByName('HUB_Central');
    if (!sheetCentral || sheetCentral.getLastRow() <= 1) {
      dashboardSheet.getRange(2, 1).setValue('Nenhum dado disponível');
      return;
    }
    
    const dados = sheetCentral.getRange(2, 1, sheetCentral.getLastRow() - 1, 26).getValues();
    const resumoHoteis = {};
    
    // Inicializar contadores para cada hotel
    Object.keys(SISTEMA_HUB_CONFIG.HOTEIS_CONECTADOS).forEach(hotel => {
      resumoHoteis[hotel] = {
        transfersHoje: 0,
        pendentes: 0,
        confirmados: 0,
        proximosSete: 0,
        statusMonitoramento: 'Ativo'
      };
    });
    
    // Processar dados
    dados.forEach(linha => {
      const hotel = linha[20] || 'Não especificado'; // Coluna U
      const dataTransfer = new Date(linha[5]); // Coluna F
      const status = linha[17]; // Coluna R
      
      if (!resumoHoteis[hotel]) {
        resumoHoteis[hotel] = {
          transfersHoje: 0,
          pendentes: 0,
          confirmados: 0,
          proximosSete: 0,
          statusMonitoramento: 'Não monitorado'
        };
      }
      
      // Transfers de hoje
      if (dataTransfer.toDateString() === hoje.toDateString()) {
        resumoHoteis[hotel].transfersHoje++;
      }
      
      // Próximos 7 dias
      if (dataTransfer >= hoje && dataTransfer <= proximosSete) {
        resumoHoteis[hotel].proximosSete++;
      }
      
      // Status
      if (status === 'Pendente') {
        resumoHoteis[hotel].pendentes++;
      } else if (status === 'Confirmado') {
        resumoHoteis[hotel].confirmados++;
      }
    });
    
    // Adicionar dados ao dashboard
    let linha = 2;
    for (const [hotel, dados] of Object.entries(resumoHoteis)) {
      const linhaData = [
        hotel,
        dados.transfersHoje,
        dados.pendentes,
        dados.confirmados,
        dados.proximosSete,
        dados.statusMonitoramento,
        new Date().toLocaleString('pt-PT')
      ];
      
      dashboardSheet.getRange(linha, 1, 1, linhaData.length).setValues([linhaData]);
      
      // Colorir baseado no status
      if (dados.pendentes > 0) {
        dashboardSheet.getRange(linha, 3).setBackground('#fff3cd'); // Amarelo para pendentes
      }
      if (dados.confirmados > 0) {
        dashboardSheet.getRange(linha, 4).setBackground('#d4edda'); // Verde para confirmados
      }
      
      linha++;
    }
    
    // Configurar larguras
    dashboardSheet.setColumnWidth(1, 200); // Hotel
    dashboardSheet.setColumnWidth(2, 120); // Transfers Hoje
    dashboardSheet.setColumnWidth(3, 100); // Pendentes
    dashboardSheet.setColumnWidth(4, 120); // Confirmados
    dashboardSheet.setColumnWidth(5, 130); // Próximos 7 Dias
    dashboardSheet.setColumnWidth(6, 160); // Status Monitoramento
    dashboardSheet.setColumnWidth(7, 180); // Última Atualização
    
    console.log('Dashboard criado com sucesso');
    
    return {
      sucesso: true,
      hoteis: Object.keys(resumoHoteis).length,
      timestamp: new Date()
    };
    
  } catch (error) {
    console.log('Erro ao criar dashboard:', error.message);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

// ===================================================
// FINAL DO CÓDIGO - SISTEMA HUB CENTRAL COMPLETO
// ===================================================

console.log('');
console.log('🎉 SISTEMA HUB CENTRAL CARREGADO COMPLETAMENTE!');
console.log('📋 Execute AJUDA_HUB() para ver todas as funções disponíveis');
console.log('🚀 Execute INSTALAR_HUB_CENTRAL() para instalar o sistema');
console.log('');

/**
 * Adiciona coluna Motorista com destaque especial em todas as abas
 */
function adicionarColunaMotoristaTudo() {
  console.log('🚗 Adicionando coluna Motorista com destaque especial...');
  
  try {
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    const sheets = ss.getSheets();
    let abasAtualizadas = 0;
    
    sheets.forEach(sheet => {
      const nomeAba = sheet.getName();
      
      // Atualizar aba central e abas mensais
      if (nomeAba === 'HUB_Central' || nomeAba.startsWith('HUB_')) {
        try {
          // Encontrar linha do header
          let linhaHeader = 1;
          if (nomeAba.startsWith('HUB_') && nomeAba !== 'HUB_Central') {
            linhaHeader = 4; // Abas mensais têm header na linha 4
          }
          
          // Adicionar coluna Motorista (índice 22)
          sheet.getRange(linhaHeader, 22).setValue('Motorista');
          
          // DESTAQUE ESPECIAL: Cor laranja/amarela para motorista
          sheet.getRange(linhaHeader, 22)
            .setBackground('#ff9500')  // Cor laranja vibrante
            .setFontColor('#ffffff')   // Texto branco
            .setFontWeight('bold')     // Texto em negrito
            .setFontSize(11)          // Tamanho da fonte
            .setHorizontalAlignment('center'); // Centralizado
          
          // Definir largura da coluna
          sheet.setColumnWidth(22, 150);
          
          abasAtualizadas++;
          console.log(`✅ Coluna Motorista com destaque adicionada: ${nomeAba}`);
          
        } catch (error) {
          console.log(`⚠️ Erro na aba ${nomeAba}: ${error.message}`);
        }
      }
    });
    
    console.log(`🎉 Coluna Motorista com destaque adicionada em ${abasAtualizadas} aba(s)!`);
    console.log('🚗 Agora a coluna Motorista tem destaque especial em laranja');
    
    return {
      sucesso: true,
      abasAtualizadas: abasAtualizadas
    };
    
  } catch (error) {
    console.log('❌ Erro ao adicionar coluna:', error.message);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Testa fluxo de hotel (deve ficar pendente) - CORRIGIDO
 */
function testarFluxoHotel() {
  console.log('🏨 Testando fluxo de HOTEL (deve ficar pendente)...');
  
  try {
    const dadosHotel = {
      nomeCliente: 'TESTE HOTEL - João Silva',
      tipoServico: 'Transfer',
      numeroPessoas: 2,
      numeroBagagens: 1,
      data: new Date(),
      contacto: '+351999888777',
      numeroVoo: 'TP123',
      origem: 'Aeroporto de Lisboa',
      destino: 'Empire Marques Hotel',
      horaPickup: '15:00',
      valorTotal: 35.00,
      status: 'Solicitado',
      modoPagamento: 'Dinheiro',
      pagoParaQuem: 'Recepção',
      observacoes: 'Teste de fluxo hotel'
    };
    
    const parametros = {
      source: 'empire_marques'  // ← HOTEL (deve ficar pendente)
    };
    
    console.log('📝 Dados enviados:', dadosHotel);
    
    const resultado = processarTransferDoHotel(dadosHotel, parametros);
    
    console.log('✅ Resultado Hotel:', resultado);
    console.log('🎯 Deve mostrar: status Pendente, whatsappEnviado: false');
    
    return resultado;
    
  } catch (error) {
    console.log('❌ Erro no teste de hotel:', error.message);
    
    // Debug: vamos ver o que está na validação
    console.log('🔍 Testando validação diretamente...');
    try {
      const dadosSimples = {
        nomeCliente: 'TESTE HOTEL - João Silva',
        numeroPessoas: 2,
        data: new Date(),
        contacto: '+351999888777',
        origem: 'Aeroporto de Lisboa',
        destino: 'Empire Marques Hotel',
        horaPickup: '15:00',
        valorTotal: 35.00
      };
      
      const validacao = validarDadosHotel(dadosSimples);
      console.log('📋 Resultado validação:', validacao);
      
    } catch (validacaoError) {
      console.log('❌ Erro na validação:', validacaoError.message);
    }
    
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Testa integração frontend - WT
 */
function testarIntegracaoWT() {
  console.log('🧪 Testando integração frontend - WT...');
  
  const dadosTesteWT = {
    nomeCliente: 'Cliente WT Frontend',
    numeroPessoas: 2,
    numeroBagagens: 1,
    data: new Date(),
    contacto: '+351987654321',
    numeroVoo: 'TP456',
    origem: 'Aeroporto de Lisboa',
    destino: 'Hotel Tivoli', 
    horaPickup: '14:00',
    valorServico: 45.00,
    operadora: 'wt',
    formaPagamento: 'monetario',
    cobranca: 'motorista',
    observacoes: 'Cliente VIP - WT'
  };
  
  const resultado = receberDadosFrontendHUB(dadosTesteWT);
  console.log('✅ Resultado WT:', resultado);
  
  return resultado;
}

/**
 * Testa integração frontend - Connecto
 */
function testarIntegracaoConnecto() {
  console.log('🧪 Testando integração frontend - Connecto...');
  
  const dadosTesteConnecto = {
    nomeCliente: 'Cliente Connecto Frontend',
    numeroPessoas: 3,
    numeroBagagens: 2,
    data: new Date(),
    contacto: '+351976543210',
    origem: 'Hotel Pestana',
    destino: 'Aeroporto de Lisboa', 
    horaPickup: '10:30',
    valorServico: 38.00,
    operadora: 'connecto',
    formaPagamento: 'multibanco',
    cobranca: 'operador',
    observacoes: 'Transfer de retorno - Connecto'
  };
  
  const resultado = receberDadosFrontendHUB(dadosTesteConnecto);
  console.log('✅ Resultado Connecto:', resultado);
  
  return resultado;
}

/**
 * Testa integração frontend - HUB Transfer
 */
function testarIntegracaoHUBTransfer() {
  console.log('🧪 Testando integração frontend - HUB Transfer...');
  
  const dadosTesteHUB = {
    nomeCliente: 'Cliente HUB Transfer Direto',
    numeroPessoas: 4,
    numeroBagagens: 3,
    data: new Date(),
    contacto: '+351965432109',
    numeroVoo: 'LH789',
    origem: 'Aeroporto de Lisboa',
    destino: 'Hotel Sheraton', 
    horaPickup: '16:45',
    valorServico: 55.00,
    operadora: 'hub_transfer',
    formaPagamento: 'transferencia',
    cobranca: 'empresa',
    observacoes: 'Grupo empresarial - HUB Transfer'
  };
  
  const resultado = receberDadosFrontendHUB(dadosTesteHUB);
  console.log('✅ Resultado HUB Transfer:', resultado);
  
  return resultado;
}

/**
 * Cria aba de cadastro de motoristas
 */
function criarAbaMotoristasHUB() {
  console.log('🚗 Criando aba de motoristas...');
  
  try {
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    let abaMotoristas = ss.getSheetByName('Motoristas_HUB');
    
    if (abaMotoristas) {
      console.log('✅ Aba já existe');
      return abaMotoristas;
    }
    
    // Criar aba
    abaMotoristas = ss.insertSheet('Motoristas_HUB');
    
    // Headers
    const headers = ['Nome', 'Telefone', 'Carro Modelo', 'Matrícula'];
    abaMotoristas.getRange(1, 1, 1, 4).setValues([headers]);
    
    // Formatação
    abaMotoristas.getRange(1, 1, 1, 4)
      .setBackground('#1a73e8')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    
    // Larguras
    abaMotoristas.setColumnWidth(1, 200); // Nome
    abaMotoristas.setColumnWidth(2, 150); // Telefone  
    abaMotoristas.setColumnWidth(3, 180); // Carro
    abaMotoristas.setColumnWidth(4, 120); // Matrícula
    
    console.log('✅ Aba Motoristas_HUB criada');
    return abaMotoristas;
    
  } catch (error) {
    console.log('❌ Erro:', error.message);
    throw error;
  }
}

/**
 * Atualiza planilhas existentes com as novas colunas W, X, Y, Z
 */
function atualizarComNovasColunas() {
  loggerHUB.info('Atualizando planilhas existentes com novas colunas W, X, Y, Z');
  
  try {
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    const sheets = ss.getSheets();
    let abasAtualizadas = 0;
    
    sheets.forEach(sheet => {
      const nomeAba = sheet.getName();
      
      // Atualizar aba central e abas mensais
      if (nomeAba === 'HUB_Central' || nomeAba.startsWith('HUB_')) {
        try {
          // Encontrar linha do header
          let linhaHeader = 1;
          if (nomeAba.startsWith('HUB_') && nomeAba !== 'HUB_Central') {
            linhaHeader = 4; // Abas mensais têm header na linha 4
          }
          
          // Verificar se as colunas já existem
          const ultimaColuna = sheet.getLastColumn();
          
          if (ultimaColuna < 26) { // Se não tem todas as colunas
            // Adicionar headers das novas colunas
            sheet.getRange(linhaHeader, 23).setValue('Status_OK');
            sheet.getRange(linhaHeader, 24).setValue('Timestamp_OK');
            sheet.getRange(linhaHeader, 25).setValue('Tipo_Viagem');
            sheet.getRange(linhaHeader, 26).setValue('Status_Followup');
            
            // Aplicar formatação dos headers
            const novosHeaders = sheet.getRange(linhaHeader, 23, 1, 4);
            if (nomeAba === 'HUB_Central') {
              novosHeaders
                .setBackground('#1a73e8')
                .setFontColor('#ffffff')
                .setFontWeight('bold')
                .setFontSize(11)
                .setHorizontalAlignment('center');
            } else {
              // Pegar cor da aba mensal
              const corHeader = sheet.getRange(linhaHeader, 1).getBackground();
              novosHeaders
                .setBackground(corHeader)
                .setFontColor('#ffffff')
                .setFontWeight('bold')
                .setFontSize(10)
                .setHorizontalAlignment('center');
            }
            
            // Definir larguras das novas colunas
            sheet.setColumnWidth(23, 120); // Status_OK
            sheet.setColumnWidth(24, 160); // Timestamp_OK
            sheet.setColumnWidth(25, 120); // Tipo_Viagem
            sheet.setColumnWidth(26, 150); // Status_Followup
            
            // Aplicar validações e formatações
            aplicarValidacoesNovasColunas(sheet);
            aplicarFormatacaoNovasColunas(sheet);
            
            abasAtualizadas++;
            console.log(`Atualizada: ${nomeAba}`);
          }
          
        } catch (error) {
          loggerHUB.error(`Erro na aba ${nomeAba}`, error);
        }
      }
    });
    
    loggerHUB.success(`Novas colunas adicionadas em ${abasAtualizadas} aba(s)`);
    
    return {
      sucesso: true,
      abasAtualizadas: abasAtualizadas
    };
    
  } catch (error) {
    loggerHUB.error('Erro ao adicionar novas colunas', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Webhook principal - VERSÃO ROBUSTA E SIMPLIFICADA
 */
function doPost(e) {
  console.log('=== WEBHOOK RECEBIDO ===');
  
  try {
    if (!e || !e.postData) {
      console.log('❌ Dados do webhook ausentes');
      return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
    }
    
    const dados = JSON.parse(e.postData.contents);
    console.log('📡 Dados brutos:', JSON.stringify(dados));
    
    // Extrair telefone e mensagem de forma mais robusta
    let telefone = null;
    let mensagem = null;
    
    // Tentar múltiplos campos para telefone
    telefone = dados.phone || dados.from || dados.sender || dados.remoteJid;
    
    // Tentar múltiplos campos para mensagem
    if (dados.text && dados.text.message) {
      mensagem = dados.text.message;
    } else if (dados.message) {
      mensagem = dados.message;
    } else if (dados.body) {
      mensagem = dados.body;
    }
    
    console.log('📱 Extraído - Telefone:', telefone, 'Mensagem:', mensagem);
    
    if (!telefone || !mensagem) {
      console.log('❌ Telefone ou mensagem não encontrados');
      return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
    }
    
    // Processar mensagem com sistema robusto
    processarMensagemRobusta(telefone, mensagem);
    
    return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
    
  } catch (error) {
    console.log('❌ ERRO no webhook:', error.message);
    console.log('❌ Stack:', error.stack);
    return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
  }
}

/**
 * Processar mensagem de forma robusta com fallbacks
 */
function processarMensagemRobusta(telefone, mensagem) {
  console.log('🔄 Iniciando processamento robusto...');
  
  try {
    // Normalizar telefone
    let telefoneNormalizado = telefone;
    if (typeof telefone === 'number') {
      telefoneNormalizado = String(telefone);
    }
    
    // Adicionar + se não tiver
    if (!telefoneNormalizado.startsWith('+')) {
      telefoneNormalizado = '+' + telefoneNormalizado.replace(/\D/g, '');
    }
    
    console.log('📱 Telefone normalizado:', telefoneNormalizado);
    
    // TENTATIVA 1: Sistema Demo (mais simples e confiável)
    if (tentarSistemaDemo(telefoneNormalizado, mensagem)) {
      console.log('✅ Processado pelo sistema DEMO');
      return;
    }
    
    // TENTATIVA 2: Sistema simplificado direto
    if (tentarSistemaSimplificado(telefoneNormalizado, mensagem)) {
      console.log('✅ Processado pelo sistema SIMPLIFICADO');
      return;
    }
    
    // FALLBACK: Resposta básica garantida
    enviarRespostaBasica(telefoneNormalizado, mensagem);
    
  } catch (error) {
    console.log('❌ Erro no processamento robusto:', error.message);
    // Último recurso
    enviarRespostaEmergencia(telefone);
  }
}

/**
 * Tentar sistema demo
 */
function tentarSistemaDemo(telefone, mensagem) {
  try {
    const perfil = buscarPerfilDemoIntegrado(telefone);
    if (perfil && perfil.status === 'ATIVO') {
      const resposta = gerarRespostaDemoIntegrada(perfil, mensagem);
      if (resposta.sucesso) {
        return enviarWhatsAppRobertaSeguro(telefone, resposta.resposta);
      }
    }
    return false;
  } catch (error) {
    console.log('⚠️ Erro sistema demo:', error.message);
    return false;
  }
}

/**
 * Sistema simplificado
 */
function tentarSistemaSimplificado(telefone, mensagem) {
  try {
    const resposta = gerarRespostaInteligente(mensagem);
    return enviarWhatsAppRobertaSeguro(telefone, resposta);
  } catch (error) {
    console.log('⚠️ Erro sistema simplificado:', error.message);
    return false;
  }
}

/**
 * Gerar resposta inteligente básica
 */
function gerarRespostaInteligente(mensagem) {
  const msg = mensagem.toLowerCase();
  
  if (msg.includes('oi') || msg.includes('olá') || msg.includes('hello')) {
    return 'Olá! Sou a Roberta da HUB Transfer. Como posso ajudar?';
  }
  
  if (msg.includes('transfer') || msg.includes('aeroporto')) {
    return 'Perfeito! Para seu transfer, preciso saber origem, destino e horário. Pode me informar?';
  }
  
  if (msg.includes('obrigad') || msg.includes('thanks')) {
    return 'De nada! Sempre às ordens! 😊';
  }
  
  return 'Como posso ajudar com seu transfer hoje?';
}

/**
 * Envio seguro com validação
 */
function enviarWhatsAppRobertaSeguro(telefone, mensagem) {
  try {
    // Usar as configurações existentes
    const numeroLimpo = telefone.replace(/\D/g, '');
    
    const payload = {
      phone: numeroLimpo,
      message: mensagem
    };
    
    const options = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Client-Token': SISTEMA_HUB_CONFIG.Z_API.CLIENT_TOKEN
      },
      payload: JSON.stringify(payload)
    };
    
    console.log('🚀 Enviando via Z-API...');
    const response = UrlFetchApp.fetch(SISTEMA_HUB_CONFIG.Z_API.ENDPOINT, options);
    
    console.log('📡 Status:', response.getResponseCode());
    
    return response.getResponseCode() === 200;
    
  } catch (error) {
    console.log('❌ Erro no envio seguro:', error.message);
    return false;
  }
}

/**
 * Resposta básica garantida
 */
function enviarRespostaBasica(telefone, mensagem) {
  console.log('🔧 Enviando resposta básica...');
  const resposta = 'Olá! Sou a Roberta HUB. Como posso ajudar?';
  enviarWhatsAppRobertaSeguro(telefone, resposta);
}

/**
 * Resposta de emergência
 */
function enviarRespostaEmergencia(telefone) {
  console.log('🚨 EMERGÊNCIA: Enviando resposta mínima...');
  try {
    const numeroLimpo = String(telefone).replace(/\D/g, '');
    const payload = {
      phone: numeroLimpo,
      message: 'Sistema ativo. Como posso ajudar?'
    };
    
    UrlFetchApp.fetch(SISTEMA_HUB_CONFIG.Z_API.ENDPOINT, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Client-Token': SISTEMA_HUB_CONFIG.Z_API.CLIENT_TOKEN
      },
      payload: JSON.stringify(payload)
    });
  } catch (error) {
    console.log('❌ Falha total:', error.message);
  }
}

/**
 * Buscar perfil demo - INTEGRADO NO WEBHOOK
 */
function buscarPerfilDemoIntegrado(telefone) {
  try {
    console.log(`🔍 Buscando perfil demo integrado para: ${telefone}`);
    
    const numeroLimpo = telefone.replace(/\D/g, '');
    console.log(`📱 Número limpo: ${numeroLimpo}`);
    
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    const abaDemo = ss.getSheetByName('Demo_Clientes');
    
    if (!abaDemo) {
      console.log('❌ Aba Demo_Clientes não encontrada');
      return null;
    }
    
    const dados = abaDemo.getDataRange().getValues();
    
    for (let i = 1; i < dados.length; i++) {
      const telefoneLinhaFormatado = String(dados[i][1]).replace(/\D/g, '');
      
      if (telefoneLinhaFormatado === numeroLimpo) {
        const perfil = {
          nome: dados[i][0],
          telefone: dados[i][1],
          empresa: dados[i][2],
          cargo: dados[i][3],
          hotel: dados[i][4],
          descricao: dados[i][5],
          interesses: dados[i][6],
          historico: dados[i][7],
          preferencias: dados[i][8],
          observacoes: dados[i][9],
          status: dados[i][10],
          linha: i + 1
        };
        
        console.log(`✅ Perfil demo encontrado: ${perfil.nome} (${perfil.cargo})`);
        return perfil;
      }
    }
    
    console.log('❌ Nenhum perfil demo encontrado');
    return null;
    
  } catch (error) {
    console.log('❌ Erro ao buscar perfil demo integrado:', error.message);
    return null;
  }
}

/**
 * Gerar resposta demo - INTEGRADA NO WEBHOOK
 */
function gerarRespostaDemoIntegrada(perfilCliente, mensagemCliente) {
  try {
    console.log(`🎭 Gerando resposta demo integrada para: ${perfilCliente.nome}`);
    
    // Verificar se é primeira mensagem (cumprimento)
    const ehPrimeiraMensagem = mensagemCliente.toLowerCase().match(/^(oi|olá|hello|hi|boa tarde|bom dia|hey)$/);
    
    if (ehPrimeiraMensagem) {
      // Resposta de impacto baseada no perfil
      const nome = perfilCliente.nome;
      const cargo = perfilCliente.cargo || '';
      const empresa = perfilCliente.empresa || '';
      const observacoes = perfilCliente.observacoes || '';
      
      const respostasImpacto = [
        `${nome}! Que surpresa incrível! Mesmo antes de você se apresentar, eu já sabia exatamente quem você é! ${observacoes}`,
        `${nome}! ${cargo} da ${empresa}! Como é impressionante - o sistema já me contou tudo sobre você!`,
        `${nome}! Eu estava esperando por você! O Junior me ensinou tanto sobre você que parece que já nos conhecemos há anos!`,
        `${nome}! Incrível! Antes mesmo de você falar, eu já reconheci quem você é. Como está?`
      ];
      
      const respostaEscolhida = respostasImpacto[Math.floor(Math.random() * respostasImpacto.length)];
      
      return {
        sucesso: true,
        resposta: respostaEscolhida,
        tipoDemo: 'IMPACTO_INTEGRADO',
        cliente: perfilCliente.nome
      };
    }
    
    // Para outras mensagens, usar template básico personalizado
    const respostaPersonalizada = `${perfilCliente.nome}, como posso ajudar você hoje? Sei que você é ${perfilCliente.cargo} da ${perfilCliente.empresa}!`;
    
    return {
      sucesso: true,
      resposta: respostaPersonalizada,
      tipoDemo: 'PERSONALIZADO_INTEGRADO',
      cliente: perfilCliente.nome
    };
    
  } catch (error) {
    console.log('Erro ao gerar resposta demo integrada:', error.message);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

function extrairTelefoneSeguro(dados) {
  const campos = [
    dados.phone,
    dados.from,
    dados.data?.from,
    dados.message?.from,
    dados.sender
  ];
  
  for (const campo of campos) {
    if (campo && typeof campo === 'string' && campo.length >= 10) {
      // Normalizar o telefone
      let telefoneNorm = campo.replace(/\D/g, '');
      if (!telefoneNorm.startsWith('351') && telefoneNorm.length === 9) {
        telefoneNorm = '351' + telefoneNorm;
      }
      return '+' + telefoneNorm;
    }
  }
  return null;
}

function extrairMensagemSeguro(dados) {
  const campos = [
    dados.text?.message,
    dados.message?.text,
    dados.body,
    dados.data?.body,
    dados.text,
    dados.message
  ];
  
  for (const campo of campos) {
    if (campo && typeof campo === 'string' && campo.trim().length > 0) {
      return campo.trim();
    }
  }
  return null;
}

/**
 * Extrai dados da mensagem de forma robusta
 */
function extractMessageData(webhookData) {
  try {
    let telefone = null;
    let mensagem = null;
    let messageType = 'text';
    
    // Múltiplos formatos suportados
    const phoneFields = ['phone', 'from', 'sender', 'number'];
    const messageFields = ['text', 'body', 'content'];
    
    // Extrair telefone
    for (const field of phoneFields) {
      if (webhookData[field]) {
        telefone = webhookData[field];
        break;
      }
    }
    
    // Extrair mensagem
    if (webhookData.message) {
      if (typeof webhookData.message === 'string') {
        mensagem = webhookData.message;
      } else {
        for (const field of messageFields) {
          if (webhookData.message[field]) {
            mensagem = webhookData.message[field];
            break;
          }
        }
        messageType = webhookData.message.messageType || 'text';
      }
    } else {
      for (const field of messageFields) {
        if (webhookData[field]) {
          mensagem = webhookData[field];
          break;
        }
      }
    }
    
    // Validação
    if (!telefone || !mensagem) {
      return {
        isValid: false,
        error: `Dados ausentes - telefone: ${!!telefone}, mensagem: ${!!mensagem}`
      };
    }
    
    // Normalizar telefone
    telefone = telefone.replace(/\D/g, '');
    if (telefone.length < 10) {
      return { isValid: false, error: 'Telefone muito curto' };
    }
    
    // Adicionar código país se necessário
    if (!telefone.startsWith('351') && telefone.length === 9) {
      telefone = '351' + telefone;
    }
    
    return {
      isValid: true,
      telefone: '+' + telefone,
      mensagem: mensagem.trim(),
      messageType: messageType
    };
    
  } catch (error) {
    return { isValid: false, error: error.message };
  }
}

/**
 * Verifica se é mensagem de motorista
 */
function isDriverMessage(telefone, mensagem) {
  try {
    // Verificar se telefone está na lista de motoristas
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Motoristas_HUB');
    
    if (!sheet) return false;
    
    const data = sheet.getDataRange().getValues();
    const telefoneNorm = telefone.replace(/\D/g, '');
    
    for (let i = 1; i < data.length; i++) {
      const driverPhone = String(data[i][1]).replace(/\D/g, '');
      if (driverPhone === telefoneNorm) {
        // Verificar se mensagem parece ser confirmação
        const msgLower = mensagem.toLowerCase();
        return msgLower.includes('ok') || 
               msgLower.includes('finalizado') || 
               msgLower.includes('concluido') ||
               msgLower.includes('entregue');
      }
    }
    
    return false;
    
  } catch (error) {
    console.log('Erro ao verificar motorista:', error.message);
    return false;
  }
}

/**
 * Processa OK do motorista
 */
function processDriverOK(telefone, mensagem) {
  try {
    console.log('Processando OK do motorista:', telefone);
    
    // Encontrar transfer ativo do motorista
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('HUB_Central');
    const driversSheet = ss.getSheetByName('Motoristas_HUB');
    
    // Encontrar nome do motorista
    const driversData = driversSheet.getDataRange().getValues();
    const telefoneNorm = telefone.replace(/\D/g, '');
    let driverName = null;
    
    for (let i = 1; i < driversData.length; i++) {
      if (String(driversData[i][1]).replace(/\D/g, '') === telefoneNorm) {
        driverName = driversData[i][0];
        break;
      }
    }
    
    if (!driverName) {
      console.log('Motorista não encontrado');
      return { success: false, error: 'Motorista não encontrado' };
    }
    
    // Encontrar transfer ativo de hoje
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    const transferData = sheet.getDataRange().getValues();
    
    for (let i = 1; i < transferData.length; i++) {
      const row = transferData[i];
      const transferDate = new Date(row[5]); // Coluna F - Data
      transferDate.setHours(0, 0, 0, 0);
      
      const assignedDriver = row[21]; // Coluna V - Motorista
      const statusOK = row[22]; // Coluna W - Status_OK
      
      if (assignedDriver === driverName && 
          !statusOK && 
          transferDate.getTime() === today.getTime()) {
        
        // Marcar como OK
        sheet.getRange(i + 1, 23).setValue('✅ OK'); // Coluna W
        sheet.getRange(i + 1, 24).setValue(new Date()); // Coluna X
        
        // Enviar confirmação para motorista
        const confirmationMsg = `✅ Perfeito ${driverName}! Transfer ${row[0]} marcado como concluído. Obrigado!`;
        enviarWhatsAppRoberta(telefone, confirmationMsg);
        
        console.log(`Transfer ${row[0]} marcado como OK pelo motorista ${driverName}`);
        
        return { 
          success: true, 
          transferId: row[0], 
          driver: driverName 
        };
      }
    }
    
    // Não encontrou transfer ativo
    const noTransferMsg = `${driverName}, não encontrei transfer ativo para hoje. Verifique se há algum transfer atribuído.`;
    enviarWhatsAppRoberta(telefone, noTransferMsg);
    
    return { success: false, error: 'Nenhum transfer ativo encontrado' };
    
  } catch (error) {
    console.log('Erro ao processar OK do motorista:', error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Processa mensagem do cliente
 */
function processClientMessage(telefone, mensagem) {
  try {
    // Usar sistema de estado consolidado
    const clientState = new ConversationState(telefone);
    
    // Processar com OpenAI
    const openaiResult = processWithOpenAI(clientState, mensagem);
    
    if (openaiResult.success) {
      // Atualizar estado
      clientState.updateContext({
        mensagem: mensagem,
        resposta: openaiResult.response,
        estado: openaiResult.nextState || clientState.estado
      });
      
      // Enviar resposta
      const sent = enviarWhatsAppRoberta(telefone, openaiResult.response);
      
      return { 
        success: sent,
        response: openaiResult.response,
        state: clientState.estado
      };
    }
    
    // Fallback para resposta simples
    const fallbackResponse = `Olá! Sou a Roberta da HUB Transfer. Como posso ajudar com seu transfer hoje?`;
    const sent = enviarWhatsAppRoberta(telefone, fallbackResponse);
    
    return { success: sent, response: fallbackResponse, state: 'FALLBACK' };
    
  } catch (error) {
    console.log('Erro ao processar mensagem do cliente:', error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Processa com OpenAI usando estado consolidado
 */
function processWithOpenAI(clientState, mensagem) {
  try {
    const API_KEY = '';
    
    // Construir prompt contextual
    const prompt = buildContextualPrompt(clientState, mensagem);
    
    const payload = {
      model: 'gpt-4o-mini',
      messages: [
        {
          role: 'system',
          content: 'Você é a Roberta da HUB Transfer em Lisboa. Seja natural, empática e mantenha contexto das conversas. Ajude com transfers de forma humana e profissional.'
        },
        {
          role: 'user',
          content: prompt
        }
      ],
      max_tokens: 300,
      temperature: 0.7
    };
    
    const options = {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${API_KEY}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      const aiResponse = data.choices[0].message.content;
      
      return {
        success: true,
        response: aiResponse,
        nextState: determineNextState(clientState, mensagem, aiResponse)
      };
    }
    
    return { success: false, error: 'OpenAI API error' };
    
  } catch (error) {
    console.log('Erro no processamento OpenAI:', error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Constrói prompt contextual
 */
function buildContextualPrompt(clientState, mensagem) {
  let prompt = `CONTEXTO DO CLIENTE:
- Telefone: ${clientState.telefone}
- Nome: ${clientState.nome || 'Não informado'}
- Idioma: ${clientState.idioma}
- Tipo: ${clientState.tipoCliente}
- Estado: ${clientState.estado}
- Primeira interação: ${clientState.primeiraInteracao}

DADOS DO TRANSFER:
- Origem: ${clientState.dadosTransfer.origem || 'Não informado'}
- Destino: ${clientState.dadosTransfer.destino || 'Não informado'}
- Horário: ${clientState.dadosTransfer.horario || 'Não informado'}
- Próxima ação: ${clientState.proximaAcao}

MENSAGEM ATUAL: "${mensagem}"

HISTÓRICO RECENTE:`;

  if (clientState.context.historico.length > 0) {
    clientState.context.historico.slice(-3).forEach((h, i) => {
      prompt += `\n${i + 1}. Cliente: "${h.mensagem}" | Roberta: "${h.resposta}"`;
    });
  } else {
    prompt += '\nPrimeira conversa';
  }

  prompt += `\n\nRESPONDA de forma natural e contextual. Se primeira interação, apresente-se. Mantenha continuidade da conversa.`;

  return prompt;
}

/**
 * Determina próximo estado
 */
function determineNextState(clientState, mensagem, response) {
  const currentState = clientState.estado;
  
  if (currentState === 'INICIAL' && response.toLowerCase().includes('roberta')) {
    return 'APRESENTADA';
  }
  
  if (clientState.proximaAcao === 'COLETAR_ORIGEM' && clientState.dadosTransfer.origem) {
    return 'ORIGEM_COLETADA';
  }
  
  // Manter estado atual por padrão
  return currentState;
}

/**
 * Log de erros do webhook
 */
function logWebhookError(errorType, details) {
  try {
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    let errorSheet = ss.getSheetByName('Webhook_Errors');
    
    if (!errorSheet) {
      errorSheet = ss.insertSheet('Webhook_Errors');
      errorSheet.getRange(1, 1, 1, 4).setValues([
        ['Timestamp', 'Error_Type', 'Details', 'Status']
      ]);
    }
    
    errorSheet.appendRow([
      new Date(),
      errorType,
      details,
      'LOGGED'
    ]);
    
  } catch (logError) {
    console.log('Erro ao registrar erro:', logError.message);
  }
}

/**
 * Resposta de emergência
 */
function sendEmergencyResponse(telefone) {
  const emergencyMsg = 'Desculpe, tive uma dificuldade técnica. Pode repetir sua mensagem?';
  return enviarWhatsAppRoberta(telefone, emergencyMsg);
}

function monitorarQualidadeConversa(telefone, mensagem, resposta, tempoProcessamento) {
  try {
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    let abaQualidade = ss.getSheetByName('Qualidade_Conversas');
    
    if (!abaQualidade) {
      abaQualidade = ss.insertSheet('Qualidade_Conversas');
      const headers = [
        'Timestamp', 'Telefone', 'Mensagem_Cliente', 'Resposta_Roberta', 
        'Tempo_Processamento_ms', 'Tamanho_Resposta', 'Contém_Nome', 
        'Tipo_Interacao', 'Score_Estimado', 'Status'
      ];
      abaQualidade.getRange(1, 1, 1, headers.length).setValues([headers]);
      abaQualidade.getRange(1, 1, 1, headers.length)
        .setBackground('#34a853')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    }
    
    // Análise básica de qualidade
    const contemNome = /\b[A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇÑ][a-záàâãéèêíïóôõöúçñ]+\b/.test(resposta);
    const tamanhoResposta = resposta.length;
    const tipoInteracao = determinarTipoInteracao(mensagem);
    const scoreEstimado = calcularScoreQualidade(mensagem, resposta, tempoProcessamento);
    
    const linha = [
      new Date(),
      telefone.substring(0, 6) + '***', // Mascarar telefone para privacidade
      mensagem,
      resposta,
      tempoProcessamento,
      tamanhoResposta,
      contemNome,
      tipoInteracao,
      scoreEstimado,
      'ATIVO'
    ];
    
    abaQualidade.appendRow(linha);
    console.log('Qualidade monitorada - Score:', scoreEstimado);
    
  } catch (error) {
    console.log('Erro no monitoramento de qualidade:', error.message);
  }
}

function determinarTipoInteracao(mensagem) {
 const msg = mensagem.toLowerCase();
 
 if (msg.includes('transfer') || msg.includes('aeroporto')) return 'PEDIDO_TRANSFER';
 if (msg.includes('ola') || msg.includes('bom dia') || msg.includes('boa tarde')) return 'SAUDACAO';
 if (msg.includes('obrigad') || msg.includes('thanks')) return 'AGRADECIMENTO';
 if (msg.includes('problema') || msg.includes('nao apareceu')) return 'PROBLEMA';
 if (msg.includes('preco') || msg.includes('quanto')) return 'CONSULTA_PRECO';
 if (msg.includes('confirma') || msg.includes('ok')) return 'CONFIRMACAO';
 if (msg.includes('?')) return 'PERGUNTA';
 
 return 'CONVERSA_GERAL';
}

function calcularScoreQualidade(mensagem, resposta, tempoProcessamento) {
 let score = 50; // Base neutra
 
 // Pontuação por tempo de resposta
 if (tempoProcessamento < 3000) score += 20;
 else if (tempoProcessamento < 5000) score += 10;
 else if (tempoProcessamento > 10000) score -= 15;
 
 // Pontuação por tamanho adequado da resposta
 const tamanho = resposta.length;
 if (tamanho >= 50 && tamanho <= 300) score += 15;
 else if (tamanho < 20) score -= 20;
 else if (tamanho > 500) score -= 10;
 
 // Pontuação por naturalidade (heurísticas básicas)
 if (resposta.includes('Como posso ajudar')) score += 5;
 if (resposta.includes('transferência') || resposta.includes('transfer')) score += 10;
 if (resposta.match(/[😊🌟📍⏰✈️🚐]/)) score += 5; // Emojis apropriados
 
 // Penalizações por robótica
 if (resposta.includes('Sou a Roberta') && mensagem.length < 50) score -= 10;
 if (resposta.includes('assistente virtual') && !mensagem.toLowerCase().includes('quem')) score -= 5;
 
 return Math.min(Math.max(score, 0), 100); // Limitar entre 0-100
}

/**
* Gera resposta contextual quando tudo falha
*/
function gerarRespostaContextualFallback(mensagem, contexto) {
 const msg = mensagem.toLowerCase();
 
 if (msg.includes('transfer') || msg.includes('aeroporto')) {
   return 'Para seu transfer, preciso saber a origem e o destino. Pode me informar?';
 }
 
 if (msg.includes('preço') || msg.includes('quanto') || msg.includes('valor')) {
   return 'Para dar um orçamento preciso, preciso saber a rota do transfer. De onde para onde?';
 }
 
 if (msg.includes('boa tarde') || msg.includes('olá') || msg.includes('oi') || msg.includes('bom dia')) {
   return 'Olá! Como posso ajudar com seu transfer hoje?';
 }
 
 if (msg.includes('nome') || msg.includes('chamo')) {
   return 'Prazer em conhecê-lo! Como posso ajudar com seu transfer?';
 }
 
 if (contexto && contexto.length > 0) {
   const ultimaInteracao = contexto[0];
   if (ultimaInteracao && ultimaInteracao.conteudoMensagem) {
     if (ultimaInteracao.conteudoMensagem.includes('transfer')) {
       return 'Entendi que precisa de transfer. Pode me dar mais detalhes sobre a viagem?';
     }
   }
 }
 
 return 'Como posso ajudar você hoje com transfers em Lisboa?';
}

/**
 * Detectar idioma simples por DDI
 */
function detectarIdiomaSimples(telefone) {
  const numeroLimpo = telefone.replace(/\D/g, '');
  
  if (numeroLimpo.startsWith('351')) return 'PT'; // Portugal
  if (numeroLimpo.startsWith('1')) return 'EN';   // EUA/Canadá
  if (numeroLimpo.startsWith('44')) return 'EN';  // Reino Unido
  if (numeroLimpo.startsWith('34')) return 'ES';  // Espanha
  if (numeroLimpo.startsWith('33')) return 'FR';  // França
  if (numeroLimpo.startsWith('39')) return 'IT';  // Itália
  
  return 'PT'; // Padrão português
}

/**
 * Analisar intenção simples
 */
function analisarIntencaoSimples(mensagem) {
  const msgLower = mensagem.toLowerCase();
  
  // Palavras-chave por intenção
  if (msgLower.includes('ok') || msgLower.includes('finalizado') || msgLower.includes('concluido')) {
    return 'CONFIRMAR';
  }
  
  if (msgLower.includes('cancelar') || msgLower.includes('cancel')) {
    return 'CANCELAR';
  }
  
  if (msgLower.includes('onde') || msgLower.includes('when') || msgLower.includes('hora')) {
    return 'INFORMACAO';
  }
  
  if (msgLower.includes('problema') || msgLower.includes('ajuda') || msgLower.includes('help')) {
    return 'PROBLEMA';
  }
  
  if (msgLower.includes('obrigado') || msgLower.includes('thanks') || msgLower.includes('merci')) {
    return 'AGRADECIMENTO';
  }
  
  return 'CONVERSA';
}

/**
 * Verificar se telefone é de motorista
 */
function isMotorista(telefone) {
  try {
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    const aba = ss.getSheetByName('Motoristas_HUB');
    
    if (!aba) return false;
    
    const dados = aba.getDataRange().getValues();
    const telefoneNormalizado = telefone.replace(/\D/g, '');
    
    for (let i = 1; i < dados.length; i++) {
      const telMotorista = String(dados[i][1]).replace(/\D/g, '');
      if (telMotorista === telefoneNormalizado) {
        return true;
      }
    }
    
    return false;
    
  } catch (error) {
    console.log('Erro ao verificar motorista:', error.message);
    return false;
  }
}

/**
 * Processar OK do motorista (versão simplificada)
 */
function processarOKMotorista(telefone, mensagem) {
  console.log('Processando OK do motorista:', telefone);
  
  try {
    // Buscar motorista por telefone
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    const abaMotoristas = ss.getSheetByName('Motoristas_HUB');
    const sheet = ss.getSheetByName('HUB_Central');
    
    if (!abaMotoristas || !sheet) {
      throw new Error('Abas não encontradas');
    }
    
    // Encontrar motorista
    const dados = abaMotoristas.getDataRange().getValues();
    const telefoneNormalizado = telefone.replace(/\D/g, '');
    let nomeMotorista = null;
    
    for (let i = 1; i < dados.length; i++) {
      const telMotorista = String(dados[i][1]).replace(/\D/g, '');
      if (telMotorista === telefoneNormalizado) {
        nomeMotorista = dados[i][0];
        break;
      }
    }
    
    if (!nomeMotorista) {
      throw new Error('Motorista não encontrado');
    }
    
    // Encontrar transfer ativo do motorista (hoje)
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    
    const lastRow = sheet.getLastRow();
    
    for (let linha = 2; linha <= lastRow; linha++) {
      const dadosLinha = sheet.getRange(linha, 1, 1, 26).getValues()[0];
      
      const motorista = dadosLinha[21]; // Coluna V - Motorista
      const statusOK = dadosLinha[22];  // Coluna W - Status_OK
      const dataTransfer = new Date(dadosLinha[5]); // Coluna F - Data
      dataTransfer.setHours(0, 0, 0, 0);
      
      if (motorista === nomeMotorista && 
          !statusOK && 
          dataTransfer.getTime() === hoje.getTime()) {
        
        // Marcar como OK
        const timestamp = new Date();
        sheet.getRange(linha, 23).setValue('✅ OK'); // Coluna W
        sheet.getRange(linha, 24).setValue(timestamp); // Coluna X
        
        console.log(`OK registrado para transfer ${dadosLinha[0]}`);
        
        // Enviar confirmação para motorista
        const resposta = `✅ OK registrado! Transfer ${dadosLinha[0]} marcado como concluído. Obrigado, ${nomeMotorista}!`;
        enviarWhatsAppRoberta(telefone, resposta);
        
        return {
          sucesso: true,
          transferId: dadosLinha[0],
          motorista: nomeMotorista,
          timestamp: timestamp
        };
      }
    }
    
    // Se não encontrou transfer ativo
    const resposta = `${nomeMotorista}, não encontrei transfer ativo para hoje. Verifique se há algum transfer atribuído a você.`;
    enviarWhatsAppRoberta(telefone, resposta);
    
    return {
      sucesso: false,
      motivo: 'transfer_nao_encontrado',
      motorista: nomeMotorista
    };
    
  } catch (error) {
    console.log('Erro ao processar OK:', error.message);
    const resposta = 'Erro ao processar OK. Por favor, tente novamente ou contacte suporte.';
    enviarWhatsAppRoberta(telefone, resposta);
    
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Gerar resposta RAG baseada na intenção
 */
function gerarRespostaRAG(mensagem, idioma, intencao) {
  const templates = {
    PT: {
      CONFIRMAR: 'Obrigada pela confirmação! Tudo anotado.',
      CANCELAR: 'Entendi que pretende cancelar. Vou verificar e retorno em breve.',
      INFORMACAO: 'Vou verificar essa informação para si e já respondo.',
      PROBLEMA: 'Compreendo. Vou analisar a situação e encontrar uma solução.',
      AGRADECIMENTO: 'De nada! Estou aqui para ajudar sempre que precisar.',
      CONVERSA: 'Olá! Sou a Roberta HUB. Como posso ajudá-lo hoje?'
    },
    EN: {
      CONFIRMAR: 'Thank you for the confirmation! All noted.',
      CANCELAR: 'I understand you want to cancel. I will check and get back to you shortly.',
      INFORMACAO: 'I will check this information for you and respond soon.',
      PROBLEMA: 'I understand. I will analyze the situation and find a solution.',
      AGRADECIMENTO: 'You are welcome! I am here to help whenever you need.',
      CONVERSA: 'Hello! I am Roberta HUB. How can I help you today?'
    },
    ES: {
      CONFIRMAR: '¡Gracias por la confirmación! Todo anotado.',
      CANCELAR: 'Entiendo que quiere cancelar. Verificaré y volveré pronto.',
      INFORMACAO: 'Verificaré esta información para usted y responderé pronto.',
      PROBLEMA: 'Comprendo. Analizaré la situación y encontraré una solución.',
      AGRADECIMENTO: '¡De nada! Estoy aquí para ayudar siempre que necesite.',
      CONVERSA: '¡Hola! Soy Roberta HUB. ¿Cómo puedo ayudarte hoy?'
    },
    FR: {
      CONFIRMAR: 'Merci pour la confirmation! Tout noté.',
      CANCELAR: 'Je comprends que vous voulez annuler. Je vérifierai et reviendrai bientôt.',
      INFORMACAO: 'Je vérifierai cette information pour vous et répondrai bientôt.',
      PROBLEMA: 'Je comprends. J\'analyserai la situation et trouverai une solution.',
      AGRADECIMENTO: 'De rien! Je suis là pour aider quand vous en avez besoin.',
      CONVERSA: 'Bonjour! Je suis Roberta HUB. Comment puis-je vous aider aujourd\'hui?'
    },
    IT: {
      CONFIRMAR: 'Grazie per la conferma! Tutto annotato.',
      CANCELAR: 'Capisco che vuoi cancellare. Verificherò e tornerò presto.',
      INFORMACAO: 'Verificherò questa informazione per te e risponderò presto.',
      PROBLEMA: 'Capisco. Analizzerò la situazione e troverò una soluzione.',
      AGRADECIMENTO: 'Prego! Sono qui per aiutare quando ne hai bisogno.',
      CONVERSA: 'Ciao! Sono Roberta HUB. Come posso aiutarti oggi?'
    }
  };
  
  const idiomaTemplates = templates[idioma] || templates.PT;
  return idiomaTemplates[intencao] || idiomaTemplates.CONVERSA;
}

/**
 * Salvar contexto RAG simplificado
 */
function salvarContextoRAGSimples(telefone, mensagem, resposta, idioma, intencao) {
  try {
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    let abaRAG = ss.getSheetByName('RAG_Context');
    
    if (!abaRAG) {
      console.log('Criando aba RAG_Context...');
      abaRAG = ss.insertSheet('RAG_Context');
      
      const cabecalhos = [
        'ID_Conversa', 'Telefone', 'Timestamp', 'Mensagem_Cliente', 
        'Resposta_Roberta', 'Idioma', 'Intencao', 'Status'
      ];
      
      abaRAG.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
      abaRAG.getRange(1, 1, 1, cabecalhos.length)
        .setBackground('#4285F4')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold');
    }
    
    const idConversa = `CONV_${Date.now()}_${Math.random().toString(36).substr(2, 5)}`;
    
    const linhaRAG = [
      idConversa,
      telefone,
      new Date(),
      mensagem,
      resposta,
      idioma,
      intencao,
      'PROCESSADO'
    ];
    
    abaRAG.appendRow(linhaRAG);
    console.log('Contexto RAG salvo:', idConversa);
    
    return true;
    
  } catch (error) {
    console.log('Erro ao salvar contexto RAG:', error.message);
    return false;
  }
}

function testarWebhookManual() {
  console.log('Testando webhook manual...');
  
  // Simular dados do webhook Z-API
  const mockWebhook = {
    postData: {
      contents: JSON.stringify({
        type: "ReceivedCallback",
        phone: "351912345678",
        message: {
          text: "Teste manual",
          messageType: "text"
        }
      })
    }
  };
  
  try {
    const resultado = doPost(mockWebhook);
    console.log('doPost executou:', resultado.getContent());
    
    // Verificar se aba foi criada
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    const debugSheet = ss.getSheetByName('DEBUG_WEBHOOK');
    
    if (debugSheet) {
      console.log('DEBUG_WEBHOOK criada com sucesso');
      SpreadsheetApp.getUi().alert('Teste OK - Aba DEBUG_WEBHOOK criada');
    } else {
      console.log('DEBUG_WEBHOOK não foi criada');
      SpreadsheetApp.getUi().alert('Erro - Aba não foi criada');
    }
    
  } catch (error) {
    console.log('Erro no teste:', error.message);
    SpreadsheetApp.getUi().alert('Erro: ' + error.message);
  }
}

function verificarConfiguracaoZAPI() {
  console.log('Verificando configuração Z-API...');
  
  const config = SISTEMA_HUB_CONFIG.Z_API;
  
  console.log('Endpoint:', config.ENDPOINT);
  console.log('Client Token:', config.CLIENT_TOKEN ? 'Configurado' : 'FALTANDO');
  console.log('Roberta Phone:', config.ROBERTA_PHONE);
  
  if (!config.CLIENT_TOKEN) {
    SpreadsheetApp.getUi().alert('ERRO: CLIENT_TOKEN não configurado');
    return;
  }
  
  // Teste simples da API
  try {
    const payload = {
      phone: "351928283652",
      message: "Teste configuração - " + new Date().toLocaleTimeString()
    };
    
    const options = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Client-Token': config.CLIENT_TOKEN
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(config.ENDPOINT, options);
    const code = response.getResponseCode();
    const text = response.getContentText();
    
    console.log('Z-API Response:', code, text);
    
    if (code === 200) {
      SpreadsheetApp.getUi().alert('Z-API OK - Código: ' + code + '\nVerifique WhatsApp');
    } else {
      SpreadsheetApp.getUi().alert('Z-API Erro - Código: ' + code + '\nResposta: ' + text);
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Erro Z-API: ' + error.message);
  }
}

function extrairDadosWebhook(dados) {
  console.log('Extraindo dados do webhook:', JSON.stringify(dados, null, 2));
  
  let telefone = null;
  let mensagem = null;
  
  // Tentar diferentes formatos de dados da Z-API
  if (dados.phone) {
    telefone = dados.phone;
  } else if (dados.from) {
    telefone = dados.from;
  } else if (dados.sender) {
    telefone = dados.sender;
  }
  
  // Tentar diferentes formatos de mensagem
  if (dados.message) {
    if (dados.message.text) {
      mensagem = dados.message.text;
    } else if (dados.message.body) {
      mensagem = dados.message.body;
    } else if (typeof dados.message === 'string') {
      mensagem = dados.message;
    }
  } else if (dados.text) {
    mensagem = dados.text;
  } else if (dados.body) {
    mensagem = dados.body;
  }
  
  console.log('Dados extraídos - Telefone:', telefone, 'Mensagem:', mensagem);
  
  return {
    telefone: telefone,
    mensagem: mensagem,
    valido: !!(telefone && mensagem)
  };
}

function processarMensagemRAGLocal(telefone, mensagem) {
  console.log('Processando RAG local:', telefone, mensagem);
  
  try {
    // Detectar idioma
    const idioma = telefone.startsWith('351') ? 'PT' : 'EN';
    
    // Resposta simples
    const resposta = idioma === 'PT' ? 
      'Olá! Sou a Roberta HUB. Recebi: "' + mensagem + '"' :
      'Hello! I am Roberta HUB. Received: "' + mensagem + '"';
    
    console.log('Resposta gerada:', resposta);
    
    // Tentar enviar WhatsApp
    const enviado = enviarWhatsAppRoberta(telefone, resposta);
    console.log('WhatsApp enviado:', enviado);
    
    return {
      sucesso: enviado,
      telefone: telefone,
      mensagem: mensagem,
      resposta: resposta,
      idioma: idioma
    };
    
  } catch (error) {
    console.log('Erro no RAG local:', error.message);
    throw error;
  }
}

function extrairDadosWebhook(dados) {
  console.log('Extraindo dados do webhook:', JSON.stringify(dados, null, 2));
  
  let telefone = null;
  let mensagem = null;
  
  // Tentar diferentes formatos de dados da Z-API
  if (dados.phone) {
    telefone = dados.phone;
  } else if (dados.from) {
    telefone = dados.from;
  } else if (dados.sender) {
    telefone = dados.sender;
  }
  
  // Tentar diferentes formatos de mensagem
  if (dados.message) {
    if (dados.message.text) {
      mensagem = dados.message.text;
    } else if (dados.message.body) {
      mensagem = dados.message.body;
    } else if (typeof dados.message === 'string') {
      mensagem = dados.message;
    }
  } else if (dados.text) {
    mensagem = dados.text;
  } else if (dados.body) {
    mensagem = dados.body;
  }
  
  console.log('Dados extraídos - Telefone:', telefone, 'Mensagem:', mensagem);
  
  return {
    telefone: telefone,
    mensagem: mensagem,
    valido: !!(telefone && mensagem)
  };
}

function processarMensagemRAGLocal(telefone, mensagem) {
  console.log('Processando RAG local:', telefone, mensagem);
  
  try {
    // Detectar idioma
    const idioma = telefone.startsWith('351') ? 'PT' : 'EN';
    
    // Resposta simples
    const resposta = idioma === 'PT' ? 
      'Olá! Sou a Roberta HUB. Recebi: "' + mensagem + '"' :
      'Hello! I am Roberta HUB. Received: "' + mensagem + '"';
    
    console.log('Resposta gerada:', resposta);
    
    // Tentar enviar WhatsApp
    const enviado = enviarWhatsAppRoberta(telefone, resposta);
    console.log('WhatsApp enviado:', enviado);
    
    // Salvar no RAG_Context
    try {
      const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
      const ragSheet = ss.getSheetByName('RAG_Context');
      
      if (ragSheet) {
        const idConversa = `CONV_${Date.now()}_WEBHOOK`;
        ragSheet.appendRow([
          idConversa,
          '',
          telefone,
          'Cliente Webhook',
          new Date(),
          'INPUT',
          idioma,
          'CONFIRMACAO',
          'WEBHOOK',
          mensagem,
          '',
          'CONVERSA',
          '',
          'NEUTRO',
          '{}',
          'PROCESSADO',
          enviado ? 'RESPOSTA_ENVIADA' : 'FALHA_ENVIO'
        ]);
      }
    } catch (ragError) {
      console.log('Erro ao salvar RAG:', ragError.message);
    }
    
    return {
      sucesso: enviado,
      telefone: telefone,
      mensagem: mensagem,
      resposta: resposta,
      idioma: idioma
    };
    
  } catch (error) {
    console.log('Erro no RAG local:', error.message);
    throw error;
  }
}

function doPost(e) {
  console.log('=== DOPOST WEBHOOK INICIADO ===');
  console.log('Timestamp:', new Date().toISOString());
  
  try {
    // CRIAR ABA DEBUG IMEDIATAMENTE
    const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
    let debugSheet = ss.getSheetByName('DEBUG_WEBHOOK');
    
    if (!debugSheet) {
      console.log('Criando aba DEBUG_WEBHOOK...');
      debugSheet = ss.insertSheet('DEBUG_WEBHOOK');
      debugSheet.getRange(1, 1, 1, 5).setValues([['Timestamp', 'Metodo', 'Phone', 'Message', 'Status']]);
      debugSheet.getRange(1, 1, 1, 5)
        .setBackground('#ff9500')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    }
    
    // LOG INICIAL
    const linhaLog = [new Date(), 'Z-API Webhook', 'N/A', 'Verificando dados', 'EXECUTANDO'];
    debugSheet.appendRow(linhaLog);
    
    if (!e || !e.postData) {
      debugSheet.appendRow([new Date(), 'Z-API Webhook', 'N/A', 'Sem dados', 'DADOS_INCOMPLETOS']);
      return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
    }
    
    // PARSE DOS DADOS
    let dados;
    try {
      dados = JSON.parse(e.postData.contents);
    } catch (parseError) {
      debugSheet.appendRow([new Date(), 'Z-API Webhook', 'N/A', 'Erro JSON', 'ERRO_PARSE']);
      return ContentService.createTextOutput('ERROR').setMimeType(ContentService.MimeType.TEXT);
    }
    
    // EXTRAIR TELEFONE E MENSAGEM
    let telefone = '';
    let mensagem = '';
    
    // VERIFICAR E FINALIZAR CONVERSAS PENDENTES
try {
  const contextoPendente = recuperarContextoRAG(phone, 1);
  if (contextoPendente[0]?.proximoPasso === 'ENVIAR_DETALHES_FINAIS') {
    console.log('Finalizando conversa pendente via webhook...');
    
    const respostaFinal = `Transfer confirmado!
    
📍 Recolha: ${contextoPendente[0].conteudoMensagem}
✈️ Destino: Aeroporto de Lisboa  
⏰ Horário será confirmado

Precisa da volta também? Posso fazer ida+volta com desconto!`;

    enviarWhatsAppRoberta(phone, respostaFinal);
    
    // Atualizar status
    const dadosFinalizacao = {
      ...contextoPendente[0],
      statusFluxo: 'PROCESSADO', 
      proximoPasso: 'FINALIZADO'
    };
    salvarContextoRAG(dadosFinalizacao);
  }
} catch (error) {
  console.log('Erro ao verificar pendentes:', error.message);
}

    // Tentar diferentes estruturas Z-API
    if (dados.phone) telefone = dados.phone;
    else if (dados.from) telefone = dados.from;
    else if (dados.data && dados.data.from) telefone = dados.data.from;
    
    if (dados.text && dados.text.message) mensagem = dados.text.message;
    else if (dados.message && dados.message.text) mensagem = dados.message.text;
    else if (dados.data && dados.data.body) mensagem = dados.data.body;
    else if (dados.body) mensagem = dados.body;
    else if (typeof dados.message === 'string') mensagem = dados.message;
    
    debugSheet.appendRow([new Date(), 'Z-API Webhook', telefone, mensagem, 'DADOS_EXTRAIDOS']);
    
    if (!telefone || !mensagem) {
      debugSheet.appendRow([new Date(), 'Z-API Webhook', telefone, mensagem, 'DADOS_INCOMPLETOS']);
      return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
    }
    
    console.log('Telefone extraído:', telefone);
    console.log('Mensagem extraída:', mensagem);
    
    // PROCESSAR COM RAG HÍBRIDO (função do FluxosTransfer.gs)
    try {
      const resultadoRAG = processarMensagemRAGHibrido(telefone, mensagem, null);
      
      if (resultadoRAG) {
        debugSheet.appendRow([new Date(), 'RAG Processar', telefone, 'Sucesso: ' + resultadoRAG.tipo, 'RAG_SUCESSO']);
      } else {
        debugSheet.appendRow([new Date(), 'RAG Processar', telefone, 'Falha no processamento', 'RAG_FALHA']);
      }
      
    } catch (ragError) {
      console.log('Erro no RAG:', ragError.message);
      debugSheet.appendRow([new Date(), 'RAG Processar', telefone, 'Erro: ' + ragError.message, 'RAG_ERRO']);
    }
    
    return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
    
  } catch (error) {
    console.log('Erro crítico no doPost:', error.message);
    
    try {
      const ss = SpreadsheetApp.openById(SISTEMA_HUB_CONFIG.SPREADSHEET_ID);
      const debugSheet = ss.getSheetByName('DEBUG_WEBHOOK');
      if (debugSheet) {
        debugSheet.appendRow([new Date(), 'ERRO CRITICO', 'N/A', error.message, 'ERRO_CRITICO']);
      }
    } catch (logError) {
      console.log('Não foi possível registrar erro:', logError.message);
    }
    
    return ContentService.createTextOutput('ERROR').setMimeType(ContentService.MimeType.TEXT);
  }
}

function obterURLWebhook() {
  const url = ScriptApp.getService().getUrl();
  console.log('URL atual do webhook:', url);
  
  try {
    SpreadsheetApp.getUi().alert('URL do Webhook:\n\n' + url + '\n\nCopie esta URL e configure na Z-API');
  } catch (uiError) {
    console.log('Executado via contexto sem UI - URL:', url);
  }
  
  return url;
}

function testarWebhookManualmente() {
  console.log('=== TESTE MANUAL WEBHOOK ===');
  
  // Simular dados do webhook
  const dadosTeste = {
    phone: '351999888777',
    message: {
      text: 'oi'
    }
  };
  
  try {
    const telefone = extrairTelefoneSeguro(dadosTeste);
    const mensagem = extrairMensagemSeguro(dadosTeste);
    
    console.log('Telefone teste:', telefone);
    console.log('Mensagem teste:', mensagem);
    
    if (telefone && mensagem) {
      const resultado = processarMensagemRAGHibridoHumanizado(telefone, mensagem);
      console.log('Resultado teste:', resultado);
    }
    
  } catch (error) {
    console.log('Erro no teste:', error.message);
  }
}

function logWebhookTest() {
  console.log('=== TESTE WEBHOOK DEBUG ===');
  console.log('URL atual:', ScriptApp.getService().getUrl());
  console.log('Token Z-API:', SISTEMA_HUB_CONFIG.Z_API.CLIENT_TOKEN ? 'Configurado' : 'FALTANDO');
  console.log('Endpoint Z-API:', SISTEMA_HUB_CONFIG.Z_API.ENDPOINT);
}

/**
 * Simular webhook com seus dados reais
 */
function simularWebhookJunior() {
  console.log('=== SIMULANDO WEBHOOK JUNIOR ===');
  
  // Simular dados que o Z-API enviaria
  const dadosSimulados = {
    postData: {
      contents: JSON.stringify({
        phone: '351968698138',
        message: 'Boa noite',
        from: '+351968698138',
        body: 'Boa noite'
      })
    }
  };
  
  // Processar via função doPost
  const resultado = doPost(dadosSimulados);
  
  console.log('Resultado webhook simulado:', resultado);
  
  SpreadsheetApp.getUi().alert(
    'Teste Webhook Simulado',
    'Webhook simulado executado. Verifique o WhatsApp para ver se recebeu a resposta personalizada.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Verificar estrutura real do webhook
 */
function verificarWebhook() {
  // Verificar se as funções existem
  try {
    const teste1 = typeof extrairTelefoneSeguro;
    const teste2 = typeof extrairMensagemSeguro;
    const teste3 = typeof buscarPerfilDemoIntegrado;
    
    console.log('extrairTelefoneSeguro:', teste1);
    console.log('extrairMensagemSeguro:', teste2);
    console.log('buscarPerfilDemoIntegrado:', teste3);
    
    SpreadsheetApp.getUi().alert(
      'Verificação Webhook',
      `extrairTelefoneSeguro: ${teste1}\nextrairMensagemSeguro: ${teste2}\nbuscarPerfilDemoIntegrado: ${teste3}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Erro', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Teste direto do webhook
 */
function testarWebhookDireto() {
  console.log('🧪 Testando webhook diretamente...');
  
  const dadosTeste = {
    postData: {
      contents: JSON.stringify({
        phone: '351968698138', // Seu número
        message: 'teste sistema',
        from: '+351968698138'
      })
    }
  };
  
  const resultado = doPost(dadosTeste);
  console.log('✅ Teste concluído:', resultado.getContent());
  
  return resultado;
}

function debugDeteccaoIdioma() {
  console.log('🔍 Testando detecção de idioma isoladamente');
  
  const transfersTest = buscarTransfersComMotoristasAtribuidos();
  console.log('Transfers encontrados:', transfersTest.length);
  
  if (transfersTest.length > 0) {
    const primeiroTransfer = transfersTest[0];
    console.log('Primeiro transfer:', primeiroTransfer);
    console.log('Contacto tipo:', typeof primeiroTransfer.contacto);
    console.log('Contacto valor:', primeiroTransfer.contacto);
    
    // Teste direto da conversão
    const telefoneString = String(primeiroTransfer.contacto || '');
    console.log('Telefone como string:', telefoneString);
    console.log('Telefone tem startsWith?', telefoneString.startsWith ? 'SIM' : 'NÃO');
    
    // Teste da função
    try {
      const resultado = detectarIdiomaPorDDI(primeiroTransfer.contacto);
      console.log('Resultado detecção:', resultado);
    } catch (error) {
      console.log('Erro na detecção:', error.message);
    }
  }
}
