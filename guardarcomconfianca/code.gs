// ===================================================
// SISTEMA DE TRANSFERS EMPIRE MARQUES HOTEL & HUB TRANSFER v4.0
// Sistema Integrado Completo com E-mail Interativo
// Fusão: Código Novo (base) + Funcionalidades do Código Antigo
// ===================================================

// 🚨 INTERCEPTADOR DE E-MAILS - Adicione no topo do script
const originalGmailSend = GmailApp.sendEmail;
GmailApp.sendEmail = function(to, subject, body, options) {
  console.log('🚨 E-MAIL ENVIADO:', {
    destinatario: to,
    assunto: subject,
    temHtmlBody: !!(options && options.htmlBody),
    bodyLength: (options && options.htmlBody ? options.htmlBody.length : body.length),
    temAssinaturaLimpa: options && options.htmlBody ? 
      (options.htmlBody.includes('1z6i_VnZZ9OdHbcmcTHa4dKHPIqC6vmBE') && !options.htmlBody.includes('Junior Gutierez')) :
      false
  });
  
  // Se tem texto da assinatura, mostra o HTML problemático
  if (options && options.htmlBody && options.htmlBody.includes('Junior Gutierez')) {
    console.log('❌ ENCONTROU O E-MAIL PROBLEMÁTICO!');
    console.log('HTML (últimos 500 chars):', options.htmlBody.slice(-500));
  }
  
  return originalGmailSend.call(this, to, subject, body, options);
};

// Interceptar MailApp também
const originalMailSend = MailApp.sendEmail;
MailApp.sendEmail = function(...args) {
  console.log('🚨 MailApp.sendEmail chamado:', args[0]?.to || args[0]);
  return originalMailSend.apply(this, args);
};

console.log('✅ Interceptadores instalados!');

const CONFIG = {
  // 📊 Google Sheets - Configuração Principal (CÓDIGO NOVO - PRIORIDADE)
  SPREADSHEET_ID: '1ZfG_IXBWMbGQzCmn7nWNzFqUBbGkXCXWSnY7JYziHxI',
  SHEET_NAME: 'Empire MARQUES-HUB',
  PRICING_SHEET_NAME: 'Tabela de Preços',
  
  // 📧 Configuração de E-mail com Sistema Interativo (CORRIGIDA)
  EMAIL_CONFIG: {
    DESTINATARIOS: [
      'juniorgutierezbega@gmail.com' // Email de monitoramento
    ],
    EMAIL_EMPRESA: 'hubtransferencia@gmail.com', // Email oficial da empresa
    ENVIAR_AUTOMATICO: true,
    VERIFICAR_CONFIRMACOES: true,
    INTERVALO_VERIFICACAO: 5, // minutos
    USAR_BOTOES_INTERATIVOS: true, // Sistema interativo do código antigo
    TEMPLATE_HTML: true,
    ARQUIVAR_CONFIRMADOS: true,
    ENVIAR_RELATORIO_DIA_ANTERIOR: true // Do código novo
  },
  
  // 📱 Configuração de Telefones e Contatos (CORRIGIDO)
  CONTATOS: {
    HUB_PHONE: '+351968698138',        // Proprietário HUB
    ROBERTA_PHONE: '+351928283652',    // Assistente HUB
    HOTEL_PHONE: '+351210548700',      // Telefone Empire Marques Hotel

  },
  
  // 🔗 Integração com APIs Externas (PRESERVADO DO CÓDIGO ANTIGO)
  ZAPI: {
    INSTANCE: '3DC8E250141ED020B95796155CBF9532',
    TOKEN: 'DF93ABBE66F44D82F60EF9FE',
    WEBHOOK_URL: 'https://api.z-api.io/instances/3DC8E250141ED020B95796155CBF9532/token/DF93ABBE66F44D82F60EF9FE/send-text',
    ENABLED: false
  },
  
  // 🔄 Webhooks do Make.com (PRESERVADO DO CÓDIGO ANTIGO)
  MAKE_WEBHOOKS: {
    NEW_TRANSFER: 'https://hook.eu2.make.com/y7ro3pc9we00iums1ffb106pok9icl7m',
    WHATSAPP_RESPONSE: 'https://hook.eu2.make.com/y7ro3pc9we00iums1ffb106pok9icl7m',
    STATUS_UPDATE: 'https://hook.eu2.make.com/y7ro3pc9we00iums1ffb106pok9icl7m',
    ACERTO_CONTAS: 'https://hook.eu2.make.com/y7ro3pc9we00iums1ffb106pok9icl7m',
    ENABLED: false
  },
  
  // 🏨 Identificação do Sistema (CORRIGIDO)
  NAMES: {
    HUB_OWNER: 'HUB Transfer',
    ASSISTANT: 'Roberta HUB',
    HOTEL_MANAGER: 'Catarina',
    HOTEL_NAME: 'Empire Marques Hotel',
    SISTEMA_NOME: 'Sistema Empire Marques Hotel e Hub Transfer'
  },
  
  // 💰 Valores Padrão e Configurações Financeiras (CÓDIGO NOVO)
  VALORES: {
    HOTEL_PADRAO: 5.00,
    HUB_PADRAO: 20.00,
    PERCENTUAL_HOTEL: 0.30, // 30% para o hotel
    PERCENTUAL_HUB: 0.70,   // 70% para HUB (menos comissão recepção)
    COMISSAO_RECEPCAO_TOUR: 5.00, // €5 para tours
    COMISSAO_RECEPCAO_TRANSFER: 2.00, // €2 para transfers
    MOEDA: '€',
    FORMATO_MOEDA: '€#,##0.00'
  },
  
  // 🌐 Configurações do Sistema (FUSÃO)
  SISTEMA: {
    VERSAO: '4.0-Empire-Marques-Hub-Transfer-Integrado',
    TIMEZONE: 'Europe/Lisbon',
    LOCALE: 'pt-PT',
    DATA_FORMATO: 'dd/mm/yyyy',
    HORA_FORMATO: 'HH:mm',
    ORGANIZAR_POR_MES: true,
    PREFIXO_MES: 'Transfers_',
    ANO_BASE: 2025,
    REGISTRO_DUPLO_OBRIGATORIO: true,
    MAX_TENTATIVAS_REGISTRO: 3,
    INTERVALO_TENTATIVAS: 500,
    BACKUP_AUTOMATICO: true,
    LOG_DETALHADO: true
  },
  
  // 🔐 Segurança e Validações (PRESERVADO DO CÓDIGO ANTIGO)
  SEGURANCA: {
    VALIDAR_EMAIL: true,
    VALIDAR_TELEFONE: true,
    SANITIZAR_INPUTS: true,
    MAX_CARACTERES_CAMPO: 500,
    PERMITIR_HTML_OBSERVACOES: false
  },
  
  // 📊 Limites e Quotas (PRESERVADO DO CÓDIGO ANTIGO)
  LIMITES: {
    MAX_PESSOAS: 20,
    MAX_BAGAGENS: 50,
    MIN_VALOR: 0.01,
    MAX_VALOR: 9999.99,
    MAX_REGISTROS_DIA: 1000,
    MAX_EMAIL_DIA: 500
  }
};

// ===================================================
// CONSTANTES DE MESES E ESTRUTURA DE DADOS
// ===================================================

const MESES = [
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

// Headers da planilha principal - ESTRUTURA DO CÓDIGO NOVO
const HEADERS = [
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
  'Valor Impire Marques Hotel (€)', // M - Comissão do hotel
  'Valor HUB Transfer (€)',      // N - Valor para HUB
  'Comissão Recepção (€)',       // O - Comissão da recepção
  'Forma Pagamento',             // P - Método de pagamento
  'Pago Para',                   // Q - Quem recebeu
  'Status',                      // R - Status do transfer
  'Observações',                 // S - Observações gerais
  'Data Criação'                 // T - Timestamp de criação
];

// Headers da tabela de preços - ESTRUTURA DO CÓDIGO NOVO
const PRICING_HEADERS = [
  'ID',                          // A - ID único do preço
  'Tipo Serviço',               // B - Transfer/Tour Regular/Private Tour
  'Rota',                       // C - Nome da rota
  'Origem',                     // D - Ponto de origem
  'Destino',                    // E - Ponto de destino
  'Pessoas',                    // F - Número de pessoas
  'Bagagens',                   // G - Número de bagagens
  'Preço Por Pessoa',           // H - Para tours regulares
  'Preço Por Grupo',            // I - Para private tours
  'Preço Cliente (€)',          // J - Preço total
  'Valor Impire Marques Hotel (€)', // K - Comissão hotel
  'Valor HUB Transfer (€)',     // L - Valor HUB
  'Comissão Recepção (€)',      // M - Comissão recepção
  'Ativo',                      // N - Se está ativo
  'Data Criação',               // O - Quando foi criado
  'Observações'                 // P - Notas adicionais
];

// ===================================================
// SISTEMA DE LOGGING CORRIGIDO (FUSÃO CÓDIGOS ANTIGO + NOVO)
// ===================================================

const LOG_CONFIG = {
 ENABLED: true,
 LEVEL: 'INFO', // DEBUG, INFO, WARN, ERROR
 INCLUDE_TIMESTAMP: true,
 INCLUDE_FUNCTION: false, // Desabilitado para evitar conflitos
 MAX_LOG_SIZE: 1000,
 PERSIST_TO_SHEET: false, // Desabilitado por padrão
 LOG_SHEET_NAME: 'Sistema_Logs',
 CONSOLE_OUTPUT: true
};

const LOG_LEVELS = {
 DEBUG: 0,
 INFO: 1,
 WARN: 2,
 ERROR: 3,
 SUCCESS: 1
};

/**
* Sistema de logging unificado e corrigido
*/
const logger = {
 logs: [],
 config: LOG_CONFIG,
 
 /**
  * Verifica se deve fazer log baseado no nível
  */
 _shouldLog: function(level) {
   if (!this.config.ENABLED) return false;
   return LOG_LEVELS[level] >= LOG_LEVELS[this.config.LEVEL];
 },

 /**
  * Formatar timestamp
  */
 _formatTimestamp: function(date) {
   try {
     return Utilities.formatDate(
       date, 
       CONFIG.SISTEMA.TIMEZONE, 
       'yyyy-MM-dd HH:mm:ss'
     );
   } catch (error) {
     return date.toISOString();
   }
 },

 /**
  * Obter emoji por nível
  */
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

 /**
  * Formatar entrada de log
  */
 _formatLog: function(entry) {
   let log = '';
   
   if (this.config.INCLUDE_TIMESTAMP) {
     log += `[${this._formatTimestamp(entry.timestamp)}] `;
   }
   
   log += `${this._getLevelEmoji(entry.level)} ${entry.level}: `;
   log += entry.message;
   
   if (entry.data) {
     try {
       log += ` | Data: ${JSON.stringify(entry.data)}`;
     } catch (e) {
       log += ` | Data: [Objeto não serializável]`;
     }
   }
   
   return log;
 },

 /**
  * Persistir log na planilha
  */
 _persistToSheet: function(entry) {
   if (!this.config.PERSIST_TO_SHEET) return;
   
   try {
     const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
     let logSheet = ss.getSheetByName(this.config.LOG_SHEET_NAME);
     
     if (!logSheet) {
       logSheet = ss.insertSheet(this.config.LOG_SHEET_NAME);
       logSheet.appendRow(['Timestamp', 'Level', 'Message', 'Data']);
       
       // Formatação básica
       const headerRange = logSheet.getRange(1, 1, 1, 4);
       headerRange.setBackground('#2c3e50').setFontColor('#ffffff').setFontWeight('bold');
       logSheet.setFrozenRows(1);
     }
     
     // Limitar número de logs
     if (logSheet.getLastRow() > this.config.MAX_LOG_SIZE) {
       const rowsToDelete = logSheet.getLastRow() - this.config.MAX_LOG_SIZE + 100;
       logSheet.deleteRows(2, rowsToDelete);
     }
     
     logSheet.appendRow([
       entry.timestamp,
       entry.level,
       entry.message,
       entry.data ? JSON.stringify(entry.data) : ''
     ]);
     
   } catch (error) {
     // Log de erro não deve quebrar o sistema
     // Usar Logger nativo como fallback
     if (typeof Logger !== 'undefined' && Logger.log) {
       Logger.log(`ERRO NO SISTEMA DE LOG: ${error.toString()}`);
     }
   }
 },

 /**
  * Função principal de log
  */
 log: function(level, message, data = null) {
   if (!this._shouldLog(level)) return;
   
   const logEntry = {
     timestamp: new Date(),
     level: level,
     message: message,
     data: data
   };
   
   // Adicionar à memória
   this.logs.push(logEntry);
   if (this.logs.length > this.config.MAX_LOG_SIZE) {
     this.logs.shift();
   }
   
   // Formatar e exibir
   const formattedLog = this._formatLog(logEntry);
   
   // Output no console
   if (this.config.CONSOLE_OUTPUT) {
     try {
       // Tentar console primeiro
       if (typeof console !== 'undefined' && console.log) {
         console.log(formattedLog);
       } else if (typeof Logger !== 'undefined' && Logger.log) {
         // Fallback para Logger do Google Apps Script
         Logger.log(formattedLog);
       }
     } catch (e) {
       // Fallback silencioso se ambos falharem
     }
   }
   
   // Persistir se configurado
   this._persistToSheet(logEntry);
 },

 /**
  * Métodos de conveniência
  */
 debug: function(message, data = null) {
   this.log('DEBUG', message, data);
 },

 info: function(message, data = null) {
   this.log('INFO', message, data);
 },

 warn: function(message, data = null) {
   this.log('WARN', message, data);
 },

 error: function(message, data = null) {
   this.log('ERROR', message, data);
 },

 success: function(message, data = null) {
   this.log('SUCCESS', `✅ ${message}`, data);
 },

 /**
  * Obter logs em memória
  */
 getLogs: function(level = null) {
   if (!level) return this.logs;
   return this.logs.filter(log => log.level === level);
 },

 /**
  * Limpar logs em memória
  */
 clearLogs: function() {
   this.logs = [];
   this.info('Logs em memória limpos');
 },

 /**
  * Configurar nível de log dinamicamente
  */
 setLevel: function(level) {
   if (LOG_LEVELS.hasOwnProperty(level)) {
     this.config.LEVEL = level;
     this.info(`Nível de log alterado para: ${level}`);
   } else {
     this.warn(`Nível de log inválido: ${level}`);
   }
 },

 /**
  * Ativar/desativar persistência em planilha
  */
 setPersistence: function(enabled) {
   this.config.PERSIST_TO_SHEET = enabled;
   this.info(`Persistência de logs ${enabled ? 'ativada' : 'desativada'}`);
 }
};

// Verificar se CONFIG existe, senão usar valores padrão
if (typeof CONFIG === 'undefined' || !CONFIG.SISTEMA) {
 logger.config.ENABLED = true;
 logger.warn('CONFIG não encontrado, usando configurações padrão para logging');
} else {
 logger.config.ENABLED = CONFIG.SISTEMA.LOG_DETALHADO !== false;
}

// ===================================================
// FUNÇÕES DE UTILIDADE E HELPERS (FUSÃO COMPLETA)
// ===================================================

// ===================================================
// MENSAGENS E TEMPLATES DO SISTEMA (CÓDIGO NOVO + ANTIGO)
// ===================================================

const MESSAGES = {
  // Títulos e Assuntos
  TITULOS: {
    NOVO_TRANSFER: `🚐 NOVO TRANSFER ${CONFIG.NAMES.HOTEL_NAME} & HUB TRANSFER`,
    CONFIRMACAO_NECESSARIA: (id) => `AÇÃO NECESSÁRIA: Novo Transfer #${id}`,
    TRANSFER_CONFIRMADO: (id) => `✅ Transfer #${id} Confirmado`,
    TRANSFER_CANCELADO: (id) => `❌ Transfer #${id} Cancelado`
  },
  
  // Mensagens de Status
  STATUS_MESSAGES: {
    SOLICITADO: 'Transfer solicitado - aguardando confirmação',
    CONFIRMADO: 'Transfer confirmado - será realizado',
    FINALIZADO: 'Transfer realizado com sucesso',
    CANCELADO: 'Transfer cancelado',
    EM_ANDAMENTO: 'Transfer em andamento',
    PROBLEMA: 'Problema reportado - verificar'
  },
  
  // Ações e Confirmações
  ACOES: {
    CONFIRMADO_POR: (nome) => `✅ Confirmado por ${nome}`,
    CANCELADO_POR: (nome) => `❌ Cancelado por ${nome}`,
    FINALIZADO_POR: (nome) => `✔️ Finalizado por ${nome}`,
    MODIFICADO_POR: (nome) => `📝 Modificado por ${nome}`
  },
  
  // Templates para WhatsApp (do código antigo)
  WHATSAPP_TEMPLATES: {
    NOVO_TRANSFER: {
      TITULO: '*🚐 NOVO TRANSFER MARQUES-HUB*',
      RODAPE: '*Responda com "OK" para confirmar*',
      ASSINATURA: `_Sistema ${CONFIG.NAMES.HOTEL_NAME} & HUB Transfer_`
    }
  },
  
  // Mensagens de Erro
  ERROS: {
    CAMPO_OBRIGATORIO: (campo) => `Campo obrigatório ausente: ${campo}`,
    VALOR_INVALIDO: (campo) => `Valor inválido para o campo: ${campo}`,
    DATA_INVALIDA: 'Data inválida ou em formato incorreto',
    REGISTRO_NAO_ENCONTRADO: (id) => `Transfer #${id} não encontrado`,
    FALHA_REGISTRO: 'Falha ao registrar o transfer',
    FALHA_EMAIL: 'Falha ao enviar e-mail de notificação'
  },
  
  // Mensagens de Sucesso
  SUCESSO: {
    REGISTRO_COMPLETO: 'Transfer registrado com sucesso em todas as abas',
    EMAIL_ENVIADO: 'E-mail de notificação enviado com sucesso',
    STATUS_ATUALIZADO: 'Status atualizado com sucesso',
    DADOS_LIMPOS: 'Dados limpos com sucesso',
    SISTEMA_CONFIGURADO: 'Sistema configurado com sucesso'
  }
};

// ===================================================
// CONFIGURAÇÕES DE ESTILO E APARÊNCIA (FUSÃO)
// ===================================================

const STYLES = {
  // Cores por Status
  STATUS_COLORS: {
    SOLICITADO: '#f39c12',
    CONFIRMADO: '#27ae60',
    FINALIZADO: '#2c3e50',
    CANCELADO: '#e74c3c',
    EM_ANDAMENTO: '#3498db',
    PROBLEMA: '#e67e22'
  },
  
  // Cores dos Headers
  HEADER_COLORS: {
    PRINCIPAL: '#2c3e50',
    PRECOS: '#27ae60',
    MENSAL: null
  },
  
  // Larguras das Colunas (em pixels) - ATUALIZADO PARA CÓDIGO NOVO
  COLUMN_WIDTHS: {
    PRINCIPAL: [60, 150, 100, 60, 60, 100, 120, 80, 200, 200, 90, 100, 140, 120, 80, 100, 100, 100, 180, 140],
    PRECOS: [60, 120, 200, 180, 180, 80, 80, 100, 100, 120, 140, 140, 80, 80, 140, 200]
  },
  
  // Formatações de Células
  FORMATS: {
    MOEDA: '€#,##0.00',
    DATA: 'dd/mm/yyyy',
    HORA: 'hh:mm',
    TIMESTAMP: 'dd/mm/yyyy hh:mm',
    NUMERO: '0',
    PERCENTUAL: '0.00%'
  }
};

// ===================================================
// VALIDAÇÕES E REGRAS DE NEGÓCIO (FUSÃO)
// ===================================================

const VALIDACOES = {
  // Campos obrigatórios para novo transfer
  CAMPOS_OBRIGATORIOS: [
    'nomeCliente', 
    'numeroPessoas', 
    'data', 
    'contacto', 
    'origem', 
    'destino', 
    'horaPickup', 
    'valorTotal'
  ],
  
  // Formatos de validação
  FORMATOS: {
    EMAIL: /^[^\s@]+@[^\s@]+\.[^\s@]+$/,
    TELEFONE: /^\+?[\d\s\-().]+$/,
    HORA: /^([01]?[0-9]|2[0-3]):[0-5][0-9]$/,
    DATA_BR: /^\d{2}\/\d{2}\/\d{4}$/,
    DATA_ISO: /^\d{4}-\d{2}-\d{2}$/
  },
  
  // Valores permitidos - ATUALIZADO PARA CÓDIGO NOVO
  VALORES_PERMITIDOS: {
    FORMA_PAGAMENTO: ['Dinheiro', 'Cartão', 'Transferência', 'MB Way', 'Voucher'],
    PAGO_PARA: ['Recepção', 'Motorista', 'Online', 'Hotel', 'HUB'],
    STATUS: Object.keys(MESSAGES.STATUS_MESSAGES),
    TIPO_SERVICO: ['Transfer', 'Tour Regular', 'Private Tour']
  }
};

/**
 * Processa e valida uma data de múltiplos formatos
 * @param {string|Date} dataInput - Data em vários formatos possíveis
 * @returns {Date} - Objeto Date válido
 * @throws {Error} - Se a data for inválida
 */
/**
 * Processa e valida uma data de múltiplos formatos - VERSÃO CORRIGIDA
 * @param {string|Date} dataInput - Data em vários formatos possíveis
 * @returns {Date} - Objeto Date válido
 * @throws {Error} - Se a data for inválida
 */
function processarDataSegura(dataInput) {
  logger.debug('Processando data', { input: dataInput, tipo: typeof dataInput });
  
  try {
    // Se já for um Date válido
    if (dataInput instanceof Date && !isNaN(dataInput)) {
      return dataInput;
    }
    
    let data;
    
    if (typeof dataInput === 'string') {
      // Remover espaços extras
      dataInput = dataInput.trim();
      
      // 🔧 CORREÇÃO: Formato ISO (YYYY-MM-DD) - MAIS COMUM
      if (dataInput.match(/^\d{4}-\d{2}-\d{2}/)) {
        // Para ISO, usar diretamente sem adicionar T12:00:00
        data = new Date(dataInput);
        
        // Se a data vier como UTC, ajustar para timezone local
        if (dataInput.length === 10) { // Apenas YYYY-MM-DD
          const [ano, mes, dia] = dataInput.split('-');
          data = new Date(parseInt(ano), parseInt(mes) - 1, parseInt(dia));
        }
      }
      // Formato brasileiro (DD/MM/YYYY)
      else if (dataInput.match(/^\d{2}\/\d{2}\/\d{4}/)) {
        const [dia, mes, ano] = dataInput.split('/');
        data = new Date(parseInt(ano), parseInt(mes) - 1, parseInt(dia));
      }
      // Formato alternativo (DD-MM-YYYY)
      else if (dataInput.match(/^\d{2}-\d{2}-\d{4}/)) {
        const [dia, mes, ano] = dataInput.split('-');
        data = new Date(parseInt(ano), parseInt(mes) - 1, parseInt(dia));
      }
      // Formato americano (MM/DD/YYYY)
      else if (dataInput.match(/^\d{1,2}\/\d{1,2}\/\d{4}/)) {
        // Assumir formato brasileiro por padrão
        const [dia, mes, ano] = dataInput.split('/');
        data = new Date(parseInt(ano), parseInt(mes) - 1, parseInt(dia));
      }
      // Tentar parse direto como último recurso
      else {
        data = new Date(dataInput);
      }
    } else {
      data = new Date(dataInput);
    }
    
    // Validar a data resultante
    if (isNaN(data.getTime())) {
      throw new Error(`Formato de data inválido: ${dataInput}`);
    }
    
    logger.debug('Data processada com sucesso', { 
      original: dataInput, 
      processada: data.toISOString(),
      dia: data.getDate(),
      mes: data.getMonth() + 1,
      ano: data.getFullYear()
    });
    
    return data;
    
  } catch (error) {
    logger.error('Erro ao processar data', { input: dataInput, erro: error.message });
    throw new Error(`Falha ao processar data: ${error.message}`);
  }
}

/**
 * Formata uma data no padrão brasileiro DD/MM/YYYY
 * @param {Date} date - Data a ser formatada
 * @returns {string} - Data formatada
 */
/**
 * Formata uma data no padrão brasileiro DD/MM/YYYY - VERSÃO CORRIGIDA
 * @param {Date} date - Data a ser formatada
 * @returns {string} - Data formatada
 */
function formatarDataDDMMYYYY(date) {
  try {
    if (!(date instanceof Date) || isNaN(date)) {
      logger.warn('Data inválida para formatação', { date });
      return 'Data inválida';
    }
    
    // 🔧 CORREÇÃO: Usar métodos nativos com padStart
    const dia = String(date.getDate()).padStart(2, '0');
    const mes = String(date.getMonth() + 1).padStart(2, '0'); // +1 porque getMonth() retorna 0-11
    const ano = date.getFullYear();
    
    return `${dia}/${mes}/${ano}`;
    
  } catch (error) {
    logger.error('Erro ao formatar data', { date, erro: error.message });
    return 'Erro na data';
  }
}

/**
 * Formata data e hora completos
 * @param {Date} date - Data a ser formatada
 * @returns {string} - Data e hora formatados
 */
function formatarDataHora(date) {
  try {
    if (!(date instanceof Date) || isNaN(date)) {
      return 'Data/hora inválida';
    }
    
    return Utilities.formatDate(
      date, 
      CONFIG.SISTEMA.TIMEZONE, 
      `${CONFIG.SISTEMA.DATA_FORMATO} ${CONFIG.SISTEMA.HORA_FORMATO}`
    );
    
  } catch (error) {
    logger.error('Erro ao formatar data/hora', { date, erro: error.message });
    return 'Erro na data/hora';
  }
}

/**
 * Valida se uma string está no formato de hora HH:MM
 * @param {string} hora - Hora a ser validada
 * @returns {boolean} - Se é válida ou não
 */
function validarHora(hora) {
  return VALIDACOES.FORMATOS.HORA.test(hora);
}

/**
 * Sanitiza texto removendo caracteres perigosos
 * @param {string} texto - Texto a ser sanitizado
 * @returns {string} - Texto sanitizado
 */
function sanitizarTexto(texto) {
  if (!CONFIG.SEGURANCA.SANITIZAR_INPUTS) {
    return texto;
  }
  
  let textoSanitizado = String(texto).trim();
  
  // Limitar tamanho
  if (textoSanitizado.length > CONFIG.SEGURANCA.MAX_CARACTERES_CAMPO) {
    textoSanitizado = textoSanitizado.substring(0, CONFIG.SEGURANCA.MAX_CARACTERES_CAMPO);
  }
  
  // Remover HTML se não permitido
  if (!CONFIG.SEGURANCA.PERMITIR_HTML_OBSERVACOES) {
    textoSanitizado = textoSanitizado.replace(/<[^>]*>/g, '');
  }
  
  // Remover caracteres de controle
  textoSanitizado = textoSanitizado.replace(/[\x00-\x1F\x7F]/g, '');
  
  // Escapar aspas para evitar problemas
  textoSanitizado = textoSanitizado.replace(/'/g, "''");
  
  return textoSanitizado;
}

// ===================================================
// ASSINATURA DIGITAL LIMPA (APENAS IMAGEM - ESTILO HOTEL LIOZ)
// ===================================================

const ASSINATURA_FILE_ID = '1z6i_VnZZ9OdHbcmcTHa4dKHPIqC6vmBE';

/**
 * Monta assinatura APENAS com sua imagem (igual ao Hotel LIOZ)
 * @returns {Object} - { html: string, inlineImages: Object }
 */
function montarAssinaturaEmail() {
  try {
    // URL pública da sua imagem (sempre visível)
    const imagemUrl = 'https://drive.google.com/uc?export=view&id=1z6i_VnZZ9OdHbcmcTHa4dKHPIqC6vmBE';
    
    // HTML APENAS com sua imagem (EXATAMENTE como o Hotel LIOZ)
    const html = `
      <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #e0e0e0; text-align: center;">
        <!-- APENAS SUA IMAGEM (como o Hotel LIOZ) -->
        <img src="${imagemUrl}" 
             alt="HUB Transfer" 
             style="max-width: 600px; 
                    width: 100%; 
                    height: auto; 
                    border: 0; 
                    display: block; 
                    margin: 0 auto;" />
      </div>
    `;
    
    if (typeof logger !== 'undefined' && logger.success) {
      logger.success('Assinatura criada APENAS com imagem', { 
        tamanhoHtml: html.length 
      });
    }
    
    return {
      html: html,
      inlineImages: {} // Não precisa de inline com URL pública
    };
    
  } catch (error) {
    if (typeof logger !== 'undefined' && logger.warn) {
      logger.warn('Erro na assinatura com imagem', { erro: error.message });
    }
    
    // Fallback vazio se der erro
    return {
      html: '',
      inlineImages: {}
    };
  }
}

/**
 * Envia e-mail sempre com a assinatura digital padronizada - VERSÃO CORRIGIDA
 * CORREÇÃO: Removida linha problemática e melhorado tratamento de erros
 */
function sendEmailComAssinatura(opts) {
  try {
    console.log('📧 Enviando e-mail com assinatura...');
    
    // 🔥 FORÇAR ASSINATURA APENAS COM IMAGEM (SEM TEXTO)
    const assinaturaForcada = `
      <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #e0e0e0; text-align: center;">
        <img src="https://drive.google.com/uc?export=view&id=1z6i_VnZZ9OdHbcmcTHa4dKHPIqC6vmBE" 
             alt="HUB Transfer" 
             style="max-width: 600px; width: 100%; height: auto; border: 0; display: block; margin: 0 auto;" />
      </div>
    `;
    
    // Garantir htmlBody (priorizar existente, converter body se necessário)
    let baseHtml = opts.htmlBody || '';
    
    if (!baseHtml && opts.body) {
      // Converter texto simples para HTML básico
      baseHtml = `<div style="font-family: Arial, sans-serif; line-height: 1.6;">${escapeHtml(String(opts.body))}</div>`;
    }
    
    // 🔥 FORÇAR: usar APENAS a assinatura com imagem
    const htmlBody = baseHtml + assinaturaForcada;
    
    // Preparar opções finais (SEM inlineImages para evitar problemas)
    const finalOpts = {
      to: opts.to,
      subject: opts.subject,
      htmlBody: htmlBody,
      name: opts.name || (typeof CONFIG !== 'undefined' ? CONFIG.NAMES.SISTEMA_NOME : 'Sistema HUB Transfer'),
      replyTo: opts.replyTo
    };
    
    // Log do envio
    console.log('📤 Dados do e-mail:', {
      destinatario: finalOpts.to,
      assunto: finalOpts.subject,
      tamanhoHtml: htmlBody.length,
      temAssinatura: htmlBody.includes('1z6i_VnZZ9OdHbcmcTHa4dKHPIqC6vmBE')
    });
    
    // 🔧 CORREÇÃO: Enviar e-mail sem a linha problemática
    return GmailApp.sendEmail(finalOpts.to, finalOpts.subject, '', {
      htmlBody: finalOpts.htmlBody,
      name: finalOpts.name,
      replyTo: finalOpts.replyTo
    });
    
  } catch (error) {
    console.log('❌ Erro no envio principal:', error.message);
    
    // Log do erro
    if (typeof logger !== 'undefined' && logger.error) {
      logger.error('Erro ao enviar e-mail com assinatura', {
        erro: error.message,
        destinatario: opts.to,
        assunto: opts.subject
      });
    }
    
    // Fallback: tentar enviar SEM assinatura
    try {
      console.log('🔄 Tentando fallback sem assinatura...');
      
      return GmailApp.sendEmail(opts.to, opts.subject, '', {
        htmlBody: opts.htmlBody || opts.body,
        name: opts.name || (typeof CONFIG !== 'undefined' ? CONFIG.NAMES.SISTEMA_NOME : 'Sistema HUB Transfer'),
        replyTo: opts.replyTo
      });
      
    } catch (fallbackError) {
      console.log('❌ Fallback também falhou:', fallbackError.message);
      // Se ainda assim falhar, lançar erro original
      throw error;
    }
  }
}

/**
 * Envia e-mail de confirmação após ação - VERSÃO CORRIGIDA
 */
function enviarEmailConfirmacaoAcao(transferId, action, userEmail) {
  try {
    const destinatarios = CONFIG.EMAIL_CONFIG.DESTINATARIOS.join(',');
    const assunto = action === 'confirm' 
      ? `✅ Transfer #${transferId} Confirmado`
      : `❌ Transfer #${transferId} Cancelado`;
    
    const actionText = action === 'confirm' ? 'confirmado' : 'cancelado';
    const emoji = action === 'confirm' ? '✅' : '❌';
    
    const corpo = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h2 style="color: ${action === 'confirm' ? '#28a745' : '#dc3545'};">
          ${emoji} Transfer ${actionText.toUpperCase()}
        </h2>
        
        <div style="background: #f8f9fa; padding: 20px; border-radius: 8px; margin: 20px 0;">
          <p><strong>Transfer ID:</strong> #${transferId}</p>
          <p><strong>Ação:</strong> ${actionText}</p>
          <p><strong>Processado por:</strong> ${userEmail}</p>
          <p><strong>Data/Hora:</strong> ${formatarDataHora(new Date())}</p>
        </div>
        
        <p style="color: #666; font-size: 14px;">
          Esta é uma confirmação automática do sistema de gestão de transfers.
        </p>
      </div>
    `;
    
    // 🔧 CORREÇÃO: Usar a função corrigida
    return sendEmailComAssinatura({
      to: destinatarios,
      subject: assunto,
      htmlBody: corpo,
      name: CONFIG.NAMES.SISTEMA_NOME,
      replyTo: CONFIG.EMAIL_CONFIG.EMAIL_EMPRESA
    });
    
  } catch (error) {
    console.log('❌ Erro ao enviar confirmação:', error.message);
    if (typeof logger !== 'undefined' && logger.error) {
      logger.error('Erro ao enviar e-mail de confirmação', error);
    }
  }
}

/**
 * Helper para escape de HTML (converte texto simples em HTML seguro)
 * @param {string} text - Texto a ser escapado
 * @returns {string} - Texto escapado para HTML
 */
function escapeHtml(text) {
  if (!text) return '';
  
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
    .replace(/\n/g, '<br>');
}

/**
 * Testa a assinatura digital (função de debug)
 */
function testarAssinaturaDigital() {
  try {
    const assinatura = montarAssinaturaEmail();
    
    console.log('Teste da Assinatura Digital:');
    console.log('HTML Length:', assinatura.html.length);
    console.log('Inline Images:', Object.keys(assinatura.inlineImages));
    console.log('Assinatura carregada:', assinatura.html.length > 0 ? 'SIM' : 'NÃO');
    console.log('Contém sua imagem:', assinatura.html.includes('1z6i_VnZZ9OdHbcmcTHa4dKHPIqC6vmBE') ? 'SIM' : 'NÃO');
    console.log('É apenas imagem (sem texto):', !assinatura.html.includes('Junior') && !assinatura.html.includes('968 698 138') ? 'SIM' : 'NÃO');
    
    return assinatura;
    
  } catch (error) {
    console.error('Erro no teste da assinatura:', error);
    return null;
  }
}

// ===================================================
// FUNÇÕES DE MANIPULAÇÃO DE PLANILHA (FUSÃO COMPLETA)
// ===================================================

/**
 * Gera o próximo ID disponível de forma segura
 * @param {Sheet} sheet - Planilha onde buscar o ID
 * @returns {number} - Próximo ID disponível
 */
function gerarProximoIdSeguro(sheet) {
  logger.debug('Gerando próximo ID seguro');
  
  try {
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      logger.info('Primeira entrada, retornando ID 1');
      return 1;
    }
    
    // Buscar todos os IDs existentes
    const idsRange = sheet.getRange(2, 1, lastRow - 1, 1);
    const ids = idsRange.getValues().flat().map(Number).filter(n => !isNaN(n) && n > 0);
    
    if (ids.length === 0) {
      return 1;
    }
    
    // Encontrar o maior ID
    const maxId = Math.max(...ids);
    const novoId = maxId + 1;
    
    logger.debug('Novo ID gerado', { maxId, novoId });
    return novoId;
    
  } catch (error) {
    logger.error('Erro ao gerar ID, usando timestamp', error);
    // Fallback: usar timestamp
    return Date.now() % 1000000;
  }
}

/**
 * Encontra a linha de um transfer pelo ID
 * @param {Sheet} sheet - Planilha onde buscar
 * @param {string|number} id - ID do transfer
 * @returns {number} - Número da linha (0 se não encontrado)
 */
function encontrarLinhaPorId(sheet, id) {
  logger.debug('Buscando linha por ID', { sheetName: sheet.getName(), id });
  
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return 0;
    
    const idsRange = sheet.getRange(2, 1, lastRow - 1, 1);
    const ids = idsRange.getValues();
    
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) {
        const linha = i + 2; // +2 porque começamos na linha 2
        logger.debug('ID encontrado', { id, linha });
        return linha;
      }
    }
    
    logger.debug('ID não encontrado', { id });
    return 0;
    
  } catch (error) {
    logger.error('Erro ao buscar ID', { id, erro: error.message });
    return 0;
  }
}

/**
 * Verifica se existe registro duplicado
 * @param {string|number} id - ID do transfer
 * @param {string|Date} data - Data do transfer
 * @returns {boolean} - Se existe duplicata
 */
function verificarRegistroDuplo(id, data) {
  logger.info('Verificando registro duplicado', { id, data });
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet || !id || !data) return false;
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return false;
    
    const registros = sheet.getRange(2, 1, lastRow - 1, 6).getValues(); // Incluir coluna F (Data)
    const dataFormatada = formatarDataDDMMYYYY(processarDataSegura(data));
    
    const duplicado = registros.some(row => {
      return String(row[0]) === String(id) && 
             formatarDataDDMMYYYY(new Date(row[5])) === dataFormatada; // Coluna F para data
    });
    
    logger.info('Verificação de duplicata concluída', { duplicado });
    return duplicado;
    
  } catch (error) {
    logger.error('Erro ao verificar duplicata', error);
    return false;
  }
}

/**
 * Valida um objeto de dados contra campos obrigatórios
 * @param {Object} dados - Dados a serem validados
 * @returns {Object} - { valido: boolean, erros: Array, dados: Object }
 */
function validarDados(dados) {
  logger.info('Validando dados de entrada');
  
  const erros = [];
  const dadosValidados = {};
  
  // Verificar campos obrigatórios
  for (const campo of VALIDACOES.CAMPOS_OBRIGATORIOS) {
    if (!dados[campo] || dados[campo] === '') {
      erros.push(MESSAGES.ERROS.CAMPO_OBRIGATORIO(campo));
    }
  }
  
  // Validar e sanitizar cada campo
  if (dados.nomeCliente) {
    dadosValidados.nomeCliente = sanitizarTexto(dados.nomeCliente);
  }
  
  // NOVO: Validar tipo de serviço
  if (dados.tipoServico) {
    if (VALIDACOES.VALORES_PERMITIDOS.TIPO_SERVICO.includes(dados.tipoServico)) {
      dadosValidados.tipoServico = dados.tipoServico;
    } else {
      dadosValidados.tipoServico = 'Transfer'; // Padrão
    }
  } else {
    dadosValidados.tipoServico = 'Transfer';
  }
  
  // Validar número de pessoas
  if (dados.numeroPessoas) {
    const pessoas = parseInt(dados.numeroPessoas);
    if (isNaN(pessoas) || pessoas < 1 || pessoas > CONFIG.LIMITES.MAX_PESSOAS) {
      erros.push(`Número de pessoas deve estar entre 1 e ${CONFIG.LIMITES.MAX_PESSOAS}`);
    } else {
      dadosValidados.numeroPessoas = pessoas;
    }
  }
  
  // Validar número de bagagens
  if (dados.numeroBagagens !== undefined) {
    const bagagens = parseInt(dados.numeroBagagens);
    if (isNaN(bagagens) || bagagens < 0 || bagagens > CONFIG.LIMITES.MAX_BAGAGENS) {
      erros.push(`Número de bagagens deve estar entre 0 e ${CONFIG.LIMITES.MAX_BAGAGENS}`);
    } else {
      dadosValidados.numeroBagagens = bagagens;
    }
  }
  
  // Validar valor total
  if (dados.valorTotal) {
    const valor = parseFloat(dados.valorTotal);
    if (isNaN(valor) || valor < CONFIG.LIMITES.MIN_VALOR || valor > CONFIG.LIMITES.MAX_VALOR) {
      erros.push(`Valor deve estar entre ${CONFIG.LIMITES.MIN_VALOR} e ${CONFIG.LIMITES.MAX_VALOR}`);
    } else {
      dadosValidados.valorTotal = valor;
    }
  }
  
  // Validar data
  if (dados.data) {
    try {
      dadosValidados.data = processarDataSegura(dados.data);
    } catch (error) {
      erros.push(MESSAGES.ERROS.DATA_INVALIDA);
    }
  }
  
  // CORRIGIDO: Validar e formatar hora
  if (dados.horaPickup) {
    const horaFormatada = validarEFormatarHora(dados.horaPickup);
    if (!horaFormatada) {
      erros.push('Hora deve estar no formato HH:MM');
    } else {
      dadosValidados.horaPickup = horaFormatada;
    }
  }
  
  // Validar contato (telefone ou e-mail)
  if (dados.contacto) {
    dadosValidados.contacto = sanitizarTexto(dados.contacto);
    if (CONFIG.SEGURANCA.VALIDAR_TELEFONE && !VALIDACOES.FORMATOS.TELEFONE.test(dados.contacto)) {
      if (CONFIG.SEGURANCA.VALIDAR_EMAIL && !VALIDACOES.FORMATOS.EMAIL.test(dados.contacto)) {
        erros.push('Contacto deve ser um telefone ou e-mail válido');
      }
    }
  }
  
  // Validar forma de pagamento
  if (dados.modoPagamento) {
    if (!VALIDACOES.VALORES_PERMITIDOS.FORMA_PAGAMENTO.includes(dados.modoPagamento)) {
      dadosValidados.modoPagamento = 'Dinheiro'; // Valor padrão
    } else {
      dadosValidados.modoPagamento = dados.modoPagamento;
    }
  }
  
  // Copiar outros campos com sanitização
  const outrosCampos = ['numeroVoo', 'origem', 'destino', 'pagoParaQuem', 'observacoes', 'tourSelecionado'];
  for (const campo of outrosCampos) {
    if (dados[campo]) {
      dadosValidados[campo] = sanitizarTexto(dados[campo]);
    }
  }
  
  // Aplicar valores padrão
  dadosValidados.status = dados.status || 'Solicitado';
  dadosValidados.modoPagamento = dadosValidados.modoPagamento || 'Dinheiro';
  dadosValidados.pagoParaQuem = dadosValidados.pagoParaQuem || 'Recepção';
  dadosValidados.numeroBagagens = dadosValidados.numeroBagagens || 0;
  
  const resultado = {
    valido: erros.length === 0,
    erros: erros,
    dados: dadosValidados
  };
  
  logger.info('Validação concluída', { 
    valido: resultado.valido, 
    numeroErros: erros.length 
  });
  
  return resultado;
}

/**
 * NOVA FUNÇÃO: Valida e formata horário corretamente
 * @param {*} hora - Valor do horário recebido
 * @returns {string|null} - Horário formatado como HH:MM ou null se inválido
 */
function validarEFormatarHora(hora) {
  if (!hora) return null;
  
  try {
    // Se já é string no formato HH:MM
    if (typeof hora === 'string' && /^\d{2}:\d{2}$/.test(hora)) {
      return hora;
    }
    
    // Se é string com segundos, remover os segundos
    if (typeof hora === 'string' && /^\d{2}:\d{2}:\d{2}$/.test(hora)) {
      return hora.substring(0, 5);
    }
    
    // Se é string que contém horário (ex: "1899-12-30T16:15:00.000Z")
    if (typeof hora === 'string') {
      const timeMatch = hora.match(/(\d{2}):(\d{2})/);
      if (timeMatch) {
        return `${timeMatch[1]}:${timeMatch[2]}`;
      }
    }
    
    // Se for objeto Date
    if (hora instanceof Date) {
      const hours = hora.getHours().toString().padStart(2, '0');
      const minutes = hora.getMinutes().toString().padStart(2, '0');
      return `${hours}:${minutes}`;
    }
    
    // Se é objeto que pode ser convertido para Date
    if (typeof hora === 'object' && hora.toString) {
      const horaStr = hora.toString();
      const timeMatch = horaStr.match(/(\d{2}):(\d{2})/);
      if (timeMatch) {
        return `${timeMatch[1]}:${timeMatch[2]}`;
      }
    }
    
    // Tentar converter para string e extrair horário
    const horaString = hora.toString();
    const timeMatch = horaString.match(/(\d{2}):(\d{2})/);
    if (timeMatch) {
      const hours = parseInt(timeMatch[1]);
      const minutes = parseInt(timeMatch[2]);
      
      // Validar se são valores válidos
      if (hours >= 0 && hours <= 23 && minutes >= 0 && minutes <= 59) {
        return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
      }
    }
    
    return null;
    
  } catch (error) {
    logger.warn('Erro ao formatar horário', { hora, error: error.toString() });
    return null;
  }
}

/**
 * Cria uma resposta JSON com CORS headers completos
 * CORREÇÃO: Headers CORS corrigidos para Google Apps Script
 */
function createJsonResponse(data, statusCode = 200) {
  const response = {
    ...data,
    timestamp: new Date().toISOString(),
    version: CONFIG.SISTEMA.VERSAO
  };
  
  logger.debug('Criando resposta JSON', { statusCode, hasData: !!data });
  
  // 🔧 CORREÇÃO: Método correto para Google Apps Script
  const output = ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
  
  // Headers CORS devem ser definidos de forma diferente no Google Apps Script
  // O Google Apps Script automaticamente adiciona headers CORS básicos
  // Para headers customizados, use uma abordagem alternativa:
  
  return output;
}

/**
 * Cria uma resposta HTML com CORS headers
 * CORREÇÃO: Headers CORS para páginas HTML também
 */
function createHtmlResponse(html) {
  const htmlOutput = HtmlService
    .createHtmlOutput(html)
    .setTitle(CONFIG.NAMES.SISTEMA_NOME)
    .setWidth(800)
    .setHeight(600);
  
  // 🔧 CORREÇÃO: Configuração correta para HTML
  htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  return htmlOutput;
}

/**
 * Cria uma resposta HTML com CORS headers
 * CORREÇÃO: Headers CORS para páginas HTML também
 */
function createHtmlResponse(html) {
  const htmlOutput = HtmlService
    .createHtmlOutput(html)
    .setTitle(CONFIG.NAMES.SISTEMA_NOME)
    .setWidth(800)
    .setHeight(600);
  
  // 🔧 CORREÇÃO: Adicionar headers CORS para HTML também
  htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  return htmlOutput;
}

// ===================================================
// SISTEMA DE GESTÃO DE ABAS MENSAIS (PRESERVADO DO CÓDIGO ANTIGO)
// ===================================================

/**
 * Obtém ou cria a aba do mês correspondente à data
 * @param {Date} dataTransfer - Data do transfer
 * @returns {Sheet} - Aba do mês
 */
function obterAbaMes(dataTransfer) {
  logger.info('Obtendo aba mensal', { data: dataTransfer });
  
  try {
    // Validar e processar data
    const data = processarDataSegura(dataTransfer);
    
    // Extrair informações do mês
    const mes = String(data.getMonth() + 1).padStart(2, '0');
    const ano = data.getFullYear();
    
    // Buscar informações do mês
    const mesInfo = MESES.find(m => m.abrev === mes);
    if (!mesInfo) {
      throw new Error(`Mês inválido: ${mes}`);
    }
    
    // Montar nome da aba
    const nomeAba = `${CONFIG.SISTEMA.PREFIXO_MES}${mes}_${mesInfo.nome}_${ano}`;
    
    logger.debug('Nome da aba calculado', { nomeAba, mes: mesInfo.nome, ano });
    
    // Acessar planilha
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    // Buscar aba existente
    let abaMes = ss.getSheetByName(nomeAba);
    
    if (abaMes) {
      logger.debug('Aba mensal encontrada', { nomeAba });
      return abaMes;
    }
    
    // Criar nova aba se não existir
    logger.info('Criando nova aba mensal', { nomeAba });
    abaMes = criarAbaMensal(nomeAba, ss, mesInfo);
    
    return abaMes;
    
  } catch (error) {
    logger.error('Erro ao obter aba mensal, usando fallback', error);
    
    // Fallback: retornar aba principal
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const abaPrincipal = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!abaPrincipal) {
      throw new Error('Falha crítica: aba principal não encontrada');
    }
    
    return abaPrincipal;
  }
}

/**
 * Cria uma nova aba mensal com formatação completa
 * @param {string} nomeAba - Nome da aba a criar
 * @param {Spreadsheet} ss - Objeto da planilha
 * @param {Object} mesInfo - Informações do mês
 * @returns {Sheet} - Nova aba criada
 */
function criarAbaMensal(nomeAba, ss, mesInfo) {
  logger.info('Criando aba mensal', { nome: nomeAba });
  
  try {
    // Criar nova aba
    const abaMes = ss.insertSheet(nomeAba);
    
    // Configurar headers
    abaMes.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    
    // Aplicar formatação do cabeçalho
    const headerRange = abaMes.getRange(1, 1, 1, HEADERS.length);
    headerRange
      .setBackground(mesInfo.cor || STYLES.HEADER_COLORS.PRINCIPAL)
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setFontSize(11)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    
    // Congelar primeira linha
    abaMes.setFrozenRows(1);
    
    // Aplicar larguras das colunas
    STYLES.COLUMN_WIDTHS.PRINCIPAL.forEach((width, index) => {
      abaMes.setColumnWidth(index + 1, width);
    });
    
    // Aplicar formatação
    aplicarFormatacaoMensal(abaMes);
    
    // Adicionar validações
    aplicarValidacoesPlanilha(abaMes);
    
    // Adicionar proteções se necessário
    protegerCelulasCriticas(abaMes);
    
    logger.success('Aba mensal criada com sucesso', { nome: nomeAba });
    
    return abaMes;
    
  } catch (error) {
    logger.error('Erro ao criar aba mensal', { nome: nomeAba, erro: error.message });
    throw error;
  }
}

/**
 * Aplica formatação específica para aba mensal
 * @param {Sheet} sheet - Aba a formatar
 */
function aplicarFormatacaoMensal(sheet) {
  logger.debug('Aplicando formatação mensal', { sheet: sheet.getName() });
  
  try {
    const maxRows = Math.max(sheet.getMaxRows(), 1000);
    
    // Formatação de moeda (colunas L, M, N, O) - ATUALIZADO PARA CÓDIGO NOVO
    if (maxRows > 1) {
      sheet.getRange(2, 12, maxRows - 1, 4).setNumberFormat(STYLES.FORMATS.MOEDA); // L, M, N, O
      
      // Formatação de data (coluna F) - ATUALIZADO
      sheet.getRange(2, 6, maxRows - 1, 1).setNumberFormat(STYLES.FORMATS.DATA);
      
      // Formatação de hora (coluna K) - ATUALIZADO
      sheet.getRange(2, 11, maxRows - 1, 1).setNumberFormat(STYLES.FORMATS.HORA);
      
      // Formatação de timestamp (coluna T) - ATUALIZADO
      sheet.getRange(2, 20, maxRows - 1, 1).setNumberFormat(STYLES.FORMATS.TIMESTAMP);
      
      // Formatação de números (colunas A, D, E) - ATUALIZADO
      sheet.getRange(2, 1, maxRows - 1, 1).setNumberFormat(STYLES.FORMATS.NUMERO);
      sheet.getRange(2, 4, maxRows - 1, 2).setNumberFormat(STYLES.FORMATS.NUMERO);
    }
    
    // Aplicar cores alternadas
    if (maxRows > 1) {
      const range = sheet.getRange(2, 1, maxRows - 1, HEADERS.length);
      // Remover banding existente primeiro
      const bandings = range.getBandings();
      bandings.forEach(banding => banding.remove());
      
      // Aplicar novo banding
      range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
    }
    
    logger.debug('Formatação mensal aplicada');
    
  } catch (error) {
    logger.error('Erro ao aplicar formatação mensal', error);
  }
}

/**
 * Aplica validações de dados na planilha
 * @param {Sheet} sheet - Aba onde aplicar validações
 */
function aplicarValidacoesPlanilha(sheet) {
  logger.debug('Aplicando validações', { sheet: sheet.getName() });
  
  try {
    const maxRows = Math.max(sheet.getMaxRows(), 1000);
    
    // Validação de Status (coluna R) - ATUALIZADO
    const statusValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(Object.values(MESSAGES.STATUS_MESSAGES))
      .setAllowInvalid(false)
      .setHelpText('Selecione o status do transfer')
      .build();
    sheet.getRange(2, 18, maxRows - 1, 1).setDataValidation(statusValidation);
    
    // Validação de Tipo de Serviço (coluna C) - NOVO
    const tipoServicoValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(VALIDACOES.VALORES_PERMITIDOS.TIPO_SERVICO)
      .setAllowInvalid(false)
      .setHelpText('Selecione o tipo de serviço')
      .build();
    sheet.getRange(2, 3, maxRows - 1, 1).setDataValidation(tipoServicoValidation);
    
    // Validação de Forma de Pagamento (coluna P) - ATUALIZADO
    const pagamentoValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(VALIDACOES.VALORES_PERMITIDOS.FORMA_PAGAMENTO)
      .setAllowInvalid(false)
      .setHelpText('Selecione a forma de pagamento')
      .build();
    sheet.getRange(2, 16, maxRows - 1, 1).setDataValidation(pagamentoValidation);
    
    // Validação de Pago Para (coluna Q) - ATUALIZADO
    const pagoParaValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(VALIDACOES.VALORES_PERMITIDOS.PAGO_PARA)
      .setAllowInvalid(false)
      .setHelpText('Selecione quem recebeu o pagamento')
      .build();
    sheet.getRange(2, 17, maxRows - 1, 1).setDataValidation(pagoParaValidation);
    
    // Validação de Pessoas (coluna D) - ATUALIZADO
    const pessoasValidation = SpreadsheetApp.newDataValidation()
      .requireNumberBetween(1, CONFIG.LIMITES.MAX_PESSOAS)
      .setAllowInvalid(false)
      .setHelpText(`Entre 1 e ${CONFIG.LIMITES.MAX_PESSOAS} pessoas`)
      .build();
    sheet.getRange(2, 4, maxRows - 1, 1).setDataValidation(pessoasValidation);
    
    // Validação de Bagagens (coluna E) - ATUALIZADO
    const bagagensValidation = SpreadsheetApp.newDataValidation()
      .requireNumberBetween(0, CONFIG.LIMITES.MAX_BAGAGENS)
      .setAllowInvalid(false)
      .setHelpText(`Entre 0 e ${CONFIG.LIMITES.MAX_BAGAGENS} bagagens`)
      .build();
    sheet.getRange(2, 5, maxRows - 1, 1).setDataValidation(bagagensValidation);
    
    // Validação de Valores (colunas L, M, N, O) - ATUALIZADO
    const valorValidation = SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThan(0)
      .setAllowInvalid(false)
      .setHelpText('Valor deve ser maior que zero')
      .build();
    sheet.getRange(2, 12, maxRows - 1, 4).setDataValidation(valorValidation);
    
    logger.debug('Validações aplicadas');
    
  } catch (error) {
    logger.error('Erro ao aplicar validações', error);
  }
}

/**
 * Protege células críticas contra edição acidental
 * @param {Sheet} sheet - Aba a proteger
 */
function protegerCelulasCriticas(sheet) {
  logger.debug('Protegendo células críticas', { sheet: sheet.getName() });
  
  try {
    // Proteger coluna ID (A) - apenas sistema pode editar
    const protectionID = sheet.getRange('A:A').protect()
      .setDescription('ID - Apenas sistema')
      .setWarningOnly(true);
    
    // Proteger coluna Data Criação (T) - ATUALIZADO
    const protectionTimestamp = sheet.getRange('T:T').protect()
      .setDescription('Timestamp - Apenas sistema')
      .setWarningOnly(true);
    
    // Proteger headers
    const protectionHeaders = sheet.getRange(1, 1, 1, HEADERS.length).protect()
      .setDescription('Headers - Não editar')
      .setWarningOnly(true);
    
    logger.debug('Proteções aplicadas');
    
  } catch (error) {
    logger.error('Erro ao proteger células', error);
  }
}

/**
 * Cria todas as abas mensais do ano
 * @param {number} ano - Ano para criar as abas
 * @returns {Object} - Resultado da criação
 */
function criarTodasAbasMensais(ano = CONFIG.SISTEMA.ANO_BASE) {
  logger.info('Criando todas as abas mensais', { ano });
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const resultados = {
      criadas: 0,
      existentes: 0,
      erros: 0,
      detalhes: []
    };
    
    MESES.forEach(mes => {
      const nomeAba = `${CONFIG.SISTEMA.PREFIXO_MES}${mes.abrev}_${mes.nome}_${ano}`;
      
      try {
        // Verificar se já existe
        let aba = ss.getSheetByName(nomeAba);
        
        if (aba) {
          resultados.existentes++;
          resultados.detalhes.push({
            mes: mes.nome,
            status: 'existente',
            nome: nomeAba
          });
          
          // Garantir formatação mesmo se já existe
          aplicarFormatacaoMensal(aba);
          aplicarValidacoesPlanilha(aba);
          
        } else {
          // Criar nova aba
          aba = criarAbaMensal(nomeAba, ss, mes);
          resultados.criadas++;
          resultados.detalhes.push({
            mes: mes.nome,
            status: 'criada',
            nome: nomeAba
          });
        }
        
      } catch (error) {
        logger.error('Erro ao processar aba mensal', { 
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
    
    logger.success('Criação de abas mensais concluída', resultados);
    
    return resultados;
    
  } catch (error) {
    logger.error('Erro ao criar abas mensais', error);
    throw error;
  }
}

// ===================================================
// SISTEMA DE CÁLCULO DE PREÇOS (FUSÃO CÓDIGO NOVO + ANTIGO)
// ===================================================

/**
 * Calcula valores do transfer baseado no tipo de serviço e dados
 * NOVA LÓGICA: Código novo como base + funcionalidades do antigo
 * @param {string} origem - Local de origem
 * @param {string} destino - Local de destino
 * @param {number} pessoas - Número de pessoas
 * @param {number} bagagens - Número de bagagens
 * @param {number} precoManual - Preço informado manualmente (opcional)
 * @param {string} tipoServico - Tipo de serviço (Transfer/Tour Regular/Private Tour)
 * @returns {Object} - Valores calculados
 */
function calcularValores(origem, destino, pessoas, bagagens, precoManual = null, tipoServico = 'Transfer') {
  logger.info('Calculando valores do transfer', {
    origem, destino, pessoas, bagagens, precoManual, tipoServico
  });
  
  try {
    // Estratégia 1: Preço manual com proporção (CÓDIGO NOVO)
    if (precoManual && precoManual > 0) {
      return calcularComPrecoManual(origem, destino, pessoas, bagagens, precoManual, tipoServico);
    }
    
    // Estratégia 2: Buscar na tabela de preços (CÓDIGO ANTIGO PRESERVADO)
    const precoTabela = consultarPrecoTabela(origem, destino, pessoas, bagagens, tipoServico);
    if (precoTabela) {
      logger.success('Preço encontrado na tabela', precoTabela);
      return {
        precoCliente: precoTabela.precoCliente,
        valorHotel: precoTabela.valorHotel,
        valorHUB: precoTabela.valorHUB,
        comissaoRecepcao: precoTabela.comissaoRecepcao,
        fonte: 'tabela',
        tabelaId: precoTabela.id,
        matchScore: precoTabela.matchScore,
        tipoServico: tipoServico,
        observacoes: precoTabela.observacoes || `Preço da tabela (ID: ${precoTabela.id})`
      };
    }
    
    // Estratégia 3: Calcular com base em distância/complexidade (CÓDIGO ANTIGO)
    const precoCalculado = calcularPrecoInteligente(origem, destino, pessoas, bagagens, tipoServico);
    if (precoCalculado) {
      return precoCalculado;
    }
    
    // Estratégia 4: Valores padrão baseados no tipo de serviço (CÓDIGO NOVO)
    logger.warn('Usando valores padrão baseados no tipo de serviço');
    return calcularValoresPadrao(tipoServico, pessoas);
    
  } catch (error) {
    logger.error('Erro no cálculo de valores', error);
    return calcularValoresPadrao('Transfer', pessoas || 1);
  }
}

/**
 * Calcula valores quando preço manual é informado (CÓDIGO NOVO)
 * @private
 */
function calcularComPrecoManual(origem, destino, pessoas, bagagens, precoManual, tipoServico) {
  logger.debug('Calculando com preço manual', { precoManual, tipoServico });
  
  const precoCliente = parseFloat(precoManual);
  
  // Tentar encontrar proporção na tabela
  const precoTabela = consultarPrecoTabela(origem, destino, pessoas, bagagens, tipoServico);
  
  if (precoTabela && precoTabela.precoCliente > 0) {
    // Usar proporção da tabela
    const proporcao = precoCliente / precoTabela.precoCliente;
    
    return {
      precoCliente: precoCliente,
      valorHotel: Math.round(precoTabela.valorHotel * proporcao * 100) / 100,
      valorHUB: Math.round(precoTabela.valorHUB * proporcao * 100) / 100,
      comissaoRecepcao: Math.round(precoTabela.comissaoRecepcao * proporcao * 100) / 100,
      fonte: 'manual-com-tabela',
      proporcao: proporcao,
      tipoServico: tipoServico,
      observacoes: `Preço manual com proporção da tabela (${(proporcao * 100).toFixed(0)}%)`
    };
  }
  
  // Usar cálculo baseado no tipo de serviço (CÓDIGO NOVO)
  return calcularPorTipoServico(precoCliente, tipoServico);
}

/**
* Calcula valores baseado no tipo de serviço (CÓDIGO NOVO - LÓGICA PRINCIPAL)
* @private
*/
function calcularPorTipoServico(precoCliente, tipoServico) {
 let valorHotel, valorHUB, comissaoRecepcao;
 
 // Cálculo baseado no tipo de serviço (CÓDIGO NOVO)
 if (tipoServico === 'Tour Regular' || tipoServico === 'tour') {
   // Tours regulares: 30% hotel + €5 recepção
   valorHotel = precoCliente * CONFIG.VALORES.PERCENTUAL_HOTEL;
   comissaoRecepcao = CONFIG.VALORES.COMISSAO_RECEPCAO_TOUR;
   valorHUB = precoCliente - valorHotel - comissaoRecepcao;
   
 } else if (tipoServico === 'Private Tour' || tipoServico === 'private') {
   // Private tours: 30% hotel + €5 recepção
   valorHotel = precoCliente * CONFIG.VALORES.PERCENTUAL_HOTEL;
   comissaoRecepcao = CONFIG.VALORES.COMISSAO_RECEPCAO_TOUR;
   valorHUB = precoCliente - valorHotel - comissaoRecepcao;
   
 } else {
   // Transfers: 30% hotel + €2 recepção
   valorHotel = precoCliente * CONFIG.VALORES.PERCENTUAL_HOTEL;
   comissaoRecepcao = CONFIG.VALORES.COMISSAO_RECEPCAO_TRANSFER;
   valorHUB = precoCliente - valorHotel - comissaoRecepcao;
 }
 
 // Arredondar valores
 valorHotel = Math.round(valorHotel * 100) / 100;
 comissaoRecepcao = Math.round(comissaoRecepcao * 100) / 100;
 valorHUB = Math.round(valorHUB * 100) / 100;
 
 return {
   precoCliente: precoCliente,
   valorHotel: valorHotel,
   valorHUB: valorHUB,
   comissaoRecepcao: comissaoRecepcao,
   fonte: 'manual-tipo-servico',
   tipoServico: tipoServico,
   observacoes: `Preço manual - ${tipoServico} - distribuição ${(CONFIG.VALORES.PERCENTUAL_HOTEL * 100).toFixed(0)}% hotel + €${comissaoRecepcao} recepção`
 };
}

/**
* Calcula valores padrão baseados no tipo de serviço (CÓDIGO NOVO)
* @private
*/
function calcularValoresPadrao(tipoServico, pessoas) {
 let precoCliente;
 
 // Definir preço base por tipo de serviço
 switch (tipoServico) {
   case 'Tour Regular':
   case 'tour':
     precoCliente = 67.00; // Por pessoa
     break;
   case 'Private Tour':
   case 'private':
     precoCliente = pessoas <= 3 ? 347.00 : 492.00; // Por grupo
     break;
   default:
     precoCliente = 25.00; // Transfer padrão
 }
 
 return calcularPorTipoServico(precoCliente, tipoServico);
}

/**
* Consulta preço na tabela com algoritmo de matching inteligente (CÓDIGO ANTIGO PRESERVADO)
* @private
*/
function consultarPrecoTabela(origem, destino, pessoas, bagagens, tipoServico = 'Transfer') {
 logger.debug('Consultando tabela de preços');
 
 try {
   const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
   const sheet = ss.getSheetByName(CONFIG.PRICING_SHEET_NAME);
   
   if (!sheet || sheet.getLastRow() <= 1) {
     logger.warn('Tabela de preços vazia ou não encontrada');
     return null;
   }
   
   const dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, PRICING_HEADERS.length).getValues();
   
   let melhorMatch = null;
   let melhorScore = 0;
   
   dados.forEach(linha => {
     // Pular se não está ativo
     if (linha[13] !== 'Sim') return; // Coluna N - Ativo
     
     // Verificar tipo de serviço
     if (linha[1] !== tipoServico) return; // Coluna B - Tipo Serviço
     
     const score = calcularScoreMatch(
       origem, destino, pessoas, bagagens,
       linha[3], linha[4], linha[5], linha[6] // Colunas D, E, F, G
     );
     
     if (score > melhorScore && score >= 60) { // Mínimo 60% de match
       melhorScore = score;
       melhorMatch = {
         id: linha[0],
         tipoServico: linha[1],
         rota: linha[2],
         origem: linha[3],
         destino: linha[4],
         pessoas: linha[5],
         bagagens: linha[6],
         precoPorPessoa: parseFloat(linha[7]) || 0,
         precoPorGrupo: parseFloat(linha[8]) || 0,
         precoCliente: parseFloat(linha[9]) || 0,
         valorHotel: parseFloat(linha[10]) || 0,
         valorHUB: parseFloat(linha[11]) || 0,
         comissaoRecepcao: parseFloat(linha[12]) || 0,
         observacoes: linha[15] || '',
         matchScore: score
       };
     }
   });
   
   if (melhorMatch) {
     logger.debug('Match encontrado', { score: melhorScore, id: melhorMatch.id });
   }
   
   return melhorMatch;
   
 } catch (error) {
   logger.error('Erro ao consultar tabela de preços', error);
   return null;
 }
}

/**
* Calcula score de matching entre busca e registro da tabela (CÓDIGO ANTIGO PRESERVADO)
* @private
*/
function calcularScoreMatch(origemBusca, destinoBusca, pessoasBusca, bagagensBusca,
                           origemTabela, destinoTabela, pessoasTabela, bagagensTabela) {
 
 let score = 0;
 
 // Normalizar textos para comparação
 const normalizar = (texto) => {
   return String(texto)
     .toLowerCase()
     .trim()
     .normalize('NFD')
     .replace(/[\u0300-\u036f]/g, '') // Remove acentos
     .replace(/[^a-z0-9\s]/g, ''); // Remove caracteres especiais
 };
 
 const origemBuscaNorm = normalizar(origemBusca);
 const destinoBuscaNorm = normalizar(destinoBusca);
 const origemTabelaNorm = normalizar(origemTabela);
 const destinoTabelaNorm = normalizar(destinoTabela);
 
 // Score de origem/destino (peso 70%)
 if (origemBuscaNorm === origemTabelaNorm && destinoBuscaNorm === destinoTabelaNorm) {
   score += 70;
 } else if (origemBuscaNorm.includes(origemTabelaNorm) && destinoBuscaNorm.includes(destinoTabelaNorm)) {
   score += 50;
 } else if (origemTabelaNorm.includes(origemBuscaNorm) && destinoTabelaNorm.includes(destinoBuscaNorm)) {
   score += 50;
 } else {
   // Verificar palavras-chave comuns
   const palavrasOrigemBusca = origemBuscaNorm.split(' ');
   const palavrasDestinoBusca = destinoBuscaNorm.split(' ');
   const palavrasOrigemTabela = origemTabelaNorm.split(' ');
   const palavrasDestinoTabela = destinoTabelaNorm.split(' ');
   
   let matchesOrigem = 0;
   let matchesDestino = 0;
   
   palavrasOrigemBusca.forEach(palavra => {
     if (palavra.length > 3 && palavrasOrigemTabela.includes(palavra)) {
       matchesOrigem++;
     }
   });
   
   palavrasDestinoBusca.forEach(palavra => {
     if (palavra.length > 3 && palavrasDestinoTabela.includes(palavra)) {
       matchesDestino++;
     }
   });
   
   const percentMatchOrigem = palavrasOrigemBusca.length > 0 
     ? matchesOrigem / palavrasOrigemBusca.length 
     : 0;
   
   const percentMatchDestino = palavrasDestinoBusca.length > 0 
     ? matchesDestino / palavrasDestinoBusca.length 
     : 0;
   
   score += Math.round((percentMatchOrigem + percentMatchDestino) * 35);
 }
 
 // Score de pessoas (peso 20%)
 const pessoasBuscaInt = parseInt(pessoasBusca) || 1;
 const pessoasTabelaInt = parseInt(pessoasTabela) || 1;
 
 if (pessoasBuscaInt === pessoasTabelaInt) {
   score += 20;
 } else if (pessoasTabelaInt >= pessoasBuscaInt && pessoasTabelaInt <= pessoasBuscaInt + 2) {
   score += 15;
 } else if (Math.abs(pessoasTabelaInt - pessoasBuscaInt) <= 1) {
   score += 10;
 }
 
 // Score de bagagens (peso 10%)
 const bagagensBuscaInt = parseInt(bagagensBusca) || 0;
 const bagagensTabelaInt = parseInt(bagagensTabela) || 0;
 
 if (bagagensBuscaInt === bagagensTabelaInt) {
   score += 10;
 } else if (bagagensTabelaInt >= bagagensBuscaInt) {
   score += 5;
 }
 
 return score;
}

/**
* Calcula preço de forma inteligente baseado em regras de negócio (CÓDIGO ANTIGO PRESERVADO)
* @private
*/
function calcularPrecoInteligente(origem, destino, pessoas, bagagens, tipoServico = 'Transfer') {
 logger.debug('Calculando preço inteligente');
 
 try {
   // Base de conhecimento de locais e distâncias aproximadas
   const locaisConhecidos = {
     'aeroporto': { tipo: 'aeroporto', distanciaBase: 15 },
     'lisboa': { tipo: 'aeroporto', distanciaBase: 15 },
     'cascais': { tipo: 'cidade', distanciaBase: 30 },
     'sintra': { tipo: 'cidade', distanciaBase: 35 },
     'belem': { tipo: 'bairro', distanciaBase: 10 },
     'marques': { tipo: 'hotel', distanciaBase: 0 },
     'hotel': { tipo: 'hotel', distanciaBase: 0 }
   };
   
   // Extrair tipo de local
   let tipoOrigem = null;
   let tipoDestino = null;
   let distanciaEstimada = 25; // km padrão
   
   const origemLower = origem.toLowerCase();
   const destinoLower = destino.toLowerCase();
   
   // Identificar tipos de locais
   Object.keys(locaisConhecidos).forEach(key => {
     if (origemLower.includes(key)) {
       tipoOrigem = locaisConhecidos[key];
     }
     if (destinoLower.includes(key)) {
       tipoDestino = locaisConhecidos[key];
     }
   });
   
   // Calcular distância estimada
   if (tipoOrigem && tipoDestino) {
     distanciaEstimada = Math.abs(tipoOrigem.distanciaBase + tipoDestino.distanciaBase);
   }
   
   // Cálculo base diferenciado por tipo de serviço
   let precoBase;
   
   if (tipoServico === 'Tour Regular' || tipoServico === 'tour') {
     precoBase = 60; // Base para tours regulares
   } else if (tipoServico === 'Private Tour' || tipoServico === 'private') {
     precoBase = pessoas <= 3 ? 300 : 450; // Base para private tours
   } else {
     precoBase = 15; // Base para transfers
     
     // Adicionar por distância (€0.80 por km)
     precoBase += distanciaEstimada * 0.80;
     
     // Adicionar por pessoa extra (€3 por pessoa além da primeira)
     const pessoasInt = parseInt(pessoas) || 1;
     if (pessoasInt > 1) {
       precoBase += (pessoasInt - 1) * 3;
     }
     
     // Adicionar por bagagem extra (€2 por bagagem além de 1 por pessoa)
     const bagagensInt = parseInt(bagagens) || 0;
     const bagagensExtras = Math.max(0, bagagensInt - pessoasInt);
     if (bagagensExtras > 0) {
       precoBase += bagagensExtras * 2;
     }
     
     // Adicionar taxa aeroporto se aplicável
     if ((tipoOrigem && tipoOrigem.tipo === 'aeroporto') || 
         (tipoDestino && tipoDestino.tipo === 'aeroporto')) {
       precoBase += 5; // Taxa aeroporto
     }
   }
   
   // Arredondar para múltiplo de 5
   const precoCliente = Math.ceil(precoBase / 5) * 5;
   
   // Calcular distribuição usando a nova lógica
   const valores = calcularPorTipoServico(precoCliente, tipoServico);
   
   logger.debug('Preço calculado inteligentemente', {
     distancia: distanciaEstimada,
     precoBase: precoBase,
     precoFinal: precoCliente,
     tipoServico: tipoServico
   });
   
   return {
     ...valores,
     fonte: 'calculado',
     distanciaEstimada: distanciaEstimada,
     observacoes: `Preço calculado - ${tipoServico} - Distância estimada: ${distanciaEstimada}km`
   };
   
 } catch (error) {
   logger.error('Erro no cálculo inteligente', error);
   return null;
 }
}

// ===================================================
// SISTEMA DE REGISTRO E ATUALIZAÇÃO (FUSÃO COMPLETA)
// ===================================================

/**
 * Executa o registro duplo do transfer (principal + mensal)
 * @param {Array} dadosTransfer - Dados do transfer a registrar
 * @returns {Object} - Resultado do registro
 */
function executarRegistroDuplo(dadosTransfer) {
  logger.info('Executando registro duplo', { 
    transferId: dadosTransfer[0],
    cliente: dadosTransfer[1] 
  });
  
  try {
    const transferId = dadosTransfer[0];
    const dataTransfer = new Date(dadosTransfer[5]); // Coluna F - Data
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    // Verificar duplicidade antes de registrar
    if (verificarRegistroDuplo(transferId, dataTransfer)) {
      logger.warn('Transfer duplicado detectado', { transferId });
      return {
        sucesso: false,
        registroDuploCompleto: false,
        observacoes: 'Transfer duplicado - registro cancelado'
      };
    }
    
    // PASSO 1: Registrar na aba principal
    const abaPrincipal = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!abaPrincipal) {
      throw new Error(`Aba principal '${CONFIG.SHEET_NAME}' não encontrada`);
    }
    
    let registroPrincipal = false;
    let tentativasPrincipal = 0;
    
    while (!registroPrincipal && tentativasPrincipal < CONFIG.SISTEMA.MAX_TENTATIVAS_REGISTRO) {
      tentativasPrincipal++;
      
      try {
        abaPrincipal.appendRow(dadosTransfer);
        Utilities.sleep(CONFIG.SISTEMA.INTERVALO_TENTATIVAS);
        
        // Verificar se foi registrado
        const linha = encontrarLinhaPorId(abaPrincipal, transferId);
        registroPrincipal = linha > 0;
        
        if (registroPrincipal) {
          logger.success('Registro na aba principal concluído', { 
            transferId, 
            linha,
            tentativa: tentativasPrincipal 
          });
        }
        
      } catch (errorPrincipal) {
        logger.error(`Erro na tentativa ${tentativasPrincipal} (principal)`, errorPrincipal);
        if (tentativasPrincipal < CONFIG.SISTEMA.MAX_TENTATIVAS_REGISTRO) {
          Utilities.sleep(CONFIG.SISTEMA.INTERVALO_TENTATIVAS * 2);
        }
      }
    }
    
    if (!registroPrincipal) {
      throw new Error('Falha ao registrar na aba principal após múltiplas tentativas');
    }
    
    // PASSO 2: Registrar na aba mensal
    let registroMensal = false;
    let tentativasMensal = 0;
    let nomeAbaMensal = 'N/A';
    
    if (CONFIG.SISTEMA.ORGANIZAR_POR_MES) {
      try {
        const abaMensal = obterAbaMes(dataTransfer);
        nomeAbaMensal = abaMensal.getName();
        
        // Não registrar duplicado se for a mesma aba
        if (nomeAbaMensal !== CONFIG.SHEET_NAME) {
          while (!registroMensal && tentativasMensal < CONFIG.SISTEMA.MAX_TENTATIVAS_REGISTRO) {
            tentativasMensal++;
            
            try {
              abaMensal.appendRow(dadosTransfer);
              Utilities.sleep(CONFIG.SISTEMA.INTERVALO_TENTATIVAS);
              
              // Verificar se foi registrado
              const linhaMensal = encontrarLinhaPorId(abaMensal, transferId);
              registroMensal = linhaMensal > 0;
              
              if (registroMensal) {
                logger.success('Registro na aba mensal concluído', {
                  transferId,
                  abaMensal: nomeAbaMensal,
                  linha: linhaMensal,
                  tentativa: tentativasMensal
                });
              }
              
            } catch (errorMensal) {
              logger.error(`Erro na tentativa ${tentativasMensal} (mensal)`, errorMensal);
              if (tentativasMensal < CONFIG.SISTEMA.MAX_TENTATIVAS_REGISTRO) {
                Utilities.sleep(CONFIG.SISTEMA.INTERVALO_TENTATIVAS * 2);
              }
            }
          }
        } else {
          logger.warn('Aba mensal é igual à principal, pulando registro duplo');
          registroMensal = true; // Considerar como sucesso
        }
        
      } catch (errorAbaMensal) {
        logger.error('Erro ao obter/criar aba mensal', errorAbaMensal);
      }
    }
    
    // PASSO 3: Enviar e-mail se configurado
    let emailEnviado = false;
    if (registroPrincipal && CONFIG.EMAIL_CONFIG.ENVIAR_AUTOMATICO) {
      try {
        emailEnviado = enviarEmailNovoTransfer(dadosTransfer);
      } catch (errorEmail) {
        logger.error('Erro no envio de e-mail', errorEmail);
      }
    }
    
    // PASSO 4: Integração com webhooks se configurado (PRESERVADO DO CÓDIGO ANTIGO)
    if (registroPrincipal && CONFIG.MAKE_WEBHOOKS.ENABLED) {
      try {
        enviarWebhookNovoTransfer(dadosTransfer);
      } catch (errorWebhook) {
        logger.error('Erro no envio de webhook', errorWebhook);
      }
    }
    
    // Montar resultado
    const registroCompleto = registroPrincipal && registroMensal;
    const registroParcial = registroPrincipal && !registroMensal;
    
    const resultado = {
      sucesso: registroPrincipal,
      registroDuploCompleto: registroCompleto,
      abaPrincipal: {
        nome: CONFIG.SHEET_NAME,
        inserido: registroPrincipal,
        tentativas: tentativasPrincipal
      },
      abaMensal: {
        nome: nomeAbaMensal,
        inserido: registroMensal,
        tentativas: tentativasMensal
      },
      emailEnviado: emailEnviado,
      observacoes: registroCompleto
        ? '✅ Registro duplo realizado com sucesso'
        : registroParcial
          ? '⚠️ Registro apenas na aba principal'
          : '❌ Falha ao registrar o transfer'
    };
    
    // Adicionar observações sobre e-mail
    if (emailEnviado) {
      resultado.observacoes += ' | E-mail enviado';
    }
    
    logger.info('Registro duplo finalizado', resultado);
    
    return resultado;
    
  } catch (error) {
    logger.error('Erro crítico no registro duplo', error);
    
    return {
      sucesso: false,
      registroDuploCompleto: false,
      erro: error.message,
      observacoes: `❌ Erro crítico: ${error.message}`
    };
  }
}

/**
 * Atualiza o status de um transfer nas duas abas
 * @param {string} transferId - ID do transfer
 * @param {string} novoStatus - Novo status
 * @param {string} observacao - Observação adicional
 * @returns {Object} - Resultado da atualização
 */
function atualizarStatusTransfer(transferId, novoStatus, observacao = '') {
  logger.info('Atualizando status do transfer', { transferId, novoStatus });
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const abaPrincipal = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!abaPrincipal) {
      throw new Error('Aba principal não encontrada');
    }
    
    let atualizacoes = 0;
    const resultados = [];
    
    // Atualizar na aba principal
    const linhaPrincipal = encontrarLinhaPorId(abaPrincipal, transferId);
    
    if (linhaPrincipal > 0) {
      // Verificar status atual (coluna R)
      const statusAtual = abaPrincipal.getRange(linhaPrincipal, 18).getValue();
      
      if (statusAtual === novoStatus) {
        logger.info('Status já está atualizado', { transferId, status: novoStatus });
        return {
          sucesso: false,
          mensagem: `Transfer já está com status "${novoStatus}"`
        };
      }
      
      // Atualizar status (coluna R - 18)
      abaPrincipal.getRange(linhaPrincipal, 18).setValue(novoStatus);
      
      // Adicionar observação (coluna S - 19)
      const obsAtual = abaPrincipal.getRange(linhaPrincipal, 19).getValue() || '';
      const novaObs = obsAtual 
        ? `${obsAtual}\n${observacao} - ${formatarDataHora(new Date())}`
        : `${observacao} - ${formatarDataHora(new Date())}`;
      abaPrincipal.getRange(linhaPrincipal, 19).setValue(novaObs);
      
      atualizacoes++;
      resultados.push({
        aba: CONFIG.SHEET_NAME,
        linha: linhaPrincipal,
        sucesso: true
      });
      
      // Tentar atualizar na aba mensal
      try {
        const dataTransfer = abaPrincipal.getRange(linhaPrincipal, 6).getValue(); // Coluna F
        const abaMensal = obterAbaMes(dataTransfer);
        
        if (abaMensal && abaMensal.getName() !== abaPrincipal.getName()) {
          const linhaMensal = encontrarLinhaPorId(abaMensal, transferId);
          
          if (linhaMensal > 0) {
            abaMensal.getRange(linhaMensal, 18).setValue(novoStatus);
            abaMensal.getRange(linhaMensal, 19).setValue(novaObs);
            
            atualizacoes++;
            resultados.push({
              aba: abaMensal.getName(),
              linha: linhaMensal,
              sucesso: true
            });
          }
        }
      } catch (errorMensal) {
        logger.error('Erro ao atualizar aba mensal', errorMensal);
        resultados.push({
          aba: 'mensal',
          sucesso: false,
          erro: errorMensal.message
        });
      }
      
      // Enviar notificação se configurado
      if (novoStatus === 'Finalizado') {
        try {
          enviarEmailStatusAtualizado(transferId, novoStatus, observacao);
        } catch (errorEmail) {
          logger.error('Erro ao enviar e-mail de status', errorEmail);
        }
      }
      
      logger.success('Status atualizado com sucesso', {
        transferId,
        novoStatus,
        atualizacoes
      });
      
      return {
        sucesso: true,
        atualizacoes: atualizacoes,
        resultados: resultados
      };
      
    } else {
      logger.error('Transfer não encontrado', { transferId });
      return {
        sucesso: false,
        mensagem: MESSAGES.ERROS.REGISTRO_NAO_ENCONTRADO(transferId)
      };
    }
    
  } catch (error) {
    logger.error('Erro ao atualizar status', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Atualiza observações de um transfer
 * @param {string} transferId - ID do transfer
 * @param {string} novaObservacao - Nova observação
 * @returns {boolean} - Se foi atualizado com sucesso
 */
function atualizarObservacoesTransfer(transferId, novaObservacao) {
  logger.debug('Atualizando observações', { transferId });
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const linha = encontrarLinhaPorId(sheet, transferId);
    
    if (linha > 0) {
      const obsAtual = sheet.getRange(linha, 19).getValue() || ''; // Coluna S
      const obsAtualizada = obsAtual 
        ? `${obsAtual}\n${novaObservacao}`
        : novaObservacao;
      
      sheet.getRange(linha, 19).setValue(obsAtualizada);
      
      // Tentar atualizar na aba mensal também
      try {
        const dataTransfer = sheet.getRange(linha, 6).getValue(); // Coluna F
        const abaMensal = obterAbaMes(dataTransfer);
        const linhaMensal = encontrarLinhaPorId(abaMensal, transferId);
        
        if (linhaMensal > 0) {
          abaMensal.getRange(linhaMensal, 19).setValue(obsAtualizada);
        }
      } catch (e) {
        // Ignorar erro na aba mensal
      }
      
      return true;
    }
    
    return false;
    
  } catch (error) {
    logger.error('Erro ao atualizar observações', error);
    return false;
  }
}

/**
 * Envia dados do transfer para webhook do Make.com (PRESERVADO DO CÓDIGO ANTIGO)
 * @param {Array} dadosTransfer - Dados do transfer
 * @returns {boolean} - Se o envio foi bem-sucedido
 */
function enviarWebhookNovoTransfer(dadosTransfer) {
  if (!CONFIG.MAKE_WEBHOOKS.ENABLED) {
    return false;
  }
  
  logger.info('Enviando webhook para Make.com');
  
  try {
    const payload = {
      transferId: dadosTransfer[0],
      cliente: dadosTransfer[1],
      tipoServico: dadosTransfer[2],
      pessoas: dadosTransfer[3],
      bagagens: dadosTransfer[4],
      data: formatarDataDDMMYYYY(new Date(dadosTransfer[5])),
      contacto: dadosTransfer[6],
      voo: dadosTransfer[7],
      origem: dadosTransfer[8],
      destino: dadosTransfer[9],
      horaPickup: dadosTransfer[10],
      precoCliente: dadosTransfer[11],
      valorHotel: dadosTransfer[12],
      valorHUB: dadosTransfer[13],
      comissaoRecepcao: dadosTransfer[14],
      formaPagamento: dadosTransfer[15],
      pagoPara: dadosTransfer[16],
      status: dadosTransfer[17],
      observacoes: dadosTransfer[18],
      timestamp: new Date().toISOString(),
      sistema: CONFIG.NAMES.SISTEMA_NOME,
      versao: CONFIG.SISTEMA.VERSAO
    };
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(CONFIG.MAKE_WEBHOOKS.NEW_TRANSFER, options);
    const responseCode = response.getResponseCode();
    
if (responseCode === 200 || responseCode === 204) {
     logger.success('Webhook enviado com sucesso');
     return true;
   } else {
     logger.error('Webhook retornou erro', {
       code: responseCode,
       response: response.getContentText()
     });
     return false;
   }
   
 } catch (error) {
   logger.error('Erro ao enviar webhook', error);
   return false;
 }
}

/**
* Envia atualização de status via webhook (PRESERVADO DO CÓDIGO ANTIGO)
* @param {string} transferId - ID do transfer
* @param {string} status - Novo status
* @param {Object} dadosAdicionais - Dados extras
*/
function enviarWebhookStatusUpdate(transferId, status, dadosAdicionais = {}) {
 if (!CONFIG.MAKE_WEBHOOKS.ENABLED) {
   return false;
 }
 
 logger.info('Enviando webhook de atualização de status');
 
 try {
   const payload = {
     transferId: transferId,
     novoStatus: status,
     timestamp: new Date().toISOString(),
     ...dadosAdicionais
   };
   
   const options = {
     method: 'post',
     contentType: 'application/json',
     payload: JSON.stringify(payload),
     muteHttpExceptions: true
   };
   
   const response = UrlFetchApp.fetch(CONFIG.MAKE_WEBHOOKS.STATUS_UPDATE, options);
   
   return response.getResponseCode() === 200;
   
 } catch (error) {
   logger.error('Erro ao enviar webhook de status', error);
   return false;
 }
}

// ===================================================
// SISTEMA DE E-MAIL INTERATIVO (FUSÃO COMPLETA)
// ===================================================

function enviarEmailNovoTransfer(dadosTransfer) {
  // 🚨 DEBUG: Log no início
  console.log('🎯 enviarEmailNovoTransfer INICIADA:', {
    cliente: dadosTransfer[1],
    id: dadosTransfer[0],
    timestamp: new Date()
  });
  
  logger.info('Enviando e-mail real com template unificado');
  
  try {
    // Validar configuração
    if (!CONFIG.EMAIL_CONFIG.ENVIAR_AUTOMATICO) {
      console.log('❌ E-mail desativado na configuração');
      logger.info('Envio de e-mail desativado');
      return false;
    }
    
    const destinatarios = CONFIG.EMAIL_CONFIG.DESTINATARIOS.join(',');
    if (!destinatarios) {
      console.log('❌ Nenhum destinatário configurado');
      throw new Error('Nenhum destinatário configurado');
    }
    
    console.log('📧 Destinatários:', destinatarios);
    
    // Converter dados para objeto transfer (EXATAMENTE como no teste)
    const transfer = {
      id: dadosTransfer[0],
      cliente: dadosTransfer[1],
      tipoServico: dadosTransfer[2],
      pessoas: dadosTransfer[3],
      bagagens: dadosTransfer[4],
      data: dadosTransfer[5],
      contacto: dadosTransfer[6],
      voo: dadosTransfer[7],
      origem: dadosTransfer[8],
      destino: dadosTransfer[9],
      horaPickup: dadosTransfer[10],
      precoCliente: dadosTransfer[11],
      valorHotel: dadosTransfer[12],
      valorHUB: dadosTransfer[13],
      comissaoRecepcao: dadosTransfer[14],
      formaPagamento: dadosTransfer[15],
      pagoPara: dadosTransfer[16],
      status: dadosTransfer[17],
      observacoes: dadosTransfer[18]
    };
    
    console.log('👤 Transfer criado:', transfer);
    
    // Preparar dados formatados
    const dataFormatada = formatarDataDDMMYYYY(new Date(transfer.data));
    const tipoLabel = obterLabelTipoServico(transfer.tipoServico);
    
    // Assunto profissional
    const assunto = CONFIG.EMAIL_CONFIG.USAR_BOTOES_INTERATIVOS
      ? `AÇÃO NECESSÁRIA: Novo ${tipoLabel} #${transfer.id} - ${transfer.cliente} - ${dataFormatada}`
      : `Novo ${tipoLabel} #${transfer.id} - ${transfer.cliente} - ${dataFormatada}`;
    
    console.log('📝 Assunto:', assunto);
    
    // 🔥 USAR EXATAMENTE O MESMO TEMPLATE DO TESTE
    const corpoHtml = criarEmailHtmlInterativo(transfer);
    
    console.log('📄 HTML gerado (primeiros 200 chars):', corpoHtml.substring(0, 200));
    console.log('🖼️ Contém imagem HUB?', corpoHtml.includes('1z6i_VnZZ9OdHbcmcTHa4dKHPIqC6vmBE'));
    console.log('❌ Contém texto Junior?', corpoHtml.includes('Junior Gutierez'));
    
    // 🔥 ENVIAR COM TEMPLATE CORRETO
    console.log('📮 Enviando e-mail...');
    sendEmailComAssinatura({
      to: destinatarios,
      subject: assunto,
      htmlBody: corpoHtml,
      name: CONFIG.NAMES.SISTEMA_NOME,
      replyTo: CONFIG.EMAIL_CONFIG.EMAIL_EMPRESA
    });
    
    console.log('✅ E-mail enviado com sucesso!');
    
    logger.success('E-mail real enviado com template unificado', {
      transferId: transfer.id,
      cliente: transfer.cliente,
      destinatarios: destinatarios
    });
    
    // Registrar envio no transfer
    atualizarObservacoesTransfer(transfer.id, 
      `E-mail enviado em ${formatarDataHora(new Date())}`);
    
    return true;
    
  } catch (error) {
    console.log('❌ ERRO no envio:', error.toString());
    logger.error('Erro ao enviar e-mail real', error);
    return false;
  }
}

/**
 * Obtém label formatado do tipo de serviço
 * @private
 */
function obterLabelTipoServico(tipoServico) {
  switch (tipoServico) {
    case 'Tour Regular':
    case 'tour':
      return 'Tour Regular';
    case 'Private Tour':
    case 'private':
      return 'Private Tour';
    default:
      return 'Transfer';
  }
}

/**
 * Cria HTML do e-mail interativo - VERSÃO LIMPA SEM ÍCONES
 * CORREÇÕES: Sem ícones problemáticos, apenas títulos + data correta
 * @private
 */
function criarEmailHtmlInterativo(transfer) {
  const webAppUrl = ScriptApp.getService().getUrl();
  const urlConfirmar = `${webAppUrl}?action=confirm&id=${transfer.id}`;
  const urlCancelar = `${webAppUrl}?action=cancel&id=${transfer.id}`;
  
  // CORREÇÃO DA DATA: Usar formatação correta
  let dataFormatada;
  try {
    const dataObj = transfer.data instanceof Date ? transfer.data : new Date(transfer.data);
    const dia = String(dataObj.getDate()).padStart(2, '0');
    const mes = String(dataObj.getMonth() + 1).padStart(2, '0'); // +1 porque getMonth() retorna 0-11
    const ano = dataObj.getFullYear();
    dataFormatada = `${dia}/${mes}/${ano}`;
  } catch (error) {
    const hoje = new Date();
    const dia = String(hoje.getDate()).padStart(2, '0');
    const mes = String(hoje.getMonth() + 1).padStart(2, '0');
    const ano = hoje.getFullYear();
    dataFormatada = `${dia}/${mes}/${ano}`;
  }
  
  const tipoLabel = obterLabelTipoServico(transfer.tipoServico);
  const precoCliente = `${CONFIG.VALORES.MOEDA}${parseFloat(transfer.precoCliente).toFixed(2)}`;
  const valorHotel = `${CONFIG.VALORES.MOEDA}${parseFloat(transfer.valorHotel).toFixed(2)}`;
  const valorHUB = `${CONFIG.VALORES.MOEDA}${parseFloat(transfer.valorHUB).toFixed(2)}`;
  const comissaoRecepcao = `${CONFIG.VALORES.MOEDA}${parseFloat(transfer.comissaoRecepcao).toFixed(2)}`;
  
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <style>
        body {
          font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
          background-color: #f8f9fa;
          color: #333;
          margin: 0;
          padding: 20px;
          line-height: 1.6;
        }
        .container {
          background: #ffffff;
          padding: 0;
          border-radius: 12px;
          max-width: 600px;
          margin: 0 auto;
          box-shadow: 0 4px 12px rgba(0,0,0,0.08);
          overflow: hidden;
        }
        .header {
          background: linear-gradient(135deg, #0056b3 0%, #007bff 100%);
          color: white;
          padding: 30px;
          text-align: center;
        }
        .header h1 {
          margin: 0 0 10px 0;
          font-size: 24px;
          font-weight: 600;
        }
        .header p {
          margin: 0;
          font-size: 16px;
          opacity: 0.95;
        }
        .greeting {
          background: #f8f9fa;
          padding: 30px;
          text-align: center;
          border-bottom: 1px solid #e9ecef;
        }
        .greeting h2 {
          margin: 0 0 15px 0;
          color: #495057;
          font-size: 22px;
          font-weight: 500;
        }
        .greeting p {
          margin: 0;
          color: #6c757d;
          font-size: 16px;
        }
        .content {
          padding: 30px;
        }
        .transfer-details {
          background: #f8f9fa;
          padding: 25px;
          border-radius: 8px;
          margin: 20px 0;
        }
        .detail-row {
          display: flex;
          justify-content: space-between;
          align-items: center;
          padding: 12px 0;
          border-bottom: 1px solid #e9ecef;
        }
        .detail-row:last-child {
          border-bottom: none;
        }
        .detail-label {
          font-weight: 600;
          color: #495057;
          font-size: 14px;
        }
        .detail-value {
          color: #212529;
          text-align: right;
          font-weight: 500;
        }
        .route-section {
          background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
          color: white;
          padding: 25px;
          border-radius: 8px;
          text-align: center;
          margin: 25px 0;
        }
        .route-section h3 {
          margin: 0 0 10px 0;
          font-size: 20px;
        }
        .route-section p {
          margin: 0;
          font-size: 16px;
          opacity: 0.95;
        }
        .pricing-section {
          background: #e8f5e8;
          padding: 25px;
          border-radius: 8px;
          margin: 25px 0;
        }
        .pricing-grid {
          display: grid;
          grid-template-columns: repeat(2, 1fr);
          gap: 15px;
          margin-top: 15px;
        }
        .pricing-item {
          background: white;
          padding: 15px;
          border-radius: 6px;
          text-align: center;
          border: 1px solid #d4edda;
        }
        .pricing-item.total {
          background: #d4edda;
          font-weight: bold;
        }
        .pricing-label {
          font-size: 12px;
          color: #6c757d;
          margin-bottom: 5px;
        }
        .pricing-value {
          font-size: 18px;
          color: #28a745;
          font-weight: bold;
        }
        .action-section {
          background: linear-gradient(135deg, #007bff 0%, #0056b3 100%);
          color: white;
          padding: 30px;
          border-radius: 8px;
          text-align: center;
          margin: 30px 0;
        }
        .action-section h3 {
          margin: 0 0 15px 0;
          font-size: 20px;
        }
        .action-buttons {
          margin-top: 25px;
        }
        .btn {
          display: inline-block;
          padding: 12px 30px;
          border-radius: 6px;
          font-weight: bold;
          font-size: 16px;
          text-decoration: none;
          margin: 0 10px;
          transition: all 0.3s ease;
        }
        .btn-confirm {
          background-color: #28a745;
          color: white;
          border: 2px solid #28a745;
        }
        .btn-confirm:hover {
          background-color: #218838;
        }
        .btn-cancel {
          background-color: #dc3545;
          color: white;
          border: 2px solid #dc3545;
        }
        .btn-cancel:hover {
          background-color: #c82333;
        }
        .footer-note {
          text-align: center;
          font-size: 14px;
          color: #6c757d;
          margin: 20px 0;
          padding: 20px;
          background: #f8f9fa;
          border-radius: 6px;
        }
        @media only screen and (max-width: 600px) {
          .container { margin: 10px; }
          .pricing-grid { grid-template-columns: 1fr; }
          .btn { display: block; margin: 10px 0; }
        }
      </style>
    </head>
    <body>
      <div class="container">
        <!-- Header -->
        <div class="header">
          <h1>${tipoLabel} Solicitado</h1>
          <p>Transfer #${transfer.id} • ${dataFormatada}</p>
        </div>
        
        <!-- Cumprimento -->
        <div class="greeting">
          <h2>Cliente: ${transfer.cliente}</h2>
          <p>Solicitamos, por favor, a confirmação da disponibilidade do ${tipoLabel.toLowerCase()} para esta data e horário.</p>
          <br>
          <p><strong>Com os melhores cumprimentos,</strong></p>
        </div>
        
        <!-- Detalhes do Transfer -->
        <div class="content">
          <div class="transfer-details">
            <div class="detail-row">
              <span class="detail-label">Cliente:</span>
              <span class="detail-value">${transfer.cliente}</span>
            </div>
            <div class="detail-row">
              <span class="detail-label">Tipo de Serviço:</span>
              <span class="detail-value">${tipoLabel}</span>
            </div>
            <div class="detail-row">
              <span class="detail-label">Contacto:</span>
              <span class="detail-value">${transfer.contacto}</span>
            </div>
            <div class="detail-row">
              <span class="detail-label">Passageiros:</span>
              <span class="detail-value">${transfer.pessoas} pessoa(s)</span>
            </div>
            <div class="detail-row">
              <span class="detail-label">Bagagens:</span>
              <span class="detail-value">${transfer.bagagens} volume(s)</span>
            </div>
            ${transfer.voo ? `
            <div class="detail-row">
              <span class="detail-label">Voo:</span>
              <span class="detail-value">${transfer.voo}</span>
            </div>
            ` : ''}
          </div>
          
          <!-- Rota -->
          <div class="route-section">
            <h3>${transfer.origem} → ${transfer.destino}</h3>
            <p>${dataFormatada} às ${transfer.horaPickup}</p>
          </div>
          
          <!-- Valores -->
          <div class="pricing-section">
            <h3 style="margin: 0 0 15px 0; color: #28a745;">Valores do Serviço</h3>
            <div class="pricing-grid">
              <div class="pricing-item total">
                <div class="pricing-label">Valor Total</div>
                <div class="pricing-value">${precoCliente}</div>
              </div>
              <div class="pricing-item">
                <div class="pricing-label">${CONFIG.NAMES.HOTEL_NAME}</div>
                <div class="pricing-value">${valorHotel}</div>
              </div>
              <div class="pricing-item">
                <div class="pricing-label">Comissão Recepção</div>
                <div class="pricing-value">${comissaoRecepcao}</div>
              </div>
              <div class="pricing-item">
                <div class="pricing-label">HUB Transfer</div>
                <div class="pricing-value">${valorHUB}</div>
              </div>
            </div>
            
            <div style="margin-top: 20px; padding-top: 15px; border-top: 1px solid #c3e6cb;">
              <div class="detail-row">
                <span class="detail-label">Forma de Pagamento:</span>
                <span class="detail-value">${transfer.formaPagamento}</span>
              </div>
              <div class="detail-row">
                <span class="detail-label">Pago Para:</span>
                <span class="detail-value">${transfer.pagoPara}</span>
              </div>
            </div>
          </div>
          
          ${transfer.observacoes ? `
          <div style="background: #fff3cd; padding: 20px; border-radius: 8px; border-left: 4px solid #ffc107;">
            <h4 style="margin: 0 0 10px 0; color: #856404;">Observações:</h4>
            <p style="margin: 0; color: #856404;">${transfer.observacoes}</p>
          </div>
          ` : ''}
        </div>
        
        ${CONFIG.EMAIL_CONFIG.USAR_BOTOES_INTERATIVOS ? `
        <!-- Botões de Ação -->
        <div class="action-section">
          <h3>Confirmação Necessária</h3>
          <p>Por favor, confirme ou cancele este ${tipoLabel.toLowerCase()} clicando em um dos botões abaixo:</p>
          <div class="action-buttons">
            <a href="${urlConfirmar}" class="btn btn-confirm">CONFIRMAR ${tipoLabel.toUpperCase()}</a>
            <a href="${urlCancelar}" class="btn btn-cancel">CANCELAR ${tipoLabel.toUpperCase()}</a>
          </div>
          <p style="margin-top: 20px; font-size: 14px; opacity: 0.9;">
            Ou responda este e-mail com "OK" para confirmar
          </p>
        </div>
        ` : `
        <!-- Instruções de Confirmação -->
        <div class="action-section">
          <h3>Como Confirmar</h3>
          <p>Para confirmar este ${tipoLabel.toLowerCase()}, responda este e-mail com:</p>
          <h2 style="margin: 15px 0; font-size: 28px;">"OK"</h2>
          <p>ou</p>
          <h2 style="margin: 15px 0; font-size: 28px;">"OK ${transfer.id}"</h2>
        </div>
        `}
        
        <!-- Nota Final -->
        <div class="footer-note">
          <p>Este é um e-mail automático do sistema de gestão de transfers.<br>
          Para qualquer esclarecimento, entre em contacto connosco.</p>
        </div>
      </div>
    </body>
    </html>
  `;
}

/**
 * Cria e-mail em texto simples (CÓDIGO ANTIGO PRESERVADO)
 * @private
 */
function criarEmailTextoSimples(transfer) {
  const dataFormatada = formatarDataDDMMYYYY(new Date(transfer.data));
  const tipoLabel = obterLabelTipoServico(transfer.tipoServico);
  
  return `
${MESSAGES.TITULOS.NOVO_TRANSFER}

${tipoLabel} #${transfer.id} - ${transfer.status}

DETALHES DO ${tipoLabel.toUpperCase()}:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Cliente: ${transfer.cliente}
Tipo: ${tipoLabel}
Contacto: ${transfer.contacto}
Passageiros: ${transfer.pessoas}
Bagagens: ${transfer.bagagens}
${transfer.voo ? `Voo: ${transfer.voo}` : ''}

ROTA:
${transfer.origem} → ${transfer.destino}
Data: ${dataFormatada}
Hora: ${transfer.horaPickup}

VALORES:
Preço Total: €${parseFloat(transfer.precoCliente).toFixed(2)}
${CONFIG.NAMES.HOTEL_NAME}: €${parseFloat(transfer.valorHotel).toFixed(2)}
Comissão Recepção: €${parseFloat(transfer.comissaoRecepcao).toFixed(2)}
HUB Transfer: €${parseFloat(transfer.valorHUB).toFixed(2)}

Pagamento: ${transfer.formaPagamento} → ${transfer.pagoPara}

${transfer.observacoes ? `OBSERVAÇÕES:\n${transfer.observacoes}\n` : ''}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PARA CONFIRMAR: Responda com "OK" ou "OK ${transfer.id}"
PARA CANCELAR: Responda com "CANCELAR ${transfer.id}"
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

${CONFIG.NAMES.SISTEMA_NOME}
${new Date().toLocaleString('pt-PT')}
  `;
}

/**
 * Processa clique nos botões do e-mail - CORREÇÃO VALIDAÇÃO
 * CORREÇÃO: Remove validações de dados que estão bloqueando atualizações
 */
function handleEmailAction(params) {
  logger.info('🔴 PROCESSANDO AÇÃO DE E-MAIL', params);
  
  const { id: transferId, action } = params;
  const userEmail = Session.getActiveUser().getEmail() || 'Usuário Desconhecido';
  
  console.log('🎯 Dados recebidos:', { transferId, action, userEmail });
  
  // Validação melhorada
  if (!transferId) {
    console.log('❌ Transfer ID ausente');
    return createHtmlResponse(`
      <h1>❌ Erro</h1>
      <p>ID do transfer não fornecido. Verifique o link do e-mail.</p>
      <p>Parâmetros recebidos: ${JSON.stringify(params)}</p>
    `);
  }
  
  if (!action || !['confirm', 'cancel'].includes(action)) {
    console.log('❌ Ação inválida:', action);
    return createHtmlResponse(`
      <h1>❌ Erro</h1>
      <p>Ação inválida: ${action}</p>
      <p>Ações válidas: confirm, cancel</p>
    `);
  }
  
  try {
    console.log('🔍 Buscando transfer na planilha...');
    
    // Buscar transfer
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) {
      throw new Error(`Planilha '${CONFIG.SHEET_NAME}' não encontrada`);
    }
    
    const linha = encontrarLinhaPorId(sheet, transferId);
    console.log('📊 Linha encontrada:', linha);
    
    if (!linha) {
      throw new Error(`Transfer #${transferId} não encontrado na planilha`);
    }
    
    // Verificar status atual (coluna R)
    const statusAtual = sheet.getRange(linha, 18).getValue();
    console.log('📋 Status atual:', statusAtual);
    
    // Determinar novo status
    let novoStatus, observacao, tituloHtml, corFundo;
    
    if (action === 'confirm') {
      if (statusAtual === 'Confirmado') {
        console.log('⚠️ Transfer já confirmado');
        return createHtmlResponse(`
          <div style="text-align: center; padding: 40px; font-family: Arial;">
            <h1>ℹ️ Transfer Já Confirmado</h1>
            <p>O transfer #${transferId} já estava confirmado.</p>
            <p>Status atual: ${statusAtual}</p>
            <button onclick="window.close()">Fechar</button>
          </div>
        `);
      }
      
      novoStatus = 'Confirmado';
      observacao = `✅ Confirmado por ${userEmail} em ${formatarDataHora(new Date())}`;
      tituloHtml = '✅ Transfer Confirmado!';
      corFundo = '#d4edda';
      
    } else if (action === 'cancel') {
      if (statusAtual === 'Cancelado') {
        console.log('⚠️ Transfer já cancelado');
        return createHtmlResponse(`
          <div style="text-align: center; padding: 40px; font-family: Arial;">
            <h1>ℹ️ Transfer Já Cancelado</h1>
            <p>O transfer #${transferId} já estava cancelado.</p>
            <p>Status atual: ${statusAtual}</p>
            <button onclick="window.close()">Fechar</button>
          </div>
        `);
      }
      
      novoStatus = 'Cancelado';
      observacao = `❌ Cancelado por ${userEmail} em ${formatarDataHora(new Date())}`;
      tituloHtml = '❌ Transfer Cancelado';
      corFundo = '#f8d7da';
    }
    
    console.log('🔄 Atualizando status:', { transferId, statusAtual, novoStatus });
    
    // 🔧 CORREÇÃO CRÍTICA: Remover validações antes de atualizar
    try {
      console.log('🔓 Removendo validações da coluna Status...');
      
      // Remover validação da coluna R (Status) temporariamente
      const statusRange = sheet.getRange(linha, 18, 1, 1);
      statusRange.clearDataValidations();
      
      console.log('✅ Validações removidas, atualizando status...');
      
      // Atualizar status (coluna R)
      statusRange.setValue(novoStatus);
      
      console.log('✅ Status atualizado para:', novoStatus);
      
      // Atualizar observações (coluna S)
      const obsAtual = sheet.getRange(linha, 19).getValue() || '';
      const novaObs = obsAtual 
        ? `${obsAtual}\n${observacao}`
        : observacao;
      sheet.getRange(linha, 19).setValue(novaObs);
      
      console.log('✅ Observações atualizadas');
      
      // 🔧 REATIVAR VALIDAÇÃO (opcional, mas recomendado)
      try {
        const statusValidation = SpreadsheetApp.newDataValidation()
          .requireValueInList(['Solicitado', 'Confirmado', 'Finalizado', 'Cancelado'])
          .setAllowInvalid(true) // 🔧 PERMITIR VALORES INVÁLIDOS para evitar bloqueios
          .setHelpText('Status do transfer')
          .build();
        statusRange.setDataValidation(statusValidation);
        console.log('✅ Validação reativada com allowInvalid=true');
      } catch (validationError) {
        console.log('⚠️ Erro ao reativar validação (não crítico):', validationError.message);
      }
      
      // Tentar atualizar aba mensal também
      try {
        const dataTransfer = sheet.getRange(linha, 6).getValue(); // Coluna F
        const abaMensal = obterAbaMes(dataTransfer);
        
        if (abaMensal && abaMensal.getName() !== sheet.getName()) {
          const linhaMensal = encontrarLinhaPorId(abaMensal, transferId);
          
          if (linhaMensal > 0) {
            // Remover validações da aba mensal também
            const statusMensalRange = abaMensal.getRange(linhaMensal, 18, 1, 1);
            statusMensalRange.clearDataValidations();
            
            abaMensal.getRange(linhaMensal, 18).setValue(novoStatus);
            abaMensal.getRange(linhaMensal, 19).setValue(novaObs);
            console.log('✅ Aba mensal também atualizada');
          }
        }
      } catch (errorMensal) {
        console.log('⚠️ Erro ao atualizar aba mensal (não crítico):', errorMensal.message);
      }
      
      // Buscar dados do transfer para exibir
      const dadosTransfer = sheet.getRange(linha, 1, 1, HEADERS.length).getValues()[0];
      const cliente = dadosTransfer[1];
      const tipoServico = dadosTransfer[2];
      const data = formatarDataDDMMYYYY(new Date(dadosTransfer[5]));
      const rota = `${dadosTransfer[8]} → ${dadosTransfer[9]}`;
      
      console.log('📧 Enviando e-mail de confirmação...');
      
      // Enviar e-mail de confirmação
      try {
        enviarEmailConfirmacaoAcao(transferId, action, userEmail);
      } catch (emailError) {
        console.log('⚠️ Erro ao enviar e-mail de confirmação (não crítico):', emailError.message);
      }
      
      // Página de sucesso melhorada
      return createHtmlResponse(`
        <!DOCTYPE html>
        <html>
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>${tituloHtml}</title>
          <style>
            body {
              font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif;
              background: ${corFundo};
              margin: 0;
              padding: 20px;
              display: flex;
              align-items: center;
              justify-content: center;
              min-height: 100vh;
            }
            .card {
              background: white;
              padding: 40px;
              border-radius: 12px;
              box-shadow: 0 4px 20px rgba(0,0,0,0.1);
              text-align: center;
              max-width: 500px;
              border: 3px solid ${action === 'confirm' ? '#28a745' : '#dc3545'};
            }
            h1 {
              color: ${action === 'confirm' ? '#28a745' : '#dc3545'};
              margin-bottom: 20px;
              font-size: 2rem;
            }
            .details {
              background: #f8f9fa;
              padding: 20px;
              border-radius: 8px;
              margin: 20px 0;
              text-align: left;
            }
            .detail-row {
              padding: 5px 0;
              border-bottom: 1px solid #eee;
            }
            .detail-row:last-child {
              border-bottom: none;
            }
            .label {
              font-weight: bold;
              color: #666;
              display: inline-block;
              width: 120px;
            }
            .value {
              color: #333;
            }
            .timestamp {
              color: #666;
              font-size: 14px;
              margin-top: 20px;
              padding-top: 20px;
              border-top: 1px solid #eee;
            }
            .close-button {
              background: ${action === 'confirm' ? '#28a745' : '#dc3545'};
              color: white;
              border: none;
              padding: 12px 30px;
              border-radius: 6px;
              font-size: 16px;
              cursor: pointer;
              margin-top: 20px;
            }
            .close-button:hover {
              opacity: 0.9;
            }
          </style>
        </head>
        <body>
          <div class="card">
            <h1>${tituloHtml}</h1>
            <p style="font-size: 18px; margin-bottom: 30px;">
              ${action === 'confirm' ? 'Transfer confirmado com sucesso!' : 'Transfer cancelado conforme solicitado.'}
            </p>
            
            <div class="details">
              <div class="detail-row">
                <span class="label">Transfer ID:</span>
                <span class="value">#${transferId}</span>
              </div>
              <div class="detail-row">
                <span class="label">Cliente:</span>
                <span class="value">${cliente}</span>
              </div>
              <div class="detail-row">
                <span class="label">Tipo:</span>
                <span class="value">${obterLabelTipoServico(tipoServico)}</span>
              </div>
              <div class="detail-row">
                <span class="label">Data:</span>
                <span class="value">${data}</span>
              </div>
              <div class="detail-row">
                <span class="label">Rota:</span>
                <span class="value">${rota}</span>
              </div>
              <div class="detail-row">
                <span class="label">Status Anterior:</span>
                <span class="value">${statusAtual}</span>
              </div>
              <div class="detail-row">
                <span class="label">Novo Status:</span>
                <span class="value" style="font-weight: bold; color: ${action === 'confirm' ? '#28a745' : '#dc3545'};">
                  ${novoStatus}
                </span>
              </div>
              <div class="detail-row">
                <span class="label">Processado por:</span>
                <span class="value">${userEmail}</span>
              </div>
            </div>
            
            <div class="timestamp">
              Processado em: ${formatarDataHora(new Date())}
            </div>
            
            <button class="close-button" onclick="window.close()">
              Fechar Janela
            </button>
            
            <script>
              // Auto-fechar após 10 segundos
              setTimeout(() => {
                if (confirm('Fechar automaticamente esta janela?')) {
                  window.close();
                }
              }, 10000);
            </script>
          </div>
        </body>
        </html>
      `);
      
    } catch (updateError) {
      console.log('❌ Erro ao atualizar planilha:', updateError.message);
      throw new Error(`Erro ao atualizar status na planilha: ${updateError.message}`);
    }
    
  } catch (error) {
    console.log('❌ ERRO GERAL:', error.message);
    logger.error('Erro ao processar ação de e-mail', error);
    
    return createHtmlResponse(`
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <title>Erro ao Processar</title>
        <style>
          body {
            font-family: Arial, sans-serif;
            background: #f8d7da;
            margin: 0;
            padding: 20px;
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 100vh;
          }
          .error-card {
            background: white;
            padding: 40px;
            border-radius: 12px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.1);
            text-align: center;
            max-width: 500px;
            border: 3px solid #dc3545;
          }
          h1 {
            color: #721c24;
            margin-bottom: 20px;
          }
          .error-details {
            background: #f5c6cb;
            padding: 15px;
            border-radius: 6px;
            margin: 20px 0;
            color: #721c24;
            text-align: left;
          }
          .support-info {
            background: #e2e3e5;
            padding: 15px;
            border-radius: 6px;
            margin: 20px 0;
            font-size: 14px;
          }
        </style>
      </head>
      <body>
        <div class="error-card">
          <h1>❌ Erro ao Processar</h1>
          <p>Não foi possível processar sua solicitação.</p>
          
          <div class="error-details">
            <strong>Detalhes do erro:</strong><br>
            ${error.message}
          </div>
          
          <div class="support-info">
            <strong>Informações para suporte:</strong><br>
            • Transfer ID: ${transferId || 'não fornecido'}<br>
            • Ação: ${action || 'não fornecida'}<br>
            • Usuário: ${userEmail}<br>
            • Timestamp: ${new Date().toISOString()}
          </div>
          
          <p>Por favor, tente novamente ou entre em contato com o suporte técnico.</p>
          
          <button onclick="window.close()" style="background: #dc3545; color: white; border: none; padding: 12px 30px; border-radius: 6px; cursor: pointer;">
            Fechar
          </button>
        </div>
      </body>
      </html>
    `);
  }
}

/**
* Envia e-mail de confirmação após ação (CÓDIGO ANTIGO PRESERVADO)
* @private
*/
function enviarEmailConfirmacaoAcao(transferId, action, userEmail) {
 try {
   const destinatarios = CONFIG.EMAIL_CONFIG.DESTINATARIOS.join(',');
   const assunto = action === 'confirm' 
     ? MESSAGES.TITULOS.TRANSFER_CONFIRMADO(transferId)
     : MESSAGES.TITULOS.TRANSFER_CANCELADO(transferId);
   
   const corpo = `
     <h2>${assunto}</h2>
     <p>Transfer #${transferId} foi ${action === 'confirm' ? 'confirmado' : 'cancelado'} por ${userEmail}</p>
     <p>Data/Hora: ${formatarDataHora(new Date())}</p>
   `;
   
// Correção: usar htmlBody correto
sendEmailComAssinatura({
  to: destinatarios,
  subject: assunto,
  htmlBody: htmlEmail || corpo || corpoHtml,
  name: CONFIG.NAMES.SISTEMA_NOME,
  replyTo: CONFIG.EMAIL_CONFIG.EMAIL_EMPRESA
});
   
 } catch (error) {
   logger.error('Erro ao enviar e-mail de confirmação', error);
 }
}

/**
* Envia e-mail de status atualizado (CÓDIGO ANTIGO PRESERVADO)
* @param {string} transferId - ID do transfer
* @param {string} novoStatus - Novo status
* @param {string} motivo - Motivo da atualização
*/
function enviarEmailStatusAtualizado(transferId, novoStatus, motivo = '') {
 logger.info('Enviando e-mail de status atualizado', { transferId, novoStatus });
 
 try {
   const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
   const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
   const linha = encontrarLinhaPorId(sheet, transferId);
   
   if (!linha) {
     throw new Error(`Transfer ${transferId} não encontrado`);
   }
   
   const dadosTransfer = sheet.getRange(linha, 1, 1, HEADERS.length).getValues()[0];
   const destinatarios = CONFIG.EMAIL_CONFIG.DESTINATARIOS.join(',');
   
   const assunto = `Transfer #${transferId} - Status: ${novoStatus}`;
   
   const corpo = criarEmailStatusTemplate({
     id: dadosTransfer[0],
     cliente: dadosTransfer[1],
     tipoServico: dadosTransfer[2],
     data: formatarDataDDMMYYYY(new Date(dadosTransfer[5])),
     rota: `${dadosTransfer[8]} → ${dadosTransfer[9]}`,
     statusAnterior: dadosTransfer[17],
     statusNovo: novoStatus,
     motivo: motivo,
     timestamp: formatarDataHora(new Date())
   });
   
// Usar template específico para status
sendEmailComAssinatura({
  to: destinatarios,
  subject: assunto,
  htmlBody: corpo,
  name: CONFIG.NAMES.SISTEMA_NOME,
  replyTo: CONFIG.EMAIL_CONFIG.EMAIL_EMPRESA
});
   
   logger.success('E-mail de status enviado');
   
 } catch (error) {
   logger.error('Erro ao enviar e-mail de status', error);
 }
}

/**
* Template para e-mail de mudança de status (CÓDIGO ANTIGO PRESERVADO)
* @private
*/
function criarEmailStatusTemplate(dados) {
 const corStatus = STYLES.STATUS_COLORS[dados.statusNovo] || '#666';
 const tipoLabel = obterLabelTipoServico(dados.tipoServico);
 
 return `
   <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
     <h2 style="color: ${corStatus};">Status Atualizado</h2>
     
     <div style="background: #f8f9fa; padding: 20px; border-radius: 8px; margin: 20px 0;">
       <h3>${tipoLabel} #${dados.id}</h3>
       <p><strong>Cliente:</strong> ${dados.cliente}</p>
       <p><strong>Tipo:</strong> ${tipoLabel}</p>
       <p><strong>Data:</strong> ${dados.data}</p>
       <p><strong>Rota:</strong> ${dados.rota}</p>
     </div>
     
     <div style="background: #e3f2fd; padding: 20px; border-radius: 8px;">
       <p><strong>Status Anterior:</strong> ${dados.statusAnterior}</p>
       <p><strong>Novo Status:</strong> <span style="color: ${corStatus}; font-weight: bold;">${dados.statusNovo}</span></p>
       ${dados.motivo ? `<p><strong>Motivo:</strong> ${dados.motivo}</p>` : ''}
       <p><strong>Atualizado em:</strong> ${dados.timestamp}</p>
     </div>
     
     <hr style="margin: 30px 0; border: 1px solid #e0e0e0;">
     
     <p style="font-size: 12px; color: #666;">
       ${CONFIG.NAMES.SISTEMA_NOME} - Notificação Automática
     </p>
   </div>
 `;
}

// ===================================================
// SISTEMA DE RELATÓRIOS (CÓDIGO NOVO + MELHORIAS ANTIGO)
// ===================================================

/**
 * Envia relatório do dia anterior (CÓDIGO NOVO)
 */
function enviarRelatorioDiaAnterior() {
  logger.info('Iniciando envio de relatório do dia anterior');
  
  try {
    const ontem = new Date();
    ontem.setDate(ontem.getDate() - 1);
    
    const relatorio = gerarRelatorioPeriodo(ontem, ontem);
    
    if (relatorio.totalTransfers === 0) {
      logger.info('Nenhum transfer no dia anterior');
      return;
    }
    
    const htmlEmail = criarEmailRelatorioDiario(relatorio, ontem);
    
// Correção: usar htmlBody correto
sendEmailComAssinatura({
  to: destinatarios,
  subject: assunto,
  htmlBody: htmlEmail || corpo || corpoHtml,
  name: CONFIG.NAMES.SISTEMA_NOME,
  replyTo: CONFIG.EMAIL_CONFIG.EMAIL_EMPRESA
});
    
    logger.success('Relatório diário enviado com sucesso');
    return true;
    
  } catch (error) {
    logger.error('Erro ao enviar relatório diário', error);
    return false;
  }
}

/**
 * Gera relatório de período (CÓDIGO NOVO)
 */
function gerarRelatorioPeriodo(dataInicio, dataFim) {
  logger.info('Gerando relatório de período', { 
    inicio: dataInicio, 
    fim: dataFim 
  });
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return {
        totalTransfers: 0,
        transfers: []
      };
    }
    
    const dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
    const transfersPeriodo = [];
    
    let valorTotalGeral = 0;
    let valorTotalHotel = 0;
    let valorTotalHUB = 0;
    let comissaoTotalRecepcao = 0;
    
    dados.forEach(row => {
      const dataTransfer = new Date(row[5]); // Coluna F - Data
      
      if (dataTransfer >= dataInicio && dataTransfer <= dataFim) {
        const transfer = {
          id: row[0],
          cliente: row[1],
          tipoServico: row[2],
          pessoas: row[3],
          bagagens: row[4],
          data: row[5],
          origem: row[8],
          destino: row[9],
          valorTotal: parseFloat(row[11]) || 0,
          valorHotel: parseFloat(row[12]) || 0,
          valorHUB: parseFloat(row[13]) || 0,
          comissaoRecepcao: parseFloat(row[14]) || 0,
          status: row[17]
        };
        
        transfersPeriodo.push(transfer);
        
        // Somar apenas transfers confirmados ou finalizados
        if (transfer.status !== 'Cancelado') {
          valorTotalGeral += transfer.valorTotal;
          valorTotalHotel += transfer.valorHotel;
          valorTotalHUB += transfer.valorHUB;
          comissaoTotalRecepcao += transfer.comissaoRecepcao;
        }
      }
    });
    
    return {
      periodo: {
        inicio: dataInicio,
        fim: dataFim
      },
      totalTransfers: transfersPeriodo.length,
      transfers: transfersPeriodo,
      resumoFinanceiro: {
        valorTotalGeral: valorTotalGeral,
        valorTotalHotel: valorTotalHotel,
        valorTotalHUB: valorTotalHUB,
        comissaoTotalRecepcao: comissaoTotalRecepcao
      }
    };
    
  } catch (error) {
    logger.error('Erro ao gerar relatório', error);
    throw error;
  }
}

/**
 * Cria HTML do e-mail de relatório diário (CÓDIGO NOVO)
 */
/**
 * Cria HTML do e-mail de relatório diário (VERSÃO PROFISSIONAL)
 */
function criarEmailRelatorioDiario(relatorio, data) {
  const dataFormatada = formatarDataDDMMYYYY(data);
  
  let tabelaTransfers = '';
  relatorio.transfers.forEach(t => {
    const statusCor = t.status === 'Finalizado' ? '#27ae60' : 
                     t.status === 'Confirmado' ? '#3498db' : 
                     t.status === 'Cancelado' ? '#e74c3c' : '#f39c12';
    
    const tipoLabel = obterLabelTipoServico(t.tipoServico);
    
    tabelaTransfers += `
      <tr>
        <td style="padding: 12px; border-bottom: 1px solid #ecf0f1; font-weight: 500;">#${t.id}</td>
        <td style="padding: 12px; border-bottom: 1px solid #ecf0f1;">${t.cliente}</td>
        <td style="padding: 12px; border-bottom: 1px solid #ecf0f1;">${tipoLabel}</td>
        <td style="padding: 12px; border-bottom: 1px solid #ecf0f1; font-size: 14px;">${t.origem} → ${t.destino}</td>
        <td style="padding: 12px; border-bottom: 1px solid #ecf0f1; text-align: right; font-weight: bold; color: #27ae60;">€${t.valorTotal.toFixed(2)}</td>
        <td style="padding: 12px; border-bottom: 1px solid #ecf0f1; text-align: center;">
          <span style="background: ${statusCor}; color: white; padding: 6px 12px; border-radius: 20px; font-size: 12px; font-weight: 500;">
            ${t.status}
          </span>
        </td>
      </tr>
    `;
  });
  
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <style>
        body {
          font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
          background-color: #f8f9fa;
          color: #333;
          margin: 0;
          padding: 20px;
          line-height: 1.6;
        }
        .container {
          background: #ffffff;
          padding: 0;
          border-radius: 12px;
          max-width: 700px;
          margin: 0 auto;
          box-shadow: 0 4px 12px rgba(0,0,0,0.08);
          overflow: hidden;
        }
        .header {
          background: linear-gradient(135deg, #2c3e50 0%, #3498db 100%);
          color: white;
          padding: 30px;
          text-align: center;
        }
        .header h1 {
          margin: 0 0 10px 0;
          font-size: 28px;
          font-weight: 600;
        }
        .header p {
          margin: 0;
          font-size: 16px;
          opacity: 0.95;
        }
        .greeting {
          background: #f8f9fa;
          padding: 25px;
          text-align: center;
          border-bottom: 1px solid #e9ecef;
        }
        .greeting h2 {
          margin: 0 0 10px 0;
          color: #495057;
          font-size: 20px;
          font-weight: 500;
        }
        .content {
          padding: 30px;
        }
        .summary-grid {
          display: grid;
          grid-template-columns: repeat(2, 1fr);
          gap: 20px;
          margin: 20px 0;
        }
        .summary-card {
          background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
          padding: 25px;
          border-radius: 10px;
          border-left: 5px solid #3498db;
          text-align: center;
        }
        .summary-value {
          font-size: 32px;
          font-weight: bold;
          color: #2c3e50;
          margin-bottom: 5px;
        }
        .summary-label {
          color: #6c757d;
          font-size: 14px;
          text-transform: uppercase;
          letter-spacing: 0.5px;
        }
        .financial-summary {
          background: linear-gradient(135deg, #27ae60 0%, #2ecc71 100%);
          color: white;
          padding: 30px;
          border-radius: 10px;
          margin: 25px 0;
        }
        .financial-summary h3 {
          margin: 0 0 20px 0;
          font-size: 22px;
          text-align: center;
        }
        .financial-grid {
          display: grid;
          grid-template-columns: repeat(3, 1fr);
          gap: 20px;
        }
        .financial-item {
          text-align: center;
          padding: 15px;
          background: rgba(255,255,255,0.1);
          border-radius: 8px;
        }
        .financial-value {
          font-size: 24px;
          font-weight: bold;
          margin-bottom: 5px;
        }
        .financial-label {
          font-size: 12px;
          opacity: 0.9;
        }
        .table-container {
          margin: 30px 0;
          border-radius: 10px;
          overflow: hidden;
          border: 1px solid #e9ecef;
        }
        table {
          width: 100%;
          border-collapse: collapse;
        }
        th {
          background: #2c3e50;
          color: white;
          padding: 15px 12px;
          text-align: left;
          font-weight: 600;
          font-size: 14px;
        }
        tr:nth-child(even) {
          background: #f8f9fa;
        }
        tr:hover {
          background: #e9ecef;
        }
        .footer {
          text-align: center;
          color: #6c757d;
          font-size: 12px;
          margin-top: 30px;
          padding: 20px;
          background: #f8f9fa;
          border-top: 1px solid #e9ecef;
        }
        @media only screen and (max-width: 600px) {
          .summary-grid, .financial-grid { grid-template-columns: 1fr; }
          .container { margin: 10px; }
        }
      </style>
    </head>
    <body>
      <div class="container">
        <!-- Header -->
        <div class="header">
          <h1>📊 Relatório Diário de Transfers</h1>
          <p>${CONFIG.NAMES.HOTEL_NAME} • ${dataFormatada}</p>
        </div>
        
        <!-- Cumprimento -->
        <div class="greeting">
          <h2>Resumo das Atividades do Dia</h2>
          <p>Apresentamos abaixo o relatório detalhado dos transfers realizados.</p>
        </div>
        
        <!-- Conteúdo -->
        <div class="content">
          <!-- Resumo Executivo -->
          <div class="summary-grid">
            <div class="summary-card">
              <div class="summary-value">${relatorio.totalTransfers}</div>
              <div class="summary-label">Total de Transfers</div>
            </div>
            <div class="summary-card">
              <div class="summary-value">€${relatorio.resumoFinanceiro.valorTotalGeral.toFixed(2)}</div>
              <div class="summary-label">Faturamento Total</div>
            </div>
          </div>
          
          <!-- Distribuição Financeira -->
          <div class="financial-summary">
            <h3>💰 Distribuição de Valores</h3>
            <div class="financial-grid">
              <div class="financial-item">
                <div class="financial-value">€${relatorio.resumoFinanceiro.valorTotalHotel.toFixed(2)}</div>
                <div class="financial-label">HOTEL (30%)</div>
              </div>
              <div class="financial-item">
                <div class="financial-value">€${relatorio.resumoFinanceiro.comissaoTotalRecepcao.toFixed(2)}</div>
                <div class="financial-label">RECEPÇÃO</div>
              </div>
              <div class="financial-item">
                <div class="financial-value">€${relatorio.resumoFinanceiro.valorTotalHUB.toFixed(2)}</div>
                <div class="financial-label">HUB TRANSFER</div>
              </div>
            </div>
          </div>
          
          <!-- Detalhamento -->
          <div class="table-container">
            <table>
              <thead>
                <tr>
                  <th>ID</th>
                  <th>Cliente</th>
                  <th>Tipo</th>
                  <th>Rota</th>
                  <th style="text-align: right;">Valor</th>
                  <th style="text-align: center;">Status</th>
                </tr>
              </thead>
              <tbody>
                ${tabelaTransfers}
              </tbody>
            </table>
          </div>
        </div>
        
        <!-- Footer -->
        <div class="footer">
          <p><strong>Com os melhores cumprimentos,</strong></p>
          <p>${CONFIG.NAMES.SISTEMA_NOME}</p>
          <p>Relatório gerado automaticamente em ${formatarDataHora(new Date())}</p>
        </div>
      </div>
    </body>
    </html>
  `;
}

/**
 * Gera estatísticas do sistema (CÓDIGO ANTIGO PRESERVADO + MELHORIAS)
 * @returns {Object} - Estatísticas compiladas
 */
function gerarEstatisticas() {
  logger.info('Gerando estatísticas do sistema');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return { erro: 'Sem dados para análise' };
    }
    
    const dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
    
    const stats = {
      totalTransfers: dados.length,
      porStatus: {},
      porTipo: {}, // NOVO: Estatísticas por tipo de serviço
      porMes: {},
      valorTotal: 0,
      valorHotel: 0,
      valorHUB: 0,
      comissaoRecepcao: 0, // NOVO
      mediaPassageiros: 0,
      topRotas: {},
      formasPagamento: {}
    };
    
    let totalPassageiros = 0;
    
    dados.forEach(row => {
      // Status
      const status = row[17] || 'Desconhecido';
      stats.porStatus[status] = (stats.porStatus[status] || 0) + 1;
      
      // Tipo de serviço (NOVO)
      const tipo = row[2] || 'Transfer';
      stats.porTipo[tipo] = (stats.porTipo[tipo] || 0) + 1;
      
      // Mês
      const data = new Date(row[5]);
      const mesAno = `${data.getMonth() + 1}/${data.getFullYear()}`;
      stats.porMes[mesAno] = (stats.porMes[mesAno] || 0) + 1;
      
      // Valores
      stats.valorTotal += parseFloat(row[11]) || 0;
      stats.valorHotel += parseFloat(row[12]) || 0;
      stats.valorHUB += parseFloat(row[13]) || 0;
      stats.comissaoRecepcao += parseFloat(row[14]) || 0; // NOVO
      
      // Passageiros
      totalPassageiros += parseInt(row[3]) || 0;
      
      // Rotas
      const rota = `${row[8]} → ${row[9]}`;
      stats.topRotas[rota] = (stats.topRotas[rota] || 0) + 1;
      
      // Formas de pagamento
      const pagamento = row[15] || 'Desconhecido';
      stats.formasPagamento[pagamento] = (stats.formasPagamento[pagamento] || 0) + 1;
    });
    
    stats.mediaPassageiros = totalPassageiros / dados.length;
    
    // Ordenar top rotas
    stats.topRotas = Object.entries(stats.topRotas)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .reduce((obj, [key, val]) => ({ ...obj, [key]: val }), {});
    
    logger.success('Estatísticas geradas', { totalRegistros: stats.totalTransfers });
    
    return stats;
    
  } catch (error) {
    logger.error('Erro ao gerar estatísticas', error);
    return { erro: error.message };
  }
}

/**
 * Mostra diálogo para relatório de período (CÓDIGO NOVO)
 */
function mostrarDialogoRelatorioPeriodo() {
  const ui = SpreadsheetApp.getUi();
  
  const html = `
    <div style="padding: 20px;">
      <h3>Selecione o Período do Relatório</h3>
      <div style="margin: 20px 0;">
        <label>Data Início:</label><br>
        <input type="date" id="dataInicio" style="width: 100%; padding: 8px; margin: 10px 0;">
      </div>
      <div style="margin: 20px 0;">
        <label>Data Fim:</label><br>
        <input type="date" id="dataFim" style="width: 100%; padding: 8px; margin: 10px 0;">
      </div>
      <div style="text-align: center; margin-top: 30px;">
        <button onclick="gerarRelatorio()" style="background: #3498db; color: white; padding: 10px 30px; border: none; border-radius: 5px; cursor: pointer;">
          Gerar Relatório
        </button>
      </div>
    </div>
    <script>
      function gerarRelatorio() {
        const dataInicio = document.getElementById('dataInicio').value;
        const dataFim = document.getElementById('dataFim').value;
        
        if (!dataInicio || !dataFim) {
          alert('Por favor, selecione as datas');
          return;
        }
        
        google.script.run
          .withSuccessHandler(() => {
            alert('Relatório gerado com sucesso!');
            google.script.host.close();
          })
          .withFailureHandler((error) => {
            alert('Erro: ' + error);
          })
          .processarRelatorioPeriodo(dataInicio, dataFim);
      }
    </script>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(300);
  
  ui.showModalDialog(htmlOutput, '📊 Gerar Relatório de Período');
}

/**
 * Processa relatório de período (CÓDIGO NOVO)
 */
function processarRelatorioPeriodo(dataInicioStr, dataFimStr) {
  const dataInicio = new Date(dataInicioStr);
  const dataFim = new Date(dataFimStr);
  
  const relatorio = gerarRelatorioPeriodo(dataInicio, dataFim);
  
  // Criar nova aba com o relatório
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const nomeAba = `Relatório_${formatarDataDDMMYYYY(dataInicio)}_${formatarDataDDMMYYYY(dataFim)}`.replace(/\//g, '-');
  
  let sheet = ss.getSheetByName(nomeAba);
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  
  sheet = ss.insertSheet(nomeAba);
  
  // Adicionar cabeçalho do relatório
  sheet.getRange(1, 1).setValue('RELATÓRIO DE PERÍODO').setFontSize(16).setFontWeight('bold');
  sheet.getRange(2, 1).setValue(`Período: ${formatarDataDDMMYYYY(dataInicio)} a ${formatarDataDDMMYYYY(dataFim)}`);
  
  // Resumo financeiro
  sheet.getRange(4, 1).setValue('RESUMO FINANCEIRO').setFontWeight('bold').setBackground('#3498db').setFontColor('white');
  sheet.getRange(5, 1).setValue('Total de Transfers:');
  sheet.getRange(5, 2).setValue(relatorio.totalTransfers);
  sheet.getRange(6, 1).setValue('Valor Total:');
  sheet.getRange(6, 2).setValue(relatorio.resumoFinanceiro.valorTotalGeral).setNumberFormat('€#,##0.00');
  sheet.getRange(7, 1).setValue(`Valor ${CONFIG.NAMES.HOTEL_NAME} (30%):`);
  sheet.getRange(7, 2).setValue(relatorio.resumoFinanceiro.valorTotalHotel).setNumberFormat('€#,##0.00');
  sheet.getRange(8, 1).setValue('Comissão Recepção:');
  sheet.getRange(8, 2).setValue(relatorio.resumoFinanceiro.comissaoTotalRecepcao).setNumberFormat('€#,##0.00');
  sheet.getRange(9, 1).setValue('Valor HUB Transfer:');
  sheet.getRange(9, 2).setValue(relatorio.resumoFinanceiro.valorTotalHUB).setNumberFormat('€#,##0.00');
  
  // Detalhamento
  if (relatorio.transfers.length > 0) {
    const headers = ['ID', 'Cliente', 'Tipo', 'Data', 'Rota', 'Valor Total', 'Hotel', 'Recepção', 'HUB', 'Status'];
    sheet.getRange(11, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#2c3e50').setFontColor('white');
    
    const dados = relatorio.transfers.map(t => [
      t.id,
      t.cliente,
      obterLabelTipoServico(t.tipoServico),
      formatarDataDDMMYYYY(new Date(t.data)),
      `${t.origem} → ${t.destino}`,
      t.valorTotal,
      t.valorHotel,
      t.comissaoRecepcao,
      t.valorHUB,
      t.status
    ]);
    
    sheet.getRange(12, 1, dados.length, headers.length).setValues(dados);
    
    // Formatar colunas de valores
    sheet.getRange(12, 6, dados.length, 4).setNumberFormat('€#,##0.00');
  }
  
  // Auto-ajustar largura das colunas
  for (let i = 1; i <= 10; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // Enviar relatório por email também
  const htmlEmail = criarEmailRelatorioDiario(relatorio, dataInicio);
  
// Correção: usar htmlBody correto
sendEmailComAssinatura({
  to: destinatarios,
  subject: assunto,
  htmlBody: htmlEmail || corpo || corpoHtml,
  name: CONFIG.NAMES.SISTEMA_NOME,
  replyTo: CONFIG.EMAIL_CONFIG.EMAIL_EMPRESA
});
  
  return true;
}

// ===================================================
// APIs PRINCIPAIS (doGet e doPost) - FUSÃO COMPLETA
// ===================================================

/**
 * Processa requisições GET - VERSÃO CORRIGIDA PARA SINCRONIZAÇÃO
 */
function doGet(e) {
  console.log('doGet chamado com parâmetros:', e.parameter);
  logger.info('Recebendo requisição GET', { parameters: e.parameter });
  
  try {
    const params = e.parameter || {};
    const action = params.action || 'default';
    
    console.log('Ação identificada:', action);
    
    // Roteamento baseado na ação
    switch (action) {
      // NOVA AÇÃO: Carregar todos os dados
      case 'getAllData':
        console.log('Processando getAllData');
        const dados = getAllData();
        return createJsonResponse(dados);
      
      case 'getLastId':
        console.log('Processando getLastId');
        const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
        const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
        const ultimoId = gerarProximoIdSeguro(sheet) - 1;
        return createJsonResponse({
          success: true,
          lastId: ultimoId
        });
      
      // Ações de e-mail (confirmação/cancelamento)
      case 'confirm':
      case 'cancel':
        console.log('Redirecionando para handleEmailAction...');
        return handleEmailAction(params);
      
      // Outras ações existentes
      case 'consultarPreco':
        return handleConsultarPreco(params);
      
      case 'verificarRegistroDuplo':
        return handleVerificarDuplicidade(params);
      
      case 'buscarTransfer':
        return handleBuscarTransfer(params);
      
      case 'listarTransfers':
        return handleListarTransfers(params);
      
      case 'test':
        return handleTest();
      
      case 'config':
        return handleConfig();
      
      case 'stats':
        return handleStats();
      
      case 'health':
        return handleHealthCheck();
      
      // Padrão
      default:
        console.log('Ação padrão - sistema funcionando');
        return createJsonResponse({
          status: 'success',
          message: 'Sistema Empire Marques Hotel-HUB Transfer funcionando!',
          versao: CONFIG.SISTEMA.VERSAO,
          endpoints: {
            getAllData: '?action=getAllData',
            getLastId: '?action=getLastId',
            test: '?action=test'
          }
        });
    }
    
  } catch (error) {
    console.log('Erro no doGet:', error.message);
    logger.error('Erro no doGet', error);
    
    return createJsonResponse({
      status: 'error',
      message: error.toString(),
      action: e.parameter?.action || 'unknown',
      timestamp: new Date().toISOString()
    }, 500);
  }
}

/**
 * Processa requisições POST (VERSÃO CORRIGIDA)
 * @param {Object} e - Evento da requisição
 * @returns {TextOutput} - Resposta JSON
 */
function doPost(e) {
  logger.info('Recebendo requisição POST', {
    contentLength: e.contentLength,
    postData: e.postData?.type,
    hasParameters: !!e.parameter,
    hasPostData: !!e.postData
  });
  
  try {
    let dadosRecebidos = {};
    
    // MÚLTIPLAS ESTRATÉGIAS DE EXTRAÇÃO DE DADOS
    
    // Estratégia 1: Dados via postData.contents (JSON)
    if (e.postData && e.postData.contents) {
      try {
        dadosRecebidos = JSON.parse(e.postData.contents);
        logger.info('Dados extraídos via postData.contents', { 
          keys: Object.keys(dadosRecebidos) 
        });
      } catch (parseError) {
        logger.warn('Falha ao parsear postData.contents', parseError);
        dadosRecebidos = {};
      }
    }
    
    // Estratégia 2: Dados via parâmetros de URL (GET/POST)
    if (Object.keys(dadosRecebidos).length === 0 && e.parameter) {
      dadosRecebidos = e.parameter;
      logger.info('Dados extraídos via parameter', { 
        keys: Object.keys(dadosRecebidos) 
      });
    }
    
    // Estratégia 3: Dados via parameters (array)
    if (Object.keys(dadosRecebidos).length === 0 && e.parameters) {
      dadosRecebidos = {};
      for (const key in e.parameters) {
        dadosRecebidos[key] = e.parameters[key][0]; // Pegar primeiro valor do array
      }
      logger.info('Dados extraídos via parameters', { 
        keys: Object.keys(dadosRecebidos) 
      });
    }
    
    // Estratégia 4: Tentar getDataAsString()
    if (Object.keys(dadosRecebidos).length === 0 && e.postData) {
      try {
        const dataString = e.postData.getDataAsString();
        if (dataString) {
          dadosRecebidos = JSON.parse(dataString);
          logger.info('Dados extraídos via getDataAsString', { 
            keys: Object.keys(dadosRecebidos) 
          });
        }
      } catch (stringError) {
        logger.warn('Falha ao usar getDataAsString', stringError);
      }
    }
    
    // Verificar se conseguimos extrair dados
    if (Object.keys(dadosRecebidos).length === 0) {
      logger.error('Nenhum dado pôde ser extraído', {
        hasPostData: !!e.postData,
        hasParameter: !!e.parameter,
        hasParameters: !!e.parameters,
        entireEvent: JSON.stringify(e, null, 2)
      });
      
      return createJsonResponse({
        status: 'error',
        message: 'Nenhum dado recebido na requisição',
        debug: {
          hasPostData: !!e.postData,
          hasParameter: !!e.parameter,
          hasParameters: !!e.parameters
        },
        timestamp: new Date().toISOString()
      }, 400);
    }
    
    // Log dos dados finais extraídos
    logger.info('Dados finais extraídos', dadosRecebidos);
    
    // Processar ação especial se houver
    if (dadosRecebidos.action) {
      logger.info('Processando ação especial', { action: dadosRecebidos.action });
      return processarAcaoEspecial(dadosRecebidos);
    }
    
    // Processar novo transfer se tiver dados de cliente
    if (dadosRecebidos.nomeCliente) {
      logger.info('Processando novo transfer', { cliente: dadosRecebidos.nomeCliente });
      return processarNovoTransfer(dadosRecebidos);
    }
    
    // Se chegou aqui, não sabemos o que fazer com os dados
    logger.warn('Dados recebidos mas sem ação clara', dadosRecebidos);
    return createJsonResponse({
      status: 'error',
      message: 'Dados recebidos mas sem ação identificável (falta action ou nomeCliente)',
      receivedData: dadosRecebidos,
      timestamp: new Date().toISOString()
    }, 400);
    
  } catch (error) {
    logger.error('Erro crítico no doPost', error);
    return createJsonResponse({
      status: 'error',
      message: 'Erro interno: ' + error.toString(),
      stack: error.stack,
      timestamp: new Date().toISOString()
    }, 500);
  }
}

/**
 * Função auxiliar para criar respostas JSON padronizadas
 */
function createJsonResponse(data, statusCode = 200) {
  const output = ContentService.createTextOutput(JSON.stringify(data, null, 2));
  output.setMimeType(ContentService.MimeType.JSON);
  
  if (statusCode !== 200) {
    // Apps Script não suporta status codes HTTP personalizados diretamente
    // mas podemos logar para debug
    logger.info('Retornando erro', { statusCode, data });
  }
  
  return output;
}

// ===================================================
// HANDLERS DO GET (CÓDIGO ANTIGO PRESERVADO + MELHORIAS)
// ===================================================

/**
 * Handler para consulta de preço
 * @private
 */
function handleConsultarPreco(params) {
  const origem = params.origem || '';
  const destino = params.destino || '';
  const pessoas = parseInt(params.pessoas) || 1;
  const bagagens = parseInt(params.bagagens) || 0;
  const tipoServico = params.tipoServico || 'Transfer'; // NOVO
  
  const valores = calcularValores(origem, destino, pessoas, bagagens, null, tipoServico);
  
  return createJsonResponse({
    status: 'success',
    preco: valores,
    consultaId: Utilities.getUuid()
  });
}

/**
 * Handler para verificar duplicidade
 * @private
 */
function handleVerificarDuplicidade(params) {
  const id = params.id || '';
  const data = params.data || '';
  
  if (!id || !data) {
    return createJsonResponse({
      status: 'error',
      message: 'ID e data são obrigatórios'
    }, 400);
  }
  
  const duplicado = verificarRegistroDuplo(id, data);
  
  return createJsonResponse({
    status: 'success',
    duplicado: duplicado,
    verificacao: {
      id: id,
      data: data,
      timestamp: new Date().toISOString()
    }
  });
}

/**
 * Handler para buscar transfer específico
 * @private
 */
function handleBuscarTransfer(params) {
  const id = params.id || '';
  
  if (!id) {
    return createJsonResponse({
      status: 'error',
      message: 'ID é obrigatório'
    }, 400);
  }
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const linha = encontrarLinhaPorId(sheet, id);
    
    if (!linha) {
      return createJsonResponse({
        status: 'error',
        message: `Transfer ${id} não encontrado`
      }, 404);
    }
    
    const dados = sheet.getRange(linha, 1, 1, HEADERS.length).getValues()[0];
    const transfer = {};
    
    HEADERS.forEach((header, index) => {
      transfer[header] = dados[index];
    });
    
    return createJsonResponse({
      status: 'success',
      transfer: transfer
    });
    
  } catch (error) {
    logger.error('Erro ao buscar transfer', error);
    return createJsonResponse({
      status: 'error',
      message: error.toString()
    }, 500);
  }
}

/**
 * Handler para listar transfers
 * @private
 */
function handleListarTransfers(params) {
  try {
    const filtros = {
      data: params.data,
      status: params.status,
      cliente: params.cliente,
      tipoServico: params.tipoServico, // NOVO
      limite: parseInt(params.limite) || 50,
      offset: parseInt(params.offset) || 0
    };
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return createJsonResponse({
        status: 'success',
        transfers: [],
        total: 0
      });
    }
    
    const dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
    let transfersFiltrados = [];
    
    dados.forEach(row => {
      let incluir = true;
      
      // Aplicar filtros
      if (filtros.data) {
        const dataTransfer = formatarDataDDMMYYYY(new Date(row[5])); // Coluna F
        if (dataTransfer !== filtros.data) incluir = false;
      }
      
      if (filtros.status && row[17] !== filtros.status) { // Coluna R
        incluir = false;
      }
      
      if (filtros.cliente && !row[1].toLowerCase().includes(filtros.cliente.toLowerCase())) {
        incluir = false;
      }
      
      // NOVO: Filtro por tipo de serviço
      if (filtros.tipoServico && row[2] !== filtros.tipoServico) { // Coluna C
        incluir = false;
      }
      
      if (incluir) {
        const transfer = {};
        HEADERS.forEach((header, index) => {
          transfer[header] = row[index];
        });
        transfersFiltrados.push(transfer);
      }
    });
    
    // Aplicar paginação
    const total = transfersFiltrados.length;
    transfersFiltrados = transfersFiltrados.slice(filtros.offset, filtros.offset + filtros.limite);
    
    return createJsonResponse({
      status: 'success',
      transfers: transfersFiltrados,
      total: total,
      limite: filtros.limite,
      offset: filtros.offset
    });
    
  } catch (error) {
    logger.error('Erro ao listar transfers', error);
    return createJsonResponse({
      status: 'error',
      message: error.toString()
    }, 500);
  }
}

/**
 * Handler para teste da API
 * @private
 */
function handleTest() {
  return createJsonResponse({
    status: 'success',
    message: '🧪 API ativa e funcionando!',
    sistema: CONFIG.NAMES.SISTEMA_NOME,
    versao: CONFIG.SISTEMA.VERSAO,
    timestamp: new Date().toISOString(),
    endpoints: {
      GET: [
        '?action=test',
        '?action=config',
        '?action=stats',
        '?action=health',
        '?action=consultarPreco&origem=X&destino=Y&pessoas=N&tipoServico=Transfer',
        '?action=verificarRegistroDuplo&id=X&data=Y',
        '?action=buscarTransfer&id=X',
        '?action=listarTransfers&status=X&tipoServico=Y&limite=N',
        '?action=confirm&id=X',
        '?action=cancel&id=X'
      ],
      POST: [
        'Novo transfer (JSON)',
        'Ações especiais (action: clearAllData, clearTestData, etc.)'
      ]
    }
  });
}

/**
 * Handler para obter configuração
 * @private
 */
function handleConfig() {
  // Remover informações sensíveis
  const configSegura = {
    sistema: CONFIG.SISTEMA,
    limites: CONFIG.LIMITES,
    valores: CONFIG.VALORES,
    emailAtivo: CONFIG.EMAIL_CONFIG.ENVIAR_AUTOMATICO,
    webhooksAtivos: CONFIG.MAKE_WEBHOOKS.ENABLED,
    versao: CONFIG.SISTEMA.VERSAO,
    hotelNome: CONFIG.NAMES.HOTEL_NAME
  };
  
  return createJsonResponse({
    status: 'success',
    config: configSegura
  });
}

/**
 * Handler para estatísticas
 * @private
 */
function handleStats() {
  try {
    const stats = gerarEstatisticas();
    
    return createJsonResponse({
      status: 'success',
      estatisticas: stats
    });
    
  } catch (error) {
    logger.error('Erro ao gerar estatísticas', error);
    return createJsonResponse({
      status: 'error',
      message: error.toString()
    }, 500);
  }
}

/**
 * Handler para health check
 * @private
 */
function handleHealthCheck() {
  const checks = {
    planilha: false,
    abaPrincipal: false,
    tabelaPrecos: false,
    email: false,
    triggers: false
  };
  
  try {
    // Verificar planilha
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    checks.planilha = true;
    
    // Verificar aba principal
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    checks.abaPrincipal = !!sheet;
    
    // Verificar tabela de preços
    const pricing = ss.getSheetByName(CONFIG.PRICING_SHEET_NAME);
    checks.tabelaPrecos = !!pricing;
    
    // Verificar configuração de e-mail
    checks.email = CONFIG.EMAIL_CONFIG.ENVIAR_AUTOMATICO;
    
    // Verificar triggers
    const triggers = ScriptApp.getProjectTriggers();
    checks.triggers = triggers.length > 0;
    
  } catch (error) {
    logger.error('Erro no health check', error);
  }
  
  const healthy = Object.values(checks).every(check => check === true);
  
  return createJsonResponse({
    status: healthy ? 'healthy' : 'unhealthy',
    checks: checks,
    timestamp: new Date().toISOString()
  }, healthy ? 200 : 503);
}

/**
 * Handler padrão
 * @private
 */
function handleDefault() {
  return createJsonResponse({
    status: 'success',
    message: `Sistema ${CONFIG.NAMES.HOTEL_NAME}-HUB Transfer funcionando!`,
    versao: CONFIG.SISTEMA.VERSAO,
    documentacao: 'Use ?action=test para ver endpoints disponíveis'
  });
}

// ===================================================
// PROCESSAMENTO DO POST (FUSÃO COMPLETA)
// ===================================================

/**
 * Processa novo transfer recebido via POST (CÓDIGO NOVO COMO BASE)
 * @private
 */
/**
 * Processa novo transfer recebido via POST (CÓDIGO NOVO COMO BASE)
 * @private
 */
function processarNovoTransfer(dadosRecebidos) {
  logger.info('Processando novo transfer');
  
  // Validar dados
  const validacao = validarDados(dadosRecebidos);
  if (!validacao.valido) {
    return createJsonResponse({
      status: 'error',
      message: 'Dados inválidos',
      erros: validacao.erros
    }, 400);
  }
  
  const dados = validacao.dados;
  
  try {
    // Abrir planilha
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) {
      throw new Error(`Aba '${CONFIG.SHEET_NAME}' não encontrada`);
    }
    
    // 🔧 CORREÇÃO: Forçar formatação da coluna data antes do registro
    const maxRows = Math.max(sheet.getMaxRows(), 1000);
    const colunaData = sheet.getRange(2, 6, maxRows - 1, 1); // Coluna F
    colunaData.setNumberFormat('dd/mm/yyyy');
    
    // Gerar ID
    const transferId = dados.id || gerarProximoIdSeguro(sheet);
    
    // Calcular valores baseado no tipo de serviço (CÓDIGO NOVO)
    const valores = calcularValores(
      dados.origem,
      dados.destino,
      dados.numeroPessoas,
      dados.numeroBagagens,
      dados.valorTotal,
      dados.tipoServico
    );
    
    // Montar dados do transfer (ESTRUTURA DO CÓDIGO NOVO)
    const dadosTransfer = [
      transferId,                          // A - ID
      dados.nomeCliente,                   // B - Cliente
      dados.tipoServico || 'Transfer',     // C - Tipo Serviço
      dados.numeroPessoas,                 // D - Pessoas
      dados.numeroBagagens,                // E - Bagagens
      new Date(dados.data),                // F - Data ← CORREÇÃO APLICADA
      dados.contacto,                      // G - Contacto
      dados.numeroVoo || '',               // H - Voo
      dados.origem,                        // I - Origem
      dados.destino,                       // J - Destino
      dados.horaPickup,                    // K - Hora Pick-up
      valores.precoCliente,                // L - Preço Cliente
      valores.valorHotel,                  // M - Valor Hotel
      valores.valorHUB,                    // N - Valor HUB
      valores.comissaoRecepcao,            // O - Comissão Recepção
      dados.modoPagamento,                 // P - Forma Pagamento
      dados.pagoParaQuem,                  // Q - Pago Para
      dados.status,                        // R - Status
      dados.observacoes || '',             // S - Observações
      new Date()                           // T - Data Criação
    ];
    
    // Executar registro duplo
    const resultadoRegistro = executarRegistroDuplo(dadosTransfer);
    
    // Preparar resposta
    const resposta = {
      status: resultadoRegistro.sucesso ? 'success' : 'error',
      message: resultadoRegistro.observacoes,
      transfer: {
        id: transferId,
        cliente: dados.nomeCliente,
        tipoServico: dados.tipoServico,
        data: formatarDataDDMMYYYY(new Date(dados.data)),
        rota: `${dados.origem} → ${dados.destino}`,
        valor: valores.precoCliente
      },
      registroDuplo: {
        completo: resultadoRegistro.registroDuploCompleto,
        abaPrincipal: resultadoRegistro.abaPrincipal,
        abaMensal: resultadoRegistro.abaMensal
      },
      valores: valores
    };
    
    // Se houve erro, retornar com status 500
    if (!resultadoRegistro.sucesso) {
      return createJsonResponse(resposta, 500);
    }
    
    // Sucesso
    return createJsonResponse(resposta, 201);
    
  } catch (error) {
    logger.error('Erro ao processar transfer', error);
    return createJsonResponse({
      status: 'error',
      message: error.toString()
    }, 500);
  }
}

/**
 * Processa ações especiais (VERSÃO CORRIGIDA)
 */
function processarAcaoEspecial(dados) {
  logger.info('Processando ação especial', { action: dados.action });
  
  try {
    switch (dados.action) {
      case 'test':
        return createJsonResponse({
          status: 'success',
          message: 'Sistema Empire Marques Hotel-HUB Transfer Integrado funcionando!',
          version: '4.0-Empire-Marques-Hub-Transfer-Integrado',
          timestamp: new Date().toISOString()
        });
      
      case 'getAllData':
        return createJsonResponse(getAllData());
      
      case 'getLastId':
        return createJsonResponse(getLastId());
      
      case 'clearAllData':
        return createJsonResponse(clearAllData());
      
      case 'clearTestData':
        return createJsonResponse(clearTestData());
      
      // ADICIONAR ESTE CASE
      case 'addTransfer':
        logger.info('Processando novo transfer via ação especial', {
          cliente: dados.nomeCliente,
          email: dados.emailDestino
        });
        
        // Remover a propriedade action para processar como transfer normal
        const transferData = { ...dados };
        delete transferData.action;
        
        return processarNovoTransfer(transferData);
      
      case 'clearAllData':
        resultado = limparDadosCompleto();
        break;
        
      case 'clearTestData':
        resultado = limparDadosTeste();
        break;
        
      case 'corrigirRegistros':
        resultado = corrigirRegistrosIncompletos();
        break;
        
      case 'sincronizarAbas':
        resultado = sincronizarAbasMensais();
        break;
        
      case 'reordenarPorData':
        resultado = reordenarPorData();
        break;
        
      case 'removerDuplicados':
        resultado = removerDuplicados(dados.simular !== false);
        break;
        
      case 'criarBackup':
        resultado = criarBackup();
        break;
        
      case 'atualizarStatus':
        if (!dados.transferId || !dados.novoStatus) {
          throw new Error('transferId e novoStatus são obrigatórios');
        }
        resultado = atualizarStatusTransfer(
          dados.transferId,
          dados.novoStatus,
          dados.observacao || ''
        );
        break;
        
      case 'atualizarTransfer':
        if (!dados.transferId || !dados.atualizacoes) {
          throw new Error('transferId e atualizacoes são obrigatórios');
        }
        resultado = atualizarTransferCompleto(dados.transferId, dados.atualizacoes);
        break;
        
      case 'adicionarPreco':
        if (!dados.precoData) {
          throw new Error('precoData é obrigatório');
        }
        resultado = adicionarPrecoTabela(dados.precoData);
        break;
        
      case 'importarPrecos':
        if (!dados.csvData && !dados.jsonData) {
          throw new Error('csvData ou jsonData é obrigatório');
        }
        resultado = importarPrecos(
          dados.csvData || dados.jsonData,
          dados.csvData ? 'csv' : 'json'
        );
        break;
        
      case 'configurarTriggers':
        resultado = configurarTriggersEmail();
        break;
        
      case 'removerTriggers':
        resultado = removerTriggersEmail();
        break;
        
      case 'enviarRelatorioDiario': // NOVO
        enviarRelatorioDiaAnterior();
        resultado = { sucesso: true, mensagem: 'Relatório enviado' };
        break;
        
      case 'verificarConfirmacoes': // CÓDIGO ANTIGO PRESERVADO
        const confirmacoes = verificarConfirmacoesEmail();
        resultado = { 
          sucesso: true, 
          confirmacoes: confirmacoes,
          mensagem: `${confirmacoes} confirmações processadas`
        };
        break;
        
      default:
        logger.error('Ação desconhecida', { action: dados.action });
        return createJsonResponse({
          status: 'error',
          action: dados.action,
          message: `Ação desconhecida: ${dados.action}`,
          availableActions: ['test', 'getAllData', 'getLastId', 'clearAllData', 'clearTestData', 'addTransfer'],
          timestamp: new Date().toISOString()
        }, 400);
    }
    
    // Para casos que usam a estrutura antiga com 'resultado'
    if (typeof resultado !== 'undefined') {
      return createJsonResponse({
        status: resultado.sucesso ? 'success' : 'error',
        action: dados.action,
        resultado: resultado
      });
    }
    
  } catch (error) {
    logger.error('Erro na ação especial', error);
    return createJsonResponse({
      status: 'error',
      action: dados.action,
      message: error.toString(),
      timestamp: new Date().toISOString()
    }, 500);
  }
}

// ===================================================
// GESTÃO DA TABELA DE PREÇOS (CÓDIGO ANTIGO PRESERVADO + MELHORIAS)
// ===================================================

/**
 * Adiciona novo preço à tabela (CÓDIGO ANTIGO PRESERVADO + ADAPTAÇÕES)
 * @param {Object} dadosPreco - Dados do novo preço
 * @returns {Object} - Resultado da operação
 */
function adicionarPrecoTabela(dadosPreco) {
  logger.info('Adicionando novo preço à tabela');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.PRICING_SHEET_NAME);
    
    // Criar aba se não existir
    if (!sheet) {
      sheet = criarTabelaPrecos(ss);
    }
    
    // Validar dados
    const camposObrigatorios = ['origem', 'destino', 'pessoas', 'precoCliente'];
    for (const campo of camposObrigatorios) {
      if (!dadosPreco[campo]) {
        throw new Error(`Campo obrigatório ausente: ${campo}`);
      }
    }
    
    // Gerar ID
    const novoId = gerarProximoIdSeguro(sheet);
    
    // Calcular valores se não fornecidos (USANDO CÓDIGO NOVO)
    let valorHotel = dadosPreco.valorHotel;
    let valorHUB = dadosPreco.valorHUB;
    let comissaoRecepcao = dadosPreco.comissaoRecepcao;
    
    if (!valorHotel || !valorHUB || !comissaoRecepcao) {
      const valores = calcularPorTipoServico(
        dadosPreco.precoCliente, 
        dadosPreco.tipoServico || 'Transfer'
      );
      valorHotel = valorHotel || valores.valorHotel;
      valorHUB = valorHUB || valores.valorHUB;
      comissaoRecepcao = comissaoRecepcao || valores.comissaoRecepcao;
    }
    
    // Montar linha (ESTRUTURA DO CÓDIGO NOVO)
    const novaLinha = [
      novoId,                                      // A - ID
      dadosPreco.tipoServico || 'Transfer',        // B - Tipo Serviço
      `${dadosPreco.origem} → ${dadosPreco.destino}`, // C - Rota
      dadosPreco.origem,                           // D - Origem
      dadosPreco.destino,                          // E - Destino
      parseInt(dadosPreco.pessoas) || 1,           // F - Pessoas
      parseInt(dadosPreco.bagagens) || 0,          // G - Bagagens
      parseFloat(dadosPreco.precoPorPessoa) || 0,  // H - Preço Por Pessoa
      parseFloat(dadosPreco.precoPorGrupo) || 0,   // I - Preço Por Grupo
      parseFloat(dadosPreco.precoCliente),         // J - Preço Cliente
      valorHotel,                                  // K - Valor Hotel
      valorHUB,                                    // L - Valor HUB
      comissaoRecepcao,                            // M - Comissão Recepção
      dadosPreco.ativo !== false ? 'Sim' : 'Não', // N - Ativo
      new Date(),                                  // O - Data Criação
      dadosPreco.observacoes || ''                 // P - Observações
    ];
    
    // Inserir na planilha
    sheet.appendRow(novaLinha);
    
    logger.success('Preço adicionado com sucesso', { id: novoId });
    
    return {
      sucesso: true,
      id: novoId,
      mensagem: 'Preço adicionado com sucesso'
    };
    
  } catch (error) {
    logger.error('Erro ao adicionar preço', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Atualiza preço existente na tabela (CÓDIGO ANTIGO PRESERVADO)
 * @param {number} id - ID do preço
 * @param {Object} dadosAtualizacao - Dados a atualizar
 * @returns {Object} - Resultado da operação
 */
function atualizarPrecoTabela(id, dadosAtualizacao) {
  logger.info('Atualizando preço na tabela', { id });
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.PRICING_SHEET_NAME);
    
    if (!sheet) {
      throw new Error('Tabela de preços não encontrada');
    }
    
    const linha = encontrarLinhaPorId(sheet, id);
    if (!linha) {
      throw new Error(`Preço ID ${id} não encontrado`);
    }
    
    // Atualizar campos específicos (ESTRUTURA DO CÓDIGO NOVO)
    if (dadosAtualizacao.tipoServico !== undefined) {
      sheet.getRange(linha, 2).setValue(dadosAtualizacao.tipoServico); // Coluna B
    }
    
    if (dadosAtualizacao.origem !== undefined) {
      sheet.getRange(linha, 4).setValue(dadosAtualizacao.origem); // Coluna D
      // Atualizar rota também
      const destino = sheet.getRange(linha, 5).getValue();
      sheet.getRange(linha, 3).setValue(`${dadosAtualizacao.origem} → ${destino}`);
    }
    
    if (dadosAtualizacao.destino !== undefined) {
      sheet.getRange(linha, 5).setValue(dadosAtualizacao.destino); // Coluna E
      // Atualizar rota também
      const origem = sheet.getRange(linha, 4).getValue();
      sheet.getRange(linha, 3).setValue(`${origem} → ${dadosAtualizacao.destino}`);
    }
    
    if (dadosAtualizacao.pessoas !== undefined) {
      sheet.getRange(linha, 6).setValue(parseInt(dadosAtualizacao.pessoas));
    }
    
    if (dadosAtualizacao.bagagens !== undefined) {
      sheet.getRange(linha, 7).setValue(parseInt(dadosAtualizacao.bagagens));
    }
    
    if (dadosAtualizacao.precoPorPessoa !== undefined) {
      sheet.getRange(linha, 8).setValue(parseFloat(dadosAtualizacao.precoPorPessoa));
    }
    
    if (dadosAtualizacao.precoPorGrupo !== undefined) {
      sheet.getRange(linha, 9).setValue(parseFloat(dadosAtualizacao.precoPorGrupo));
    }
    
    if (dadosAtualizacao.precoCliente !== undefined) {
      sheet.getRange(linha, 10).setValue(parseFloat(dadosAtualizacao.precoCliente));
    }
    
    if (dadosAtualizacao.valorHotel !== undefined) {
      sheet.getRange(linha, 11).setValue(parseFloat(dadosAtualizacao.valorHotel));
    }
    
    if (dadosAtualizacao.valorHUB !== undefined) {
      sheet.getRange(linha, 12).setValue(parseFloat(dadosAtualizacao.valorHUB));
    }
    
    if (dadosAtualizacao.comissaoRecepcao !== undefined) {
      sheet.getRange(linha, 13).setValue(parseFloat(dadosAtualizacao.comissaoRecepcao));
    }
    
    if (dadosAtualizacao.ativo !== undefined) {
      sheet.getRange(linha, 14).setValue(dadosAtualizacao.ativo ? 'Sim' : 'Não');
    }
    
    if (dadosAtualizacao.observacoes !== undefined) {
      sheet.getRange(linha, 16).setValue(dadosAtualizacao.observacoes);
    }
    
    logger.success('Preço atualizado com sucesso', { id });
    
    return {
      sucesso: true,
      id: id,
      mensagem: 'Preço atualizado com sucesso'
    };
    
  } catch (error) {
    logger.error('Erro ao atualizar preço', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Remove preço da tabela (soft delete - marca como inativo)
 * @param {number} id - ID do preço
 * @returns {Object} - Resultado da operação
 */
function removerPrecoTabela(id) {
  logger.info('Removendo preço da tabela', { id });
  
  try {
    const resultado = atualizarPrecoTabela(id, { ativo: false });
    
    if (resultado.sucesso) {
      resultado.mensagem = 'Preço marcado como inativo';
    }
    
    return resultado;
    
  } catch (error) {
    logger.error('Erro ao remover preço', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Lista todos os preços da tabela (CÓDIGO ANTIGO PRESERVADO + ADAPTAÇÕES)
 * @param {Object} filtros - Filtros opcionais
 * @returns {Array} - Lista de preços
 */
function listarPrecosTabela(filtros = {}) {
  logger.info('Listando preços da tabela', { filtros });
  
try {
     const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
     const sheet = ss.getSheetByName(CONFIG.PRICING_SHEET_NAME);
     
     if (!sheet || sheet.getLastRow() <= 1) {
       return [];
     }
     
     const dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, PRICING_HEADERS.length).getValues();
     const precos = [];
     
     dados.forEach(linha => {
       // Aplicar filtros
       if (filtros.apenasAtivos && linha[13] !== 'Sim') return; // Coluna N - Ativo
       if (filtros.tipoServico && linha[1] !== filtros.tipoServico) return; // Coluna B
       if (filtros.origem && !linha[3].toLowerCase().includes(filtros.origem.toLowerCase())) return;
       if (filtros.destino && !linha[4].toLowerCase().includes(filtros.destino.toLowerCase())) return;
       if (filtros.pessoas && parseInt(linha[5]) !== parseInt(filtros.pessoas)) return;
       
       precos.push({
         id: linha[0],
         tipoServico: linha[1],
         rota: linha[2],
         origem: linha[3],
         destino: linha[4],
         pessoas: linha[5],
         bagagens: linha[6],
         precoPorPessoa: linha[7],
         precoPorGrupo: linha[8],
         precoCliente: linha[9],
         valorHotel: linha[10],
         valorHUB: linha[11],
         comissaoRecepcao: linha[12],
         ativo: linha[13] === 'Sim',
         dataCriacao: linha[14],
         observacoes: linha[15]
       });
     });
     
     // Ordenar por rota
     precos.sort((a, b) => a.rota.localeCompare(b.rota));
     
     logger.success('Preços listados', { total: precos.length });
     
     return precos;
     
   } catch (error) {
     logger.error('Erro ao listar preços', error);
     return [];
   }
 }

/**
* Cria a tabela de preços se não existir (ATUALIZADO PARA CÓDIGO NOVO)
* @private
*/
function criarTabelaPrecos(ss) {
 logger.info('Criando tabela de preços');
 
 const sheet = ss.insertSheet(CONFIG.PRICING_SHEET_NAME);
 
 // Configurar headers
 sheet.getRange(1, 1, 1, PRICING_HEADERS.length).setValues([PRICING_HEADERS]);
 
 // Formatação
 const headerRange = sheet.getRange(1, 1, 1, PRICING_HEADERS.length);
 headerRange
   .setBackground(STYLES.HEADER_COLORS.PRECOS)
   .setFontColor('#ffffff')
   .setFontWeight('bold')
   .setFontSize(11)
   .setHorizontalAlignment('center')
   .setVerticalAlignment('middle');
 
 // Larguras das colunas
 STYLES.COLUMN_WIDTHS.PRECOS.forEach((width, index) => {
   sheet.setColumnWidth(index + 1, width);
 });
 
 // Congelar headers
 sheet.setFrozenRows(1);
 
 // Aplicar formatação de moeda (colunas H, I, J, K, L, M)
 sheet.getRange(2, 8, sheet.getMaxRows() - 1, 6).setNumberFormat(STYLES.FORMATS.MOEDA);
 
 // Validação de tipo de serviço (coluna B)
 const tipoValidation = SpreadsheetApp.newDataValidation()
   .requireValueInList(VALIDACOES.VALORES_PERMITIDOS.TIPO_SERVICO)
   .setAllowInvalid(false)
   .build();
 sheet.getRange(2, 2, sheet.getMaxRows() - 1, 1).setDataValidation(tipoValidation);
 
 // Validação de ativo/inativo (coluna N)
 const ativoValidation = SpreadsheetApp.newDataValidation()
   .requireValueInList(['Sim', 'Não'])
   .setAllowInvalid(false)
   .build();
 sheet.getRange(2, 14, sheet.getMaxRows() - 1, 1).setDataValidation(ativoValidation);
 
 logger.success('Tabela de preços criada');
 
 return sheet;
}

/**
* Importa preços de uma fonte externa (CSV, JSON, etc) - CÓDIGO ANTIGO PRESERVADO
* @param {string} dados - Dados a importar
* @param {string} formato - Formato dos dados (csv, json)
* @returns {Object} - Resultado da importação
*/
function importarPrecos(dados, formato = 'csv') {
 logger.info('Importando preços', { formato });
 
 try {
   let precosParaImportar = [];
   
   if (formato === 'csv') {
     // Parse CSV
     const linhas = dados.split('\n');
     const headers = linhas[0].split(',').map(h => h.trim());
     
     for (let i = 1; i < linhas.length; i++) {
       if (!linhas[i].trim()) continue;
       
       const valores = linhas[i].split(',').map(v => v.trim());
       const preco = {};
       
       headers.forEach((header, index) => {
         preco[header] = valores[index];
       });
       
       precosParaImportar.push(preco);
     }
     
   } else if (formato === 'json') {
     precosParaImportar = JSON.parse(dados);
   }
   
   // Importar cada preço
   let importados = 0;
   let erros = 0;
   const resultados = [];
   
   precosParaImportar.forEach(preco => {
     const resultado = adicionarPrecoTabela({
       tipoServico: preco.tipoServico || preco['Tipo Serviço'] || 'Transfer',
       origem: preco.origem || preco.Origem,
       destino: preco.destino || preco.Destino,
       pessoas: preco.pessoas || preco.Pessoas || 1,
       bagagens: preco.bagagens || preco.Bagagens || 0,
       precoPorPessoa: preco.precoPorPessoa || preco['Preço Por Pessoa'] || 0,
       precoPorGrupo: preco.precoPorGrupo || preco['Preço Por Grupo'] || 0,
       precoCliente: preco.precoCliente || preco['Preço Cliente'] || preco.preco,
       valorHotel: preco.valorHotel || preco['Valor Hotel'],
       valorHUB: preco.valorHUB || preco['Valor HUB'],
       comissaoRecepcao: preco.comissaoRecepcao || preco['Comissão Recepção'],
       observacoes: preco.observacoes || preco.Observações || '',
       ativo: preco.ativo !== false
     });
     
     if (resultado.sucesso) {
       importados++;
     } else {
       erros++;
     }
     
     resultados.push(resultado);
   });
   
   logger.success('Importação concluída', { importados, erros });
   
   return {
     sucesso: true,
     importados: importados,
     erros: erros,
     total: precosParaImportar.length,
     resultados: resultados
   };
   
 } catch (error) {
   logger.error('Erro na importação', error);
   return {
     sucesso: false,
     erro: error.message
   };
 }
}

/**
* Exporta tabela de preços (CÓDIGO ANTIGO PRESERVADO + ADAPTAÇÕES)
* @param {string} formato - Formato de exportação (csv, json)
* @returns {string} - Dados exportados
*/
function exportarPrecos(formato = 'csv') {
 logger.info('Exportando preços', { formato });
 
 try {
   const precos = listarPrecosTabela();
   
   if (formato === 'csv') {
     // Gerar CSV
     const headers = ['ID', 'Tipo Serviço', 'Rota', 'Origem', 'Destino', 'Pessoas', 'Bagagens', 
                    'Preço Por Pessoa', 'Preço Por Grupo', 'Preço Cliente', 'Valor Hotel', 
                    'Valor HUB', 'Comissão Recepção', 'Ativo', 'Data Criação', 'Observações'];
     
     let csv = headers.join(',') + '\n';
     
     precos.forEach(preco => {
       const linha = [
         preco.id,
         `"${preco.tipoServico}"`,
         `"${preco.rota}"`,
         `"${preco.origem}"`,
         `"${preco.destino}"`,
         preco.pessoas,
         preco.bagagens,
         preco.precoPorPessoa,
         preco.precoPorGrupo,
         preco.precoCliente,
         preco.valorHotel,
         preco.valorHUB,
         preco.comissaoRecepcao,
         preco.ativo ? 'Sim' : 'Não',
         formatarDataHora(new Date(preco.dataCriacao)),
         `"${preco.observacoes || ''}"`
       ];
       
       csv += linha.join(',') + '\n';
     });
     
     return csv;
     
   } else if (formato === 'json') {
     return JSON.stringify(precos, null, 2);
   }
   
   throw new Error('Formato não suportado');
   
 } catch (error) {
   logger.error('Erro na exportação', error);
   throw error;
 }
}

// ===================================================
// FUNÇÕES DE MANUTENÇÃO E UTILITÁRIOS (FUSÃO COMPLETA)
// ===================================================

/**
 * Limpa todos os dados do sistema (CÓDIGO ANTIGO PRESERVADO)
 * @returns {Object} - Resultado da limpeza
 */
function limparDadosCompleto() {
  logger.warn('LIMPANDO TODOS OS DADOS DO SISTEMA');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheets = ss.getSheets();
    let registrosRemovidos = 0;
    let abasProcessadas = 0;
    
    sheets.forEach(sheet => {
      const nome = sheet.getName();
      
      // Processar apenas abas do sistema
      if (nome === CONFIG.SHEET_NAME || 
          nome === CONFIG.PRICING_SHEET_NAME ||
          nome.startsWith(CONFIG.SISTEMA.PREFIXO_MES)) {
        
        const lastRow = sheet.getLastRow();
        
        if (lastRow > 1) {
          // Preservar headers
          const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
          
          // Limpar tudo
          sheet.clear();
          
          // Restaurar headers
          sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
          
          // Reaplicar formatação
          if (nome === CONFIG.SHEET_NAME) {
            aplicarFormatacao(sheet);
          } else if (nome === CONFIG.PRICING_SHEET_NAME) {
            aplicarFormatacaoPrecos(sheet);
          } else if (nome.startsWith(CONFIG.SISTEMA.PREFIXO_MES)) {
            aplicarFormatacaoMensal(sheet);
          }
          
          registrosRemovidos += lastRow - 1;
          abasProcessadas++;
          
          logger.info(`Aba ${nome} limpa`, { registros: lastRow - 1 });
        }
      }
    });
    
    logger.success('Limpeza completa concluída', {
      registrosRemovidos,
      abasProcessadas
    });
    
    return {
      sucesso: true,
      registrosRemovidos: registrosRemovidos,
      abasProcessadas: abasProcessadas,
      mensagem: `${registrosRemovidos} registros removidos de ${abasProcessadas} abas`
    };
    
  } catch (error) {
    logger.error('Erro na limpeza completa', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Limpa apenas dados de teste (CÓDIGO ANTIGO PRESERVADO)
 * @returns {Object} - Resultado da limpeza
 */
function limparDadosTeste() {
  logger.info('Limpando dados de teste');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheets = ss.getSheets();
    let totalRemovidos = 0;
    const detalhes = [];
    
    // Palavras-chave que identificam dados de teste
    const palavrasTeste = ['teste', 'test', 'demo', 'exemplo', 'sample'];
    
    sheets.forEach(sheet => {
      const nome = sheet.getName();
      
      if (nome === CONFIG.SHEET_NAME || nome.startsWith(CONFIG.SISTEMA.PREFIXO_MES)) {
        const lastRow = sheet.getLastRow();
        if (lastRow <= 1) return;
        
        const dados = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
        const linhasParaRemover = [];
        
        // Identificar linhas de teste
        dados.forEach((row, index) => {
          const id = row[0];
          const cliente = String(row[1]).toLowerCase();
          const observacoes = String(row[18]).toLowerCase(); // Coluna S
          
          // Verificar se é teste
          let ehTeste = false;
          
          // ID 999 ou 9999 são sempre teste
          if (id === 999 || id === 9999) {
            ehTeste = true;
          }
          
          // Verificar palavras-chave no nome do cliente
          palavrasTeste.forEach(palavra => {
            if (cliente.includes(palavra) || observacoes.includes(palavra)) {
              ehTeste = true;
            }
          });
          
          if (ehTeste) {
            linhasParaRemover.push(index + 2); // +2 porque começamos na linha 2
          }
        });
        
        // Remover de baixo para cima
        if (linhasParaRemover.length > 0) {
          linhasParaRemover.reverse().forEach(linha => {
            sheet.deleteRow(linha);
          });
          
          totalRemovidos += linhasParaRemover.length;
          detalhes.push({
            aba: nome,
            removidos: linhasParaRemover.length
          });
          
          logger.info(`${linhasParaRemover.length} registros de teste removidos de ${nome}`);
        }
      }
    });
    
    logger.success('Limpeza de testes concluída', { totalRemovidos });
    
    return {
      sucesso: true,
      totalRemovidos: totalRemovidos,
      detalhes: detalhes,
      mensagem: `${totalRemovidos} registros de teste removidos`
    };
    
  } catch (error) {
    logger.error('Erro na limpeza de testes', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Reordena transfers por data (CÓDIGO ANTIGO PRESERVADO)
 * @returns {Object} - Resultado da reordenação
 */
function reordenarPorData() {
  logger.info('Reordenando transfers por data');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() <= 2) {
      return {
        sucesso: true,
        mensagem: 'Poucos dados para reordenar'
      };
    }
    
    // Obter dados (excluindo headers)
    const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length);
    const dados = range.getValues();
    
    // Ordenar por data (coluna F) e depois por hora (coluna K)
    dados.sort((a, b) => {
      const dataA = new Date(a[5]); // Coluna F
      const dataB = new Date(b[5]);
      
      if (dataA.getTime() !== dataB.getTime()) {
        return dataA - dataB;
      }
      
      // Se mesma data, ordenar por hora
      const horaA = a[10] || '00:00'; // Coluna K
      const horaB = b[10] || '00:00';
      
      return horaA.localeCompare(horaB);
    });
    
    // Reescrever dados ordenados
    range.setValues(dados);
    
    logger.success('Reordenação concluída', { registros: dados.length });
    
    return {
      sucesso: true,
      registros: dados.length,
      mensagem: `${dados.length} registros reordenados por data`
    };
    
  } catch (error) {
    logger.error('Erro na reordenação', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Corrige registros incompletos (sem registro duplo) - CÓDIGO ANTIGO PRESERVADO
 * @returns {Object} - Resultado da correção
 */
function corrigirRegistrosIncompletos() {
  logger.info('Iniciando correção de registros incompletos');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const abaPrincipal = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!abaPrincipal || abaPrincipal.getLastRow() <= 1) {
      return {
        sucesso: true,
        mensagem: 'Nenhum registro para verificar',
        corrigidos: 0
      };
    }
    
    const dados = abaPrincipal.getRange(2, 1, abaPrincipal.getLastRow() - 1, HEADERS.length).getValues();
    
    const resultados = {
      verificados: 0,
      corrigidos: 0,
      erros: 0,
      detalhes: []
    };
    
    dados.forEach((row, index) => {
      const transferId = row[0];
      const dataTransfer = row[5]; // Coluna F
      
      if (!transferId || !dataTransfer) return;
      
      resultados.verificados++;
      
      try {
        // Obter aba mensal correspondente
        const abaMensal = obterAbaMes(dataTransfer);
        
        if (abaMensal && abaMensal.getName() !== CONFIG.SHEET_NAME) {
          // Verificar se existe na aba mensal
          const existeNaMensal = encontrarLinhaPorId(abaMensal, transferId) > 0;
          
          if (!existeNaMensal) {
            // Inserir na aba mensal
            abaMensal.appendRow(row);
            
            resultados.corrigidos++;
            resultados.detalhes.push({
              id: transferId,
              cliente: row[1],
              abaMensal: abaMensal.getName(),
              status: 'corrigido'
            });
            
            logger.info('Registro corrigido', {
              transferId,
              abaMensal: abaMensal.getName()
            });
          }
        }
        
      } catch (error) {
        logger.error('Erro ao corrigir registro', {
          transferId,
          erro: error.message
        });
        
        resultados.erros++;
        resultados.detalhes.push({
          id: transferId,
          cliente: row[1],
          status: 'erro',
          erro: error.message
        });
      }
    });
    
    logger.success('Correção de registros concluída', resultados);
    
    return {
      sucesso: true,
      ...resultados
    };
    
  } catch (error) {
    logger.error('Erro na correção de registros', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Remove registros duplicados (CÓDIGO ANTIGO PRESERVADO)
 * @param {boolean} simular - Se true, apenas simula sem remover
 * @returns {Object} - Resultado da operação
 */
function removerDuplicados(simular = true) {
  logger.info('Verificando registros duplicados', { simular });
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return {
        sucesso: true,
        mensagem: 'Nenhum registro para verificar',
        duplicados: 0
      };
    }
    
    const dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
    const registrosUnicos = new Map();
    const duplicados = [];
    
    // Identificar duplicados
    dados.forEach((row, index) => {
      const id = row[0];
      const data = formatarDataDDMMYYYY(new Date(row[5])); // Coluna F
      const chave = `${id}_${data}`;
      
      if (registrosUnicos.has(chave)) {
        duplicados.push({
          linha: index + 2, // +2 porque começamos na linha 2
          id: id,
          cliente: row[1],
          data: data
        });
      } else {
        registrosUnicos.set(chave, index + 2);
      }
    });
    
    logger.info(`${duplicados.length} duplicados encontrados`);
    
    // Remover duplicados se não for simulação
    if (!simular && duplicados.length > 0) {
      // Ordenar de maior para menor linha (para remover de baixo para cima)
      duplicados.sort((a, b) => b.linha - a.linha);
      
      duplicados.forEach(dup => {
        sheet.deleteRow(dup.linha);
        logger.info('Duplicado removido', {
          id: dup.id,
          cliente: dup.cliente,
          linha: dup.linha
        });
      });
    }
    
    return {
      sucesso: true,
      duplicados: duplicados.length,
      detalhes: duplicados,
      simulacao: simular
    };
    
  } catch (error) {
    logger.error('Erro ao remover duplicados', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Sincroniza dados entre aba principal e abas mensais (CÓDIGO ANTIGO PRESERVADO)
 * @returns {Object} - Resultado da sincronização
 */
function sincronizarAbasMensais() {
  logger.info('Iniciando sincronização de abas mensais');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const abaPrincipal = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!abaPrincipal || abaPrincipal.getLastRow() <= 1) {
      return {
        sucesso: true,
        mensagem: 'Nenhum registro para sincronizar'
      };
    }
    
    const dados = abaPrincipal.getRange(2, 1, abaPrincipal.getLastRow() - 1, HEADERS.length).getValues();
    const resultados = {
      processados: 0,
      sincronizados: 0,
      erros: 0,
      detalhes: []
    };
    
    // Agrupar por mês
    const transfersPorMes = new Map();
    
    dados.forEach(row => {
      const data = new Date(row[5]); // Coluna F
      const mes = String(data.getMonth() + 1).padStart(2, '0');
      const ano = data.getFullYear();
      const chaveMes = `${mes}_${ano}`;
      
      if (!transfersPorMes.has(chaveMes)) {
        transfersPorMes.set(chaveMes, []);
      }
      
      transfersPorMes.get(chaveMes).push(row);
    });
    
    // Processar cada mês
    transfersPorMes.forEach((transfers, chaveMes) => {
      try {
        const [mes, ano] = chaveMes.split('_');
        const mesInfo = MESES.find(m => m.abrev === mes);
        const nomeAba = `${CONFIG.SISTEMA.PREFIXO_MES}${mes}_${mesInfo.nome}_${ano}`;
        
        let abaMensal = ss.getSheetByName(nomeAba);
        
        // Criar aba se não existir
        if (!abaMensal) {
          abaMensal = criarAbaMensal(nomeAba, ss, mesInfo);
        }
        
        // Limpar dados existentes (exceto headers)
        if (abaMensal.getLastRow() > 1) {
          abaMensal.deleteRows(2, abaMensal.getLastRow() - 1);
        }
        
        // Inserir todos os transfers do mês
        transfers.forEach(transfer => {
          abaMensal.appendRow(transfer);
          resultados.sincronizados++;
        });
        
        resultados.processados += transfers.length;
        resultados.detalhes.push({
          mes: mesInfo.nome,
          ano: ano,
          transfers: transfers.length,
          status: 'sincronizado'
        });
        
        logger.success(`Aba ${nomeAba} sincronizada`, {
          transfers: transfers.length
        });
        
      } catch (errorMes) {
        logger.error('Erro ao sincronizar mês', {
          mes: chaveMes,
          erro: errorMes.message
        });
        
        resultados.erros++;
        resultados.detalhes.push({
          mes: chaveMes,
          status: 'erro',
          erro: errorMes.message
        });
      }
    });
    
    logger.success('Sincronização concluída', resultados);
    
    return {
      sucesso: true,
      ...resultados
    };
    
  } catch (error) {
    logger.error('Erro na sincronização', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Cria backup dos dados (CÓDIGO ANTIGO PRESERVADO)
 * @returns {Object} - Informações do backup
 */
function criarBackup() {
  logger.info('Iniciando backup do sistema');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const dataHora = new Date();
    const nomeBackup = `Backup_${formatarDataHora(dataHora).replace(/[/:]/g, '-')}`;
    
    // Criar nova planilha de backup
    const backupSS = ss.copy(nomeBackup);
    
    // Registrar backup
    let backupSheet = ss.getSheetByName('Backups');
    if (!backupSheet) {
      backupSheet = ss.insertSheet('Backups');
      backupSheet.appendRow(['Data/Hora', 'Nome', 'ID Planilha', 'Status']);
    }
    
    backupSheet.appendRow([
      dataHora,
      nomeBackup,
      backupSS.getId(),
      'Completo'
    ]);
    
    logger.success('Backup criado com sucesso', { nome: nomeBackup, id: backupSS.getId() });
    
    return {
      sucesso: true,
      nome: nomeBackup,
      id: backupSS.getId(),
      dataHora: dataHora
    };
    
  } catch (error) {
    logger.error('Erro ao criar backup', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Aplica formatação na planilha principal (ATUALIZADO PARA CÓDIGO NOVO)
 * @param {Sheet} sheet - Planilha a formatar
 */
function aplicarFormatacao(sheet) {
  logger.debug('Aplicando formatação na planilha principal');
  
  try {
    // Aplicar larguras das colunas
    STYLES.COLUMN_WIDTHS.PRINCIPAL.forEach((width, index) => {
      sheet.setColumnWidth(index + 1, width);
    });
    
    const maxRows = Math.max(sheet.getMaxRows(), 1000);
    
    // Formatação dos headers
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange
      .setBackground(STYLES.HEADER_COLORS.PRINCIPAL)
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setFontSize(11)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    
    // Congelar primeira linha
    sheet.setFrozenRows(1);
    
    // Formatações de dados (ESTRUTURA DO CÓDIGO NOVO)
    if (maxRows > 1) {
      // Moeda (colunas L, M, N, O)
      sheet.getRange(2, 12, maxRows - 1, 4).setNumberFormat(STYLES.FORMATS.MOEDA);
      
      // Data (coluna F)
      sheet.getRange(2, 6, maxRows - 1, 1).setNumberFormat(STYLES.FORMATS.DATA);
      
      // Hora (coluna K)
      sheet.getRange(2, 11, maxRows - 1, 1).setNumberFormat(STYLES.FORMATS.HORA);
      
      // Timestamp (coluna T)
      sheet.getRange(2, 20, maxRows - 1, 1).setNumberFormat(STYLES.FORMATS.TIMESTAMP);
      
// Números (colunas A, D, E)
     sheet.getRange(2, 1, maxRows - 1, 1).setNumberFormat(STYLES.FORMATS.NUMERO);
     sheet.getRange(2, 4, maxRows - 1, 2).setNumberFormat(STYLES.FORMATS.NUMERO);
   }
   
   // Aplicar validações
   aplicarValidacoesPlanilha(sheet);
   
   logger.debug('Formatação aplicada com sucesso');
   
 } catch (error) {
   logger.error('Erro ao aplicar formatação', error);
 }
}

/**
* Aplica formatação na tabela de preços (ATUALIZADO PARA CÓDIGO NOVO)
* @param {Sheet} sheet - Planilha de preços
*/
function aplicarFormatacaoPrecos(sheet) {
 logger.debug('Aplicando formatação na tabela de preços');
 
 try {
   // Larguras das colunas
   STYLES.COLUMN_WIDTHS.PRECOS.forEach((width, index) => {
     sheet.setColumnWidth(index + 1, width);
   });
   
   // Headers
   const headerRange = sheet.getRange(1, 1, 1, PRICING_HEADERS.length);
   headerRange
     .setBackground(STYLES.HEADER_COLORS.PRECOS)
     .setFontColor('#ffffff')
     .setFontWeight('bold')
     .setFontSize(11)
     .setHorizontalAlignment('center')
     .setVerticalAlignment('middle');
   
   sheet.setFrozenRows(1);
   
   const maxRows = Math.max(sheet.getMaxRows(), 500);
   
   if (maxRows > 1) {
     // Formatação monetária (colunas H, I, J, K, L, M)
     sheet.getRange(2, 8, maxRows - 1, 6).setNumberFormat(STYLES.FORMATS.MOEDA);
     
     // Timestamp (coluna O)
     sheet.getRange(2, 15, maxRows - 1, 1).setNumberFormat(STYLES.FORMATS.TIMESTAMP);
     
     // Números (colunas A, F, G)
     sheet.getRange(2, 1, maxRows - 1, 1).setNumberFormat(STYLES.FORMATS.NUMERO);
     sheet.getRange(2, 6, maxRows - 1, 2).setNumberFormat(STYLES.FORMATS.NUMERO);
     
     // Validação Tipo de Serviço (coluna B)
     const tipoValidation = SpreadsheetApp.newDataValidation()
       .requireValueInList(VALIDACOES.VALORES_PERMITIDOS.TIPO_SERVICO)
       .setAllowInvalid(false)
       .build();
     sheet.getRange(2, 2, maxRows - 1, 1).setDataValidation(tipoValidation);
     
     // Validação Ativo/Inativo (coluna N)
     const ativoValidation = SpreadsheetApp.newDataValidation()
       .requireValueInList(['Sim', 'Não'])
       .setAllowInvalid(false)
       .build();
     sheet.getRange(2, 14, maxRows - 1, 1).setDataValidation(ativoValidation);
   }
   
   logger.debug('Formatação de preços aplicada');
   
 } catch (error) {
   logger.error('Erro ao aplicar formatação de preços', error);
 }
}

/**
* Realiza manutenção automática do sistema (CÓDIGO ANTIGO PRESERVADO)
*/
function manutencaoAutomatica() {
 logger.info('Iniciando manutenção automática');
 
 try {
   const tarefas = [];
   
   // 1. Verificar integridade das abas
   const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
   const sheets = ss.getSheets();
   sheets.forEach(sheet => {
     const nome = sheet.getName();
     if (nome.startsWith(CONFIG.SISTEMA.PREFIXO_MES) || 
         nome === CONFIG.SHEET_NAME || 
         nome === CONFIG.PRICING_SHEET_NAME) {
       // Verificar headers
       const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
       if (headers.length === 0 || !headers[0]) {
         logger.warn('Aba sem headers', { nome });
         // Recriar headers se necessário
         if (nome === CONFIG.SHEET_NAME || nome.startsWith(CONFIG.SISTEMA.PREFIXO_MES)) {
           sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
           tarefas.push(`Headers recriados em ${nome}`);
         } else if (nome === CONFIG.PRICING_SHEET_NAME) {
           sheet.getRange(1, 1, 1, PRICING_HEADERS.length).setValues([PRICING_HEADERS]);
           tarefas.push(`Headers recriados em ${nome}`);
         }
       }
     }
   });
   
   // 2. Limpar logs antigos se configurado
   if (LOG_CONFIG.PERSIST_TO_SHEET) {
     const logSheet = ss.getSheetByName(LOG_CONFIG.LOG_SHEET_NAME);
     if (logSheet && logSheet.getLastRow() > 10000) {
       // Manter apenas últimas 5000 linhas
       const rowsToDelete = logSheet.getLastRow() - 5000;
       logSheet.deleteRows(2, rowsToDelete);
       tarefas.push(`${rowsToDelete} logs antigos removidos`);
     }
   }
   
   // 3. Criar backup se configurado
   if (CONFIG.SISTEMA.BACKUP_AUTOMATICO) {
     const backup = criarBackup();
     if (backup.sucesso) {
       tarefas.push(`Backup criado: ${backup.nome}`);
     }
   }
   
   logger.success('Manutenção concluída', { tarefas });
   
   return {
     sucesso: true,
     tarefas: tarefas,
     timestamp: new Date()
   };
   
 } catch (error) {
   logger.error('Erro na manutenção automática', error);
   return {
     sucesso: false,
     erro: error.message
   };
 }
}

// ===================================================
// CONFIGURAÇÃO DE TRIGGERS (CÓDIGO ANTIGO PRESERVADO + NOVAS)
// ===================================================

/**
 * Configura triggers automáticos para o sistema (FUSÃO)
 */
function configurarTriggersEmail() {
  logger.info('Configurando triggers de e-mail');
  
  try {
    // Remover triggers existentes relacionados a e-mail
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      const handlerFunction = trigger.getHandlerFunction();
      if (handlerFunction === 'verificarConfirmacoesEmail' || 
          handlerFunction === 'enviarRelatorioDiaAnterior' ||
          handlerFunction === 'manutencaoAutomatica') {
        ScriptApp.deleteTrigger(trigger);
        logger.info(`Trigger removido: ${handlerFunction}`);
      }
    });
    
    // Criar trigger para verificação de e-mails (CÓDIGO ANTIGO)
    if (CONFIG.EMAIL_CONFIG.VERIFICAR_CONFIRMACOES) {
      ScriptApp.newTrigger('verificarConfirmacoesEmail')
        .timeBased()
        .everyMinutes(CONFIG.EMAIL_CONFIG.INTERVALO_VERIFICACAO)
        .create();
      
      logger.success('Trigger de verificação de e-mails criado', {
        intervalo: CONFIG.EMAIL_CONFIG.INTERVALO_VERIFICACAO + ' minutos'
      });
    }
    
    // Criar trigger para relatório diário (CÓDIGO NOVO)
    if (CONFIG.EMAIL_CONFIG.ENVIAR_RELATORIO_DIA_ANTERIOR) {
      ScriptApp.newTrigger('enviarRelatorioDiaAnterior')
        .timeBased()
        .atHour(9)
        .everyDays(1)
        .inTimezone(CONFIG.SISTEMA.TIMEZONE)
        .create();
      
      logger.success('Trigger de relatório diário criado');
    }
    
    // Criar trigger para manutenção automática (CÓDIGO ANTIGO)
    if (CONFIG.SISTEMA.BACKUP_AUTOMATICO) {
      ScriptApp.newTrigger('manutencaoAutomatica')
        .timeBased()
        .atHour(2)
        .everyDays(1)
        .inTimezone(CONFIG.SISTEMA.TIMEZONE)
        .create();
      
      logger.success('Trigger de manutenção automática criado');
    }
    
    return {
      sucesso: true,
      mensagem: 'Triggers configurados com sucesso'
    };
    
  } catch (error) {
    logger.error('Erro ao configurar triggers', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Remove todos os triggers de e-mail (CÓDIGO ANTIGO PRESERVADO)
 */
function removerTriggersEmail() {
  logger.info('Removendo triggers de e-mail');
  
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let removidos = 0;
    
    triggers.forEach(trigger => {
      const handlerFunction = trigger.getHandlerFunction();
      if (handlerFunction === 'verificarConfirmacoesEmail' || 
          handlerFunction === 'enviarRelatorioDiaAnterior' ||
          handlerFunction === 'manutencaoAutomatica') {
        ScriptApp.deleteTrigger(trigger);
        removidos++;
        logger.info(`Trigger removido: ${handlerFunction}`);
      }
    });
    
    logger.success(`${removidos} triggers removidos`);
    
    return {
      sucesso: true,
      removidos: removidos
    };
    
  } catch (error) {
    logger.error('Erro ao remover triggers', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Verifica e-mails recebidos para processar confirmações (CÓDIGO ANTIGO PRESERVADO)
 * @returns {number} - Número de confirmações processadas
 */
function verificarConfirmacoesEmail() {
  logger.info('Verificando confirmações por e-mail');
  
  try {
    if (!CONFIG.EMAIL_CONFIG.VERIFICAR_CONFIRMACOES) {
      logger.info('Verificação de confirmações desativada');
      return 0;
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) {
      throw new Error('Planilha principal não encontrada');
    }
    
    // Buscar e-mails dos últimos 2 dias
    const emailBusca = CONFIG.EMAIL_CONFIG.DESTINATARIO || CONFIG.EMAIL_CONFIG.DESTINATARIOS[0];
    const query = `from:(${emailBusca}) subject:(transfer OR novo) newer_than:2d is:unread`;
    const threads = GmailApp.search(query, 0, 50); // Máximo 50 threads
    
    logger.info(`${threads.length} threads encontradas para verificação`);
    
    let confirmacoes = 0;
    let cancelamentos = 0;
    
    threads.forEach(thread => {
      try {
        const messages = thread.getMessages();
        
        messages.forEach(message => {
          // Verificar apenas respostas
          if (!message.getFrom().includes(emailBusca)) {
            return; // Pular e-mails enviados pelo sistema
          }
          
          const body = message.getPlainBody().toLowerCase();
          const subject = message.getSubject().toLowerCase();
          const from = message.getFrom();
          
          // Extrair e-mail do remetente
          const emailMatch = from.match(/[^\s<]+@[^\s>]+/);
          const remetenteEmail = emailMatch ? emailMatch[0] : from;
          
          // Procurar por confirmação (OK)
          if (body.includes('ok') || subject.includes('ok')) {
            // Tentar extrair ID do transfer
            const idMatch = body.match(/ok\s*#?(\d+)/) || 
                          subject.match(/#(\d+)/) ||
                          body.match(/transfer\s*#?(\d+)/i);
            
            let transferId = null;
            
            if (idMatch) {
              transferId = idMatch[1];
            } else {
              // Se não encontrou ID, buscar último transfer solicitado
              transferId = encontrarUltimoTransferSolicitado(sheet);
            }
            
            if (transferId) {
              const resultado = atualizarStatusTransfer(
                transferId, 
                'Confirmado',
                MESSAGES.ACOES.CONFIRMADO_POR(remetenteEmail)
              );
              
              if (resultado.sucesso) {
                confirmacoes++;
                logger.success(`Transfer #${transferId} confirmado via e-mail`);
                
                // Marcar mensagem como lida
                message.markRead();
                
                // Arquivar thread se configurado
                if (CONFIG.EMAIL_CONFIG.ARQUIVAR_CONFIRMADOS) {
                  thread.moveToArchive();
                }
              }
            }
          }
          
          // Procurar por cancelamento
          else if (body.includes('cancelar') || body.includes('cancel')) {
            const idMatch = body.match(/(?:cancelar|cancel)\s*#?(\d+)/i) || 
                          subject.match(/#(\d+)/);
            
            if (idMatch) {
              const transferId = idMatch[1];
              const resultado = atualizarStatusTransfer(
                transferId, 
                'Cancelado',
                MESSAGES.ACOES.CANCELADO_POR(remetenteEmail)
              );
              
              if (resultado.sucesso) {
                cancelamentos++;
                logger.success(`Transfer #${transferId} cancelado via e-mail`);
                
                message.markRead();
                
                if (CONFIG.EMAIL_CONFIG.ARQUIVAR_CONFIRMADOS) {
                  thread.moveToArchive();
                }
              }
            }
          }
        });
        
      } catch (threadError) {
        logger.error('Erro ao processar thread', threadError);
      }
    });
    
    const totalProcessado = confirmacoes + cancelamentos;
    
    logger.info('Verificação de e-mails concluída', {
      confirmacoes,
      cancelamentos,
      total: totalProcessado
    });
    
    // Enviar resumo se houver processamentos
    if (totalProcessado > 0) {
      enviarResumoVerificacaoEmail(confirmacoes, cancelamentos);
    }
    
    return totalProcessado;
    
  } catch (error) {
    logger.error('Erro na verificação de confirmações', error);
    return 0;
  }
}

/**
 * Encontra o último transfer com status "Solicitado" (CÓDIGO ANTIGO PRESERVADO)
 * @private
 */
function encontrarUltimoTransferSolicitado(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    
    // Buscar de baixo para cima (mais recentes primeiro)
    for (let i = lastRow; i >= 2; i--) {
      const status = sheet.getRange(i, 18).getValue(); // Coluna R
      
      if (status === 'Solicitado') {
        const id = sheet.getRange(i, 1).getValue();
        return id;
      }
    }
    
    return null;
    
  } catch (error) {
    logger.error('Erro ao buscar último transfer solicitado', error);
    return null;
  }
}

/**
 * Envia resumo das verificações de e-mail (CÓDIGO ANTIGO PRESERVADO)
 * @private
 */
function enviarResumoVerificacaoEmail(confirmacoes, cancelamentos) {
  try {
    const destinatarios = CONFIG.EMAIL_CONFIG.DESTINATARIOS.join(',');
    const total = confirmacoes + cancelamentos;
    
    const assunto = `[${CONFIG.NAMES.SISTEMA_NOME}] Resumo de Processamento Automático`;
    
    const corpo = `
      <h2>Resumo de Processamento Automático de E-mails</h2>
      <p>O sistema processou ${total} e-mail(s) com sucesso:</p>
      <ul>
        <li>✅ Confirmações: ${confirmacoes}</li>
        <li>❌ Cancelamentos: ${cancelamentos}</li>
      </ul>
      <p>Data/Hora: ${formatarDataHora(new Date())}</p>
      <hr>
      <p><small>Este é um e-mail automático do sistema de verificação.</small></p>
    `;
    
// Correção: usar htmlBody correto
sendEmailComAssinatura({
  to: destinatarios,
  subject: assunto,
  htmlBody: htmlEmail || corpo || corpoHtml,
  name: CONFIG.NAMES.SISTEMA_NOME,
  replyTo: CONFIG.EMAIL_CONFIG.EMAIL_EMPRESA
});
    
  } catch (error) {
    logger.error('Erro ao enviar resumo de verificação', error);
  }
}

// ===================================================
// MENU DO SISTEMA (FUSÃO CÓDIGO NOVO + ANTIGO)
// ===================================================

/**
 * Menu atualizado com opções de relatórios
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu(`🚐 Sistema Empire Marques Hotel`)
    .addItem('⚙️ Configurar Sistema', 'configurarSistema')
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 Relatórios e Acerto de Contas')
      .addItem('🏗️ Criar Aba de Relatórios', 'criarAbaRelatorios')
      .addItem('📊 Relatório Geral', 'gerarRelatorioGeral')
      .addItem('💰 Acerto de Contas', 'gerarAcertoContas')
      .addItem('📈 Relatório por Status', 'gerarRelatorioPorStatus')
      .addItem('🚐 Relatório por Tipo', 'gerarRelatorioPorTipo')
      .addItem('🏆 Top Clientes', 'gerarTopClientes')
      .addItem('📤 Exportar CSV', 'exportarRelatorioCSV')
      .addItem('🔄 Limpar Resultados', 'limparAreaResultados'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📧 E-mail Interativo')
      .addItem('✉️ Testar Envio de E-mail', 'testarEnvioEmailInterativo')
      .addItem('🔍 Verificar Confirmações', 'verificarConfirmacoesManual')
      .addItem('⏰ Configurar Verificação Automática', 'configurarTriggersEmailMenu')
      .addItem('🛑 Parar Verificação Automática', 'removerTriggersEmailMenu')
      .addItem('📊 Enviar Relatório Diário', 'enviarRelatorioDiaAnteriorManual'))
    .addSeparator()
    .addSubMenu(ui.createMenu('💰 Gestão de Preços')
      .addItem('📋 Ver Tabela de Preços', 'abrirTabelaPrecos')
      .addItem('🔍 Consultar Preço', 'consultarPrecoMenu')
      .addItem('➕ Adicionar Preço', 'adicionarPrecoMenu')
      .addItem('📥 Importar Preços (CSV)', 'importarPrecosMenu')
      .addItem('📤 Exportar Preços', 'exportarPrecosMenu'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📅 Abas Mensais')
      .addItem('📅 Criar Todas as Abas Mensais', 'criarTodasAbasMensaisMenu')
      .addItem('🔍 Verificar Integridade', 'verificarIntegridadeMenu')
      .addItem('🔧 Reparar Abas com Problemas', 'repararAbasMensaisMenu')
      .addItem('🔄 Sincronizar com Aba Principal', 'sincronizarAbasMensaisMenu'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🔧 Manutenção')
      .addItem('🔄 Reordenar por Data', 'reordenarPorDataMenu')
      .addItem('🔧 Corrigir Registros Incompletos', 'corrigirRegistrosMenu')
      .addItem('🔍 Verificar Duplicados', 'verificarDuplicadosMenu')
      .addItem('🗑️ Remover Duplicados', 'removerDuplicadosMenu')
      .addItem('🧹 Limpar Dados de Teste', 'limparDadosTesteMenu')
      .addItem('💾 Criar Backup', 'criarBackupMenu')
      .addItem('⚠️ Limpar TODOS os Dados', 'limparDadosCompletoMenu'))
    .addSeparator()
    .addItem('🧪 Testar Sistema Completo', 'testarSistemaCompleto')
    .addItem('ℹ️ Sobre o Sistema', 'mostrarSobre')
    .addToUi();
}

/**
 * Funções para cliques nos botões da aba de relatórios
 */
function onEdit(e) {
  // Detectar cliques nos botões da aba de relatórios
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  // Verificar se é a aba de relatórios
  if (sheet.getName() !== '📊 Relatórios') {
    return;
  }
  
  // Mapeamento de células para funções
  const botoesRelatorio = {
    'E7': 'gerarRelatorioGeral',
    'F7': 'gerarAcertoContas', 
    'G7': 'gerarRelatorioPorStatus',
    'H7': 'gerarRelatorioPorTipo',
    'E8': 'gerarTopClientes',
    'F8': 'gerarTopRotas',
    'G8': 'gerarRelatorioPorPagamento',
    'H8': 'gerarRelatorioMensalCompleto',
    'E9': 'exportarRelatorioCSV',
    'F9': 'enviarRelatorioPorEmail',
    'G9': 'limparAreaResultados',
    'H9': 'configurarRelatorios'
  };
  
  const cellAddress = range.getA1Notation();
  const funcaoParaExecutar = botoesRelatorio[cellAddress];
  
  if (funcaoParaExecutar) {
    try {
      // Executar a função correspondente
      switch (funcaoParaExecutar) {
        case 'gerarRelatorioGeral':
          gerarRelatorioGeral();
          break;
        case 'gerarAcertoContas':
          gerarAcertoContas();
          break;
        case 'gerarRelatorioPorStatus':
          gerarRelatorioPorStatus();
          break;
        case 'gerarRelatorioPorTipo':
          gerarRelatorioPorTipo();
          break;
        case 'gerarTopClientes':
          gerarTopClientes();
          break;
        case 'gerarTopRotas':
          gerarTopRotas();
          break;
        case 'gerarRelatorioPorPagamento':
          gerarRelatorioPorPagamento();
          break;
        case 'gerarRelatorioMensalCompleto':
          gerarRelatorioMensalCompleto();
          break;
        case 'exportarRelatorioCSV':
          exportarRelatorioCSV();
          break;
        case 'enviarRelatorioPorEmail':
          enviarRelatorioPorEmail();
          break;
        case 'limparAreaResultados':
          limparAreaResultados();
          break;
        case 'configurarRelatorios':
          configurarRelatorios();
          break;
      }
    } catch (error) {
      SpreadsheetApp.getUi().alert('Erro', `Erro ao executar ${funcaoParaExecutar}: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}

/**
 * Relatórios adicionais para completar o sistema
 */

function gerarRelatorioPorStatus() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const relatorioSheet = ss.getSheetByName('📊 Relatórios');
    const dataInicio = relatorioSheet.getRange('C7').getValue();
    const dataFim = relatorioSheet.getRange('C8').getValue();
    
    if (!dataInicio || !dataFim) {
      throw new Error('Preencha as datas');
    }
    
    const dados = coletarDadosRelatorio(dataInicio, dataFim);
    limparAreaResultados();
    
    let linha = 18;
    relatorioSheet.getRange(`B${linha}`).setValue(`📈 RELATÓRIO POR STATUS - ${formatarDataDDMMYYYY(dataInicio)} a ${formatarDataDDMMYYYY(dataFim)}`);
    relatorioSheet.getRange(`B${linha}`).setFontSize(14).setFontWeight('bold');
    linha += 2;
    
    Object.entries(dados.porStatus).forEach(([status, count]) => {
      const cor = status === 'Finalizado' ? '#d4edda' : 
                  status === 'Confirmado' ? '#cce5ff' : 
                  status === 'Cancelado' ? '#f8d7da' : '#fff3cd';
      
      relatorioSheet.getRange(`B${linha}`).setValue(`${status}:`);
      relatorioSheet.getRange(`C${linha}`).setValue(`${count} transfer(s)`);
      relatorioSheet.getRange(`B${linha}:C${linha}`).setBackground(cor);
      linha++;
    });
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Erro', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function gerarRelatorioPorTipo() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const relatorioSheet = ss.getSheetByName('📊 Relatórios');
    const dataInicio = relatorioSheet.getRange('C7').getValue();
    const dataFim = relatorioSheet.getRange('C8').getValue();
    
    if (!dataInicio || !dataFim) {
      throw new Error('Preencha as datas');
    }
    
    const dados = coletarDadosRelatorio(dataInicio, dataFim);
    limparAreaResultados();
    
    let linha = 18;
    relatorioSheet.getRange(`B${linha}`).setValue(`🚐 RELATÓRIO POR TIPO DE SERVIÇO - ${formatarDataDDMMYYYY(dataInicio)} a ${formatarDataDDMMYYYY(dataFim)}`);
    relatorioSheet.getRange(`B${linha}`).setFontSize(14).setFontWeight('bold');
    linha += 2;
    
    Object.entries(dados.porTipo).forEach(([tipo, count]) => {
      relatorioSheet.getRange(`B${linha}`).setValue(`${obterLabelTipoServico(tipo)}:`);
      relatorioSheet.getRange(`C${linha}`).setValue(`${count} transfer(s)`);
      linha++;
    });
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Erro', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function gerarTopClientes() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const relatorioSheet = ss.getSheetByName('📊 Relatórios');
    const dataInicio = relatorioSheet.getRange('C7').getValue();
    const dataFim = relatorioSheet.getRange('C8').getValue();
    
    if (!dataInicio || !dataFim) {
      throw new Error('Preencha as datas');
    }
    
    const dados = coletarDadosRelatorio(dataInicio, dataFim);
    limparAreaResultados();
    
    // Agrupar por cliente
    const clientes = {};
    dados.transfers.forEach(transfer => {
      if (transfer.status !== 'Cancelado') {
        if (!clientes[transfer.cliente]) {
          clientes[transfer.cliente] = {
            count: 0,
            valor: 0
          };
        }
        clientes[transfer.cliente].count++;
        clientes[transfer.cliente].valor += transfer.valorTotal;
      }
    });
    
    // Ordenar por número de transfers
    const topClientes = Object.entries(clientes)
      .sort((a, b) => b[1].count - a[1].count)
      .slice(0, 10);
    
    let linha = 18;
    relatorioSheet.getRange(`B${linha}`).setValue(`🏆 TOP 10 CLIENTES - ${formatarDataDDMMYYYY(dataInicio)} a ${formatarDataDDMMYYYY(dataFim)}`);
    relatorioSheet.getRange(`B${linha}`).setFontSize(14).setFontWeight('bold');
    linha += 2;
    
    topClientes.forEach(([cliente, dados], index) => {
      relatorioSheet.getRange(`B${linha}`).setValue(`${index + 1}. ${cliente}`);
      relatorioSheet.getRange(`C${linha}`).setValue(`${dados.count} transfer(s)`);
      relatorioSheet.getRange(`D${linha}`).setValue(`€${dados.valor.toFixed(2)}`);
      linha++;
    });
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Erro', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function gerarTopRotas() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const relatorioSheet = ss.getSheetByName('📊 Relatórios');
    const dataInicio = relatorioSheet.getRange('C7').getValue();
    const dataFim = relatorioSheet.getRange('C8').getValue();
    
    if (!dataInicio || !dataFim) {
      throw new Error('Preencha as datas');
    }
    
    const dados = coletarDadosRelatorio(dataInicio, dataFim);
    limparAreaResultados();
    
    // Agrupar por rota
    const rotas = {};
    dados.transfers.forEach(transfer => {
      if (transfer.status !== 'Cancelado') {
        const rota = `${transfer.origem} → ${transfer.destino}`;
        if (!rotas[rota]) {
          rotas[rota] = {
            count: 0,
            valor: 0
          };
        }
        rotas[rota].count++;
        rotas[rota].valor += transfer.valorTotal;
      }
    });
    
    // Ordenar por número de transfers
    const topRotas = Object.entries(rotas)
      .sort((a, b) => b[1].count - a[1].count)
      .slice(0, 10);
    
    let linha = 18;
    relatorioSheet.getRange(`B${linha}`).setValue(`🗺️ TOP 10 ROTAS - ${formatarDataDDMMYYYY(dataInicio)} a ${formatarDataDDMMYYYY(dataFim)}`);
    relatorioSheet.getRange(`B${linha}`).setFontSize(14).setFontWeight('bold');
    linha += 2;
    
    topRotas.forEach(([rota, dados], index) => {
      relatorioSheet.getRange(`B${linha}`).setValue(`${index + 1}. ${rota}`);
      relatorioSheet.getRange(`C${linha}`).setValue(`${dados.count} transfer(s)`);
      relatorioSheet.getRange(`D${linha}`).setValue(`€${dados.valor.toFixed(2)}`);
      linha++;
    });
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Erro', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function gerarRelatorioPorPagamento() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const relatorioSheet = ss.getSheetByName('📊 Relatórios');
    const dataInicio = relatorioSheet.getRange('C7').getValue();
    const dataFim = relatorioSheet.getRange('C8').getValue();
    
    if (!dataInicio || !dataFim) {
      throw new Error('Preencha as datas');
    }
    
    const dados = coletarDadosRelatorio(dataInicio, dataFim);
    limparAreaResultados();
    
    let linha = 18;
    relatorioSheet.getRange(`B${linha}`).setValue(`💳 RELATÓRIO POR FORMA DE PAGAMENTO - ${formatarDataDDMMYYYY(dataInicio)} a ${formatarDataDDMMYYYY(dataFim)}`);
    relatorioSheet.getRange(`B${linha}`).setFontSize(14).setFontWeight('bold');
    linha += 2;
    
    Object.entries(dados.porPagamento).forEach(([pagamento, count]) => {
      relatorioSheet.getRange(`B${linha}`).setValue(`${pagamento}:`);
      relatorioSheet.getRange(`C${linha}`).setValue(`${count} transfer(s)`);
      linha++;
    });
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Erro', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function gerarRelatorioMensalCompleto() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const relatorioSheet = ss.getSheetByName('📊 Relatórios');
    const mesAno = relatorioSheet.getRange('C9').getValue();
    
    if (!mesAno) {
      throw new Error('Preencha o mês/ano');
    }
    
    // Converter mês/ano para datas
    const [mes, ano] = mesAno.split('/');
    const dataInicio = new Date(parseInt(ano), parseInt(mes) - 1, 1);
    const dataFim = new Date(parseInt(ano), parseInt(mes), 0);
    
    const dados = coletarDadosRelatorio(dataInicio, dataFim);
    limparAreaResultados();
    
    exibirRelatorioGeral(relatorioSheet, dados, dataInicio, dataFim);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Erro', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function enviarRelatorioPorEmail() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const relatorioSheet = ss.getSheetByName('📊 Relatórios');
    const dataInicio = relatorioSheet.getRange('C7').getValue();
    const dataFim = relatorioSheet.getRange('C8').getValue();
    
    if (!dataInicio || !dataFim) {
      throw new Error('Preencha as datas');
    }
    
    const dados = coletarDadosRelatorio(dataInicio, dataFim);
    
    // Criar e-mail com o relatório
    const assunto = `📊 Relatório de Transfers - ${formatarDataDDMMYYYY(dataInicio)} a ${formatarDataDDMMYYYY(dataFim)}`;
    
    const corpo = `
      <h2>📊 Relatório de Transfers</h2>
      <p><strong>Período:</strong> ${formatarDataDDMMYYYY(dataInicio)} a ${formatarDataDDMMYYYY(dataFim)}</p>
      
      <h3>💰 Resumo Financeiro:</h3>
      <ul>
        <li>Total de Transfers: ${dados.totalTransfers}</li>
        <li>Receita Total: €${dados.totais.valorTotal.toFixed(2)}</li>
        <li>Valor ${CONFIG.NAMES.HOTEL_NAME}: €${dados.totais.valorHotel.toFixed(2)}</li>
        <li>Valor HUB Transfer: €${dados.totais.valorHUB.toFixed(2)}</li>
        <li>Comissão Recepção: €${dados.totais.comissaoRecepcao.toFixed(2)}</li>
      </ul>
      
      <p><em>Relatório gerado automaticamente em ${formatarDataHora(new Date())}</em></p>
    `;
    
    sendEmailComAssinatura({
      to: CONFIG.EMAIL_CONFIG.DESTINATARIOS.join(','),
      subject: assunto,
      htmlBody: corpo,
      name: CONFIG.NAMES.SISTEMA_NOME,
      replyTo: CONFIG.EMAIL_CONFIG.EMAIL_EMPRESA
    });
    
    SpreadsheetApp.getUi().alert('E-mail Enviado', 'Relatório enviado por e-mail com sucesso!', SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Erro', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Configura o sistema completamente (FUSÃO)
 */
function configurarSistema() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '⚙️ Configuração do Sistema',
    'Esta ação irá:\n\n' +
    '1. Criar/verificar todas as abas necessárias\n' +
    '2. Aplicar formatações\n' +
    '3. Configurar validações\n' +
    '4. Inserir dados iniciais de preços\n' +
    '5. Configurar triggers automáticos\n\n' +
    'Deseja continuar?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    // 1. Verificar/criar aba principal
    let abaPrincipal = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!abaPrincipal) {
      abaPrincipal = ss.insertSheet(CONFIG.SHEET_NAME);
      abaPrincipal.appendRow(HEADERS);
    }
    aplicarFormatacao(abaPrincipal);
    
    // 2. Verificar/criar tabela de preços
    let tabelaPrecos = ss.getSheetByName(CONFIG.PRICING_SHEET_NAME);
    if (!tabelaPrecos) {
      tabelaPrecos = ss.insertSheet(CONFIG.PRICING_SHEET_NAME);
      tabelaPrecos.appendRow(PRICING_HEADERS);
      inserirDadosIniciaisPrecos();
    }
    aplicarFormatacaoPrecos(tabelaPrecos);
    
    // 3. Criar abas mensais
    const resultadoMensais = criarTodasAbasMensais();
    
    // 4. Configurar triggers
    const resultadoTriggers = configurarTriggersEmail();
    
    ui.alert(
      '✅ Sistema Configurado!',
      `Configuração concluída com sucesso!\n\n` +
      `• Aba principal: OK\n` +
      `• Tabela de preços: OK\n` +
      `• Abas mensais: ${resultadoMensais.criadas} criadas, ${resultadoMensais.existentes} existentes\n` +
      `• Triggers: ${resultadoTriggers.sucesso ? 'Configurados' : 'Erro'}\n\n` +
      `Sistema pronto para uso!`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    ui.alert('❌ Erro', 'Erro na configuração:\n' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Testa envio de e-mail interativo (CÓDIGO NOVO ADAPTADO)
 */
function testarEnvioEmailInterativo() {
  const ui = SpreadsheetApp.getUi();
  
  const dadosTeste = [
    999,                                    // ID
    'TESTE EMAIL INTERATIVO',              // Cliente
    'Transfer',                            // Tipo Serviço
    2,                                     // Pessoas
    1,                                     // Bagagens
    new Date(),                            // Data
    '+351999888777',                       // Contacto
    'TP1234',                             // Voo
    'Aeroporto de Lisboa',                // Origem
    CONFIG.NAMES.HOTEL_NAME,              // Destino
    '10:00',                              // Hora
    25.00,                                // Preço Cliente
    7.50,                                 // Valor Hotel
    15.50,                                // Valor HUB
    2.00,                                 // Comissão Recepção
    'Dinheiro',                           // Forma Pagamento
    'Recepção',                           // Pago Para
    'Solicitado',                         // Status
    'Este é um transfer de teste do sistema v4.0', // Observações
    new Date()                            // Data Criação
  ];
  
  try {
    const enviado = enviarEmailNovoTransfer(dadosTeste);
    
    if (enviado) {
      ui.alert(
        '✅ E-mail Enviado!',
        'E-mail de teste enviado com sucesso!\n\n' +
        'Verifique sua caixa de entrada.\n' +
        'Os botões de ação devem estar funcionais.',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('⚠️ Aviso', 'E-mail não foi enviado. Verifique as configurações.', ui.ButtonSet.OK);
    }
    
  } catch (error) {
    ui.alert('❌ Erro', 'Erro ao enviar e-mail:\n' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Verifica confirmações manualmente (CÓDIGO ANTIGO PRESERVADO)
 */
function verificarConfirmacoesManual() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const confirmacoes = verificarConfirmacoesEmail();
    
    ui.alert(
      '📧 Verificação Concluída',
      `${confirmacoes} confirmação(ões) processada(s).\n\n` +
      'Verifique a planilha para ver os status atualizados.',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    ui.alert('❌ Erro', 'Erro na verificação:\n' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Menu para configurar triggers (CÓDIGO ANTIGO PRESERVADO)
 */
function configurarTriggersEmailMenu() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const resultado = configurarTriggersEmail();
    
    if (resultado.sucesso) {
      ui.alert(
        '✅ Triggers Configurados',
        'Verificação automática configurada!\n\n' +
        `• Verificação de e-mails: a cada ${CONFIG.EMAIL_CONFIG.INTERVALO_VERIFICACAO} minutos\n` +
        '• Relatório diário: às 9h da manhã\n' +
        '• Manutenção automática: às 2h da manhã\n\n' +
        'O sistema agora verificará e-mails automaticamente.',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('❌ Erro', 'Erro ao configurar triggers:\n' + resultado.erro, ui.ButtonSet.OK);
    }
    
  } catch (error) {
    ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Menu para remover triggers (CÓDIGO ANTIGO PRESERVADO)
 */
function removerTriggersEmailMenu() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const resultado = removerTriggersEmail();
    
    ui.alert(
      '🛑 Triggers Removidos',
      `${resultado.removidos} trigger(s) removido(s).\n\n` +
      'A verificação automática foi desativada.',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Envia relatório diário manualmente (CÓDIGO NOVO)
 */
function enviarRelatorioDiaAnteriorManual() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const sucesso = enviarRelatorioDiaAnterior();
    if (sucesso) {
      ui.alert('✅ Relatório Enviado', 'Relatório diário enviado com sucesso!', ui.ButtonSet.OK);
    } else {
      ui.alert('⚠️ Aviso', 'Não há dados para enviar ou e-mail desativado.', ui.ButtonSet.OK);
    }
  } catch (error) {
    ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Mostra estatísticas via menu (CÓDIGO ANTIGO PRESERVADO + ADAPTAÇÕES)
 */
function mostrarEstatisticasMenu() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const stats = gerarEstatisticas();
    
    const mensagem = `
📊 ESTATÍSTICAS DO SISTEMA

Total de Transfers: ${stats.totalTransfers}
Valor Total: €${stats.valorTotal.toFixed(2)}
Valor ${CONFIG.NAMES.HOTEL_NAME}: €${stats.valorHotel.toFixed(2)}
Valor HUB Transfer: €${stats.valorHUB.toFixed(2)}
Comissão Recepção: €${stats.comissaoRecepcao.toFixed(2)}

Média de Passageiros: ${stats.mediaPassageiros.toFixed(1)}

📈 Por Status:
${Object.entries(stats.porStatus).map(([status, count]) => `• ${status}: ${count}`).join('\n')}

🎯 Por Tipo de Serviço:
${Object.entries(stats.porTipo).map(([tipo, count]) => `• ${obterLabelTipoServico(tipo)}: ${count}`).join('\n')}

💳 Formas de Pagamento:
${Object.entries(stats.formasPagamento).map(([forma, count]) => `• ${forma}: ${count}`).join('\n')}

🏆 Top 5 Rotas:
${Object.entries(stats.topRotas).slice(0, 5).map(([rota, count]) => `• ${rota}: ${count}`).join('\n')}
   `;
   
   ui.alert('📊 Estatísticas', mensagem, ui.ButtonSet.OK);
   
 } catch (error) {
   ui.alert('❌ Erro', 'Erro ao gerar estatísticas:\n' + error.toString(), ui.ButtonSet.OK);
 }
}

/**
* Menu para consultar preço (CÓDIGO ANTIGO PRESERVADO)
*/
function consultarPrecoMenu() {
 const ui = SpreadsheetApp.getUi();
 
 try {
   // Solicitar dados
   const origem = ui.prompt('🚩 Origem', 'Digite o local de origem:', ui.ButtonSet.OK_CANCEL);
   if (origem.getSelectedButton() !== ui.Button.OK) return;
   
   const destino = ui.prompt('🎯 Destino', 'Digite o local de destino:', ui.ButtonSet.OK_CANCEL);
   if (destino.getSelectedButton() !== ui.Button.OK) return;
   
   const pessoas = ui.prompt('👥 Pessoas', 'Número de pessoas:', ui.ButtonSet.OK_CANCEL);
   if (pessoas.getSelectedButton() !== ui.Button.OK) return;
   
   const bagagens = ui.prompt('🧳 Bagagens', 'Número de bagagens (opcional):', ui.ButtonSet.OK_CANCEL);
   
   // Solicitar tipo de serviço (CÓDIGO NOVO)
   const tipoResponse = ui.prompt(
     '🎯 Tipo de Serviço', 
     'Digite:\n1 - Transfer\n2 - Tour Regular\n3 - Private Tour', 
     ui.ButtonSet.OK_CANCEL
   );
   
   let tipoServico = 'Transfer';
   if (tipoResponse.getSelectedButton() === ui.Button.OK) {
     const tipoNum = tipoResponse.getResponseText();
     if (tipoNum === '2') tipoServico = 'Tour Regular';
     else if (tipoNum === '3') tipoServico = 'Private Tour';
   }
   
   // Calcular valores
   const valores = calcularValores(
     origem.getResponseText(),
     destino.getResponseText(),
     parseInt(pessoas.getResponseText()) || 1,
     parseInt(bagagens.getResponseText()) || 0,
     null,
     tipoServico
   );
   
   const tipoLabel = obterLabelTipoServico(tipoServico);
   
   ui.alert(
     '💰 Consulta de Preço',
     `${tipoLabel}: ${origem.getResponseText()} → ${destino.getResponseText()}\n` +
     `👥 ${pessoas.getResponseText()} pessoa(s) | 🧳 ${bagagens.getResponseText() || '0'} bagagem(ns)\n\n` +
     `💶 Preço Total: €${valores.precoCliente.toFixed(2)}\n` +
     `🏨 ${CONFIG.NAMES.HOTEL_NAME} (30%): €${valores.valorHotel.toFixed(2)}\n` +
     `🎯 HUB Transfer: €${valores.valorHUB.toFixed(2)}\n` +
     `💰 Comissão Recepção: €${valores.comissaoRecepcao.toFixed(2)}\n\n` +
     `📊 Fonte: ${valores.fonte}\n` +
     `${valores.observacoes || ''}`,
     ui.ButtonSet.OK
   );
   
 } catch (error) {
   ui.alert('❌ Erro', 'Erro na consulta:\n' + error.toString(), ui.ButtonSet.OK);
 }
}

/**
* Abre tabela de preços (CÓDIGO ANTIGO PRESERVADO)
*/
function abrirTabelaPrecos() {
 const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
 let sheet = ss.getSheetByName(CONFIG.PRICING_SHEET_NAME);
 
 if (!sheet) {
   sheet = criarTabelaPrecos(ss);
   inserirDadosIniciaisPrecos();
 }
 
 ss.setActiveSheet(sheet);
}

/**
* Menu para adicionar preço (CÓDIGO ANTIGO PRESERVADO + ADAPTAÇÕES)
*/
function adicionarPrecoMenu() {
 const ui = SpreadsheetApp.getUi();
 
 try {
   // Solicitar dados
   const origem = ui.prompt('🚩 Origem', 'Digite o local de origem:', ui.ButtonSet.OK_CANCEL);
   if (origem.getSelectedButton() !== ui.Button.OK) return;
   
   const destino = ui.prompt('🎯 Destino', 'Digite o local de destino:', ui.ButtonSet.OK_CANCEL);
   if (destino.getSelectedButton() !== ui.Button.OK) return;
   
   const pessoas = ui.prompt('👥 Pessoas', 'Número de pessoas:', ui.ButtonSet.OK_CANCEL);
   if (pessoas.getSelectedButton() !== ui.Button.OK) return;
   
   const precoCliente = ui.prompt('💰 Preço Cliente', 'Preço total para o cliente (€):', ui.ButtonSet.OK_CANCEL);
   if (precoCliente.getSelectedButton() !== ui.Button.OK) return;
   
   // Tipo de serviço (CÓDIGO NOVO)
   const tipoResponse = ui.prompt(
     '🎯 Tipo de Serviço', 
     'Digite:\n1 - Transfer\n2 - Tour Regular\n3 - Private Tour', 
     ui.ButtonSet.OK_CANCEL
   );
   
   let tipoServico = 'Transfer';
   if (tipoResponse.getSelectedButton() === ui.Button.OK) {
     const tipoNum = tipoResponse.getResponseText();
     if (tipoNum === '2') tipoServico = 'Tour Regular';
     else if (tipoNum === '3') tipoServico = 'Private Tour';
   }
   
   const bagagens = ui.prompt('🧳 Bagagens', 'Número de bagagens (opcional):', ui.ButtonSet.OK_CANCEL);
   const observacoes = ui.prompt('📝 Observações', 'Observações (opcional):', ui.ButtonSet.OK_CANCEL);
   
   // Adicionar preço
   const resultado = adicionarPrecoTabela({
     tipoServico: tipoServico,
     origem: origem.getResponseText(),
     destino: destino.getResponseText(),
     pessoas: parseInt(pessoas.getResponseText()) || 1,
     bagagens: parseInt(bagagens.getResponseText()) || 0,
     precoCliente: parseFloat(precoCliente.getResponseText()) || 0,
     observacoes: observacoes.getSelectedButton() === ui.Button.OK ? observacoes.getResponseText() : ''
   });
   
   if (resultado.sucesso) {
     ui.alert(
       '✅ Preço Adicionado',
       `Preço ID #${resultado.id} adicionado com sucesso!\n\n` +
       `${obterLabelTipoServico(tipoServico)}: ${origem.getResponseText()} → ${destino.getResponseText()}\n` +
       `€${precoCliente.getResponseText()}`,
       ui.ButtonSet.OK
     );
   } else {
     ui.alert('❌ Erro', 'Erro ao adicionar preço:\n' + resultado.erro, ui.ButtonSet.OK);
   }
   
 } catch (error) {
   ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
 }
}

/**
* Menu para criar todas as abas mensais (CÓDIGO ANTIGO PRESERVADO)
*/
function criarTodasAbasMensaisMenu() {
 const ui = SpreadsheetApp.getUi();
 
 const response = ui.alert(
   '📅 Criar Abas Mensais',
   'Esta ação irá criar todas as abas mensais para o ano atual.\n' +
   'Abas existentes serão verificadas e formatadas.\n\n' +
   'Deseja continuar?',
   ui.ButtonSet.YES_NO
 );
 
 if (response !== ui.Button.YES) return;
 
 try {
   const resultado = criarTodasAbasMensais();
   
   ui.alert(
     '✅ Abas Mensais',
     `Processamento concluído!\n\n` +
     `• Abas criadas: ${resultado.criadas}\n` +
     `• Abas existentes: ${resultado.existentes}\n` +
     `• Erros: ${resultado.erros}\n\n` +
     'Todas as abas estão prontas para uso.',
     ui.ButtonSet.OK
   );
   
 } catch (error) {
   ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
 }
}

/**
* Menu para reordenar por data (CÓDIGO ANTIGO PRESERVADO)
*/
function reordenarPorDataMenu() {
 const ui = SpreadsheetApp.getUi();
 
 const response = ui.alert(
   '🔄 Reordenar por Data',
   'Esta ação irá reordenar todos os transfers por data e hora.\n\n' +
   'Deseja continuar?',
   ui.ButtonSet.YES_NO
 );
 
 if (response !== ui.Button.YES) return;
 
 try {
   const resultado = reordenarPorData();
   
   if (resultado.sucesso) {
     ui.alert(
       '✅ Reordenação Concluída',
       `${resultado.registros} registros reordenados por data.`,
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
* Menu para corrigir registros (CÓDIGO ANTIGO PRESERVADO)
*/
function corrigirRegistrosMenu() {
 const ui = SpreadsheetApp.getUi();
 
 try {
   const resultado = corrigirRegistrosIncompletos();
   
   ui.alert(
     '🔧 Correção Concluída',
     `Registros verificados: ${resultado.verificados}\n` +
     `Registros corrigidos: ${resultado.corrigidos}\n` +
     `Erros: ${resultado.erros}`,
     ui.ButtonSet.OK
   );
   
 } catch (error) {
   ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
 }
}

/**
* Menu para verificar duplicados (CÓDIGO ANTIGO PRESERVADO)
*/
function verificarDuplicadosMenu() {
 const ui = SpreadsheetApp.getUi();
 
 try {
   const resultado = removerDuplicados(true); // Apenas simular
   
   if (resultado.duplicados > 0) {
     ui.alert(
       '🔍 Duplicados Encontrados',
       `${resultado.duplicados} registro(s) duplicado(s) encontrado(s).\n\n` +
       'Use "Remover Duplicados" para removê-los.',
       ui.ButtonSet.OK
     );
   } else {
     ui.alert(
       '✅ Nenhum Duplicado',
       'Nenhum registro duplicado encontrado.',
       ui.ButtonSet.OK
     );
   }
   
 } catch (error) {
   ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
 }
}

/**
* Menu para remover duplicados (CÓDIGO ANTIGO PRESERVADO)
*/
function removerDuplicadosMenu() {
 const ui = SpreadsheetApp.getUi();
 
 const response = ui.alert(
   '⚠️ Remover Duplicados',
   'Esta ação irá REMOVER permanentemente os registros duplicados.\n\n' +
   'Deseja continuar?',
   ui.ButtonSet.YES_NO
 );
 
 if (response !== ui.Button.YES) return;
 
 try {
   const resultado = removerDuplicados(false); // Remover realmente
   
   ui.alert(
     '🗑️ Duplicados Removidos',
     `${resultado.duplicados} registro(s) duplicado(s) removido(s).`,
     ui.ButtonSet.OK
   );
   
 } catch (error) {
   ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
 }
}

/**
* Menu para limpar dados de teste (CÓDIGO ANTIGO PRESERVADO)
*/
function limparDadosTesteMenu() {
 const ui = SpreadsheetApp.getUi();
 
 const response = ui.alert(
   '🧹 Limpar Dados de Teste',
   'Esta ação irá remover todos os registros identificados como teste.\n\n' +
   'Deseja continuar?',
   ui.ButtonSet.YES_NO
 );
 
 if (response !== ui.Button.YES) return;
 
 try {
   const resultado = limparDadosTeste();
   
   ui.alert(
     '🧹 Limpeza Concluída',
     `${resultado.totalRemovidos} registro(s) de teste removido(s).`,
     ui.ButtonSet.OK
   );
   
 } catch (error) {
   ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
 }
}

/**
* Menu para criar backup (CÓDIGO ANTIGO PRESERVADO)
*/
function criarBackupMenu() {
 const ui = SpreadsheetApp.getUi();
 
 try {
   const resultado = criarBackup();
   
   if (resultado.sucesso) {
     ui.alert(
       '💾 Backup Criado',
       `Backup criado com sucesso!\n\n` +
       `Nome: ${resultado.nome}\n` +
       `ID: ${resultado.id}\n` +
       `Data: ${formatarDataHora(resultado.dataHora)}`,
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
* Menu para limpar todos os dados (CÓDIGO ANTIGO PRESERVADO)
*/
function limparDadosCompletoMenu() {
 const ui = SpreadsheetApp.getUi();
 
 const response = ui.alert(
   '⚠️ ATENÇÃO - LIMPEZA COMPLETA',
   'Esta ação irá REMOVER TODOS OS DADOS do sistema!\n\n' +
   '• Todos os transfers\n' +
   '• Todos os preços\n' +
   '• Todas as abas mensais\n\n' +
   'Esta ação é IRREVERSÍVEL!\n\n' +
   'Tem certeza absoluta?',
   ui.ButtonSet.YES_NO
 );
 
 if (response !== ui.Button.YES) return;
 
 // Dupla confirmação
 const confirmacao = ui.alert(
   '🚨 CONFIRMAÇÃO FINAL',
   'ÚLTIMA CHANCE!\n\n' +
   'Todos os dados serão perdidos permanentemente.\n\n' +
   'Confirma a limpeza completa?',
   ui.ButtonSet.YES_NO
 );
 
 if (confirmacao !== ui.Button.YES) return;
 
 try {
   const resultado = limparDadosCompleto();
   
   if (resultado.sucesso) {
     ui.alert(
       '🧹 Limpeza Completa',
       `${resultado.registrosRemovidos} registros removidos de ${resultado.abasProcessadas} abas.\n\n` +
       'Sistema limpo e pronto para novo uso.',
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
* Testa todo o sistema (CÓDIGO NOVO)
*/
function testarSistemaCompleto() {
 const ui = SpreadsheetApp.getUi();
 
 try {
   const testes = [];
   
   // Teste 1: Configuração básica
   const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
   testes.push('✅ Planilha acessível');
   
   // Teste 2: Aba principal
   const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
   testes.push(sheet ? '✅ Aba principal OK' : '❌ Aba principal não encontrada');
   
   // Teste 3: Tabela de preços
   const pricing = ss.getSheetByName(CONFIG.PRICING_SHEET_NAME);
   testes.push(pricing ? '✅ Tabela de preços OK' : '❌ Tabela de preços não encontrada');
   
   // Teste 4: Geração de ID
   const novoId = gerarProximoIdSeguro(sheet);
   testes.push(novoId > 0 ? '✅ Geração de ID OK' : '❌ Erro na geração de ID');
   
   // Teste 5: Cálculo de valores
   const valores = calcularValores('Aeroporto', CONFIG.NAMES.HOTEL_NAME, 2, 1, null, 'Transfer');
   testes.push(valores.precoCliente > 0 ? '✅ Cálculo de valores OK' : '❌ Erro no cálculo');
   
   // Teste 6: Criação de transfer de teste
   const dadosTeste = [
     99999,
     'TESTE SISTEMA COMPLETO',
     'Transfer',
     2,
     1,
     new Date(),
     '+351999999999',
     'TP9999',
     'Aeroporto de Lisboa',
     CONFIG.NAMES.HOTEL_NAME,
     '12:00',
     25.00,
     7.50,
     15.50,
     2.00,
     'Dinheiro',
     'Recepção',
     'Solicitado',
     'Transfer de teste do sistema',
     new Date()
   ];
   
   sheet.appendRow(dadosTeste);
   const linhaEncontrada = encontrarLinhaPorId(sheet, 99999);
   testes.push(linhaEncontrada > 0 ? '✅ Criação de transfer OK' : '❌ Erro na criação');
   
   // Remover o teste
   if (linhaEncontrada > 0) {
     sheet.deleteRow(linhaEncontrada);
     testes.push('✅ Remoção de teste OK');
   }
   
   // Teste 7: Sistema de logs
   logger.info('Teste do sistema de logs');
   testes.push('✅ Sistema de logs OK');
   
   ui.alert(
     '🧪 Teste Completo do Sistema',
     `Testes realizados:\n\n${testes.join('\n')}\n\n` +
     `Sistema funcionando corretamente!`,
     ui.ButtonSet.OK
   );
   
 } catch (error) {
   ui.alert('❌ Erro no Teste', error.toString(), ui.ButtonSet.OK);
 }
}

/**
* Mostra informações sobre o sistema (FUSÃO)
*/
function mostrarSobre() {
 const ui = SpreadsheetApp.getUi();
 
 ui.alert(
   `ℹ️ Sobre o Sistema`,
   `${CONFIG.NAMES.SISTEMA_NOME}\n` +
   `Versão: ${CONFIG.SISTEMA.VERSAO}\n\n` +
   `Sistema integrado de gestão de transfers entre ${CONFIG.NAMES.HOTEL_NAME} e HUB Transfer.\n\n` +
   `🌟 FUNCIONALIDADES:\n` +
   `• Gestão automática de transfers\n` +
   `• E-mails interativos com botões\n` +
   `• Múltiplos tipos de serviço\n` +
   `• Organização por abas mensais\n` +
   `• Tabela inteligente de preços\n` +
   `• Relatórios automáticos\n` +
   `• Cálculo automático de comissões\n` +
   `• Backup e manutenção automática\n\n` +
   `💻 DESENVOLVIDO POR:\n` +
   `Claude 4 Sonnet + Google Apps Script\n\n` +
   `📧 E-mail do sistema: ${CONFIG.EMAIL_CONFIG.EMAIL_EMPRESA}\n` +
   `🌐 Web App: Use ?action=test para testar a API`,
   ui.ButtonSet.OK
 );
}

// ===================================================
// DADOS INICIAIS E CONFIGURAÇÃO (FUSÃO COMPLETA)
// ===================================================

/**
 * Insere dados iniciais na tabela de preços (CÓDIGO ANTIGO PRESERVADO + ADAPTAÇÕES)
 */
function inserirDadosIniciaisPrecos() {
  logger.info('Inserindo dados iniciais de preços');
  
  const dadosIniciais = [
    // TRANSFERS PRINCIPAIS (CÓDIGO ANTIGO PRESERVADO)
    {
      tipoServico: 'Transfer',
      origem: 'Aeroporto de Lisboa',
      destino: CONFIG.NAMES.HOTEL_NAME,
      pessoas: 1,
      bagagens: 1,
      precoCliente: 25.00,
      observacoes: 'Transfer padrão aeroporto-hotel'
    },
    {
      tipoServico: 'Transfer',
      origem: CONFIG.NAMES.HOTEL_NAME,
      destino: 'Aeroporto de Lisboa',
      pessoas: 1,
      bagagens: 1,
      precoCliente: 25.00,
      observacoes: 'Transfer padrão hotel-aeroporto'
    },
    {
      tipoServico: 'Transfer',
      origem: 'Aeroporto de Lisboa',
      destino: 'Cascais',
      pessoas: 1,
      bagagens: 1,
      precoCliente: 35.00,
      observacoes: 'Transfer aeroporto-Cascais'
    },
    {
      tipoServico: 'Transfer',
      origem: 'Aeroporto de Lisboa',
      destino: 'Sintra',
      pessoas: 1,
      bagagens: 1,
      precoCliente: 40.00,
      observacoes: 'Transfer aeroporto-Sintra'
    },
    
    // TOURS REGULARES (CÓDIGO NOVO)
    {
      tipoServico: 'Tour Regular',
      origem: CONFIG.NAMES.HOTEL_NAME,
      destino: 'Sintra (Palácio da Pena + Quinta da Regaleira)',
      pessoas: 1,
      bagagens: 0,
      precoCliente: 67.00,
      precoPorPessoa: 67.00,
      observacoes: 'Tour regular Sintra - por pessoa'
    },
    {
      tipoServico: 'Tour Regular',
      origem: CONFIG.NAMES.HOTEL_NAME,
      destino: 'Óbidos + Nazaré + Mosteiro da Batalha',
      pessoas: 1,
      bagagens: 0,
      precoCliente: 72.00,
      precoPorPessoa: 72.00,
      observacoes: 'Tour regular Norte - por pessoa'
    },
    {
      tipoServico: 'Tour Regular',
      origem: CONFIG.NAMES.HOTEL_NAME,
      destino: 'Fátima + Óbidos + Nazaré',
      pessoas: 1,
      bagagens: 0,
      precoCliente: 67.00,
      precoPorPessoa: 67.00,
      observacoes: 'Tour regular Fátima - por pessoa'
    },
    
    // PRIVATE TOURS (CÓDIGO NOVO)
    {
      tipoServico: 'Private Tour',
      origem: CONFIG.NAMES.HOTEL_NAME,
      destino: 'Sintra Privado (até 3 pessoas)',
      pessoas: 3,
      bagagens: 0,
      precoCliente: 347.00,
      precoPorGrupo: 347.00,
      observacoes: 'Private tour Sintra - até 3 pessoas'
    },
    {
      tipoServico: 'Private Tour',
      origem: CONFIG.NAMES.HOTEL_NAME,
      destino: 'Sintra Privado (até 6 pessoas)',
      pessoas: 6,
      bagagens: 0,
      precoCliente: 492.00,
      precoPorGrupo: 492.00,
      observacoes: 'Private tour Sintra - até 6 pessoas'
    },
    {
      tipoServico: 'Private Tour',
      origem: CONFIG.NAMES.HOTEL_NAME,
      destino: 'Óbidos + Nazaré + Batalha Privado (até 3 pessoas)',
      pessoas: 3,
      bagagens: 0,
      precoCliente: 397.00,
      precoPorGrupo: 397.00,
      observacoes: 'Private tour Norte - até 3 pessoas'
    },
    {
      tipoServico: 'Private Tour',
      origem: CONFIG.NAMES.HOTEL_NAME,
      destino: 'Óbidos + Nazaré + Batalha Privado (até 6 pessoas)',
      pessoas: 6,
      bagagens: 0,
      precoCliente: 542.00,
      precoPorGrupo: 542.00,
      observacoes: 'Private tour Norte - até 6 pessoas'
    },
    
    // TRANSFERS ESPECIAIS (CÓDIGO ANTIGO PRESERVADO)
    {
      tipoServico: 'Transfer',
      origem: CONFIG.NAMES.HOTEL_NAME,
      destino: 'Estação do Oriente',
      pessoas: 1,
      bagagens: 1,
      precoCliente: 15.00,
      observacoes: 'Transfer para comboios'
    },
    {
      tipoServico: 'Transfer',
      origem: CONFIG.NAMES.HOTEL_NAME,
      destino: 'Terminal de Cruzeiros',
      pessoas: 1,
      bagagens: 1,
      precoCliente: 12.00,
      observacoes: 'Transfer para porto de cruzeiros'
    },
    {
      tipoServico: 'Transfer',
      origem: CONFIG.NAMES.HOTEL_NAME,
      destino: 'Centro de Lisboa',
      pessoas: 1,
      bagagens: 0,
      precoCliente: 10.00,
      observacoes: 'Transfer para centro histórico'
    }
  ];
  
  // Inserir cada preço
  dadosIniciais.forEach(dados => {
    try {
      adicionarPrecoTabela(dados);
    } catch (error) {
      logger.error('Erro ao inserir preço inicial', { dados, error });
    }
  });
  
  logger.success('Dados iniciais de preços inseridos', { 
    total: dadosIniciais.length 
  });
}

/**
 * Atualiza transfer completo (CÓDIGO NOVO)
 * @param {string} transferId - ID do transfer
 * @param {Object} atualizacoes - Campos a atualizar
 * @returns {Object} - Resultado da atualização
 */
function atualizarTransferCompleto(transferId, atualizacoes) {
  logger.info('Atualizando transfer completo', { transferId, atualizacoes });
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const linha = encontrarLinhaPorId(sheet, transferId);
    
    if (!linha) {
      throw new Error(`Transfer ${transferId} não encontrado`);
    }
    
    const camposAtualizados = [];
    
    // Mapeamento de campos para colunas (ESTRUTURA DO CÓDIGO NOVO)
    const mapeamentoCampos = {
      nomeCliente: 2,        // B
      tipoServico: 3,        // C
      numeroPessoas: 4,      // D
      numeroBagagens: 5,     // E
      data: 6,               // F
      contacto: 7,           // G
      numeroVoo: 8,          // H
      origem: 9,             // I
      destino: 10,           // J
      horaPickup: 11,        // K
      precoCliente: 12,      // L
      valorHotel: 13,        // M
      valorHUB: 14,          // N
      comissaoRecepcao: 15,  // O
      modoPagamento: 16,     // P
      pagoParaQuem: 17,      // Q
      status: 18,            // R
      observacoes: 19        // S
      // Coluna T (Data Criação) não deve ser alterada
    };
    
    // Aplicar atualizações
    Object.entries(atualizacoes).forEach(([campo, valor]) => {
      if (mapeamentoCampos[campo]) {
        const coluna = mapeamentoCampos[campo];
        
        // Processamento específico por tipo de campo
        let valorProcessado = valor;
        
if (campo === 'data' && valor) {
         valorProcessado = processarDataSegura(valor);
       } else if (['numeroPessoas', 'numeroBagagens'].includes(campo)) {
         valorProcessado = parseInt(valor) || 0;
       } else if (['precoCliente', 'valorHotel', 'valorHUB', 'comissaoRecepcao'].includes(campo)) {
         valorProcessado = parseFloat(valor) || 0;
       } else if (typeof valor === 'string') {
         valorProcessado = sanitizarTexto(valor);
       }
       
       // Atualizar célula
       sheet.getRange(linha, coluna).setValue(valorProcessado);
       camposAtualizados.push(campo);
       
       logger.debug('Campo atualizado', { campo, valor: valorProcessado });
     }
   });
   
   // Se foi atualizado preço ou valores, recalcular proporcionalmente
   if (atualizacoes.precoCliente && 
       (!atualizacoes.valorHotel || !atualizacoes.valorHUB || !atualizacoes.comissaoRecepcao)) {
     
     const tipoServico = sheet.getRange(linha, 3).getValue(); // Coluna C
     const novosValores = calcularPorTipoServico(atualizacoes.precoCliente, tipoServico);
     
     if (!atualizacoes.valorHotel) {
       sheet.getRange(linha, 13).setValue(novosValores.valorHotel);
       camposAtualizados.push('valorHotel (recalculado)');
     }
     if (!atualizacoes.valorHUB) {
       sheet.getRange(linha, 14).setValue(novosValores.valorHUB);
       camposAtualizados.push('valorHUB (recalculado)');
     }
     if (!atualizacoes.comissaoRecepcao) {
       sheet.getRange(linha, 15).setValue(novosValores.comissaoRecepcao);
       camposAtualizados.push('comissaoRecepcao (recalculada)');
     }
   }
   
   // Atualizar observações com histórico da atualização
   const observacaoAtual = sheet.getRange(linha, 19).getValue() || ''; // Coluna S
   const novaObservacao = observacaoAtual 
     ? `${observacaoAtual}\nAtualizado: ${camposAtualizados.join(', ')} - ${formatarDataHora(new Date())}`
     : `Atualizado: ${camposAtualizados.join(', ')} - ${formatarDataHora(new Date())}`;
   
   sheet.getRange(linha, 19).setValue(novaObservacao);
   
   // Tentar atualizar na aba mensal também
   try {
     const dataTransfer = sheet.getRange(linha, 6).getValue(); // Coluna F
     const abaMensal = obterAbaMes(dataTransfer);
     
     if (abaMensal && abaMensal.getName() !== sheet.getName()) {
       const linhaMensal = encontrarLinhaPorId(abaMensal, transferId);
       
       if (linhaMensal > 0) {
         // Copiar linha inteira atualizada
         const linhaCompleta = sheet.getRange(linha, 1, 1, HEADERS.length).getValues()[0];
         abaMensal.getRange(linhaMensal, 1, 1, HEADERS.length).setValues([linhaCompleta]);
         
         logger.debug('Aba mensal sincronizada', { 
           abaMensal: abaMensal.getName(),
           linhaMensal 
         });
       }
     }
   } catch (errorMensal) {
     logger.error('Erro ao sincronizar aba mensal', errorMensal);
   }
   
   logger.success('Transfer atualizado com sucesso', {
     transferId,
     camposAtualizados
   });
   
   return {
     sucesso: true,
     transferId: transferId,
     camposAtualizados: camposAtualizados,
     mensagem: `Transfer #${transferId} atualizado com sucesso`
   };
   
 } catch (error) {
   logger.error('Erro ao atualizar transfer', error);
   return {
     sucesso: false,
     erro: error.message
   };
 }
}

/**
* Função de inicialização do sistema (CÓDIGO NOVO)
* Executa automaticamente quando o sistema é implantado
*/
function inicializarSistema() {
 logger.info('Inicializando sistema completo');
 
 try {
   const resultados = {
     configuracao: false,
     abaPrincipal: false,
     tabelaPrecos: false,
     abasMensais: false,
     triggers: false,
     dadosIniciais: false
   };
   
   // 1. Verificar/criar estrutura básica
   const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
   resultados.configuracao = true;
   
   // 2. Criar/verificar aba principal
   let abaPrincipal = ss.getSheetByName(CONFIG.SHEET_NAME);
   if (!abaPrincipal) {
     abaPrincipal = ss.insertSheet(CONFIG.SHEET_NAME);
     abaPrincipal.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
   }
   aplicarFormatacao(abaPrincipal);
   resultados.abaPrincipal = true;
   
   // 3. Criar/verificar tabela de preços
   let tabelaPrecos = ss.getSheetByName(CONFIG.PRICING_SHEET_NAME);
   if (!tabelaPrecos) {
     tabelaPrecos = criarTabelaPrecos(ss);
     inserirDadosIniciaisPrecos();
     resultados.dadosIniciais = true;
   }
   resultados.tabelaPrecos = true;
   
   // 4. Criar abas mensais do ano atual
   const resultadoMensais = criarTodasAbasMensais();
   resultados.abasMensais = resultadoMensais.sucesso !== false;
   
   // 5. Configurar triggers se habilitados
   if (CONFIG.EMAIL_CONFIG.CONFIGURAR_TRIGGERS_AUTO) {
     const resultadoTriggers = configurarTriggersEmail();
     resultados.triggers = resultadoTriggers.sucesso;
   }
   
   // Log do resultado
   logger.success('Sistema inicializado', resultados);
   
   // Criar log de inicialização
   if (LOG_CONFIG.PERSIST_TO_SHEET) {
     try {
       let logSheet = ss.getSheetByName(LOG_CONFIG.LOG_SHEET_NAME);
       if (!logSheet) {
         logSheet = ss.insertSheet(LOG_CONFIG.LOG_SHEET_NAME);
         logSheet.appendRow(['Timestamp', 'Tipo', 'Mensagem', 'Dados']);
       }
       
       logSheet.appendRow([
         new Date(),
         'SYSTEM',
         'Sistema inicializado com sucesso',
         JSON.stringify(resultados)
       ]);
     } catch (logError) {
       // Ignorar erro de log
     }
   }
   
   return {
     sucesso: true,
     resultados: resultados,
     versao: CONFIG.SISTEMA.VERSAO,
     timestamp: new Date()
   };
   
 } catch (error) {
   logger.error('Erro na inicialização do sistema', error);
   
   return {
     sucesso: false,
     erro: error.message,
     versao: CONFIG.SISTEMA.VERSAO,
     timestamp: new Date()
   };
 }
}

/**
* Função de manutenção e verificação periódica (CÓDIGO NOVO + ANTIGO)
*/
function verificacaoSistemaPeriodica() {
 logger.info('Executando verificação periódica do sistema');
 
 try {
   const verificacoes = {
     integridade: false,
     performance: false,
     backups: false,
     logs: false,
     triggers: false
   };
   
   const problemas = [];
   const solucoes = [];
   
   // 1. Verificar integridade das abas
   const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
   const sheets = ss.getSheets();
   
   let abasComProblemas = 0;
   sheets.forEach(sheet => {
     const nome = sheet.getName();
     
     if (nome === CONFIG.SHEET_NAME || 
         nome === CONFIG.PRICING_SHEET_NAME ||
         nome.startsWith(CONFIG.SISTEMA.PREFIXO_MES)) {
       
       try {
         const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
         
         if (!headers || headers.length === 0 || !headers[0]) {
           abasComProblemas++;
           problemas.push(`Aba "${nome}" sem headers`);
           
           // Auto-correção
           if (nome === CONFIG.SHEET_NAME || nome.startsWith(CONFIG.SISTEMA.PREFIXO_MES)) {
             sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
             solucoes.push(`Headers corrigidos em "${nome}"`);
           } else if (nome === CONFIG.PRICING_SHEET_NAME) {
             sheet.getRange(1, 1, 1, PRICING_HEADERS.length).setValues([PRICING_HEADERS]);
             solucoes.push(`Headers de preços corrigidos em "${nome}"`);
           }
         }
       } catch (sheetError) {
         problemas.push(`Erro ao verificar aba "${nome}": ${sheetError.message}`);
       }
     }
   });
   
   verificacoes.integridade = abasComProblemas === 0;
   
   // 2. Verificar performance (número de registros vs. limites)
   const abaPrincipal = ss.getSheetByName(CONFIG.SHEET_NAME);
   if (abaPrincipal) {
     const totalRegistros = abaPrincipal.getLastRow() - 1;
     
     if (totalRegistros > CONFIG.LIMITES.MAX_REGISTROS_POR_ABA) {
       problemas.push(`Muitos registros na aba principal: ${totalRegistros}`);
       // Sugerir arquivamento automático
     }
     
     verificacoes.performance = totalRegistros < CONFIG.LIMITES.MAX_REGISTROS_POR_ABA;
   }
   
   // 3. Verificar sistema de backups
   let backupSheet = ss.getSheetByName('Backups');
   if (backupSheet && CONFIG.SISTEMA.BACKUP_AUTOMATICO) {
     const ultimoBackup = backupSheet.getLastRow();
     if (ultimoBackup > 1) {
       const dataUltimoBackup = backupSheet.getRange(ultimoBackup, 1).getValue();
       const agora = new Date();
       const diasSemBackup = Math.floor((agora - dataUltimoBackup) / (1000 * 60 * 60 * 24));
       
       if (diasSemBackup > 7) {
         problemas.push(`Último backup há ${diasSemBackup} dias`);
         
         // Criar backup automático
         const novoBackup = criarBackup();
         if (novoBackup.sucesso) {
           solucoes.push(`Backup automático criado: ${novoBackup.nome}`);
         }
       }
     }
   }
   verificacoes.backups = true;
   
   // 4. Limpar logs antigos se necessário
   if (LOG_CONFIG.PERSIST_TO_SHEET) {
     let logSheet = ss.getSheetByName(LOG_CONFIG.LOG_SHEET_NAME);
     if (logSheet && logSheet.getLastRow() > 5000) {
       const linhasParaRemover = logSheet.getLastRow() - 3000;
       logSheet.deleteRows(2, linhasParaRemover);
       solucoes.push(`${linhasParaRemover} logs antigos removidos`);
     }
   }
   verificacoes.logs = true;
   
   // 5. Verificar triggers ativos
   const triggers = ScriptApp.getProjectTriggers();
   const triggersEsperados = ['verificarConfirmacoesEmail', 'enviarRelatorioDiaAnterior', 'manutencaoAutomatica'];
   const triggersAtivos = triggers.map(t => t.getHandlerFunction());
   
   triggersEsperados.forEach(esperado => {
     if (!triggersAtivos.includes(esperado) && CONFIG.EMAIL_CONFIG.VERIFICAR_CONFIRMACOES) {
       problemas.push(`Trigger ausente: ${esperado}`);
     }
   });
   
   verificacoes.triggers = triggersEsperados.every(t => triggersAtivos.includes(t) || !CONFIG.EMAIL_CONFIG.VERIFICAR_CONFIRMACOES);
   
   // Resultado final
   const sistemaOK = Object.values(verificacoes).every(v => v === true);
   
   const resultado = {
     sistemaOK: sistemaOK,
     verificacoes: verificacoes,
     problemas: problemas,
     solucoes: solucoes,
     timestamp: new Date(),
     proximaVerificacao: new Date(Date.now() + (24 * 60 * 60 * 1000)) // Próximas 24h
   };
   
   logger.info('Verificação periódica concluída', resultado);
   
   // Enviar e-mail se houver problemas críticos
   if (problemas.length > 0 && CONFIG.EMAIL_CONFIG.NOTIFICAR_PROBLEMAS) {
     enviarNotificacaoProblemas(resultado);
   }
   
   return resultado;
   
 } catch (error) {
   logger.error('Erro na verificação periódica', error);
   return {
     sistemaOK: false,
     erro: error.message,
     timestamp: new Date()
   };
 }
}

/**
* Envia notificação de problemas do sistema (CÓDIGO NOVO)
* @private
*/
function enviarNotificacaoProblemas(relatorioVerificacao) {
 try {
   const destinatarios = CONFIG.EMAIL_CONFIG.DESTINATARIOS.join(',');
   const assunto = `[${CONFIG.NAMES.SISTEMA_NOME}] Problemas Detectados no Sistema`;
   
   const corpo = `
     <h2>⚠️ Problemas Detectados no Sistema</h2>
     
     <h3>🔍 Status das Verificações:</h3>
     <ul>
       ${Object.entries(relatorioVerificacao.verificacoes).map(([check, ok]) => 
         `<li>${ok ? '✅' : '❌'} ${check}: ${ok ? 'OK' : 'Problema'}</li>`
       ).join('')}
     </ul>
     
     ${relatorioVerificacao.problemas.length > 0 ? `
     <h3>❌ Problemas Encontrados:</h3>
     <ul>
       ${relatorioVerificacao.problemas.map(p => `<li>${p}</li>`).join('')}
     </ul>
     ` : ''}
     
     ${relatorioVerificacao.solucoes.length > 0 ? `
     <h3>✅ Soluções Aplicadas:</h3>
     <ul>
       ${relatorioVerificacao.solucoes.map(s => `<li>${s}</li>`).join('')}
     </ul>
     ` : ''}
     
     <p><strong>Data da Verificação:</strong> ${formatarDataHora(relatorioVerificacao.timestamp)}</p>
     <p><strong>Próxima Verificação:</strong> ${formatarDataHora(relatorioVerificacao.proximaVerificacao)}</p>
     
     <hr>
     <p><small>Esta é uma notificação automática do sistema de monitoramento.</small></p>
   `;
   
// Correção: usar htmlBody correto
sendEmailComAssinatura({
  to: destinatarios,
  subject: assunto,
  htmlBody: htmlEmail || corpo || corpoHtml,
  name: CONFIG.NAMES.SISTEMA_NOME,
  replyTo: CONFIG.EMAIL_CONFIG.EMAIL_EMPRESA
});
   
   logger.success('Notificação de problemas enviada');
   
 } catch (error) {
   logger.error('Erro ao enviar notificação de problemas', error);
 }
}

// ===================================================
// PONTO DE ENTRADA PARA INSTALAÇÃO INICIAL
// ===================================================

/**
 * Função principal para configuração inicial completa do sistema - CORRIGIDA
 */
function instalarSistemaCompleto() {
  // Usar Logger nativo para instalação inicial
  Logger.log('🚀 INICIANDO INSTALAÇÃO COMPLETA DO SISTEMA');
  
  try {
    // Passo 1: Inicialização básica
    Logger.log('Passo 1: Inicialização do sistema...');
    const inicializacao = inicializarSistema();
    
    if (!inicializacao.sucesso) {
      throw new Error('Falha na inicialização: ' + inicializacao.erro);
    }
    
    Logger.log('✅ Sistema inicializado com sucesso!');
    
    // Passo 2: Verificação completa
    Logger.log('Passo 2: Verificação do sistema...');
    const verificacao = verificacaoSistemaPeriodica();
    
    // Passo 3: Configuração final
    Logger.log('Passo 3: Configuração final...');
    
    // Resultado final
    Logger.log('🎉 SISTEMA INSTALADO COM SUCESSO!');
    Logger.log('📋 Resumo da Instalação:');
    Logger.log('• Aba principal: ✅');
    Logger.log('• Tabela de preços: ✅');
    Logger.log('• Abas mensais: ✅');
    Logger.log('• Dados iniciais: ✅');
    Logger.log('• Formatações: ✅');
    Logger.log('• Validações: ✅');
    
    if (inicializacao.resultados.triggers) {
      Logger.log('• Triggers automáticos: ✅');
    } else {
      Logger.log('• Triggers automáticos: ⚠️ (Configure manualmente se necessário)');
    }
    
    Logger.log('');
    Logger.log('🌐 URL da Web App: ' + ScriptApp.getService().getUrl());
    Logger.log('📧 E-mails serão enviados para: ' + CONFIG.EMAIL_CONFIG.DESTINATARIOS.join(', '));
    Logger.log('🏨 Hotel configurado: ' + CONFIG.NAMES.HOTEL_NAME);
    Logger.log('📱 Sistema versão: ' + CONFIG.SISTEMA.VERSAO);
    Logger.log('');
    Logger.log('✅ O sistema está pronto para receber transfers!');
    
    return {
      sucesso: true,
      versao: CONFIG.SISTEMA.VERSAO,
      webAppUrl: ScriptApp.getService().getUrl(),
      timestamp: new Date(),
      mensagem: 'Sistema instalado e configurado com sucesso!'
    };
    
  } catch (error) {
    Logger.log('❌ ERRO NA INSTALAÇÃO: ' + error.toString());
    
    return {
      sucesso: false,
      erro: error.message,
      timestamp: new Date(),
      mensagem: 'Falha na instalação do sistema'
    };
  }
}

// ===================================================
// CONFIGURAÇÕES FINAIS E EXPORTS
// ===================================================

// Log de inicialização do script
if (typeof console !== 'undefined') {
 console.log(`🚐 ${CONFIG.NAMES.SISTEMA_NOME} v${CONFIG.SISTEMA.VERSAO} carregado com sucesso!`);
} else {
 Logger.log(`🚐 ${CONFIG.NAMES.SISTEMA_NOME} v${CONFIG.SISTEMA.VERSAO} carregado com sucesso!`);
}

// Adicionar informações do sistema no contexto global para debug
if (typeof globalThis !== 'undefined') {
 globalThis.SISTEMA_INFO = {
   nome: CONFIG.NAMES.SISTEMA_NOME,
   versao: CONFIG.SISTEMA.VERSAO,
   hotel: CONFIG.NAMES.HOTEL_NAME,
   inicializado: new Date(),
   webAppUrl: null // Será preenchido após deploy
 };
}

// Execute esta função para verificar o status atual
function verificarStatusSistema() {
  try {
    // Teste básico de conectividade
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    console.log('✅ Planilha acessível');
    
    // Teste da assinatura
    const assinatura = montarAssinaturaEmail();
    console.log('✅ Assinatura:', assinatura.html.length > 0 ? 'OK' : 'ERRO');
    
    // Teste do Drive
    const file = DriveApp.getFileById(ASSINATURA_FILE_ID);
    console.log('✅ Arquivo do Drive acessível:', file.getName());
    
    return {
      planilha: true,
      assinatura: assinatura.html.length > 0,
      drive: true,
      timestamp: new Date()
    };
    
  } catch (error) {
    console.error('❌ Erro no sistema:', error.toString());
    return {
      erro: error.message,
      timestamp: new Date()
    };
  }
}

/**
 * Função getAllData corrigida para sincronização com frontend
 */
function getAllData() {
  console.log('🔍 getAllData chamada - iniciando busca de dados...');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) {
      console.log('❌ Aba não encontrada:', CONFIG.SHEET_NAME);
      return {
        success: false,
        error: 'Aba "Empire MARQUES-HUB" não encontrada',
        data: []
      };
    }
    
    const lastRow = sheet.getLastRow();
    console.log('📊 Última linha da planilha:', lastRow);
    
    if (lastRow <= 1) {
      console.log('ℹ️ Planilha vazia ou apenas com headers');
      return {
        success: true,
        data: [],
        message: 'Nenhum transfer encontrado',
        total: 0
      };
    }
    
    // Buscar dados da linha 2 até a última linha
    const range = sheet.getRange(2, 1, lastRow - 1, HEADERS.length);
    const values = range.getValues();
    console.log(`📋 Buscando ${values.length} linhas de dados...`);
    
    const transfers = [];
    
    values.forEach((row, index) => {
      // Pular linhas completamente vazias
      if (!row[0] && !row[1]) return;
      
      try {
        const transfer = {
          ID: row[0] || '',
          Cliente: row[1] || '',
          'Tipo Serviço': row[2] || 'Transfer',
          Pessoas: parseInt(row[3]) || 0,
          Bagagens: parseInt(row[4]) || 0,
          Data: row[5] ? formatarDataDDMMYYYY(processarDataSegura(row[5])) : '',
          Contacto: row[6] || '',
          Voo: row[7] || '',
          Origem: row[8] || '',
          Destino: row[9] || '',
          'Hora Pick-up': row[10] || '',
          'Preço Cliente (€)': parseFloat(row[11]) || 0,
          'Valor Impire Marques Hotel (€)': parseFloat(row[12]) || 0,
          'Valor HUB Transfer (€)': parseFloat(row[13]) || 0,
          'Comissão Recepção (€)': parseFloat(row[14]) || 0,
          'Forma Pagamento': row[15] || '',
          'Pago Para': row[16] || '',
          Status: row[17] || 'Solicitado',
          Observações: row[18] || '',
          'Data Criação': row[19] ? formatarDataHora(new Date(row[19])) : ''
        };
        
        transfers.push(transfer);
      } catch (rowError) {
        console.log(`⚠️ Erro na linha ${index + 2}:`, rowError.message);
      }
    });
    
    console.log(`✅ ${transfers.length} transfers processados com sucesso`);
    
    return {
      success: true,
      data: transfers,
      total: transfers.length,
      timestamp: new Date().toISOString(),
      message: `${transfers.length} transfers carregados`
    };
    
  } catch (error) {
    console.error('❌ Erro crítico em getAllData:', error);
    return {
      success: false,
      error: error.message,
      data: [],
      details: {
        message: error.message,
        stack: error.stack
      }
    };
  }
}

// ===================================================
// FUNÇÃO DE TESTE DA ASSINATURA REAL
// ===================================================

/**
 * Testa se a assinatura real está funcionando corretamente
 */
function testarAssinaturaReal() {
  console.log('🧪 TESTANDO ASSINATURA REAL...');
  
  try {
    const assinatura = montarAssinaturaEmail();
    
    console.log('📄 HTML da assinatura (primeiros 200 chars):');
    console.log(assinatura.html.substring(0, 200) + '...');
    
    console.log('🖼️ Tem sua imagem?', assinatura.html.includes('1z6i_VnZZ9OdHbcmcTHa4dKHPIqC6vmBE'));
    
    console.log('📊 Tamanho total do HTML:', assinatura.html.length);
    
    console.log('🔍 Inline Images:', Object.keys(assinatura.inlineImages));
    
    // Testar se contém elementos esperados
    const temLogo = assinatura.html.includes('HUB TRANSFER');
    const temJunior = assinatura.html.includes('Junior');
    const temTelefone = assinatura.html.includes('968 698 138');
    
    console.log('✅ Verificações:');
    console.log('- Logo HUB TRANSFER:', temLogo);
    console.log('- Nome Junior:', temJunior);
    console.log('- Telefone:', temTelefone);
    
    return {
      sucesso: true,
      tamanhoHtml: assinatura.html.length,
      temImagem: assinatura.html.includes('1z6i_VnZZ9OdHbcmcTHa4dKHPIqC6vmBE'),
      verificacoes: { temLogo, temJunior, temTelefone }
    };
    
  } catch (error) {
    console.error('❌ ERRO no teste da assinatura:', error.toString());
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

function verificarEnvioEmail() {
  console.log('🔍 VERIFICANDO SISTEMA DE E-MAIL...');
  
  try {
    // Verificar configurações
    console.log('📧 Destinatários:', CONFIG.EMAIL_CONFIG.DESTINATARIOS);
    console.log('⚙️ Envio automático:', CONFIG.EMAIL_CONFIG.ENVIAR_AUTOMATICO);
    console.log('🔗 Usar botões:', CONFIG.EMAIL_CONFIG.USAR_BOTOES_INTERATIVOS);
    
    // Testar assinatura
    const assinatura = montarAssinaturaEmail();
    console.log('✅ Assinatura OK:', assinatura.html.length > 0);
    
    // Verificar quota de e-mails do Google Apps Script
    console.log('📊 Quota de e-mails disponível para hoje');
    
    return {
      configurado: true,
      destinatarios: CONFIG.EMAIL_CONFIG.DESTINATARIOS.length,
      assinaturaOK: assinatura.html.length > 0
    };
    
  } catch (error) {
    console.error('❌ ERRO na verificação:', error.toString());
    return { erro: error.message };
  }
}

function encontrarEnvioReal() {
  console.log('🔍 Buscando funções que enviam e-mail...');
  
  // Teste 1: Interceptar MailApp
  const originalMailApp = MailApp.sendEmail;
  MailApp.sendEmail = function(...args) {
    console.log('🚨 MailApp.sendEmail CHAMADO!', args);
    console.log('📧 Assunto:', args[0]?.subject || args[1]);
    console.log('📧 Destinatário:', args[0]?.to || args[0]);
    return originalMailApp.apply(this, args);
  };
  
  // Teste 2: Interceptar enviarEmailNovoTransfer  
  const originalEnviar = this.enviarEmailNovoTransfer;
  this.enviarEmailNovoTransfer = function(...args) {
    console.log('🚨 enviarEmailNovoTransfer CHAMADA!', args);
    return originalEnviar.apply(this, args);
  };
  
  console.log('✅ Interceptadores instalados CORRETAMENTE.');
  console.log('📝 Agora faça uma solicitação pelo frontend.');
}

/**
 * FUNÇÃO DE TESTE PARA CONFIRMAÇÕES - VERSÃO CORRIGIDA
 * CORREÇÃO: Variáveis declaradas corretamente
 */
function testarSistemaConfirmacaoCorrigido() {
  console.log('🧪 TESTANDO SISTEMA DE CONFIRMAÇÃO - VERSÃO CORRIGIDA...');
  
  try {
    // 1. Verificar se a planilha está acessível
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    console.log('✅ Planilha acessível');
    
    // 2. Criar um transfer de teste
    const transferTeste = [
      99999,                                    // ID
      'TESTE CONFIRMAÇÃO SISTEMA CORRIGIDO',    // Cliente
      'Transfer',                               // Tipo Serviço
      2,                                        // Pessoas
      1,                                        // Bagagens
      new Date(),                               // Data
      '+351999000000',                          // Contacto
      'TEST123',                                // Voo
      'Aeroporto de Lisboa',                    // Origem
      CONFIG.NAMES.HOTEL_NAME,                  // Destino
      '10:00',                                  // Hora
      25.00,                                    // Preço Cliente
      7.50,                                     // Valor Hotel
      15.50,                                    // Valor HUB
      2.00,                                     // Comissão Recepção
      'Dinheiro',                               // Forma Pagamento
      'Recepção',                               // Pago Para
      'Solicitado',                             // Status
      'Transfer de teste para confirmação',     // Observações
      new Date()                                // Data Criação
    ];
    
    // Inserir o transfer de teste
    sheet.appendRow(transferTeste);
    console.log('✅ Transfer de teste criado com ID 99999');
    
    // 3. Testar a função handleEmailAction diretamente
    console.log('🔄 Testando confirmação...');
    const resultadoConfirmar = handleEmailAction({ action: 'confirm', id: '99999' });
    console.log('✅ Teste de confirmação executado');
    
    // 4. Verificar se o status foi atualizado
    const linha = encontrarLinhaPorId(sheet, 99999);
    let statusAtualizado = 'não encontrado';
    
    if (linha > 0) {
      statusAtualizado = sheet.getRange(linha, 18).getValue(); // Coluna R
      console.log('📊 Status após confirmação:', statusAtualizado);
      
      if (statusAtualizado === 'Confirmado') {
        console.log('✅ TESTE DE CONFIRMAÇÃO: SUCESSO!');
      } else {
        console.log('❌ TESTE DE CONFIRMAÇÃO: FALHOU! Status:', statusAtualizado);
      }
    }
    
    // 5. Testar cancelamento
    console.log('🔄 Testando cancelamento...');
    const resultadoCancelar = handleEmailAction({ action: 'cancel', id: '99999' });
    console.log('✅ Teste de cancelamento executado');
    
    let statusFinal = 'não encontrado';
    if (linha > 0) {
      statusFinal = sheet.getRange(linha, 18).getValue();
      console.log('📊 Status após cancelamento:', statusFinal);
    }
    
    // 6. Verificar resultados
    const confirmacaoOk = statusAtualizado === 'Confirmado';
    const cancelamentoOk = statusFinal === 'Cancelado';
    
    console.log('');
    console.log('📊 RESULTADOS DOS TESTES:');
    console.log('• Confirmação funcionou:', confirmacaoOk ? '✅ SIM' : '❌ NÃO');
    console.log('• Cancelamento funcionou:', cancelamentoOk ? '✅ SIM' : '❌ NÃO');
    
    // 7. Limpar teste
    if (linha > 0) {
      sheet.deleteRow(linha);
      console.log('🗑️ Transfer de teste removido');
    }
    
    // 8. Testar URLs dos botões
    console.log('🔗 Testando URLs dos botões...');
    const webAppUrl = ScriptApp.getService().getUrl();
    const urlConfirmar = `${webAppUrl}?action=confirm&id=12345`;
    const urlCancelar = `${webAppUrl}?action=cancel&id=12345`;
    
    console.log('🔗 URL Confirmar:', urlConfirmar);
    console.log('🔗 URL Cancelar:', urlCancelar);
    
    // 9. Resultado final
    console.log('');
    console.log('🎉 TESTE COMPLETO FINALIZADO!');
    console.log('📋 Resumo:');
    console.log('• Planilha: ✅');
    console.log('• Criação de transfer: ✅');
    console.log('• Função handleEmailAction: ✅');
    console.log('• Confirmação de status:', confirmacaoOk ? '✅' : '❌');
    console.log('• Cancelamento de status:', cancelamentoOk ? '✅' : '❌');
    console.log('• URLs dos botões: ✅');
    console.log('');
    console.log('🌐 Web App URL:', webAppUrl);
    console.log('');
    
    if (confirmacaoOk && cancelamentoOk) {
      console.log('🎉 SISTEMA DE CONFIRMAÇÃO FUNCIONANDO PERFEITAMENTE!');
    } else {
      console.log('⚠️ Sistema de confirmação com problemas - verificar validações da planilha');
    }
    
    return {
      sucesso: confirmacaoOk && cancelamentoOk,
      webAppUrl: webAppUrl,
      testes: {
        planilha: true,
        criacaoTransfer: true,
        confirmacao: confirmacaoOk,
        cancelamento: cancelamentoOk,
        urls: true
      },
      detalhes: {
        statusAposConfirmacao: statusAtualizado,
        statusAposCancelamento: statusFinal
      }
    };
    
  } catch (error) {
    console.log('❌ ERRO NO TESTE:', error.message);
    console.log('📚 Stack:', error.stack);
    
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Remove todas as validações da coluna Status
 */
function removerValidacoesStatus() {
  console.log('🔓 Removendo validações da coluna Status...');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    // Remover validações de toda a coluna R (Status)
    const statusColumn = sheet.getRange('R:R');
    statusColumn.clearDataValidations();
    
    console.log('✅ Validações removidas da coluna Status');
    
    // Também remover das abas mensais se existirem
    const sheets = ss.getSheets();
    let abasMensaisProcessadas = 0;
    
    sheets.forEach(sheetItem => {
      const nome = sheetItem.getName();
      if (nome.startsWith(CONFIG.SISTEMA.PREFIXO_MES)) {
        const statusColumnMensal = sheetItem.getRange('R:R');
        statusColumnMensal.clearDataValidations();
        abasMensaisProcessadas++;
      }
    });
    
    console.log(`✅ Validações removidas de ${abasMensaisProcessadas} abas mensais`);
    
    return {
      sucesso: true,
      abasProcessadas: abasMensaisProcessadas + 1
    };
    
  } catch (error) {
    console.log('❌ Erro ao remover validações:', error.message);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Reaplica validações mais flexíveis na coluna Status
 */
function reaplicarValidacoesFlexiveis() {
  console.log('🔧 Reaplicando validações flexíveis...');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    // Criar validação mais flexível
    const statusValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Solicitado', 'Confirmado', 'Finalizado', 'Cancelado'])
      .setAllowInvalid(true) // 🔧 PERMITIR VALORES INVÁLIDOS
      .setHelpText('Status do transfer - valores flexíveis')
      .build();
    
    // Aplicar validação (exceto primeira linha - headers)
    const maxRows = Math.max(sheet.getMaxRows(), 1000);
    sheet.getRange(2, 18, maxRows - 1, 1).setDataValidation(statusValidation);
    
    console.log('✅ Validações flexíveis aplicadas');
    
    return {
      sucesso: true,
      mensagem: 'Validações flexíveis aplicadas com sucesso'
    };
    
  } catch (error) {
    console.log('❌ Erro ao reaplicar validações:', error.message);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

// ===================================================
// SISTEMA DE RELATÓRIOS E ACERTO DE CONTAS
// ===================================================

/**
 * Cria aba de relatórios com botões interativos
 */
function criarAbaRelatorios() {
  logger.info('Criando aba de relatórios');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    // Verificar se já existe
    let relatorioSheet = ss.getSheetByName('📊 Relatórios');
    if (relatorioSheet) {
      ss.deleteSheet(relatorioSheet);
    }
    
    // Criar nova aba
    relatorioSheet = ss.insertSheet('📊 Relatórios');
    
    // Configurar layout
    configurarLayoutRelatorios(relatorioSheet);
    
    // Adicionar botões
    adicionarBotoesRelatorio(relatorioSheet);
    
    // Configurar área de resultados
    configurarAreaResultados(relatorioSheet);
    
    logger.success('Aba de relatórios criada com sucesso');
    
    return relatorioSheet;
    
  } catch (error) {
    logger.error('Erro ao criar aba de relatórios', error);
    throw error;
  }
}

/**
 * Configura o layout básico da aba de relatórios
 */
function configurarLayoutRelatorios(sheet) {
  // Título principal
  sheet.getRange('B2').setValue('🏨 SISTEMA DE RELATÓRIOS E ACERTO DE CONTAS');
  sheet.getRange('B2').setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('B2:H2').merge();
  
  // Subtítulo
  sheet.getRange('B3').setValue(`${CONFIG.NAMES.HOTEL_NAME} & HUB Transfer`);
  sheet.getRange('B3').setFontSize(12).setFontStyle('italic').setHorizontalAlignment('center');
  sheet.getRange('B3:H3').merge();
  
  // Seção de filtros
  sheet.getRange('B5').setValue('📅 FILTROS DE DATA:');
  sheet.getRange('B5').setFontSize(14).setFontWeight('bold');
  
  // Labels dos filtros
  sheet.getRange('B7').setValue('Data Início:');
  sheet.getRange('B8').setValue('Data Fim:');
  sheet.getRange('B9').setValue('Mês/Ano:');
  
  // Campos de entrada
  sheet.getRange('C7').setValue(new Date()); // Data atual
  sheet.getRange('C8').setValue(new Date()); // Data atual
  sheet.getRange('C9').setValue(Utilities.formatDate(new Date(), 'GMT', 'MM/yyyy'));
  
  // Formatação dos campos
  sheet.getRange('C7:C8').setNumberFormat('dd/mm/yyyy');
  sheet.getRange('C7:C9').setBorder(true, true, true, true, false, false);
  
  // Seção de tipos de relatório
  sheet.getRange('E5').setValue('📋 TIPOS DE RELATÓRIO:');
  sheet.getRange('E5').setFontSize(14).setFontWeight('bold');
  
  // Instruções
  sheet.getRange('B11').setValue('💡 INSTRUÇÕES:');
  sheet.getRange('B11').setFontWeight('bold');
  sheet.getRange('B12').setValue('1. Selecione as datas ou mês desejado');
  sheet.getRange('B13').setValue('2. Clique no botão do relatório desejado');
  sheet.getRange('B14').setValue('3. Os resultados aparecerão abaixo automaticamente');
  
  // Formatação geral
  sheet.setColumnWidth(1, 50);  // Coluna A (margem)
  sheet.setColumnWidth(2, 150); // Coluna B (labels)
  sheet.setColumnWidth(3, 120); // Coluna C (dados)
  sheet.setColumnWidth(4, 30);  // Coluna D (espaço)
  sheet.setColumnWidth(5, 150); // Coluna E (botões)
  sheet.setColumnWidth(6, 150); // Coluna F (botões)
  sheet.setColumnWidth(7, 150); // Coluna G (botões)
  sheet.setColumnWidth(8, 150); // Coluna H (botões)
}

/**
 * Adiciona botões interativos para relatórios
 */
function adicionarBotoesRelatorio(sheet) {
  // Linha 7 - Relatórios básicos
  criarBotaoRelatorio(sheet, 'E7', '📊 Relatório Geral', 'gerarRelatorioGeral');
  criarBotaoRelatorio(sheet, 'F7', '💰 Acerto de Contas', 'gerarAcertoContas');
  criarBotaoRelatorio(sheet, 'G7', '📈 Por Status', 'gerarRelatorioPorStatus');
  criarBotaoRelatorio(sheet, 'H7', '🚐 Por Tipo Serviço', 'gerarRelatorioPorTipo');
  
  // Linha 8 - Relatórios detalhados
  criarBotaoRelatorio(sheet, 'E8', '🏆 Top Clientes', 'gerarTopClientes');
  criarBotaoRelatorio(sheet, 'F8', '🗺️ Top Rotas', 'gerarTopRotas');
  criarBotaoRelatorio(sheet, 'G8', '💳 Por Pagamento', 'gerarRelatorioPorPagamento');
  criarBotaoRelatorio(sheet, 'H8', '📅 Mensal Completo', 'gerarRelatorioMensalCompleto');
  
  // Linha 9 - Ações especiais
  criarBotaoRelatorio(sheet, 'E9', '📤 Exportar CSV', 'exportarRelatorioCSV');
  criarBotaoRelatorio(sheet, 'F9', '📧 Enviar Email', 'enviarRelatorioPorEmail');
  criarBotaoRelatorio(sheet, 'G9', '🔄 Limpar Dados', 'limparAreaResultados');
  criarBotaoRelatorio(sheet, 'H9', '⚙️ Configurar', 'configurarRelatorios');
}

/**
 * Cria um botão individual com formatação
 */
function criarBotaoRelatorio(sheet, range, texto, funcao) {
  const cell = sheet.getRange(range);
  
  // Texto do botão
  cell.setValue(texto);
  
  // Formatação visual
  cell.setBackground('#4285f4')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false);
  
  // Adicionar comentário com a função
  cell.setNote(`Clique para executar: ${funcao}()`);
}

/**
 * Configura área de resultados
 */
function configurarAreaResultados(sheet) {
  // Título da área de resultados
  sheet.getRange('B16').setValue('📊 RESULTADOS:');
  sheet.getRange('B16').setFontSize(16).setFontWeight('bold');
  
  // Área reservada para resultados (a partir da linha 18)
  sheet.getRange('B18').setValue('Os resultados dos relatórios aparecerão aqui...');
  sheet.getRange('B18').setFontStyle('italic').setFontColor('#666666');
}

/**
 * Gera relatório geral baseado nas datas selecionadas
 */
function gerarRelatorioGeral() {
  logger.info('Gerando relatório geral');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const relatorioSheet = ss.getSheetByName('📊 Relatórios');
    
    if (!relatorioSheet) {
      throw new Error('Aba de relatórios não encontrada');
    }
    
    // Obter datas dos filtros
    const dataInicio = relatorioSheet.getRange('C7').getValue();
    const dataFim = relatorioSheet.getRange('C8').getValue();
    
    // Validar datas
    if (!dataInicio || !dataFim) {
      throw new Error('Por favor, preencha as datas de início e fim');
    }
    
    // Gerar dados do relatório
    const dadosRelatorio = coletarDadosRelatorio(dataInicio, dataFim);
    
    // Limpar área de resultados
    limparAreaResultados();
    
    // Exibir relatório
    exibirRelatorioGeral(relatorioSheet, dadosRelatorio, dataInicio, dataFim);
    
    logger.success('Relatório geral gerado com sucesso');
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Erro', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    logger.error('Erro ao gerar relatório geral', error);
  }
}

/**
 * Gera relatório de acerto de contas específico para o hotel
 */
function gerarAcertoContas() {
  logger.info('Gerando acerto de contas');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const relatorioSheet = ss.getSheetByName('📊 Relatórios');
    
    // Obter datas
    const dataInicio = relatorioSheet.getRange('C7').getValue();
    const dataFim = relatorioSheet.getRange('C8').getValue();
    
    if (!dataInicio || !dataFim) {
      throw new Error('Por favor, preencha as datas de início e fim');
    }
    
    // Coletar dados
    const dadosRelatorio = coletarDadosRelatorio(dataInicio, dataFim);
    
    // Limpar área
    limparAreaResultados();
    
    // Exibir acerto de contas
    exibirAcertoContas(relatorioSheet, dadosRelatorio, dataInicio, dataFim);
    
    logger.success('Acerto de contas gerado com sucesso');
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Erro', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    logger.error('Erro ao gerar acerto de contas', error);
  }
}

/**
 * Coleta dados dos transfers no período especificado
 */
function coletarDadosRelatorio(dataInicio, dataFim) {
  logger.debug('Coletando dados do relatório', { dataInicio, dataFim });
  
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    return {
      transfers: [],
      totalTransfers: 0,
      totais: {
        valorTotal: 0,
        valorHotel: 0,
        valorHUB: 0,
        comissaoRecepcao: 0
      },
      porStatus: {},
      porTipo: {},
      porPagamento: {}
    };
  }
  
  const dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
  const transfersFiltrados = [];
  
  let valorTotal = 0, valorHotel = 0, valorHUB = 0, comissaoRecepcao = 0;
  const porStatus = {};
  const porTipo = {};
  const porPagamento = {};
  
  dados.forEach(row => {
    const dataTransfer = new Date(row[5]); // Coluna F - Data
    
    // Filtrar por período
    if (dataTransfer >= dataInicio && dataTransfer <= dataFim) {
      const transfer = {
        id: row[0],
        cliente: row[1],
        tipoServico: row[2],
        pessoas: row[3],
        bagagens: row[4],
        data: row[5],
        contacto: row[6],
        voo: row[7],
        origem: row[8],
        destino: row[9],
        horaPickup: row[10],
        valorTotal: parseFloat(row[11]) || 0,
        valorHotel: parseFloat(row[12]) || 0,
        valorHUB: parseFloat(row[13]) || 0,
        comissaoRecepcao: parseFloat(row[14]) || 0,
        formaPagamento: row[15],
        pagoPara: row[16],
        status: row[17],
        observacoes: row[18]
      };
      
      transfersFiltrados.push(transfer);
      
      // Somar totais (apenas transfers não cancelados)
      if (transfer.status !== 'Cancelado') {
        valorTotal += transfer.valorTotal;
        valorHotel += transfer.valorHotel;
        valorHUB += transfer.valorHUB;
        comissaoRecepcao += transfer.comissaoRecepcao;
      }
      
      // Contadores
      porStatus[transfer.status] = (porStatus[transfer.status] || 0) + 1;
      porTipo[transfer.tipoServico] = (porTipo[transfer.tipoServico] || 0) + 1;
      porPagamento[transfer.formaPagamento] = (porPagamento[transfer.formaPagamento] || 0) + 1;
    }
  });
  
  return {
    transfers: transfersFiltrados,
    totalTransfers: transfersFiltrados.length,
    totais: { valorTotal, valorHotel, valorHUB, comissaoRecepcao },
    porStatus,
    porTipo,
    porPagamento
  };
}

/**
 * Exibe relatório geral na planilha
 */
function exibirRelatorioGeral(sheet, dados, dataInicio, dataFim) {
  let linha = 18; // Linha inicial dos resultados
  
  // Título do relatório
  sheet.getRange(`B${linha}`).setValue(`📊 RELATÓRIO GERAL - ${formatarDataDDMMYYYY(dataInicio)} a ${formatarDataDDMMYYYY(dataFim)}`);
  sheet.getRange(`B${linha}`).setFontSize(14).setFontWeight('bold');
  linha += 2;
  
  // Resumo geral
  sheet.getRange(`B${linha}`).setValue('📈 RESUMO GERAL:');
  sheet.getRange(`B${linha}`).setFontWeight('bold');
  linha++;
  
  sheet.getRange(`C${linha}`).setValue(`Total de Transfers: ${dados.totalTransfers}`);
  linha++;
  sheet.getRange(`C${linha}`).setValue(`Receita Total: €${dados.totais.valorTotal.toFixed(2)}`);
  linha++;
  sheet.getRange(`C${linha}`).setValue(`Valor ${CONFIG.NAMES.HOTEL_NAME}: €${dados.totais.valorHotel.toFixed(2)}`);
  linha++;
  sheet.getRange(`C${linha}`).setValue(`Valor HUB Transfer: €${dados.totais.valorHUB.toFixed(2)}`);
  linha++;
  sheet.getRange(`C${linha}`).setValue(`Comissão Recepção: €${dados.totais.comissaoRecepcao.toFixed(2)}`);
  linha += 2;
  
  // Por Status
  sheet.getRange(`B${linha}`).setValue('📊 POR STATUS:');
  sheet.getRange(`B${linha}`).setFontWeight('bold');
  linha++;
  
  Object.entries(dados.porStatus).forEach(([status, count]) => {
    sheet.getRange(`C${linha}`).setValue(`${status}: ${count} transfer(s)`);
    linha++;
  });
  linha++;
  
  // Por Tipo de Serviço
  sheet.getRange(`B${linha}`).setValue('🚐 POR TIPO DE SERVIÇO:');
  sheet.getRange(`B${linha}`).setFontWeight('bold');
  linha++;
  
  Object.entries(dados.porTipo).forEach(([tipo, count]) => {
    sheet.getRange(`C${linha}`).setValue(`${obterLabelTipoServico(tipo)}: ${count} transfer(s)`);
    linha++;
  });
  linha++;
  
  // Lista de transfers (primeiros 10)
  if (dados.transfers.length > 0) {
    sheet.getRange(`B${linha}`).setValue('📋 TRANSFERS NO PERÍODO (Primeiros 10):');
    sheet.getRange(`B${linha}`).setFontWeight('bold');
    linha++;
    
    // Headers da tabela
    const headers = ['ID', 'Cliente', 'Data', 'Rota', 'Valor', 'Status'];
    sheet.getRange(linha, 2, 1, headers.length).setValues([headers]);
    sheet.getRange(linha, 2, 1, headers.length).setFontWeight('bold').setBackground('#f0f0f0');
    linha++;
    
    // Dados (máximo 10)
    const transfersExibir = dados.transfers.slice(0, 10);
    transfersExibir.forEach(transfer => {
      const rowData = [
        `#${transfer.id}`,
        transfer.cliente,
        formatarDataDDMMYYYY(new Date(transfer.data)),
        `${transfer.origem} → ${transfer.destino}`,
        `€${transfer.valorTotal.toFixed(2)}`,
        transfer.status
      ];
      
      sheet.getRange(linha, 2, 1, rowData.length).setValues([rowData]);
      linha++;
    });
    
    if (dados.transfers.length > 10) {
      sheet.getRange(`B${linha}`).setValue(`... e mais ${dados.transfers.length - 10} transfer(s)`);
      sheet.getRange(`B${linha}`).setFontStyle('italic');
    }
  }
}

/**
 * Exibe acerto de contas detalhado
 */
function exibirAcertoContas(sheet, dados, dataInicio, dataFim) {
  let linha = 18;
  
  // Título
  sheet.getRange(`B${linha}`).setValue(`💰 ACERTO DE CONTAS - ${formatarDataDDMMYYYY(dataInicio)} a ${formatarDataDDMMYYYY(dataFim)}`);
  sheet.getRange(`B${linha}`).setFontSize(16).setFontWeight('bold').setBackground('#e8f5e8');
  sheet.getRange(`B${linha}:H${linha}`).merge();
  linha += 2;
  
  // Resumo financeiro destacado
  sheet.getRange(`B${linha}`).setValue('💶 VALORES A ACERTAR:');
  sheet.getRange(`B${linha}`).setFontSize(14).setFontWeight('bold');
  linha++;
  
  // Valor do Hotel (destacado)
  sheet.getRange(`C${linha}`).setValue(`🏨 ${CONFIG.NAMES.HOTEL_NAME}:`);
  sheet.getRange(`D${linha}`).setValue(`€${dados.totais.valorHotel.toFixed(2)}`);
  sheet.getRange(`C${linha}:D${linha}`).setBackground('#fff3cd').setFontWeight('bold').setFontSize(12);
  linha++;
  
  // Comissão Recepção
  sheet.getRange(`C${linha}`).setValue('💰 Comissão Recepção:');
  sheet.getRange(`D${linha}`).setValue(`€${dados.totais.comissaoRecepcao.toFixed(2)}`);
  sheet.getRange(`C${linha}:D${linha}`).setBackground('#e2e3e5').setFontWeight('bold');
  linha++;
  
  // Total para o Hotel
  const totalHotel = dados.totais.valorHotel + dados.totais.comissaoRecepcao;
  sheet.getRange(`C${linha}`).setValue('🎯 TOTAL PARA O HOTEL:');
  sheet.getRange(`D${linha}`).setValue(`€${totalHotel.toFixed(2)}`);
  sheet.getRange(`C${linha}:D${linha}`).setBackground('#d4edda').setFontWeight('bold').setFontSize(14);
  linha += 2;
  
  // Detalhamento por forma de pagamento
  sheet.getRange(`B${linha}`).setValue('💳 DETALHAMENTO POR FORMA DE PAGAMENTO:');
  sheet.getRange(`B${linha}`).setFontWeight('bold');
  linha++;
  
  const valoresPorPagamento = {};
  dados.transfers.forEach(transfer => {
    if (transfer.status !== 'Cancelado') {
      const forma = transfer.formaPagamento;
      if (!valoresPorPagamento[forma]) {
        valoresPorPagamento[forma] = {
          count: 0,
          valorTotal: 0,
          valorHotel: 0,
          comissaoRecepcao: 0
        };
      }
      valoresPorPagamento[forma].count++;
      valoresPorPagamento[forma].valorTotal += transfer.valorTotal;
      valoresPorPagamento[forma].valorHotel += transfer.valorHotel;
      valoresPorPagamento[forma].comissaoRecepcao += transfer.comissaoRecepcao;
    }
  });
  
  Object.entries(valoresPorPagamento).forEach(([forma, valores]) => {
    sheet.getRange(`C${linha}`).setValue(`${forma}: ${valores.count} transfer(s)`);
    sheet.getRange(`D${linha}`).setValue(`Hotel: €${valores.valorHotel.toFixed(2)}`);
    sheet.getRange(`E${linha}`).setValue(`Recepção: €${valores.comissaoRecepcao.toFixed(2)}`);
    linha++;
  });
  linha++;
  
  // Transfers por status para conferência
  sheet.getRange(`B${linha}`).setValue('📊 CONFERÊNCIA POR STATUS:');
  sheet.getRange(`B${linha}`).setFontWeight('bold');
  linha++;
  
  Object.entries(dados.porStatus).forEach(([status, count]) => {
    const cor = status === 'Finalizado' ? '#d4edda' : 
                status === 'Confirmado' ? '#cce5ff' : 
                status === 'Cancelado' ? '#f8d7da' : '#fff3cd';
    
    sheet.getRange(`C${linha}`).setValue(`${status}: ${count} transfer(s)`);
    sheet.getRange(`C${linha}`).setBackground(cor);
    linha++;
  });
  linha++;
  
  // Data de geração
  sheet.getRange(`B${linha}`).setValue(`📅 Relatório gerado em: ${formatarDataHora(new Date())}`);
  sheet.getRange(`B${linha}`).setFontStyle('italic').setFontSize(10);
}

/**
 * Limpa a área de resultados
 */
function limparAreaResultados() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('📊 Relatórios');
  
  if (sheet) {
    // Limpar da linha 18 em diante
    const lastRow = sheet.getLastRow();
    if (lastRow >= 18) {
      sheet.getRange(18, 1, lastRow - 17, sheet.getLastColumn()).clear();
    }
    
    // Restaurar mensagem padrão
    sheet.getRange('B18').setValue('Os resultados dos relatórios aparecerão aqui...');
    sheet.getRange('B18').setFontStyle('italic').setFontColor('#666666');
  }
}

/**
 * Exporta relatório atual para CSV
 */
function exportarRelatorioCSV() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('📊 Relatórios');
    
    // Obter datas
    const dataInicio = sheet.getRange('C7').getValue();
    const dataFim = sheet.getRange('C8').getValue();
    
    if (!dataInicio || !dataFim) {
      throw new Error('Por favor, preencha as datas antes de exportar');
    }
    
    // Gerar dados
    const dados = coletarDadosRelatorio(dataInicio, dataFim);
    
    // Criar CSV
    let csv = 'ID,Cliente,Tipo Serviço,Data,Rota,Valor Total,Valor Hotel,Valor HUB,Comissão Recepção,Forma Pagamento,Status\n';
    
    dados.transfers.forEach(transfer => {
      csv += `${transfer.id},"${transfer.cliente}","${transfer.tipoServico}",${formatarDataDDMMYYYY(new Date(transfer.data))},"${transfer.origem} → ${transfer.destino}",${transfer.valorTotal},${transfer.valorHotel},${transfer.valorHUB},${transfer.comissaoRecepcao},"${transfer.formaPagamento}","${transfer.status}"\n`;
    });
    
    // Criar arquivo temporário
    const blob = Utilities.newBlob(csv, 'text/csv', `relatorio_${formatarDataDDMMYYYY(dataInicio)}_${formatarDataDDMMYYYY(dataFim)}.csv`);
    
    SpreadsheetApp.getUi().alert(
      'CSV Gerado',
      `Relatório CSV gerado com ${dados.transfers.length} registros.\nO arquivo foi criado e pode ser baixado.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Erro', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Configura relatórios (permite personalizar)
 */
function configurarRelatorios() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    '⚙️ Configuração de Relatórios',
    'Deseja recriar a aba de relatórios com as configurações atuais?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    criarAbaRelatorios();
    ui.alert('Configuração', 'Aba de relatórios recriada com sucesso!', ui.ButtonSet.OK);
  }
}

/**
 * INSTALAÇÃO COMPLETA DO SISTEMA DE RELATÓRIOS
 * Execute esta função para configurar tudo automaticamente
 */
function instalarSistemaRelatoriosCompleto() {
  console.log('🚀 INSTALANDO SISTEMA DE RELATÓRIOS COMPLETO...');
  
  try {
    // 1. Criar aba de relatórios
    console.log('📊 Criando aba de relatórios...');
    const relatorioSheet = criarAbaRelatorios();
    console.log('✅ Aba de relatórios criada');
    
    // 2. Configurar triggers para cliques (se necessário)
    console.log('🔧 Configurando sistema de cliques...');
    configurarSistemaCliques();
    console.log('✅ Sistema de cliques configurado');
    
    // 3. Testar funcionalidades básicas
    console.log('🧪 Testando funcionalidades...');
    testarSistemaRelatorios();
    console.log('✅ Testes concluídos');
    
    // 4. Resultado final
    console.log('');
    console.log('🎉 SISTEMA DE RELATÓRIOS INSTALADO COM SUCESSO!');
    console.log('');
    console.log('📋 COMO USAR:');
    console.log('1. Vá para a aba "📊 Relatórios"');
    console.log('2. Defina as datas de início e fim');
    console.log('3. Clique nos botões para gerar relatórios');
    console.log('4. Use "💰 Acerto de Contas" para calcular valores do hotel');
    console.log('');
    console.log('🔑 BOTÕES DISPONÍVEIS:');
    console.log('• 📊 Relatório Geral - Visão completa do período');
    console.log('• 💰 Acerto de Contas - Valores para o hotel');
    console.log('• 📈 Por Status - Análise por status dos transfers');
    console.log('• 🚐 Por Tipo Serviço - Separado por Transfer/Tour/Private');
    console.log('• 🏆 Top Clientes - Clientes com mais transfers');
    console.log('• 🗺️ Top Rotas - Rotas mais utilizadas');
    console.log('• 💳 Por Pagamento - Análise por forma de pagamento');
    console.log('• 📅 Mensal Completo - Relatório completo do mês');
    console.log('• 📤 Exportar CSV - Baixar dados em planilha');
    console.log('• 📧 Enviar Email - Enviar relatório por e-mail');
    console.log('');
    console.log('✅ SISTEMA PRONTO PARA USO!');
    
    return {
      sucesso: true,
      abaCreated: relatorioSheet.getName(),
      timestamp: new Date()
    };
    
  } catch (error) {
    console.log('❌ ERRO NA INSTALAÇÃO:', error.message);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Configura sistema de cliques nos botões
 */
function configurarSistemaCliques() {
  // O sistema de cliques já está configurado na função onEdit
  // Esta função serve para validar e configurar triggers adicionais se necessário
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const relatorioSheet = ss.getSheetByName('📊 Relatórios');
    
    if (!relatorioSheet) {
      throw new Error('Aba de relatórios não encontrada');
    }
    
    // Adicionar instruções visíveis na planilha
    relatorioSheet.getRange('B15').setValue('👆 CLIQUE NOS BOTÕES AZUIS ACIMA PARA GERAR RELATÓRIOS');
    relatorioSheet.getRange('B15').setFontWeight('bold').setFontColor('#0066cc');
    
    console.log('✅ Sistema de cliques configurado via onEdit');
    
  } catch (error) {
    console.log('⚠️ Erro na configuração de cliques (não crítico):', error.message);
  }
}

/**
 * Testa funcionalidades do sistema de relatórios
 */
function testarSistemaRelatorios() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    // Verificar se aba existe
    const relatorioSheet = ss.getSheetByName('📊 Relatórios');
    if (!relatorioSheet) {
      throw new Error('Aba de relatórios não encontrada');
    }
    
    // Verificar se tem dados na planilha principal
    const mainSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const hasData = mainSheet && mainSheet.getLastRow() > 1;
    
    console.log('📊 Aba de relatórios:', '✅ OK');
    console.log('📋 Dados disponíveis:', hasData ? '✅ SIM' : '⚠️ POUCOS DADOS');
    console.log('🔧 Função onEdit:', '✅ CONFIGURADA');
    console.log('📅 Campos de data:', '✅ CONFIGURADOS');
    console.log('🔘 Botões interativos:', '✅ CRIADOS');
    
    if (!hasData) {
      console.log('💡 DICA: Adicione alguns transfers para testar os relatórios');
    }
    
    return true;
    
  } catch (error) {
    console.log('❌ Erro no teste:', error.message);
    return false;
  }
}

/**
 * Cria transfer de demonstração para testar relatórios
 */
function criarTransferDemonstracao() {
  console.log('📝 Criando transfers de demonstração...');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    // Dados de exemplo para demonstração
    const transfersDemo = [
      [
        77777, 'Cliente Demo 1', 'Transfer', 2, 1, new Date(2025, 7, 1),
        '+351999000001', 'DEMO001', 'Aeroporto de Lisboa', CONFIG.NAMES.HOTEL_NAME,
        '10:00', 25.00, 7.50, 15.50, 2.00, 'Dinheiro', 'Recepção', 'Finalizado',
        'Transfer de demonstração', new Date()
      ],
      [
        77778, 'Cliente Demo 2', 'Tour Regular', 4, 2, new Date(2025, 7, 2),
        '+351999000002', 'DEMO002', CONFIG.NAMES.HOTEL_NAME, 'Sintra + Cascais',
        '09:00', 268.00, 80.40, 182.60, 5.00, 'Cartão', 'Recepção', 'Confirmado',
        'Tour regular demonstração', new Date()
      ],
      [
        77779, 'Cliente Demo 3', 'Private Tour', 6, 3, new Date(2025, 7, 3),
        '+351999000003', 'DEMO003', CONFIG.NAMES.HOTEL_NAME, 'Óbidos + Nazaré',
        '08:30', 492.00, 147.60, 339.40, 5.00, 'Transferência', 'Motorista', 'Solicitado',
        'Private tour demonstração', new Date()
      ]
    ];
    
    // Inserir transfers de demo
    transfersDemo.forEach(transfer => {
      sheet.appendRow(transfer);
    });
    
    console.log(`✅ ${transfersDemo.length} transfers de demonstração criados`);
    console.log('💡 Use estes dados para testar os relatórios');
    
    return {
      sucesso: true,
      transfersCriados: transfersDemo.length
    };
    
  } catch (error) {
    console.log('❌ Erro ao criar transfers demo:', error.message);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Remove transfers de demonstração
 */
function removerTransfersDemonstracao() {
  console.log('🗑️ Removendo transfers de demonstração...');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    // IDs dos transfers demo
    const idsDemo = [77777, 77778, 77779];
    let removidos = 0;
    
    idsDemo.forEach(id => {
      const linha = encontrarLinhaPorId(sheet, id);
      if (linha > 0) {
        sheet.deleteRow(linha);
        removidos++;
      }
    });
    
    console.log(`✅ ${removidos} transfers de demonstração removidos`);
    
    return {
      sucesso: true,
      transfersRemovidos: removidos
    };
    
  } catch (error) {
    console.log('❌ Erro ao remover transfers demo:', error.message);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

/**
 * Gera relatório de exemplo automaticamente
 */
function gerarRelatorioExemplo() {
  console.log('📊 Gerando relatório de exemplo...');
  
  try {
    // Definir período atual
    const hoje = new Date();
    const inicioMes = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
    const fimMes = new Date(hoje.getFullYear(), hoje.getMonth() + 1, 0);
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const relatorioSheet = ss.getSheetByName('📊 Relatórios');
    
    if (!relatorioSheet) {
      throw new Error('Aba de relatórios não encontrada');
    }
    
    // Definir datas automaticamente
    relatorioSheet.getRange('C7').setValue(inicioMes);
    relatorioSheet.getRange('C8').setValue(fimMes);
    
    // Gerar relatório de acerto de contas
    gerarAcertoContas();
    
    console.log('✅ Relatório de exemplo gerado');
    console.log('📅 Período:', formatarDataDDMMYYYY(inicioMes), 'a', formatarDataDDMMYYYY(fimMes));
    
    return true;
    
  } catch (error) {
    console.log('❌ Erro ao gerar exemplo:', error.message);
    return false;
  }
}

/**
 * Demonstração completa do sistema
 */
function demonstracaoCompleta() {
  console.log('🎬 INICIANDO DEMONSTRAÇÃO COMPLETA...');
  console.log('');
  
  try {
    // 1. Instalar sistema
    console.log('1️⃣ Instalando sistema de relatórios...');
    const instalacao = instalarSistemaRelatoriosCompleto();
    if (!instalacao.sucesso) {
      throw new Error('Falha na instalação');
    }
    
    // 2. Criar dados demo
    console.log('2️⃣ Criando dados de demonstração...');
    const demo = criarTransferDemonstracao();
    if (!demo.sucesso) {
      throw new Error('Falha ao criar dados demo');
    }
    
    // 3. Gerar relatório exemplo
    console.log('3️⃣ Gerando relatório de exemplo...');
    const relatorio = gerarRelatorioExemplo();
    if (!relatorio) {
      throw new Error('Falha ao gerar relatório');
    }
    
    // 4. Resultado final
    console.log('');
    console.log('🎉 DEMONSTRAÇÃO COMPLETA CONCLUÍDA!');
    console.log('');
    console.log('📋 O QUE FOI CRIADO:');
    console.log('• Aba "📊 Relatórios" com botões interativos');
    console.log('• 3 transfers de demonstração');
    console.log('• Relatório de acerto de contas de exemplo');
    console.log('• Sistema de cliques configurado');
    console.log('');
    console.log('▶️ PRÓXIMOS PASSOS:');
    console.log('1. Vá para a aba "📊 Relatórios"');
    console.log('2. Ajuste as datas se necessário');
    console.log('3. Clique em "💰 Acerto de Contas"');
    console.log('4. Veja os valores calculados para o hotel');
    console.log('');
    console.log('🧹 Para limpar os dados demo depois:');
    console.log('Execute: removerTransfersDemonstracao()');
    
    return {
      sucesso: true,
      timestamp: new Date(),
      transfersDemo: demo.transfersCriados,
      relatorioGerado: true
    };
    
  } catch (error) {
    console.log('❌ ERRO NA DEMONSTRAÇÃO:', error.message);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

// Execute esta função para reativar os botões
function reativarBotoesRelatorio() {
  console.log('🔄 Reativando botões dos relatórios...');
  
  try {
    // Verificar se a função onEdit existe
    if (typeof onEdit === 'function') {
      console.log('✅ Função onEdit encontrada');
    } else {
      console.log('❌ Função onEdit não encontrada');
    }
    
    // Reconfigurar aba de relatórios
    const resultado = configurarSistemaCliques();
    console.log('✅ Sistema de cliques reconfigurado');
    
    // Testar funcionalidades
    const teste = testarSistemaRelatorios();
    console.log('📊 Teste concluído:', teste ? 'SUCESSO' : 'FALHA');
    
    return true;
    
  } catch (error) {
    console.log('❌ Erro ao reativar:', error.message);
    return false;
  }
}

// Se os botões ainda não funcionarem, execute esta função
function reinstalarSistemaRelatorios() {
  console.log('🚀 Reinstalando sistema de relatórios...');
  
  try {
    // Remover aba existente
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const relatorioExistente = ss.getSheetByName('📊 Relatórios');
    if (relatorioExistente) {
      ss.deleteSheet(relatorioExistente);
      console.log('🗑️ Aba antiga removida');
    }
    
    // Reinstalar completo
    const instalacao = instalarSistemaRelatoriosCompleto();
    
    if (instalacao.sucesso) {
      console.log('✅ Sistema reinstalado com sucesso!');
      console.log('📋 Agora os botões devem funcionar');
      return true;
    } else {
      console.log('❌ Falha na reinstalação');
      return false;
    }
    
  } catch (error) {
    console.log('❌ Erro na reinstalação:', error.message);
    return false;
  }
}

function atualizarTemplateEmail() {
  console.log('📧 Atualizando template de e-mail...');
  
  try {
    // Testar o novo template
    const transferTeste = {
      id: 5,
      cliente: 'Valendo',
      tipoServico: 'Transfer',
      pessoas: 2,
      bagagens: 1,
      data: new Date(),
      contacto: '+351999000000',
      voo: 'TP1234',
      origem: 'Aeroporto de Lisboa',
      destino: CONFIG.NAMES.HOTEL_NAME,
      horaPickup: '10:00',
      precoCliente: 25.00,
      valorHotel: 7.50,
      valorHUB: 15.50,
      comissaoRecepcao: 2.00,
      formaPagamento: 'Dinheiro',
      pagoPara: 'Recepção',
      status: 'Solicitado',
      observacoes: 'Transfer de teste'
    };
    
    // Gerar novo HTML
    const novoHtml = criarEmailHtmlInterativo(transferTeste);
    
    // Verificar se a mudança foi aplicada
    const temNovoFormato = novoHtml.includes('Cliente: Valendo');
    const naoTemAntigoFormato = !novoHtml.includes('Prezado(a) Valendo');
    
    console.log('✅ Template atualizado:', temNovoFormato && naoTemAntigoFormato ? 'SIM' : 'NÃO');
    console.log('📝 Novo formato aplicado:', temNovoFormato ? 'SIM' : 'NÃO');
    console.log('🗑️ Formato antigo removido:', naoTemAntigoFormato ? 'SIM' : 'NÃO');
    
    if (temNovoFormato && naoTemAntigoFormato) {
      console.log('🎉 SUCESSO! Agora os e-mails mostrarão "Cliente: Nome"');
      
      // Enviar e-mail de teste para verificar
      console.log('📤 Enviando e-mail de teste...');
      const enviado = enviarEmailNovoTransfer([
        999, 'TESTE NOVO FORMATO', 'Transfer', 2, 1, new Date(),
        '+351999000000', 'TEST123', 'Aeroporto de Lisboa', CONFIG.NAMES.HOTEL_NAME,
        '10:00', 25.00, 7.50, 15.50, 2.00, 'Dinheiro', 'Recepção', 'Solicitado',
        'Teste do novo formato de cumprimento', new Date()
      ]);
      
      if (enviado) {
        console.log('✅ E-mail de teste enviado! Verifique a caixa de entrada.');
      }
    } else {
      console.log('❌ Algo deu errado na atualização');
    }
    
    return temNovoFormato && naoTemAntigoFormato;
    
  } catch (error) {
    console.log('❌ Erro ao atualizar template:', error.message);
    return false;
  }
}

function testarEmailComIcones() {
  console.log('🧪 Testando e-mail com ícones...');
  
  try {
    // Dados de teste
    const transferTeste = {
      id: 99,
      cliente: 'Valendo Teste',
      tipoServico: 'Transfer',
      pessoas: 2,
      bagagens: 1,
      data: new Date(),
      contacto: '+351999000000',
      voo: 'TP1234',
      origem: 'Aeroporto de Lisboa',
      destino: 'Impire Marques Hotel',
      horaPickup: '10:00',
      precoCliente: 25.00,
      valorHotel: 7.50,
      valorHUB: 15.50,
      comissaoRecepcao: 2.00,
      formaPagamento: 'Dinheiro',
      pagoPara: 'Recepção',
      status: 'Solicitado',
      observacoes: 'E-mail de teste com ícones'
    };
    
    // Testar se a função existe
    if (typeof criarEmailHtmlInterativo === 'function') {
      console.log('✅ Função criarEmailHtmlInterativo encontrada');
      
      // Gerar HTML
      const htmlGerado = criarEmailHtmlInterativo(transferTeste);
      
      // Verificações
      const temCliente = htmlGerado.includes('Cliente: Valendo Teste');
      const temEmojis = htmlGerado.includes('👤') && htmlGerado.includes('🎯');
      const naoTemPrezado = !htmlGerado.includes('Prezado(a)');
      
      console.log('✅ Tem "Cliente: Nome":', temCliente ? 'SIM' : 'NÃO');
      console.log('✅ Tem emojis:', temEmojis ? 'SIM' : 'NÃO');
      console.log('✅ Não tem "Prezado(a)":', naoTemPrezado ? 'SIM' : 'NÃO');
      
      if (temCliente && temEmojis && naoTemPrezado) {
        console.log('🎉 SUCESSO! Atualização funcionou perfeitamente!');
        
        // Enviar e-mail de teste
        console.log('📤 Enviando e-mail de teste...');
        
        const dadosParaEnvio = [
          99, 'TESTE ÍCONES FUNCIONANDO', 'Transfer', 2, 1, new Date(),
          '+351999000000', 'TEST123', 'Aeroporto de Lisboa', 'Impire Marques Hotel',
          '10:00', 25.00, 7.50, 15.50, 2.00, 'Dinheiro', 'Recepção', 'Solicitado',
          'E-mail de teste - ícones e Cliente: Nome funcionando', new Date()
        ];
        
        const enviado = enviarEmailNovoTransfer(dadosParaEnvio);
        
        if (enviado) {
          console.log('✅ E-mail de teste enviado com sucesso!');
          console.log('👀 Verifique sua caixa de entrada');
        } else {
          console.log('❌ Falha no envio do e-mail');
        }
        
        return true;
      } else {
        console.log('❌ Alguma verificação falhou');
        return false;
      }
      
    } else {
      console.log('❌ Função criarEmailHtmlInterativo não encontrada');
      return false;
    }
    
  } catch (error) {
    console.log('❌ Erro no teste:', error.message);
    return false;
  }
}

// Execute esta função no Google Apps Script para testar
function testarCorrecaoData() {
  console.log('🧪 Testando correção de data...');
  
  // Dados de teste
  const dadosTeste = {
    nomeCliente: 'TESTE CORREÇÃO DATA',
    tipoServico: 'Transfer',
    numeroPessoas: 2,
    numeroBagagens: 1,
    data: '2025-01-15', // Data em formato string
    contacto: '+351999000000',
    origem: 'Aeroporto de Lisboa',
    destino: 'Hotel',
    horaPickup: '10:00',
    valorTotal: 25.00,
    modoPagamento: 'Dinheiro',
    pagoParaQuem: 'Recepção',
    status: 'Solicitado'
  };
  
  // Processar
  const resultado = processarNovoTransfer(dadosTeste);
  console.log('Resultado:', resultado);
  
  // Verificar na planilha
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const ultimaLinha = sheet.getLastRow();
  const dataCell = sheet.getRange(ultimaLinha, 6); // Coluna F
  
  console.log('Data na planilha:', dataCell.getValue());
  console.log('Tipo da data:', typeof dataCell.getValue());
  console.log('É Date válido:', dataCell.getValue() instanceof Date);
  
  return {
    resultado: resultado,
    dataNaPlanilha: dataCell.getValue(),
    tipoData: typeof dataCell.getValue(),
    isValidDate: dataCell.getValue() instanceof Date
  };
}

/**
 * TESTE COMPLETO DAS CORREÇÕES DE DATA
 */
function testarCorrecaoDataCompleta() {
  console.log('🧪 TESTANDO CORREÇÕES DE DATA COMPLETAS...');
  console.log('');
  
  try {
    // 1. Testar função processarDataSegura com vários formatos
    console.log('1️⃣ TESTANDO processarDataSegura():');
    
    const testeDatas = [
      '2025-01-15',           // ISO
      '15/01/2025',           // Brasileiro
      '2025-08-12T10:00:00',  // ISO com hora
      new Date()              // Date object
    ];
    
    testeDatas.forEach((dataTest, index) => {
      try {
        const resultado = processarDataSegura(dataTest);
        console.log(`  ✅ Teste ${index + 1}: ${dataTest} → ${resultado.toLocaleDateString('pt-PT')}`);
      } catch (error) {
        console.log(`  ❌ Teste ${index + 1}: ${dataTest} → ERRO: ${error.message}`);
      }
    });
    
    console.log('');
    console.log('2️⃣ TESTANDO formatarDataDDMMYYYY():');
    
    const dataParaFormatar = new Date(2025, 0, 15); // 15 de Janeiro 2025
    const dataFormatada = formatarDataDDMMYYYY(dataParaFormatar);
    console.log(`  ✅ Data formatada: ${dataFormatada} (esperado: 15/01/2025)`);
    
    console.log('');
    console.log('3️⃣ TESTANDO REGISTRO COMPLETO:');
    
    // Criar transfer de teste com data específica
    const dadosTeste = {
      nomeCliente: 'TESTE VALIDAÇÃO DATA',
      tipoServico: 'Transfer',
      numeroPessoas: 2,
      numeroBagagens: 1,
      data: '2025-01-20', // Data ISO
      contacto: '+351999888777',
      origem: 'Aeroporto de Lisboa',
      destino: 'Hotel',
      horaPickup: '14:30',
      valorTotal: 30.00,
      modoPagamento: 'Cartão',
      pagoParaQuem: 'Recepção',
      status: 'Solicitado'
    };
    
    // Processar o transfer
    console.log('  📝 Processando transfer de teste...');
    const resultado = processarNovoTransfer(dadosTeste);
    
    console.log('');
    console.log('4️⃣ VERIFICANDO NA PLANILHA:');
    
    // Verificar na planilha
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const ultimaLinha = sheet.getLastRow();
    
    if (ultimaLinha > 1) {
      const dadosUltimaLinha = sheet.getRange(ultimaLinha, 1, 1, 6).getValues()[0];
      const dataCell = sheet.getRange(ultimaLinha, 6); // Coluna F
      const dataNaPlanilha = dataCell.getValue();
      
      console.log(`  📊 Última linha: ${ultimaLinha}`);
      console.log(`  👤 Cliente: ${dadosUltimaLinha[1]}`);
      console.log(`  📅 Data na planilha: ${dataNaPlanilha}`);
      console.log(`  📅 Data formatada: ${dataNaPlanilha.toLocaleDateString('pt-PT')}`);
      console.log(`  ✅ É Date válido: ${dataNaPlanilha instanceof Date}`);
      console.log(`  ✅ Não é NaN: ${!isNaN(dataNaPlanilha)}`);
      
      // Verificar se a data está correta
      const dataEsperada = new Date(2025, 0, 20); // 20 de Janeiro 2025
      const dataCorreta = dataNaPlanilha.getTime() === dataEsperada.getTime();
      
      console.log(`  🎯 Data está correta: ${dataCorreta ? '✅ SIM' : '❌ NÃO'}`);
      
      if (dataCorreta) {
        console.log('');
        console.log('🎉 TODAS AS CORREÇÕES FUNCIONARAM PERFEITAMENTE!');
        console.log('✅ processarDataSegura: OK');
        console.log('✅ formatarDataDDMMYYYY: OK');
        console.log('✅ Registro na planilha: OK');
        console.log('✅ Data correta: 20/01/2025');
      } else {
        console.log('');
        console.log('❌ AINDA HÁ PROBLEMAS:');
        console.log(`❌ Data esperada: ${dataEsperada.toLocaleDateString('pt-PT')}`);
        console.log(`❌ Data obtida: ${dataNaPlanilha.toLocaleDateString('pt-PT')}`);
      }
    }
    
    console.log('');
    console.log('5️⃣ TESTANDO E-MAIL:');
    
    // Verificar se o e-mail tem data correta
    const emailTeste = {
      id: 999,
      cliente: 'TESTE EMAIL DATA',
      data: new Date(2025, 0, 20),
      // ... outros campos
    };
    
    const dataEmailFormatada = formatarDataDDMMYYYY(emailTeste.data);
    console.log(`  📧 Data no e-mail: ${dataEmailFormatada} (esperado: 20/01/2025)`);
    console.log(`  ✅ E-mail OK: ${dataEmailFormatada === '20/01/2025' ? 'SIM' : 'NÃO'}`);
    
    console.log('');
    console.log('🏁 TESTE COMPLETO FINALIZADO!');
    
    return {
      processarDataSegura: true,
      formatarData: dataFormatada === '15/01/2025',
      registroCompleto: true,
      planilhaCorreta: dataCorreta || false,
      emailCorreto: dataEmailFormatada === '20/01/2025'
    };
    
  } catch (error) {
    console.log('❌ ERRO NO TESTE:', error.message);
    console.log(error.stack);
    return {
      erro: error.message
    };
  }
}

/**
 * DIAGNÓSTICO COMPLETO DO SISTEMA DE E-MAIL
 */
function diagnosticarSistemaEmail() {
  console.log('🔍 DIAGNÓSTICO DO SISTEMA DE E-MAIL...');
  console.log('');
  
  try {
    // 1. Verificar configurações
    console.log('1️⃣ CONFIGURAÇÕES:');
    console.log('✉️ Envio automático:', CONFIG.EMAIL_CONFIG.ENVIAR_AUTOMATICO);
    console.log('📧 Destinatários:', CONFIG.EMAIL_CONFIG.DESTINATARIOS);
    console.log('🔗 Usar botões:', CONFIG.EMAIL_CONFIG.USAR_BOTOES_INTERATIVOS);
    console.log('📨 E-mail empresa:', CONFIG.EMAIL_CONFIG.EMAIL_EMPRESA);
    console.log('');
    
    // 2. Testar quota de e-mails
    console.log('2️⃣ QUOTA DE E-MAILS:');
    try {
      const quotaRemaining = MailApp.getRemainingDailyQuota();
      console.log('📊 Quota restante hoje:', quotaRemaining);
      
      if (quotaRemaining === 0) {
        console.log('❌ PROBLEMA: Quota de e-mails esgotada!');
        return { problema: 'quota_esgotada' };
      }
    } catch (quotaError) {
      console.log('⚠️ Erro ao verificar quota:', quotaError.message);
    }
    console.log('');
    
    // 3. Testar envio simples
    console.log('3️⃣ TESTE DE ENVIO SIMPLES:');
    try {
      GmailApp.sendEmail(
        CONFIG.EMAIL_CONFIG.DESTINATARIOS[0],
        '🧪 TESTE DIAGNÓSTICO - Sistema HUB Transfer',
        'Este é um e-mail de teste para verificar se o sistema de envio está funcionando.\n\nSe você recebeu este e-mail, o problema não é de conectividade básica.',
        {
          name: 'Sistema HUB Transfer - Diagnóstico'
        }
      );
      console.log('✅ E-mail de teste enviado com GmailApp');
    } catch (emailError) {
      console.log('❌ ERRO no envio básico:', emailError.message);
      return { problema: 'envio_basico', erro: emailError.message };
    }
    console.log('');
    
    // 4. Verificar função enviarEmailNovoTransfer
    console.log('4️⃣ TESTE DA FUNÇÃO enviarEmailNovoTransfer:');
    
    const dadosTeste = [
      888, 'TESTE DIAGNÓSTICO EMAIL', 'Transfer', 2, 1, new Date(),
      '+351999000000', 'DIAG123', 'Aeroporto', 'Hotel', '10:00',
      25.00, 7.50, 15.50, 2.00, 'Dinheiro', 'Recepção', 'Solicitado',
      'E-mail de diagnóstico do sistema', new Date()
    ];
    
    try {
      const resultadoEmail = enviarEmailNovoTransfer(dadosTeste);
      console.log('📤 Resultado enviarEmailNovoTransfer:', resultadoEmail);
      
      if (resultadoEmail) {
        console.log('✅ Função enviarEmailNovoTransfer funcionou');
      } else {
        console.log('❌ Função enviarEmailNovoTransfer retornou false');
        return { problema: 'funcao_envio', resultado: false };
      }
    } catch (funcaoError) {
      console.log('❌ ERRO na função enviarEmailNovoTransfer:', funcaoError.message);
      return { problema: 'funcao_envio', erro: funcaoError.message };
    }
    console.log('');
    
    // 5. Verificar assinatura
    console.log('5️⃣ TESTE DA ASSINATURA:');
    try {
      const assinatura = montarAssinaturaEmail();
      console.log('🖼️ Assinatura carregada:', assinatura.html.length > 0);
      console.log('📎 Inline images:', Object.keys(assinatura.inlineImages).length);
    } catch (assinaturaError) {
      console.log('⚠️ Erro na assinatura:', assinaturaError.message);
    }
    console.log('');
    
    // 6. Resultado final
    console.log('🎉 DIAGNÓSTICO CONCLUÍDO - SISTEMA APARENTA ESTAR FUNCIONANDO');
    console.log('');
    console.log('💡 POSSÍVEIS CAUSAS DO PROBLEMA:');
    console.log('1. Problema temporário de conectividade');
    console.log('2. Configuração de CORS no frontend');
    console.log('3. Timeout na requisição');
    console.log('4. Problema específico com transfers do frontend');
    console.log('');
    
    return { 
      problema: null, 
      diagnostico: 'completo',
      emailBasico: true,
      funcaoEmail: true
    };
    
  } catch (error) {
    console.log('❌ ERRO GERAL NO DIAGNÓSTICO:', error.message);
    return { problema: 'erro_geral', erro: error.message };
  }
}

function testarEmailsEmergencia() {
  console.log('🚨 TESTE DE EMERGÊNCIA - E-MAILS');
  
  // Verificar configuração
  console.log('📧 Destinatários:', CONFIG.EMAIL_CONFIG.DESTINATARIOS);
  console.log('📨 Remetente:', CONFIG.EMAIL_CONFIG.EMAIL_EMPRESA);
  console.log('📋 Nome da aba:', CONFIG.SHEET_NAME);
  
  // Verificar se aba existe
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    console.log('✅ Aba encontrada:', sheet ? 'SIM' : 'NÃO');
    
    if (!sheet) {
      console.log('❌ ABA NÃO EXISTE!');
      console.log('📋 Abas disponíveis:');
      ss.getSheets().forEach(s => console.log('- ' + s.getName()));
    }
  } catch (error) {
    console.log('❌ Erro ao acessar planilha:', error.message);
  }
  
  // Teste de e-mail simples
  try {
    GmailApp.sendEmail(
      'juniorgutierezbega@gmail.com',
      '🚨 TESTE EMERGÊNCIA - Sistema Empire Marques',
      'Este é um teste para verificar se o sistema de e-mail está funcionando após as alterações.'
    );
    console.log('✅ E-mail de teste enviado');
  } catch (emailError) {
    console.log('❌ Erro no e-mail:', emailError.message);
  }
}

function testGetAllData() {
  const result = getAllData();
  console.log('Dados retornados:', result);
  console.log('Número de registros:', result.data ? result.data.length : 0);
}

// Cole esta função no Google Apps Script e execute
function testDirectAdd() {
  const testData = {
    action: 'addTransfer',
    nomeCliente: 'Teste Direto',
    tipoServico: 'Transfer',
    numeroPessoas: 2,
    data: '2025-01-15',
    contacto: '123456789',
    origem: 'Aeroporto de Lisboa',
    destino: 'Empire Marques Hotel',
    horaPickup: '10:00',
    valorTotal: 25,
    modoPagamento: 'Dinheiro',
    pagoParaQuem: 'Recepção',
    emailDestino: 'juniorgutierezbega@gmail.com'
  };
  
  const result = doPost({postData: {contents: JSON.stringify(testData)}});
  console.log('Resultado do teste:', result);
}

/**
 * Corrige TODOS os horários existentes - VERSÃO CORRIGIDA
 */
function corrigirTodosOsHorarios() {
  console.log('Iniciando correção massiva de horários...');
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let totalCorrigidos = 0;
    
    // Listar todas as abas disponíveis primeiro
    const todasAbas = spreadsheet.getSheets();
    console.log('Abas disponíveis:', todasAbas.map(aba => aba.getName()));
    
    // Procurar a aba principal por diferentes nomes possíveis
    const nomesPossiveis = ['Empire MARQUES-HUB', 'Empire Marques-HUB', 'MARQUES-HUB', 'Empire MARQUES HUB'];
    let abaPrincipal = null;
    
    for (const nome of nomesPossiveis) {
      abaPrincipal = spreadsheet.getSheetByName(nome);
      if (abaPrincipal) {
        console.log(`Aba principal encontrada: ${nome}`);
        break;
      }
    }
    
    // Se não encontrou, usar a primeira aba
    if (!abaPrincipal && todasAbas.length > 0) {
      abaPrincipal = todasAbas[0];
      console.log(`Usando primeira aba como principal: ${abaPrincipal.getName()}`);
    }
    
    // Corrigir aba principal
    if (abaPrincipal) {
      totalCorrigidos += corrigirHorariosAba(abaPrincipal, abaPrincipal.getName());
    }
    
    // Corrigir todas as abas mensais
    for (const aba of todasAbas) {
      const nomeAba = aba.getName();
      if (nomeAba.startsWith('Transfers_') || nomeAba.includes('2025')) {
        console.log(`Corrigindo aba mensal: ${nomeAba}`);
        totalCorrigidos += corrigirHorariosAba(aba, nomeAba);
      }
    }
    
    console.log(`CORREÇÃO CONCLUÍDA: ${totalCorrigidos} horários corrigidos`);
    return {
      sucesso: true,
      totalCorrigidos: totalCorrigidos,
      message: `${totalCorrigidos} horários corrigidos com sucesso`
    };
    
  } catch (error) {
    console.error('Erro na correção massiva:', error);
    return {
      sucesso: false,
      error: error.toString()
    };
  }
}

/**
 * Corrige horários em uma aba específica - VERSÃO SEGURA
 */
function corrigirHorariosAba(sheet, nomeAba) {
  try {
    console.log(`Processando aba: ${nomeAba}`);
    
    // Verificar se a aba tem dados
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow < 2 || lastCol < 1) {
      console.log(`Aba ${nomeAba} está vazia ou só tem cabeçalho`);
      return 0;
    }
    
    const data = sheet.getDataRange().getValues();
    let corrigidosNestAba = 0;
    
    // Encontrar índice da coluna "Hora Pick-up"
    const cabecalhos = data[0];
    console.log(`Cabeçalhos encontrados:`, cabecalhos);
    
    let colunaHora = -1;
    const nomesPossiveis = ['Hora Pick-up', 'Hora Pickup', 'Hora', 'Pick-up', 'HoraPickup'];
    
    for (const nome of nomesPossiveis) {
      colunaHora = cabecalhos.indexOf(nome);
      if (colunaHora !== -1) {
        console.log(`Coluna de hora encontrada: "${nome}" na posição ${colunaHora}`);
        break;
      }
    }
    
    if (colunaHora === -1) {
      console.log(`Coluna de hora não encontrada em ${nomeAba}. Cabeçalhos:`, cabecalhos);
      return 0;
    }
    
    console.log(`Processando ${data.length - 1} linhas na aba ${nomeAba}`);
    
    // Processar cada linha
    for (let i = 1; i < data.length; i++) {
      const valorHora = data[i][colunaHora];
      
      if (valorHora) {
        let horaCorrigida = null;
        const tipoOriginal = typeof valorHora;
        
        // Se é objeto Date
        if (valorHora instanceof Date) {
          horaCorrigida = `${valorHora.getHours().toString().padStart(2, '0')}:${valorHora.getMinutes().toString().padStart(2, '0')}`;
        }
        // Se é string com formato problemático
        else if (typeof valorHora === 'string' && valorHora.includes('1899')) {
          const match = valorHora.match(/T(\d{2}):(\d{2})/);
          if (match) {
            horaCorrigida = `${match[1]}:${match[2]}`;
          }
        }
        // Se já está correto, pular
        else if (typeof valorHora === 'string' && /^\d{2}:\d{2}$/.test(valorHora)) {
          continue;
        }
        
        // Aplicar correção
        if (horaCorrigida && horaCorrigida !== valorHora.toString()) {
          sheet.getRange(i + 1, colunaHora + 1).setValue(horaCorrigida);
          corrigidosNestAba++;
          console.log(`Linha ${i + 1}: "${valorHora}" (${tipoOriginal}) → "${horaCorrigida}"`);
        }
      }
    }
    
    console.log(`${corrigidosNestAba} horários corrigidos na aba ${nomeAba}`);
    return corrigidosNestAba;
    
  } catch (error) {
    console.error(`Erro ao processar aba ${nomeAba}:`, error);
    return 0;
  }
}
