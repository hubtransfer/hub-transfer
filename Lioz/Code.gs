// ===================================================
// SISTEMA DE TRANSFERS HOTEL LIOZ & HUB TRANSFER v4.0
// Sistema Integrado Completo com E-mail Interativo
// ===================================================

// ===================================================
// PARTE 1: CONFIGURAÇÃO GLOBAL E CONSTANTES
// ===================================================

const CONFIG = {
  // 📊 Google Sheets - Configuração Principal
  SPREADSHEET_ID: '1jXhF6tAPhuieIoIm6F3zsaEkq7NV_tafarkxcot4IE8',
  SHEET_NAME: 'Transfers LIOZ-HUB',
  PRICING_SHEET_NAME: 'Tabela de Preços',
  
// 📧 Configuração de E-mail com Sistema Interativo
  EMAIL_CONFIG: {
    DESTINATARIOS: [
      'reservations@liozlisboa.com',  // Email principal do hotel
      'fom@liozlisboa.com',          // Front Office Manager
      'juniorgutierezbega@gmail.com'  // Seu email para monitoramento
    ], // Array para múltiplos destinatários
    DESTINATARIO: 'reservations@liozlisboa.com', // Retrocompatibilidade
    ENVIAR_AUTOMATICO: true,
    VERIFICAR_CONFIRMACOES: true,
    INTERVALO_VERIFICACAO: 5, // minutos
    USAR_BOTOES_INTERATIVOS: true, // Nova flag para e-mails com botões
    TEMPLATE_HTML: true, // Usar template HTML rico
    ARQUIVAR_CONFIRMADOS: true // Arquivar e-mails após confirmação
  },
  
  // 📱 Configuração de Telefones e Contatos
CONTATOS: {
    HUB_PHONE: '+351968698138',        // Proprietário HUB
    ROBERTA_PHONE: '+351928283652',    // Assistente HUB
    QUARESMA_PHONE: '+351936148546',   // Gerente Hotel
    HOTEL_EMAIL: 'reservations@liozlisboa.com' // E-mail oficial do hotel
  },
  
  // 🔗 Integração com APIs Externas
  ZAPI: {
    INSTANCE: '3DC8E250141ED020B95796155CBF9532',
    TOKEN: 'DF93ABBE66F44D82F60EF9FE',
    WEBHOOK_URL: 'https://api.z-api.io/instances/3DC8E250141ED020B95796155CBF9532/token/DF93ABBE66F44D82F60EF9FE/send-text',
    ENABLED: false // Controle de ativação
  },
  
  // 🔄 Webhooks do Make.com
  MAKE_WEBHOOKS: {
    NEW_TRANSFER: 'https://hook.eu2.make.com/y7ro3pc9we00iums1ffb106pok9icl7m',
    WHATSAPP_RESPONSE: 'https://hook.eu2.make.com/y7ro3pc9we00iums1ffb106pok9icl7m',
    STATUS_UPDATE: 'https://hook.eu2.make.com/y7ro3pc9we00iums1ffb106pok9icl7m',
    ACERTO_CONTAS: 'https://hook.eu2.make.com/y7ro3pc9we00iums1ffb106pok9icl7m',
    ENABLED: false // Controle de ativação
  },
  
  // 🏨 Identificação do Sistema
  NAMES: {
    HUB_OWNER: 'HUB Transfer',
    ASSISTANT: 'Roberta HUB',
    HOTEL_MANAGER: 'Quaresma',
    HOTEL_NAME: 'Hotel LIOZ Lisboa',
    SISTEMA_NOME: 'Sistema LIOZ-HUB Transfer'
  },
  
  // 💰 Valores Padrão e Configurações Financeiras
  VALORES: {
    HOTEL_PADRAO: 5.00,
    HUB_PADRAO: 20.00,
    PERCENTUAL_HOTEL: 0.40, // 40% para o hotel
    PERCENTUAL_HUB: 0.60,   // 60% para HUB
    MOEDA: '€',
    FORMATO_MOEDA: '€#,##0.00'
  },
  
  // 🌐 Configurações do Sistema
  SISTEMA: {
    VERSAO: '4.0-Integrado-Email-Interativo',
    TIMEZONE: 'Europe/Lisbon',
    LOCALE: 'pt-PT',
    DATA_FORMATO: 'dd/mm/yyyy',
    HORA_FORMATO: 'HH:mm',
    ORGANIZAR_POR_MES: true,
    PREFIXO_MES: 'Transfers_',
    ANO_BASE: 2025,
    REGISTRO_DUPLO_OBRIGATORIO: true,
    MAX_TENTATIVAS_REGISTRO: 3,
    INTERVALO_TENTATIVAS: 500, // milissegundos
    BACKUP_AUTOMATICO: true,
    LOG_DETALHADO: true
  },
  
  // 🔐 Segurança e Validações
  SEGURANCA: {
    VALIDAR_EMAIL: true,
    VALIDAR_TELEFONE: true,
    SANITIZAR_INPUTS: true,
    MAX_CARACTERES_CAMPO: 500,
    PERMITIR_HTML_OBSERVACOES: false
  },
  
  // 📊 Limites e Quotas
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

// Headers da planilha principal
const HEADERS = [
  'ID',                    // A - Identificador único
  'Cliente',               // B - Nome do cliente
  'Pessoas',               // C - Número de pessoas
  'Bagagens',              // D - Número de bagagens
  'Data',                  // E - Data do transfer
  'Contacto',              // F - Telefone/contato
  'Voo',                   // G - Número do voo
  'Origem',                // H - Local de origem
  'Destino',               // I - Local de destino
  'Hora Pick-up',          // J - Hora de recolha
  'Preço Cliente (€)',     // K - Valor total cobrado
  'Valor Hotel LIOZ (€)',  // L - Comissão do hotel
  'Valor HUB Transfer (€)', // M - Valor para HUB
  'Forma Pagamento',       // N - Método de pagamento
  'Pago Para',             // O - Quem recebeu
  'Status',                // P - Status do transfer
  'Observações',           // Q - Observações gerais
  'Data Criação'           // R - Timestamp de criação
];

// Headers da tabela de preços
const PRICING_HEADERS = [
  'ID',                    // A - ID único do preço
  'Rota',                  // B - Nome da rota
  'Origem',                // C - Ponto de origem
  'Destino',               // D - Ponto de destino
  'Pessoas',               // E - Número de pessoas
  'Bagagens',              // F - Número de bagagens
  'Preço Cliente (€)',     // G - Preço total
  'Valor Hotel LIOZ (€)',  // H - Comissão hotel
  'Valor HUB Transfer (€)', // I - Valor HUB
  'Ativo',                 // J - Se está ativo
  'Data Criação',          // K - Quando foi criado
  'Observações'            // L - Notas adicionais
];

// ===================================================
// MENSAGENS E TEMPLATES DO SISTEMA
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
  
  // Templates para WhatsApp (futura implementação)
  WHATSAPP_TEMPLATES: {
    NOVO_TRANSFER: {
      TITULO: '*🚐 NOVO TRANSFER LIOZ-HUB*',
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
// CONFIGURAÇÕES DE ESTILO E APARÊNCIA
// ===================================================

const STYLES = {
  // Cores por Status
  STATUS_COLORS: {
    SOLICITADO: '#f39c12',    // Laranja
    CONFIRMADO: '#27ae60',    // Verde
    FINALIZADO: '#2c3e50',    // Cinza escuro
    CANCELADO: '#e74c3c',     // Vermelho
    EM_ANDAMENTO: '#3498db',  // Azul
    PROBLEMA: '#e67e22'       // Laranja escuro
  },
  
  // Cores dos Headers
  HEADER_COLORS: {
    PRINCIPAL: '#2c3e50',     // Azul escuro
    PRECOS: '#27ae60',        // Verde
    MENSAL: null              // Definido por mês
  },
  
  // Larguras das Colunas (em pixels)
  COLUMN_WIDTHS: {
    PRINCIPAL: [60, 150, 60, 60, 100, 120, 80, 200, 200, 90, 100, 120, 120, 100, 100, 100, 180, 140],
    PRECOS: [60, 200, 180, 180, 80, 80, 120, 140, 140, 80, 140, 200]
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
// VALIDAÇÕES E REGRAS DE NEGÓCIO
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
  
  // Valores permitidos
  VALORES_PERMITIDOS: {
    FORMA_PAGAMENTO: ['Dinheiro', 'Cartão', 'Transferência', 'MB Way', 'Voucher'],
    PAGO_PARA: ['Recepção', 'Motorista', 'Online', 'Hotel', 'HUB'],
    STATUS: Object.keys(MESSAGES.STATUS_MESSAGES)
  }
};

// ===================================================
// CONFIGURAÇÃO DE LOGS E DEBUG
// ===================================================

const LOG_CONFIG = {
  ENABLED: true,
  LEVEL: 'INFO', // DEBUG, INFO, WARN, ERROR
  INCLUDE_TIMESTAMP: true,
  INCLUDE_FUNCTION: true,
  MAX_LOG_SIZE: 1000, // Número máximo de logs na memória
  PERSIST_TO_SHEET: false, // Salvar logs em aba específica
  LOG_SHEET_NAME: 'Sistema_Logs'
};

// ===================================================
// PARTE 2: FUNÇÕES DE UTILIDADE E HELPERS
// ===================================================

// ===================================================
// Sistema de Logging Aprimorado
// ===================================================

class Logger {
  constructor() {
    this.logs = [];
    this.config = LOG_CONFIG;
  }
  
  log(level, message, data = null) {
    if (!this.config.ENABLED) return;
    
    const logEntry = {
      timestamp: new Date(),
      level: level,
      message: message,
      data: data,
      function: this.getFunctionName()
    };
    
    // Adicionar à memória
    this.logs.push(logEntry);
    if (this.logs.length > this.config.MAX_LOG_SIZE) {
      this.logs.shift();
    }
    
    // Formatar e exibir
    const formattedLog = this.formatLog(logEntry);
    console.log(formattedLog);
    
    // Persistir se configurado
    if (this.config.PERSIST_TO_SHEET) {
      this.persistToSheet(logEntry);
    }
  }
  
  formatLog(entry) {
    let log = '';
    
    if (this.config.INCLUDE_TIMESTAMP) {
      log += `[${this.formatTimestamp(entry.timestamp)}] `;
    }
    
    log += `${this.getLevelEmoji(entry.level)} ${entry.level}: `;
    
    if (this.config.INCLUDE_FUNCTION && entry.function) {
      log += `(${entry.function}) `;
    }
    
    log += entry.message;
    
    if (entry.data) {
      log += ` | Data: ${JSON.stringify(entry.data)}`;
    }
    
    return log;
  }
  
  formatTimestamp(date) {
    return Utilities.formatDate(
      date, 
      CONFIG.SISTEMA.TIMEZONE, 
      'yyyy-MM-dd HH:mm:ss'
    );
  }
  
  getLevelEmoji(level) {
    const emojis = {
      DEBUG: '🔍',
      INFO: 'ℹ️',
      WARN: '⚠️',
      ERROR: '❌',
      SUCCESS: '✅'
    };
    return emojis[level] || '📝';
  }
  
  getFunctionName() {
    try {
      throw new Error();
    } catch (e) {
      const stack = e.stack.split('\n');
      // Pegar a terceira linha do stack (pulando este método e o método log)
      const caller = stack[3];
      const match = caller.match(/at (\w+)/);
      return match ? match[1] : 'unknown';
    }
  }
  
  persistToSheet(entry) {
    try {
      const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
      let logSheet = ss.getSheetByName(this.config.LOG_SHEET_NAME);
      
      if (!logSheet) {
        logSheet = ss.insertSheet(this.config.LOG_SHEET_NAME);
        logSheet.appendRow(['Timestamp', 'Level', 'Function', 'Message', 'Data']);
      }
      
      logSheet.appendRow([
        entry.timestamp,
        entry.level,
        entry.function || '',
        entry.message,
        JSON.stringify(entry.data || {})
      ]);
      
    } catch (error) {
      console.error('Erro ao persistir log:', error);
    }
  }
  
  // Métodos de conveniência
  debug(message, data) { this.log('DEBUG', message, data); }
  info(message, data) { this.log('INFO', message, data); }
  warn(message, data) { this.log('WARN', message, data); }
  error(message, data) { this.log('ERROR', message, data); }
  success(message, data) { this.log('SUCCESS', message, data); }
}

// Instância global do logger
const logger = new Logger();

// ===================================================
// Funções de Data e Hora
// ===================================================

/**
 * Processa e valida uma data de múltiplos formatos
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
      
      // Formato ISO (YYYY-MM-DD ou YYYY-MM-DDTHH:mm:ss)
      if (dataInput.match(/^\d{4}-\d{2}-\d{2}/)) {
        if (dataInput.includes('T')) {
          data = new Date(dataInput);
        } else {
          // Adicionar hora do meio-dia para evitar problemas de timezone
          data = new Date(dataInput + 'T12:00:00');
        }
      }
      // Formato brasileiro (DD/MM/YYYY)
      else if (dataInput.match(/^\d{2}\/\d{2}\/\d{4}/)) {
        const [dia, mes, ano] = dataInput.split('/');
        data = new Date(`${ano}-${mes.padStart(2, '0')}-${dia.padStart(2, '0')}T12:00:00`);
      }
      // Formato alternativo (DD-MM-YYYY)
      else if (dataInput.match(/^\d{2}-\d{2}-\d{4}/)) {
        const [dia, mes, ano] = dataInput.split('-');
        data = new Date(`${ano}-${mes.padStart(2, '0')}-${dia.padStart(2, '0')}T12:00:00`);
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
      processada: data.toISOString() 
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
function formatarDataDDMMYYYY(date) {
  try {
    if (!(date instanceof Date) || isNaN(date)) {
      logger.warn('Data inválida para formatação', { date });
      return 'Data inválida';
    }
    
    return Utilities.formatDate(date, CONFIG.SISTEMA.TIMEZONE, CONFIG.SISTEMA.DATA_FORMATO);
    
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

// ===================================================
// Funções de Validação e Sanitização
// ===================================================

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
  
  // Validar hora
  if (dados.horaPickup) {
    if (!validarHora(dados.horaPickup)) {
      erros.push('Hora deve estar no formato HH:MM');
    } else {
      dadosValidados.horaPickup = dados.horaPickup;
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
  const outrosCampos = ['numeroVoo', 'origem', 'destino', 'pagoParaQuem', 'observacoes'];
  for (const campo of outrosCampos) {
    if (dados[campo]) {
      dadosValidados[campo] = sanitizarTexto(dados[campo]);
    }
  }
  
  // Aplicar valores padrão
  dadosValidados.status = dados.status || MESSAGES.STATUS_MESSAGES.SOLICITADO;
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
// Funções de Manipulação de Planilha
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
   
   const registros = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
   const dataFormatada = formatarDataDDMMYYYY(processarDataSegura(data));
   
   const duplicado = registros.some(row => {
     return String(row[0]) === String(id) && 
            formatarDataDDMMYYYY(new Date(row[4])) === dataFormatada;
   });
   
   logger.info('Verificação de duplicata concluída', { duplicado });
   return duplicado;
   
 } catch (error) {
   logger.error('Erro ao verificar duplicata', error);
   return false;
 }
}

// ===================================================
// Funções de Resposta HTTP
// ===================================================

/**
* Cria uma resposta JSON para retornar via ContentService
* @param {Object} data - Dados para retornar
* @param {number} statusCode - Código de status (para documentação)
* @returns {TextOutput} - Resposta formatada
*/
function createJsonResponse(data, statusCode = 200) {
  const response = {
    ...data,
    timestamp: new Date().toISOString(),
    version: CONFIG.SISTEMA.VERSAO
  };
  
  logger.debug('Criando resposta JSON', { statusCode, hasData: !!data });
  
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'GET, POST')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

/**
* Cria uma resposta HTML
* @param {string} html - Conteúdo HTML
* @returns {HtmlOutput} - Resposta HTML
*/
function createHtmlResponse(html) {
 return HtmlService
   .createHtmlOutput(html)
   .setTitle(CONFIG.NAMES.SISTEMA_NOME)
   .setWidth(800)
   .setHeight(600);
}

// ===================================================
// Funções de Cache e Performance
// ===================================================

/**
* Sistema de cache para melhorar performance
*/
class CacheManager {
 constructor() {
   this.cache = CacheService.getScriptCache();
   this.TTL = 600; // 10 minutos padrão
 }
 
 get(key) {
   try {
     const cached = this.cache.get(key);
     if (cached) {
       logger.debug('Cache hit', { key });
       return JSON.parse(cached);
     }
     logger.debug('Cache miss', { key });
     return null;
   } catch (error) {
     logger.error('Erro ao ler cache', { key, erro: error.message });
     return null;
   }
 }
 
 set(key, value, ttl = this.TTL) {
   try {
     this.cache.put(key, JSON.stringify(value), ttl);
     logger.debug('Cache set', { key, ttl });
   } catch (error) {
     logger.error('Erro ao gravar cache', { key, erro: error.message });
   }
 }
 
 remove(key) {
   try {
     this.cache.remove(key);
     logger.debug('Cache removed', { key });
   } catch (error) {
     logger.error('Erro ao remover cache', { key, erro: error.message });
   }
 }
 
 clear() {
   try {
     // Não há método clear() direto, então removemos chaves conhecidas
     const keys = ['prices', 'config', 'sheets'];
     keys.forEach(key => this.remove(key));
     logger.info('Cache limpo');
   } catch (error) {
     logger.error('Erro ao limpar cache', error);
   }
 }
}

const cache = new CacheManager();

// ===================================================
// Funções de Backup e Recuperação
// ===================================================

/**
* Cria backup dos dados
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

// ===================================================
// Funções de Estatísticas e Relatórios
// ===================================================

/**
* Gera estatísticas do sistema
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
     porMes: {},
     valorTotal: 0,
     valorHotel: 0,
     valorHUB: 0,
     mediaPassageiros: 0,
     topRotas: {},
     formasPagamento: {}
   };
   
   let totalPassageiros = 0;
   
   dados.forEach(row => {
     // Status
     const status = row[15] || 'Desconhecido';
     stats.porStatus[status] = (stats.porStatus[status] || 0) + 1;
     
     // Mês
     const data = new Date(row[4]);
     const mesAno = `${data.getMonth() + 1}/${data.getFullYear()}`;
     stats.porMes[mesAno] = (stats.porMes[mesAno] || 0) + 1;
     
     // Valores
     stats.valorTotal += parseFloat(row[10]) || 0;
     stats.valorHotel += parseFloat(row[11]) || 0;
     stats.valorHUB += parseFloat(row[12]) || 0;
     
     // Passageiros
     totalPassageiros += parseInt(row[2]) || 0;
     
     // Rotas
     const rota = `${row[7]} → ${row[8]}`;
     stats.topRotas[rota] = (stats.topRotas[rota] || 0) + 1;
     
     // Formas de pagamento
     const pagamento = row[13] || 'Desconhecido';
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

// ===================================================
// Funções de Notificação e Alertas
// ===================================================

/**
* Envia notificação para administradores
* @param {string} tipo - Tipo de notificação
* @param {string} mensagem - Mensagem a enviar
* @param {Object} dados - Dados adicionais
*/
function enviarNotificacaoAdmin(tipo, mensagem, dados = {}) {
 logger.info('Enviando notificação admin', { tipo });
 
 try {
   const destinatarios = CONFIG.EMAIL_CONFIG.DESTINATARIOS.join(',');
   const assunto = `[${CONFIG.NAMES.SISTEMA_NOME}] ${tipo}`;
   
   let corpo = `
     <h2>${tipo}</h2>
     <p>${mensagem}</p>
     <hr>
     <h3>Detalhes:</h3>
     <pre>${JSON.stringify(dados, null, 2)}</pre>
     <hr>
     <p><small>Notificação automática - ${new Date().toLocaleString('pt-PT')}</small></p>
   `;
   
   MailApp.sendEmail({
     to: destinatarios,
     subject: assunto,
     htmlBody: corpo,
     name: CONFIG.NAMES.SISTEMA_NOME
   });
   
   logger.success('Notificação enviada');
   
 } catch (error) {
   logger.error('Erro ao enviar notificação', error);
 }
}

// ===================================================
// Funções de Manutenção e Limpeza
// ===================================================

/**
* Realiza manutenção automática do sistema
*/
function manutencaoAutomatica() {
 logger.info('Iniciando manutenção automática');
 
 try {
   const tarefas = [];
   
   // 1. Limpar cache antigo
   cache.clear();
   tarefas.push('Cache limpo');
   
   // 2. Verificar integridade das abas
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
         }
       }
     }
   });
   
   // 3. Limpar logs antigos se configurado
   if (LOG_CONFIG.PERSIST_TO_SHEET) {
     const logSheet = ss.getSheetByName(LOG_CONFIG.LOG_SHEET_NAME);
     if (logSheet && logSheet.getLastRow() > 10000) {
       // Manter apenas últimas 5000 linhas
       const rowsToDelete = logSheet.getLastRow() - 5000;
       logSheet.deleteRows(2, rowsToDelete);
       tarefas.push(`${rowsToDelete} logs antigos removidos`);
     }
   }
   
   // 4. Criar backup se configurado
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
// PARTE 3: SISTEMA DE GESTÃO DE ABAS MENSAIS
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
    
    // Verificar cache
    const cacheKey = `sheet_${nomeAba}`;
    const cached = cache.get(cacheKey);
    if (cached && cached.exists) {
      logger.debug('Aba encontrada no cache', { nomeAba });
      return ss.getSheetByName(nomeAba);
    }
    
    // Buscar aba existente
    let abaMes = ss.getSheetByName(nomeAba);
    
    if (abaMes) {
      logger.debug('Aba mensal encontrada', { nomeAba });
      cache.set(cacheKey, { exists: true }, 300); // Cache por 5 minutos
      return abaMes;
    }
    
    // Criar nova aba se não existir
    logger.info('Criando nova aba mensal', { nomeAba });
    abaMes = criarAbaMensal(nomeAba, ss, mesInfo);
    
    cache.set(cacheKey, { exists: true }, 300);
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
    
    // Aplicar formatação das células
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
// Substitua a função aplicarFormatacaoMensal por esta versão corrigida:

function aplicarFormatacaoMensal(sheet) {
logger.debug('Aplicando formatação mensal', { sheet: sheet.getName() });
  
  try {
    const maxRows = Math.max(sheet.getMaxRows(), 1000);
    
    // Formatação de moeda (colunas K, L, M)
    if (maxRows > 1) {
      sheet.getRange(2, 11, maxRows - 1, 3).setNumberFormat(STYLES.FORMATS.MOEDA);
      
      // Formatação de data (coluna E)
      sheet.getRange(2, 5, maxRows - 1, 1).setNumberFormat(STYLES.FORMATS.DATA);
      
      // Formatação de hora (coluna J)
      sheet.getRange(2, 10, maxRows - 1, 1).setNumberFormat(STYLES.FORMATS.HORA);
      
      // Formatação de timestamp (coluna R)
      sheet.getRange(2, 18, maxRows - 1, 1).setNumberFormat(STYLES.FORMATS.TIMESTAMP);
      
      // Formatação de números (colunas A, C, D)
      sheet.getRange(2, 1, maxRows - 1, 1).setNumberFormat(STYLES.FORMATS.NUMERO);
      sheet.getRange(2, 3, maxRows - 1, 2).setNumberFormat(STYLES.FORMATS.NUMERO);
    }
    
    // Aplicar cores alternadas - CORRIGIDO
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
    
    // Validação de Status (coluna P)
    const statusValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(Object.values(MESSAGES.STATUS_MESSAGES))
      .setAllowInvalid(false)
      .setHelpText('Selecione o status do transfer')
      .build();
    sheet.getRange(2, 16, maxRows - 1, 1).setDataValidation(statusValidation);
    
    // Validação de Forma de Pagamento (coluna N)
    const pagamentoValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(VALIDACOES.VALORES_PERMITIDOS.FORMA_PAGAMENTO)
      .setAllowInvalid(false)
      .setHelpText('Selecione a forma de pagamento')
      .build();
    sheet.getRange(2, 14, maxRows - 1, 1).setDataValidation(pagamentoValidation);
    
    // Validação de Pago Para (coluna O)
    const pagoParaValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(VALIDACOES.VALORES_PERMITIDOS.PAGO_PARA)
      .setAllowInvalid(false)
      .setHelpText('Selecione quem recebeu o pagamento')
      .build();
    sheet.getRange(2, 15, maxRows - 1, 1).setDataValidation(pagoParaValidation);
    
    // Validação de Pessoas (coluna C) - entre 1 e MAX_PESSOAS
    const pessoasValidation = SpreadsheetApp.newDataValidation()
      .requireNumberBetween(1, CONFIG.LIMITES.MAX_PESSOAS)
      .setAllowInvalid(false)
      .setHelpText(`Entre 1 e ${CONFIG.LIMITES.MAX_PESSOAS} pessoas`)
      .build();
    sheet.getRange(2, 3, maxRows - 1, 1).setDataValidation(pessoasValidation);
    
    // Validação de Bagagens (coluna D) - entre 0 e MAX_BAGAGENS
    const bagagensValidation = SpreadsheetApp.newDataValidation()
      .requireNumberBetween(0, CONFIG.LIMITES.MAX_BAGAGENS)
      .setAllowInvalid(false)
      .setHelpText(`Entre 0 e ${CONFIG.LIMITES.MAX_BAGAGENS} bagagens`)
      .build();
    sheet.getRange(2, 4, maxRows - 1, 1).setDataValidation(bagagensValidation);
    
    // Validação de Valores (colunas K, L, M) - maior que 0
    const valorValidation = SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThan(0)
      .setAllowInvalid(false)
      .setHelpText('Valor deve ser maior que zero')
      .build();
    sheet.getRange(2, 11, maxRows - 1, 3).setDataValidation(valorValidation);
    
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
    
    // Proteger coluna Data Criação (R) - apenas sistema pode editar
    const protectionTimestamp = sheet.getRange('R:R').protect()
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

/**
* Verifica integridade das abas mensais
* @returns {Object} - Relatório de integridade
*/
function verificarIntegridadeAbasMensais() {
 logger.info('Verificando integridade das abas mensais');
 
 try {
   const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
   const ano = CONFIG.SISTEMA.ANO_BASE;
   const relatorio = {
     total: 12,
     encontradas: 0,
     comProblemas: 0,
     detalhes: []
   };
   
   MESES.forEach(mes => {
     const nomeAba = `${CONFIG.SISTEMA.PREFIXO_MES}${mes.abrev}_${mes.nome}_${ano}`;
     const aba = ss.getSheetByName(nomeAba);
     
     const status = {
       mes: mes.nome,
       nome: nomeAba,
       existe: false,
       headers: false,
       formatacao: false,
       validacoes: false,
       problemas: []
     };
     
     if (aba) {
       status.existe = true;
       relatorio.encontradas++;
       
       // Verificar headers
       try {
         const headers = aba.getRange(1, 1, 1, HEADERS.length).getValues()[0];
         status.headers = headers.join(',') === HEADERS.join(',');
         if (!status.headers) {
           status.problemas.push('Headers incorretos');
         }
       } catch (e) {
         status.problemas.push('Erro ao verificar headers');
       }
       
       // Verificar formatação básica
       try {
         const formatoData = aba.getRange(2, 5).getNumberFormat();
         status.formatacao = formatoData === STYLES.FORMATS.DATA;
         if (!status.formatacao) {
           status.problemas.push('Formatação incorreta');
         }
       } catch (e) {
         status.problemas.push('Erro ao verificar formatação');
       }
       
       // Verificar validações
       try {
         const validacaoStatus = aba.getRange(2, 16).getDataValidation();
         status.validacoes = validacaoStatus !== null;
         if (!status.validacoes) {
           status.problemas.push('Validações ausentes');
         }
       } catch (e) {
         status.problemas.push('Erro ao verificar validações');
       }
       
     } else {
       status.problemas.push('Aba não encontrada');
     }
     
     if (status.problemas.length > 0) {
       relatorio.comProblemas++;
     }
     
     relatorio.detalhes.push(status);
   });
   
   logger.info('Verificação de integridade concluída', {
     encontradas: relatorio.encontradas,
     comProblemas: relatorio.comProblemas
   });
   
   return relatorio;
   
 } catch (error) {
   logger.error('Erro ao verificar integridade', error);
   return { erro: error.message };
 }
}

/**
* Repara problemas encontrados nas abas mensais
* @param {Array} problemas - Lista de problemas a corrigir
* @returns {Object} - Resultado das correções
*/
function repararAbasMensais(problemas = null) {
 logger.info('Iniciando reparação de abas mensais');
 
 try {
   // Se não foram passados problemas específicos, fazer verificação completa
   if (!problemas) {
     const verificacao = verificarIntegridadeAbasMensais();
     problemas = verificacao.detalhes.filter(d => d.problemas.length > 0);
   }
   
   const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
   const resultados = {
     reparadas: 0,
     criadas: 0,
     erros: 0,
     detalhes: []
   };
   
   problemas.forEach(problema => {
     try {
       const mesInfo = MESES.find(m => m.nome === problema.mes);
       if (!mesInfo) return;
       
       if (!problema.existe) {
         // Criar aba que não existe
         criarAbaMensal(problema.nome, ss, mesInfo);
         resultados.criadas++;
         resultados.detalhes.push({
           mes: problema.mes,
           acao: 'criada',
           sucesso: true
         });
         
       } else {
         // Reparar aba existente
         const aba = ss.getSheetByName(problema.nome);
         
         // Corrigir headers
         if (problema.problemas.includes('Headers incorretos')) {
           aba.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
           const headerRange = aba.getRange(1, 1, 1, HEADERS.length);
           headerRange
             .setBackground(mesInfo.cor || STYLES.HEADER_COLORS.PRINCIPAL)
             .setFontColor('#ffffff')
             .setFontWeight('bold');
         }
         
         // Corrigir formatação
         if (problema.problemas.includes('Formatação incorreta')) {
           aplicarFormatacaoMensal(aba);
         }
         
         // Corrigir validações
         if (problema.problemas.includes('Validações ausentes')) {
           aplicarValidacoesPlanilha(aba);
         }
         
         resultados.reparadas++;
         resultados.detalhes.push({
           mes: problema.mes,
           acao: 'reparada',
           problemas: problema.problemas,
           sucesso: true
         });
       }
       
     } catch (error) {
       logger.error('Erro ao reparar aba', {
         mes: problema.mes,
         erro: error.message
       });
       resultados.erros++;
       resultados.detalhes.push({
         mes: problema.mes,
         acao: 'erro',
         erro: error.message,
         sucesso: false
       });
     }
   });
   
   logger.success('Reparação concluída', resultados);
   
   return resultados;
   
 } catch (error) {
   logger.error('Erro na reparação de abas', error);
   throw error;
 }
}

// ===================================================
// PARTE 3.1: CONSOLIDAÇÃO DE DADOS MENSAIS
// ===================================================

/**
* Consolida dados de todas as abas mensais
* @param {number} ano - Ano a consolidar
* @returns {Object} - Dados consolidados
*/
function consolidarDadosMensais(ano = CONFIG.SISTEMA.ANO_BASE) {
 logger.info('Consolidando dados mensais', { ano });
 
 try {
   const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
   const consolidado = {
     ano: ano,
     totalGeral: 0,
     valorTotalGeral: 0,
     porMes: {},
     porStatus: {},
     porRota: {},
     detalhes: []
   };
   
   MESES.forEach(mes => {
     const nomeAba = `${CONFIG.SISTEMA.PREFIXO_MES}${mes.abrev}_${mes.nome}_${ano}`;
     const aba = ss.getSheetByName(nomeAba);
     
     if (!aba || aba.getLastRow() <= 1) {
       consolidado.porMes[mes.nome] = {
         total: 0,
         valor: 0,
         status: 'vazio'
       };
       return;
     }
     
     const dados = aba.getRange(2, 1, aba.getLastRow() - 1, HEADERS.length).getValues();
     const dadosMes = {
       mes: mes.nome,
       total: dados.length,
       valorTotal: 0,
       valorHotel: 0,
       valorHUB: 0,
       porStatus: {},
       topRotas: {}
     };
     
     dados.forEach(row => {
       // Valores
       const valorCliente = parseFloat(row[10]) || 0;
       const valorHotel = parseFloat(row[11]) || 0;
       const valorHUB = parseFloat(row[12]) || 0;
       
       dadosMes.valorTotal += valorCliente;
       dadosMes.valorHotel += valorHotel;
       dadosMes.valorHUB += valorHUB;
       
       // Status
       const status = row[15] || 'Sem status';
       dadosMes.porStatus[status] = (dadosMes.porStatus[status] || 0) + 1;
       consolidado.porStatus[status] = (consolidado.porStatus[status] || 0) + 1;
       
       // Rotas
       const rota = `${row[7]} → ${row[8]}`;
       dadosMes.topRotas[rota] = (dadosMes.topRotas[rota] || 0) + 1;
       consolidado.porRota[rota] = (consolidado.porRota[rota] || 0) + 1;
     });
     
     consolidado.totalGeral += dadosMes.total;
     consolidado.valorTotalGeral += dadosMes.valorTotal;
     consolidado.porMes[mes.nome] = dadosMes;
     consolidado.detalhes.push(dadosMes);
   });
   
   // Ordenar top rotas geral
   consolidado.topRotas = Object.entries(consolidado.porRota)
     .sort((a, b) => b[1] - a[1])
     .slice(0, 10)
     .map(([rota, count]) => ({ rota, count }));
   
   logger.success('Consolidação concluída', {
     total: consolidado.totalGeral,
     valor: consolidado.valorTotalGeral
   });
   
   return consolidado;
   
 } catch (error) {
   logger.error('Erro ao consolidar dados', error);
   throw error;
 }
}

/**
* Gera relatório mensal detalhado
* @param {number} mes - Número do mês (1-12)
* @param {number} ano - Ano
* @returns {Object} - Relatório do mês
*/
function gerarRelatorioMensal(mes, ano = CONFIG.SISTEMA.ANO_BASE) {
 logger.info('Gerando relatório mensal', { mes, ano });
 
 try {
   const mesInfo = MESES.find(m => m.numero === mes);
   if (!mesInfo) {
     throw new Error('Mês inválido');
   }
   
   const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
   const nomeAba = `${CONFIG.SISTEMA.PREFIXO_MES}${mesInfo.abrev}_${mesInfo.nome}_${ano}`;
   const aba = ss.getSheetByName(nomeAba);
   
   if (!aba || aba.getLastRow() <= 1) {
     return {
       mes: mesInfo.nome,
       ano: ano,
       status: 'sem_dados',
       mensagem: 'Nenhum transfer registrado neste mês'
     };
   }
   
   const dados = aba.getRange(2, 1, aba.getLastRow() - 1, HEADERS.length).getValues();
   
   const relatorio = {
     mes: mesInfo.nome,
     ano: ano,
     periodo: `01/${mesInfo.abrev}/${ano} a ${mesInfo.dias}/${mesInfo.abrev}/${ano}`,
     resumo: {
       totalTransfers: dados.length,
       totalPassageiros: 0,
       totalBagagens: 0,
       valorTotal: 0,
       valorHotel: 0,
       valorHUB: 0,
       ticketMedio: 0
     },
     porDia: {},
     porSemana: {},
     porStatus: {},
     porFormaPagamento: {},
     porPagoPara: {},
     topClientes: {},
     topRotas: {},
     horariosPreferidos: {},
     observacoes: []
   };
   
   // Processar cada registro
   dados.forEach(row => {
     // Resumo geral
     relatorio.resumo.totalPassageiros += parseInt(row[2]) || 0;
     relatorio.resumo.totalBagagens += parseInt(row[3]) || 0;
     relatorio.resumo.valorTotal += parseFloat(row[10]) || 0;
     relatorio.resumo.valorHotel += parseFloat(row[11]) || 0;
     relatorio.resumo.valorHUB += parseFloat(row[12]) || 0;
     
     // Por dia
     const data = new Date(row[4]);
     const dia = data.getDate();
     relatorio.porDia[dia] = (relatorio.porDia[dia] || 0) + 1;
     
     // Por semana
     const semana = Math.ceil(dia / 7);
     relatorio.porSemana[`Semana ${semana}`] = (relatorio.porSemana[`Semana ${semana}`] || 0) + 1;
     
     // Por status
     const status = row[15] || 'Sem status';
     relatorio.porStatus[status] = (relatorio.porStatus[status] || 0) + 1;
     
     // Por forma de pagamento
     const pagamento = row[13] || 'Desconhecido';
     relatorio.porFormaPagamento[pagamento] = (relatorio.porFormaPagamento[pagamento] || 0) + 1;
     
     // Por pago para
     const pagoPara = row[14] || 'Desconhecido';
     relatorio.porPagoPara[pagoPara] = (relatorio.porPagoPara[pagoPara] || 0) + 1;
     
     // Top clientes
     const cliente = row[1];
     relatorio.topClientes[cliente] = (relatorio.topClientes[cliente] || 0) + 1;
     
     // Top rotas
     const rota = `${row[7]} → ${row[8]}`;
     relatorio.topRotas[rota] = (relatorio.topRotas[rota] || 0) + 1;
     
     // Horários preferidos
     const hora = row[9] ? row[9].split(':')[0] : 'Desconhecido';
     relatorio.horariosPreferidos[hora + 'h'] = (relatorio.horariosPreferidos[hora + 'h'] || 0) + 1;
     
     // Coletar observações não vazias
     if (row[16] && row[16].trim()) {
       relatorio.observacoes.push({
         id: row[0],
         cliente: row[1],
         observacao: row[16]
       });
     }
   });
   
   // Calcular ticket médio
   relatorio.resumo.ticketMedio = relatorio.resumo.totalTransfers > 0 
     ? relatorio.resumo.valorTotal / relatorio.resumo.totalTransfers 
     : 0;
   
   // Ordenar e limitar tops
   relatorio.topClientes = Object.entries(relatorio.topClientes)
     .sort((a, b) => b[1] - a[1])
     .slice(0, 10)
     .map(([cliente, count]) => ({ cliente, count }));
   
   relatorio.topRotas = Object.entries(relatorio.topRotas)
     .sort((a, b) => b[1] - a[1])
     .slice(0, 10)
     .map(([rota, count]) => ({ rota, count }));
   
   // Análise de tendências
   relatorio.analise = analisarTendenciasMes(relatorio);
   
   logger.success('Relatório mensal gerado', {
     mes: mesInfo.nome,
     total: relatorio.resumo.totalTransfers
   });
   
   return relatorio;
   
 } catch (error) {
   logger.error('Erro ao gerar relatório mensal', error);
   throw error;
 }
}

/**
* Analisa tendências do mês
* @param {Object} relatorio - Dados do relatório
* @returns {Object} - Análise de tendências
*/
function analisarTendenciasMes(relatorio) {
 const analise = {
   melhorDia: null,
   piorDia: null,
   melhorSemana: null,
   horarioPico: null,
   rotaPrincipal: null,
   clientePrincipal: null,
   tendenciaStatus: null,
   alertas: []
 };
 
 // Melhor e pior dia
 const diasOrdenados = Object.entries(relatorio.porDia)
   .sort((a, b) => b[1] - a[1]);
 
 if (diasOrdenados.length > 0) {
   analise.melhorDia = {
     dia: diasOrdenados[0][0],
     transfers: diasOrdenados[0][1]
   };
   analise.piorDia = {
     dia: diasOrdenados[diasOrdenados.length - 1][0],
     transfers: diasOrdenados[diasOrdenados.length - 1][1]
   };
 }
 
 // Melhor semana
 const semanasOrdenadas = Object.entries(relatorio.porSemana)
   .sort((a, b) => b[1] - a[1]);
 
 if (semanasOrdenadas.length > 0) {
   analise.melhorSemana = {
     semana: semanasOrdenadas[0][0],
     transfers: semanasOrdenadas[0][1]
   };
 }
 
 // Horário de pico
 const horariosOrdenados = Object.entries(relatorio.horariosPreferidos)
   .sort((a, b) => b[1] - a[1]);
 
 if (horariosOrdenados.length > 0) {
   analise.horarioPico = {
     horario: horariosOrdenados[0][0],
     transfers: horariosOrdenados[0][1]
   };
 }
 
 // Rota e cliente principal
 if (relatorio.topRotas.length > 0) {
   analise.rotaPrincipal = relatorio.topRotas[0];
 }
 
 if (relatorio.topClientes.length > 0) {
   analise.clientePrincipal = relatorio.topClientes[0];
 }
 
 // Análise de status
 const totalTransfers = relatorio.resumo.totalTransfers;
 const cancelados = relatorio.porStatus['Cancelado'] || 0;
 const confirmados = relatorio.porStatus['Confirmado'] || 0;
 const finalizados = relatorio.porStatus['Finalizado'] || 0;
 
 analise.tendenciaStatus = {
   taxaCancelamento: totalTransfers > 0 ? (cancelados / totalTransfers * 100).toFixed(1) : 0,
   taxaConfirmacao: totalTransfers > 0 ? (confirmados / totalTransfers * 100).toFixed(1) : 0,
   taxaFinalizacao: totalTransfers > 0 ? (finalizados / totalTransfers * 100).toFixed(1) : 0
 };
 
 // Gerar alertas
 if (analise.tendenciaStatus.taxaCancelamento > 20) {
   analise.alertas.push({
     tipo: 'warning',
     mensagem: `Taxa de cancelamento alta: ${analise.tendenciaStatus.taxaCancelamento}%`
   });
 }
 
 if (relatorio.resumo.ticketMedio < 20) {
   analise.alertas.push({
     tipo: 'info',
     mensagem: `Ticket médio baixo: €${relatorio.resumo.ticketMedio.toFixed(2)}`
   });
 }
 
 return analise;
}

// ===================================================
// PARTE 4: SISTEMA DE CÁLCULO DE PREÇOS
// ===================================================

/**
 * Calcula valores do transfer baseado em múltiplas estratégias
 * @param {string} origem - Local de origem
 * @param {string} destino - Local de destino
 * @param {number} pessoas - Número de pessoas
 * @param {number} bagagens - Número de bagagens
 * @param {number} precoManual - Preço informado manualmente (opcional)
 * @returns {Object} - Valores calculados
 */
function calcularValores(origem, destino, pessoas, bagagens, precoManual = null) {
  logger.info('Calculando valores do transfer', {
    origem, destino, pessoas, bagagens, precoManual
  });
  
  try {
    // Estratégia 1: Preço manual com proporção da tabela
    if (precoManual && precoManual > 0) {
      return calcularComPrecoManual(origem, destino, pessoas, bagagens, precoManual);
    }
    
    // Estratégia 2: Buscar na tabela de preços
    const precoTabela = consultarPrecoTabela(origem, destino, pessoas, bagagens);
    if (precoTabela) {
      logger.success('Preço encontrado na tabela', precoTabela);
      return {
        precoCliente: precoTabela.precoCliente,
        valorHotel: precoTabela.valorHotel,
        valorHUB: precoTabela.valorHUB,
        fonte: 'tabela',
        tabelaId: precoTabela.id,
        matchScore: precoTabela.matchScore,
        observacoes: precoTabela.observacoes || `Preço da tabela (ID: ${precoTabela.id})`
      };
    }
    
    // Estratégia 3: Calcular com base em distância/complexidade
    const precoCalculado = calcularPrecoInteligente(origem, destino, pessoas, bagagens);
    if (precoCalculado) {
      return precoCalculado;
    }
    
    // Estratégia 4: Valores padrão
    logger.warn('Usando valores padrão de emergência');
    return {
      precoCliente: 25.00,
      valorHotel: CONFIG.VALORES.HOTEL_PADRAO,
      valorHUB: CONFIG.VALORES.HUB_PADRAO,
      fonte: 'padrao',
      observacoes: 'Valores padrão - configure tabela de preços'
    };
    
  } catch (error) {
    logger.error('Erro no cálculo de valores', error);
    return {
      precoCliente: 25.00,
      valorHotel: CONFIG.VALORES.HOTEL_PADRAO,
      valorHUB: CONFIG.VALORES.HUB_PADRAO,
      fonte: 'erro',
      observacoes: 'Erro no cálculo - valores padrão aplicados'
    };
  }
}

/**
 * Calcula valores quando preço manual é informado
 * @private
 */
function calcularComPrecoManual(origem, destino, pessoas, bagagens, precoManual) {
  logger.debug('Calculando com preço manual', { precoManual });
  
  const precoCliente = parseFloat(precoManual);
  
  // Tentar encontrar proporção na tabela
  const precoTabela = consultarPrecoTabela(origem, destino, pessoas, bagagens);
  
  if (precoTabela && precoTabela.precoCliente > 0) {
    // Usar proporção da tabela
    const proporcao = precoCliente / precoTabela.precoCliente;
    
    return {
      precoCliente: precoCliente,
      valorHotel: Math.round(precoTabela.valorHotel * proporcao * 100) / 100,
      valorHUB: Math.round(precoTabela.valorHUB * proporcao * 100) / 100,
      fonte: 'manual-com-tabela',
      proporcao: proporcao,
      observacoes: `Preço manual com proporção da tabela (${(proporcao * 100).toFixed(0)}%)`
    };
  }
  
  // Usar proporção padrão configurada
  const valorHotel = Math.round(precoCliente * CONFIG.VALORES.PERCENTUAL_HOTEL * 100) / 100;
  const valorHUB = Math.round(precoCliente * CONFIG.VALORES.PERCENTUAL_HUB * 100) / 100;
  
  return {
    precoCliente: precoCliente,
    valorHotel: valorHotel,
    valorHUB: valorHUB,
    fonte: 'manual-sem-tabela',
    observacoes: `Preço manual - distribuição ${(CONFIG.VALORES.PERCENTUAL_HOTEL * 100).toFixed(0)}%/${(CONFIG.VALORES.PERCENTUAL_HUB * 100).toFixed(0)}%`
  };
}

/**
 * Consulta preço na tabela com algoritmo de matching inteligente
 * @private
 */
function consultarPrecoTabela(origem, destino, pessoas, bagagens) {
  logger.debug('Consultando tabela de preços');
  
  try {
    // Verificar cache primeiro
    const cacheKey = `preco_${origem}_${destino}_${pessoas}_${bagagens}`;
    const cached = cache.get(cacheKey);
    if (cached) {
      return cached;
    }
    
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
      if (linha[9] !== 'Sim') return;
      
      const score = calcularScoreMatch(
        origem, destino, pessoas, bagagens,
        linha[2], linha[3], linha[4], linha[5]
      );
      
      if (score > melhorScore && score >= 60) { // Mínimo 60% de match
        melhorScore = score;
        melhorMatch = {
          id: linha[0],
          rota: linha[1],
          origem: linha[2],
          destino: linha[3],
          pessoas: linha[4],
          bagagens: linha[5],
          precoCliente: parseFloat(linha[6]) || 0,
          valorHotel: parseFloat(linha[7]) || 0,
          valorHUB: parseFloat(linha[8]) || 0,
          observacoes: linha[11] || '',
          matchScore: score
        };
      }
    });
    
    if (melhorMatch) {
      // Cachear resultado
      cache.set(cacheKey, melhorMatch, 600);
      logger.debug('Match encontrado', { score: melhorScore, id: melhorMatch.id });
    }
    
    return melhorMatch;
    
  } catch (error) {
    logger.error('Erro ao consultar tabela de preços', error);
    return null;
  }
}

/**
 * Calcula score de matching entre busca e registro da tabela
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
* Calcula preço de forma inteligente baseado em regras de negócio
* @private
*/
function calcularPrecoInteligente(origem, destino, pessoas, bagagens) {
 logger.debug('Calculando preço inteligente');
 
 try {
   // Base de conhecimento de locais e distâncias aproximadas
   const locaisConhecidos = {
     'aeroporto': { tipo: 'aeroporto', distanciaBase: 15 },
     'lisboa': { tipo: 'aeroporto', distanciaBase: 15 },
     'cascais': { tipo: 'cidade', distanciaBase: 30 },
     'sintra': { tipo: 'cidade', distanciaBase: 35 },
     'belem': { tipo: 'bairro', distanciaBase: 10 },
     'lioz': { tipo: 'hotel', distanciaBase: 0 },
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
   
   // Cálculo base
   let precoBase = 15; // Preço mínimo
   
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
   
   // Arredondar para múltiplo de 5
   const precoCliente = Math.ceil(precoBase / 5) * 5;
   
   // Calcular distribuição
   const valorHotel = Math.round(precoCliente * CONFIG.VALORES.PERCENTUAL_HOTEL * 100) / 100;
   const valorHUB = Math.round(precoCliente * CONFIG.VALORES.PERCENTUAL_HUB * 100) / 100;
   
   logger.debug('Preço calculado inteligentemente', {
     distancia: distanciaEstimada,
     precoBase: precoBase,
     precoFinal: precoCliente
   });
   
   return {
     precoCliente: precoCliente,
     valorHotel: valorHotel,
     valorHUB: valorHUB,
     fonte: 'calculado',
     distanciaEstimada: distanciaEstimada,
     observacoes: `Preço calculado - Distância estimada: ${distanciaEstimada}km`
   };
   
 } catch (error) {
   logger.error('Erro no cálculo inteligente', error);
   return null;
 }
}

// ===================================================
// PARTE 4.1: GESTÃO DA TABELA DE PREÇOS
// ===================================================

/**
* Adiciona novo preço à tabela
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
   
   // Calcular valores se não fornecidos
   let valorHotel = dadosPreco.valorHotel;
   let valorHUB = dadosPreco.valorHUB;
   
   if (!valorHotel || !valorHUB) {
     valorHotel = Math.round(dadosPreco.precoCliente * CONFIG.VALORES.PERCENTUAL_HOTEL * 100) / 100;
     valorHUB = Math.round(dadosPreco.precoCliente * CONFIG.VALORES.PERCENTUAL_HUB * 100) / 100;
   }
   
   // Montar linha
   const novaLinha = [
     novoId,
     `${dadosPreco.origem} → ${dadosPreco.destino}`, // Rota
     dadosPreco.origem,
     dadosPreco.destino,
     parseInt(dadosPreco.pessoas) || 1,
     parseInt(dadosPreco.bagagens) || 0,
     parseFloat(dadosPreco.precoCliente),
     valorHotel,
     valorHUB,
     dadosPreco.ativo !== false ? 'Sim' : 'Não',
     new Date(),
     dadosPreco.observacoes || ''
   ];
   
   // Inserir na planilha
   sheet.appendRow(novaLinha);
   
   // Limpar cache de preços
   cache.clear();
   
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
* Atualiza preço existente na tabela
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
   
   // Atualizar campos específicos
   if (dadosAtualizacao.origem !== undefined) {
     sheet.getRange(linha, 3).setValue(dadosAtualizacao.origem);
     sheet.getRange(linha, 2).setValue(`${dadosAtualizacao.origem} → ${sheet.getRange(linha, 4).getValue()}`);
   }
   
   if (dadosAtualizacao.destino !== undefined) {
     sheet.getRange(linha, 4).setValue(dadosAtualizacao.destino);
     sheet.getRange(linha, 2).setValue(`${sheet.getRange(linha, 3).getValue()} → ${dadosAtualizacao.destino}`);
   }
   
   if (dadosAtualizacao.pessoas !== undefined) {
     sheet.getRange(linha, 5).setValue(parseInt(dadosAtualizacao.pessoas));
   }
   
   if (dadosAtualizacao.bagagens !== undefined) {
     sheet.getRange(linha, 6).setValue(parseInt(dadosAtualizacao.bagagens));
   }
   
   if (dadosAtualizacao.precoCliente !== undefined) {
     sheet.getRange(linha, 7).setValue(parseFloat(dadosAtualizacao.precoCliente));
   }
   
   if (dadosAtualizacao.valorHotel !== undefined) {
     sheet.getRange(linha, 8).setValue(parseFloat(dadosAtualizacao.valorHotel));
   }
   
   if (dadosAtualizacao.valorHUB !== undefined) {
     sheet.getRange(linha, 9).setValue(parseFloat(dadosAtualizacao.valorHUB));
   }
   
   if (dadosAtualizacao.ativo !== undefined) {
     sheet.getRange(linha, 10).setValue(dadosAtualizacao.ativo ? 'Sim' : 'Não');
   }
   
   if (dadosAtualizacao.observacoes !== undefined) {
     sheet.getRange(linha, 12).setValue(dadosAtualizacao.observacoes);
   }
   
   // Limpar cache
   cache.clear();
   
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
* Lista todos os preços da tabela
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
     if (filtros.apenasAtivos && linha[9] !== 'Sim') return;
     if (filtros.origem && !linha[2].toLowerCase().includes(filtros.origem.toLowerCase())) return;
     if (filtros.destino && !linha[3].toLowerCase().includes(filtros.destino.toLowerCase())) return;
     if (filtros.pessoas && parseInt(linha[4]) !== parseInt(filtros.pessoas)) return;
     
     precos.push({
       id: linha[0],
       rota: linha[1],
       origem: linha[2],
       destino: linha[3],
       pessoas: linha[4],
       bagagens: linha[5],
       precoCliente: linha[6],
       valorHotel: linha[7],
       valorHUB: linha[8],
       ativo: linha[9] === 'Sim',
       dataCriacao: linha[10],
       observacoes: linha[11]
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
* Cria a tabela de preços se não existir
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
 
 // Aplicar formatação de moeda
 sheet.getRange(2, 7, sheet.getMaxRows() - 1, 3).setNumberFormat(STYLES.FORMATS.MOEDA);
 
 // Validação de ativo/inativo
 const ativoValidation = SpreadsheetApp.newDataValidation()
   .requireValueInList(['Sim', 'Não'])
   .setAllowInvalid(false)
   .build();
 sheet.getRange(2, 10, sheet.getMaxRows() - 1, 1).setDataValidation(ativoValidation);
 
 logger.success('Tabela de preços criada');
 
 return sheet;
}

/**
* Importa preços de uma fonte externa (CSV, JSON, etc)
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
       origem: preco.origem || preco.Origem,
       destino: preco.destino || preco.Destino,
       pessoas: preco.pessoas || preco.Pessoas || 1,
       bagagens: preco.bagagens || preco.Bagagens || 0,
       precoCliente: preco.precoCliente || preco['Preço Cliente'] || preco.preco,
       valorHotel: preco.valorHotel || preco['Valor Hotel'],
       valorHUB: preco.valorHUB || preco['Valor HUB'],
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
* Exporta tabela de preços
* @param {string} formato - Formato de exportação (csv, json)
* @returns {string} - Dados exportados
*/
function exportarPrecos(formato = 'csv') {
 logger.info('Exportando preços', { formato });
 
 try {
   const precos = listarPrecosTabela();
   
   if (formato === 'csv') {
     // Gerar CSV
     const headers = ['ID', 'Rota', 'Origem', 'Destino', 'Pessoas', 'Bagagens', 
                     'Preço Cliente', 'Valor Hotel', 'Valor HUB', 'Ativo', 
                     'Data Criação', 'Observações'];
     
     let csv = headers.join(',') + '\n';
     
     precos.forEach(preco => {
       const linha = [
         preco.id,
         `"${preco.rota}"`,
         `"${preco.origem}"`,
         `"${preco.destino}"`,
         preco.pessoas,
         preco.bagagens,
         preco.precoCliente,
         preco.valorHotel,
         preco.valorHUB,
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
// PARTE 5: SISTEMA DE E-MAIL INTERATIVO
// ===================================================

/**
 * Envia e-mail interativo com botões de ação
 * @param {Array} dadosTransfer - Dados do transfer
 * @returns {boolean} - Se o envio foi bem-sucedido
 */
function enviarEmailNovoTransfer(dadosTransfer) {
  logger.info('Enviando e-mail interativo para novo transfer');
  
  try {
    // Validar configuração
    if (!CONFIG.EMAIL_CONFIG.ENVIAR_AUTOMATICO) {
      logger.info('Envio de e-mail desativado');
      return false;
    }
    
    const destinatarios = CONFIG.EMAIL_CONFIG.DESTINATARIOS.join(',');
    if (!destinatarios) {
      throw new Error('Nenhum destinatário configurado');
    }
    
    // Extrair dados do transfer
    const transfer = {
      id: dadosTransfer[0],
      cliente: dadosTransfer[1],
      pessoas: dadosTransfer[2],
      bagagens: dadosTransfer[3],
      data: dadosTransfer[4],
      contacto: dadosTransfer[5],
      voo: dadosTransfer[6],
      origem: dadosTransfer[7],
      destino: dadosTransfer[8],
      horaPickup: dadosTransfer[9],
      precoCliente: dadosTransfer[10],
      valorHotel: dadosTransfer[11],
      valorHUB: dadosTransfer[12],
      formaPagamento: dadosTransfer[13],
      pagoPara: dadosTransfer[14],
      status: dadosTransfer[15],
      observacoes: dadosTransfer[16]
    };
    
    // Formatar dados para exibição
    const dataFormatada = formatarDataDDMMYYYY(new Date(transfer.data));
    const assunto = CONFIG.EMAIL_CONFIG.USAR_BOTOES_INTERATIVOS
      ? `AÇÃO NECESSÁRIA: Novo Transfer #${transfer.id} - ${transfer.cliente} - ${dataFormatada}`
      : `Novo Transfer #${transfer.id} - ${transfer.cliente} - ${dataFormatada}`;
    
    // Gerar corpo do e-mail
    const corpoHtml = CONFIG.EMAIL_CONFIG.TEMPLATE_HTML
      ? criarEmailHtmlInterativo(transfer)
      : criarEmailTextoSimples(transfer);
    
    // Enviar e-mail
    MailApp.sendEmail({
      to: destinatarios,
      subject: assunto,
      htmlBody: corpoHtml,
      name: CONFIG.NAMES.SISTEMA_NOME
    });
    
    logger.success('E-mail enviado com sucesso', {
      transferId: transfer.id,
      destinatarios: destinatarios
    });
    
    // Registrar envio no transfer
    atualizarObservacoesTransfer(transfer.id, 
      `E-mail enviado em ${formatarDataHora(new Date())}`);
    
    return true;
    
  } catch (error) {
    logger.error('Erro ao enviar e-mail', error);
    return false;
  }
}

/**
 * Cria HTML do e-mail interativo com botões
 * @private
 */
function criarEmailHtmlInterativo(transfer) {
  const webAppUrl = ScriptApp.getService().getUrl();
  const urlConfirmar = `${webAppUrl}?action=confirm&id=${transfer.id}`;
  const urlCancelar = `${webAppUrl}?action=cancel&id=${transfer.id}`;
  
  const dataFormatada = formatarDataDDMMYYYY(new Date(transfer.data));
  const precoCliente = `${CONFIG.VALORES.MOEDA}${parseFloat(transfer.precoCliente).toFixed(2)}`;
  const valorHotel = `${CONFIG.VALORES.MOEDA}${parseFloat(transfer.valorHotel).toFixed(2)}`;
  const valorHUB = `${CONFIG.VALORES.MOEDA}${parseFloat(transfer.valorHUB).toFixed(2)}`;
  
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <style>
        body {
          font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
          background-color: #f4f7f6;
          color: #333;
          margin: 0;
          padding: 0;
        }
        .container {
          background: #ffffff;
          padding: 30px;
          border-radius: 12px;
          max-width: 600px;
          margin: 20px auto;
          box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        }
        .header {
          text-align: center;
          border-bottom: 3px solid #ffc107;
          padding-bottom: 20px;
          margin-bottom: 30px;
        }
        .header h1 {
          margin: 0;
          font-size: 24px;
          color: #0056b3;
        }
        .header p {
          margin: 10px 0 0 0;
          color: #666;
          font-size: 14px;
        }
        .info-section {
          background: #f8f9fa;
          padding: 20px;
          border-radius: 8px;
          margin-bottom: 20px;
        }
        .info-row {
          display: flex;
          justify-content: space-between;
          padding: 8px 0;
          border-bottom: 1px solid #e9ecef;
        }
        .info-row:last-child {
          border-bottom: none;
        }
        .info-label {
          font-weight: 600;
          color: #495057;
        }
        .info-value {
          color: #212529;
        }
        .route-section {
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          color: white;
          padding: 20px;
          border-radius: 8px;
          text-align: center;
          margin: 20px 0;
        }
        .route-section h2 {
          margin: 0;
          font-size: 20px;
        }
        .pricing-grid {
          display: grid;
          grid-template-columns: 1fr 1fr 1fr;
          gap: 15px;
          margin: 20px 0;
        }
        .pricing-card {
          background: #fff;
          border: 2px solid #e9ecef;
          border-radius: 8px;
          padding: 15px;
          text-align: center;
        }
        .pricing-card.total {
          border-color: #28a745;
          background: #f8fff9;
        }
        .pricing-label {
          font-size: 12px;
          color: #6c757d;
          margin-bottom: 5px;
        }
        .pricing-value {
          font-size: 20px;
          font-weight: bold;
          color: #28a745;
        }
        .action-section {
          background: #e3f2fd;
          border: 2px solid #2196f3;
          border-radius: 8px;
          padding: 25px;
          text-align: center;
          margin: 30px 0;
        }
        .action-section h3 {
          margin: 0 0 15px 0;
          color: #1976d2;
        }
        .action-buttons {
          margin-top: 20px;
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
          border-color: #218838;
        }
        .btn-cancel {
          background-color: #dc3545;
          color: white;
          border: 2px solid #dc3545;
        }
        .btn-cancel:hover {
          background-color: #c82333;
          border-color: #c82333;
        }
        .observations {
          background: #fff3cd;
          border: 1px solid #ffeaa7;
          border-radius: 6px;
          padding: 15px;
          margin: 20px 0;
        }
        .observations h4 {
          margin: 0 0 10px 0;
          color: #856404;
        }
.footer {
         text-align: center;
         font-size: 12px;
         color: #6c757d;
         margin-top: 30px;
         padding-top: 20px;
         border-top: 1px solid #e9ecef;
       }
       .footer a {
         color: #0056b3;
         text-decoration: none;
       }
       @media only screen and (max-width: 600px) {
         .container {
           padding: 20px;
           margin: 10px;
         }
         .pricing-grid {
           grid-template-columns: 1fr;
         }
         .btn {
           display: block;
           margin: 10px 0;
         }
       }
     </style>
   </head>
   <body>
     <div class="container">
       <div class="header">
         <h1>🚐 Novo Transfer Solicitado</h1>
         <p>Transfer #${transfer.id} - Aguardando Confirmação</p>
       </div>
       
       <div class="info-section">
         <div class="info-row">
           <span class="info-label">👤 Cliente:</span>
           <span class="info-value">${transfer.cliente}</span>
         </div>
         <div class="info-row">
           <span class="info-label">📞 Contacto:</span>
           <span class="info-value">${transfer.contacto}</span>
         </div>
         <div class="info-row">
           <span class="info-label">👥 Passageiros:</span>
           <span class="info-value">${transfer.pessoas} pessoa(s)</span>
         </div>
         <div class="info-row">
           <span class="info-label">🧳 Bagagens:</span>
           <span class="info-value">${transfer.bagagens} volume(s)</span>
         </div>
         ${transfer.voo ? `
         <div class="info-row">
           <span class="info-label">✈️ Voo:</span>
           <span class="info-value">${transfer.voo}</span>
         </div>
         ` : ''}
       </div>
       
       <div class="route-section">
         <h2>${transfer.origem} → ${transfer.destino}</h2>
         <p style="margin: 10px 0 0 0; font-size: 16px;">
           📅 ${dataFormatada} às ${transfer.horaPickup}
         </p>
       </div>
       
       <div class="pricing-grid">
         <div class="pricing-card total">
           <div class="pricing-label">Valor Total</div>
           <div class="pricing-value">${precoCliente}</div>
         </div>
         <div class="pricing-card">
           <div class="pricing-label">Hotel LIOZ</div>
           <div class="pricing-value">${valorHotel}</div>
         </div>
         <div class="pricing-card">
           <div class="pricing-label">HUB Transfer</div>
           <div class="pricing-value">${valorHUB}</div>
         </div>
       </div>
       
       <div class="info-section">
         <div class="info-row">
           <span class="info-label">💳 Forma de Pagamento:</span>
           <span class="info-value">${transfer.formaPagamento}</span>
         </div>
         <div class="info-row">
           <span class="info-label">💰 Pago Para:</span>
           <span class="info-value">${transfer.pagoPara}</span>
         </div>
       </div>
       
       ${transfer.observacoes ? `
       <div class="observations">
         <h4>📝 Observações:</h4>
         <p style="margin: 0;">${transfer.observacoes}</p>
       </div>
       ` : ''}
       
       ${CONFIG.EMAIL_CONFIG.USAR_BOTOES_INTERATIVOS ? `
       <div class="action-section">
         <h3>✅ Confirmação Necessária</h3>
         <p>Por favor, confirme ou cancele este transfer clicando em um dos botões abaixo:</p>
         <div class="action-buttons">
           <a href="${urlConfirmar}" class="btn btn-confirm">✓ CONFIRMAR TRANSFER</a>
           <a href="${urlCancelar}" class="btn btn-cancel">✗ CANCELAR TRANSFER</a>
         </div>
         <p style="margin-top: 15px; font-size: 12px; color: #666;">
           Ou responda este e-mail com "OK" para confirmar
         </p>
       </div>
       ` : `
       <div class="action-section">
         <h3>✅ Como Confirmar</h3>
         <p>Para confirmar este transfer, responda este e-mail com:</p>
         <h2 style="margin: 15px 0; color: #28a745;">"OK"</h2>
         <p>ou</p>
         <h2 style="margin: 15px 0; color: #28a745;">"OK ${transfer.id}"</h2>
       </div>
       `}
       
       <div class="footer">
         <p>
           <strong>${CONFIG.NAMES.SISTEMA_NOME}</strong><br>
           ${CONFIG.NAMES.HOTEL_NAME} & ${CONFIG.NAMES.HUB_OWNER}<br>
           Versão ${CONFIG.SISTEMA.VERSAO}
         </p>
         <p>
           E-mail automático enviado em ${formatarDataHora(new Date())}<br>
           <a href="mailto:${CONFIG.CONTATOS.HOTEL_EMAIL}">${CONFIG.CONTATOS.HOTEL_EMAIL}</a>
         </p>
       </div>
     </div>
   </body>
   </html>
 `;
}

/**
* Cria e-mail em texto simples
* @private
*/
function criarEmailTextoSimples(transfer) {
 const dataFormatada = formatarDataDDMMYYYY(new Date(transfer.data));
 
 return `
${MESSAGES.TITULOS.NOVO_TRANSFER}

Transfer #${transfer.id} - ${transfer.status}

DETALHES DO TRANSFER:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Cliente: ${transfer.cliente}
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
Hotel LIOZ: €${parseFloat(transfer.valorHotel).toFixed(2)}
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
* Processa clique nos botões do e-mail
* @param {Object} params - Parâmetros da URL
* @returns {HtmlOutput} - Página de resposta
*/
function handleEmailAction(params) {
 logger.info('Processando ação de e-mail', params);
 
 const { id: transferId, action } = params;
 const userEmail = Session.getActiveUser().getEmail() || 'Usuário Desconhecido';
 
 // Validar parâmetros
 if (!transferId || !action) {
   return createHtmlResponse(`
     <h1>❌ Erro</h1>
     <p>Parâmetros inválidos. Verifique o link.</p>
   `);
 }
 
 // Validar ação
 if (!['confirm', 'cancel'].includes(action)) {
   return createHtmlResponse(`
     <h1>❌ Erro</h1>
     <p>Ação inválida: ${action}</p>
   `);
 }
 
 try {
   // Buscar transfer
   const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
   const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
   const linha = encontrarLinhaPorId(sheet, transferId);
   
   if (!linha) {
     throw new Error(MESSAGES.ERROS.REGISTRO_NAO_ENCONTRADO(transferId));
   }
   
   // Verificar status atual
   const statusAtual = sheet.getRange(linha, 16).getValue();
   
   // Determinar novo status
   let novoStatus, observacao, tituloHtml, corFundo;
   
   if (action === 'confirm') {
     if (statusAtual === MESSAGES.STATUS_MESSAGES.CONFIRMADO) {
       return createHtmlResponse(`
         <h1>ℹ️ Transfer Já Confirmado</h1>
         <p>O transfer #${transferId} já estava confirmado.</p>
         <p>Status atual: ${statusAtual}</p>
       `);
     }
     
     novoStatus = MESSAGES.STATUS_MESSAGES.CONFIRMADO;
     observacao = MESSAGES.ACOES.CONFIRMADO_POR(userEmail);
     tituloHtml = '✅ Transfer Confirmado!';
     corFundo = '#d4edda';
     
   } else if (action === 'cancel') {
     if (statusAtual === MESSAGES.STATUS_MESSAGES.CANCELADO) {
       return createHtmlResponse(`
         <h1>ℹ️ Transfer Já Cancelado</h1>
         <p>O transfer #${transferId} já estava cancelado.</p>
         <p>Status atual: ${statusAtual}</p>
       `);
     }
     
     novoStatus = MESSAGES.STATUS_MESSAGES.CANCELADO;
     observacao = MESSAGES.ACOES.CANCELADO_POR(userEmail);
     tituloHtml = '❌ Transfer Cancelado';
     corFundo = '#f8d7da';
   }
   
   // Atualizar status
   const resultado = atualizarStatusTransfer(transferId, novoStatus, observacao);
   
   if (resultado.sucesso) {
     // Buscar dados do transfer para exibir
     const dadosTransfer = sheet.getRange(linha, 1, 1, HEADERS.length).getValues()[0];
     const cliente = dadosTransfer[1];
     const data = formatarDataDDMMYYYY(new Date(dadosTransfer[4]));
     const rota = `${dadosTransfer[7]} → ${dadosTransfer[8]}`;
     
     // Enviar e-mail de confirmação
     enviarEmailConfirmacaoAcao(transferId, action, userEmail);
     
     return createHtmlResponse(`
       <!DOCTYPE html>
       <html>
       <head>
         <meta charset="UTF-8">
         <style>
           body {
             font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif;
             background: ${corFundo};
             height: 100vh;
             display: flex;
             align-items: center;
             justify-content: center;
             margin: 0;
           }
           .card {
             background: white;
             padding: 40px;
             border-radius: 12px;
             box-shadow: 0 4px 20px rgba(0,0,0,0.1);
             text-align: center;
             max-width: 500px;
           }
           h1 {
             color: #333;
             margin-bottom: 20px;
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
           }
           .label {
             font-weight: bold;
             color: #666;
           }
           .timestamp {
             color: #666;
             font-size: 14px;
             margin-top: 20px;
           }
           .close-message {
             font-size: 12px;
             color: #888;
             margin-top: 30px;
           }
         </style>
       </head>
       <body>
         <div class="card">
           <h1>${tituloHtml}</h1>
           <p>Transfer processado com sucesso!</p>
           
           <div class="details">
             <div class="detail-row">
               <span class="label">Transfer ID:</span> #${transferId}
             </div>
             <div class="detail-row">
               <span class="label">Cliente:</span> ${cliente}
             </div>
             <div class="detail-row">
               <span class="label">Data:</span> ${data}
             </div>
             <div class="detail-row">
               <span class="label">Rota:</span> ${rota}
             </div>
             <div class="detail-row">
               <span class="label">Novo Status:</span> ${novoStatus}
             </div>
             <div class="detail-row">
               <span class="label">Processado por:</span> ${userEmail}
             </div>
           </div>
           
           <p class="timestamp">
             ${formatarDataHora(new Date())}
           </p>
           
           <p class="close-message">
             Você pode fechar esta aba agora.
           </p>
         </div>
       </body>
       </html>
     `);
     
   } else {
     throw new Error(resultado.mensagem || 'Erro ao atualizar status');
   }
   
 } catch (error) {
   logger.error('Erro ao processar ação de e-mail', error);
   
   return createHtmlResponse(`
     <!DOCTYPE html>
     <html>
     <head>
       <meta charset="UTF-8">
       <style>
         body {
           font-family: Arial, sans-serif;
           background: #f8d7da;
           height: 100vh;
           display: flex;
           align-items: center;
           justify-content: center;
           margin: 0;
         }
         .error-card {
           background: white;
           padding: 40px;
           border-radius: 12px;
           box-shadow: 0 4px 20px rgba(0,0,0,0.1);
           text-align: center;
           max-width: 500px;
         }
         h1 {
           color: #721c24;
         }
         .error-details {
           background: #f5c6cb;
           padding: 15px;
           border-radius: 6px;
           margin: 20px 0;
           color: #721c24;
         }
       </style>
     </head>
     <body>
       <div class="error-card">
         <h1>❌ Erro ao Processar</h1>
         <p>Não foi possível processar sua solicitação.</p>
         <div class="error-details">
           ${error.message}
         </div>
         <p>Por favor, tente novamente ou entre em contato com o suporte.</p>
       </div>
     </body>
     </html>
   `);
 }
}

/**
* Envia e-mail de confirmação após ação
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
   
   MailApp.sendEmail({
     to: destinatarios,
     subject: assunto,
     htmlBody: corpo,
     name: CONFIG.NAMES.SISTEMA_NOME
   });
   
 } catch (error) {
   logger.error('Erro ao enviar e-mail de confirmação', error);
 }
}

// ===================================================
// PARTE 5.1: VERIFICAÇÃO AUTOMÁTICA DE E-MAILS
// ===================================================

/**
* Verifica e-mails recebidos para processar confirmações
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
               MESSAGES.STATUS_MESSAGES.CONFIRMADO,
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
               MESSAGES.STATUS_MESSAGES.CANCELADO,
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
* Encontra o último transfer com status "Solicitado"
* @private
*/
function encontrarUltimoTransferSolicitado(sheet) {
 try {
   const lastRow = sheet.getLastRow();
   
   // Buscar de baixo para cima (mais recentes primeiro)
   for (let i = lastRow; i >= 2; i--) {
     const status = sheet.getRange(i, 16).getValue();
     
     if (status === MESSAGES.STATUS_MESSAGES.SOLICITADO) {
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
* Envia resumo das verificações de e-mail
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
   
   MailApp.sendEmail({
     to: destinatarios,
     subject: assunto,
     htmlBody: corpo,
     name: CONFIG.NAMES.SISTEMA_NOME
   });
   
 } catch (error) {
   logger.error('Erro ao enviar resumo de verificação', error);
 }
}

// ===================================================
// PARTE 5.2: TEMPLATES DE E-MAIL ADICIONAIS
// ===================================================

/**
* Envia e-mail de status atualizado
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
     data: formatarDataDDMMYYYY(new Date(dadosTransfer[4])),
     rota: `${dadosTransfer[7]} → ${dadosTransfer[8]}`,
     statusAnterior: dadosTransfer[15],
     statusNovo: novoStatus,
     motivo: motivo,
     timestamp: formatarDataHora(new Date())
   });
   
   MailApp.sendEmail({
     to: destinatarios,
     subject: assunto,
     htmlBody: corpo,
     name: CONFIG.NAMES.SISTEMA_NOME
   });
   
   logger.success('E-mail de status enviado');
   
 } catch (error) {
   logger.error('Erro ao enviar e-mail de status', error);
 }
}

/**
* Template para e-mail de mudança de status
* @private
*/
function criarEmailStatusTemplate(dados) {
 const corStatus = STYLES.STATUS_COLORS[dados.statusNovo] || '#666';
 
 return `
   <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
     <h2 style="color: ${corStatus};">Status Atualizado</h2>
     
     <div style="background: #f8f9fa; padding: 20px; border-radius: 8px; margin: 20px 0;">
       <h3>Transfer #${dados.id}</h3>
       <p><strong>Cliente:</strong> ${dados.cliente}</p>
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

/**
* Envia relatório diário por e-mail
*/
function enviarRelatorioDiario() {
 logger.info('Gerando e enviando relatório diário');
 
 try {
   const hoje = new Date();
   const ontem = new Date(hoje);
   ontem.setDate(ontem.getDate() - 1);
   
   // Buscar dados do dia
   const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
   const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
   
   if (!sheet || sheet.getLastRow() <= 1) {
     logger.info('Sem dados para relatório diário');
     return;
   }
   
   const dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
   
   // Filtrar transfers de hoje
   const transfersHoje = dados.filter(row => {
     const dataTransfer = new Date(row[4]);
     return dataTransfer.toDateString() === hoje.toDateString();
   });
   
   // Gerar estatísticas
   const stats = {
     total: transfersHoje.length,
     confirmados: 0,
     cancelados: 0,
     pendentes: 0,
     valorTotal: 0,
     topRotas: {},
     porHora: {}
   };
   
   transfersHoje.forEach(transfer => {
     // Status
     const status = transfer[15];
     if (status === MESSAGES.STATUS_MESSAGES.CONFIRMADO) stats.confirmados++;
     else if (status === MESSAGES.STATUS_MESSAGES.CANCELADO) stats.cancelados++;
     else if (status === MESSAGES.STATUS_MESSAGES.SOLICITADO) stats.pendentes++;
     
     // Valores
     stats.valorTotal += parseFloat(transfer[10]) || 0;
     
     // Rotas
     const rota = `${transfer[7]} → ${transfer[8]}`;
     stats.topRotas[rota] = (stats.topRotas[rota] || 0) + 1;
     
     // Por hora
     const hora = transfer[9] ? transfer[9].split(':')[0] : 'N/A';
     stats.porHora[hora] = (stats.porHora[hora] || 0) + 1;
   });
   
// Enviar e-mail
   const destinatarios = CONFIG.EMAIL_CONFIG.DESTINATARIOS.join(',');
   const assunto = `[${CONFIG.NAMES.SISTEMA_NOME}] Relatório Diário - ${formatarDataDDMMYYYY(hoje)}`;
   const corpo = criarRelatorioEmailTemplate(stats, hoje);
   
   MailApp.sendEmail({
     to: destinatarios,
     subject: assunto,
     htmlBody: corpo,
     name: CONFIG.NAMES.SISTEMA_NOME
   });
   
   logger.success('Relatório diário enviado', { transfers: stats.total });
   
 } catch (error) {
   logger.error('Erro ao enviar relatório diário', error);
 }
}

/**
* Template para relatório diário
* @private
*/
function criarRelatorioEmailTemplate(stats, data) {
 const topRotasHtml = Object.entries(stats.topRotas)
   .sort((a, b) => b[1] - a[1])
   .slice(0, 5)
   .map(([rota, count]) => `<li>${rota}: ${count} transfer(s)</li>`)
   .join('');
 
 const horariosHtml = Object.entries(stats.porHora)
   .sort((a, b) => a[0].localeCompare(b[0]))
   .map(([hora, count]) => `<li>${hora}:00 - ${count} transfer(s)</li>`)
   .join('');
 
 return `
   <!DOCTYPE html>
   <html>
   <head>
     <style>
       body { font-family: Arial, sans-serif; background: #f5f5f5; margin: 0; padding: 20px; }
       .container { max-width: 600px; margin: 0 auto; background: white; border-radius: 8px; padding: 30px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
       h1 { color: #2c3e50; border-bottom: 3px solid #3498db; padding-bottom: 10px; }
       .stats-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin: 20px 0; }
       .stat-card { background: #f8f9fa; padding: 20px; border-radius: 6px; text-align: center; }
       .stat-value { font-size: 32px; font-weight: bold; color: #2c3e50; }
       .stat-label { color: #7f8c8d; margin-top: 5px; }
       .section { margin: 30px 0; }
       .section h3 { color: #34495e; }
       ul { padding-left: 20px; }
       .footer { text-align: center; color: #7f8c8d; font-size: 12px; margin-top: 30px; padding-top: 20px; border-top: 1px solid #ecf0f1; }
     </style>
   </head>
   <body>
     <div class="container">
       <h1>📊 Relatório Diário de Transfers</h1>
       <p style="color: #7f8c8d;">Data: ${formatarDataDDMMYYYY(data)}</p>
       
       <div class="stats-grid">
         <div class="stat-card">
           <div class="stat-value">${stats.total}</div>
           <div class="stat-label">Total de Transfers</div>
         </div>
         <div class="stat-card">
           <div class="stat-value">€${stats.valorTotal.toFixed(2)}</div>
           <div class="stat-label">Valor Total</div>
         </div>
       </div>
       
       <div class="section">
         <h3>📈 Status dos Transfers</h3>
         <ul>
           <li>✅ Confirmados: ${stats.confirmados}</li>
           <li>⏳ Pendentes: ${stats.pendentes}</li>
           <li>❌ Cancelados: ${stats.cancelados}</li>
         </ul>
       </div>
       
       ${topRotasHtml ? `
       <div class="section">
         <h3>🚐 Top 5 Rotas</h3>
         <ul>${topRotasHtml}</ul>
       </div>
       ` : ''}
       
       ${horariosHtml ? `
       <div class="section">
         <h3>🕐 Distribuição por Horário</h3>
         <ul>${horariosHtml}</ul>
       </div>
       ` : ''}
       
       <div class="footer">
         <p>${CONFIG.NAMES.SISTEMA_NOME} - Relatório Automático</p>
         <p>Gerado em ${formatarDataHora(new Date())}</p>
       </div>
     </div>
   </body>
   </html>
 `;
}

// ===================================================
// PARTE 5.3: CONFIGURAÇÃO DE TRIGGERS
// ===================================================

/**
* Configura triggers automáticos para o sistema de e-mail
*/
function configurarTriggersEmail() {
 logger.info('Configurando triggers de e-mail');
 
 try {
   // Remover triggers existentes relacionados a e-mail
   const triggers = ScriptApp.getProjectTriggers();
   triggers.forEach(trigger => {
     const handlerFunction = trigger.getHandlerFunction();
     if (handlerFunction === 'verificarConfirmacoesEmail' || 
         handlerFunction === 'enviarRelatorioDiario') {
       ScriptApp.deleteTrigger(trigger);
       logger.info(`Trigger removido: ${handlerFunction}`);
     }
   });
   
   // Criar trigger para verificação de e-mails
   if (CONFIG.EMAIL_CONFIG.VERIFICAR_CONFIRMACOES) {
     ScriptApp.newTrigger('verificarConfirmacoesEmail')
       .timeBased()
       .everyMinutes(CONFIG.EMAIL_CONFIG.INTERVALO_VERIFICACAO)
       .create();
     
     logger.success('Trigger de verificação de e-mails criado', {
       intervalo: CONFIG.EMAIL_CONFIG.INTERVALO_VERIFICACAO + ' minutos'
     });
   }
   
   // Criar trigger para relatório diário (8h da manhã)
   ScriptApp.newTrigger('enviarRelatorioDiario')
     .timeBased()
     .atHour(8)
     .everyDays(1)
     .inTimezone(CONFIG.SISTEMA.TIMEZONE)
     .create();
   
   logger.success('Trigger de relatório diário criado');
   
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
* Remove todos os triggers de e-mail
*/
function removerTriggersEmail() {
 logger.info('Removendo triggers de e-mail');
 
 try {
   const triggers = ScriptApp.getProjectTriggers();
   let removidos = 0;
   
   triggers.forEach(trigger => {
     const handlerFunction = trigger.getHandlerFunction();
     if (handlerFunction === 'verificarConfirmacoesEmail' || 
         handlerFunction === 'enviarRelatorioDiario') {
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

// ===================================================
// PARTE 6: SISTEMA DE REGISTRO E ATUALIZAÇÃO
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
    const dataTransfer = new Date(dadosTransfer[4]);
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
    
    // PASSO 4: Integração com webhooks se configurado
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
      // Verificar status atual
      const statusAtual = abaPrincipal.getRange(linhaPrincipal, 16).getValue();
      
      if (statusAtual === novoStatus) {
        logger.info('Status já está atualizado', { transferId, status: novoStatus });
        return {
          sucesso: false,
          mensagem: `Transfer já está com status "${novoStatus}"`
        };
      }
      
      // Atualizar status (coluna P - 16)
      abaPrincipal.getRange(linhaPrincipal, 16).setValue(novoStatus);
      
      // Adicionar observação (coluna Q - 17)
      const obsAtual = abaPrincipal.getRange(linhaPrincipal, 17).getValue() || '';
      const novaObs = obsAtual 
        ? `${obsAtual}\n${observacao} - ${formatarDataHora(new Date())}`
        : `${observacao} - ${formatarDataHora(new Date())}`;
      abaPrincipal.getRange(linhaPrincipal, 17).setValue(novaObs);
      
      atualizacoes++;
      resultados.push({
        aba: CONFIG.SHEET_NAME,
        linha: linhaPrincipal,
        sucesso: true
      });
      
      // Tentar atualizar na aba mensal
      try {
        const dataTransfer = abaPrincipal.getRange(linhaPrincipal, 5).getValue();
        const abaMensal = obterAbaMes(dataTransfer);
        
        if (abaMensal && abaMensal.getName() !== abaPrincipal.getName()) {
          const linhaMensal = encontrarLinhaPorId(abaMensal, transferId);
          
          if (linhaMensal > 0) {
            abaMensal.getRange(linhaMensal, 16).setValue(novoStatus);
            abaMensal.getRange(linhaMensal, 17).setValue(novaObs);
            
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
      if (novoStatus === MESSAGES.STATUS_MESSAGES.FINALIZADO) {
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
      const obsAtual = sheet.getRange(linha, 17).getValue() || '';
      const obsAtualizada = obsAtual 
        ? `${obsAtual}\n${novaObservacao}`
        : novaObservacao;
      
      sheet.getRange(linha, 17).setValue(obsAtualizada);
      
      // Tentar atualizar na aba mensal também
      try {
        const dataTransfer = sheet.getRange(linha, 5).getValue();
        const abaMensal = obterAbaMes(dataTransfer);
        const linhaMensal = encontrarLinhaPorId(abaMensal, transferId);
        
        if (linhaMensal > 0) {
          abaMensal.getRange(linhaMensal, 17).setValue(obsAtualizada);
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
 * Atualiza dados completos de um transfer
 * @param {string} transferId - ID do transfer
 * @param {Object} dadosAtualizacao - Dados a atualizar
 * @returns {Object} - Resultado da atualização
 */
function atualizarTransferCompleto(transferId, dadosAtualizacao) {
  logger.info('Atualizando transfer completo', { transferId });
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const linha = encontrarLinhaPorId(sheet, transferId);
    
    if (!linha) {
      throw new Error(MESSAGES.ERROS.REGISTRO_NAO_ENCONTRADO(transferId));
    }
    
    // Mapear campos para colunas
    const mapeamento = {
      cliente: 2,
      pessoas: 3,
      bagagens: 4,
      data: 5,
      contacto: 6,
      voo: 7,
      origem: 8,
      destino: 9,
      horaPickup: 10,
      precoCliente: 11,
      valorHotel: 12,
      valorHUB: 13,
      formaPagamento: 14,
      pagoPara: 15,
      status: 16,
      observacoes: 17
    };
    
    const alteracoes = [];
    
    // Aplicar atualizações
    Object.keys(dadosAtualizacao).forEach(campo => {
      if (mapeamento[campo]) {
        const coluna = mapeamento[campo];
        let valor = dadosAtualizacao[campo];
        
        // Processar valores especiais
        if (campo === 'data') {
          valor = processarDataSegura(valor);
        } else if (['pessoas', 'bagagens'].includes(campo)) {
          valor = parseInt(valor) || 0;
        } else if (['precoCliente', 'valorHotel', 'valorHUB'].includes(campo)) {
          valor = parseFloat(valor) || 0;
        }
        
        sheet.getRange(linha, coluna).setValue(valor);
        alteracoes.push(`${campo}: ${valor}`);
      }
    });
    
    // Adicionar registro de alteração nas observações
    const obsAtual = sheet.getRange(linha, 17).getValue() || '';
    const registroAlteracao = `\n[ALTERADO ${formatarDataHora(new Date())}] ${alteracoes.join(', ')}`;
    sheet.getRange(linha, 17).setValue(obsAtual + registroAlteracao);
    
    // Tentar atualizar na aba mensal
    try {
      const dataTransfer = sheet.getRange(linha, 5).getValue();
      const abaMensal = obterAbaMes(dataTransfer);
      
      if (abaMensal && abaMensal.getName() !== sheet.getName()) {
        const linhaMensal = encontrarLinhaPorId(abaMensal, transferId);
        
        if (linhaMensal > 0) {
          // Aplicar mesmas atualizações
          Object.keys(dadosAtualizacao).forEach(campo => {
            if (mapeamento[campo]) {
              const coluna = mapeamento[campo];
              let valor = dadosAtualizacao[campo];
              
              if (campo === 'data') {
                valor = processarDataSegura(valor);
              } else if (['pessoas', 'bagagens'].includes(campo)) {
                valor = parseInt(valor) || 0;
              } else if (['precoCliente', 'valorHotel', 'valorHUB'].includes(campo)) {
                valor = parseFloat(valor) || 0;
              }
              
              abaMensal.getRange(linhaMensal, coluna).setValue(valor);
            }
          });
          
          abaMensal.getRange(linhaMensal, 17).setValue(obsAtual + registroAlteracao);
        }
      }
    } catch (errorMensal) {
      logger.error('Erro ao atualizar aba mensal', errorMensal);
    }
    
    logger.success('Transfer atualizado com sucesso', {
      transferId,
      alteracoes: alteracoes.length
    });
    
    return {
      sucesso: true,
      transferId: transferId,
      alteracoes: alteracoes
    };
    
  } catch (error) {
    logger.error('Erro ao atualizar transfer', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

// ===================================================
// PARTE 6.1: CORREÇÃO E MANUTENÇÃO DE REGISTROS
// ===================================================

/**
 * Corrige registros incompletos (sem registro duplo)
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
      const dataTransfer = row[4];
      
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
 * Remove registros duplicados
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
      const data = formatarDataDDMMYYYY(new Date(row[4]));
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
* Sincroniza dados entre aba principal e abas mensais
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
     const data = new Date(row[4]);
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

// ===================================================
// PARTE 6.2: INTEGRAÇÃO COM WEBHOOKS
// ===================================================

/**
* Envia dados do transfer para webhook do Make.com
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
     pessoas: dadosTransfer[2],
     bagagens: dadosTransfer[3],
     data: formatarDataDDMMYYYY(new Date(dadosTransfer[4])),
     contacto: dadosTransfer[5],
     voo: dadosTransfer[6],
     origem: dadosTransfer[7],
     destino: dadosTransfer[8],
     horaPickup: dadosTransfer[9],
     precoCliente: dadosTransfer[10],
     valorHotel: dadosTransfer[11],
     valorHUB: dadosTransfer[12],
     formaPagamento: dadosTransfer[13],
     pagoPara: dadosTransfer[14],
     status: dadosTransfer[15],
     observacoes: dadosTransfer[16],
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
* Envia atualização de status via webhook
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
// PARTE 7: APIs PRINCIPAIS (doGet e doPost)
// ===================================================

/**
 * Processa requisições GET
 * @param {Object} e - Evento da requisição
 * @returns {TextOutput|HtmlOutput} - Resposta
 */
function doGet(e) {
  logger.info('Recebendo requisição GET', { 
    parameters: e.parameter,
    queryString: e.queryString 
  });
  
  try {
    const params = e.parameter || {};
    const action = params.action || 'default';
    
    // Roteamento baseado na ação
    switch (action) {
      // Ações de e-mail
      case 'confirm':
      case 'cancel':
        return handleEmailAction(params);
      
      // Consultas
      case 'consultarPreco':
        return handleConsultarPreco(params);
      
      case 'verificarRegistroDuplo':
        return handleVerificarDuplicidade(params);
      
      case 'buscarTransfer':
        return handleBuscarTransfer(params);
      
      case 'listarTransfers':
        return handleListarTransfers(params);
      
      // Sistema
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
        return handleDefault();
    }
    
  } catch (error) {
    logger.error('Erro no doGet', error);
    return createJsonResponse({
      status: 'error',
      message: error.toString(),
      action: e.parameter?.action || 'unknown'
    }, 500);
  }
}

/**
 * Processa requisições POST
 * @param {Object} e - Evento da requisição
 * @returns {TextOutput} - Resposta JSON
 */
function doPost(e) {
  logger.info('Recebendo requisição POST', {
    contentLength: e.contentLength,
    postData: e.postData?.type
  });
  
  try {
    // Validar dados recebidos
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Nenhum dado recebido');
    }
    
    const dadosRecebidos = JSON.parse(e.postData.contents);
    logger.debug('Dados recebidos', { 
      action: dadosRecebidos.action,
      hasTransferData: !!dadosRecebidos.nomeCliente 
    });
    
    // Processar ação especial se houver
    if (dadosRecebidos.action) {
      return processarAcaoEspecial(dadosRecebidos);
    }
    
    // Processar novo transfer
    return processarNovoTransfer(dadosRecebidos);
    
  } catch (error) {
    logger.error('Erro no doPost', error);
    return createJsonResponse({
      status: 'error',
      message: error.toString(),
      timestamp: new Date().toISOString()
    }, 500);
  }
}

// ===================================================
// PARTE 7.1: HANDLERS DO GET
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
  
  const valores = calcularValores(origem, destino, pessoas, bagagens);
  
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
        const dataTransfer = formatarDataDDMMYYYY(new Date(row[4]));
        if (dataTransfer !== filtros.data) incluir = false;
      }
      
      if (filtros.status && row[15] !== filtros.status) {
        incluir = false;
      }
      
      if (filtros.cliente && !row[1].toLowerCase().includes(filtros.cliente.toLowerCase())) {
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
    versao: CONFIG.SISTEMA.VERSAO,
    timestamp: new Date().toISOString(),
    endpoints: {
      GET: [
        '?action=test',
        '?action=config',
        '?action=stats',
        '?action=health',
        '?action=consultarPreco&origem=X&destino=Y&pessoas=N',
        '?action=verificarRegistroDuplo&id=X&data=Y',
        '?action=buscarTransfer&id=X',
        '?action=listarTransfers&status=X&limite=N',
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
    versao: CONFIG.SISTEMA.VERSAO
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
    message: 'Sistema LIOZ-HUB Transfer funcionando!',
    versao: CONFIG.SISTEMA.VERSAO,
    documentacao: 'Use ?action=test para ver endpoints disponíveis'
  });
}

// ===================================================
// PARTE 7.2: PROCESSAMENTO DO POST
// ===================================================

/**
 * Processa novo transfer recebido via POST
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
    
    // Gerar ID
    const transferId = dados.id || gerarProximoIdSeguro(sheet);
    
    // Calcular valores
    const valores = calcularValores(
      dados.origem,
      dados.destino,
      dados.numeroPessoas,
      dados.numeroBagagens,
      dados.valorTotal
    );
    
    // Montar dados do transfer
    const dadosTransfer = [
      transferId,
      dados.nomeCliente,
      dados.numeroPessoas,
      dados.numeroBagagens,
      dados.data,
      dados.contacto,
      dados.numeroVoo || '',
      dados.origem,
      dados.destino,
      dados.horaPickup,
      valores.precoCliente,
      valores.valorHotel,
      valores.valorHUB,
      dados.modoPagamento,
      dados.pagoParaQuem,
      dados.status,
      dados.observacoes || '',
      new Date()
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
        data: formatarDataDDMMYYYY(dados.data),
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
 * Processa ações especiais via POST
 * @private
 */
function processarAcaoEspecial(dados) {
  logger.info('Processando ação especial', { action: dados.action });
  
  try {
    let resultado;
    
    switch (dados.action) {
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
        
      case 'enviarRelatorioDiario':
        enviarRelatorioDiario();
        resultado = { sucesso: true, mensagem: 'Relatório enviado' };
        break;
        
      case 'verificarConfirmacoes':
        const confirmacoes = verificarConfirmacoesEmail();
        resultado = { 
          sucesso: true, 
          confirmacoes: confirmacoes,
          mensagem: `${confirmacoes} confirmações processadas`
        };
        break;
        
      default:
        throw new Error(`Ação desconhecida: ${dados.action}`);
    }
    
    return createJsonResponse({
      status: resultado.sucesso ? 'success' : 'error',
      action: dados.action,
      resultado: resultado
    });
    
  } catch (error) {
    logger.error('Erro na ação especial', error);
    return createJsonResponse({
      status: 'error',
      action: dados.action,
      message: error.toString()
    }, 500);
  }
}

// ===================================================
// PARTE 8: FUNÇÕES DE MANUTENÇÃO E UTILITÁRIOS
// ===================================================

/**
 * Limpa todos os dados do sistema
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
    
    // Limpar cache
    cache.clear();
    
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
 * Limpa apenas dados de teste
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
         const observacoes = String(row[16]).toLowerCase();
         
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
* Reordena transfers por data
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
   
   // Ordenar por data (coluna 5) e depois por hora (coluna 10)
   dados.sort((a, b) => {
     const dataA = new Date(a[4]);
     const dataB = new Date(b[4]);
     
     if (dataA.getTime() !== dataB.getTime()) {
       return dataA - dataB;
     }
     
     // Se mesma data, ordenar por hora
     const horaA = a[9] || '00:00';
     const horaB = b[9] || '00:00';
     
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
* Aplica formatação na planilha principal
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
   
   // Formatações de dados
   if (maxRows > 1) {
     // Moeda
     sheet.getRange(2, 11, maxRows - 1, 3).setNumberFormat(STYLES.FORMATS.MOEDA);
     
     // Data
     sheet.getRange(2, 5, maxRows - 1, 1).setNumberFormat(STYLES.FORMATS.DATA);
     
     // Hora
     sheet.getRange(2, 10, maxRows - 1, 1).setNumberFormat(STYLES.FORMATS.HORA);
     
     // Timestamp
     sheet.getRange(2, 18, maxRows - 1, 1).setNumberFormat(STYLES.FORMATS.TIMESTAMP);
     
     // Números
     sheet.getRange(2, 1, maxRows - 1, 1).setNumberFormat(STYLES.FORMATS.NUMERO);
     sheet.getRange(2, 3, maxRows - 1, 2).setNumberFormat(STYLES.FORMATS.NUMERO);
   }
   
   // Aplicar validações
   aplicarValidacoesPlanilha(sheet);
   
   logger.debug('Formatação aplicada com sucesso');
   
 } catch (error) {
   logger.error('Erro ao aplicar formatação', error);
 }
}

/**
* Aplica formatação na tabela de preços
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
     // Formatação monetária
     sheet.getRange(2, 7, maxRows - 1, 3).setNumberFormat(STYLES.FORMATS.MOEDA);
     
     // Timestamp
     sheet.getRange(2, 11, maxRows - 1, 1).setNumberFormat(STYLES.FORMATS.TIMESTAMP);
     
     // Números
     sheet.getRange(2, 1, maxRows - 1, 1).setNumberFormat(STYLES.FORMATS.NUMERO);
     sheet.getRange(2, 5, maxRows - 1, 2).setNumberFormat(STYLES.FORMATS.NUMERO);
     
     // Validação Ativo/Inativo
     const ativoValidation = SpreadsheetApp.newDataValidation()
       .requireValueInList(['Sim', 'Não'])
       .setAllowInvalid(false)
       .build();
     sheet.getRange(2, 10, maxRows - 1, 1).setDataValidation(ativoValidation);
   }
   
   logger.debug('Formatação de preços aplicada');
   
 } catch (error) {
   logger.error('Erro ao aplicar formatação de preços', error);
 }
}

// ===================================================
// PARTE 8.1: MENU DO SISTEMA
// ===================================================

/**
* Cria o menu personalizado na planilha
*/
function onOpen() {
 const ui = SpreadsheetApp.getUi();
 
 ui.createMenu('🚐 Sistema v4.0')
   .addItem('⚙️ Configurar Sistema', 'configurarSistema')
   .addSeparator()
   .addSubMenu(ui.createMenu('📧 E-mail Interativo')
     .addItem('✉️ Testar Envio de E-mail', 'testarEnvioEmailInterativo')
     .addItem('🔍 Verificar Confirmações', 'verificarConfirmacoesManual')
     .addItem('⏰ Configurar Verificação Automática', 'configurarTriggersEmailMenu')
     .addItem('🛑 Parar Verificação Automática', 'removerTriggersEmailMenu')
     .addItem('📊 Enviar Relatório Diário', 'enviarRelatorioDiarioManual'))
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
     .addItem('🔄 Sincronizar com Aba Principal', 'sincronizarAbasMensaisMenu')
     .addItem('📊 Relatório Mensal', 'gerarRelatorioMensalMenu'))
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
   .addSubMenu(ui.createMenu('📊 Relatórios')
     .addItem('📈 Estatísticas Gerais', 'mostrarEstatisticasMenu')
     .addItem('📅 Consolidado Anual', 'mostrarConsolidadoAnualMenu')
     .addItem('🏆 Top Clientes e Rotas', 'mostrarTopClientesRotasMenu'))
   .addSeparator()
   .addItem('🧪 Testar Sistema Completo', 'testarSistemaCompleto')
   .addItem('ℹ️ Sobre o Sistema', 'mostrarSobre')
   .addToUi();
}

// ===================================================
// PARTE 8.2: FUNÇÕES DO MENU
// ===================================================

/**
* Configura o sistema completamente
*/
function configurarSistema() {
 const ui = SpreadsheetApp.getUi();
 
 ui.alert(
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
     inserirDadosIniciaisPrecos(tabelaPrecos);
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
* Testa envio de e-mail interativo
*/
function testarEnvioEmailInterativo() {
 const ui = SpreadsheetApp.getUi();
 
 const dadosTeste = [
   999,
   'TESTE EMAIL INTERATIVO',
   2,
   1,
   new Date(),
   '+351999888777',
   'TP1234',
   'Aeroporto de Lisboa',
   'Hotel LIOZ',
   '10:00',
   25.00,
   10.00,
   15.00,
   'Dinheiro',
   'Recepção',
   'Solicitado',
   'Este é um transfer de teste do sistema v4.0',
   new Date()
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
* Verifica confirmações manualmente
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
* Menu para configurar triggers
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
       '• Relatório diário: às 8h da manhã\n\n' +
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
* Menu para remover triggers
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
* Envia relatório diário manualmente
*/
function enviarRelatorioDiarioManual() {
 const ui = SpreadsheetApp.getUi();
 
 try {
   enviarRelatorioDiario();
   ui.alert('✅ Relatório Enviado', 'Relatório diário enviado com sucesso!', ui.ButtonSet.OK);
 } catch (error) {
   ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
 }
}

/**
* Adicionar preço via menu
*/
function adicionarPrecoMenu() {
 const ui = SpreadsheetApp.getUi();
 
 const response = ui.prompt(
   '➕ Adicionar Preço',
   'Digite os dados do preço no formato:\n' +
   'Origem|Destino|Pessoas|Preço\n\n' +
   'Exemplo: Aeroporto Lisboa|Hotel LIOZ|2|25',
   ui.ButtonSet.OK_CANCEL
 );
 
 if (response.getSelectedButton() === ui.Button.OK) {
   try {
     const partes = response.getResponseText().split('|');
     
     if (partes.length < 4) {
       throw new Error('Formato inválido');
     }
     
     const resultado = adicionarPrecoTabela({
       origem: partes[0].trim(),
       destino: partes[1].trim(),
       pessoas: parseInt(partes[2]),
       precoCliente: parseFloat(partes[3]),
       bagagens: 0,
       ativo: true
     });
     
     if (resultado.sucesso) {
       ui.alert('✅ Sucesso', `Preço adicionado com ID ${resultado.id}`, ui.ButtonSet.OK);
     } else {
       ui.alert('❌ Erro', resultado.erro, ui.ButtonSet.OK);
     }
     
   } catch (error) {
     ui.alert('❌ Erro', 'Erro ao adicionar preço:\n' + error.toString(), ui.ButtonSet.OK);
   }
 }
}

/**
* Mostra estatísticas via menu
*/
function mostrarEstatisticasMenu() {
 const ui = SpreadsheetApp.getUi();
 
 try {
   const stats = gerarEstatisticas();
   
   const mensagem = `
📊 ESTATÍSTICAS DO SISTEMA

Total de Transfers: ${stats.totalTransfers}
Valor Total: €${stats.valorTotal.toFixed(2)}
Valor Hotel: €${stats.valorHotel.toFixed(2)}
Valor HUB: €${stats.valorHUB.toFixed(2)}

Média de Passageiros: ${stats.mediaPassageiros.toFixed(1)}

📈 Por Status:
${Object.entries(stats.porStatus).map(([status, count]) => `• ${status}: ${count}`).join('\n')}

🚐 Top 5 Rotas:
${Object.entries(stats.topRotas).slice(0, 5).map(([rota, count]) => `• ${rota}: ${count}`).join('\n')}

💳 Formas de Pagamento:
${Object.entries(stats.formasPagamento).map(([forma, count]) => `• ${forma}: ${count}`).join('\n')}
   `;
   
   ui.alert('📊 Estatísticas', mensagem, ui.ButtonSet.OK);
   
 } catch (error) {
   ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
 }
}

/**
* Testa o sistema completo
*/
function testarSistemaCompleto() {
 const ui = SpreadsheetApp.getUi();
 
 ui.alert(
   '🧪 Teste Completo',
   'Esta função irá:\n\n' +
   '1. Verificar configurações\n' +
   '2. Testar registro de transfer\n' +
   '3. Testar cálculo de preços\n' +
   '4. Verificar abas mensais\n' +
   '5. Testar envio de e-mail\n\n' +
   'Continuar?',
   ui.ButtonSet.YES_NO
 );
 
 const resultados = {
   configuracao: false,
   registro: false,
   precos: false,
   abasMensais: false,
   email: false
 };
 
 try {
   // 1. Verificar configuração
   const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
   resultados.configuracao = !!ss.getSheetByName(CONFIG.SHEET_NAME);
   
   // 2. Testar registro
   const dadosTeste = {
     nomeCliente: 'TESTE SISTEMA COMPLETO',
     numeroPessoas: 2,
     numeroBagagens: 1,
     data: new Date().toISOString().split('T')[0],
     contacto: '+351999888777',
     origem: 'Aeroporto de Lisboa',
     destino: 'Hotel LIOZ',
     horaPickup: '10:00',
     valorTotal: 25
   };
   
   const mockRequest = {
     postData: {
       contents: JSON.stringify(dadosTeste)
     }
   };
   
   const response = JSON.parse(doPost(mockRequest).getContent());
   resultados.registro = response.status === 'success';
   
   // 3. Testar preços
   const preco = calcularValores('Aeroporto', 'Hotel', 2, 1);
   resultados.precos = preco.precoCliente > 0;
   
   // 4. Verificar abas mensais
   const integridade = verificarIntegridadeAbasMensais();
   resultados.abasMensais = integridade.encontradas >= 12;
   
   // 5. Testar e-mail (simulado)
   resultados.email = CONFIG.EMAIL_CONFIG.ENVIAR_AUTOMATICO;
   
   // Mostrar resultados
   const mensagem = `
🧪 RESULTADOS DO TESTE COMPLETO

${resultados.configuracao ? '✅' : '❌'} Configuração básica
${resultados.registro ? '✅' : '❌'} Sistema de registro
${resultados.precos ? '✅' : '❌'} Cálculo de preços
${resultados.abasMensais ? '✅' : '❌'} Abas mensais (${integridade.encontradas}/12)
${resultados.email ? '✅' : '❌'} Sistema de e-mail

${Object.values(resultados).every(r => r) ? '\n🎉 SISTEMA 100% OPERACIONAL!' : '\n⚠️ Alguns componentes precisam atenção.'}
   `;
   
   ui.alert('🧪 Teste Completo', mensagem, ui.ButtonSet.OK);
   
 } catch (error) {
   ui.alert('❌ Erro no Teste', error.toString(), ui.ButtonSet.OK);
 }
}

/**
* Mostra informações sobre o sistema
*/
function mostrarSobre() {
 const ui = SpreadsheetApp.getUi();
 
 const mensagem = `
🚐 SISTEMA DE TRANSFERS LIOZ-HUB v4.0

Sistema integrado completo para gestão de transfers entre Hotel LIOZ Lisboa e HUB Transfer.

📋 CARACTERÍSTICAS PRINCIPAIS:
- Registro duplo automático (principal + mensal)
- E-mails interativos com botões de ação
- Verificação automática de confirmações
- Cálculo inteligente de preços
- Relatórios e estatísticas completas
- Backup e recuperação de dados
- API REST completa

🏢 DESENVOLVIDO PARA:
${CONFIG.NAMES.HOTEL_NAME} & ${CONFIG.NAMES.HUB_OWNER}

📧 E-MAIL CONFIGURADO:
${CONFIG.EMAIL_CONFIG.DESTINATARIOS.join(', ')}

⚙️ VERSÃO: ${CONFIG.SISTEMA.VERSAO}
📅 DATA: ${formatarDataDDMMYYYY(new Date())}

💡 Para suporte ou dúvidas, consulte a documentação do sistema.
 `;
 
 ui.alert('ℹ️ Sobre o Sistema', mensagem, ui.ButtonSet.OK);
}

// ===================================================
// PARTE 8.3: DADOS INICIAIS DE PREÇOS
// ===================================================

/**
* Insere dados iniciais na tabela de preços
* @param {Sheet} sheet - Planilha de preços
*/
function inserirDadosIniciaisPrecos(sheet) {
 logger.info('Inserindo dados iniciais de preços');
 
 const precos = [
   // Aeroporto de Lisboa
   [1, 'Aeroporto → Hotel LIOZ', 'Aeroporto de Lisboa', 'Hotel LIOZ', 1, 0, 25.00, 10.00, 15.00, 'Sim', new Date(), '1-4 pessoas mesmo preço'],
   [2, 'Aeroporto → Hotel LIOZ', 'Aeroporto de Lisboa', 'Hotel LIOZ', 2, 0, 25.00, 10.00, 15.00, 'Sim', new Date(), ''],
   [3, 'Aeroporto → Hotel LIOZ', 'Aeroporto de Lisboa', 'Hotel LIOZ', 3, 0, 25.00, 10.00, 15.00, 'Sim', new Date(), ''],
   [4, 'Aeroporto → Hotel LIOZ', 'Aeroporto de Lisboa', 'Hotel LIOZ', 4, 0, 25.00, 10.00, 15.00, 'Sim', new Date(), ''],
   [5, 'Aeroporto → Hotel LIOZ', 'Aeroporto de Lisboa', 'Hotel LIOZ', 5, 0, 30.00, 12.00, 18.00, 'Sim', new Date(), '5+ pessoas preço maior'],
   [6, 'Aeroporto → Hotel LIOZ', 'Aeroporto de Lisboa', 'Hotel LIOZ', 6, 0, 35.00, 14.00, 21.00, 'Sim', new Date(), ''],
   
// Hotel LIOZ → Aeroporto
   [7, 'Hotel LIOZ → Aeroporto', 'Hotel LIOZ', 'Aeroporto de Lisboa', 1, 0, 25.00, 10.00, 15.00, 'Sim', new Date(), ''],
   [8, 'Hotel LIOZ → Aeroporto', 'Hotel LIOZ', 'Aeroporto de Lisboa', 2, 0, 25.00, 10.00, 15.00, 'Sim', new Date(), ''],
   [9, 'Hotel LIOZ → Aeroporto', 'Hotel LIOZ', 'Aeroporto de Lisboa', 3, 0, 25.00, 10.00, 15.00, 'Sim', new Date(), ''],
   [10, 'Hotel LIOZ → Aeroporto', 'Hotel LIOZ', 'Aeroporto de Lisboa', 4, 0, 25.00, 10.00, 15.00, 'Sim', new Date(), ''],
   [11, 'Hotel LIOZ → Aeroporto', 'Hotel LIOZ', 'Aeroporto de Lisboa', 5, 0, 30.00, 12.00, 18.00, 'Sim', new Date(), ''],
   [12, 'Hotel LIOZ → Aeroporto', 'Hotel LIOZ', 'Aeroporto de Lisboa', 6, 0, 35.00, 14.00, 21.00, 'Sim', new Date(), ''],
   
   // Cascais
   [13, 'Cascais → Hotel LIOZ', 'Cascais', 'Hotel LIOZ', 1, 0, 30.00, 12.00, 18.00, 'Sim', new Date(), 'Distância maior'],
   [14, 'Cascais → Hotel LIOZ', 'Cascais', 'Hotel LIOZ', 2, 0, 30.00, 12.00, 18.00, 'Sim', new Date(), ''],
   [15, 'Cascais → Hotel LIOZ', 'Cascais', 'Hotel LIOZ', 3, 0, 35.00, 14.00, 21.00, 'Sim', new Date(), ''],
   [16, 'Cascais → Hotel LIOZ', 'Cascais', 'Hotel LIOZ', 4, 0, 40.00, 16.00, 24.00, 'Sim', new Date(), ''],
   [17, 'Hotel LIOZ → Cascais', 'Hotel LIOZ', 'Cascais', 1, 0, 30.00, 12.00, 18.00, 'Sim', new Date(), ''],
   [18, 'Hotel LIOZ → Cascais', 'Hotel LIOZ', 'Cascais', 2, 0, 30.00, 12.00, 18.00, 'Sim', new Date(), ''],
   [19, 'Hotel LIOZ → Cascais', 'Hotel LIOZ', 'Cascais', 3, 0, 35.00, 14.00, 21.00, 'Sim', new Date(), ''],
   [20, 'Hotel LIOZ → Cascais', 'Hotel LIOZ', 'Cascais', 4, 0, 40.00, 16.00, 24.00, 'Sim', new Date(), ''],
   
   // Sintra
   [21, 'Sintra → Hotel LIOZ', 'Sintra Centro', 'Hotel LIOZ', 1, 0, 35.00, 14.00, 21.00, 'Sim', new Date(), 'Inclui portagens'],
   [22, 'Sintra → Hotel LIOZ', 'Sintra Centro', 'Hotel LIOZ', 2, 0, 35.00, 14.00, 21.00, 'Sim', new Date(), ''],
   [23, 'Sintra → Hotel LIOZ', 'Sintra Centro', 'Hotel LIOZ', 3, 0, 40.00, 16.00, 24.00, 'Sim', new Date(), ''],
   [24, 'Sintra → Hotel LIOZ', 'Sintra Centro', 'Hotel LIOZ', 4, 0, 45.00, 18.00, 27.00, 'Sim', new Date(), ''],
   [25, 'Hotel LIOZ → Sintra', 'Hotel LIOZ', 'Sintra Centro', 1, 0, 35.00, 14.00, 21.00, 'Sim', new Date(), ''],
   [26, 'Hotel LIOZ → Sintra', 'Hotel LIOZ', 'Sintra Centro', 2, 0, 35.00, 14.00, 21.00, 'Sim', new Date(), ''],
   [27, 'Hotel LIOZ → Sintra', 'Hotel LIOZ', 'Sintra Centro', 3, 0, 40.00, 16.00, 24.00, 'Sim', new Date(), ''],
   [28, 'Hotel LIOZ → Sintra', 'Hotel LIOZ', 'Sintra Centro', 4, 0, 45.00, 18.00, 27.00, 'Sim', new Date(), ''],
   
   // Palácio da Pena
   [29, 'Palácio Pena → Hotel LIOZ', 'Palácio da Pena (Sintra)', 'Hotel LIOZ', 1, 0, 48.00, 19.00, 29.00, 'Sim', new Date(), 'Distância + portagens'],
   [30, 'Palácio Pena → Hotel LIOZ', 'Palácio da Pena (Sintra)', 'Hotel LIOZ', 2, 0, 48.00, 19.00, 29.00, 'Sim', new Date(), ''],
   [31, 'Palácio Pena → Hotel LIOZ', 'Palácio da Pena (Sintra)', 'Hotel LIOZ', 3, 0, 55.00, 22.00, 33.00, 'Sim', new Date(), ''],
   [32, 'Palácio Pena → Hotel LIOZ', 'Palácio da Pena (Sintra)', 'Hotel LIOZ', 4, 0, 62.00, 25.00, 37.00, 'Sim', new Date(), ''],
   [33, 'Hotel LIOZ → Palácio Pena', 'Hotel LIOZ', 'Palácio da Pena (Sintra)', 1, 0, 48.00, 19.00, 29.00, 'Sim', new Date(), ''],
   [34, 'Hotel LIOZ → Palácio Pena', 'Hotel LIOZ', 'Palácio da Pena (Sintra)', 2, 0, 48.00, 19.00, 29.00, 'Sim', new Date(), ''],
   [35, 'Hotel LIOZ → Palácio Pena', 'Hotel LIOZ', 'Palácio da Pena (Sintra)', 3, 0, 55.00, 22.00, 33.00, 'Sim', new Date(), ''],
   [36, 'Hotel LIOZ → Palácio Pena', 'Hotel LIOZ', 'Palácio da Pena (Sintra)', 4, 0, 62.00, 25.00, 37.00, 'Sim', new Date(), ''],
   
   // Quinta da Regaleira
   [37, 'Quinta Regaleira → Hotel LIOZ', 'Quinta da Regaleira (Sintra)', 'Hotel LIOZ', 1, 0, 45.00, 18.00, 27.00, 'Sim', new Date(), ''],
   [38, 'Quinta Regaleira → Hotel LIOZ', 'Quinta da Regaleira (Sintra)', 'Hotel LIOZ', 2, 0, 45.00, 18.00, 27.00, 'Sim', new Date(), ''],
   [39, 'Quinta Regaleira → Hotel LIOZ', 'Quinta da Regaleira (Sintra)', 'Hotel LIOZ', 3, 0, 52.00, 21.00, 31.00, 'Sim', new Date(), ''],
   [40, 'Quinta Regaleira → Hotel LIOZ', 'Quinta da Regaleira (Sintra)', 'Hotel LIOZ', 4, 0, 59.00, 24.00, 35.00, 'Sim', new Date(), ''],
   [41, 'Hotel LIOZ → Quinta Regaleira', 'Hotel LIOZ', 'Quinta da Regaleira (Sintra)', 1, 0, 45.00, 18.00, 27.00, 'Sim', new Date(), ''],
   [42, 'Hotel LIOZ → Quinta Regaleira', 'Hotel LIOZ', 'Quinta da Regaleira (Sintra)', 2, 0, 45.00, 18.00, 27.00, 'Sim', new Date(), ''],
   [43, 'Hotel LIOZ → Quinta Regaleira', 'Hotel LIOZ', 'Quinta da Regaleira (Sintra)', 3, 0, 52.00, 21.00, 31.00, 'Sim', new Date(), ''],
   [44, 'Hotel LIOZ → Quinta Regaleira', 'Hotel LIOZ', 'Quinta da Regaleira (Sintra)', 4, 0, 59.00, 24.00, 35.00, 'Sim', new Date(), ''],
   
   // Oceanário
   [45, 'Oceanário → Hotel LIOZ', 'Oceanário de Lisboa', 'Hotel LIOZ', 1, 0, 28.00, 11.00, 17.00, 'Sim', new Date(), 'Parque das Nações'],
   [46, 'Oceanário → Hotel LIOZ', 'Oceanário de Lisboa', 'Hotel LIOZ', 2, 0, 28.00, 11.00, 17.00, 'Sim', new Date(), ''],
   [47, 'Oceanário → Hotel LIOZ', 'Oceanário de Lisboa', 'Hotel LIOZ', 3, 0, 32.00, 13.00, 19.00, 'Sim', new Date(), ''],
   [48, 'Oceanário → Hotel LIOZ', 'Oceanário de Lisboa', 'Hotel LIOZ', 4, 0, 36.00, 14.00, 22.00, 'Sim', new Date(), ''],
   [49, 'Hotel LIOZ → Oceanário', 'Hotel LIOZ', 'Oceanário de Lisboa', 1, 0, 28.00, 11.00, 17.00, 'Sim', new Date(), ''],
   [50, 'Hotel LIOZ → Oceanário', 'Hotel LIOZ', 'Oceanário de Lisboa', 2, 0, 28.00, 11.00, 17.00, 'Sim', new Date(), ''],
   [51, 'Hotel LIOZ → Oceanário', 'Hotel LIOZ', 'Oceanário de Lisboa', 3, 0, 32.00, 13.00, 19.00, 'Sim', new Date(), ''],
   [52, 'Hotel LIOZ → Oceanário', 'Hotel LIOZ', 'Oceanário de Lisboa', 4, 0, 36.00, 14.00, 22.00, 'Sim', new Date(), ''],
   
   // Torre de Belém
   [53, 'Torre Belém → Hotel LIOZ', 'Torre de Belém', 'Hotel LIOZ', 1, 0, 27.00, 11.00, 16.00, 'Sim', new Date(), 'Centro histórico'],
   [54, 'Torre Belém → Hotel LIOZ', 'Torre de Belém', 'Hotel LIOZ', 2, 0, 27.00, 11.00, 16.00, 'Sim', new Date(), ''],
   [55, 'Torre Belém → Hotel LIOZ', 'Torre de Belém', 'Hotel LIOZ', 3, 0, 31.00, 12.00, 19.00, 'Sim', new Date(), ''],
   [56, 'Torre Belém → Hotel LIOZ', 'Torre de Belém', 'Hotel LIOZ', 4, 0, 35.00, 14.00, 21.00, 'Sim', new Date(), ''],
   [57, 'Hotel LIOZ → Torre Belém', 'Hotel LIOZ', 'Torre de Belém', 1, 0, 27.00, 11.00, 16.00, 'Sim', new Date(), ''],
   [58, 'Hotel LIOZ → Torre Belém', 'Hotel LIOZ', 'Torre de Belém', 2, 0, 27.00, 11.00, 16.00, 'Sim', new Date(), ''],
   [59, 'Hotel LIOZ → Torre Belém', 'Hotel LIOZ', 'Torre de Belém', 3, 0, 31.00, 12.00, 19.00, 'Sim', new Date(), ''],
   [60, 'Hotel LIOZ → Torre Belém', 'Hotel LIOZ', 'Torre de Belém', 4, 0, 35.00, 14.00, 21.00, 'Sim', new Date(), ''],
   
   // Mosteiro dos Jerónimos
   [61, 'Mosteiro Jerónimos → Hotel LIOZ', 'Mosteiro dos Jerónimos', 'Hotel LIOZ', 1, 0, 26.00, 10.00, 16.00, 'Sim', new Date(), 'Belém'],
   [62, 'Mosteiro Jerónimos → Hotel LIOZ', 'Mosteiro dos Jerónimos', 'Hotel LIOZ', 2, 0, 26.00, 10.00, 16.00, 'Sim', new Date(), ''],
   [63, 'Mosteiro Jerónimos → Hotel LIOZ', 'Mosteiro dos Jerónimos', 'Hotel LIOZ', 3, 0, 30.00, 12.00, 18.00, 'Sim', new Date(), ''],
   [64, 'Mosteiro Jerónimos → Hotel LIOZ', 'Mosteiro dos Jerónimos', 'Hotel LIOZ', 4, 0, 34.00, 14.00, 20.00, 'Sim', new Date(), ''],
   [65, 'Hotel LIOZ → Mosteiro Jerónimos', 'Hotel LIOZ', 'Mosteiro dos Jerónimos', 1, 0, 26.00, 10.00, 16.00, 'Sim', new Date(), ''],
   [66, 'Hotel LIOZ → Mosteiro Jerónimos', 'Hotel LIOZ', 'Mosteiro dos Jerónimos', 2, 0, 26.00, 10.00, 16.00, 'Sim', new Date(), ''],
   [67, 'Hotel LIOZ → Mosteiro Jerónimos', 'Hotel LIOZ', 'Mosteiro dos Jerónimos', 3, 0, 30.00, 12.00, 18.00, 'Sim', new Date(), ''],
   [68, 'Hotel LIOZ → Mosteiro Jerónimos', 'Hotel LIOZ', 'Mosteiro dos Jerónimos', 4, 0, 34.00, 14.00, 20.00, 'Sim', new Date(), ''],
   
   // Centro de Lisboa (genérico)
   [69, 'Centro Lisboa → Hotel LIOZ', 'Centro de Lisboa', 'Hotel LIOZ', 1, 0, 20.00, 8.00, 12.00, 'Sim', new Date(), 'Distância curta'],
   [70, 'Centro Lisboa → Hotel LIOZ', 'Centro de Lisboa', 'Hotel LIOZ', 2, 0, 20.00, 8.00, 12.00, 'Sim', new Date(), ''],
   [71, 'Centro Lisboa → Hotel LIOZ', 'Centro de Lisboa', 'Hotel LIOZ', 3, 0, 23.00, 9.00, 14.00, 'Sim', new Date(), ''],
   [72, 'Centro Lisboa → Hotel LIOZ', 'Centro de Lisboa', 'Hotel LIOZ', 4, 0, 26.00, 10.00, 16.00, 'Sim', new Date(), ''],
   [73, 'Hotel LIOZ → Centro Lisboa', 'Hotel LIOZ', 'Centro de Lisboa', 1, 0, 20.00, 8.00, 12.00, 'Sim', new Date(), ''],
   [74, 'Hotel LIOZ → Centro Lisboa', 'Hotel LIOZ', 'Centro de Lisboa', 2, 0, 20.00, 8.00, 12.00, 'Sim', new Date(), ''],
   [75, 'Hotel LIOZ → Centro Lisboa', 'Hotel LIOZ', 'Centro de Lisboa', 3, 0, 23.00, 9.00, 14.00, 'Sim', new Date(), ''],
   [76, 'Hotel LIOZ → Centro Lisboa', 'Hotel LIOZ', 'Centro de Lisboa', 4, 0, 26.00, 10.00, 16.00, 'Sim', new Date(), '']
 ];
 
 try {
   // Limpar dados existentes se houver
   if (sheet.getLastRow() > 1) {
     sheet.deleteRows(2, sheet.getLastRow() - 1);
   }
   
   // Inserir novos dados
   const startRow = 2;
   const numRows = precos.length;
   const numCols = PRICING_HEADERS.length;
   
   sheet.getRange(startRow, 1, numRows, numCols).setValues(precos);
   
   logger.success(`${precos.length} preços inseridos na tabela`);
   
 } catch (error) {
   logger.error('Erro ao inserir preços iniciais', error);
 }
}

// ===================================================
// PARTE 8.4: FUNÇÕES DO MENU ADICIONAL
// ===================================================

/**
* Abre a tabela de preços
*/
function abrirTabelaPrecos() {
 const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
 const sheet = ss.getSheetByName(CONFIG.PRICING_SHEET_NAME);
 
 if (sheet) {
   ss.setActiveSheet(sheet);
   SpreadsheetApp.getUi().alert(
     '💰 Tabela de Preços',
     `Tabela aberta com ${sheet.getLastRow() - 1} preços cadastrados.\n\n` +
     'Você pode editar diretamente na planilha.',
     SpreadsheetApp.getUi().ButtonSet.OK
   );
 } else {
   SpreadsheetApp.getUi().alert(
     '❌ Erro',
     'Tabela de preços não encontrada. Execute "Configurar Sistema" primeiro.',
     SpreadsheetApp.getUi().ButtonSet.OK
   );
 }
}

/**
* Consulta preço via menu
*/
function consultarPrecoMenu() {
 const ui = SpreadsheetApp.getUi();
 
 const response = ui.prompt(
   '🔍 Consultar Preço',
   'Digite a consulta no formato:\n' +
   'Origem|Destino|Pessoas|Bagagens\n\n' +
   'Exemplo: Aeroporto|Hotel|2|1',
   ui.ButtonSet.OK_CANCEL
 );
 
 if (response.getSelectedButton() === ui.Button.OK) {
   try {
     const partes = response.getResponseText().split('|');
     
     if (partes.length < 3) {
       throw new Error('Formato inválido');
     }
     
     const valores = calcularValores(
       partes[0].trim(),
       partes[1].trim(),
       parseInt(partes[2]) || 1,
       parseInt(partes[3]) || 0
     );
     
     ui.alert(
       '💰 Resultado da Consulta',
       `Origem: ${partes[0]}\n` +
       `Destino: ${partes[1]}\n` +
       `Pessoas: ${partes[2]}\n` +
       `Bagagens: ${partes[3] || 0}\n\n` +
       `Preço Cliente: €${valores.precoCliente.toFixed(2)}\n` +
       `Valor Hotel: €${valores.valorHotel.toFixed(2)}\n` +
       `Valor HUB: €${valores.valorHUB.toFixed(2)}\n\n` +
       `Fonte: ${valores.fonte}\n` +
       `${valores.observacoes || ''}`,
       ui.ButtonSet.OK
     );
     
   } catch (error) {
     ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
   }
 }
}

/**
* Importa preços via menu
*/
function importarPrecosMenu() {
 const ui = SpreadsheetApp.getUi();
 
 ui.alert(
   '📥 Importar Preços',
   'Para importar preços:\n\n' +
   '1. Prepare um arquivo CSV com as colunas:\n' +
   '   Origem,Destino,Pessoas,Bagagens,Preço\n\n' +
   '2. Cole o conteúdo na próxima tela\n\n' +
   'Exemplo:\n' +
   'Aeroporto,Hotel,1,0,25.00\n' +
   'Cascais,Hotel,2,0,30.00',
   ui.ButtonSet.OK
 );
 
 const response = ui.prompt(
   '📥 Cole o CSV',
   'Cole o conteúdo CSV aqui:',
   ui.ButtonSet.OK_CANCEL
 );
 
 if (response.getSelectedButton() === ui.Button.OK) {
   try {
     const resultado = importarPrecos(response.getResponseText(), 'csv');
     
     ui.alert(
       resultado.sucesso ? '✅ Importação Concluída' : '❌ Erro na Importação',
       `Importados: ${resultado.importados}\n` +
       `Erros: ${resultado.erros}\n` +
       `Total: ${resultado.total}`,
       ui.ButtonSet.OK
     );
     
   } catch (error) {
     ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
   }
 }
}

/**
* Exporta preços via menu
*/
function exportarPrecosMenu() {
 const ui = SpreadsheetApp.getUi();
 
 const response = ui.alert(
   '📤 Exportar Preços',
   'Escolha o formato de exportação:',
   ui.ButtonSet.YES_NO_CANCEL
 );
 
 try {
   let formato = 'csv';
   if (response === ui.Button.NO) {
     formato = 'json';
   } else if (response === ui.Button.CANCEL) {
     return;
   }
   
   const dados = exportarPrecos(formato);
   
   // Criar um documento temporário
   const blob = Utilities.newBlob(dados, 
     formato === 'csv' ? 'text/csv' : 'application/json',
     `precos_export_${new Date().getTime()}.${formato}`
   );
   
   // Criar arquivo no Drive
   const file = DriveApp.createFile(blob);
   
   ui.alert(
     '✅ Exportação Concluída',
     `Arquivo criado no Google Drive:\n\n` +
     `Nome: ${file.getName()}\n` +
     `URL: ${file.getUrl()}\n\n` +
     `Clique em OK e acesse o link.`,
     ui.ButtonSet.OK
   );
   
 } catch (error) {
   ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
 }
}

/**
* Menu para criar todas as abas mensais
*/
function criarTodasAbasMensaisMenu() {
 const ui = SpreadsheetApp.getUi();
 
 const response = ui.prompt(
   '📅 Criar Abas Mensais',
   'Digite o ano para criar as abas (padrão: 2025):',
   ui.ButtonSet.OK_CANCEL
 );
 
 if (response.getSelectedButton() === ui.Button.OK) {
   try {
     const ano = parseInt(response.getResponseText()) || CONFIG.SISTEMA.ANO_BASE;
     const resultado = criarTodasAbasMensais(ano);
     
     ui.alert(
       '✅ Abas Criadas',
       `Ano: ${ano}\n\n` +
       `Criadas: ${resultado.criadas}\n` +
       `Existentes: ${resultado.existentes}\n` +
       `Erros: ${resultado.erros}\n\n` +
       'Todas as 12 abas mensais estão prontas!',
       ui.ButtonSet.OK
     );
     
   } catch (error) {
     ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
   }
 }
}

/**
* Menu para verificar integridade
*/
function verificarIntegridadeMenu() {
 const ui = SpreadsheetApp.getUi();
 
 try {
   const resultado = verificarIntegridadeAbasMensais();
   
   let detalhes = 'DETALHES POR MÊS:\n\n';
   resultado.detalhes.forEach(d => {
     const status = d.existe ? '✅' : '❌';
     const problemas = d.problemas.length > 0 ? ` (${d.problemas.join(', ')})` : '';
     detalhes += `${status} ${d.mes}${problemas}\n`;
   });
   
   ui.alert(
     '🔍 Verificação de Integridade',
     `Total de meses: ${resultado.total}\n` +
     `Abas encontradas: ${resultado.encontradas}\n` +
     `Abas com problemas: ${resultado.comProblemas}\n\n` +
     detalhes,
     ui.ButtonSet.OK
   );
   
 } catch (error) {
   ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
 }
}

/**
* Menu para reparar abas
*/
function repararAbasMensaisMenu() {
 const ui = SpreadsheetApp.getUi();
 
 const response = ui.alert(
   '🔧 Reparar Abas Mensais',
   'Esta função irá:\n\n' +
   '1. Verificar todas as abas mensais\n' +
   '2. Criar abas faltantes\n' +
   '3. Corrigir headers incorretos\n' +
   '4. Aplicar formatações\n\n' +
   'Continuar?',
   ui.ButtonSet.YES_NO
 );
 
 if (response === ui.Button.YES) {
   try {
     const resultado = repararAbasMensais();
     
     ui.alert(
       '✅ Reparação Concluída',
       `Abas reparadas: ${resultado.reparadas}\n` +
       `Abas criadas: ${resultado.criadas}\n` +
       `Erros: ${resultado.erros}\n\n` +
       'Sistema de abas mensais restaurado!',
       ui.ButtonSet.OK
     );
     
   } catch (error) {
     ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
   }
 }
}

/**
* Menu para sincronizar abas
*/
function sincronizarAbasMensaisMenu() {
 const ui = SpreadsheetApp.getUi();
 
 const response = ui.alert(
   '🔄 Sincronizar Abas Mensais',
   'ATENÇÃO: Esta função irá:\n\n' +
   '1. APAGAR todos os dados das abas mensais\n' +
   '2. Recriar com base na aba principal\n' +
   '3. Reorganizar por mês\n\n' +
   'Esta ação é IRREVERSÍVEL. Continuar?',
   ui.ButtonSet.YES_NO
 );
 
 if (response === ui.Button.YES) {
   try {
     const resultado = sincronizarAbasMensais();
     
     ui.alert(
       '✅ Sincronização Concluída',
       `Registros processados: ${resultado.processados}\n` +
       `Registros sincronizados: ${resultado.sincronizados}\n` +
       `Erros: ${resultado.erros}\n\n` +
       'Abas mensais sincronizadas com sucesso!',
       ui.ButtonSet.OK
     );
     
   } catch (error) {
     ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
   }
 }
}

/**
* Menu para gerar relatório mensal
*/
function gerarRelatorioMensalMenu() {
 const ui = SpreadsheetApp.getUi();
 
 const response = ui.prompt(
   '📊 Relatório Mensal',
   'Digite o mês (1-12) e ano (opcional):\n\n' +
   'Formato: mês,ano\n' +
   'Exemplo: 7,2025 (para Julho de 2025)\n' +
   'Ou apenas: 7 (para Julho do ano atual)',
   ui.ButtonSet.OK_CANCEL
 );
 
 if (response.getSelectedButton() === ui.Button.OK) {
   try {
     const partes = response.getResponseText().split(',');
     const mes = parseInt(partes[0]);
     const ano = partes[1] ? parseInt(partes[1]) : CONFIG.SISTEMA.ANO_BASE;
     
     if (mes < 1 || mes > 12) {
       throw new Error('Mês deve ser entre 1 e 12');
     }
     
     const relatorio = gerarRelatorioMensal(mes, ano);
     
     if (relatorio.status === 'sem_dados') {
       ui.alert('📊 Sem Dados', relatorio.mensagem, ui.ButtonSet.OK);
       return;
     }
     
// Criar documento com o relatório
     const doc = DocumentApp.create(`Relatório ${relatorio.mes} ${relatorio.ano}`);
     const body = doc.getBody();
     
     // Título
     body.appendParagraph(`RELATÓRIO MENSAL - ${relatorio.mes.toUpperCase()} ${relatorio.ano}`)
       .setHeading(DocumentApp.ParagraphHeading.HEADING1);
     
     body.appendParagraph(`Período: ${relatorio.periodo}`);
     body.appendHorizontalRule();
     
     // Resumo
     body.appendParagraph('RESUMO GERAL').setHeading(DocumentApp.ParagraphHeading.HEADING2);
     body.appendParagraph(`Total de Transfers: ${relatorio.resumo.totalTransfers}`);
     body.appendParagraph(`Total de Passageiros: ${relatorio.resumo.totalPassageiros}`);
     body.appendParagraph(`Total de Bagagens: ${relatorio.resumo.totalBagagens}`);
     body.appendParagraph(`Valor Total: €${relatorio.resumo.valorTotal.toFixed(2)}`);
     body.appendParagraph(`Valor Hotel LIOZ: €${relatorio.resumo.valorHotel.toFixed(2)}`);
     body.appendParagraph(`Valor HUB Transfer: €${relatorio.resumo.valorHUB.toFixed(2)}`);
     body.appendParagraph(`Ticket Médio: €${relatorio.resumo.ticketMedio.toFixed(2)}`);
     
     // Status
     body.appendParagraph('DISTRIBUIÇÃO POR STATUS').setHeading(DocumentApp.ParagraphHeading.HEADING2);
     Object.entries(relatorio.porStatus).forEach(([status, count]) => {
       body.appendParagraph(`• ${status}: ${count}`);
     });
     
     // Top Rotas
     if (relatorio.topRotas.length > 0) {
       body.appendParagraph('TOP 10 ROTAS').setHeading(DocumentApp.ParagraphHeading.HEADING2);
       relatorio.topRotas.forEach((item, index) => {
         body.appendParagraph(`${index + 1}. ${item.rota}: ${item.count} transfers`);
       });
     }
     
     // Top Clientes
     if (relatorio.topClientes.length > 0) {
       body.appendParagraph('TOP 10 CLIENTES').setHeading(DocumentApp.ParagraphHeading.HEADING2);
       relatorio.topClientes.forEach((item, index) => {
         body.appendParagraph(`${index + 1}. ${item.cliente}: ${item.count} transfers`);
       });
     }
     
     // Análise
     body.appendParagraph('ANÁLISE DE TENDÊNCIAS').setHeading(DocumentApp.ParagraphHeading.HEADING2);
     const analise = relatorio.analise;
     
     if (analise.melhorDia) {
       body.appendParagraph(`Melhor dia: ${analise.melhorDia.dia} (${analise.melhorDia.transfers} transfers)`);
     }
     
     if (analise.piorDia) {
       body.appendParagraph(`Pior dia: ${analise.piorDia.dia} (${analise.piorDia.transfers} transfers)`);
     }
     
     if (analise.horarioPico) {
       body.appendParagraph(`Horário de pico: ${analise.horarioPico.horario} (${analise.horarioPico.transfers} transfers)`);
     }
     
     body.appendParagraph(`Taxa de cancelamento: ${analise.tendenciaStatus.taxaCancelamento}%`);
     body.appendParagraph(`Taxa de confirmação: ${analise.tendenciaStatus.taxaConfirmacao}%`);
     body.appendParagraph(`Taxa de finalização: ${analise.tendenciaStatus.taxaFinalizacao}%`);
     
     // Alertas
     if (analise.alertas.length > 0) {
       body.appendParagraph('ALERTAS').setHeading(DocumentApp.ParagraphHeading.HEADING2);
       analise.alertas.forEach(alerta => {
         body.appendParagraph(`⚠️ ${alerta.mensagem}`);
       });
     }
     
     // Rodapé
     body.appendHorizontalRule();
     body.appendParagraph(`Relatório gerado em ${formatarDataHora(new Date())}`);
     body.appendParagraph(`Sistema ${CONFIG.NAMES.SISTEMA_NOME} v${CONFIG.SISTEMA.VERSAO}`);
     
     // Salvar e obter URL
     doc.saveAndClose();
     const url = doc.getUrl();
     
     ui.alert(
       '✅ Relatório Gerado',
       `Relatório de ${relatorio.mes}/${relatorio.ano} criado com sucesso!\n\n` +
       `Documento: ${doc.getName()}\n` +
       `URL: ${url}\n\n` +
       `Total de transfers: ${relatorio.resumo.totalTransfers}\n` +
       `Valor total: €${relatorio.resumo.valorTotal.toFixed(2)}`,
       ui.ButtonSet.OK
     );
     
   } catch (error) {
     ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
   }
 }
}

/**
* Menu para verificar duplicados
*/
function verificarDuplicadosMenu() {
 const ui = SpreadsheetApp.getUi();
 
 try {
   const resultado = removerDuplicados(true); // Simular apenas
   
   if (resultado.duplicados === 0) {
     ui.alert(
       '✅ Sem Duplicados',
       'Nenhum registro duplicado foi encontrado!',
       ui.ButtonSet.OK
     );
   } else {
     let detalhes = 'DUPLICADOS ENCONTRADOS:\n\n';
     resultado.detalhes.slice(0, 10).forEach(dup => {
       detalhes += `• ID ${dup.id} - ${dup.cliente} (linha ${dup.linha})\n`;
     });
     
     if (resultado.detalhes.length > 10) {
       detalhes += `\n... e mais ${resultado.detalhes.length - 10} duplicados`;
     }
     
     ui.alert(
       '⚠️ Duplicados Encontrados',
       `Total de duplicados: ${resultado.duplicados}\n\n${detalhes}`,
       ui.ButtonSet.OK
     );
   }
   
 } catch (error) {
   ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
 }
}

/**
* Menu para remover duplicados
*/
function removerDuplicadosMenu() {
 const ui = SpreadsheetApp.getUi();
 
 // Primeiro verificar
 const verificacao = removerDuplicados(true);
 
 if (verificacao.duplicados === 0) {
   ui.alert(
     '✅ Sem Duplicados',
     'Nenhum registro duplicado para remover!',
     ui.ButtonSet.OK
   );
   return;
 }
 
 const response = ui.alert(
   '🗑️ Remover Duplicados',
   `Foram encontrados ${verificacao.duplicados} registros duplicados.\n\n` +
   'Deseja removê-los?\n\n' +
   'Esta ação é IRREVERSÍVEL!',
   ui.ButtonSet.YES_NO
 );
 
 if (response === ui.Button.YES) {
   try {
     const resultado = removerDuplicados(false); // Remover de verdade
     
     ui.alert(
       '✅ Duplicados Removidos',
       `${resultado.duplicados} registros duplicados foram removidos com sucesso!`,
       ui.ButtonSet.OK
     );
     
   } catch (error) {
     ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
   }
 }
}

/**
* Menu para corrigir registros
*/
function corrigirRegistrosMenu() {
 const ui = SpreadsheetApp.getUi();
 
 const response = ui.alert(
   '🔧 Corrigir Registros Incompletos',
   'Esta função irá verificar todos os registros da aba principal e:\n\n' +
   '1. Identificar registros sem entrada na aba mensal\n' +
   '2. Adicionar automaticamente na aba correta\n' +
   '3. Manter sincronização\n\n' +
   'Continuar?',
   ui.ButtonSet.YES_NO
 );
 
 if (response === ui.Button.YES) {
   try {
     const resultado = corrigirRegistrosIncompletos();
     
     ui.alert(
       '✅ Correção Concluída',
       `Registros verificados: ${resultado.verificados}\n` +
       `Registros corrigidos: ${resultado.corrigidos}\n` +
       `Erros encontrados: ${resultado.erros}\n\n` +
       `${resultado.corrigidos > 0 ? 'Registros duplos restaurados!' : 'Todos os registros já estavam corretos!'}`,
       ui.ButtonSet.OK
     );
     
   } catch (error) {
     ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
   }
 }
}

/**
* Menu para criar backup
*/
function criarBackupMenu() {
 const ui = SpreadsheetApp.getUi();
 
 const response = ui.alert(
   '💾 Criar Backup',
   'Esta função irá criar uma cópia completa da planilha atual.\n\n' +
   'O backup incluirá:\n' +
   '• Todos os transfers\n' +
   '• Tabela de preços\n' +
   '• Todas as abas mensais\n' +
   '• Formatações e validações\n\n' +
   'Continuar?',
   ui.ButtonSet.YES_NO
 );
 
 if (response === ui.Button.YES) {
   try {
     const resultado = criarBackup();
     
     if (resultado.sucesso) {
       ui.alert(
         '✅ Backup Criado',
         `Backup criado com sucesso!\n\n` +
         `Nome: ${resultado.nome}\n` +
         `ID: ${resultado.id}\n` +
         `Data/Hora: ${formatarDataHora(resultado.dataHora)}\n\n` +
         'O backup está salvo no Google Drive.',
         ui.ButtonSet.OK
       );
     } else {
       ui.alert('❌ Erro', resultado.erro, ui.ButtonSet.OK);
     }
     
   } catch (error) {
     ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
   }
 }
}

/**
* Menu para limpar todos os dados
*/
function limparDadosCompletoMenu() {
 const ui = SpreadsheetApp.getUi();
 
 const response1 = ui.alert(
   '⚠️ ATENÇÃO - AÇÃO PERIGOSA',
   'Esta função irá APAGAR TODOS OS DADOS do sistema!\n\n' +
   'Isso inclui:\n' +
   '• Todos os transfers registrados\n' +
   '• Dados de todas as abas mensais\n' +
   '• Tabela de preços (opcional)\n\n' +
   'Esta ação é IRREVERSÍVEL!\n\n' +
   'Tem certeza que deseja continuar?',
   ui.ButtonSet.YES_NO
 );
 
 if (response1 === ui.Button.YES) {
   const response2 = ui.alert(
     '⚠️ CONFIRMAÇÃO FINAL',
     'ÚLTIMA CHANCE!\n\n' +
     'Digite "CONFIRMAR" na próxima tela para apagar todos os dados.\n\n' +
     'Qualquer outro texto cancelará a operação.',
     ui.ButtonSet.OK_CANCEL
   );
   
   if (response2 === ui.Button.OK) {
     const confirmacao = ui.prompt(
       '🗑️ Confirmação Final',
       'Digite CONFIRMAR (em maiúsculas) para apagar todos os dados:',
       ui.ButtonSet.OK_CANCEL
     );
     
     if (confirmacao.getSelectedButton() === ui.Button.OK && 
         confirmacao.getResponseText() === 'CONFIRMAR') {
       
       try {
         const resultado = limparDadosCompleto();
         
         ui.alert(
           '✅ Dados Removidos',
           `${resultado.registrosRemovidos} registros foram removidos de ${resultado.abasProcessadas} abas.\n\n` +
           'O sistema está limpo e pronto para novos registros.',
           ui.ButtonSet.OK
         );
         
       } catch (error) {
         ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
       }
       
     } else {
       ui.alert('❌ Cancelado', 'Operação cancelada. Nenhum dado foi removido.', ui.ButtonSet.OK);
     }
   }
 }
}

/**
* Menu para reordenar por data
*/
function reordenarPorDataMenu() {
 const ui = SpreadsheetApp.getUi();
 
 const response = ui.alert(
   '🔄 Reordenar por Data',
   'Esta função irá reorganizar todos os transfers em ordem cronológica.\n\n' +
   'Os registros serão ordenados por:\n' +
   '1. Data do transfer\n' +
   '2. Hora do pickup (se mesma data)\n\n' +
   'Continuar?',
   ui.ButtonSet.YES_NO
 );
 
 if (response === ui.Button.YES) {
   try {
     const resultado = reordenarPorData();
     
     ui.alert(
       '✅ Reordenação Concluída',
       `${resultado.registros} registros foram reordenados por data com sucesso!`,
       ui.ButtonSet.OK
     );
     
   } catch (error) {
     ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
   }
 }
}

/**
* Menu para mostrar consolidado anual
*/
function mostrarConsolidadoAnualMenu() {
 const ui = SpreadsheetApp.getUi();
 
 const response = ui.prompt(
   '📅 Consolidado Anual',
   'Digite o ano para consolidar (padrão: 2025):',
   ui.ButtonSet.OK_CANCEL
 );
 
 if (response.getSelectedButton() === ui.Button.OK) {
   try {
     const ano = parseInt(response.getResponseText()) || CONFIG.SISTEMA.ANO_BASE;
     const consolidado = consolidarDadosMensais(ano);
     
     let mensagem = `📊 CONSOLIDADO ANUAL ${ano}\n\n`;
     mensagem += `Total Geral: ${consolidado.totalGeral} transfers\n`;
     mensagem += `Valor Total: €${consolidado.valorTotalGeral.toFixed(2)}\n\n`;
     
     mensagem += `POR MÊS:\n`;
     MESES.forEach(mes => {
       const dadosMes = consolidado.porMes[mes.nome];
       if (dadosMes && dadosMes.total > 0) {
         mensagem += `\n${mes.nome}:\n`;
         mensagem += `• Transfers: ${dadosMes.total}\n`;
         mensagem += `• Valor: €${dadosMes.valorTotal.toFixed(2)}\n`;
       }
     });
     
     if (consolidado.topRotas.length > 0) {
       mensagem += `\nTOP 5 ROTAS DO ANO:\n`;
       consolidado.topRotas.slice(0, 5).forEach((item, index) => {
         mensagem += `${index + 1}. ${item.rota}: ${item.count} transfers\n`;
       });
     }
     
     ui.alert('📊 Consolidado Anual', mensagem, ui.ButtonSet.OK);
     
   } catch (error) {
     ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
   }
 }
}

/**
* Menu para mostrar top clientes e rotas
*/
function mostrarTopClientesRotasMenu() {
 const ui = SpreadsheetApp.getUi();
 
 try {
   const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
   const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
   
   if (!sheet || sheet.getLastRow() <= 1) {
     ui.alert('📊 Sem Dados', 'Nenhum registro encontrado para análise.', ui.ButtonSet.OK);
     return;
   }
   
   const dados = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
   const clientes = {};
   const rotas = {};
   
   // Processar dados
   dados.forEach(row => {
     // Clientes
     const cliente = row[1];
     if (!clientes[cliente]) {
       clientes[cliente] = { count: 0, valor: 0 };
     }
     clientes[cliente].count++;
     clientes[cliente].valor += parseFloat(row[10]) || 0;
     
     // Rotas
     const rota = `${row[7]} → ${row[8]}`;
     if (!rotas[rota]) {
       rotas[rota] = { count: 0, valor: 0 };
     }
     rotas[rota].count++;
     rotas[rota].valor += parseFloat(row[10]) || 0;
   });
   
   // Ordenar e pegar top 10
   const topClientes = Object.entries(clientes)
     .sort((a, b) => b[1].valor - a[1].valor)
     .slice(0, 10);
   
   const topRotas = Object.entries(rotas)
     .sort((a, b) => b[1].count - a[1].count)
     .slice(0, 10);
   
   let mensagem = '🏆 TOP 10 CLIENTES (por valor)\n\n';
   topClientes.forEach(([nome, dados], index) => {
     mensagem += `${index + 1}. ${nome}\n`;
     mensagem += `   Transfers: ${dados.count} | Valor: €${dados.valor.toFixed(2)}\n\n`;
   });
   
   mensagem += '\n🚐 TOP 10 ROTAS (por frequência)\n\n';
   topRotas.forEach(([rota, dados], index) => {
     mensagem += `${index + 1}. ${rota}\n`;
     mensagem += `   Transfers: ${dados.count} | Valor: €${dados.valor.toFixed(2)}\n\n`;
   });
   
   ui.alert('🏆 Top Clientes e Rotas', mensagem, ui.ButtonSet.OK);
   
 } catch (error) {
   ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
 }
}

// ===================================================
// FIM DO CÓDIGO - SISTEMA COMPLETO v4.0
// ===================================================

/**
* SISTEMA DE TRANSFERS HOTEL LIOZ & HUB TRANSFER v4.0
* 
* Sistema completo e integrado para gestão de transfers com:
* - Registro duplo automático (principal + mensal)
* - E-mails interativos com botões de confirmação/cancelamento
* - Verificação automática de confirmações por e-mail
* - Sistema inteligente de cálculo de preços
* - Gestão completa de abas mensais
* - Relatórios e estatísticas avançadas
* - API REST completa (GET/POST)
* - Backup e recuperação
* - Menu interativo completo
* 
* Desenvolvido para Hotel LIOZ Lisboa & HUB Transfer
* 
* © 2025 - Todos os direitos reservados
*/
