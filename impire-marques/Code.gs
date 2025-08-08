// ===================================================
// PARTE 11: MENU DO SISTEMA
// ===================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🚐 Sistema Marques Empire')
    .addItem('⚙️ Configurar Sistema', 'configurarSistema')
    .addSeparator()
    .addItem('📊 Inserir Preços Iniciais', 'inserirDadosIniciaisPrecos')
    .addItem('📧 Enviar Relatório Dia Anterior', 'enviarRelatorioDiaAnterior')
    .addItem('📈 Gerar Relatório Período', 'mostrarDialogoRelatorioPeriodo')
    .addItem('🔍 Testar Sistema', 'testarSistema')
    .addSeparator()
    .addItem('ℹ️ Sobre o Sistema', 'mostrarSobre')
    .addToUi();
}

// ===================================================
// PARTE 12: SISTEMA DE RELATÓRIOS
// ===================================================

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
    
    // Usar email da empresa como remetente
    MailApp.sendEmail({
      to: CONFIG.EMAIL_CONFIG.DESTINATARIOS.join(','),
      subject: `📊 Relatório Diário - ${formatarDataDDMMYYYY(ontem)} - ${CONFIG.NAMES.HOTEL_NAME}`,
      htmlBody: htmlEmail,
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

function criarEmailRelatorioDiario(relatorio, data) {
  const dataFormatada = formatarDataDDMMYYYY(data);
  
  let tabelaTransfers = '';
  relatorio.transfers.forEach(t => {
    const statusCor = t.status === 'Finalizado' ? '#27ae60' : 
                     t.status === 'Confirmado' ? '#3498db' : 
                     t.status === 'Cancelado' ? '#e74c3c' : '#f39c12';
    
    tabelaTransfers += `
      <tr>
        <td style="padding: 12px; border-bottom: 1px solid #ecf0f1;">#${t.id}</td>
        <td style="padding: 12px; border-bottom: 1px solid #ecf0f1;">${t.cliente}</td>
        <td style="padding: 12px; border-bottom: 1px solid #ecf0f1;">${t.origem} → ${t.destino}</td>
        <td style="padding: 12px; border-bottom: 1px solid #ecf0f1; text-align: right; font-weight: bold;">€${t.valorTotal.toFixed(2)}</td>
        <td style="padding: 12px; border-bottom: 1px solid #ecf0f1;">
          <span style="background: ${statusCor}; color: white; padding: 4px 8px; border-radius: 4px; font-size: 12px;">
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
      <style>
        body { font-family: Arial, sans-serif; background: #f4f7f6; margin: 0; padding: 20px; }
        .container { background: #fff; max-width: 700px; margin: 0 auto; padding: 30px; border-radius: 10px; }
        .header { background: linear-gradient(135deg, #2c3e50 0%, #3498db 100%); color: white; padding: 25px; border-radius: 10px 10px 0 0; margin: -30px -30px 30px -30px; }
        h1 { margin: 0 0 10px 0; font-size: 24px; }
        .subtitle { opacity: 0.9; font-size: 14px; }
        .summary-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 20px; margin: 30px 0; }
        .summary-card { background: #f8f9fa; padding: 20px; border-radius: 8px; border-left: 4px solid #3498db; }
        .summary-value { font-size: 28px; font-weight: bold; color: #2c3e50; }
        .summary-label { color: #7f8c8d; font-size: 12px; text-transform: uppercase; margin-top: 5px; }
        .financial-summary { background: linear-gradient(135deg, #27ae60 0%, #2ecc71 100%); color: white; padding: 25px; border-radius: 8px; margin: 30px 0; }
        .table-container { margin: 30px 0; }
        table { width: 100%; border-collapse: collapse; }
        th { background: #2c3e50; color: white; padding: 12px; text-align: left; }
        .footer { text-align: center; color: #95a5a6; font-size: 12px; margin-top: 30px; padding-top: 20px; border-top: 1px solid #ecf0f1; }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="header">
          <h1>📊 Relatório Diário de Transfers</h1>
          <div class="subtitle">${CONFIG.NAMES.HOTEL_NAME} | ${dataFormatada}</div>
        </div>
        
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
        
        <div class="financial-summary">
          <h3 style="margin: 0 0 20px 0;">💰 Distribuição de Valores</h3>
          <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px;">
            <div>
              <div style="font-size: 24px; font-weight: bold;">€${relatorio.resumoFinanceiro.valorTotalHotel.toFixed(2)}</div>
              <div style="opacity: 0.9; font-size: 12px;">HOTEL (30%)</div>
            </div>
            <div>
              <div style="font-size: 24px; font-weight: bold;">€${relatorio.resumoFinanceiro.comissaoTotalRecepcao.toFixed(2)}</div>
              <div style="opacity: 0.9; font-size: 12px;">RECEPÇÃO</div>
            </div>
            <div>
              <div style="font-size: 24px; font-weight: bold;">€${relatorio.resumoFinanceiro.valorTotalHUB.toFixed(2)}</div>
              <div style="opacity: 0.9; font-size: 12px;">HUB TRANSFER</div>
            </div>
          </div>
        </div>
        
        <div class="table-container">
          <h3 style="color: #2c3e50; margin-bottom: 15px;">📋 Detalhamento dos Transfers</h3>
          <table>
            <thead>
              <tr>
                <th>ID</th>
                <th>Cliente</th>
                <th>Rota</th>
                <th style="text-align: right;">Valor</th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody>
              ${tabelaTransfers}
            </tbody>
          </table>
        </div>
        
        <div class="footer">
          <p><strong>${CONFIG.NAMES.SISTEMA_NOME}</strong></p>
          <p>Este é um e-mail automático. Por favor, não responda.</p>
          <p>Enviado por: ${CONFIG.EMAIL_CONFIG.EMAIL_EMPRESA}</p>
        </div>
      </div>
    </body>
    </html>
  `;
}

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
  sheet.getRange(7, 1).setValue('Valor Hotel (30%):');
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
      t.tipoServico,
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
  const htmlEmail = criarEmailRelatorioPeriodo(relatorio, dataInicio, dataFim);
  
  MailApp.sendEmail({
    to: CONFIG.EMAIL_CONFIG.DESTINATARIOS.join(','),
    subject: `📊 Relatório de Período - ${formatarDataDDMMYYYY(dataInicio)} a ${formatarDataDDMMYYYY(dataFim)}`,
    htmlBody: htmlEmail,
    name: CONFIG.NAMES.SISTEMA_NOME,
    replyTo: CONFIG.EMAIL_CONFIG.EMAIL_EMPRESA
  });
  
  return true;
}

function criarEmailRelatorioPeriodo(relatorio, dataInicio, dataFim) {
  return criarEmailRelatorioDiario(relatorio, dataInicio); // Reutiliza o template
}

// Função para configurar trigger automático de relatório diário
function configurarTriggerRelatorioDiario() {
  // Remove triggers existentes
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'enviarRelatorioDiaAnterior') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Cria novo trigger para executar às 9h da manhã
  ScriptApp.newTrigger('enviarRelatorioDiaAnterior')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
    
  logger.success('Trigger de relatório diário configurado para 9h');
}// ===================================================
// SISTEMA DE TRANSFERS IMPIRE MARQUES HOTEL & HUB TRANSFER v1.0
// Sistema Integrado Completo com E-mail Interativo
// Customizado para Impire Marques Hotel
// ===================================================

// ===================================================
// PARTE 1: CONFIGURAÇÃO GLOBAL E CONSTANTES
// ===================================================

const CONFIG = {
  // 📊 Google Sheets - Configuração Principal
  SPREADSHEET_ID: '1ZfG_IXBWMbGQzCmn7nWNzFqUBbGkXCXWSnY7JYziHxI',
  SHEET_NAME: 'Impire MARQUES-HUB',
  PRICING_SHEET_NAME: 'Tabela de Preços',
  
  // 📧 Configuração de E-mail com Sistema Interativo
  EMAIL_CONFIG: {
    DESTINATARIOS: [
      'juniorguitierez@hubtransferencia.com', // Email principal HUB Transfer para testes
      'teste@hubtransferencia.com',           // Email de teste secundário
      'juniorgutierezbega@gmail.com'          // Email de monitoramento atualizado
    ],
    DESTINATARIO: 'juniorguitierez@hubtransferencia.com', // Email principal
    EMAIL_EMPRESA: 'hubtransferencia@gmail.com', // Email oficial da empresa para envios
    ENVIAR_AUTOMATICO: true,
    VERIFICAR_CONFIRMACOES: true,
    INTERVALO_VERIFICACAO: 5, // minutos
    USAR_BOTOES_INTERATIVOS: true,
    TEMPLATE_HTML: true,
    ARQUIVAR_CONFIRMADOS: true,
    ENVIAR_RELATORIO_DIA_ANTERIOR: true // Configuração para relatório do dia anterior
  },
  
  // 📱 Configuração de Telefones e Contatos
  CONTATOS: {
    HUB_PHONE: '+351968698138',        // Manter (proprietário HUB)
    ROBERTA_PHONE: '+351928283652',    // Manter (assistente HUB)
    HOTEL_PHONE: '+351210548700',      // Telefone Impire Marques Hotel
    HOTEL_EMAIL: 'email_pendente@exemplo.com' // Email pendente de confirmação
  },
  
  // 🔗 Integração com APIs Externas
  ZAPI: {
    INSTANCE: '3DC8E250141ED020B95796155CBF9532',
    TOKEN: 'DF93ABBE66F44D82F60EF9FE',
    WEBHOOK_URL: 'https://api.z-api.io/instances/3DC8E250141ED020B95796155CBF9532/token/DF93ABBE66F44D82F60EF9FE/send-text',
    ENABLED: false
  },
  
  // 🔄 Webhooks do Make.com
  MAKE_WEBHOOKS: {
    NEW_TRANSFER: 'https://hook.eu2.make.com/y7ro3pc9we00iums1ffb106pok9icl7m',
    WHATSAPP_RESPONSE: 'https://hook.eu2.make.com/y7ro3pc9we00iums1ffb106pok9icl7m',
    STATUS_UPDATE: 'https://hook.eu2.make.com/y7ro3pc9we00iums1ffb106pok9icl7m',
    ACERTO_CONTAS: 'https://hook.eu2.make.com/y7ro3pc9we00iums1ffb106pok9icl7m',
    ENABLED: false
  },
  
  // 🏨 Identificação do Sistema
  NAMES: {
    HUB_OWNER: 'HUB Transfer',
    ASSISTANT: 'Roberta HUB',
    HOTEL_MANAGER: 'Catarina', // Nome da gerente atualizado
    HOTEL_NAME: 'Impire Marques Hotel', // Nome do hotel atualizado
    SISTEMA_NOME: 'Sistema Impire Marques Hotel-HUB Transfer' // Nome do sistema atualizado
  },
  
  // 💰 Valores Padrão e Configurações Financeiras
  VALORES: {
    HOTEL_PADRAO: 5.00,
    HUB_PADRAO: 20.00,
    PERCENTUAL_HOTEL: 0.30, // 30% para o hotel (atualizado)
    PERCENTUAL_HUB: 0.70,   // 70% para HUB (menos comissão recepção)
    COMISSAO_RECEPCAO_TOUR: 5.00, // €5 para tours
    COMISSAO_RECEPCAO_TRANSFER: 2.00, // €2 para transfers (alterado de 2.50)
    MOEDA: '€',
    FORMATO_MOEDA: '€#,##0.00'
  },
  
  // 🌐 Configurações do Sistema
  SISTEMA: {
    VERSAO: '1.0-Marques-Empire',
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

// Headers da planilha principal - ATUALIZADO COM NOVOS CAMPOS
const HEADERS = [
  'ID',                    // A - Identificador único
  'Cliente',               // B - Nome do cliente
  'Tipo Serviço',          // C - Transfer/Tour/Private (NOVO)
  'Pessoas',               // D - Número de pessoas
  'Bagagens',              // E - Número de bagagens
  'Data',                  // F - Data do transfer
  'Contacto',              // G - Telefone/contato
  'Voo',                   // H - Número do voo
  'Origem',                // I - Local de origem
  'Destino',               // J - Local de destino
  'Hora Pick-up',          // K - Hora de recolha
  'Preço Cliente (€)',     // L - Valor total cobrado
  'Valor Impire Marques Hotel (€)', // M - Comissão do hotel (ATUALIZADO)
  'Valor HUB Transfer (€)', // N - Valor para HUB
  'Comissão Recepção (€)', // O - Comissão da recepção (NOVO)
  'Forma Pagamento',       // P - Método de pagamento
  'Pago Para',             // Q - Quem recebeu
  'Status',                // R - Status do transfer
  'Observações',           // S - Observações gerais
  'Data Criação'           // T - Timestamp de criação
];

// Headers da tabela de preços - ATUALIZADO
const PRICING_HEADERS = [
  'ID',                    // A - ID único do preço
  'Tipo Serviço',          // B - Transfer/Tour Regular/Private Tour (NOVO)
  'Rota',                  // C - Nome da rota
  'Origem',                // D - Ponto de origem
  'Destino',               // E - Ponto de destino
  'Pessoas',               // F - Número de pessoas
  'Bagagens',              // G - Número de bagagens
  'Preço Por Pessoa',      // H - Para tours regulares (NOVO)
  'Preço Por Grupo',       // I - Para private tours (NOVO)
  'Preço Cliente (€)',     // J - Preço total
  'Valor Impire Marques Hotel (€)', // K - Comissão hotel (ATUALIZADO)
  'Valor HUB Transfer (€)', // L - Valor HUB
  'Comissão Recepção (€)', // M - Comissão recepção (NOVO)
  'Ativo',                 // N - Se está ativo
  'Data Criação',          // O - Quando foi criado
  'Observações'            // P - Notas adicionais
];

// ===================================================
// MENSAGENS E TEMPLATES DO SISTEMA
// ===================================================

const MESSAGES = {
  TITULOS: {
    NOVO_TRANSFER: `🚐 NOVO TRANSFER ${CONFIG.NAMES.HOTEL_NAME} & HUB TRANSFER`,
    CONFIRMACAO_NECESSARIA: (id) => `AÇÃO NECESSÁRIA: Novo Transfer #${id}`,
    TRANSFER_CONFIRMADO: (id) => `✅ Transfer #${id} Confirmado`,
    TRANSFER_CANCELADO: (id) => `❌ Transfer #${id} Cancelado`
  },
  
  STATUS_MESSAGES: {
    SOLICITADO: 'Transfer solicitado - aguardando confirmação',
    CONFIRMADO: 'Transfer confirmado - será realizado',
    FINALIZADO: 'Transfer realizado com sucesso',
    CANCELADO: 'Transfer cancelado',
    EM_ANDAMENTO: 'Transfer em andamento',
    PROBLEMA: 'Problema reportado - verificar'
  },
  
  ACOES: {
    CONFIRMADO_POR: (nome) => `✅ Confirmado por ${nome}`,
    CANCELADO_POR: (nome) => `❌ Cancelado por ${nome}`,
    FINALIZADO_POR: (nome) => `✔️ Finalizado por ${nome}`,
    MODIFICADO_POR: (nome) => `📝 Modificado por ${nome}`
  },
  
  ERROS: {
    CAMPO_OBRIGATORIO: (campo) => `Campo obrigatório ausente: ${campo}`,
    VALOR_INVALIDO: (campo) => `Valor inválido para o campo: ${campo}`,
    DATA_INVALIDA: 'Data inválida ou em formato incorreto',
    REGISTRO_NAO_ENCONTRADO: (id) => `Transfer #${id} não encontrado`,
    FALHA_REGISTRO: 'Falha ao registrar o transfer',
    FALHA_EMAIL: 'Falha ao enviar e-mail de notificação'
  },
  
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
  STATUS_COLORS: {
    SOLICITADO: '#f39c12',
    CONFIRMADO: '#27ae60',
    FINALIZADO: '#2c3e50',
    CANCELADO: '#e74c3c',
    EM_ANDAMENTO: '#3498db',
    PROBLEMA: '#e67e22'
  },
  
  HEADER_COLORS: {
    PRINCIPAL: '#2c3e50',
    PRECOS: '#27ae60',
    MENSAL: null
  },
  
  COLUMN_WIDTHS: {
    PRINCIPAL: [60, 150, 100, 60, 60, 100, 120, 80, 200, 200, 90, 100, 140, 120, 80, 100, 100, 100, 180, 140],
    PRECOS: [60, 120, 200, 180, 180, 80, 80, 100, 100, 120, 140, 140, 80, 80, 140, 200]
  },
  
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
  
  FORMATOS: {
    EMAIL: /^[^\s@]+@[^\s@]+\.[^\s@]+$/,
    TELEFONE: /^\+?[\d\s\-().]+$/,
    HORA: /^([01]?[0-9]|2[0-3]):[0-5][0-9]$/,
    DATA_BR: /^\d{2}\/\d{2}\/\d{4}$/,
    DATA_ISO: /^\d{4}-\d{2}-\d{2}$/
  },
  
  VALORES_PERMITIDOS: {
    FORMA_PAGAMENTO: ['Dinheiro', 'Cartão', 'Transferência', 'MB Way', 'Voucher'],
    PAGO_PARA: ['Recepção', 'Motorista', 'Online', 'Hotel', 'HUB'],
    STATUS: Object.keys(MESSAGES.STATUS_MESSAGES),
    TIPO_SERVICO: ['Transfer', 'Tour Regular', 'Private Tour']
  }
};

// ===================================================
// PARTE 2: SISTEMA DE LOGGING
// ===================================================

class Logger {
  constructor() {
    this.logs = [];
    this.enabled = CONFIG.SISTEMA.LOG_DETALHADO;
  }
  
  log(level, message, data = null) {
    if (!this.enabled) return;
    
    const logEntry = {
      timestamp: new Date(),
      level: level,
      message: message,
      data: data
    };
    
    this.logs.push(logEntry);
    console.log(`[${level}] ${message}`, data || '');
  }
  
  debug(message, data) { this.log('DEBUG', message, data); }
  info(message, data) { this.log('INFO', message, data); }
  warn(message, data) { this.log('WARN', message, data); }
  error(message, data) { this.log('ERROR', message, data); }
  success(message, data) { this.log('SUCCESS', message, data); }
}

const logger = new Logger();

// ===================================================
// PARTE 3: FUNÇÕES DE UTILIDADE
// ===================================================

function processarDataSegura(dataInput) {
  logger.debug('Processando data', { input: dataInput });
  
  try {
    if (dataInput instanceof Date && !isNaN(dataInput)) {
      return dataInput;
    }
    
    let data;
    
    if (typeof dataInput === 'string') {
      dataInput = dataInput.trim();
      
      if (dataInput.match(/^\d{4}-\d{2}-\d{2}/)) {
        if (dataInput.includes('T')) {
          data = new Date(dataInput);
        } else {
          data = new Date(dataInput + 'T12:00:00');
        }
      }
      else if (dataInput.match(/^\d{2}\/\d{2}\/\d{4}/)) {
        const [dia, mes, ano] = dataInput.split('/');
        data = new Date(`${ano}-${mes.padStart(2, '0')}-${dia.padStart(2, '0')}T12:00:00`);
      }
      else if (dataInput.match(/^\d{2}-\d{2}-\d{4}/)) {
        const [dia, mes, ano] = dataInput.split('-');
        data = new Date(`${ano}-${mes.padStart(2, '0')}-${dia.padStart(2, '0')}T12:00:00`);
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
    logger.error('Erro ao processar data', { input: dataInput, erro: error.message });
    throw new Error(`Falha ao processar data: ${error.message}`);
  }
}

function formatarDataDDMMYYYY(date) {
  try {
    if (!(date instanceof Date) || isNaN(date)) {
      return 'Data inválida';
    }
    return Utilities.formatDate(date, CONFIG.SISTEMA.TIMEZONE, CONFIG.SISTEMA.DATA_FORMATO);
  } catch (error) {
    logger.error('Erro ao formatar data', { date, erro: error.message });
    return 'Erro na data';
  }
}

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

function validarDados(dados) {
  logger.info('Validando dados de entrada');
  
  const erros = [];
  const dadosValidados = {};
  
  for (const campo of VALIDACOES.CAMPOS_OBRIGATORIOS) {
    if (!dados[campo] || dados[campo] === '') {
      erros.push(MESSAGES.ERROS.CAMPO_OBRIGATORIO(campo));
    }
  }
  
  if (dados.nomeCliente) {
    dadosValidados.nomeCliente = sanitizarTexto(dados.nomeCliente);
  }
  
  if (dados.tipoServico) {
    dadosValidados.tipoServico = dados.tipoServico;
  } else {
    dadosValidados.tipoServico = 'Transfer';
  }
  
  if (dados.numeroPessoas) {
    const pessoas = parseInt(dados.numeroPessoas);
    if (isNaN(pessoas) || pessoas < 1 || pessoas > CONFIG.LIMITES.MAX_PESSOAS) {
      erros.push(`Número de pessoas deve estar entre 1 e ${CONFIG.LIMITES.MAX_PESSOAS}`);
    } else {
      dadosValidados.numeroPessoas = pessoas;
    }
  }
  
  if (dados.numeroBagagens !== undefined) {
    const bagagens = parseInt(dados.numeroBagagens);
    if (isNaN(bagagens) || bagagens < 0 || bagagens > CONFIG.LIMITES.MAX_BAGAGENS) {
      erros.push(`Número de bagagens deve estar entre 0 e ${CONFIG.LIMITES.MAX_BAGAGENS}`);
    } else {
      dadosValidados.numeroBagagens = bagagens;
    }
  }
  
  if (dados.valorTotal) {
    const valor = parseFloat(dados.valorTotal);
    if (isNaN(valor) || valor < CONFIG.LIMITES.MIN_VALOR || valor > CONFIG.LIMITES.MAX_VALOR) {
      erros.push(`Valor deve estar entre ${CONFIG.LIMITES.MIN_VALOR} e ${CONFIG.LIMITES.MAX_VALOR}`);
    } else {
      dadosValidados.valorTotal = valor;
    }
  }
  
  if (dados.data) {
    try {
      dadosValidados.data = processarDataSegura(dados.data);
    } catch (error) {
      erros.push(MESSAGES.ERROS.DATA_INVALIDA);
    }
  }
  
  const outrosCampos = ['horaPickup', 'contacto', 'numeroVoo', 'origem', 'destino', 
                        'modoPagamento', 'pagoParaQuem', 'observacoes', 'tourSelecionado'];
  for (const campo of outrosCampos) {
    if (dados[campo]) {
      dadosValidados[campo] = sanitizarTexto(dados[campo]);
    }
  }
  
  dadosValidados.status = dados.status || 'Solicitado';
  dadosValidados.modoPagamento = dadosValidados.modoPagamento || 'Dinheiro';
  dadosValidados.pagoParaQuem = dadosValidados.pagoParaQuem || 'Recepção';
  dadosValidados.numeroBagagens = dadosValidados.numeroBagagens || 0;
  
  return {
    valido: erros.length === 0,
    erros: erros,
    dados: dadosValidados
  };
}

function sanitizarTexto(texto) {
  if (!CONFIG.SEGURANCA.SANITIZAR_INPUTS) {
    return texto;
  }
  
  let textoSanitizado = String(texto).trim();
  
  if (textoSanitizado.length > CONFIG.SEGURANCA.MAX_CARACTERES_CAMPO) {
    textoSanitizado = textoSanitizado.substring(0, CONFIG.SEGURANCA.MAX_CARACTERES_CAMPO);
  }
  
  if (!CONFIG.SEGURANCA.PERMITIR_HTML_OBSERVACOES) {
    textoSanitizado = textoSanitizado.replace(/<[^>]*>/g, '');
  }
  
  textoSanitizado = textoSanitizado.replace(/[\x00-\x1F\x7F]/g, '');
  textoSanitizado = textoSanitizado.replace(/'/g, "''");
  
  return textoSanitizado;
}

// ===================================================
// PARTE 4: SISTEMA DE CÁLCULO DE PREÇOS ATUALIZADO
// ===================================================

function calcularValores(origem, destino, pessoas, bagagens, precoManual = null, tipoServico = 'Transfer') {
  logger.info('Calculando valores do transfer', {
    origem, destino, pessoas, bagagens, precoManual, tipoServico
  });
  
  try {
    let precoCliente = precoManual ? parseFloat(precoManual) : 25.00;
    let valorHotel, valorHUB, comissaoRecepcao;
    
    // Cálculo baseado no tipo de serviço
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
      // Transfers: 30% hotel + €2,50 recepção
      valorHotel = precoCliente * CONFIG.VALORES.PERCENTUAL_HOTEL;
      comissaoRecepcao = CONFIG.VALORES.COMISSAO_RECEPCAO_TRANSFER;
      valorHUB = precoCliente - valorHotel - comissaoRecepcao;
    }
    
    // Arredondar valores
    valorHotel = Math.round(valorHotel * 100) / 100;
    comissaoRecepcao = Math.round(comissaoRecepcao * 100) / 100;
    valorHUB = Math.round(valorHUB * 100) / 100;
    
    logger.success('Valores calculados', {
      precoCliente,
      valorHotel,
      valorHUB,
      comissaoRecepcao,
      tipoServico
    });
    
    return {
      precoCliente: precoCliente,
      valorHotel: valorHotel,
      valorHUB: valorHUB,
      comissaoRecepcao: comissaoRecepcao,
      fonte: 'calculado',
      tipoServico: tipoServico
    };
    
  } catch (error) {
    logger.error('Erro no cálculo de valores', error);
    return {
      precoCliente: 25.00,
      valorHotel: 7.50,
      valorHUB: 15.00,
      comissaoRecepcao: 2.50,
      fonte: 'padrao'
    };
  }
}

// ===================================================
// PARTE 5: MANIPULAÇÃO DE PLANILHA
// ===================================================

function gerarProximoIdSeguro(sheet) {
  logger.debug('Gerando próximo ID seguro');
  
  try {
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      logger.info('Primeira entrada, retornando ID 1');
      return 1;
    }
    
    const idsRange = sheet.getRange(2, 1, lastRow - 1, 1);
    const ids = idsRange.getValues().flat().map(Number).filter(n => !isNaN(n) && n > 0);
    
    if (ids.length === 0) {
      return 1;
    }
    
    const maxId = Math.max(...ids);
    const novoId = maxId + 1;
    
    logger.debug('Novo ID gerado', { maxId, novoId });
    return novoId;
    
  } catch (error) {
    logger.error('Erro ao gerar ID, usando timestamp', error);
    return Date.now() % 1000000;
  }
}

function encontrarLinhaPorId(sheet, id) {
  logger.debug('Buscando linha por ID', { sheetName: sheet.getName(), id });
  
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return 0;
    
    const idsRange = sheet.getRange(2, 1, lastRow - 1, 1);
    const ids = idsRange.getValues();
    
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) {
        const linha = i + 2;
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

function obterAbaMes(dataTransfer) {
  logger.info('Obtendo aba mensal', { data: dataTransfer });
  
  try {
    const data = processarDataSegura(dataTransfer);
    const mes = String(data.getMonth() + 1).padStart(2, '0');
    const ano = data.getFullYear();
    
    const mesInfo = MESES.find(m => m.abrev === mes);
    if (!mesInfo) {
      throw new Error(`Mês inválido: ${mes}`);
    }
    
    const nomeAba = `${CONFIG.SISTEMA.PREFIXO_MES}${mes}_${mesInfo.nome}_${ano}`;
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    let abaMes = ss.getSheetByName(nomeAba);
    
    if (abaMes) {
      logger.debug('Aba mensal encontrada', { nomeAba });
      return abaMes;
    }
    
    logger.info('Criando nova aba mensal', { nomeAba });
    abaMes = criarAbaMensal(nomeAba, ss, mesInfo);
    
    return abaMes;
    
  } catch (error) {
    logger.error('Erro ao obter aba mensal, usando fallback', error);
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const abaPrincipal = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!abaPrincipal) {
      throw new Error('Falha crítica: aba principal não encontrada');
    }
    
    return abaPrincipal;
  }
}

function criarAbaMensal(nomeAba, ss, mesInfo) {
  logger.info('Criando aba mensal', { nome: nomeAba });
  
  try {
    const abaMes = ss.insertSheet(nomeAba);
    
    abaMes.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    
    const headerRange = abaMes.getRange(1, 1, 1, HEADERS.length);
    headerRange
      .setBackground(mesInfo.cor || STYLES.HEADER_COLORS.PRINCIPAL)
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setFontSize(11)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    
    abaMes.setFrozenRows(1);
    
    STYLES.COLUMN_WIDTHS.PRINCIPAL.forEach((width, index) => {
      abaMes.setColumnWidth(index + 1, width);
    });
    
    logger.success('Aba mensal criada com sucesso', { nome: nomeAba });
    
    return abaMes;
    
  } catch (error) {
    logger.error('Erro ao criar aba mensal', { nome: nomeAba, erro: error.message });
    throw error;
  }
}

// ===================================================
// PARTE 6: SISTEMA DE REGISTRO E ATUALIZAÇÃO
// ===================================================

function executarRegistroDuplo(dadosTransfer) {
  logger.info('Executando registro duplo', { 
    transferId: dadosTransfer[0],
    cliente: dadosTransfer[1] 
  });
  
  try {
    const transferId = dadosTransfer[0];
    const dataTransfer = new Date(dadosTransfer[5]);
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    // Registrar na aba principal
    const abaPrincipal = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!abaPrincipal) {
      throw new Error(`Aba principal '${CONFIG.SHEET_NAME}' não encontrada`);
    }
    
    abaPrincipal.appendRow(dadosTransfer);
    logger.success('Registro na aba principal concluído', { transferId });
    
    // Registrar na aba mensal se configurado
    if (CONFIG.SISTEMA.ORGANIZAR_POR_MES) {
      try {
        const abaMensal = obterAbaMes(dataTransfer);
        if (abaMensal && abaMensal.getName() !== CONFIG.SHEET_NAME) {
          abaMensal.appendRow(dadosTransfer);
          logger.success('Registro na aba mensal concluído', {
            transferId,
            abaMensal: abaMensal.getName()
          });
        }
      } catch (errorAbaMensal) {
        logger.error('Erro ao registrar na aba mensal', errorAbaMensal);
      }
    }
    
    // Enviar e-mail se configurado
    if (CONFIG.EMAIL_CONFIG.ENVIAR_AUTOMATICO) {
      try {
        enviarEmailNovoTransfer(dadosTransfer);
      } catch (errorEmail) {
        logger.error('Erro no envio de e-mail', errorEmail);
      }
    }
    
    return {
      sucesso: true,
      transferId: transferId,
      mensagem: MESSAGES.SUCESSO.REGISTRO_COMPLETO
    };
    
  } catch (error) {
    logger.error('Erro crítico no registro duplo', error);
    return {
      sucesso: false,
      erro: error.message
    };
  }
}

// ===================================================
// PARTE 7: SISTEMA DE E-MAIL
// ===================================================

function enviarEmailNovoTransfer(dadosTransfer) {
  logger.info('Enviando e-mail para novo transfer');
  
  try {
    if (!CONFIG.EMAIL_CONFIG.ENVIAR_AUTOMATICO) {
      logger.info('Envio de e-mail desativado');
      return false;
    }
    
    const destinatarios = CONFIG.EMAIL_CONFIG.DESTINATARIOS.join(',');
    if (!destinatarios) {
      throw new Error('Nenhum destinatário configurado');
    }
    
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
    
    const dataFormatada = formatarDataDDMMYYYY(new Date(transfer.data));
    const assunto = `NOVO: ${transfer.tipoServico} #${transfer.id} - ${transfer.cliente} - ${dataFormatada}`;
    
    const corpoHtml = criarEmailHtml(transfer);
    
    MailApp.sendEmail({
      to: destinatarios,
      subject: assunto,
      htmlBody: corpoHtml,
      name: CONFIG.NAMES.SISTEMA_NOME,
      replyTo: CONFIG.EMAIL_CONFIG.EMAIL_EMPRESA // Usa email da empresa
    });
    
    logger.success('E-mail enviado com sucesso', {
      transferId: transfer.id,
      destinatarios: destinatarios
    });
    
    return true;
    
  } catch (error) {
    logger.error('Erro ao enviar e-mail', error);
    return false;
  }
}

function criarEmailHtml(transfer) {
  const dataFormatada = formatarDataDDMMYYYY(new Date(transfer.data));
  const tipoLabel = transfer.tipoServico === 'tour' ? 'Tour Regular' : 
                   transfer.tipoServico === 'private' ? 'Private Tour' : 
                   'Transfer';
  
  // Assinatura digital do HUB Transfer
  const assinaturaHUB = `
    <table style="width: 100%; max-width: 600px; margin-top: 40px; border-top: 2px solid #ffd700;">
      <tr>
        <td style="padding: 20px 0;">
          <table>
            <tr>
              <td style="vertical-align: top; padding-right: 20px;">
                <img src="https://hubtransferencia.com/logo.png" alt="HUB Transfer" style="width: 150px;">
              </td>
              <td style="vertical-align: top; color: #333;">
                <p style="margin: 0; font-weight: bold; color: #ffd700; font-size: 18px;">Junior Gutierez</p>
                <p style="margin: 5px 0; color: #666;">Diretor</p>
                <p style="margin: 5px 0;">
                  <img src="https://cdn-icons-png.flaticon.com/16/724/724664.png" style="width: 14px; vertical-align: middle;">
                  <a href="https://wa.me/351968698138" style="color: #25d366; text-decoration: none;">+351 968 698 138</a>
                </p>
                <p style="margin: 5px 0;">
                  <img src="https://cdn-icons-png.flaticon.com/16/2504/2504727.png" style="width: 14px; vertical-align: middle;">
                  <a href="https://www.hubtransferencia.com" style="color: #0066cc; text-decoration: none;">www.hubtransferencia.com</a>
                </p>
                <p style="margin: 5px 0;">
                  <img src="https://cdn-icons-png.flaticon.com/16/732/732200.png" style="width: 14px; vertical-align: middle;">
                  <a href="mailto:juniorguitierez@hubtransferencia.com" style="color: #0066cc; text-decoration: none;">juniorguitierez@hubtransferencia.com</a>
                </p>
              </td>
            </tr>
          </table>
          <div style="margin-top: 20px; padding: 15px; background: #000; border-radius: 8px;">
            <p style="color: #ffd700; font-weight: bold; margin: 0 0 10px 0;">HUB TRANSFER</p>
            <p style="color: #fff; font-size: 12px; margin: 0;">TRANSFER AND TOURISM</p>
          </div>
        
        ${assinaturaHUB}
      </div>
    </body>
    </html>
  `;
}

// ===================================================
// PARTE 8: APIs PRINCIPAIS (doGet e doPost)
// ===================================================

function doGet(e) {
  logger.info('Recebendo requisição GET', { 
    parameters: e.parameter,
    queryString: e.queryString 
  });
  
  try {
    const params = e.parameter || {};
    const action = params.action || 'default';
    
    switch (action) {
      case 'test':
        return ContentService
          .createTextOutput(JSON.stringify({
            status: 'success',
            message: 'API funcionando!',
            sistema: CONFIG.NAMES.SISTEMA_NOME,
            versao: CONFIG.SISTEMA.VERSAO
          }))
          .setMimeType(ContentService.MimeType.JSON);
      
      default:
        return ContentService
          .createTextOutput(JSON.stringify({
            status: 'success',
            message: 'Sistema Marques Empire-HUB Transfer funcionando!'
          }))
          .setMimeType(ContentService.MimeType.JSON);
    }
    
  } catch (error) {
    logger.error('Erro no doGet', error);
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  logger.info('Recebendo requisição POST');
  
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Nenhum dado recebido');
    }
    
    const dadosRecebidos = JSON.parse(e.postData.contents);
    logger.debug('Dados recebidos', dadosRecebidos);
    
    // Processar ação especial se houver
    if (dadosRecebidos.action) {
      return processarAcaoEspecial(dadosRecebidos);
    }
    
    // Processar novo transfer
    return processarNovoTransfer(dadosRecebidos);
    
  } catch (error) {
    logger.error('Erro no doPost', error);
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function processarNovoTransfer(dadosRecebidos) {
  logger.info('Processando novo transfer');
  
  const validacao = validarDados(dadosRecebidos);
  if (!validacao.valido) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: 'Dados inválidos',
        erros: validacao.erros
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  const dados = validacao.dados;
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) {
      throw new Error(`Aba '${CONFIG.SHEET_NAME}' não encontrada`);
    }
    
    const transferId = dados.id || gerarProximoIdSeguro(sheet);
    
    // Calcular valores com base no tipo de serviço
    const valores = calcularValores(
      dados.origem,
      dados.destino,
      dados.numeroPessoas,
      dados.numeroBagagens,
      dados.valorTotal,
      dados.tipoServico
    );
    
    // Montar dados do transfer com a nova estrutura
    const dadosTransfer = [
      transferId,                          // ID
      dados.nomeCliente,                   // Cliente
      dados.tipoServico || 'Transfer',     // Tipo Serviço
      dados.numeroPessoas,                 // Pessoas
      dados.numeroBagagens,                // Bagagens
      dados.data,                          // Data
      dados.contacto,                      // Contacto
      dados.numeroVoo || '',               // Voo
      dados.origem,                        // Origem
      dados.destino,                       // Destino
      dados.horaPickup,                    // Hora Pick-up
      valores.precoCliente,                // Preço Cliente
      valores.valorHotel,                  // Valor Hotel Marques Empire
      valores.valorHUB,                    // Valor HUB Transfer
      valores.comissaoRecepcao,            // Comissão Recepção
      dados.modoPagamento,                 // Forma Pagamento
      dados.pagoParaQuem,                  // Pago Para
      dados.status,                        // Status
      dados.observacoes || '',             // Observações
      new Date()                           // Data Criação
    ];
    
    const resultadoRegistro = executarRegistroDuplo(dadosTransfer);
    
    const resposta = {
      status: resultadoRegistro.sucesso ? 'success' : 'error',
      message: resultadoRegistro.mensagem || resultadoRegistro.erro,
      transfer: {
        id: transferId,
        cliente: dados.nomeCliente,
        tipoServico: dados.tipoServico,
        data: formatarDataDDMMYYYY(dados.data),
        rota: `${dados.origem} → ${dados.destino}`,
        valor: valores.precoCliente
      },
      valores: valores
    };
    
    return ContentService
      .createTextOutput(JSON.stringify(resposta))
      .setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    logger.error('Erro ao processar transfer', error);
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

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
        
      default:
        throw new Error(`Ação desconhecida: ${dados.action}`);
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({
        status: resultado.sucesso ? 'success' : 'error',
        action: dados.action,
        resultado: resultado
      }))
      .setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    logger.error('Erro na ação especial', error);
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        action: dados.action,
        message: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ===================================================
// PARTE 9: FUNÇÕES DE MANUTENÇÃO
// ===================================================

function limparDadosCompleto() {
  logger.warn('LIMPANDO TODOS OS DADOS DO SISTEMA');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheets = ss.getSheets();
    let registrosRemovidos = 0;
    let abasProcessadas = 0;
    
    sheets.forEach(sheet => {
      const nome = sheet.getName();
      
      if (nome === CONFIG.SHEET_NAME || 
          nome === CONFIG.PRICING_SHEET_NAME ||
          nome.startsWith(CONFIG.SISTEMA.PREFIXO_MES)) {
        
        const lastRow = sheet.getLastRow();
        
        if (lastRow > 1) {
          const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
          sheet.clear();
          sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
          
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

function limparDadosTeste() {
  logger.info('Limpando dados de teste');
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheets = ss.getSheets();
    let totalRemovidos = 0;
    const detalhes = [];
    
    const palavrasTeste = ['teste', 'test', 'demo', 'exemplo', 'sample'];
    
    sheets.forEach(sheet => {
      const nome = sheet.getName();
      
      if (nome === CONFIG.SHEET_NAME || nome.startsWith(CONFIG.SISTEMA.PREFIXO_MES)) {
        const lastRow = sheet.getLastRow();
        if (lastRow <= 1) return;
        
        const dados = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
        const linhasParaRemover = [];
        
        dados.forEach((row, index) => {
          const id = row[0];
          const cliente = String(row[1]).toLowerCase();
          const observacoes = String(row[18]).toLowerCase();
          
          let ehTeste = false;
          
          if (id === 999 || id === 9999) {
            ehTeste = true;
          }
          
          palavrasTeste.forEach(palavra => {
            if (cliente.includes(palavra) || observacoes.includes(palavra)) {
              ehTeste = true;
            }
          });
          
          if (ehTeste) {
            linhasParaRemover.push(index + 2);
          }
        });
        
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

// ===================================================
// PARTE 10: DADOS INICIAIS DE PREÇOS
// ===================================================

function inserirDadosIniciaisPrecos() {
  logger.info('Inserindo dados iniciais de preços');
  
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.PRICING_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.PRICING_SHEET_NAME);
    sheet.getRange(1, 1, 1, PRICING_HEADERS.length).setValues([PRICING_HEADERS]);
  }
  
  const precos = [
    // TOURS REGULARES (Por Pessoa)
    [1, 'Tour Regular', 'Tour Sintra e Cascais', 'Hotel Marques Empire', 'Sintra/Cascais Tour', 1, 0, 67, 0, 67, 20.10, 41.90, 5.00, 'Sim', new Date(), 'Tour 8h - 30% hotel + €5 recepção'],
    [2, 'Tour Regular', 'Tour Fátima e Óbidos', 'Hotel Marques Empire', 'Fátima/Óbidos Tour', 1, 0, 83, 0, 83, 24.90, 53.10, 5.00, 'Sim', new Date(), 'Tour 8h - 30% hotel + €5 recepção'],
    [3, 'Tour Regular', 'Combo 2 Dias Tours', 'Hotel Marques Empire', 'Combo Tours 2 Dias', 1, 0, 138, 0, 138, 41.40, 91.60, 5.00, 'Sim', new Date(), 'Tour 16h - 30% hotel + €5 recepção'],
    
    // PRIVATE TOURS (Por Grupo 1-3 pessoas)
    [4, 'Private Tour', 'Private Palaces Sintra', 'Hotel Marques Empire', 'Private Sintra Palaces', 3, 0, 0, 347, 347, 104.10, 237.90, 5.00, 'Sim', new Date(), 'Private 1-3 pax - 30% hotel + €5 recepção'],
    [5, 'Private Tour', 'Private Évora City', 'Hotel Marques Empire', 'Private Évora', 3, 0, 0, 347, 347, 104.10, 237.90, 5.00, 'Sim', new Date(), 'Private 1-3 pax - 30% hotel + €5 recepção'],
    [6, 'Private Tour', 'Private Templars', 'Hotel Marques Empire', 'Private Templars', 3, 0, 0, 347, 347, 104.10, 237.90, 5.00, 'Sim', new Date(), 'Private 1-3 pax - 30% hotel + €5 recepção'],
    [7, 'Private Tour', 'Private Arrábida', 'Hotel Marques Empire', 'Private Arrábida', 3, 0, 0, 347, 347, 104.10, 237.90, 5.00, 'Sim', new Date(), 'Private 1-3 pax - 30% hotel + €5 recepção'],
    
    // PRIVATE TOURS (Por Grupo 4-8 pessoas)
    [8, 'Private Tour', 'Private Palaces Sintra', 'Hotel Marques Empire', 'Private Sintra Palaces', 8, 0, 0, 492, 492, 147.60, 339.40, 5.00, 'Sim', new Date(), 'Private 4-8 pax - 30% hotel + €5 recepção'],
    [9, 'Private Tour', 'Private Évora City', 'Hotel Marques Empire', 'Private Évora', 8, 0, 0, 492, 492, 147.60, 339.40, 5.00, 'Sim', new Date(), 'Private 4-8 pax - 30% hotel + €5 recepção'],
    [10, 'Private Tour', 'Private Templars', 'Hotel Marques Empire', 'Private Templars', 8, 0, 0, 492, 492, 147.60, 339.40, 5.00, 'Sim', new Date(), 'Private 4-8 pax - 30% hotel + €5 recepção'],
    [11, 'Private Tour', 'Private Arrábida', 'Hotel Marques Empire', 'Private Arrábida', 8, 0, 0, 492, 492, 147.60, 339.40, 5.00, 'Sim', new Date(), 'Private 4-8 pax - 30% hotel + €5 recepção'],
    
    // TRANSFERS POPULARES
    [12, 'Transfer', 'Airport → Hotel', 'Lisbon Airport', 'Hotel Marques Empire', 4, 0, 0, 0, 25, 7.50, 15.50, 2.00, 'Sim', new Date(), 'Transfer até 4 pessoas'],
    [13, 'Transfer', 'Hotel → Airport', 'Hotel Marques Empire', 'Lisbon Airport', 4, 0, 0, 0, 25, 7.50, 15.50, 2.00, 'Sim', new Date(), 'Transfer até 4 pessoas'],
    [14, 'Transfer', 'Belém Tower', 'Hotel Marques Empire', 'Torre de Belém', 4, 0, 0, 0, 27, 8.10, 16.90, 2.00, 'Sim', new Date(), 'Transfer até 4 pessoas'],
    [15, 'Transfer', 'Lisbon Oceanarium', 'Hotel Marques Empire', 'Oceanário de Lisboa', 4, 0, 0, 0, 28, 8.40, 17.60, 2.00, 'Sim', new Date(), 'Transfer até 4 pessoas'],
    [16, 'Transfer', 'Pena Palace', 'Hotel Marques Empire', 'Palácio da Pena (Sintra)', 4, 0, 0, 0, 48, 14.40, 31.60, 2.00, 'Sim', new Date(), 'Transfer até 4 pessoas'],
    [17, 'Transfer', 'Quinta da Regaleira', 'Hotel Marques Empire', 'Quinta da Regaleira (Sintra)', 4, 0, 0, 0, 48, 14.40, 31.60, 2.00, 'Sim', new Date(), 'Transfer até 4 pessoas']
  ];
  
  try {
    if (sheet.getLastRow() > 1) {
      sheet.deleteRows(2, sheet.getLastRow() - 1);
    }
    
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
// PARTE 11: MENU DO SISTEMA
// ===================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🚐 Sistema Marques Empire')
    .addItem('⚙️ Configurar Sistema', 'configurarSistema')
    .addSeparator()
    .addItem('📊 Inserir Preços Iniciais', 'inserirDadosIniciaisPrecos')
    .addItem('🔍 Testar Sistema', 'testarSistema')
    .addSeparator()
    .addItem('ℹ️ Sobre o Sistema', 'mostrarSobre')
    .addToUi();
}

function configurarSistema() {
  const ui = SpreadsheetApp.getUi();
  
  ui.alert(
    '⚙️ Configuração do Sistema',
    'Esta ação irá:\n\n' +
    '1. Criar/verificar todas as abas necessárias\n' +
    '2. Aplicar formatações\n' +
    '3. Inserir dados iniciais de preços\n\n' +
    'Deseja continuar?',
    ui.ButtonSet.YES_NO
  );
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    // Verificar/criar aba principal
    let abaPrincipal = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!abaPrincipal) {
      abaPrincipal = ss.insertSheet(CONFIG.SHEET_NAME);
      abaPrincipal.appendRow(HEADERS);
    }
    
    // Verificar/criar tabela de preços
    let tabelaPrecos = ss.getSheetByName(CONFIG.PRICING_SHEET_NAME);
    if (!tabelaPrecos) {
      tabelaPrecos = ss.insertSheet(CONFIG.PRICING_SHEET_NAME);
      tabelaPrecos.appendRow(PRICING_HEADERS);
      inserirDadosIniciaisPrecos();
    }
    
    ui.alert(
      '✅ Sistema Configurado!',
      'Sistema Hotel Marques Empire & HUB Transfer configurado com sucesso!',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    ui.alert('❌ Erro', 'Erro na configuração:\n' + error.toString(), ui.ButtonSet.OK);
  }
}

function testarSistema() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) {
      ui.alert('❌ Erro', 'Aba principal não encontrada. Execute "Configurar Sistema" primeiro.', ui.ButtonSet.OK);
      return;
    }
    
    ui.alert(
      '✅ Sistema OK!',
      'Sistema funcionando corretamente!\n\n' +
      `Versão: ${CONFIG.SISTEMA.VERSAO}\n` +
      `Hotel: ${CONFIG.NAMES.HOTEL_NAME}\n` +
      `Aba Principal: ${CONFIG.SHEET_NAME}`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    ui.alert('❌ Erro', error.toString(), ui.ButtonSet.OK);
  }
}

function mostrarSobre() {
  const ui = SpreadsheetApp.getUi();
  
  const mensagem = `
🚐 SISTEMA DE TRANSFERS MARQUES EMPIRE-HUB v1.0

Sistema customizado para gestão de transfers entre Hotel Marques Empire e HUB Transfer.

📋 CARACTERÍSTICAS:
- Cálculo automático de comissões (30% hotel)
- Comissão recepção: €5 (tours) / €2,50 (transfers)
- Suporte para Tours Regulares e Private Tours
- Registro duplo automático
- Sistema de e-mail automático

🏢 DESENVOLVIDO PARA:
${CONFIG.NAMES.HOTEL_NAME} & ${CONFIG.NAMES.HUB_OWNER}

⚙️ VERSÃO: ${CONFIG.SISTEMA.VERSAO}
📅 DATA: ${formatarDataDDMMYYYY(new Date())}
`;
  
  ui.alert('ℹ️ Sobre o Sistema', mensagem, ui.ButtonSet.OK);
}

// ===================================================
// FIM DO CÓDIGO - SISTEMA CUSTOMIZADO HOTEL MARQUES EMPIRE
// ===================================================
        </td>
      </tr>
    </table>
  `;
  
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <style>
        body { font-family: Arial, sans-serif; background: #f4f7f6; margin: 0; padding: 20px; }
        .container { background: #fff; max-width: 600px; margin: 0 auto; padding: 30px; border-radius: 10px; }
        h1 { color: #2c3e50; border-bottom: 3px solid #ffd700; padding-bottom: 10px; }
        .info-row { display: flex; justify-content: space-between; padding: 10px 0; border-bottom: 1px solid #ecf0f1; }
        .label { font-weight: bold; color: #7f8c8d; }
        .value { color: #2c3e50; }
        .pricing { background: #ecf0f1; padding: 20px; border-radius: 8px; margin: 20px 0; }
        .footer { text-align: center; color: #95a5a6; font-size: 12px; margin-top: 30px; padding-top: 20px; border-top: 1px solid #ecf0f1; }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>🚐 Novo ${tipoLabel} Solicitado</h1>
        
        <div class="info-row">
          <span class="label">ID:</span>
          <span class="value">#${transfer.id}</span>
        </div>
        
        <div class="info-row">
          <span class="label">Cliente:</span>
          <span class="value">${transfer.cliente}</span>
        </div>
        
        <div class="info-row">
          <span class="label">Tipo de Serviço:</span>
          <span class="value">${tipoLabel}</span>
        </div>
        
        <div class="info-row">
          <span class="label">Data:</span>
          <span class="value">${dataFormatada} às ${transfer.horaPickup}</span>
        </div>
        
        <div class="info-row">
          <span class="label">Rota:</span>
          <span class="value">${transfer.origem} → ${transfer.destino}</span>
        </div>
        
        <div class="info-row">
          <span class="label">Passageiros:</span>
          <span class="value">${transfer.pessoas} pessoa(s)</span>
        </div>
        
        <div class="info-row">
          <span class="label">Contacto:</span>
          <span class="value">${transfer.contacto}</span>
        </div>
        
        <div class="pricing">
          <h3>💰 Valores</h3>
          <div class="info-row">
            <span class="label">Valor Total:</span>
            <span class="value">€${transfer.precoCliente.toFixed(2)}</span>
          </div>
          <div class="info-row">
            <span class="label">Impire Marques Hotel:</span>
            <span class="value">€${transfer.valorHotel.toFixed(2)}</span>
          </div>
          <div class="info-row">
            <span class="label">HUB Transfer:</span>
            <span class="value">€${transfer.valorHUB.toFixed(2)}</span>
          </div>
          <div class="info-row">
            <span class="label">Comissão Recepção:</span>
            <span class="value">€${transfer.comissaoRecepcao.toFixed(2)}</span>
          </div>
        </div>
