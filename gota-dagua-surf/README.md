# Sistema de Transfers - Gota d'água Surf Camps & HUB Transfer

## 🏨 Sistema de Gestão de Transfers

Sistema desenvolvido para gestão de transfers entre **Gota d'água Surf Camps** e **HUB Transfer**.

### 📊 Características Principais

- **Valor padrão**: €35 por transfer (Aeroporto ↔ Gota d'água)
- **Distribuição**: €6.50 Gota d'água + €28.50 HUB Transfer
- **Gestão completa** de transfers, tours regulares e private tours
- **E-mails interativos** com botões de confirmação
- **Relatórios automáticos** diários e por período
- **Abas mensais** organizadas automaticamente
- **Sistema de preços híbrido** (fixos + personalizados)

### 🔗 Links de Acesso

- **Planilha Google Sheets**: [Abrir Planilha](https://docs.google.com/spreadsheets/d/1Zo0er2QaKszT3sYV8F0MZ50r3Pgj5gi50stqNrAnizY/)
- **API Sistema**: Disponível via Google Apps Script
- **Web App**: https://script.google.com/macros/s/AKfycbw6N5Ia6ev4OrI7JlI-oiVzqUPHwVFZ9RS7ATouMMDhJzpkUmjBRZ5F0YLyvf-55WdD0A/exec

### 👥 Contatos

- **Gerente Hotel**: Luiz Felipe
- **Telefone Hotel**: +351 939 591 704
- **Email Principal**: robertavgutierez@gmail.com
- **Email Monitoramento**: juniorgutierezbega@gmail.com
- **HUB Transfer**: +351 968 698 138
- **Assistente HUB**: Roberta (+351 928 283 652)

### 💰 Sistema de Valores

#### 🚐 **Transfers Padrão**
- **Aeroporto ↔ Gota d'água**: €35 total
  - Gota d'água: €6.50
  - HUB Transfer: €28.50

#### 🎯 **Tours Regulares**
- **Por pessoa**: €67 (Sintra, Fátima, etc.)
- **Distribuição proporcional** conforme tabela

#### 🏆 **Private Tours**
- **Até 3 pessoas**: €347
- **Até 6 pessoas**: €492
- **Rotas personalizadas disponíveis**

### 🛠️ Funcionalidades Técnicas

- **Sistema v4.0** - Gota d'água Hub Transfer Integrado
- **Triggers automáticos** para verificação de e-mails (5 min)
- **Relatório diário** enviado às 9h
- **Manutenção automática** às 2h
- **Backup automático** configurável
- **Interceptadores de e-mail** para debugging
- **Validações de dados** com sanitização
- **Suporte a múltiplos formatos de data**

### 📧 Sistema de E-mail

- **E-mails interativos** com botões HTML
- **Confirmação automática** via clique ou resposta
- **Templates profissionais** responsivos
- **Assinatura digital** com logo HUB Transfer
- **Notificações automáticas** de status

### 📊 Relatórios Disponíveis

- **📈 Relatório Geral**: Visão completa do período
- **💰 Acerto de Contas**: Valores específicos para o hotel
- **🚐 Por Tipo de Serviço**: Transfer/Tour/Private separados
- **🏆 Top Clientes**: Ranking de clientes frequentes
- **🗺️ Top Rotas**: Rotas mais utilizadas
- **📤 Exportação CSV**: Download de dados

### 🔧 Configuração e Instalação

#### **Primeira Instalação:**
```javascript
// Execute no Google Apps Script:
instalarSistemaCompleto()
```

#### **Configuração Completa:**
```javascript
// Após instalação, execute:
configurarSistema()
instalarSistemaRelatoriosCompleto()
configurarTriggersEmailMenu()
```

#### **Testes:**
```javascript
// Testar sistema:
testarSistemaCompleto()
testarEnvioEmailInterativo()
testarSistemaPrecoHibrido()
```

### 🎯 Menu do Sistema

Acesse via **Google Sheets** → **Menu "🚐 Sistema Gota d água"**:

- ⚙️ **Configurar Sistema**
- 📊 **Relatórios e Acerto de Contas**
- 📧 **E-mail Interativo**
- 💰 **Gestão de Preços**
- 📅 **Abas Mensais**
- 🔧 **Manutenção**

### 🔐 Estrutura de Dados

#### **Planilha Principal**: `Gota d agua-HUB`
- 20 colunas de dados
- Validações automáticas
- Formatação condicional por status

#### **Tabela de Preços**: `Tabela de Preços`
- Preços fixos e personalizados
- Sistema de matching inteligente
- Suporte a múltiplos tipos de serviço

#### **Abas Mensais**: `Transfers_MM_MêsName_YYYY`
- Organização automática por mês
- Sincronização com aba principal
- Formatação individual

---

**Versão**: 4.0-Gota-d-agua-Hub-Transfer-Integrado  
**Data**: Agosto 2025  
**Desenvolvido por**: Claude 4 Sonnet + Google Apps Script  
**Timezone**: Europe/Lisbon  

> **🎯 Sistema 100% funcional** com valores fixos, e-mails interativos e relatórios automáticos para otimizar a gestão de transfers entre Gota d'água Surf Camps e HUB Transfer.
