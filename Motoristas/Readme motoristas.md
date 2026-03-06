# ✈ HUB Transfer — Operations

**Sistema operacional de gestão de transfers aeroporto/hotel em Lisboa.**

---

## 🔗 Acesso Rápido

| Painel | Link |
|--------|------|
| **Painel Principal (Admin)** | [▶ Abrir](https://hubtransfer.github.io/hub-transfer/Sistema%20HUB%20Central/Hub_transfer_ops.html) |

### 🚗 Páginas dos Motoristas

| Motorista | Link |
|-----------|------|
| Victor Santiago | [▶ Abrir](https://hubtransfer.github.io/hub-transfer/Motoristas/Victor%20Santiago.html) |
| Filipe Ventura | [▶ Abrir](https://hubtransfer.github.io/hub-transfer/Motoristas/Filipe%20ventura.html) |
| Karolini Duarte | [▶ Abrir](https://hubtransfer.github.io/hub-transfer/Motoristas/Karolini%20duarte.html) |
| Igor Beltrame | [▶ Abrir](https://hubtransfer.github.io/hub-transfer/Motoristas/Igor%20Beltrame.html) |

---

## 📋 Funcionalidades

### Painel Principal
- 📅 Viagens sincronizadas com a HUB Central (Google Sheets)
- 🔄 Navegação por data — Ontem / Hoje / Amanhã / Calendário
- 🚗 Filtro por motorista — ver viagens de cada um individualmente
- ✈ Bloco de voo com aeroporto de origem, país, bandeira e tempo estimado de saída
- 🟢 Barra de progresso do voo em tempo real
- 🚗 Rota visual A→B para recolhas / 🧭 Bússola para tours
- 🏨 Nome do hotel separado do endereço para melhor navegação GPS
- 👤 Atribuição de motorista por viagem
- 💬 Template WhatsApp automático no idioma do cliente (PT/EN/FR/ES/IT/DE)
- 📱 SMS com template como fallback se cliente não tiver WhatsApp
- 📞 Telefone do cliente visível e copiável (com +)
- 🏷 Badge de idioma detectado por DDI do telefone
- 💰 Comissão do motorista automática (€9 Talixo/WT · €10 outros)
- 📊 Resumo de pagamento por motorista no topo
- ✅ Botão "Dar Baixa" que marca viagem como concluída na planilha
- 📛 Placa de nome fullscreen ao clicar no cliente (modo aeroporto)

### Páginas dos Motoristas
- 🔒 Acesso isolado — cada motorista só vê as suas viagens
- 📅 Navegação por data com histórico
- 💰 Resumo do valor a receber por dia
- 📋 Mesmo visual e informação do painel principal
- 💬 WhatsApp/SMS com template personalizado com o nome do motorista
- ✅ Dar baixa nas viagens
- 📛 Placa de nome para receção no aeroporto
- 🔄 Auto-refresh a cada 2 minutos

---

## 🏗 Arquitectura

```
GitHub Pages (Frontend)
  └── Sistema HUB Central/
        └── Hub_transfer_ops.html ─── Painel Admin
  └── Motoristas/
        ├── Victor Santiago.html
        ├── Filipe ventura.html
        ├── Karolini duarte.html
        └── Igor Beltrame.html

Google Apps Script (Backend)
  └── CONFIG_E_CONSTANTES.gs
        ├── doGet() ─── API endpoints
        ├── getViagens(data) ─── Viagens por data
        ├── getMotoristas() ─── Lista de motoristas
        └── completarViagem() ─── Dar baixa

Google Sheets (Base de dados)
  └── HUB-Central / Transfers_MM_Mês_YYYY
        ├── Colunas A-E: ID, Cliente, Tipo, Pax, Bags
        ├── Colunas F-K: Data, Telefone, Voo, Origem, Destino, Hora
        ├── Coluna R: Status (CONCLUIDA)
        ├── Coluna V: Idioma
        ├── Coluna AA: Motorista Atribuído
        ├── Coluna CF: Comissão Motorista (€9/€10)
        ├── Coluna CG: Plataforma (Talixo/WT/Empire)
        └── Coluna CH: Timestamp Conclusão
```

---

## 💰 Regras de Comissão

| Plataforma | Valor por viagem |
|------------|-----------------|
| Talixo | €9 |
| World Transfer / WT Driver | €9 |
| Empire Lisbon Hotel | €10 |
| Empire Marques Hotel | €10 |
| Gota d'água | €10 |
| Hotel Lioz | €10 |
| Directo / Outros | €10 |

---

## 🌍 Base de Aeroportos (379)

| Categoria | Tempo estimado saída | Exemplos |
|-----------|---------------------|----------|
| 🟢 Schengen | ~15-25min | Frankfurt, Paris, Barcelona, Roma |
| 🟠 Reino Unido | ~40-60min | Heathrow, Gatwick, Manchester |
| 🟡 Europa não-Schengen | ~30-45min | Dublin, Istambul, Bucareste |
| 🔵 América do Norte | ~45-60min | JFK, Miami, Toronto |
| 🟤 África | ~45-60min | Casablanca, Luanda, Cabo Verde |
| ⚫ Médio Oriente | ~45-60min | Dubai, Doha, Tel Aviv |
| 🔴 Brasil | ~60-90min | São Paulo, Rio, Brasília |
| 🟣 América Latina | ~50-75min | Buenos Aires, Bogotá, México |
| 🔵 Ásia/Oceania | ~60-90min | Tóquio, Singapura, Sydney |

---

## 🔧 Manutenção

**Adicionar novo motorista:**
1. Duplicar qualquer ficheiro da pasta `Motoristas/`
2. Mudar `const MOTORISTA = 'Nome Completo';` e o `<title>`
3. Commit no GitHub

**Mudar preço de comissão:**
- Editar directamente na coluna CF da planilha (override manual)
- Ou editar `calcDriverPrice()` no HTML para mudar a regra global

**Mudar templates WhatsApp:**
- Editar o objecto `TEMPLATES` no HTML (secção CHEGADA/RECOLHA/TOUR por idioma)

---

> **HUB Transfer** · Jornadas e Possibilidades Unipessoal LDA · RNAVT 12529
