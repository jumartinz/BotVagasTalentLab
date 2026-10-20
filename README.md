# 🤖 BotVagasTalentLab

> Automação para monitoramento, coleta e notificação de vagas de emprego da plataforma TalentLab.

[![Status](https://img.shields.io/badge/status-ativo-brightgreen)](#)
[![Python Version](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

---

## 📌 Sobre o Projeto

O **BotVagasTalentLab** é um bot desenvolvido para automatizar a busca por novas oportunidades de emprego anunciadas na plataforma **TalentLab**. 

Ele realiza a varredura periódica de vagas conforme filtros configurados (área de atuação, senioridade, localização/remoto) e envia alertas imediatos para canais de comunicação (como Telegram, Discord ou E-mail), permitindo que você se candidate rapidamente.

---

## ✨ Funcionalidades

- 🔍 **Varredura Automatizada:** Monitora novas vagas abertas na plataforma com intervalos configuráveis.
- 🎯 **Filtro Customizado:** Suporte a filtros por palavras-chave, stack tecnológica e modalidade de trabalho (Remoto, Híbrido, Presencial).
- 🔔 **Notificações em Tempo Real:** Envio de alertas detalhados (título da vaga, empresa, descrição curta e link direto).
- 🚫 **Deduplicação Inteligente:** Evita notificações repetidas de vagas já processadas.

---

## 🛠️ Tecnologias Utilizadas

- **Linguagem:** [Python 3.10+](https://www.python.org/)
- **Scraping / Requisições:** `requests` / `httpx` / `BeautifulSoup4` *(ou `Playwright` / `Selenium`)*
- **Notificações:** API do Telegram / Webhooks do Discord
- **Armazenamento:** `SQLite` *(para cache e controle de vagas enviadas)*
- **Agendamento:** `APScheduler` / `schedule` / GitHub Actions

---

## 🚀 Começando

### Pré-requisitos

- Python 3.10 ou superior instalado.
- Git instalado.
- Token/Webhook do canal de notificação (ex.: Bot no Telegram ou Webhook no Discord).

### Instalação

1. **Clone o repositório:**
   ```bash
   git clone [https://github.com/seu-usuario/BotVagasTalentLab.git](https://github.com/seu-usuario/BotVagasTalentLab.git)
   cd BotVagasTalentLab
