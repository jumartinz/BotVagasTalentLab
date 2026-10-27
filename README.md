
# Bot Vagas Talent Lab - FIAP

Programa em Python desenvolvido para automação, busca de vagas de tecnologia em fontes públicas, padronização de acordo com as regras de negócio da FIAP e exportação em lote para planilha Excel (`.xlsx`).

## 🚀 Funcionalidades
- **Interface Gráfica Responsiva:** Desenvolvida para facilitar a operação intuitiva.
- **Padronização Automática:** Aplicação das regras de títulos (limite de 58 caracteres, tags para grupos minoritários), descrição (até 2000 caracteres) e validações administrativas exigidas pelo Talent Lab da FIAP.
- **Exportação em Lote:** Geração de planilhas formatadas prontas para upload.

## 🛠️ Tecnologias Utilizadas
- Python 3.x
- `customtkinter` (Interface Gráfica)
- `pandas` & `openpyxl` (Manipulação e exportação para Excel)
- `beautifulsoup4` / `requests` (Busca e extração de vagas)

## 📌 Como Executar
1. Instale as dependências:
   ```bash
   pip install pandas beautifulsoup4 customtkinter openpyxl

2. Execute a interface principal:
   ```bash
   pip install -r requirements.txt
   python fiap_vagas_tech.py              # abre a interface
   python fiap_vagas_tech.py --selftest   # teste offline: gera um Excel com vagas de exemplo
