# 🤖 Bot Vagas Talent Lab

> Automação para coleta, limpeza, classificação e exportação de vagas de emprego.

[![Status](https://img.shields.io/badge/status-ativo-brightgreen)](#)
[![Python Version](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

---

## 📌 Sobre o Projeto

O **Bot Vagas Talent Lab** é um script desenvolvido em **Python** para automatizar a busca, tratamento e organização de oportunidades de emprego de forma estruturada.

---

## ✨ Principais Funções

### 1. Coleta Automatizada de Vagas
* Conecta-se diretamente às APIs públicas de 3 plataformas:
  * **Gupy:** Vagas nacionais.
  * **Remotive:** Vagas remotas internacionais.
  * **Remote OK:** Vagas remotas internacionais.
* Não faz raspagem direta de HTML; recorre aos *endpoints* e dados estruturados em JSON das plataformas[cite: 1].

### 2. Duplicação e Limpeza de Dados
* Remove vagas repetidas com base no link ou na combinação de título e nome da empresa[cite: 1].
* Converte e remove tags HTML das descrições[cite: 1].

### 3. Classificação por Palavras-Chave
O script analisa o título e a descrição de cada vaga para caracterizá-la em 5 eixos[cite: 1]:
* **Tipo de contratação** *(Padrão Symplicity)*[cite: 1]
* **Modelo de trabalho** *(Padrão Symplicity)*[cite: 1]
* **Nível da vaga** *(Padrão Symplicity)*[cite: 1]
* **Cluster** *(Padrão Symplicity)*[cite: 1]
* **Vaga afirmativa** *(Diferencial)*[cite: 1]

### 4. Exportação para Excel (`vagas.xlsx`)
* **Aba "vagas":** Lista detalhada com colunas, resumo e links clicáveis[cite: 2].
* **Aba "resumo":** Cria um painel gerencial dinâmico usando fórmulas do próprio Excel[cite: 2].

---

## 🛠️ Tecnologias Utilizadas

* **Linguagem:** [Python 3.10+](https://www.python.org/)

---

## 🚀 Começando

### Pré-requisitos

* Python 3.10 ou superior instalado.
* Git instalado.

### Instalação

1. **Clone o repositório:**
   ```bash
   git clone [https://github.com/seu-usuario/BotVagasTalentLab.git](https://github.com/seu-usuario/BotVagasTalentLab.git)
   cd BotVagasTalentLab
2. **Para rodar o projeto:**
   ```bash
   python C:/Users/anna.martins/Desktop/BotVagasTalentLab/AnalisadorDeVagas.py --saida C:/Users/anna.martins/Desktop/BotVagasTalentLab/vagas.xlsx
