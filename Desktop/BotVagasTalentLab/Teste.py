# Importando a biblioteca necessária para ler a planilha

from openpyxl import load_workbook
print("Lendo a planilha de empresas...")

# Abre a planilha
caminho = "empresas.xlsx"
wb = load_workbook(caminho, read_only=True)
ws = wb.active

# Lê linha por linha
for row in ws.iter_rows(min_row=1, values_only=True):
    nome_empresa = row[0]
    url_portal = row[1]
    if nome_empresa and url_portal:
        print(f"Empresa encontrada: {nome_empresa} -> Link: {url_portal}")
print("Leitura concluída!")

# Importando a biblioteca necessária para automatizar a navegação web
import time
from datetime import datetime
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By

print("=== INICIANDO O ROBÔ COM PAGINAÇÃO DA GUPY CORRIGIDA ===")

# 1. Pega a data de hoje para o filtro
data_hoje = datetime.now().strftime("%Y-%m-%d")
print(f"Data de referência (hoje): {data_hoje}\n")

# 2. Lê a planilha de empresas
caminho_planilha = "empresas.xlsx"
wb = load_workbook(caminho_planilha, read_only=True)
ws = wb.active

# 3. Inicializa o navegador
navegador = webdriver.Chrome()

try:
    for row in ws.iter_rows(min_row=1, values_only=True):
        nome_empresa = row[0]
        url_portal = row[1]
        
        if not nome_empresa or not url_portal:
            continue
            
        print(f"\n[Empresa] A aceder ao portal de: {nome_empresa}")
        navegador.get(url_portal)
        time.sleep(4)
        
        links_vagas = []
        
        # --- LOOP DE PAGINAÇÃO REAL DA GUPY ---
        pagina_atual = 1
        while True:
            print(f" -> A varrer a página {pagina_atual}...")
            
            # Rola até o fim para garantir o carregamento completo dos elementos da tabela e rodapé
            navegador.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            
            # Coleta os links de vagas visíveis na página atual
            elementos_links = navegador.find_elements(By.TAG_NAME, "a")
            novos_links = 0
            
            for el in elementos_links:
                href = el.get_attribute("href")
                if href and "/job/" in href and href not in links_vagas:
                    links_vagas.append(href)
                    novos_links += 1
                    
            print(f"    Encontrados {novos_links} novos links nesta página.")
            
            # Tenta encontrar e avançar para a próxima página
            try:
                # Na Gupy, o botão de avançar ('>') costuma vir logo após os números ou com aria-label específico
                # Vamos procurar pelo botão que avança a paginação
                botao_proximo = navegador.find_element(By.XPATH, "//button[@aria-label='Próxima página' or @aria-label='Próxima'] | //li[not(@class*='disabled')]//button[contains(., '>')]")
                
                # Se encontrar e estiver habilitado, clica
                if botao_proximo.is_displayed() and botao_proximo.isEnabled():
                    navegador.execute_script("arguments[0].click();", botao_proximo)
                    time.sleep(3) # Aguarda a nova página renderizar as vagas
                    pagina_atual += 1
                else:
                    print(" -> Chegamos à última página (botão desativado).")
                    break
            except Exception:
                # Tentativa alternativa: procurar pelo número da página seguinte (ex: se estamos na 1, clica na 2)
                try:
                    proxima_pagina_num = str(pagina_atual + 1)
                    botao_numero = navegador.find_element(By.XPATH, f"//button[text()='{proxima_pagina_num}']")
                    navegador.execute_script("arguments[0].click();", botao_numero)
                    time.sleep(3)
                    pagina_atual += 1
                    continue
                except:
                    print(" -> Fim das páginas para esta empresa.")
                    break
                
        print(f" Total acumulado de links recolhidos para {nome_empresa}: {len(links_vagas)}")
        
        # --- VALIDAR E COLETAR DADOS DE CADA VAGA ---
        vagas_validas_hoje = 0
        
        for link in links_vagas:
            try:
                navegador.get(link)
                time.sleep(2)
                
                # Pega o texto completo da página da vaga
                texto_pagina = navegador.find_element(By.TAG_NAME, "body").text
                
                # FILTRO DE DATA: Verifica se a vaga é de hoje
                if data_hoje in texto_pagina or "Hoje" in texto_pagina:
                    print(f"    [VAGA DE HOJE ENCONTRADA!] {link}")
                    vagas_validas_hoje += 1
                    
                # Retorna para a listagem para pegar o próximo link
                navegador.back()
                time.sleep(1.5)
                
            except Exception as e:
                print(f"    [Erro ao processar vaga] {e}")
                navegador.back()
                continue
                
        print(f"--- Concluído {nome_empresa}: {vagas_validas_hoje} vagas novas identificadas. ---")

finally:
    navegador.quit()
    print("\nProcesso finalizado com sucesso!")
    