# Para funcionar use os comandos pip install selenium  || pip install pywin32
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select  # Para caixas de <select>
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# --- Configuração ---
# (Você precisará baixar o "chromedriver" e colocar o caminho aqui,
# ou deixar o Selenium 4+ baixá-lo automaticamente)
driver = webdriver.Chrome()
wait = WebDriverWait(driver, 70)  # Define um tempo de espera máximo (70 seg)

# DEFINA SEUS DADOS DE LOGIN E A URL INICIAL
URL_INICIAL = "https://agilis.mrv.com.br/HomePage.do?view_type=my_view"
EMAIL_DESTINO = "pedro.henrsilva@mrv.com.br"

try:
    # 0. ABRIR A PÁGINA E FAZER LOGIN
    driver.get(URL_INICIAL)
    print(f"Página aberta: {URL_INICIAL}")
    print("Aguardando tela de login...")

    # Tenta encontrar o botão de Login Integrado Microsoft
    # Tentativa de busca por XPATH (mais robusto)
    selector_login_integrado = (By.XPATH, "//*[text()='Login Integrado Microsoft']")
    wait.until(EC.element_to_be_clickable(selector_login_integrado)).click()
    
    print("0. Cliquei em 'Login Integrado Microsoft'.")
    print("Aguardando autenticação SSO e carregamento da página principal...")
    
    # --- O RESTO DO SEU SCRIPT CONTINUA DAQUI ---

    # 1. CLICAR EM 'RELATÓRIOS'
    # Aguarda o menu 'Relatórios' ficar visível após o login
    selector_relatorios = (By.LINK_TEXT, "Relatórios")
    wait.until(EC.element_to_be_clickable(selector_relatorios)).click()
    print("1. Cliquei em 'Relatórios'.")
    print("2. Navegando no menu...")
    
    # 2. IR EM "CONTRATOS - ADM" -> "PRODUTIVIDADE CONTRATOS - ADM"
    # (Provavelmente precisa clicar no primeiro para o segundo aparecer)
    selector_contratos_adm = (By.LINK_TEXT, "Contratos - ADM")
    wait.until(EC.element_to_be_clickable(selector_contratos_adm)).click()
    print(" - Cliquei em 'Contratos - ADM'.")
    
    selector_produtividade = (By.LINK_TEXT, "Produtividade Contratos - ADM")
    wait.until(EC.element_to_be_clickable(selector_produtividade)).click()
    print(" - Cliquei em 'Produtividade Contratos - ADM'.")

    # 3. APERTAR "EDITAR"
    # (O seletor original "linkborder" é um chute, mantido aqui)
    selector_editar = (By.CLASS_NAME, "linkborder") 
    wait.until(EC.element_to_be_clickable(selector_editar)).click()
    print("3. Cliquei em 'Editar'.")

    # ATENÇÃO: Se as próximas etapas falharem, é provável que precise
    # MUDAR O FOCO PARA DENTRO DE UM IFRAME.
    
    # 4. PASSO 1 - SELECIONAR "COLETOR DE CUSTO ADM" E MOVER
    # Selecionar o item na lista da esquerda (pelo texto)
    selector_coletor = (By.XPATH, "//option[text()='Coletor de custo ADM']")
    wait.until(EC.element_to_be_clickable(selector_coletor)).click()
    print("4. Selecionei 'Coletor de custo ADM'.")

    # Mover o item selecionado
    selector_seta_direita = (By.CLASS_NAME, "moverightButton") 
    # Usando find_element com * para desempacotar a tupla (By.CLASS_NAME, "moverightButton")
    driver.find_element(*selector_seta_direita).click()
    print(" - Cliquei na seta para mover.")

    # 4.5. Expandindo 'Passo 2: Opções de filtragem'
    print("4.5. Expandindo 'Passo 2: Opções de filtragem'...")
    try:
        selector_opcoes_filtragem = (By.ID, "rcstep2src")
        
        # Encontra o elemento (use "rcstep2src" que é o ID para o clique)
        elemento_clique = wait.until(EC.presence_of_element_located(selector_opcoes_filtragem))
        
        # Usa JavaScript para forçar o clique, pois o clique normal pode falhar em elementos de menu
        driver.execute_script("arguments[0].click();", elemento_clique)
        print(" - SUCESSO: Cliquei em 'Opções de filtragem' (via JavaScript).")
        time.sleep(1) # Pausa para a animação
        
    except TimeoutException:
        print(" - FALHA: Não foi possível encontrar 'Passo 2: Opções de filtragem' (ID: rcstep2src).")
        raise # Interrompe a execução se o filtro não puder ser expandido

    # --- PASSO 5: Selecionar o rádio 'Durante' ---
    print("5. Selecionando o filtro 'Durante'...")
    try:
        # Usando CSS_SELECTOR para encontrar pelo atributo [value='predefined']
        selector_radio_durante = (By.CSS_SELECTOR, "input[value='predefined']")
        wait.until(EC.element_to_be_clickable(selector_radio_durante)).click()
        print(" - SUCESSO: Filtro 'Durante' selecionado.")
        
    except TimeoutException:
        print(" - FALHA: Não foi possível encontrar o rádio 'Durante' (CSS_SELECTOR: input[value='predefined']).")
        raise # Interrompe a execução se o rádio não for encontrado

    # 6. APERTAR "EXECUTAR RELATÓRIO"
    selector_executar = (By.ID, "addnew223222") # ID CHUTE
    wait.until(EC.element_to_be_clickable(selector_executar)).click()
    print("6. Cliquei em 'Executar relatório'.")
    print("--- Relatório executado, aguardando carregamento...")
    time.sleep(10) # Pausa maior para garantir que o relatório carregue

    # --- 7. Enviar Relatório por E-mail ---
    print("7. Iniciando envio de e-mail...")
    try:
        # 7. APERTAR "ENVIAR POR EMAIL ESTE ARQUIVO"
        SELECIONAR_ARQ_EMAIL = (By.ID, "sendmaillink") # ID CHUTE
        wait.until(EC.element_to_be_clickable(SELECIONAR_ARQ_EMAIL)).click()
        print("7. Cliquei em 'Enviar este arquivo por email'.")

        # 7.1 - Aguardar o pop-up (modal) de e-mail aparecer
        selector_enviar_modal = (By.CSS_SELECTOR, "input[value='Enviar']")
        wait.until(EC.element_to_be_clickable(selector_enviar_modal))
        print(" - Modal de e-mail aberto.")

        # 7.2 - Selecionar "XLS" no dropdown "Formato"
        select_element = wait.until(EC.element_to_be_clickable((By.ID, "file_type")))
        select_obj = Select(select_element)
        select_obj.select_by_visible_text("XLS")
        print(" - Formato 'XLS' selecionado.")

        # 7.3 - Preencher o campo "Para"
        wait.until(EC.element_to_be_clickable((By.ID, "toEmailSearch"))).send_keys(EMAIL_DESTINO)
        print(f" - E-mail preenchido: {EMAIL_DESTINO}")

        # 7.4 - Clicar no botão "Enviar" final
        driver.find_element(*selector_enviar_modal).click()
        print(" - E-mail enviado com sucesso!")

        # Pausa para ver o resultado antes de fechar
        time.sleep(5)
        
    except TimeoutException:
        print(" - FALHA: Não foi possível encontrar um dos elementos do modal de email.")
        print(" - Verifique os seletores.")
        
    print("--- Automação concluída com sucesso! ---")
    time.sleep(5) # Pausa final para você ver o resultado

except Exception as e:
    print(f"\nERRO: A automação falhou em uma etapa crítica.")
    print(f"Tipo de Erro: {type(e).__name__}")
    print(f"Detalhes do Erro: {e}")

finally:
    # driver.quit() # Descomente esta linha para fechar o navegador automaticamente
    print("\nScript finalizado.")