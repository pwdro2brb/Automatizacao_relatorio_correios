# precisa instalar o pip install pywin32 e pip install selenium no terminal de comando
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select # Para caixas de <select>
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
import win32com.client
import os

''''''
# --- Configuração ---
# (Você precisará baixar o "chromedriver" e colocar o caminho aqui,
# ou deixar o Selenium 4+ baixá-lo automaticamente)f
driver = webdriver.Chrome()
wait = WebDriverWait(driver, 70) # Define um tempo de espera máximo (70 seg)


# ----------- Parte 1 Agilis ----------- # 
# DEFINA SEUS DADOS DE LOGIN E A URL INICIAL
URL_INICIAL = "https://agilis.mrv.com.br/HomePage.do?view_type=my_view"
try:
  # 0. ABRIR A PÁGINA E FAZER LOGIN
  driver.get(URL_INICIAL)
  print(f"Página aberta: {URL_INICIAL}")
  print("Aguardando tela de login...")

  # O seletor mais provável para esse botão é pelo texto.
  # Tentativa 1: Usando By.LINK_TEXT (se for uma tag <a>)
  try:
    selector_login_integrado = (By.LINK_TEXT, "Login Integrado Microsoft")
    wait.until(EC.element_to_be_clickable(selector_login_integrado)).click()
    
  # Tentativa 2: Usando By.XPATH (funciona para <button>, <div>, <span>, etc.)
  except:
    print("Não encontrou por LINK_TEXT. Tentando por XPATH...")
    # Este XPATH procura QUALQUER elemento que tenha o texto exato.
    selector_login_integrado = (By.XPATH, "//*[text()='Login Integrado Microsoft']")
    wait.until(EC.element_to_be_clickable(selector_login_integrado)).click()

  print("0. Cliquei em 'Login Integrado Microsoft'.")
  print("Aguardando autenticação SSO e carregamento da página principal...")

  # Preenche o e-mail (usando a variável segura)
  email_field_microsoft = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.ID, "i0116")) 
  )
  print("Preenchendo e-mail da Microsoft...")
  email_field_microsoft.send_keys("pedro.henrsilva@mrv.com.br")#Coloque o seu email aqui
  WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
  
  # Preenche a senha (usando a variável segura)
  password_field_microsoft = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.ID, "i0118")) 
  )
  print("Preenchendo senha da Microsoft...")
  password_field_microsoft.send_keys("Sua senha aqui")#Coloque a sua senha aqui
  
  # Tenta clicar no botão "Entrar" (com loop anti-stale)
  print("Procurando o botão 'Entrar'...")
  tentativas = 0
  clicado_entrar = False
  while not clicado_entrar and tentativas < 5:
    try:
      entrar_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
      entrar_button.click()
      clicado_entrar = True
      print("Botão 'Entrar' clicado.")
    except StaleElementReferenceException:
      tentativas += 1; time.sleep(0.5)
  if not clicado_entrar: raise Exception("Falha ao clicar em Entrar")

  # --- ESPERA PELO MFA MANUAL ---
  print("!!! AÇÃO MANUAL NECESSÁRIA !!!")
  print("Aguardando aprovação do MFA no seu celular (até 180s)...")
  
  tentativas = 0
  clicado_manter = False
  while not clicado_manter and tentativas < 5:
    try:
      keep_logged_in_button = WebDriverWait(driver, 180).until( 
        EC.element_to_be_clickable((By.ID, "idSIButton9"))
      )
      keep_logged_in_button.click() # Clica "Sim"
      clicado_manter = True
      print("MFA Aprovado! Botão 'Manter conectado' clicado.")
    except StaleElementReferenceException:
      tentativas += 1; time.sleep(0.5)
    except TimeoutException:
      print("Erro: Timeout após 180s. Você não aprovou o MFA a tempo?")
      clicado_manter = False; break
  if not clicado_manter: raise Exception("Falha ao clicar em Manter Conectado")

  print("Login da Microsoft concluído na janela pop-up.")

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
  print("  - Cliquei em 'Contratos - ADM'.")

  selector_produtividade = (By.LINK_TEXT, "Produtividade Contratos - ADM")
  wait.until(EC.element_to_be_clickable(selector_produtividade)).click()
  print("  - Cliquei em 'Produtividade Contratos - ADM'.")

  # 3. APERTAR "EDITAR"
  # (Pode ser por ID, NOME, ou texto. 'name' é um bom chute em apps Zoho)
  selector_editar = (By.CLASS_NAME, "linkborder") # Chute, pode ser "Editar"
  wait.until(EC.element_to_be_clickable(selector_editar)).click()
  print("3. Cliquei em 'Editar'.")

  # 4. PASSO 1 - SELECIONAR "COLETOR DE CUSTO ADM" E MOVER
  # Selecionar o item na lista da esquerda (pelo texto)
  selector_coletor = (By.XPATH, "//option[text()='Coletor de custo ADM']")
  wait.until(EC.element_to_be_clickable(selector_coletor)).click()
  print("4. Selecionei 'Coletor de custo ADM'.")
  

  # O seletor dele provavelmente é um 'class' ou 'onclick'
  selector_seta_direita = (By.CLASS_NAME, "moverightButton") # CHUTE!
  driver.find_element(*selector_seta_direita).click()
  print("  - Cliquei na seta para mover.")

  print("4.5. Expandindo 'Passo 2: Opções de filtragem'...")
  try:
    # O ID 'reportstep2' foi confirmado pela sua imagem do Inspecionar
    selector_opcoes_filtragem = (By.ID, "rcstep2src")
    # Encontra o elemento que você quer clicar (use "reportstep2" que é melhor)
    elemento_clique = wait.until(EC.presence_of_element_located((By.ID, "rcstep2src")))

    # Usa JavaScript para forçar o clique
    driver.execute_script("arguments[0].click();", elemento_clique)

    print("  - SUCESSO: Cliquei em 'Opções de filtragem' (via JavaScript).")
    
    # Pequena pausa para a animação de expandir terminar
    time.sleep(1) 

  except TimeoutException:
    print("  - FALHA: Não foi possível encontrar 'Passo 2: Opções de filtragem' (ID: reportstep2).")
    # Se este passo falhar, o próximo (clicar no rádio) também vai falhar.
    raise # 'raise' vai parar o script e pular para o bloco 'except Exception'

  # --- PASSO 5: Selecionar o rádio 'Durante' ---
  print("5. Selecionando o filtro 'Durante'...")
  try: 
    # Usando CSS_SELECTOR para encontrar pelo atributo [value='predefined']
    selector_radio_durante = (By.CSS_SELECTOR, "input[value='predefined']")
    
    # Espera o rádio ficar clicável
    wait.until(EC.element_to_be_clickable(selector_radio_durante)).click()
    print("  - SUCESSO: Filtro 'Durante' selecionado.")

  except TimeoutException:
    print("  - FALHA: Não foi possível encontrar o rádio 'Durante' (CSS_SELECTOR: input[value='predefined']).")
    raise # Para o script se não encontrar

  # 6. APERTAR "EXECUTAR RELATÓRIO"
  selector_executar = (By.ID, "addnew223222")
  wait.until(EC.element_to_be_clickable(selector_executar)).click()
  print("6. Cliquei em 'Executar relatório'.")
  print("--- Relatório executado, aguardando 10s para carregar...")
  time.sleep(3) # Pausa importante para o relatório carregar

  # --- 7. Enviar Relatório por E-mail ---
  print("7. Iniciando envio de e-mail...")
  try:  
    # 7. APERTAR "ENVIAR POR EMAIL ESTE ARQUIVO"
    SELECIONAR_ARQ_EMAIL = (By.ID, "sendmaillink") 
    wait.until(EC.element_to_be_clickable(SELECIONAR_ARQ_EMAIL)).click()
    print("7. Cliquei em 'Enviar este arquivo por email'.")

    # 7.1 - Aguardar o pop-up (modal) de e-mail aparecer
    # (Estamos esperando o botão "Enviar" de DENTRO do pop-up ficar visível)
    selector_enviar_modal = (By.CSS_SELECTOR, "input[value='Enviar']")
    wait.until(EC.element_to_be_clickable(selector_enviar_modal))
    print("  - Modal de e-mail aberto.")
    
    # 7.2 - Selecionar "XLS" no dropdown "Formato"
    select_element = wait.until(EC.element_to_be_clickable((By.ID, "file_type")))
    select_obj = Select(select_element)
    
    select_obj.select_by_visible_text("XLS")
    print("  - Formato 'XLS' selecionado.")

    # 7.3 - Preencher o campo "Para"
    email_para = "pedro.henrsilva@mrv.com.br"
    wait.until(EC.element_to_be_clickable((By.ID, "toEmailSearch"))).send_keys(email_para)
    print(f"  - E-mail preenchido: {email_para}")

    # 7.4 - Clicar no botão "Enviar" final
    driver.find_element(*selector_enviar_modal).click()
    print("  - E-mail enviado com sucesso!")

    # Pausa para ver o resultado antes de fechar
    time.sleep(20) 

  except TimeoutException:
    print("  - FALHA: Não foi possível encontrar um dos elementos do modal de e-mail.")
    print("  - Verifique os seletores ")
    pass # Continua mesmo se falhar
  
  print("--- Automação concluída com sucesso! ---")
  time.sleep(20) # Pausa para você ver o resultado

except Exception as e:
  print(f"ERRO: A automação falhou.")
  print(e)

finally:
  driver.quit() # Comente "driver.quit()" para o navegador não fechar no final 
  print("Script finalizado.")

import win32com.client
import os
import time

time.sleep(10)

# --- CONFIGURAÇÃO ---
PASTA_DOWNLOAD = os.getcwd() 
NOME_ASSUNTO = "Produtividade Contratos - ADM"
NOME_REMETENTE = "Agilis"
NOME_ANEXO = "Produtividade Contratos - ADM.xls"
NOME_ARQUIVO_FINAL = "Produtividade_EDITADO.xlsx"
# --------------------

def processar_relatorio_email():
  """
  Conecta ao Outlook, baixa o anexo, edita no Excel (deletando linhas)
  e cria uma planilha de resumo formatada.
  """
  
  arquivo_salvo_path = None
  excel = None
  wb = None
  
  # Constantes de formatação do Excel
  xlUp = -4162
  xlCellTypeVisible = 12 # Embora não usemos mais o filtro, manter pode ser útil
  xlPasteValues = -4163 
  xlOpenXMLWorkbook = 51
  xlCenter = -4108      
  xlTop = -4160       
  xlContinuous = 1      
  # -----------------------------------------------

  # --- 1. CONECTAR AO OUTLOOK E BAIXAR O ANEXO ---
  try:
    print("Conectando ao Outlook...")
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6) 
    print(f"Procurando na Caixa de Entrada por e-mails de '{NOME_REMETENTE}'...")

    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    email_encontrado = None
    for i in range(10): 
      for message in messages:
        if message.Subject == NOME_ASSUNTO and message.SenderName == NOME_REMETENTE:
          print(f"\n--- E-mail encontrado! (Assunto: {message.Subject}) ---")
          email_encontrado = message
          break
      if email_encontrado:
        break 
      print(f"E-mail ainda não encontrado. Tentando novamente em 60 segundos... ({i+1}/10)")
      time.sleep(60)

    if not email_encontrado:
      print(f"ERRO: E-mail não encontrado após 5 minutos.")
      return

    attachments = email_encontrado.Attachments
    if attachments.Count > 0:
      for attachment in attachments:
        if attachment.FileName == NOME_ANEXO:
          arquivo_salvo_path = os.path.join(PASTA_DOWNLOAD, NOME_ANEXO)
          attachment.SaveAsFile(arquivo_salvo_path)
          print(f"Anexo '{NOME_ANEXO}' salvo em: {arquivo_salvo_path}")
          break
    if not arquivo_salvo_path:
      print(f"ERRO: E-mail encontrado, mas o anexo '{NOME_ANEXO}' não estava nele.")
      return
  except Exception as e:
    print(f"ERRO ao conectar ao Outlook ou baixar o anexo: {e}")
    return

  # --- 2. ABRIR O EXCEL E EDITAR O ARQUIVO ---
  try:
    print("\n--- Iniciando edição no Excel ---")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True 
    excel.DisplayAlerts = False 

    wb = excel.Workbooks.Open(arquivo_salvo_path)
    ws = wb.Worksheets(1) # ws = Planilha de dados processada

    # --- 2.1 a 2.6. Limpeza (SEM AutoFilter) ---
    print("Editando: Excluindo linhas 1-8...")
    ws.Rows("1:8").Delete()
    
    print("Editando: Removendo logos/imagens...")
    for shape in ws.Shapes:
      shape.Delete()
    
    print("Editando: Ajustando altura e largura...")
    ws.Rows.RowHeight = 12
    ws.Columns.ColumnWidth = 12
    
    # --- (NOVO) 2.7. Deletar linhas que NÃO SÃO de "Correspondência" E/OU têm hora < 12:00 ------
    print("Editando: Removendo linhas que não são de 'Solicitação de Envio de Correspondência' OU têm hora de conclusão antes do meio-dia...")
    
    # ATENÇÃO AQUI: Se "Correspondência" for muito genérico, 
    # use "Solicitação de Envio de Correspondência"
    # --- (NOVO) 2.7. Deletar linhas que NÃO SÃO de "Correspondência" ---

    print("Editando: Removendo linhas que não são de 'Solicitação de Envio de Correspondência'...")

   

    # ATENÇÃO AQUI: Se "Correspondência" for muito genérico,

    # use "Solicitação de Envio de Correspondência"

    filtro_texto = "Solicitação de Envio de Correspondência"

   

    # Achar a última linha (baseado na Coluna H, campo 8)

    last_row = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row

    print(f"Verificando {last_row} linhas (de baixo para cima)...")



    # Loop de baixo para cima para deletar com segurança

    for i in range(last_row, 1, -1): # (de last_row até a linha 2)

      cell_value = str(ws.Cells(i, 9).Value) # Coluna H

     

      # Se o texto NÃO for encontrado, deleta a linha

      if filtro_texto not in cell_value:

        ws.Rows(i).Delete()

    print("Linhas indesejadas removidas.")

    # 2.9. Adicionar fórmulas (Agora sem SpecialCells)
    print("Editando: Adicionando fórmulas nas colunas M-R...")
    # Achar a *nova* última linha, após as deleções
    last_row = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    print(f"Última linha de dados (pós-deleção): {last_row}")

    if last_row > 1: 
      # Não precisamos mais de 'SpecialCells' ou do 'try/except' para filtro vazio
      ws.Range(f"N2:N{last_row}").FormulaLocal = "=TEXTODEPOIS(F2;\"ADM:\")"
      ws.Range(f"O2:O{last_row}").FormulaLocal = "=ESQUERDA(N2;11)"
      ws.Range(f"P2:P{last_row}").FormulaLocal = "=TEXTODEPOIS(F2;\"Correspondência:\")"
      ws.Range(f"Q2:Q{last_row}").FormulaLocal = "=TEXTOANTES(P2;\"Cidade\")"
      ws.Range(f"R2:R{last_row}").FormulaLocal = "=TEXTODEPOIS(F2;\"Documentos:\")"
      ws.Range(f"S2:S{last_row}").FormulaLocal = "=TEXTOANTES(R2;\"*\")"
      print("Fórmulas aplicadas com sucesso.")
    else:
      print("AVISO: Nenhuma linha de 'Correspondência' foi encontrada.")

    # 2.11 Criar Planilha de Resumo
    print("Criando: Planilha de Resumo...")
    ws_summary = wb.Worksheets.Add(After=ws)
    ws_summary.Name = "Resumo"
    ws_summary.Activate()

    # 2.12 Montar o Layout do Resumo
    print("Criando: Layout do Resumo (Título e Cabeçalho)...")
    today_date = time.strftime("%d/%m/%Y")
    
    title_range = ws_summary.Range("A1:D1")
    title_range.Merge()
    title_range.Value = f"MRV - DATA - {today_date}"
    title_range.Font.Bold = True
    title_range.HorizontalAlignment = xlCenter 
    
    ws_summary.Range("A2").Value = "Centro de Custo"
    ws_summary.Range("B2").Value = "Chamado"
    ws_summary.Range("C2").Value = "Serviço"
    ws_summary.Range("D2").Value = "Quantidade"
    summary_header = ws_summary.Range("A2:D2")
    summary_header.Font.Bold = True
    summary_header.Font.Color = 16777215 
    summary_header.Interior.Color = 12611584 
    summary_header.AutoFilter()
    ws_summary.Columns("A:D").ColumnWidth = 22

    # 2.13 Copiar Dados Visíveis para o Resumo (Agora sem SpecialCells)
    print("Copiando dados filtrados para o Resumo (somente valores)...")
    if last_row > 1:
      ws.Range(f"O2:O{last_row}").Copy()
      ws_summary.Range("A3").PasteSpecial(Paste=xlPasteValues)

      ws.Range(f"B2:B{last_row}").Copy()
      ws_summary.Range("B3").PasteSpecial(Paste=xlPasteValues)

      ws.Range(f"Q2:Q{last_row}").Copy()
      ws_summary.Range("C3").PasteSpecial(Paste=xlPasteValues)

      ws.Range(f"S2:S{last_row}").Copy()
      ws_summary.Range("D3").PasteSpecial(Paste=xlPasteValues)
      print("Dados copiados com sucesso.")
    else:
      print("AVISO: Nenhum dado filtrado para copiar.")
    
    excel.Application.CutCopyMode = False
    
    # 2.13.1 Adicionar Rodapé Dinâmico
    print("Criando: Rodapé do Resumo...")
    last_summary_row = ws_summary.Cells(ws_summary.Rows.Count, "A").End(xlUp).Row
    footer_row = max(last_summary_row + 2, 57)
    
    footer_cliente = ws_summary.Range(f"A{footer_row}:B{footer_row + 1}")
    footer_cliente.Merge()
    footer_cliente.Value = "CLIENTE:"
    footer_cliente.Font.Bold = True
    footer_cliente.VerticalAlignment = xlTop 

    footer_agf = ws_summary.Range(f"C{footer_row}:D{footer_row + 1}")
    footer_agf.Merge()
    footer_agf.Value = "AGF:"
    footer_agf.Font.Bold = True
    footer_agf.VerticalAlignment = xlTop 
    
    # 2.13.2 Adicionar Bordas
    print("Criando: Bordas da planilha Resumo...")
    final_used_row = footer_row + 1
    full_range = ws_summary.Range(f"A1:D{final_used_row}")
    full_range.Borders.LineStyle = xlContinuous
    
    # 2.14 Ativar resumo (NÃO precisamos mais limpar o filtro)
    ws_summary.Activate() 


    # --- 3. SALVAR O ARQUIVO FINAL ---
    caminho_final = os.path.join(PASTA_DOWNLOAD, NOME_ARQUIVO_FINAL)
    
    wb.SaveAs(caminho_final, FileFormat=xlOpenXMLWorkbook) 
    print(f"\n--- SUCESSO! ---")
    print(f"Arquivo editado e com resumo salvo como: {caminho_final}")

  except Exception as e:
    print(f"ERRO durante a edição no Excel: {e}")

'''
  finally:
    # --- 4. FECHAR O EXCEL ---
    if wb:
      wb.Close(SaveChanges=False) 
    if excel:
      excel.Quit() 
    wb = None
    excel = None
'''
if __name__ == "__main__":
  processar_relatorio_email()
