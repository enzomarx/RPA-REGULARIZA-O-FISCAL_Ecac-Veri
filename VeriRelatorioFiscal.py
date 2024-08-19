import os
import time
import yagmail
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui

# Configurações do selenium
download_dir = os.path.expanduser("~/Downloads") # diretorio de downloads do pc
chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": r"C:\Users\PC\Downloads",
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": False,
    "profile.default_content_settings.popups": 0
})

# Inicializando o WebDriver
driver = webdriver.Chrome(service=Service(), options=chrome_options)

# Função para encontrar os arquivos PDF na pasta
def find_all_pdfs():
    files = [f for f in os.listdir(download_dir) if f.endswith(".pdf")]
    return files

# Função para verificar se o arquivo foi completamente baixado
def wait_for_download(timeout=60):
    start_time = time.time()
    while time.time() - start_time < timeout:
        pdf_files = find_all_pdfs()
        if pdf_files:
            # Verifica se os arquivos estão completos (não são .crdownload)
            if all(not f.endswith('.crdownload') for f in pdf_files):
                return [os.path.join(download_dir, f) for f in pdf_files]
        time.sleep(1)
    return []

# Função para enviar e-mails com todos os PDFs
def send_email():
    recipients = ["ti2controllersbr@gmail.com", "carlos.junior@controllersbr.com", "ingrid@controllersbr.com", "joas@controllersbr.com", "juliocesar@controllersbr.com", "lucas@controllersbr.com"] # Ajuste conforme necessário
    sender_email = "redstarenzo@gmail.com" # Ajuste conforme necessário
    yag = yagmail.SMTP(sender_email, "xxxx xxxx xxxx") # Ajuste conforme necessário
    
    # Encontrar todos os arquivos PDF
    pdf_files = wait_for_download()
    if pdf_files:
        yag.send(
            to=recipients,
            subject="Relatório Automático (Regularização Fiscal)", # Ajuste conforme necessário
            contents="Por favor, veja os relatórios em anexo.",
            attachments=pdf_files,
        )
        print(f"E-mail enviado com sucesso com os anexos {pdf_files}")
        
        # Deletar todos os arquivos da pasta após o envio
        for file_path in pdf_files:
            os.remove(file_path)
        print("Todos os arquivos foram deletados.")
    else:
        print("Nenhum arquivo PDF encontrado para enviar.")

def run_automation():
    # Setup
    driver.get("https://08978175000180.portal-veri.com.br")
    driver.set_window_size(1366, 768)

    # Actions in Login
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="kt_sign_in_form"]/div[2]/input'))).send_keys("br@gmail.com") # Ajuste conforme necessário
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="kt_sign_in_form"]/div[3]/input'))).send_keys("senha") # Ajuste conforme necessário
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="kt_sign_in_form"]/div[5]/button'))).click()

    # Actions in Dashboard
    print("Tentando clicar no elemento X")
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".menu-icon > .ki-sms"))).click()
    print("Clique no elemento X realizado")         

    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".col-md-3:nth-child(3) .text-gray-700"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".buttons-pdf"))).click()
    pyautogui.press('enter')
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".menu-icon > .ki-chart-pie-simple"))).click()
    
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".row:nth-child(4) > .col-md-3:nth-child(1) .text-gray-700"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".buttons-pdf > span"))).click()
    pyautogui.press('enter')
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".menu-icon > .ki-chart-pie-simple"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".row:nth-child(4) > .col-md-3:nth-child(2) .text-gray-700"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".buttons-pdf > span"))).click()
    pyautogui.press('enter')
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".menu-icon > .ki-chart-pie-simple"))).click()  
    
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".row:nth-child(4) > .col-md-3:nth-child(3) .text-gray-700"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".buttons-pdf > span"))).click()
    pyautogui.press('enter')
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".menu-icon > .ki-chart-pie-simple"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".row:nth-child(4) > .col-md-3:nth-child(4) .text-gray-700"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".buttons-pdf > span"))).click()
    pyautogui.press('enter')
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".menu-icon > .ki-chart-pie-simple"))).click()
    
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".row:nth-child(6) > .col-md-3:nth-child(1) .text-gray-700"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".row:nth-child(6) > .col-md-3:nth-child(1) a:nth-child(1)"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".buttons-pdf"))).click()
    pyautogui.press('enter')
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".menu-icon > .ki-chart-pie-simple"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".row:nth-child(6) > .col-md-3:nth-child(1) .card-body a:nth-child(2)"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".buttons-pdf > span"))).click()
    pyautogui.press('enter')
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".menu-icon > .ki-chart-pie-simple"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".col-md-3:nth-child(1) a:nth-child(3)"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".buttons-pdf > span"))).click()
    pyautogui.press('enter')
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#\\#kt_app_sidebar_menu > .menu-item:nth-child(4) > .menu-link"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".row:nth-child(6) a:nth-child(4)"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".buttons-pdf > span"))).click()
    pyautogui.press('enter')
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#\\#kt_app_sidebar_menu > .menu-item:nth-child(4) > .menu-link"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a:nth-child(5)"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".buttons-pdf > span"))).click()
    pyautogui.press('enter')
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".menu-icon > .ki-chart-pie-simple"))).click()
    
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".row:nth-child(6) a:nth-child(6)"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".buttons-pdf > span"))).click()
    pyautogui.press('enter')
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#\\#kt_app_sidebar_menu > .menu-item:nth-child(4) > .menu-link"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a:nth-child(7)"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".buttons-pdf > span"))).click()
    pyautogui.press('enter')
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".menu-icon > .ki-chart-pie-simple"))).click()
    
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".row:nth-child(6) a:nth-child(8)"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".buttons-pdf > span"))).click()
    pyautogui.press('enter')
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".menu-icon > .ki-chart-pie-simple"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".row:nth-child(6) > .col-md-3:nth-child(3) .text-gray-700"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".buttons-pdf > span"))).click()
    pyautogui.press('enter')
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".menu-icon > .ki-chart-pie-simple"))).click()
    
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".row:nth-child(8) > .col-md-3:nth-child(1) .text-gray-700"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".buttons-pdf"))).click()
    pyautogui.press('enter')
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#\\#kt_app_sidebar_menu > .menu-item:nth-child(4) > .menu-link"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".row:nth-child(8) > .col-md-3:nth-child(2) .text-gray-700"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".buttons-pdf > span"))).click()
    pyautogui.press('enter')
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".menu-icon > .ki-chart-pie-simple"))).click()
    
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".row:nth-child(8) > .col-md-3:nth-child(3) .text-gray-700"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".buttons-pdf > span"))).click()
    pyautogui.press('enter')
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".menu-icon > .ki-chart-pie-simple"))).click()
    

    # Aguardar download ser concluído
    time.sleep(10)
    
    # Envia o e-mail já com documentos baixados
    send_email()
    
    # Fechar
    driver.quit()

# Agendamento
# def job():
#     run_automation()
# schedule.every().monday.at("12:00").do(job)

# Mantém o script rodando para cumprir agendamento
# while True:
#     schedule.run_pending()
#     time.sleep(1)

run_automation()
