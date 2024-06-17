from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# Configurar o serviço do ChromeDriver usando webdriver_manager
service = Service(ChromeDriverManager().install())

# Inicializar o driver do Chrome com o serviço configurado
driver = webdriver.Chrome(service=service)

# Abre uma página
driver.get("https://fornecedor2.procon.sp.gov.br/u/m/atendimentos/f/70a58963-4da2-eb11-b1ac-0022483741ab/cips/eyJhYmVydG9zIjp0cnVlfQ--")


