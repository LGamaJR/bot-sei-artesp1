import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

class SeleniumHandler:
    def __init__(self):
        self.driver = None
        self.wait = None

    def _inicializar_driver(self):
        print("🔗 Conectando ao Chrome (Porta 9222)...")
        opts = ChromeOptions()
        opts.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        try:
            service = ChromeService(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=opts)
            self.wait = WebDriverWait(self.driver, 10)
            print("✅ Conectado!")
        except Exception as e:
            print(f"❌ Erro ao conectar: {e}")
            exit()

    def __enter__(self):
        self._inicializar_driver()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.driver = None

    def lidar_com_alertas(self):
        """Fecha pop-ups (como 'Link sem assinatura') que travam o navegador."""
        try:
            alert = self.driver.switch_to.alert
            print(f"   ⚠️ Fechando alerta: {alert.text}")
            alert.accept()
        except:
            pass

    def buscar_processo(self, processo: str) -> bool:
        try:
            self.driver.switch_to.default_content()
            self.lidar_com_alertas()
            
            # Força o foco no campo de busca via JavaScript
            js_foco = "let c = document.getElementById('txtPesquisaRapida'); if(c){c.focus(); c.click();}"
            self.driver.execute_script(js_foco)
            time.sleep(0.5)

            campo = self.wait.until(EC.element_to_be_clickable((By.ID, "txtPesquisaRapida")))
            campo.send_keys(Keys.CONTROL + "a")
            campo.send_keys(Keys.BACKSPACE)
            
            print(f"   ⌨️ Digitando: {processo}")
            campo.send_keys(processo)
            campo.send_keys(Keys.ENTER)
            
            time.sleep(4)
            return True
        except Exception as e:
            print(f"   ⚠️ Falha na busca: {str(e)[:50]}")
            self.driver.get("https://sei.sp.gov.br/sei/controlador.php?acao=procedimento_controlar")
            return False

    def clicar_e_extrair(self) -> str:
        try:
            self.driver.switch_to.default_content()
            self.lidar_com_alertas()
            
            # Varredura de frames para achar o botão de alteração (engrenagem)
            iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
            clicou = False
            for i in range(len(iframes)):
                try:
                    self.driver.switch_to.default_content()
                    frames = self.driver.find_elements(By.TAG_NAME, "iframe")
                    self.driver.switch_to.frame(frames[i])
                    btn = self.driver.find_element(By.XPATH, "//a[contains(@href, 'procedimento_alterar')] | //a[contains(@title, 'Alterar Processo')]")
                    self.driver.execute_script("arguments[0].click();", btn)
                    clicou = True
                    break
                except: continue

            if not clicou: return "⚠️ Botão não localizado"

            time.sleep(4)
            
            # Extração profunda (JavaScript recursivo)
            self.driver.switch_to.default_content()
            js_scan = r"""
            function scan(win) {
                const regex = /L\d{2}\s*-\s*[^|\n\r\t]+/i;
                try {
                    let txt = win.document.body.innerText;
                    let ins = win.document.querySelectorAll('input, textarea');
                    for (let i of ins) { txt += " " + i.value; }
                    let m = txt.match(regex);
                    if (m) return m[0].trim();
                    for (let j = 0; j < win.frames.length; j++) {
                        let r = scan(win.frames[j]);
                        if (r) return r;
                    }
                } catch(e) {}
                return null;
            }
            return scan(window);
            """
            resultado = self.driver.execute_script(js_scan)
            
            return resultado if resultado else "⚠️ Lote não preenchido"
        except:
            return "❌ Falha na extração"