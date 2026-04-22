import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class SeleniumHandler:
    def __init__(self):
        self.driver = None
        self.wait = None

    def _inicializar_driver(self):
        print("🔗 Conectando ao Chrome (Porta 9222)...")
        opts = ChromeOptions()
        opts.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        try:
            self.driver = webdriver.Chrome(options=opts)
            self.wait = WebDriverWait(self.driver, 15)
            print("✅ Conectado com sucesso!")
        except Exception as e:
            print(f"❌ Erro ao conectar: {e}")
            exit()

    def __enter__(self):
        self._inicializar_driver()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.driver = None

    def lidar_com_alertas(self):
        try:
            alert = self.driver.switch_to.alert
            alert.accept()
        except: pass

    def buscar_processo(self, processo: str) -> bool:
        try:
            self.driver.switch_to.default_content()
            self.lidar_com_alertas()
            campo = self.wait.until(EC.presence_of_element_located((By.ID, "txtPesquisaRapida")))
            self.driver.execute_script("arguments[0].focus(); arguments[0].click();", campo)
            campo.send_keys(Keys.CONTROL + "a")
            campo.send_keys(Keys.BACKSPACE)
            campo.send_keys(processo)
            time.sleep(0.5)
            campo.send_keys(Keys.ENTER)
            time.sleep(5) 
            return True
        except: return False

    def clicar_e_extrair(self) -> str:
        try:
            self.driver.switch_to.default_content()
            self.lidar_com_alertas()
            
            # Tenta clicar no botão Alterar ou Consultar
            try:
                self.driver.switch_to.frame("ifrVisualizacao")
                btn = self.driver.find_element(By.XPATH, "//a[contains(@href, 'procedimento_alterar') or contains(@href, 'procedimento_consultar') or contains(@title, 'Alterar') or contains(@title, 'Consultar')]")
                self.driver.execute_script("arguments[0].click();", btn)
            except:
                # Fallback: Varre frames
                self.driver.switch_to.default_content()
                iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
                for i in range(len(iframes)):
                    try:
                        self.driver.switch_to.default_content()
                        self.driver.switch_to.frame(i)
                        btn = self.driver.find_element(By.XPATH, "//a[contains(@href, 'procedimento_alterar') or contains(@href, 'procedimento_consultar')]")
                        self.driver.execute_script("arguments[0].click();", btn)
                        break
                    except: continue

            time.sleep(4)
            self.driver.switch_to.default_content()
            
            js_script = r"""
            function fetchRealValue(win) {
                try {
                    let targets = ['txtDescricao', 'txtEspecificacao', 'lblEspecificacao', 'divEspecificacao'];
                    for (let id of targets) {
                        let el = win.document.getElementById(id);
                        if (el) {
                            let val = el.value || el.innerText;
                            if (val && val.trim() !== "") return val.trim();
                        }
                    }
                    let spans = win.document.getElementsByTagName('span');
                    for (let s of spans) {
                        if (s.innerText.includes('Especificação:')) {
                            let next = s.nextElementSibling || s.parentElement.nextElementSibling;
                            if (next && next.innerText.trim() !== "") return next.innerText.trim();
                        }
                    }
                    for (let j = 0; j < win.frames.length; j++) {
                        let r = fetchRealValue(win.frames[j]);
                        if (r) return r;
                    }
                } catch(e) {}
                return null;
            }
            return fetchRealValue(window);
            """
            for _ in range(5):
                res = self.driver.execute_script(js_script)
                if res: return res
                time.sleep(1)
            return "⚠️ Dado não encontrado"
        except Exception as e:
            return f"❌ Erro: {str(e)[:15]}"