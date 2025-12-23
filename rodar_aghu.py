from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException,
)
import pandas as pd
import time
import os

import shutil
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service


def criar_driver():
    options = Options()

    # Modo headless é obrigatório na nuvem
    options.add_argument("--headless=new")  # se der erro, trocar para "--headless"
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")

    # Caminho do binário do Chromium instalado via packages.txt
    chrome_bin = (
        shutil.which("chromium")
        or shutil.which("chromium-browser")
        or shutil.which("google-chrome")
    )
    if chrome_bin:
        options.binary_location = chrome_bin

    # Caminho do chromedriver instalado via packages.txt
    driver_path = shutil.which("chromedriver")

    if driver_path:
        service = Service(driver_path)
        driver = webdriver.Chrome(service=service, options=options)
    else:
        # Fallback para quando você rodar LOCALMENTE com webdriver_manager
        from webdriver_manager.chrome import ChromeDriverManager

        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)

    return driver

# ======================================
# CONFIGURAÇÕES
# ======================================
# URL que você já usava – ela mesma leva para a tela de login, se ainda não estiver logada
URL = "https://aghu.hc-ufpe.ebserh/aghu/pages/casca/casca.xhtml"

EXCEL_ENTRADA = r"C:\Users\camila.melo.3\Desktop\aplicativo\Planilha Teste.xlsx"
EXCEL_SAIDA   = r"C:\Users\camila.melo.3\Desktop\aplicativo\AGHU_CONSUMO_ATUALIZADO.xlsx"

COL_CODIGO = "Código AGHU"
COL_VALOR  = "Preço"
COL_FLAG   = "Encontrado no AGHU"

ID_CAMPO_MATERIAL = "sbMaterial:sbMaterial:suggestion_input"
XPATH_CAMPO_MATERIAL_ABS = (
    "/html/body/div[1]/div[1]/form[3]/span/div[1]/div/div[1]/div/div[2]/div[1]/span/input[1]"
)

# botão de limpar (X)
XPATH_CLEAR_BUTTON = (
    "/html/body/div[1]/div[1]/form[3]/span/div[1]/div/div[1]/div/div[2]/div[3]/button/span[1]"
)

XPATH_BOTAO_PESQUISAR = '//*[@id="bt_pesquisar:button"]'
XPATH_VALOR_ULTIMA_COMPRA = (
    '//*[@id="valorUnitárioUltimaCompra:valorUnitárioUltimaCompra:inputId"]'
)

# índice fixo do iframe onde está a tela (0 = primeiro, 1 = segundo, etc.)
FRAME_INDEX_MATERIAL = 1

# XPaths de login (mesmos do teste_login_aghu.py)
XPATH_AGHU_USUARIO = '//*[@id="usuario:usuario:inputId"]'
XPATH_AGHU_SENHA = '//*[@id="password:inputId"]'
XPATH_AGHU_ENTRAR = '/html/body/div[1]/div/div/div/div/form/fieldset/div[4]/button/span'


# ======================================
# FUNÇÃO PARA QUEBRAR TELA "NÃO SEGURO"
# ======================================
def bypass_cert_warning(driver, timeout=5):
    try:
        wait_local = WebDriverWait(driver, timeout)
        details = wait_local.until(
            EC.element_to_be_clickable((By.ID, "details-button"))
        )
        details.click()
        proceed = wait_local.until(
            EC.element_to_be_clickable((By.ID, "proceed-link"))
        )
        proceed.click()
        print("⚠️ Tela de certificado inválido detectada e ignorada.")
    except TimeoutException:
        print("✅ Nenhuma tela de certificado para ignorar (ou já foi ignorada).")
    except Exception as e:
        print(f"⚠️ Não foi possível tratar a tela de certificado: {e}")


# ======================================
# LOGIN AGHU AUTOMÁTICO (reaproveitando o teste)
# ======================================
def fazer_login_aghu(driver, usuario: str, senha: str, timeout: int = 20):
    """
    Preenche usuário/senha e clica em 'Entrar' na tela de login do AGHU.
    Assume que a página de login já está carregada.
    """
    if not usuario or not senha:
        raise ValueError("Usuário e senha do AGHU são obrigatórios para login automático.")

    wait = WebDriverWait(driver, timeout)

    # usuário
    inp_user = wait.until(
        EC.presence_of_element_located((By.XPATH, XPATH_AGHU_USUARIO))
    )
    inp_user.clear()
    inp_user.send_keys(usuario)

    # senha
    inp_pass = wait.until(
        EC.presence_of_element_located((By.XPATH, XPATH_AGHU_SENHA))
    )
    inp_pass.clear()
    inp_pass.send_keys(senha)

    # botão entrar
    btn = wait.until(
        EC.element_to_be_clickable((By.XPATH, XPATH_AGHU_ENTRAR))
    )
    btn.click()

    # espera o header da aplicação aparecer (mesmo critério que você usava)
    wait.until(EC.presence_of_element_located((By.XPATH, "//header//ul")))
    print("✅ Login AGHU realizado com sucesso.")


# ======================================
# MUDAR PARA O FRAME FIXO DO MATERIAL
# ======================================
def switch_to_material_frame(driver) -> bool:
    """
    Vai direto para o iframe de índice FRAME_INDEX_MATERIAL.
    Não varre nem espera nada, só pega o frame pelo índice.
    """
    driver.switch_to.default_content()
    frames = driver.find_elements(By.TAG_NAME, "iframe")

    if len(frames) <= FRAME_INDEX_MATERIAL:
        print(f"❌ Frame índice {FRAME_INDEX_MATERIAL} não existe. Total de frames: {len(frames)}")
        return False

    driver.switch_to.frame(frames[FRAME_INDEX_MATERIAL])
    return True


# ======================================
# FUNÇÃO PRINCIPAL: RODAR AGHU
# ======================================
def rodar_aghu(usuario: str, senha: str):
    """
    Lê a planilha de entrada, faz login automático no AGHU,
    busca o preço (valor unitário da última compra)
    e salva a planilha de saída com as colunas 'Preço' e
    'Encontrado no AGHU' atualizadas.
    """

    # ---------- ABRIR PLANILHA ----------
    df = pd.read_excel(EXCEL_ENTRADA)

    if COL_VALOR not in df.columns:
        df[COL_VALOR] = ""
    if COL_FLAG not in df.columns:
        df[COL_FLAG] = ""

    # ---------- INICIAR NAVEGADOR ----------
    options = webdriver.ChromeOptions()
    options.add_argument("--ignore-certificate-errors")
    options.set_capability("acceptInsecureCerts", True)

    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    wait = WebDriverWait(driver, 20)

    # Abre a URL, trata certificado e faz login
    driver.get(URL)
    bypass_cert_warning(driver)
    fazer_login_aghu(driver, usuario, senha)

    # ---------- NAVEGAÇÃO: SUPRIMENTOS → ESTOQUES → CONSULTA → ESTATÍSTICA DE CONSUMO ----------
    try:
        suprimentos = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "/html/body/header/div[2]/ul/li[3]/a/span")
            )
        )
        suprimentos.click()
        time.sleep(0.2)

        estoques = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "/html/body/header/div[2]/ul/li[3]/ul/li/a/span")
            )
        )
        estoques.click()
        time.sleep(0.2)

        consulta = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "/html/body/header/div[2]/ul/li[3]/ul/li/ul/li[4]/a")
            )
        )
        consulta.click()
        time.sleep(0.2)

        estatistica = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "/html/body/header/div[2]/ul/li[3]/ul/li/ul/li[4]/ul/li[3]/a")
            )
        )
        estatistica.click()
        print("📊 Tela 'Estatística de consumo' aberta.")
        time.sleep(0.7)

    except TimeoutException:
        print("❌ Não consegui navegar no menu.")
        driver.quit()
        raise SystemExit("Verifique se o caminho do menu continua o mesmo.")

    # ---------- LOOP DOS MATERIAIS (USANDO FRAME FIXO) ----------
    """
    for idx, row in df.iterrows():
        codigo = str(row[COL_CODIGO]).strip()
        print(f"\n🔎 Buscando Código AGHU: {codigo}")

        try:
            # sempre troca para o frame fixo, sem procurar
            if not switch_to_material_frame(driver):
                df.at[idx, COL_VALOR] = 1
                df.at[idx, COL_FLAG] = "NÃO"
                print("   ⚠️ Erro de frame → marcado como NÃO encontrado (valor 1).")
                continue

            # 1) Clicar no botão de limpar (X)
            try:
                clear_btn = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((By.XPATH, XPATH_CLEAR_BUTTON))
                )
                clear_btn.click()
                time.sleep(0.1)
                print("   🧹 Campo limpo pelo botão 'X'.")
            except TimeoutException:
                print("   ℹ️ Botão 'X' não encontrado (provavelmente primeiro item).")

            # 2) Campo do material
            try:
                campo = WebDriverWait(driver, 4).until(
                    EC.element_to_be_clickable((By.ID, ID_CAMPO_MATERIAL))
                )
            except TimeoutException:
                campo = WebDriverWait(driver, 4).until(
                    EC.element_to_be_clickable((By.XPATH, XPATH_CAMPO_MATERIAL_ABS))
                )

            campo.clear()
            campo.send_keys(codigo)
            campo.send_keys(Keys.ENTER)

            # 3) Botão pesquisar
            botao_pesquisar = WebDriverWait(driver, 6).until(
                EC.element_to_be_clickable((By.XPATH, XPATH_BOTAO_PESQUISAR))
            )
            botao_pesquisar.click()

            # 4) Valor da última compra
            valor_raw = WebDriverWait(driver, 6).until(
                EC.presence_of_element_located(
                    (By.XPATH, XPATH_VALOR_ULTIMA_COMPRA)
                )
            ).get_attribute("value")

            # trata valor: troca ponto por vírgula
            if valor_raw:
                valor = valor_raw.replace(".", ",")
            else:
                valor = ""

            df.at[idx, COL_VALOR] = valor
            df.at[idx, COL_FLAG] = "SIM"
            print(f"   💰 Encontrado: {valor}")

        except TimeoutException:
            df.at[idx, COL_VALOR] = 1
            df.at[idx, COL_FLAG] = "NÃO"
            print("   ❌ Não encontrado → valor 1.")
        except Exception as e:
            df.at[idx, COL_VALOR] = 1
            df.at[idx, COL_FLAG] = "NÃO"
            print(f"   ⚠️ Erro inesperado ({e}) → valor 1.")

        # pausa mínima entre itens
        time.sleep(0.2)
     """
    for idx, row in df.iterrows():
        codigo = str(row[COL_CODIGO]).strip()
        print(f"\n🔎 Buscando Código AGHU: {codigo}")

        sucesso_item = False

        # até 2 tentativas para o mesmo código (caso dê StaleElementReference)
        for tentativa in range(2):
            try:
                # sempre troca para o frame fixo, sem procurar
                if not switch_to_material_frame(driver):
                    raise RuntimeError("Erro de frame (FRAME_INDEX_MATERIAL fora do range)")

                # 1) Clicar no botão de limpar (X)
                try:
                    clear_btn = WebDriverWait(driver, 3).until(
                        EC.element_to_be_clickable((By.XPATH, XPATH_CLEAR_BUTTON))
                    )
                    clear_btn.click()
                    time.sleep(0.1)
                    print("   🧹 Campo limpo pelo botão 'X'.")
                except TimeoutException:
                    print("   ℹ️ Botão 'X' não encontrado (provavelmente primeiro item).")

                # 2) Campo do material
                try:
                    campo = WebDriverWait(driver, 4).until(
                        EC.element_to_be_clickable((By.ID, ID_CAMPO_MATERIAL))
                    )
                except TimeoutException:
                    campo = WebDriverWait(driver, 4).until(
                        EC.element_to_be_clickable((By.XPATH, XPATH_CAMPO_MATERIAL_ABS))
                    )

                campo.clear()
                campo.send_keys(codigo)
                campo.send_keys(Keys.ENTER)

                # 3) Botão pesquisar
                botao_pesquisar = WebDriverWait(driver, 6).until(
                    EC.element_to_be_clickable((By.XPATH, XPATH_BOTAO_PESQUISAR))
                )
                botao_pesquisar.click()

                # 4) Valor da última compra
                valor_raw = WebDriverWait(driver, 6).until(
                    EC.presence_of_element_located(
                        (By.XPATH, XPATH_VALOR_ULTIMA_COMPRA)
                    )
                ).get_attribute("value")

                valor = valor_raw.replace(".", ",") if valor_raw else ""

                df.at[idx, COL_VALOR] = valor
                df.at[idx, COL_FLAG] = "SIM"
                print(f"   💰 Encontrado: {valor}")

                sucesso_item = True
                break  # sai das tentativas para esse código

            except StaleElementReferenceException:
                print(f"   ⚠️ StaleElementReferenceException na tentativa {tentativa+1}. "
                      "Vou tentar de novo este item...")
                time.sleep(0.5)
                # volta para o for tentativa (tenta novamente)
                continue

            except TimeoutException:
                print("   ❌ Timeout em algum elemento (campo, botão ou valor).")
                break  # não adianta tentar de novo exatamente igual

            except Exception as e:
                print(f"   ⚠️ Erro inesperado ({e})")
                break

        # se depois das tentativas ainda não deu certo:
        if not sucesso_item:
            df.at[idx, COL_VALOR] = 1
            df.at[idx, COL_FLAG] = "NÃO"
            print("   → valor 1 (não encontrado ou erro após tentativas).")

        # pausa mínima entre itens
        time.sleep(0.2)   

    # ---------- SALVAR PLANILHA FINAL ----------
    driver.switch_to.default_content()
    df.to_excel(EXCEL_SAIDA, index=False)

    print("\n📁 Planilha salva com sucesso!")
    print("✅ Processo finalizado")

    driver.quit()

    # se você quiser usar no main.py
    return EXCEL_SAIDA


if __name__ == "__main__":
    # para rodar direto do terminal você pode passar usuário/senha por variável de ambiente
    user = os.getenv("AGHU_USER", "")
    pwd = os.getenv("AGHU_PASS", "")
    if not user or not pwd:
        raise SystemExit("Defina AGHU_USER e AGHU_PASS nas variáveis de ambiente ou chame rodar_aghu(usuario, senha).")
    rodar_aghu(user, pwd)
