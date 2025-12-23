import re
import time
import unicodedata
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException

# =========================================================
# CONFIG
# =========================================================
EXCEL_PATH = r"C:\Users\camila.melo.3\Desktop\aplicativo\AGHU_CONSUMO_ATUALIZADO.xlsx"
RELATORIO_SAIDA = r"C:\Users\camila.melo.3\Desktop\aplicativo\RELATORIO_IRP_RESULTADO.xlsx"

LOGIN_URL = "https://www.comprasnet.gov.br/seguro/loginPortalUASG.asp"
IRP_NUMERO = "155022 - 00080/2025"
MUNICIPIO_BUSCA = "recife"

DEFAULT_TIMEOUT = 25
KEEP_BROWSER_OPEN = True

# Colunas do Excel
COL_CATMAT = "código catmat"
COL_PRECO = "preço"
COL_UNIDADE = "unidade de fornecimento"
COL_QTD = "quantidade"

# IDs dos campos de login (iguais ao teste_login_comprasnet)
ID_CPF = "txtLogin"
ID_SENHA = "txtSenha"

# =========================================================
# XPATHS
# =========================================================
XPATH_CARD_IRP = "/html/body/app-root/div/main/app-dashboard-governo/div/app-hub-acesso-sistemas/div[2]/div/div/p-dataview/div/div/div/div[6]/app-redirect-sistemas/span/span/div/p[1]/img"

# Passo 4-7 (menus)
XPATH_MENU_IRP = "/html/body/div[2]/table/tbody/tr[2]/td/div[1]/a[6]"
XPATH_MENU_ABRIR_INTENCAO = "/html/body/div[2]/table/tbody/tr[2]/td/div[1]/a[8]"
XPATH_TABELA_LISTA_IRP = "/html/body/div[2]/table/tbody/tr[2]/td/div[3]/form/table/tbody/tr[3]/td/table"

# Loop itens - Tab Item
XPATH_TAB_ITEM = "/html/body/div[2]/table/tbody/tr[2]/td/div[3]/form/table[1]/tbody/tr[1]/td[3]"

# Primeiro clique para abrir a tela de busca (muitas vezes abre o CATMAT)
XPATH_BTN_INCLUIR_ITEM = "/html/body/div[2]/table/tbody/tr[2]/td/div[3]/form/table[1]/tbody/tr[2]/td/div[3]/table[2]/tbody/tr/td/input[1]"

# ✅ PASSO QUE FALTAVA (depois de clicar "item" input[3], a página atualiza e você clica nisso)
XPATH_BTN_APOS_ITEM_ABRIR = "/html/body/div[2]/table/tbody/tr[2]/td/div[3]/form/table[1]/tbody/tr[2]/td/div[3]/table[3]/tbody/tr/td/input[1]"

# CATMAT (Siasgnet)
XPATH_INPUT_CATMAT = "/html/body/app-root/div/main/app-busca/div[1]/div/div/div[4]/div/div/p-autocomplete/span/input"
XPATH_BTN_LUPA = "/html/body/app-root/div/main/app-busca/div[1]/div/div/div[4]/div/div/button"
XPATH_SELECT_UNIDADE = "/html/body/app-root/div/main/app-busca/app-detalhe-material-siasgnet-lote/div/div[2]/div[1]/div[2]/select"
XPATH_BTN_ADICIONAR = "/html/body/app-root/div/main/app-busca/app-detalhe-material-siasgnet-lote/div/div[2]/div[2]/p-table/div/div/table/tbody/tr/td[4]/button"
XPATH_BTN_CARRINHO = "/html/body/app-root/div/main/app-busca/div[1]/div/div/div[2]/div[2]/button/i"
XPATH_BTN_ADICIONAR_AO_SIASGNET = "/html/body/app-root/div/main/app-busca/app-exibir-selecionados-siasgnet-lote/div/div[1]/div/div[3]/button"

# OK popup
XPATH_BTN_OK_POPUP = "/html/body/form/div/fieldset/div[2]/input"

# Comprasnet - tabela itens / alterar
XPATH_TABELA_ITENS = "/html/body/div[2]/table/tbody/tr[2]/td/div[3]/form/table[1]/tbody/tr[2]/td/div[3]/table[2]"

# 🔸 Paginação da tabela de itens – botão "ir para a última página"
XPATH_PAGINACAO_ULTIMA_PAGINA = (
    "/html/body/div[2]/table/tbody/tr[2]/td/div[3]/form/table[1]/tbody/tr[2]/td/div[3]/div/span[2]/a[last()]/img"
)

# Alterar/preencher
XPATH_INPUT_PRECO = "/html/body/div[2]/table/tbody/tr[2]/td/div[3]/form/table[1]/tbody/tr[2]/td/div[3]/table[3]/tbody/tr[2]/td[1]/input"
XPATH_CHECK_MARCAR = "/html/body/div[2]/table/tbody/tr[2]/td/div[3]/form/table[1]/tbody/tr[2]/td/div[3]/table[3]/tbody/tr[2]/td[2]/div/input[1]"
XPATH_BTN_LOCALIZAR = "/html/body/div[2]/table/tbody/tr[2]/td/div[3]/form/table[1]/tbody/tr[2]/td/div[3]/fieldset/table/tbody/tr[2]/td[2]/input"

# Popup localidade
XPATH_POPUP_LOCAL_INPUT = "/html/body/table/tbody/tr/td/div[2]/form/table/tbody/tr/td/table[1]/tbody/tr[2]/td/input"
XPATH_POPUP_LOCAL_PESQUISAR = "/html/body/table/tbody/tr/td/div[2]/form/table/tbody/tr/td/table[2]/tbody/tr/td/input[1]"
XPATH_POPUP_LOCAL_SELECIONAR = "/html/body/table/tbody/tr/td/div[2]/form/table/tbody/tr[2]/td/table/tbody/tr/td[3]/a"

# Quantidade / incluir / salvar
XPATH_INPUT_QTD = "/html/body/div[2]/table/tbody/tr[2]/td/div[3]/form/table[1]/tbody/tr[2]/td/div[3]/fieldset/table/tbody/tr[2]/td[3]/input"
XPATH_BTN_INCLUIR_QTD = "/html/body/div[2]/table/tbody/tr[2]/td/div[3]/form/table[1]/tbody/tr[2]/td/div[3]/fieldset/table/tbody/tr[2]/td[4]/input"
XPATH_BTN_SALVAR_ITEM = "/html/body/div[2]/table/tbody/tr[2]/td/div[3]/form/table[1]/tbody/tr[2]/td/div[3]/table[4]/tbody/tr/td/input[1]"

# Botão "item" após salvar (para seguir o loop)
XPATH_BTN_ITEM_DEPOIS_SALVAR = "/html/body/div[2]/table/tbody/tr[2]/td/div[3]/form/table[1]/tbody/tr[2]/td/div[3]/table[4]/tbody/tr/td/input[3]"


# =========================================================
# HELPERS COMUNS
# =========================================================
def log(msg: str):
    print(f"[{time.strftime('%H:%M:%S')}] {msg}")


def wait_click(driver, xpath: str, timeout: int = DEFAULT_TIMEOUT):
    el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.XPATH, xpath)))
    el.click()
    return el


def wait_presence(driver, xpath: str, timeout: int = DEFAULT_TIMEOUT):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, xpath)))


def type_clear(driver, xpath: str, text: str, timeout: int = DEFAULT_TIMEOUT):
    el = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, xpath)))
    el.clear()
    el.send_keys(text)
    return el


def accept_alert_if_any(driver) -> bool:
    try:
        driver.switch_to.alert.accept()
        return True
    except Exception:
        return False


def click_ok_popup(driver) -> bool:
    if accept_alert_if_any(driver):
        return True
    try:
        WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, XPATH_BTN_OK_POPUP))).click()
        return True
    except Exception:
        return False


def bypass_privacy_error_chrome(driver) -> bool:
    """Avançadas -> Prosseguir; fallback thisisunsafe."""
    time.sleep(0.6)
    try:
        src = (driver.page_source or "").lower()
        if ("sua conexão não é particular" not in src) and ("err_cert" not in src) and ("não seguro" not in src):
            return False
    except Exception:
        pass

    try:
        WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, "details-button"))).click()
        time.sleep(0.2)
    except Exception:
        pass

    try:
        WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, "proceed-link"))).click()
        time.sleep(0.8)
        return True
    except Exception:
        pass

    try:
        driver.switch_to.active_element.send_keys("thisisunsafe")
        time.sleep(0.8)
        return True
    except Exception:
        return False


def switch_to_new_tab_if_any(driver, before_handles: List[str], timeout: int = 10) -> bool:
    """Se clicar algo e abrir nova aba, troca para ela."""
    end = time.time() + timeout
    while time.time() < end:
        now = driver.window_handles
        if len(now) > len(before_handles):
            newh = [h for h in now if h not in before_handles]
            if newh:
                driver.switch_to.window(newh[0])
                return True
        time.sleep(0.2)
    return False


def ensure_on_work_area_after_irp_click(driver):
    """Garante estar na tela onde existem os menus /a[6] e /a[8]."""
    try:
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, XPATH_MENU_IRP)))
        return
    except Exception:
        pass

    driver.switch_to.window(driver.window_handles[-1])
    time.sleep(0.6)
    bypass_privacy_error_chrome(driver)
    WebDriverWait(driver, DEFAULT_TIMEOUT).until(EC.presence_of_element_located((By.XPATH, XPATH_MENU_IRP)))


def norm_text(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.strip().lower().replace(",", ".")
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s).strip()
    s = s.replace(" mililitro", " ml").replace(" mililitros", " ml")
    s = s.replace(" litro", " l").replace(" litros", " l")
    return s


def format_preco_4casas(preco: Any) -> str:
    if preco is None or str(preco).strip() == "":
        return "0,0000"
    raw = str(preco).strip()
    raw = raw.replace(".", "").replace(",", ".")
    try:
        v = float(raw)
    except Exception:
        v = 0.0
    return f"{v:.4f}".replace(".", ",")


def click_ok_item_e_abrir_proximo(driver, base_handle: str, timeout: int = DEFAULT_TIMEOUT):
    """
    Depois de SALVAR -> OK -> clicar ITEM (input[3]) -> página atualiza -> 
    clicar BTN_APOS_ITEM_ABRIR (input[1]) para iniciar o próximo item.
    """
    # OK
    _ = click_ok_popup(driver)

    # garante foco na janela base
    try:
        driver.switch_to.window(base_handle)
    except Exception:
        if driver.window_handles:
            driver.switch_to.window(driver.window_handles[0])

    # clicar ITEM (input[3]) com retry/stale-safe
    end = time.time() + timeout
    last_err = None
    while time.time() < end:
        try:
            btn_item = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, XPATH_BTN_ITEM_DEPOIS_SALVAR))
            )
            btn_item.click()
            break
        except (StaleElementReferenceException, TimeoutException) as e:
            last_err = e
            time.sleep(0.4)
    else:
        raise TimeoutException(
            f"Não consegui clicar no botão ITEM (input[3]) após OK. Último erro: {last_err}"
        )

    # após clicar ITEM, a página atualiza e você precisa clicar no BTN_APOS_ITEM_ABRIR (input[1])
    before = driver.window_handles[:]
    end = time.time() + timeout
    last_err = None
    while time.time() < end:
        try:
            btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, XPATH_BTN_APOS_ITEM_ABRIR))
            )
            btn.click()
            break
        except (StaleElementReferenceException, TimeoutException) as e:
            last_err = e
            time.sleep(0.4)
    else:
        raise TimeoutException(
            f"Não consegui clicar no botão após ITEM (input[1]). Último erro: {last_err}"
        )

    # se abriu nova aba/janela do CATMAT, troca
    switch_to_new_tab_if_any(driver, before, timeout=8)
    bypass_privacy_error_chrome(driver)


def is_on_catmat_page(driver) -> bool:
    """Detecta se já estamos na tela do CATMAT (input do autocomplete)."""
    try:
        driver.find_element(By.XPATH, XPATH_INPUT_CATMAT)
        return True
    except Exception:
        return False


# =========================================================
# LOGIN COMPRASNET AUTOMÁTICO
# =========================================================
def login_comprasnet(driver, cpf: str, senha: str, timeout: int = DEFAULT_TIMEOUT):
    """
    Usa os campos txtLogin/txtSenha para fazer login automático no ComprasNet.
    """
    if not cpf or not senha:
        raise ValueError("CPF e senha do ComprasNet são obrigatórios para login automático.")

    wait = WebDriverWait(driver, timeout)

    # campo CPF
    input_cpf = wait.until(EC.presence_of_element_located((By.ID, ID_CPF)))
    input_cpf.clear()
    input_cpf.send_keys(cpf)

    # campo senha
    input_senha = wait.until(EC.presence_of_element_located((By.ID, ID_SENHA)))
    input_senha.clear()
    input_senha.send_keys(senha)

    # botão Entrar no mesmo form
    try:
        form = input_cpf.find_element(By.XPATH, "./ancestor::form")
        botoes = form.find_elements(By.TAG_NAME, "button")
        botao_entrar = None
        for btn in botoes:
            txt = (btn.text or "").strip().lower()
            tipo = (btn.get_attribute("type") or "").lower()
            if "entrar" in txt or tipo == "submit":
                botao_entrar = btn
                break

        if botao_entrar is None:
            raise TimeoutException("Não encontrei botão 'Entrar' no formulário de login do ComprasNet.")

        botao_entrar.click()
    except Exception as e:
        raise TimeoutException(f"Não consegui acionar o botão de login do ComprasNet: {e}")

    # depois do login, esperamos o card IRP ficar visível
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.XPATH, XPATH_CARD_IRP))
    )
    log("✅ Login ComprasNet realizado com sucesso.")


# =========================================================
# PLANILHA
# =========================================================
@dataclass
class ItemPlanilha:
    idx: int
    catmat: str
    preco: str
    unidade: str
    qtd: int


def load_planilha(path: str) -> List[ItemPlanilha]:
    df = pd.read_excel(path)
    df.columns = [str(c).strip().lower() for c in df.columns]
    missing = [c for c in [COL_CATMAT, COL_PRECO, COL_UNIDADE, COL_QTD] if c not in df.columns]
    if missing:
        raise ValueError(f"Colunas faltando no Excel: {missing}. Colunas encontradas: {list(df.columns)}")

    items: List[ItemPlanilha] = []
    for i, row in df.iterrows():
        catmat = str(row.get(COL_CATMAT, "")).strip()
        unidade = str(row.get(COL_UNIDADE, "")).strip()
        preco = format_preco_4casas(row.get(COL_PRECO, "0"))
        try:
            qtd = int(row.get(COL_QTD, 1))
        except Exception:
            qtd = 1
        if catmat:
            items.append(ItemPlanilha(idx=i, catmat=catmat, preco=preco, unidade=unidade, qtd=qtd))
    return items


# =========================================================
# IRP SELEÇÃO
# =========================================================
def selecionar_irp_na_tabela(driver, irp_numero: str):
    tabela = wait_presence(driver, XPATH_TABELA_LISTA_IRP, DEFAULT_TIMEOUT)
    alvo = re.sub(r"\s+", " ", str(irp_numero).strip())
    rows = tabela.find_elements(By.XPATH, ".//tbody/tr")
    for r in rows:
        txt = re.sub(r"\s+", " ", (r.text or "").strip())
        if alvo in txt:
            r.find_element(By.XPATH, ".//a[contains(.,'Selecionar') or contains(.,'selecionar')]").click()
            return
    raise TimeoutException(f"IRP '{irp_numero}' não encontrada na tabela.")


# =========================================================
# LOOP ITENS (resto do arquivo igual ao seu atual)
# =========================================================
def step_click_item_tab(driver):
    wait_click(driver, XPATH_TAB_ITEM, DEFAULT_TIMEOUT)
    time.sleep(0.4)


def step_incluir_item_abrir_catmat(driver):
    """
    Clica no incluir item e troca para nova aba se abrir.
    """
    before = driver.window_handles[:]
    wait_click(driver, XPATH_BTN_INCLUIR_ITEM, DEFAULT_TIMEOUT)
    time.sleep(0.6)
    switch_to_new_tab_if_any(driver, before, timeout=8)
    bypass_privacy_error_chrome(driver)


def step_catmat_add(driver, catmat: str, unidade: str) -> (bool, str):
    try:
        inp = wait_presence(driver, XPATH_INPUT_CATMAT, DEFAULT_TIMEOUT)
    except TimeoutException:
        return False, "Não adicionado porque o catmat não foi encontrado"

    inp.clear()
    inp.send_keys(str(catmat))
    wait_click(driver, XPATH_BTN_LUPA, DEFAULT_TIMEOUT)
    time.sleep(1.0)

    try:
        sel_el = wait_presence(driver, XPATH_SELECT_UNIDADE, DEFAULT_TIMEOUT)
    except TimeoutException:
        return False, "Não adicionado porque a unidade de fornecimento não foi encontrada"

    sel = Select(sel_el)
    alvo = norm_text(unidade)

    chosen = None
    for opt in sel.options:
        if norm_text(opt.text) == alvo:
            chosen = opt.get_attribute("value")
            break
    if chosen is None:
        for opt in sel.options:
            if alvo and alvo in norm_text(opt.text):
                chosen = opt.get_attribute("value")
                break

    if chosen is None:
        return False, "Não adicionado porque a unidade de fornecimento não foi encontrada"

    sel.select_by_value(chosen)

    wait_click(driver, XPATH_BTN_ADICIONAR, DEFAULT_TIMEOUT)
    time.sleep(0.5)
    wait_click(driver, XPATH_BTN_CARRINHO, DEFAULT_TIMEOUT)
    time.sleep(0.5)
    wait_click(driver, XPATH_BTN_ADICIONAR_AO_SIASGNET, DEFAULT_TIMEOUT)
    time.sleep(0.8)

    click_ok_popup(driver)
    time.sleep(0.4)
    return True, "OK"


def voltar_para_comprasnet(driver, main_handle: str):
    """Volta pra janela principal após CATMAT."""
    try:
        driver.switch_to.window(main_handle)
        return
    except Exception:
        pass
    if driver.window_handles:
        driver.switch_to.window(driver.window_handles[0])


def go_to_last_page_items_table(driver):
    """
    Se existir paginação (mais de 20 itens), clica no botão de ir para a última página.
    Se não existir (menos de 20 itens), simplesmente segue na página atual.
    """
    try:
        log("Tentando ir para a última página da tabela de itens (se existir)...")
        btn_last = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.XPATH, XPATH_PAGINACAO_ULTIMA_PAGINA))
        )
        btn_last.click()
        time.sleep(0.8)  # dá tempo da página recarregar
    except TimeoutException:
        log("Paginação não encontrada ou não há última página; permanecendo na página atual.")
    except Exception as e:
        log(f"Falha ao tentar mudar para a última página da tabela de itens: {e}")


def step_click_last_alterar(driver):
    """
    Vai para a última página da tabela (se houver várias páginas) e
    clica no último link 'Alterar' da tabela de itens.
    """
    go_to_last_page_items_table(driver)

    tbl = wait_presence(driver, XPATH_TABELA_ITENS, DEFAULT_TIMEOUT)
    rows = tbl.find_elements(By.XPATH, ".//tbody/tr")
    if not rows:
        raise TimeoutException("Tabela de itens sem linhas.")
    rows[-1].find_element(By.XPATH, ".//td[9]/a").click()
    time.sleep(0.6)


def step_localidade_popup(driver, municipio: str):
    type_clear(driver, XPATH_POPUP_LOCAL_INPUT, municipio, DEFAULT_TIMEOUT)
    wait_click(driver, XPATH_POPUP_LOCAL_PESQUISAR, DEFAULT_TIMEOUT)
    time.sleep(0.6)
    wait_click(driver, XPATH_POPUP_LOCAL_SELECIONAR, DEFAULT_TIMEOUT)
    time.sleep(0.6)


def step_preencher_salvar_e_preparar_proximo(driver, preco: str, municipio: str, qtd: int):
    """
    Preenche e salva; depois faz:
    OK -> ITEM (input[3]) -> BTN_APOS_ITEM_ABRIR (input[1])
    """
    type_clear(driver, XPATH_INPUT_PRECO, preco, DEFAULT_TIMEOUT)

    cb = WebDriverWait(driver, DEFAULT_TIMEOUT).until(
        EC.element_to_be_clickable((By.XPATH, XPATH_CHECK_MARCAR))
    )
    if not cb.is_selected():
        cb.click()

    base_handles = driver.window_handles[:]
    base_handle = driver.current_window_handle

    wait_click(driver, XPATH_BTN_LOCALIZAR, DEFAULT_TIMEOUT)
    time.sleep(0.8)

    if len(driver.window_handles) > len(base_handles):
        driver.switch_to.window(driver.window_handles[-1])

    step_localidade_popup(driver, municipio)

    try:
        driver.switch_to.window(base_handle)
    except Exception:
        if driver.window_handles:
            driver.switch_to.window(driver.window_handles[0])

    type_clear(driver, XPATH_INPUT_QTD, str(qtd), DEFAULT_TIMEOUT)
    wait_click(driver, XPATH_BTN_INCLUIR_QTD, DEFAULT_TIMEOUT)
    time.sleep(0.3)

    wait_click(driver, XPATH_BTN_SALVAR_ITEM, DEFAULT_TIMEOUT)
    time.sleep(0.8)

    click_ok_item_e_abrir_proximo(driver, base_handle=base_handle, timeout=DEFAULT_TIMEOUT)
    time.sleep(0.4)


# =========================================================
# MAIN
# =========================================================
#def rodar_irp(excel_path: Optional[str] = None, cpf: str = "", senha: str = ""):
def rodar_irp(
    excel_path: Optional[str] = None,
    cpf: str = "",
    senha: str = "",
    irp_numero: str = "",
):

    if excel_path is not None:
        global EXCEL_PATH
        EXCEL_PATH = excel_path

    if not cpf or not senha:
        raise ValueError("CPF e senha do ComprasNet são obrigatórios para rodar a IRP automaticamente.")
    
    # Se o usuário não passar uma IRP, usa o padrão definido no topo do arquivo
    global IRP_NUMERO
    if irp_numero:
        IRP_NUMERO = irp_numero

    items = load_planilha(EXCEL_PATH)

    prev_ok_catmat: Optional[str] = None
    prev_ok_unidade: Optional[str] = None

    report: List[Dict[str, Any]] = []

    opts = webdriver.ChromeOptions()
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-popup-blocking")
    opts.add_argument("--ignore-certificate-errors")
    opts.add_argument("--allow-insecure-localhost")

    driver = webdriver.Chrome(options=opts)
    driver.set_page_load_timeout(60)

    try:
        # 1) login automático
        driver.get(LOGIN_URL)
        bypass_privacy_error_chrome(driver)
        login_comprasnet(driver, cpf, senha)

        # 2) clicar no card IRP (pode abrir nova aba)
        before = driver.window_handles[:]
        wait_click(driver, XPATH_CARD_IRP, DEFAULT_TIMEOUT)
        time.sleep(1.0)

        switch_to_new_tab_if_any(driver, before, timeout=10)
        bypass_privacy_error_chrome(driver)
        ensure_on_work_area_after_irp_click(driver)

        # 3) menus /a[6] e /a[8]
        wait_click(driver, XPATH_MENU_IRP, DEFAULT_TIMEOUT)
        wait_click(driver, XPATH_MENU_ABRIR_INTENCAO, DEFAULT_TIMEOUT)

        # 4) selecionar IRP correta na tabela
        selecionar_irp_na_tabela(driver, IRP_NUMERO)

        # 5) LOOP
        main_handle = driver.current_window_handle

        for idx, it in enumerate(items):
            row = {
                "linha_excel": it.idx,
                "catmat_planilha": it.catmat,
                "unidade_planilha": it.unidade,
                "preco_planilha": it.preco,
                "qtd_planilha": it.qtd,
                "catmat_adicionado": "",
                "unidade_adicionada": "",
                "status": "",
                "motivo": "",
            }

            try:
                if not is_on_catmat_page(driver):
                    step_click_item_tab(driver)
                    step_incluir_item_abrir_catmat(driver)

                ok, motivo = step_catmat_add(driver, it.catmat, it.unidade)

                if ok:
                    row["catmat_adicionado"] = it.catmat
                    row["unidade_adicionada"] = it.unidade
                    prev_ok_catmat = it.catmat
                    prev_ok_unidade = it.unidade
                else:
                    if prev_ok_catmat and prev_ok_unidade:
                        ok2, motivo2 = step_catmat_add(driver, prev_ok_catmat, prev_ok_unidade)
                        if ok2:
                            row["catmat_adicionado"] = prev_ok_catmat
                            row["unidade_adicionada"] = prev_ok_unidade
                            row["motivo"] = motivo
                            ok = True
                        else:
                            row["status"] = "NAO_ADICIONADO"
                            row["motivo"] = motivo if motivo else motivo2
                            report.append(row)
                            continue
                    else:
                        row["status"] = "NAO_ADICIONADO"
                        row["motivo"] = motivo
                        report.append(row)
                        continue

                voltar_para_comprasnet(driver, main_handle)

                step_click_last_alterar(driver)

                step_preencher_salvar_e_preparar_proximo(driver, it.preco, MUNICIPIO_BUSCA, it.qtd)

                row["status"] = "ADICIONADO"
                if not row["motivo"]:
                    row["motivo"] = "OK"
                report.append(row)

            except Exception as e:
                row["status"] = "ERRO"
                row["motivo"] = f"{type(e).__name__}: {e}"
                report.append(row)
                click_ok_popup(driver)

        pd.DataFrame(report).to_excel(RELATORIO_SAIDA, index=False)
        log(f"Relatório salvo em: {RELATORIO_SAIDA}")

    finally:
        if not KEEP_BROWSER_OPEN:
            driver.quit()
        else:
            log("KEEP_BROWSER_OPEN=True -> navegador ficará aberto para inspeção.")


if __name__ == "__main__":
    # para rodar direto, você pode preencher CPF/SENHA aqui ou pegar de variável de ambiente
    import os

    cpf_env = os.getenv("COMPRASNET_CPF", "")
    senha_env = os.getenv("COMPRASNET_SENHA", "")
    if not cpf_env or not senha_env:
        raise SystemExit("Defina COMPRASNET_CPF e COMPRASNET_SENHA nas variáveis de ambiente ou chame rodar_irp(cpf=..., senha=...).")
    rodar_irp(cpf=cpf_env, senha=senha_env)
