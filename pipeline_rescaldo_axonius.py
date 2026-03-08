import os
import subprocess
import sys
import time
from contextlib import contextmanager
from datetime import date, datetime
from getpass import getuser

import pandas as pd
import win32com.client as win32
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
from rich.console import Console

# Carrega variáveis de ambiente do arquivo .env
load_dotenv()

# ======================================================
# CONFIGURAÇÕES GERAIS
# ======================================================

console = Console(log_time=False, log_path=False, markup=True)

# URLs e IDs de Execução
AXONIUS_URL = os.getenv("AXONIUS_URL", "https://sua-url-axonius.corp")
AZURE_RESOURCE_MANAGER_URL = (
    "https://portal.azure.com/#view/HubsExtension/ServiceMenuBlade/~/"
    "resourcegraphexplorer/extension/Microsoft_Azure_Resources/menuId/"
    "ResourceManager/itemId/resourcegraphexplorer"
)
QUERIES_AXONIUS = {
    "Bitlocker": os.getenv("AXONIUS_QUERY_BITLOCKER", "({{QueryID=Exemplo_Bitlocker}})"),
    "Trellix": os.getenv("AXONIUS_QUERY_TRELLIX", "({{QueryID=Exemplo_Trellix}})"),
    "Zscaler": os.getenv("AXONIUS_QUERY_ZSCALER", "({{QueryID=Exemplo_Zscaler}})"),
}

# Caminhos base
EDGE_USER_DATA_DIR = os.getenv(
    "EDGE_USER_DATA_DIR", 
    rf"C:\Users\{getuser()}\AppData\Local\Microsoft\Edge\User Data\Profile 1"
)
BASE_DIR = os.getenv("BASE_DIR", r"C:\Caminho\Para\Documentos\Rescaldo Axonius")

# Diretórios
AXONIUS_EXTRACTIONS_DIR = os.path.join(BASE_DIR, "00_Entrada_Dados", "Axonius")
BASES_DIR = os.path.join(BASE_DIR, "00_Entrada_Dados", "Bases")
PROCESSAMENTO_DIR = os.path.join(BASE_DIR, "01_Processamento")
ATUACAO_DIR = os.path.join(BASE_DIR, "02_Atuação")
HISTORICO_DIR = os.path.join(BASE_DIR, "03_Histórico", "snapshots")

# Arquivos
AXONIUS_PROCESSAMENTO_XLSX = os.path.join(
    PROCESSAMENTO_DIR, "Axonius_Processamento.xlsx"
)
ATUACAO_XLSX = os.path.join(ATUACAO_DIR, "Atuação.xlsx")
HISTORICO_PROCESSAMENTO_XLSX = os.path.join(
    BASE_DIR, "03_Histórico", "Histórico_Processamento.xlsx"
)

PS_EXE_64 = r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"

# ======================================================
# UTILITÁRIOS
# ======================================================


def info(msg: str):
    console.print(f"[cyan][INFO][/cyan] {msg}")


def ok(msg: str):
    console.print(f"[green][OK][/green] {msg}")


def warn(msg: str):
    console.print(f"[yellow][WARN][/yellow] {msg}")


def error(msg: str):
    console.print(f"[bold red][ERRO][/bold red] {msg}")


@contextmanager
def etapa_status(titulo: str):
    """Exibe um spinner indicando a execução da etapa atual."""
    with console.status(f"[bold cyan]{titulo}...", spinner="dots"):
        yield


# ======================================================
# FLUXOS PRINCIPAIS
# ======================================================


def validar_pre_requisitos():
    """Valida se todas as pastas e arquivos base existem (Fail-Fast)."""
    if not os.path.exists(BASE_DIR):
        raise FileNotFoundError(f"Diretório base não encontrado: {BASE_DIR}")

    for planilha in [
        AXONIUS_PROCESSAMENTO_XLSX,
        ATUACAO_XLSX,
        HISTORICO_PROCESSAMENTO_XLSX,
    ]:
        if not os.path.exists(planilha):
            raise FileNotFoundError(f"Planilha obrigatória não encontrada: {planilha}")

    ok("Pré-requisitos estruturais validados.")


def extrair_ad() -> str:
    """Extrai computadores do AD e salva na pasta de bases, limpando rastros locais."""
    info("Exportando base do AD...")
    arquivo_saida = os.path.join(BASES_DIR, "admaquina.csv")

    # Referencia o script PowerShell dentro da pasta execution/ do projeto
    current_dir = os.path.dirname(os.path.abspath(__file__))
    ps1 = os.path.join(current_dir, "execution", "export_ad.ps1")

    if not os.path.exists(ps1):
        raise FileNotFoundError(f"Script de execução não encontrado: {ps1}")

    try:
        subprocess.run(
            [
                PS_EXE_64,
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                ps1,
                "-OutputPath",
                arquivo_saida,
            ],
            check=True,
            capture_output=True,
        )
        ok("Base do AD exportada com sucesso.")
        return arquivo_saida
    except subprocess.CalledProcessError as e:
        error(f"Erro na execução do script PowerShell: {e.stderr.decode('utf-8', errors='ignore') if e.stderr else e}")
        raise


def _exportar_query_axonius(page, query_id: str, nome_arquivo: str):
    """Rotina isolada para buscar de forma passiva os relatórios no Axonius."""
    info(f"Exportando {nome_arquivo}...")
    try:
        search = page.get_by_role("textbox", name="Search for assets or saved")
        search.wait_for(state="visible", timeout=10000)
        search.fill("")
        search.fill(query_id)
        page.keyboard.press("Enter")

        btn_export = page.get_by_role("button", name="export table to csv")
        btn_export.wait_for(timeout=10000)
        btn_export.click()

        with page.expect_download(timeout=20000) as download_info:
            page.get_by_role("button", name="Export", exact=True).click()

        destino = os.path.join(AXONIUS_EXTRACTIONS_DIR, nome_arquivo)
        os.makedirs(os.path.dirname(destino), exist_ok=True)
        download_info.value.save_as(destino)
        ok(f"{nome_arquivo} salvo com sucesso.")
    except Exception as e:
        error(f"Falha na exportação de {nome_arquivo}: {e}")
        raise


def _exportar_vms_azure(page):
    """Extração dedicada das VDI Base no Azure Resource Manager."""
    info("Exportando Azure VMs...")
    try:
        page.goto(AZURE_RESOURCE_MANAGER_URL, wait_until="domcontentloaded")

        btn_open_query = page.get_by_role("button", name="Open a query")
        btn_open_query.wait_for(timeout=15000)
        btn_open_query.click()

        page.get_by_role("link", name="Get-All-VMs").click()
        page.get_by_role("button", name="Run query").click()

        with page.expect_download(timeout=20000) as download_info:
            page.get_by_role("button", name="Download as CSV").click()

        caminho = os.path.join(BASES_DIR, "vdis_azure.csv")
        os.makedirs(os.path.dirname(caminho), exist_ok=True)
        download_info.value.save_as(caminho)
        ok("Azure VMs exportado com sucesso.")
    except Exception as e:
        error(f"Falha na exportação do Azure VMs: {e}")
        raise


def extrair_bases_web():
    """Orquestra as requisições da persistência local de navegador e faz os retornos determinísticos."""
    info("Iniciando navegador web...")
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=EDGE_USER_DATA_DIR,
            channel="msedge",
            headless=False,
            accept_downloads=True,
            args=["--no-first-run"],
        )
        try:
            page = context.pages[0] if context.pages else context.new_page()

            # Step 1: Axonius
            info("Acessando Axonius...")
            page.goto(AXONIUS_URL, wait_until="domcontentloaded")
            page.get_by_test_id("assets").wait_for(timeout=10000)
            page.get_by_test_id("assets").click()

            for nome, query_id in QUERIES_AXONIUS.items():
                _exportar_query_axonius(page, query_id, f"Extração_{nome}.csv")

            # Step 2: Azure VMs
            _exportar_vms_azure(page)
        finally:
            context.close()


def atualizar_power_query(workbook_path: str):
    """Executa a atualização de conexões desativando o BackgroundQuery para evitar concorrência ou crashs."""
    nome_arquivo = os.path.basename(workbook_path)
    info(f"Atualizando conexões em lote para o arquivo: {nome_arquivo}")

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(workbook_path)

        for conn in wb.Connections:
            try:
                if conn.Type == 1:  # OLEDB
                    conn.OLEDBConnection.BackgroundQuery = False
                elif conn.Type == 2:  # ODBC
                    conn.ODBCConnection.BackgroundQuery = False
            except Exception as e:
                warn(
                    f"Falha não-deslizante durante o parsing condicional da connection ({conn.Name}): {e}"
                )

        wb.RefreshAll()

        # Hook assíncrono básico aguardando cálculo final
        while excel.CalculationState != 0:  # 0 = xlDone
            time.sleep(0.5)

        wb.Save()
        wb.Close()
        ok(f"Conexões atualizadas em {nome_arquivo}.")
    finally:
        excel.Quit()


def gerar_snapshot_diario():
    """Lê a persistência estática do Excel Processado gerando logs diários blindados (snapshot)."""
    hoje = date.today().isoformat()
    arquivo_snapshot = os.path.join(HISTORICO_DIR, f"snapshot_{hoje}.csv")

    if os.path.exists(arquivo_snapshot):
        info(f"Snapshot omitido (já gerado na data atual {hoje}).")
        return

    info("Gerando snapshot diário...")
    df = pd.read_excel(
        AXONIUS_PROCESSAMENTO_XLSX,
        sheet_name="Extração_Axonius_Processada",
        engine="openpyxl",
    )
    df.columns = [str(c).strip() for c in df.columns]
    df["data_snapshot"] = hoje

    os.makedirs(os.path.dirname(arquivo_snapshot), exist_ok=True)
    df.to_csv(arquivo_snapshot, index=False, encoding="utf-8-sig")
    ok(f"Snapshot finalizado e salvo em: {arquivo_snapshot}")


def main():
    inicio_total = datetime.now()
    console.rule("[bold cyan]PIPELINE RESCALDO AXONIUS[/bold cyan]")

    try:
        with etapa_status("Validação de Pré-requisitos"):
            validar_pre_requisitos()

        with etapa_status("Exportando bases de Web Apps"):
            extrair_bases_web()

        with etapa_status("Exportando base do AD via Script PS1"):
            extrair_ad()

        with etapa_status("Atualizando PQ - Processamento"):
            atualizar_power_query(AXONIUS_PROCESSAMENTO_XLSX)

        with etapa_status("Atualizando PQ - Atuação"):
            atualizar_power_query(ATUACAO_XLSX)

        with etapa_status("Atualizando PQ - Histórico"):
            atualizar_power_query(HISTORICO_PROCESSAMENTO_XLSX)

        with etapa_status("Gerando Registro Histórico snapshot"):
            gerar_snapshot_diario()

        console.rule("[bold green]PROCESSO FINALIZADO COM SUCESSO[/bold green]")

    except Exception as e:
        error(f"Erro central diagnosticado e com pipeline abortado: {e}")
        console.rule("[bold red]PROCESSO FINALIZADO COM ERROS[/bold red]")
        sys.exit(1)

    finally:
        duracao = (datetime.now() - inicio_total).total_seconds()
        mins, secs = divmod(int(duracao), 60)
        console.print(
            f"[bold green]Tempo total decorrido:[/bold green] {mins}m {secs}s"
        )
        time.sleep(3)


if __name__ == "__main__":
    main()
