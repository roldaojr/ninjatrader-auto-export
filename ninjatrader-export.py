import tomllib
from pywinauto.application import Application, ProcessNotFoundError, WindowSpecification


with open("settings.toml", "rb") as f:
    settings = tomllib.load(f)


NINJATRADER_PATH = settings.get(
    "NINJATRADER_PATH", r"C:\Program Files\NinjaTrader 8\bin\NinjaTrader.exe"
)
NINJATRADER_USERNAME = settings.get("NINJATRADER_USERNAME", "")
NINJATRADER_PASSWORD = settings.get("NINJATRADER_PASSWORD", "")
NINJASCRIPT_FILES = settings.get("NINJASCRIPT_FILES", [])
PROTECT_ASSEMBLY = settings.get("PROTECT_ASSEMBLY", False)
PRODUCT_NAME = settings.get("PRODUCT_NAME", "NinjaScriptPackage")
PRODUCT_VERSION = settings.get("PRODUCT_VERSION", "1.0.0.0")
OUTPUT_PATH = settings.get("OUTPUT_PATH", "NinjaScriptPackage.zip")


def main():
    if not NINJATRADER_PATH:
        raise Exception("NINJATRADER_PATH não definido.")
    if not NINJASCRIPT_FILES:
        raise Exception("NINJASCRIPT_FILES não definido.")

    print("Iniciando NinjaTrader...")
    # Conectar ao aplicativo já aberto ou iniciar
    app = Application(backend="uia")
    try:
        app.connect(path=NINJATRADER_PATH, timeout=5)
    except ProcessNotFoundError:
        app.start(NINJATRADER_PATH, timeout=20)

    welcome_win = app.window(title_re="Welcome")
    if welcome_win.exists():
        # estamos na tela de login inserir crendenciais
        realizar_login(app)
    # aguardar tela principal
    main_win = app.window(title="", control_type="Custom")
    main_win.wait("ready", timeout=60)
    print("Iniciando Exportação...")
    # abrir tela de exportar
    open_export_window(main_win)
    export_win = app.window(auto_id="ExportNinjaScriptWindow")
    export_win.wait("ready", timeout=10)
    # configurar exportação
    configure_export(export_win)
    select_dlg = export_win.child_window(title="Select", control_type="Window")
    select_ninjascript_files(select_dlg)
    # executar exportação
    do_export(export_win)
    # fechar
    export_win.close()
    print("Exportação finalizada...")


def realizar_login(win: WindowSpecification):
    if not NINJATRADER_USERNAME or not NINJATRADER_PASSWORD:
        raise Exception("NINJATRADER_USERNAME ou NINJATRADER_PASSWORD não definidos")
    print("Executando login...")
    win.child_window(auto_id="tbUserName").set_text(NINJATRADER_USERNAME)
    win.child_window(auto_id="passwordBox").set_text(NINJATRADER_PASSWORD)
    win.child_window(auto_id="btnLogin").click()


def open_export_window(win: WindowSpecification):
    win.menu_select("Tools->Export->NinjaScript Add-On...")


def configure_export(win: WindowSpecification):
    # Markar exportar compilado
    compile_check = win.child_window(
        auto_id="chkExportCompiledAssembly", control_type="CheckBox"
    )
    if not compile_check.get_toggle_state():
        compile_check.toggle()
    # Markar proteção de assembly
    protect_check = win.child_window(
        auto_id="chkProtectCompiledAssembly", control_type="CheckBox"
    )
    if not protect_check.get_toggle_state() and PROTECT_ASSEMBLY:
        protect_check.toggle()
    # Definir nome do produto
    win.child_window(auto_id="txtProduct", control_type="Edit").set_text(PRODUCT_NAME)
    # Definir versão
    version0, version1, version2, version3 = PRODUCT_VERSION.split(".")
    win.child_window(auto_id="txtVersion0", control_type="Edit").set_text(version0)
    win.child_window(auto_id="txtVersion1", control_type="Edit").set_text(version1)
    win.child_window(auto_id="txtVersion2", control_type="Edit").set_text(version2)
    win.child_window(auto_id="txtVersion3", control_type="Edit").set_text(version3)
    # apertar em adicionar
    win.child_window(auto_id="NinjaScriptAdd", control_type="Button").click()


def select_ninjascript_files(win: WindowSpecification):
    # encontrar listbox
    listbox: WindowSpecification = win.child_window(
        auto_id="lbExportItems", control_type="List"
    )
    prev_length = 0
    while True:
        prev_length = len(listbox.items())
        listbox.scroll(direction="down", amount="page", count=1)
        for item in listbox.descendants(control_type="CheckBox"):
            filename = item.window_text().replace(" - ", "/")
            if filename in NINJASCRIPT_FILES and not item.get_toggle_state():
                item.toggle()
        length = len(listbox.items())
        if length == prev_length:
            break
    win.child_window(auto_id="btnOk", control_type="Button").click()


def do_export(win: WindowSpecification):
    win.child_window(auto_id="btnExport", control_type="Button").click()
    export_dlg = win.child_window(best_match="ExportDialog")
    export_dlg.wait("ready", timeout=10)
    # inserir nome do arquivo
    filename_input = export_dlg.child_window(
        auto_id="FileNameControlHost", control_type="ComboBox"
    )
    filename_input.child_window(control_type="Edit").set_text(OUTPUT_PATH)
    export_dlg.child_window(auto_id="1", control_type="Button").click()
    overwrite_dlg = export_dlg.child_window(best_match="ExportWindow2")
    if overwrite_dlg.exists(timeout=5):
        overwrite_dlg.child_window(auto_id="6", control_type="Button").click()
    messagebox_button_click(win, "NTMessageBoxYesButton", timeout=20)
    messagebox_button_click(win, "OkButton", timeout=2)
    messagebox_button_click(win, "OkButton", timeout=20)


def messagebox_button_click(win: WindowSpecification, button_id: str, timeout=1):
    message_box = win.child_window(auto_id="NTMessageBox", control_type="Window")
    if message_box.wait("ready", timeout=timeout):
        message_box.child_window(best_match=button_id).click()


if __name__ == "__main__":
    main()
