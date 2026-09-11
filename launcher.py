# -*- coding: utf-8 -*-
"""
Arranque do Arquivo Emails como aplicação de janela nativa (sem browser).

Uso em desenvolvimento:  python launcher.py
Empacotamento:           duplo clique em criar_exe.bat
                         (resultado: dist\\ArquivoEmails.exe)
"""
import socket
import threading
import time

import webview

from app import app, PORT

# Necessário para eventuais downloads dentro da janela — pywebview >= 5.
webview.settings["ALLOW_DOWNLOADS"] = True


def _porta_ocupada():
    try:
        with socket.create_connection(("127.0.0.1", PORT), timeout=1):
            return True
    except OSError:
        return False


def _erro(msg):
    """Mostra o erro numa caixa de diálogo do Windows (ou na consola)."""
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(None, msg,
                                         "Arquivo Emails — Prospectiva", 0x10)
    except Exception:
        print(msg)
        input("Prima Enter para sair...")


def _run_server():
    app.run(host="127.0.0.1", port=PORT, debug=False, use_reloader=False)


def _wait_for_server(timeout=15):
    """Espera que o servidor Flask aceite ligações antes de abrir a janela."""
    limit = time.time() + timeout
    while time.time() < limit:
        try:
            with socket.create_connection(("127.0.0.1", PORT), timeout=1):
                return True
        except OSError:
            time.sleep(0.2)
    return False


class Api:
    """Funções expostas ao JavaScript da página (window.pywebview.api)."""

    def choose_folder(self):
        """Diálogo nativo de escolha de pasta; devolve o caminho ou None."""
        result = webview.windows[0].create_file_dialog(webview.FOLDER_DIALOG)
        if result:
            return result[0] if isinstance(result, (list, tuple)) else result
        return None

    def choose_msg_files(self):
        """Diálogo nativo de escolha de ficheiros .msg (selecção múltipla)."""
        result = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG, allow_multiple=True,
            file_types=("Emails do Outlook (*.msg)",))
        if not result:
            return []
        return list(result) if isinstance(result, (list, tuple)) else [result]


if __name__ == "__main__":
    if _porta_ocupada():
        _erro("Já existe uma instância do Arquivo Emails (ou outra aplicação) "
              f"a usar a porta {PORT}.\n\nFeche a execução anterior — janela "
              "preta de consola, possivelmente minimizada, ou processos "
              "python.exe/ArquivoEmails.exe no Gestor de Tarefas — e volte a "
              "abrir.")
        raise SystemExit(1)
    threading.Thread(target=_run_server, daemon=True).start()
    if not _wait_for_server():
        _erro("O servidor interno não arrancou. Volte a tentar; se persistir, "
              "contacte o suporte interno.")
        raise SystemExit(1)
    webview.create_window(
        "Arquivo Emails — Prospectiva",
        f"http://127.0.0.1:{PORT}",
        width=1180,
        height=820,
        min_size=(960, 640),
        js_api=Api(),
    )
    webview.start()
