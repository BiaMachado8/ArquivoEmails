# -*- coding: utf-8 -*-
"""
Arquivo Emails — aplicação web local (mesmo padrão do Conversor PDF/A e
do Renumeração). Arquiva emails do Outlook em pastas de projecto, com
nome normalizado e índice pesquisável partilhável em rede.

Arranque:  python app.py   →  http://localhost:8323
"""
import json
import os
import socket
import webbrowser
from threading import Timer

from flask import Flask, jsonify, render_template, request, send_file

import engine

app = Flask(__name__)

PORT = 8323  # Conversor PDF/A: 8321 · Renumeração: 8322 — podem coexistir


@app.route("/")
def index():
    return render_template("index.html", tools=engine.check_tools(),
                           config=engine.ler_config())


@app.route("/logo")
def logo():
    """Serve o logotipo do programa, se existir em resources/logo.png."""
    for base in (engine.APP_DIR, engine.BASE_DIR):
        p = os.path.join(base, "resources", "logo.png")
        if os.path.isfile(p):
            return send_file(p, mimetype="image/png")
    return ("", 404)


@app.route("/api/tools")
def api_tools():
    return jsonify(engine.check_tools())


# ------------------------------------------------------------- configuração
@app.route("/api/config", methods=["GET", "POST"])
def api_config():
    try:
        if request.method == "POST":
            cfg = engine.gravar_config(request.form.get("pasta_raiz"))
        else:
            cfg = engine.ler_config()
        return jsonify(cfg)
    except engine.EngineError as e:
        return jsonify({"error": str(e)}), 400


@app.route("/api/projectos", methods=["GET", "POST"])
def api_projectos():
    try:
        if request.method == "POST":
            engine.criar_projecto(request.form.get("nome"))
        return jsonify({"projectos": engine.listar_projectos()})
    except engine.EngineError as e:
        return jsonify({"error": str(e)}), 400


# ------------------------------------------------------------------ Outlook
@app.route("/api/outlook/pastas", methods=["POST"])
def api_pastas():
    try:
        return jsonify({"pastas": engine.listar_pastas_outlook()})
    except engine.EngineError as e:
        return jsonify({"error": str(e)}), 400


@app.route("/api/outlook/emails", methods=["POST"])
def api_emails():
    try:
        try:
            limite = max(1, min(500, int(request.form.get("limite") or 50)))
        except ValueError:
            raise engine.EngineError("O n.º de emails a listar deve ser "
                                     "um número inteiro.")
        res = engine.listar_emails(
            (request.form.get("entry_id") or "").strip(),
            (request.form.get("store_id") or "").strip(),
            limite=limite,
            filtro=request.form.get("filtro") or "")
        return jsonify(res)
    except engine.EngineError as e:
        return jsonify({"error": str(e)}), 400


@app.route("/api/arquivar", methods=["POST"])
def api_arquivar():
    try:
        try:
            ids = json.loads(request.form.get("ids") or "[]")
        except ValueError:
            raise engine.EngineError("Pedido inválido — volte a listar os "
                                     "emails e repita.")
        res = engine.arquivar_emails(
            ids,
            request.form.get("projecto") or "",
            permitir_repetidos=(request.form.get("repetidos") == "1"),
            marcar_categoria=(request.form.get("marcar") != "0"))
        return jsonify({"ok": True, **res})
    except engine.EngineError as e:
        return jsonify({"error": str(e)}), 400


@app.route("/api/arquivar-ficheiros", methods=["POST"])
def api_arquivar_ficheiros():
    try:
        try:
            caminhos = json.loads(request.form.get("caminhos") or "[]")
        except ValueError:
            raise engine.EngineError("Pedido inválido.")
        res = engine.arquivar_ficheiros(
            caminhos,
            request.form.get("projecto") or "",
            permitir_repetidos=(request.form.get("repetidos") == "1"))
        return jsonify({"ok": True, **res})
    except engine.EngineError as e:
        return jsonify({"error": str(e)}), 400


# ----------------------------------------------------------------- pesquisa
@app.route("/api/pesquisar", methods=["POST"])
def api_pesquisar():
    try:
        res = engine.pesquisar(request.form.get("termo"),
                               request.form.get("projecto") or None)
        return jsonify(res)
    except engine.EngineError as e:
        return jsonify({"error": str(e)}), 400


@app.route("/api/abrir", methods=["POST"])
def api_abrir():
    try:
        engine.abrir_caminho(request.form.get("caminho"))
        return jsonify({"ok": True})
    except engine.EngineError as e:
        return jsonify({"error": str(e)}), 400


# ----------------------------------------------------------------- arranque
def _open_browser():
    webbrowser.open(f"http://localhost:{PORT}")


def _porta_ocupada():
    try:
        with socket.create_connection(("127.0.0.1", PORT), timeout=1):
            return True
    except OSError:
        return False


if __name__ == "__main__":
    if _porta_ocupada():
        print("=" * 60)
        print("  ERRO: ja existe uma instancia do Arquivo Emails (ou outra")
        print(f"  aplicacao) a usar a porta {PORT}.")
        print()
        print("  Feche a janela preta (consola) da execucao anterior —")
        print("  pode estar minimizada — ou termine os processos")
        print('  "python.exe" no Gestor de Tarefas, e volte a iniciar.')
        print("=" * 60)
        input("\nPrima Enter para sair...")
        raise SystemExit(1)
    tools = engine.check_tools()
    missing = [k for k, v in tools.items() if not v]
    print("=" * 60)
    print("  Arquivo Emails — Prospectiva")
    print("=" * 60)
    for k, v in tools.items():
        print(f"  {'OK ' if v else 'EM FALTA'}  {k}: {v or '---'}")
    if missing:
        print("\n  AVISO: dependências em falta — correr: "
              "python -m pip install -r requirements.txt")
    print(f"\n  A abrir http://localhost:{PORT} no browser...")
    Timer(1.5, _open_browser).start()
    app.run(host="127.0.0.1", port=PORT, debug=False)
