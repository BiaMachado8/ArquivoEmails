# -*- coding: utf-8 -*-
"""
Arquivo Emails — motor (mesmo padrão do Conversor PDF/A e do Renumeração).

Liga ao Outlook clássico via COM (pywin32), grava emails seleccionados como
ficheiros .msg com nome normalizado (AAAA-MM-DD_Remetente_Assunto.msg) na
pasta do projecto escolhido, e mantém um índice pesquisável por projecto
(_indice_emails.csv — abre directamente no Excel).

A pasta raiz do arquivo é configurável (local para testes; de rede para
uso partilhado pela equipa). Cada projecto é uma subpasta da raiz.
"""
import csv
import json
import os
import re
import sys
import time
import unicodedata
from datetime import datetime


class EngineError(Exception):
    """Erro de utilização/ambiente, com mensagem apresentável ao utilizador."""


# ---------------------------------------------------------------- caminhos --
# APP_DIR: recursos empacotados (PyInstaller extrai para _MEIPASS)
APP_DIR = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
# BASE_DIR: pasta do .exe (ou do código-fonte, em desenvolvimento)
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".arquivo_emails.json")

INDICE_NOME = "_indice_emails.csv"
CAMPOS_INDICE = ["DataEmail", "Remetente", "Destinatarios", "Assunto",
                 "Anexos", "Ficheiro", "Projecto", "ArquivadoPor",
                 "DataArquivo", "MessageID"]
CATEGORIA_ARQUIVADO = "Arquivado DEP"

OL_MAIL_ITEM = 43        # olMail
OL_FORMATO_MSG = 3       # olMSG (formato Unicode .msg)
PROP_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"


# ------------------------------------------------------------- diagnóstico --
def check_tools():
    """Versões das dependências (para o aviso na página inicial)."""
    tools = {}
    try:
        import flask
        tools["flask"] = flask.__version__
    except Exception:
        tools["flask"] = None
    try:
        import win32com  # noqa: F401
        tools["pywin32"] = "ok"
    except Exception:
        tools["pywin32"] = None
    return tools


# --------------------------------------------------------------- configuração
def ler_config():
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    except Exception:
        cfg = {}
    cfg.setdefault("pasta_raiz", "")
    return cfg


def gravar_config(pasta_raiz):
    pasta_raiz = (pasta_raiz or "").strip().strip('"')
    if pasta_raiz and not os.path.isdir(pasta_raiz):
        raise EngineError(f"A pasta indicada não existe: {pasta_raiz}")
    cfg = ler_config()
    cfg["pasta_raiz"] = pasta_raiz
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    return cfg


def _pasta_raiz():
    raiz = ler_config().get("pasta_raiz") or ""
    if not raiz:
        raise EngineError("Configure primeiro a pasta raiz do arquivo "
                          "(secção 1).")
    if not os.path.isdir(raiz):
        raise EngineError(f"A pasta raiz do arquivo não está acessível: "
                          f"{raiz}. Verifique a ligação à rede.")
    return raiz


# ----------------------------------------------------------------- projectos
_RE_PROJECTO = re.compile(r"^[\w\-. ]{2,80}$", re.UNICODE)


def listar_projectos():
    raiz = _pasta_raiz()
    out = []
    for nome in sorted(os.listdir(raiz), key=str.casefold):
        p = os.path.join(raiz, nome)
        if os.path.isdir(p) and not nome.startswith(("_", ".")):
            out.append(nome)
    return out


def criar_projecto(nome):
    nome = (nome or "").strip()
    if not _RE_PROJECTO.match(nome) or nome.startswith(("_", ".")):
        raise EngineError("Nome de projecto inválido. Use letras, números, "
                          "hífen, ponto e espaço (ex.: P3191_ANA).")
    raiz = _pasta_raiz()
    p = os.path.join(raiz, nome)
    if os.path.isdir(p):
        raise EngineError(f"O projecto «{nome}» já existe.")
    os.makedirs(p)
    return nome


def _pasta_projecto(projecto):
    raiz = _pasta_raiz()
    p = os.path.abspath(os.path.join(raiz, projecto or ""))
    # impedir fugas fora da raiz (ex.: "..\..")
    if os.path.commonpath([p, os.path.abspath(raiz)]) != os.path.abspath(raiz) \
            or p == os.path.abspath(raiz):
        raise EngineError("Escolha um projecto válido.")
    if not os.path.isdir(p):
        raise EngineError(f"O projecto «{projecto}» não existe na pasta raiz.")
    return p


# ------------------------------------------------------- nomes normalizados --
def _sem_acentos(s):
    return unicodedata.normalize("NFKD", s or "").encode(
        "ascii", "ignore").decode("ascii")


def _limpar_assunto(assunto):
    """Remove prefixos RE:/FW:/ENC: em cadeia, mantendo o assunto útil."""
    s = (assunto or "").strip()
    while True:
        novo = re.sub(r"^\s*(re|fw|fwd|enc|rv|res)\s*:\s*", "", s,
                      flags=re.IGNORECASE)
        if novo == s:
            return s
        s = novo


def nome_normalizado(data, remetente, assunto, max_assunto=60):
    """AAAA-MM-DD_Remetente_Assunto.msg (sem acentos nem caracteres
    inválidos no Windows; ordena cronologicamente no Explorer)."""
    rem = re.sub(r"[^A-Za-z0-9]+", "", _sem_acentos(remetente)) or "Remetente"
    rem = rem[:25]
    ass = re.sub(r"[^A-Za-z0-9]+", "_",
                 _sem_acentos(_limpar_assunto(assunto))).strip("_")
    ass = (ass or "Sem_assunto")[:max_assunto].rstrip("_")
    return f"{data:%Y-%m-%d}_{rem}_{ass}.msg"


def _caminho_livre(pasta, nome):
    """Evita colisões: acrescenta _2, _3, … antes da extensão."""
    base, ext = os.path.splitext(nome)
    caminho = os.path.join(pasta, nome)
    n = 2
    while os.path.exists(caminho):
        caminho = os.path.join(pasta, f"{base}_{n}{ext}")
        n += 1
    return caminho


# --------------------------------------------------------------------- índice
def _indice_path(pasta_projecto):
    return os.path.join(pasta_projecto, INDICE_NOME)


def _ler_indice(pasta_projecto):
    path = _indice_path(pasta_projecto)
    if not os.path.isfile(path):
        return []
    try:
        with open(path, "r", encoding="utf-8-sig", newline="") as f:
            return list(csv.DictReader(f, delimiter=";"))
    except Exception:
        return []


def _msgids_arquivados(pasta_projecto):
    return {r.get("MessageID") for r in _ler_indice(pasta_projecto)
            if r.get("MessageID")}


def _registar_indice(pasta_projecto, linha, tentativas=12, espera=0.5):
    """Acrescenta uma linha ao índice do projecto. Em pasta de rede
    partilhada o ficheiro pode estar momentaneamente bloqueado por um
    colega — repete com espera antes de desistir."""
    path = _indice_path(pasta_projecto)
    novo = not os.path.isfile(path)
    ultima = None
    for _ in range(tentativas):
        try:
            with open(path, "a", encoding="utf-8-sig", newline="") as f:
                w = csv.DictWriter(f, fieldnames=CAMPOS_INDICE, delimiter=";")
                if novo:
                    w.writeheader()
                w.writerow({k: linha.get(k, "") for k in CAMPOS_INDICE})
            return
        except (PermissionError, OSError) as e:
            ultima = e
            time.sleep(espera)
    raise EngineError(f"Não foi possível actualizar o índice "
                      f"({INDICE_NOME}) — ficheiro bloqueado? Feche-o no "
                      f"Excel e repita. Detalhe: {ultima}")


# --------------------------------------------------------------- Outlook COM
def _outlook():
    """Liga ao Outlook clássico. Cada pedido do Flask corre numa thread
    própria, pelo que o COM tem de ser inicializado em cada chamada."""
    try:
        import pythoncom
        pythoncom.CoInitialize()
        import win32com.client
    except ImportError:
        raise EngineError("Dependência pywin32 em falta — correr: "
                          "python -m pip install -r requirements.txt")
    try:
        app = win32com.client.Dispatch("Outlook.Application")
        return app.GetNamespace("MAPI")
    except Exception as e:
        raise EngineError(
            "Não foi possível ligar ao Outlook. Esta ligação requer o "
            "Outlook CLÁSSICO instalado e configurado (o «Novo Outlook» "
            "não a suporta — desactive o interruptor «Experimentar o novo "
            "Outlook»). Abra o Outlook e volte a tentar. "
            f"Detalhe técnico: {e}")


def listar_pastas_outlook(max_nivel=4):
    """Árvore de pastas de correio, achatada com indentação por nível."""
    ns = _outlook()
    out = []

    def _walk(folder, caminho, nivel):
        if nivel > max_nivel:
            return
        try:
            if getattr(folder, "DefaultItemType", 0) != 0:  # 0 = correio
                return
        except Exception:
            return
        try:
            total = folder.Items.Count
        except Exception:
            total = None
        out.append({"entry_id": folder.EntryID,
                    "store_id": folder.StoreID,
                    "nome": folder.Name,
                    "caminho": caminho,
                    "nivel": nivel,
                    "total": total})
        try:
            subs = list(folder.Folders)
        except Exception:
            subs = []
        for sub in sorted(subs, key=lambda x: str(x.Name).casefold()):
            _walk(sub, f"{caminho} / {sub.Name}", nivel + 1)

    for raiz in ns.Folders:
        try:
            subs = list(raiz.Folders)
        except Exception:
            continue
        for sub in subs:
            _walk(sub, f"{raiz.Name} / {sub.Name}", 1)
    if not out:
        raise EngineError("Nenhuma pasta de correio encontrada no Outlook.")
    return out


def _fmt_data(dt):
    try:
        return dt.strftime("%Y-%m-%d %H:%M")
    except Exception:
        return str(dt or "")


def listar_emails(entry_id, store_id, limite=50, filtro="",
                  max_varridos=1500):
    """Emails da pasta, do mais recente para o mais antigo. O filtro é
    aplicado ao remetente e ao assunto (sem distinguir maiúsculas)."""
    ns = _outlook()
    try:
        pasta = ns.GetFolderFromID(entry_id, store_id)
    except Exception:
        raise EngineError("Pasta do Outlook não encontrada — volte a "
                          "carregar em «Ler pastas do Outlook».")
    itens = pasta.Items
    itens.Sort("[ReceivedTime]", True)
    filtro = (filtro or "").strip().casefold()
    out, varridos = [], 0
    for item in itens:
        varridos += 1
        if varridos > max_varridos:
            break
        try:
            if getattr(item, "Class", None) != OL_MAIL_ITEM:
                continue
            remetente = str(getattr(item, "SenderName", "") or "")
            assunto = str(getattr(item, "Subject", "") or "")
            if filtro and filtro not in remetente.casefold() \
                    and filtro not in assunto.casefold():
                continue
            cats = str(getattr(item, "Categories", "") or "")
            out.append({
                "entry_id": item.EntryID,
                "store_id": store_id,
                "data": _fmt_data(getattr(item, "ReceivedTime", None)),
                "remetente": remetente,
                "assunto": assunto,
                "anexos": int(getattr(item.Attachments, "Count", 0) or 0),
                "arquivado": CATEGORIA_ARQUIVADO in cats,
            })
        except Exception:
            continue
        if len(out) >= limite:
            break
    return {"emails": out, "varridos": min(varridos, max_varridos),
            "truncado": varridos > max_varridos}


def _message_id(item):
    try:
        pa = item.PropertyAccessor
        return str(pa.GetProperty(PROP_MESSAGE_ID) or "").strip()
    except Exception:
        return ""


def _arquivar_item(item, pasta_projecto, projecto, ja_arquivados,
                   permitir_repetidos, marcar_categoria):
    """Arquiva um MailItem; devolve ('ok'|'repetido', info)."""
    msgid = _message_id(item)
    if msgid and not permitir_repetidos and msgid in ja_arquivados:
        return "repetido", str(getattr(item, "Subject", "") or "")
    data = getattr(item, "ReceivedTime", None) or datetime.now()
    remetente = str(getattr(item, "SenderName", "") or "")
    assunto = str(getattr(item, "Subject", "") or "")
    try:
        dest = str(getattr(item, "To", "") or "")
    except Exception:
        dest = ""
    nome = nome_normalizado(data, remetente, assunto)
    caminho = _caminho_livre(pasta_projecto, nome)
    item.SaveAs(caminho, OL_FORMATO_MSG)
    _registar_indice(pasta_projecto, {
        "DataEmail": _fmt_data(data),
        "Remetente": remetente,
        "Destinatarios": dest,
        "Assunto": assunto,
        "Anexos": int(getattr(item.Attachments, "Count", 0) or 0),
        "Ficheiro": os.path.basename(caminho),
        "Projecto": projecto,
        "ArquivadoPor": os.environ.get("USERNAME", ""),
        "DataArquivo": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "MessageID": msgid,
    })
    if msgid:
        ja_arquivados.add(msgid)
    if marcar_categoria:
        try:
            cats = [c.strip() for c in
                    str(getattr(item, "Categories", "") or "").split(";")
                    if c.strip()]
            if CATEGORIA_ARQUIVADO not in cats:
                cats.append(CATEGORIA_ARQUIVADO)
                item.Categories = "; ".join(cats)
                item.Save()
        except Exception:
            pass  # marcação é conveniência; não impede o arquivo
    return "ok", os.path.basename(caminho)


def arquivar_emails(ids, projecto, permitir_repetidos=False,
                    marcar_categoria=True):
    """ids: lista de {entry_id, store_id} vindos de listar_emails."""
    if not ids:
        raise EngineError("Seleccione pelo menos um email.")
    pasta_projecto = _pasta_projecto(projecto)
    ns = _outlook()
    ja = _msgids_arquivados(pasta_projecto)
    arquivados, repetidos, erros = [], [], []
    for ref in ids:
        try:
            item = ns.GetItemFromID(ref["entry_id"], ref["store_id"])
            estado, info = _arquivar_item(item, pasta_projecto, projecto, ja,
                                          permitir_repetidos,
                                          marcar_categoria)
            (arquivados if estado == "ok" else repetidos).append(info)
        except EngineError:
            raise
        except Exception as e:
            erros.append(str(e))
    return {"arquivados": arquivados, "repetidos": repetidos, "erros": erros,
            "pasta": pasta_projecto}


def arquivar_ficheiros(caminhos, projecto, permitir_repetidos=False):
    """Alternativa: arquiva ficheiros .msg já existentes no disco
    (ex.: emails arrastados do Outlook para uma pasta). Usa o Outlook
    para ler os metadados e grava uma cópia normalizada no projecto."""
    if not caminhos:
        raise EngineError("Escolha pelo menos um ficheiro .msg.")
    pasta_projecto = _pasta_projecto(projecto)
    ns = _outlook()
    ja = _msgids_arquivados(pasta_projecto)
    arquivados, repetidos, erros = [], [], []
    for c in caminhos:
        try:
            if not str(c).lower().endswith(".msg"):
                erros.append(f"Ignorado (não é .msg): {os.path.basename(c)}")
                continue
            item = ns.OpenSharedItem(c)
            estado, info = _arquivar_item(item, pasta_projecto, projecto, ja,
                                          permitir_repetidos, False)
            (arquivados if estado == "ok" else repetidos).append(info)
        except EngineError:
            raise
        except Exception as e:
            erros.append(f"{os.path.basename(str(c))}: {e}")
    return {"arquivados": arquivados, "repetidos": repetidos, "erros": erros,
            "pasta": pasta_projecto}


# ------------------------------------------------------------------ pesquisa
def pesquisar(termo, projecto=None, max_resultados=200):
    """Pesquisa nos índices (todos os projectos, ou apenas um)."""
    termo = (termo or "").strip().casefold()
    if not termo:
        raise EngineError("Indique o texto a pesquisar.")
    raiz = _pasta_raiz()
    projectos = [projecto] if projecto else listar_projectos()
    out = []
    for prj in projectos:
        pasta = os.path.join(raiz, prj)
        if not os.path.isdir(pasta):
            continue
        for r in _ler_indice(pasta):
            alvo = " ".join([r.get("DataEmail", ""), r.get("Remetente", ""),
                             r.get("Destinatarios", ""), r.get("Assunto", ""),
                             r.get("Ficheiro", "")]).casefold()
            if termo in alvo:
                out.append({
                    "projecto": prj,
                    "data": r.get("DataEmail", ""),
                    "remetente": r.get("Remetente", ""),
                    "assunto": r.get("Assunto", ""),
                    "ficheiro": r.get("Ficheiro", ""),
                    "arquivado_por": r.get("ArquivadoPor", ""),
                    "caminho": os.path.join(pasta, r.get("Ficheiro", "")),
                    "existe": os.path.isfile(
                        os.path.join(pasta, r.get("Ficheiro", ""))),
                })
                if len(out) >= max_resultados:
                    return {"resultados": out, "truncado": True}
    out.sort(key=lambda r: r["data"], reverse=True)
    return {"resultados": out, "truncado": False}


def abrir_caminho(caminho):
    """Abre um email arquivado (ou a pasta de um projecto) no Windows.
    Só permite caminhos dentro da pasta raiz do arquivo."""
    raiz = os.path.abspath(_pasta_raiz())
    alvo = os.path.abspath((caminho or "").strip())
    if os.path.commonpath([alvo, raiz]) != raiz:
        raise EngineError("Só é possível abrir ficheiros dentro do arquivo.")
    if not os.path.exists(alvo):
        raise EngineError("O ficheiro já não existe nesse local.")
    os.startfile(alvo)  # noqa: disponível apenas no Windows
    return True
