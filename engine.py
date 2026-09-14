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
from urllib.parse import urlparse
from werkzeug.utils import secure_filename


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
CONTEXTOS_NOME = "_contextos.json"
ENTIDADES_NOME = "_entidades.json"
CAMPOS_INDICE = ["DataEmail", "Remetente", "Destinatarios", "Assunto",
                 "Anexos", "Ficheiro", "PastaAnexos", "Contexto", "Links", "Entidade", "Projecto",
                 "ArquivadoPor", "DataArquivo", "MessageID"]
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
    if pasta_raiz:
        nome = os.path.basename(os.path.normpath(pasta_raiz))
        conteudo = os.listdir(pasta_raiz)
        e_projeto = (os.path.isfile(os.path.join(pasta_raiz, INDICE_NOME))
                     or os.path.isdir(os.path.join(pasta_raiz, "Anexos"))
                     or any(str(f).lower().endswith(".msg") for f in conteudo))
        if e_projeto and re.match(r"^[\w\-. ]{2,80}$", nome, re.UNICODE):
            pasta_raiz = os.path.dirname(os.path.normpath(pasta_raiz))
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


CONTEXTOS_PADRAO = ["Informação", "Ação necessária", "Acompanhamento",
                    "Aprovação", "Decisão", "Reunião", "Contrato",
                    "Faturação", "Técnico", "Urgente", "Outro"]


def ler_contextos():
    path = os.path.join(_pasta_raiz(), CONTEXTOS_NOME)
    try:
        with open(path, "r", encoding="utf-8") as f:
            valores = json.load(f)
        if isinstance(valores, list):
            return [str(v).strip() for v in valores if str(v).strip()]
    except Exception:
        pass
    return list(CONTEXTOS_PADRAO)


def gravar_contextos(contextos):
    valores = []
    for contexto in contextos or []:
        contexto = str(contexto).strip()
        if contexto and contexto not in valores:
            valores.append(contexto)
    if not valores:
        raise EngineError("Mantenha pelo menos um contexto.")
    with open(os.path.join(_pasta_raiz(), CONTEXTOS_NOME), "w", encoding="utf-8") as f:
        json.dump(valores, f, ensure_ascii=False, indent=2)
    return valores


def ler_entidades():
    try:
        with open(os.path.join(_pasta_raiz(), ENTIDADES_NOME), "r", encoding="utf-8") as f:
            valores = json.load(f)
        return valores if isinstance(valores, dict) else {}
    except Exception:
        return {}


def gravar_entidade(remetente, entidade):
    remetente = (remetente or "").strip()
    entidade = (entidade or "").strip()
    if not remetente or not entidade:
        return
    valores = ler_entidades()
    valores[remetente.casefold()] = entidade
    with open(os.path.join(_pasta_raiz(), ENTIDADES_NOME), "w", encoding="utf-8") as f:
        json.dump(valores, f, ensure_ascii=False, indent=2)


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


def _pasta_email_projecto(projecto, criar=True):
    pasta = os.path.join(_pasta_projecto(projecto), "02-Geral", "06-Email")
    if criar:
        os.makedirs(pasta, exist_ok=True)
    return pasta


def _pastas_consulta_projecto(projecto):
    projeto = _pasta_projecto(projecto)
    email = os.path.join(projeto, "02-Geral", "06-Email")
    pastas = [email] if os.path.isdir(email) else []
    legado = projeto
    if os.path.isfile(os.path.join(legado, INDICE_NOME)):
        pastas.append(legado)
    return pastas


def _metadados_historico_projecto(projecto):
    dados = {}
    for pasta in _pastas_consulta_projecto(projecto):
        dados.update(_metadados_arquivados(pasta))
    return dados


def _msgids_historico_projecto(projecto):
    return set(_metadados_historico_projecto(projecto))


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


def _guardar_anexos(item, pasta_projecto, nome_email):
    """Guarda os anexos numa pasta própria do email e devolve o seu caminho relativo."""
    anexos = getattr(item, "Attachments", None)
    total = int(getattr(anexos, "Count", 0) or 0)
    if not total:
        return ""
    pasta = os.path.join(pasta_projecto, "Anexos", os.path.splitext(nome_email)[0])
    os.makedirs(pasta, exist_ok=True)
    for indice in range(1, total + 1):
        anexo = anexos.Item(indice)
        nome = os.path.basename(str(getattr(anexo, "FileName", "anexo") or "anexo"))
        _ = _caminho_livre(pasta, nome)
        anexo.SaveAsFile(_)
    return os.path.relpath(pasta, pasta_projecto)


def _guardar_anexos_selecionados(item, pasta_projecto, nome_pasta, indices):
    anexos = getattr(item, "Attachments", None)
    indices = {int(i) for i in (indices or [])}
    if not indices:
        return ""
    pasta = os.path.join(pasta_projecto, "Anexos", nome_pasta)
    os.makedirs(pasta, exist_ok=True)
    for indice in sorted(indices):
        if indice < 1 or indice > int(getattr(anexos, "Count", 0) or 0):
            continue
        anexo = anexos.Item(indice)
        nome = os.path.basename(str(getattr(anexo, "FileName", "anexo") or "anexo"))
        anexo.SaveAsFile(_caminho_livre(pasta, nome))
    return os.path.relpath(pasta, pasta_projecto) if os.listdir(pasta) else ""


def _pasta_anexos_sugerida(data, remetente, assunto):
    base = nome_normalizado(data, remetente, assunto).removesuffix(".msg")
    return base[:100]


def _guardar_ficheiros(ficheiros, pasta_projecto, nome_pasta):
    if not ficheiros:
        return ""
    pasta = os.path.join(pasta_projecto, "Anexos", nome_pasta)
    os.makedirs(pasta, exist_ok=True)
    for ficheiro in ficheiros:
        nome = secure_filename(os.path.basename(ficheiro.filename or ""))
        if nome:
            ficheiro.save(_caminho_livre(pasta, nome))
    return os.path.relpath(pasta, pasta_projecto)


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


def _contextos_arquivados(pasta_projecto):
    return {r.get("MessageID"): r.get("Contexto", "")
            for r in _ler_indice(pasta_projecto) if r.get("MessageID")}


def _metadados_arquivados(pasta_projecto):
    return {r.get("MessageID"): r for r in _ler_indice(pasta_projecto)
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
            if not novo:
                with open(path, "r", encoding="utf-8-sig", newline="") as f:
                    leitor = csv.reader(f, delimiter=";")
                    cabecalho = next(leitor, [])
                if cabecalho != CAMPOS_INDICE:
                    linhas = _ler_indice(pasta_projecto)
                    with open(path, "w", encoding="utf-8-sig", newline="") as f:
                        w = csv.DictWriter(f, fieldnames=CAMPOS_INDICE,
                                           delimiter=";")
                        w.writeheader()
                        w.writerows({k: r.get(k, "") for k in CAMPOS_INDICE}
                                    for r in linhas)
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


def listar_emails(entry_id, store_id, limite=50, filtro="", projecto="",
                  max_varridos=1500):
    """Emails da pasta, do mais recente para o mais antigo. O filtro é
    aplicado ao remetente e ao assunto (sem distinguir maiúsculas)."""
    ns = _outlook()
    arquivados_no_projecto = set()
    metadados_no_projecto = {}
    if projecto:
        pasta_projecto = _pasta_email_projecto(projecto)
        arquivados_no_projecto = _msgids_historico_projecto(projecto)
        metadados_no_projecto = _metadados_historico_projecto(projecto)
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
            msgid = _message_id(item)
            out.append({
                "entry_id": item.EntryID,
                "store_id": store_id,
                "data": _fmt_data(getattr(item, "ReceivedTime", None)),
                "remetente": remetente,
                "entidade": _entidade_item(item),
                "assunto": assunto,
                "anexos": int(getattr(item.Attachments, "Count", 0) or 0),
                "arquivado": bool(msgid and msgid in arquivados_no_projecto),
                "contexto": metadados_no_projecto.get(msgid, {}).get("Contexto", ""),
                "links": metadados_no_projecto.get(msgid, {}).get("Links", ""),
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


def _entidade_item(item):
    try:
        exchange_user = item.Sender.GetExchangeUser()
        empresa = str(getattr(exchange_user, "CompanyName", "") or "").strip()
        if empresa:
            return empresa
    except Exception:
        pass
    endereco = str(getattr(item, "SenderEmailAddress", "") or "").strip()
    if "@" in endereco:
        return endereco.rsplit("@", 1)[1].split(".", 1)[0].lower()
    return ""


def _dados_item(item, store_id, metadados=None):
    msgid = _message_id(item)
    metadados = metadados or {}
    data = getattr(item, "ReceivedTime", None)
    remetente = str(getattr(item, "SenderName", "") or "")
    entidade = ler_entidades().get(remetente.casefold(), _entidade_item(item))
    anexos = [{"indice": i, "nome": str(getattr(item.Attachments.Item(i), "FileName", "anexo") or "anexo")}
              for i in range(1, int(getattr(item.Attachments, "Count", 0) or 0) + 1)]
    return {
        "entry_id": item.EntryID,
        "store_id": store_id,
        "data": _fmt_data(data),
        "remetente": remetente,
        "entidade": entidade,
        "assunto": str(getattr(item, "Subject", "") or ""),
        "anexos": anexos,
        "pasta_anexos_sugerida": _pasta_anexos_sugerida(data, remetente, str(getattr(item, "Subject", "") or "")),
        "arquivado": bool(msgid and msgid in metadados),
        "contexto": metadados.get(msgid, {}).get("Contexto", ""),
        "links": metadados.get(msgid, {}).get("Links", ""),
    }


def selecionar_emails(projecto=""):
    """Obtém apenas os emails atualmente selecionados no Outlook clássico."""
    ns = _outlook()
    try:
        selection = ns.Application.ActiveExplorer().Selection
    except Exception as e:
        raise EngineError("Não foi possível ler a seleção atual do Outlook. "
                          "Abra uma pasta e seleccione os emails.") from e
    metadados = {}
    if projecto:
        metadados = _metadados_historico_projecto(projecto)
    out = []
    for indice in range(1, selection.Count + 1):
        try:
            item = selection.Item(indice)
            if getattr(item, "Class", None) == OL_MAIL_ITEM:
                out.append(_dados_item(item, getattr(item, "StoreID", "") or "",
                                       metadados))
        except Exception:
            continue
    if not out:
        raise EngineError("Não foram encontrados emails selecionados no Outlook.")
    return {"emails": out, "varridos": len(out), "truncado": False}


def _arquivar_item(item, pasta_projecto, projecto, ja_arquivados,
                   permitir_repetidos, marcar_categoria, contexto="", links="",
                   entidade="", guardar_anexos=True, anexos_selecionados=None,
                   nome_pasta_anexos="", ficheiros=None):
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
    links = (links or "").strip()
    if links and urlparse(links).scheme not in ("http", "https"):
        raise EngineError("O link do anexo deve começar por http:// ou https://.")
    caminho = _caminho_livre(pasta_projecto, nome)
    item.SaveAs(caminho, OL_FORMATO_MSG)
    pasta_anexos = ""
    prefixo_data = data.strftime("%Y-%m-%d") + "_"
    nome_pasta = nome_pasta_anexos.strip() or os.path.splitext(os.path.basename(caminho))[0]
    nome_pasta = re.sub(r"[\\/:*?\"<>|]", "_", nome_pasta).strip(" .")[:120]
    nome_pasta = prefixo_data + nome_pasta.removeprefix(prefixo_data)
    nome_pasta = re.sub(r"[\\/:*?\"<>|]", "_", nome_pasta).strip(" .")[:120]
    if guardar_anexos:
        pasta_anexos = _guardar_anexos_selecionados(
            item, pasta_projecto, nome_pasta, anexos_selecionados)
    pasta_anexos = _guardar_ficheiros(ficheiros or [], pasta_projecto,
                                      nome_pasta) or pasta_anexos
    _registar_indice(pasta_projecto, {
        "DataEmail": _fmt_data(data),
        "Remetente": remetente,
        "Destinatarios": dest,
        "Assunto": assunto,
        "Anexos": int(getattr(item.Attachments, "Count", 0) or 0),
        "Ficheiro": os.path.basename(caminho),
        "PastaAnexos": pasta_anexos,
        "Contexto": contexto.strip(),
        "Links": links.strip(),
        "Entidade": (entidade or _entidade_item(item)).strip(),
        "Projecto": projecto,
        "ArquivadoPor": os.environ.get("USERNAME", ""),
        "DataArquivo": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "MessageID": msgid,
    })
    gravar_entidade(remetente, entidade or _entidade_item(item))
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
    pasta_projecto = _pasta_email_projecto(projecto)
    ns = _outlook()
    ja = _msgids_historico_projecto(projecto)
    arquivados, repetidos, erros = [], [], []
    for ref in ids:
        try:
            if ref.get("store_id"):
                item = ns.GetItemFromID(ref["entry_id"], ref["store_id"])
            else:
                item = ns.GetItemFromID(ref["entry_id"])
            estado, info = _arquivar_item(item, pasta_projecto, projecto, ja,
                                          permitir_repetidos,
                                          marcar_categoria,
                                          ref.get("contexto", ""),
                                          ref.get("links", ""),
                                          ref.get("entidade", ""),
                                          ref.get("guardar_anexos", True),
                                          ref.get("anexos_selecionados", []),
                                          ref.get("nome_pasta_anexos", ""),
                                          ref.get("ficheiros", []))
            (arquivados if estado == "ok" else repetidos).append(info)
        except EngineError:
            raise
        except Exception as e:
            erros.append(str(e))
    return {"arquivados": arquivados, "repetidos": repetidos, "erros": erros,
            "pasta": pasta_projecto}


def arquivar_ficheiros(caminhos, projecto, permitir_repetidos=False,
                       contexto=""):
    """Alternativa: arquiva ficheiros .msg já existentes no disco
    (ex.: emails arrastados do Outlook para uma pasta). Usa o Outlook
    para ler os metadados e grava uma cópia normalizada no projecto."""
    if not caminhos:
        raise EngineError("Escolha pelo menos um ficheiro .msg.")
    pasta_projecto = _pasta_email_projecto(projecto)
    ns = _outlook()
    ja = _msgids_historico_projecto(projecto)
    arquivados, repetidos, erros = [], [], []
    for c in caminhos:
        try:
            if not str(c).lower().endswith(".msg"):
                erros.append(f"Ignorado (não é .msg): {os.path.basename(c)}")
                continue
            item = ns.OpenSharedItem(c)
            estado, info = _arquivar_item(item, pasta_projecto, projecto, ja,
                                          permitir_repetidos, False, contexto)
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
    raiz = _pasta_raiz()
    projectos = [projecto] if projecto else listar_projectos()
    out = []
    for prj in projectos:
        for pasta in _pastas_consulta_projecto(prj):
          for r in _ler_indice(pasta):
            alvo = " ".join([r.get("DataEmail", ""), r.get("Remetente", ""),
                             r.get("Destinatarios", ""), r.get("Assunto", ""),
                             r.get("Ficheiro", ""), r.get("Contexto", ""),
                             r.get("Entidade", "")]).casefold()
            if termo in alvo:
                out.append({
                    "projecto": prj,
                    "data": r.get("DataEmail", ""),
                    "remetente": r.get("Remetente", ""),
                    "entidade": r.get("Entidade", ""),
                    "assunto": r.get("Assunto", ""),
                    "contexto": r.get("Contexto", ""),
                    "pasta_anexos": os.path.join(pasta, r.get("PastaAnexos", ""))
                        if r.get("PastaAnexos") else "",
                    "ficheiro": r.get("Ficheiro", ""),
                    "arquivado_por": r.get("ArquivadoPor", ""),
                    "links": r.get("Links", ""),
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
