"""
Acompanhamento de Meta (Olympus/MGM e Sniper) — regras da CRN
Backend serverless (Vercel) — busca Pipedrive + Google Sheets e calcula
tudo que a planilha "Acompanhamento_Meta_Julho2026_V1.xlsx" fazia via fórmula.

Fontes:
- Pipedrive API v1 (deals ganhos) e v2 (forecast aberto, activities)
- Google Sheets publicados em CSV: COLAB, METAS, FERIADOS

Sem cache: cada chamada refaz tudo (decisão do time, ver /areas/acompanhamento-meta-dashboard.md).
"""

import os
import csv
import io
import json
import unicodedata
import calendar
import math
import functools
import socket
import re
import datetime as dt
from http.server import BaseHTTPRequestHandler
from urllib.request import Request, urlopen
from urllib.parse import urlencode

# ===========================================================================
# CRN — Centralização de Regras de Negócio (bloco de modelos/vercel_carregar_crn.py)
# ===========================================================================
# ── CRN: baixada do GitHub quando o painel liga ───────────────
# A Vercel não repassa o token para o instalador do build, então a CRN não entra pelo
# requirements.txt. O painel baixa a versão atual da main (com o token, que em tempo de
# execução funciona) e carrega. Cada vez que o painel liga, pega as regras mais recentes.
# No computador, se a crn já estiver instalada (instalar.bat), usa a instalada.
def _carregar_crn():
    import importlib.util
    import shutil
    import sys
    import tarfile
    import tempfile
    if importlib.util.find_spec("crn") is not None:
        return None
    import requests
    token = (os.environ.get("CRN_GITHUB_TOKEN") or os.environ.get("GITHUB_TOKEN") or "").strip()
    if not token:
        raise RuntimeError("CRN_GITHUB_TOKEN não está nas Environment Variables da Vercel.")
    r = requests.get("https://api.github.com/repos/negocios87-sketch/CRN/tarball/main",
                     headers={"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json"},
                     timeout=30)
    if r.status_code != 200:
        raise RuntimeError(f"Não consegui baixar a CRN do GitHub (HTTP {r.status_code}). Confira o CRN_GITHUB_TOKEN.")
    base = tempfile.mkdtemp(prefix="crn_")
    with tarfile.open(fileobj=io.BytesIO(r.content), mode="r:gz") as tar:
        tar.extractall(base, filter="data")
    raiz = os.path.join(base, os.listdir(base)[0])          # negocios87-sketch-CRN-<commit>
    destino = os.path.join(base, "pacote", "crn")
    shutil.copytree(os.path.join(raiz, "lib"), destino)     # mesmo mapa do pyproject da CRN
    shutil.copytree(os.path.join(raiz, "config"), os.path.join(destino, "config"))
    sys.path.insert(0, os.path.join(base, "pacote"))
    return os.path.basename(raiz).rsplit("-", 1)[-1][:7]   # commit que entrou

_ERRO_CRN = None   # se a CRN não carregar, o painel responde o motivo em JSON em vez de cair
try:
    _COMMIT_CRN = _carregar_crn()

    import crn
    from crn.base import fuso as _fuso, fontes as _fontes, pipedrive as _pd
    from crn.base.config import campo as _campo
    from crn.regras import (m1_ganho_closer as _m1, m2_faturamento_closer as _m2, m3_atingimento_closer as _m3,
                            m4_ganho_sdr as _m4, m6_reunioes_sdr as _m6, m7_pct_final_sdr as _m7,
                            t1_mtd_closer as _t1,
                            f1_forecast_faixas as _f1, fr1_reunioes_previstas as _fr1,
                            h1_hierarquia as _h1, f2_forecast_media as _f2,
                            lc1_licenciado as _lc1)
    if _COMMIT_CRN:
        crn.versao = lambda: _COMMIT_CRN                    # carimbo mostra o commit carregado
except Exception as _e:   # noqa: BLE001
    _ERRO_CRN = f"{type(_e).__name__}: {_e}"

# Nomes de variável/URL que a CRN espera, mapeados a partir do que o painel já usa na Vercel
# (nada muda nas Environment Variables atuais).
os.environ.setdefault("PIPEDRIVE_TOKEN", os.environ.get("PIPEDRIVE_API_TOKEN", ""))

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

PD_DOMAIN = "boardacademy.pipedrive.com"
PD_TOKEN = os.environ.get("PIPEDRIVE_API_TOKEN", "")

V1_BASE = f"https://{PD_DOMAIN}/api/v1"
V2_BASE = f"https://{PD_DOMAIN}/api/v2"

FILTER_FORECAST = 1490240     # pipeline aberto (forecast)
FILTER_REFERIDOS = 1562285    # indicações (não usado neste painel por ora)

CF_MULTIPLICADOR = "7e0e43c2734751f77be292a72527f638a850ad50"
CF_QUALIFICADOR = "a6f13cc27c8d041f3af4091283ce0d4fe0913875"
CF_UTM_CAMPAIGN = "ae03fa460a108b8cdfa87e97ebca24379d2779d6"
CF_PRODUTO = "8bdce76ba66f0fed0280918a4845190c92899ed5"
CF_NOME_PRODUTO = "09d57fd58b8cac693f5901417f758df746223273"
CF_DATA_ULTIMA_APLICACAO = "23de049432e523993f69ecd456a3f755c0f07f3d"
CF_UTM_SOURCE = "8fb3221ab3d91cddaf51e0a9e1bbcda34fc9d28e"
CF_CARGO_NEGOCIO = "718c8aba81211c883ffd9f4616f75ee22a10b2da"  # campo "Cargo" do negócio (diferente do Cargo da COLAB)
CF_PONTUACAO = "673fc401f8cb6d72a16ee7775464f80ec974e2b6"  # "Pontuação"/score — só roda pra LEAN e CES

# Lista MQL_BASE da skill "filtros-mqls-pipedrive" — usada pra achar o Cargo que qualifica um lead PFCC como MQL
PALAVRAS_MQL_EXATAS = [
    "socio", "sócio", "partner", "founder", "fundador", "proprietario", "proprietário", "owner",
    "empresario", "empresário", "entrepreneur", "presidente", "president",
    "ceo", "cmo", "cto", "coo", "cfo", "cpo", "ciso", "chro", "cro", "cio", "chief",
    "c-level", "c level", "vice-presidente", "vp",
    "director", "advisor", "board member",
]
PALAVRAS_MQL_PREFIXO = ["conselh", "diretor"]  # conselh* / diretor* — aceita continuação (conselheiro, diretoria...)
# mantido só pra compatibilidade de quem ainda importa PALAVRAS_MQL diretamente
PALAVRAS_MQL = PALAVRAS_MQL_EXATAS + PALAVRAS_MQL_PREFIXO

PRODUTOS = ["PFCC", "CES", "LEAN", "ABP", "COAUTORIA"]

# IDs de campanha do Google Ads que também identificam PFCC (além de conter "pfcc" na utm_campaign)
PFCC_IDS_GOOGLE = {
    "22822740395", "23869952762", "22675029313", "23874401647",
    "23397388987", "22508119133", "23963147748",
}

# Skill "filtros-mqls-pipedrive": pra ser PFCC, a utm_campaign NÃO pode conter nada disso
PFCC_EXCLUSOES_CAMPANHA = [
    "mexico", "chile", "lean", "ces",
    "hotseat", "summit", "ppc", "pesquisa", "masterclass", "coautoria", "workshop", "lic-bc",
]
# valores "vazios" que às vezes aparecem escritos por extenso no campo
VALORES_VAZIOS_TEXTO = {"", "null", "none", "undefined", "-", "n/a", "na"}

PONTUACAO_MINIMA_LEAN_CES = 6

SHEET_COLAB = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSvwO3Ag2f2cbkVgR1pJZp6fANQcbualGKlAG50fmOljuEGKZ1gJBbSAjRdO3SomXUEVQOWnTvlfHRd/pub?gid=1782440078&single=true&output=csv"
SHEET_METAS = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSvwO3Ag2f2cbkVgR1pJZp6fANQcbualGKlAG50fmOljuEGKZ1gJBbSAjRdO3SomXUEVQOWnTvlfHRd/pub?gid=0&single=true&output=csv"
SHEET_FERIADOS = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSvwO3Ag2f2cbkVgR1pJZp6fANQcbualGKlAG50fmOljuEGKZ1gJBbSAjRdO3SomXUEVQOWnTvlfHRd/pub?gid=1010928978&single=true&output=csv"

os.environ.setdefault("URL_COLAB", SHEET_COLAB)
os.environ.setdefault("URL_METAS", SHEET_METAS)
os.environ.setdefault("URL_FERIADOS", SHEET_FERIADOS)

# Mapa interno -> exibição. Chave interna "mgm" (funis Olympus + Navigator) exibe como "Olympus".
SQUAD_DISPLAY = {"mgm": "Olympus", "sniper": "Sniper", "navigator": "Navigator"}
SQUADS_FINANCEIROS = ["mgm"]     # closers (valor em R$)
SQUAD_SDR = "sniper"                       # reuniões

# ⚠️ Parâmetros que NÃO vêm de nenhuma das 3 abas do Sheets (COLAB/METAS/FERIADOS):
# percentual do "gap intermediário" e a data-prazo (ex.: 40% até 16/07/2026).
# Por ora como constantes — mover pro Sheets se o time preferir editar sem deploy.
PCT_GAP_INTERMEDIARIO = 0.40



# ---------------------------------------------------------------------------
# UTILS
# ---------------------------------------------------------------------------

def norm(s):
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = unicodedata.normalize("NFD", s).encode("ascii", "ignore").decode()
    return s


def parse_num_br(v):
    if v is None:
        return 0.0
    s = str(v).strip()
    if s == "":
        return 0.0
    s = s.replace("R$", "").strip()
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def safe_div(a, b):
    try:
        if b == 0:
            return 0.0
        v = a / b
        if v != v or v in (float("inf"), float("-inf")):  # NaN/Inf
            return 0.0
        return round(v, 4)
    except Exception:
        return 0.0


def to_brt(iso_ts):
    """R0 (CRN): Pipedrive UTC -> Brasília (America/Sao_Paulo), via crn.base.fuso. Devolve datetime sem tzinfo."""
    d = _fuso.para_brasil(iso_ts)
    return d.replace(tzinfo=None) if d else None


def hoje_brt():
    """R0 (CRN): 'hoje' em Brasília, nunca o relógio do servidor."""
    return _fuso.hoje_brasil()


from urllib.error import HTTPError, URLError

# ---------------------------------------------------------------------------
# LOGIN / CONTROLE DE ACESSO POR PAPEL — mesma lógica de login.py e me.py.
# admin: vê tudo. elite/sniper/olympus: Métricas e Produtos iguais pra todos;
# na aba Perdidos, só o próprio funil. Sem login: Perdidos fica bloqueado.
# ---------------------------------------------------------------------------
import hmac
import base64
import time as _time
import hashlib

PAPEIS_VALIDOS = {"admin", "elite", "sniper", "olympus"}
AUTH_SECRET = (os.environ.get("AUTH_SECRET", "") or "troque-este-segredo").encode()
PAPEL_PARA_SQUAD_DISPLAY = {"elite": "Elite", "sniper": "Sniper", "olympus": "Olympus"}
PAPEL_PARA_SQUAD_INTERNO = {"elite": "elite", "sniper": "sniper", "olympus": "mgm"}


def _valida_token(token):
    if not token:
        return None
    try:
        bruto = base64.urlsafe_b64decode(token.encode()).decode()
        usuario, papel, exp, assinatura = bruto.rsplit("|", 3)
        corpo = f"{usuario}|{papel}|{exp}"
        esperado = hmac.new(AUTH_SECRET, corpo.encode(), hashlib.sha256).hexdigest()
        if not hmac.compare_digest(esperado, assinatura):
            return None
        if int(exp) < int(_time.time()):
            return None
        if papel not in PAPEIS_VALIDOS:
            return None
        return {"usuario": usuario, "papel": papel}
    except Exception:
        return None


def sessao_da_requisicao(headers):
    auth = headers.get("Authorization") or headers.get("authorization") or ""
    if auth.startswith("Bearer "):
        return _valida_token(auth[7:])
    return None


def _sha256_senha(txt):
    return hashlib.sha256(txt.encode("utf-8")).hexdigest()


def _carregar_usuarios_painel():
    raw = os.environ.get("USUARIOS_PAINEL", "") or ""
    raw = raw.strip().strip('"').strip("'").strip()
    usuarios = {}
    for par in raw.split(","):
        par = par.strip()
        if not par:
            continue
        partes = par.split(":")
        if len(partes) != 3:
            continue
        u, h, papel = partes
        papel = papel.strip().lower()
        if papel not in PAPEIS_VALIDOS:
            continue
        usuarios[u.strip().lower()] = {"hash": h.strip().lower(), "papel": papel}
    return usuarios


def checa_login(usuario, senha):
    usuarios = _carregar_usuarios_painel()
    info = usuarios.get(str(usuario).strip().lower())
    if info and hmac.compare_digest(info["hash"], _sha256_senha(str(senha))):
        return info["papel"]
    return None


def gera_token(usuario, papel):
    exp = int(_time.time()) + 12 * 3600
    corpo = f"{usuario}|{papel}|{exp}"
    assinatura = hmac.new(AUTH_SECRET, corpo.encode(), hashlib.sha256).hexdigest()
    bruto = f"{corpo}|{assinatura}"
    return base64.urlsafe_b64encode(bruto.encode()).decode()


def filtrar_perdidos_por_papel(perdidos_analise, papel):
    """admin (ou sem papel reconhecido de restrição) vê tudo; elite/sniper/olympus só o próprio funil."""
    if papel == "admin":
        return perdidos_analise
    squad_nome = PAPEL_PARA_SQUAD_DISPLAY.get(papel)
    if not squad_nome:
        return None  # sem login válido -> Perdidos fica bloqueado
    vazio = {"total_mes": 0, "total_hoje": 0, "motivos": [], "auditoria": []}
    dados_funil = (perdidos_analise.get("por_funil") or {}).get(squad_nome, vazio)
    return {
        "geral": dados_funil,
        "por_funil": {squad_nome: dados_funil},
        "responsaveis": perdidos_analise.get("responsaveis", []),
    }



def http_get_json(url, headers=None, tentativas=2):
    req = Request(url, headers=headers or {})
    ultimo_erro = None
    for tentativa in range(tentativas):
        try:
            with urlopen(req, timeout=20) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except HTTPError as e:
            detalhe = e.read().decode("utf-8", errors="ignore")
            raise RuntimeError(f"Erro HTTP {e.code} em {url}: {detalhe[:600]}") from e
        except (URLError, TimeoutError, socket.timeout) as e:
            ultimo_erro = e
            continue  # tenta de novo (timeout de leitura costuma ser passageiro)
    razao = getattr(ultimo_erro, "reason", ultimo_erro)
    raise RuntimeError(f"Falha de conexão (após {tentativas} tentativas) em {url}: {razao}") from ultimo_erro


def http_get_text(url, tentativas=3):
    req = Request(url, headers={"User-Agent": "Mozilla/5.0"})
    ultimo_erro = None
    for tentativa in range(tentativas):
        try:
            with urlopen(req, timeout=30) as resp:
                return resp.read().decode("utf-8")
        except HTTPError as e:
            detalhe = e.read().decode("utf-8", errors="ignore")
            raise RuntimeError(f"Erro HTTP {e.code} em {url}: {detalhe[:600]}") from e
        except (URLError, TimeoutError, socket.timeout) as e:
            ultimo_erro = e
            continue
    razao = getattr(ultimo_erro, "reason", ultimo_erro)
    raise RuntimeError(f"Falha de conexão (após {tentativas} tentativas) em {url}: {razao}") from ultimo_erro



def find_col(fieldnames, exact=None, contains=None, contains_all=None):
    norm_map = {norm(f): f for f in fieldnames}
    if exact:
        return norm_map.get(norm(exact))
    if contains:
        for nf, orig in norm_map.items():
            if contains in nf:
                return orig
    if contains_all:
        for nf, orig in norm_map.items():
            if all(c in nf for c in contains_all):
                return orig
    return None


# ---------------------------------------------------------------------------
# GOOGLE SHEETS
# ---------------------------------------------------------------------------

def read_csv(url):
    text = http_get_text(url)
    return list(csv.DictReader(io.StringIO(text)))


# Pedido do João (26/10/2026): estas pessoas ficam fora de todas as análises deste painel.
# É um recorte do painel, não muda regra da CRN.
EXCLUIDOS_PAINEL = {"mylena oliveira", "raphaela moutinho"}

SQUAD_CRN_PARA_INTERNO = {"olympus": "mgm", "sniper": "sniper", "elite": "elite", "navigator": "navigator"}


def carregar_feriados():
    """A2 (CRN): tabela de Feriados."""
    return _fontes.carregar_feriados()


def carregar_colaboradores(mes, ano):
    """Colab do mês pela CRN (R1/H1). Devolve o mesmo formato que o painel já usava
    (chave = nome normalizado), com a equipe em chave interna: Olympus -> 'mgm'."""
    colab = _fontes.carregar_colab(ano, mes)
    out = {}
    for p in colab:
        if p["nome_norm"] in EXCLUIDOS_PAINEL:
            continue
        equipe = _m1.squad_da_pessoa(p)  # H1/M1: gerente tem equipe fixa na config
        out[p["nome_norm"]] = {
            "nome_exibicao": p["nome"],
            "subarea": SQUAD_CRN_PARA_INTERNO.get(norm(equipe), norm(equipe)),
            "cargo": norm(p.get("cargo", "")),
            "head": p["head_norm"], "lider": p["lider_norm"],
            "_pessoa": p, "_colab": colab,
        }
    return out


def carregar_metas(mes, ano):
    """Metas do mês pela CRN (R2: meta de reunião > 0 = SDR)."""
    out = {}
    for m in _fontes.carregar_metas(ano, mes):
        if m["nome_norm"] in EXCLUIDOS_PAINEL:
            continue
        if not m["meta_reunioes"] and not m["meta_financeira"]:
            continue
        out[m["nome_norm"]] = {"meta_reu": m["meta_reunioes"], "meta_fin": m["meta_financeira"],
                               "papel": m["papel"]}
    return out


# ---------------------------------------------------------------------------
# DIAS ÚTEIS (A2 — CRN)
# ---------------------------------------------------------------------------

def eh_dia_util(d, feriados):
    return d.weekday() < 5 and d not in feriados


def calcular_dias_uteis(ano, mes, feriados, hoje=None):
    """A2 (CRN): total do mês, passados (1º dia até D-1) e restantes (hoje até o fim).
    'fechados' = 'passados' (mesma coisa pela A2); mantido por compatibilidade do painel."""
    du = _fontes.dias_uteis_periodo(ano, mes, feriados, hoje or hoje_brt())
    return {"total": du["total"], "passados": du["passados"], "restantes": du["restantes"],
            "fechados": du["passados"]}


def proximo_dia_util(a_partir_de, feriados):
    d = a_partir_de + dt.timedelta(days=1)
    while not eh_dia_util(d, feriados):
        d += dt.timedelta(days=1)
    return d


def dia_util_anterior(a_partir_de, feriados):
    d = a_partir_de - dt.timedelta(days=1)
    while not eh_dia_util(d, feriados):
        d -= dt.timedelta(days=1)
    return d


# ---------------------------------------------------------------------------
# PIPEDRIVE
# ---------------------------------------------------------------------------





def pd_users():
    data = http_get_json(f"{V1_BASE}/users?{urlencode({'api_token': PD_TOKEN})}")
    return {u["id"]: u.get("name", "") for u in (data.get("data") or [])}


@functools.lru_cache(maxsize=32)
def buscar_pipeline_id_por_nome(nome_busca):
    """Acha o pipeline_id cujo nome contém nome_busca (case-insensitive)."""
    data = http_get_json(f"{V1_BASE}/pipelines?{urlencode({'api_token': PD_TOKEN})}")
    alvo = norm(nome_busca)
    for p in (data.get("data") or []):
        if alvo in norm(p.get("name", "")):
            return p.get("id")
    return None


@functools.lru_cache(maxsize=32)
def buscar_stages_pipeline(pipeline_id):
    """Mapa stage_id -> {'norm': nome normalizado, 'nome': nome original} das etapas de um pipeline."""
    data = http_get_json(f"{V1_BASE}/stages?{urlencode({'api_token': PD_TOKEN, 'pipeline_id': pipeline_id})}")
    mapa = {}
    for s in (data.get("data") or []):
        mapa[s.get("id")] = {"norm": norm(s.get("name", "")), "nome": s.get("name", "")}
    return mapa


@functools.lru_cache(maxsize=64)
def _buscar_deals_por_pipeline_cached(pipeline_id, status, desde_iso, ate_iso):
    itens, cursor = [], None
    while True:
        params = {"status": status, "pipeline_id": pipeline_id, "limit": 500}
        if cursor:
            params["cursor"] = cursor
        if desde_iso:
            params["updated_since"] = desde_iso
        if ate_iso:
            params["updated_until"] = ate_iso
        headers = {"x-api-token": PD_TOKEN}
        data = http_get_json(f"{V2_BASE}/deals?{urlencode(params)}", headers=headers)
        itens.extend(data.get("data") or [])
        cursor = (data.get("additional_data") or {}).get("next_cursor")
        if not cursor:
            break
    return tuple(itens)


def buscar_deals_por_pipeline(pipeline_id, status, desde_iso=None, ate_iso=None):
    """Negócios de um status específico (open/won/lost) de um pipeline — direto,
    sem passar por nenhum filter_id salvo. Usa API v2: a v1 ignora silenciosamente
    o parâmetro pipeline_id (por isso todos os funis vinham juntos antes).
    desde_iso/ate_iso limitam por update_time — reduz MUITO o volume pra won/lost
    (senão baixaria o histórico inteiro do funil e estourava o timeout).
    Cacheado por (pipeline_id, status, desde_iso, ate_iso): Produtos/Perdidos/MQL's
    buscam exatamente os mesmos negócios várias vezes na mesma execução — retorna
    sempre uma lista NOVA (cópia), então quem chama pode mexer nela à vontade."""
    return list(_buscar_deals_por_pipeline_cached(pipeline_id, status, desde_iso, ate_iso))






def cf_valor(deal, hash_):
    v = deal.get(hash_)
    if v is None:
        v = (deal.get("custom_fields") or {}).get(hash_)  # formato v2
    if isinstance(v, dict):
        return v.get("value")
    return v


def owner_nome(deal, users_map):
    owner = deal.get("user_id")
    if owner is None:
        owner = deal.get("owner_id")  # formato v2
    if isinstance(owner, dict):
        return owner.get("name", "")
    return users_map.get(owner, "")


def buscar_deals_ganhos(ano, mes, users_map, com_exclusoes=False):
    """M1 (CRN): ganhos do mês pelo won_time em Brasília, busca só pelo período (sem filtro salvo).
    Devolve só os válidos: saem cortesia (R4), outra moeda (R3) e ganho automático (G1).
    com_exclusoes=True devolve o dict completo {'validos','cortesias','outra_moeda','automaticos'}."""
    ini = dt.date(ano, mes, 1)
    fim = dt.date(ano, mes, calendar.monthrange(ano, mes)[1])
    deals = _pd.ganhos_do_periodo(ini, fim)
    sep = _m1.separar_ganhos(deals, ano, mes, users_map)
    return sep if com_exclusoes else sep["validos"]


@functools.lru_cache(maxsize=1)
def _funis_cached():
    return _pd.funis()


def _ganhos_por_squad(deals, colaboradores, users_map):
    """M1 (CRN): ganhos válidos agrupados pelo squad da venda (chave interna; 'mgm' = Olympus)."""
    colab_lista = next(iter(colaboradores.values()))["_colab"] if colaboradores else []
    por = _m1.ganhos_por_closer(deals, colab_lista, _funis_cached(), users_map)
    out = {}
    for r in por.values():
        for d in r["deals"]:
            sq = r["squad_por_deal"][d.get("id")]
            out.setdefault(SQUAD_CRN_PARA_INTERNO.get(norm(sq), norm(sq)), []).append(d)
    return out


def _meta_financeira_squad(squad_interno, colaboradores, metas):
    """A1: meta do squad = soma das metas financeiras de quem é Closer/IS (R2) naquele squad."""
    return sum(m["meta_fin"] for nome, m in metas.items()
               if m.get("papel") == "closer" and colaboradores.get(nome, {}).get("subarea") == squad_interno)




def squad_do_deal(deal, colaboradores, users_map):
    """M1/H1 (CRN): squad da venda pelo dono na Colab do mês (gerente = equipe fixa; head = squad do funil;
    licenciado fora). Devolve a chave interna do painel ('mgm' = Olympus) ou None."""
    c = colaboradores.get(norm(owner_nome(deal, users_map)))
    if not c:
        return None
    funis = _funis_cached()
    if _lc1.eh_licenciado(deal, funis):
        return None
    sq = _m1.squad_da_venda(deal, c["_pessoa"], c["_colab"], funis)
    return SQUAD_CRN_PARA_INTERNO.get(norm(sq), norm(sq))


























# ---------------------------------------------------------------------------
# MONTAGEM DO PAINEL
# ---------------------------------------------------------------------------

def buscar_forecast_deals():
    """F1 (CRN): deals abertos do filtro de forecast (1490240, de crn.base.config)."""
    return _pd.deals_do_filtro(crn.base.config.filtro("forecast"), "open")




def _deals_por_squad_forecast(pool, colaboradores, users_map):
    por_squad = {s: [] for s in SQUADS_FINANCEIROS}
    vistos = set()
    for d in pool:
        did = d.get("id")
        if did in vistos or d.get("status") != "open":
            continue
        vistos.add(did)
        sq = squad_do_deal(d, colaboradores, users_map)
        if sq in por_squad:
            por_squad[sq].append(d)
    return por_squad


def valor_previsto_por_squad(pool, colaboradores, users_map, data_alvo):
    """F1 + F2 (CRN): faixas 20/50/70 por squad na data prevista e média ponderada."""
    out = {}
    for sq, lista in _deals_por_squad_forecast(pool, colaboradores, users_map).items():
        f = _f1.faixas(lista, data_alvo, data_alvo)
        out[sq] = {"p20": f[20], "p50": f[50], "p70": f[70], "media": _f2.media_ponderada(f),
                   "em_aberto": f["em_aberto"]}
    return out


def valor_abertos_por_squad(pool, colaboradores, users_map, data_alvo):
    """F1 (CRN): em aberto = soma das três faixas na data prevista."""
    return {sq: v["em_aberto"] for sq, v in valor_previsto_por_squad(pool, colaboradores, users_map, data_alvo).items()}


def _match_produto_no_texto(texto):
    """Match simples por substring — usado só pro campo 'Produto'/'Nome produto' dos ganhos
    (não passa pelas exclusões/pontuação, que são regras específicas da utm_campaign)."""
    t = norm(texto or "")
    if "pfcc" in t or any(gid in t for gid in PFCC_IDS_GOOGLE):
        return "PFCC"
    if "lean" in t:
        return "LEAN"
    if "ces" in t:
        return "CES"
    return None


def _pfcc_valido(utm_raw):
    """Regras da skill 'filtros-mqls-pipedrive' pro PFCC: a utm_campaign precisa conter
    'pfcc' (ou um dos IDs do Google), não pode estar vazia/ser um placeholder tipo
    'null'/'n/a', e não pode conter nenhuma das palavras de exclusão."""
    if utm_raw is None:
        return False
    bruto = str(utm_raw).strip()
    t = norm(bruto)
    if t in VALORES_VAZIOS_TEXTO:
        return False
    if any(excl in t for excl in PFCC_EXCLUSOES_CAMPANHA):
        return False
    return "pfcc" in t or any(gid in t for gid in PFCC_IDS_GOOGLE)


def _match_abp_coautoria(deal):
    """ABP e COAUTORIA são identificados pelo campo 'Nome produto' (ou 'Produto'), não pela utm_campaign."""
    for campo in (CF_NOME_PRODUTO, CF_PRODUTO):
        valor = norm(cf_valor(deal, campo) or "")
        if "abp" in valor:
            return "ABP"
        if "coautoria" in valor:
            return "COAUTORIA"
    return None


def classificar_produto(deal):
    """PFCC: regras completas da skill (utm contém pfcc/ID do Google, sem exclusões, sem placeholder).
    LEAN/CES: utm contém a palavra E Pontuação >= 6 (score que só roda pra esses dois).
    ABP/COAUTORIA: pelo campo 'Nome produto'/'Produto'."""
    produto = _match_abp_coautoria(deal)
    if produto:
        return produto

    utm_raw = cf_valor(deal, CF_UTM_CAMPAIGN)
    if _pfcc_valido(utm_raw):
        return "PFCC"

    t = norm(utm_raw or "")
    try:
        pontuacao_raw = cf_valor(deal, CF_PONTUACAO)
        pontuacao_num = float(pontuacao_raw) if pontuacao_raw not in (None, "") else None
    except (TypeError, ValueError):
        pontuacao_num = None

    if pontuacao_num is not None and pontuacao_num >= PONTUACAO_MINIMA_LEAN_CES:
        if "lean" in t:
            return "LEAN"
        if "ces" in t:
            return "CES"
    return None


def classificar_produto_ganho(deal):
    """Só pra negócios GANHOS: prioriza o produto que o cliente efetivamente comprou
    ('Produto' / 'Nome produto'), já que pode ser diferente da campanha de origem
    (utm_campaign) — cliente entra por uma campanha mas fecha em outro produto.
    Cai pra utm_campaign só se os dois campos vierem vazios."""
    for campo in (CF_PRODUTO, CF_NOME_PRODUTO):
        produto = _match_produto_no_texto(cf_valor(deal, campo))
        if produto:
            return produto
    return classificar_produto(deal)


PRODUTOS_TODOS = PRODUTOS + ["Não classificado"]


SQUAD_NOME_PIPELINE = {"mgm": "olympus"}  # nome do funil no Pipedrive
PIPELINE_ID_SNIPER = 88  # ID direto, sem precisar buscar por nome
PIPELINE_ID_NAVIGATOR = 51  # ID direto — só aparece na aba Produtos, não em Perdidos
SQUADS_PRODUTOS = ["mgm", "sniper"]  # squads usados em Perdidos (Olympus/Sniper)
SQUADS_PRODUTOS_ABA = SQUADS_PRODUTOS + ["navigator"]  # squads que aparecem na aba Produtos - Detalhamento


def _pipeline_id_do_squad(squad_interno):
    if squad_interno == "sniper":
        return PIPELINE_ID_SNIPER
    if squad_interno == "navigator":
        return PIPELINE_ID_NAVIGATOR
    return buscar_pipeline_id_por_nome(SQUAD_NOME_PIPELINE[squad_interno])


def _cargo_eh_mql(cargo_raw):
    """Cargo contém um termo da lista MQL — com word-boundary nas siglas curtas,
    pra 'cio'/'cro'/'vp' etc não baterem dentro de palavras comuns tipo 'negócios'."""
    cargo_norm = norm(cargo_raw or "")
    if not cargo_norm:
        return False
    for termo in PALAVRAS_MQL_EXATAS:
        if re.search(r"\b" + re.escape(norm(termo)) + r"\b", cargo_norm):
            return True
    for termo in PALAVRAS_MQL_PREFIXO:
        if re.search(r"\b" + re.escape(norm(termo)), cargo_norm):
            return True
    return False


def analisar_mql_pfcc(ano, mes, hoje):
    """'NOVOS MQL's PFCC' — skill 'filtros-mqls-pipedrive': negócio precisa ser um PFCC
    válido (utm sem exclusões, sem placeholder) + utm_source != 'org' + Cargo contendo
    uma palavra da lista MQL + Data da última aplicação no dia (hoje) ou mês/ano (mês)."""
    inicio_mes = dt.date(ano, mes, 1)
    fim_mes = dt.date(ano + 1, 1, 1) - dt.timedelta(days=1) if mes == 12 else dt.date(ano, mes + 1, 1) - dt.timedelta(days=1)
    desde_iso = f"{inicio_mes.isoformat()}T00:00:00Z"
    ate_iso = f"{fim_mes.isoformat()}T23:59:59Z"

    pipelines = [PIPELINE_ID_SNIPER]
    for nome_pipeline in SQUAD_NOME_PIPELINE.values():  # olympus, elite
        pid = buscar_pipeline_id_por_nome(nome_pipeline)
        if pid is not None:
            pipelines.append(pid)

    variantes_dia = {hoje.isoformat(), hoje.strftime("%d/%m/%Y"), hoje.strftime("%d-%m-%Y")}
    variantes_mes = {f"{ano:04d}-{mes:02d}", f"{mes:02d}/{ano:04d}"}

    vistos = set()
    total_hoje = 0
    total_mes = 0
    total_mes_novos = 0
    total_mes_reaplicados = 0
    for pid in pipelines:
        deals = buscar_deals_por_pipeline(pid, "open")
        for status in ("won", "lost"):
            deals.extend(buscar_deals_por_pipeline(pid, status, desde_iso, ate_iso))
        for d in deals:
            did = d.get("id")
            if did in vistos:
                continue
            vistos.add(did)

            if not _pfcc_valido(cf_valor(d, CF_UTM_CAMPAIGN)):
                continue
            cargo = cf_valor(d, CF_CARGO_NEGOCIO)
            if not (cargo and str(cargo).strip()):
                continue
            if not _cargo_eh_mql(cargo):
                continue

            valor_data = str(cf_valor(d, CF_DATA_ULTIMA_APLICACAO) or "")
            if any(v in valor_data for v in variantes_dia):
                total_hoje += 1
            if any(v in valor_data for v in variantes_mes):
                total_mes += 1
                # Novo: criado no mesmo dia da aplicação. Reaplicação: datas diferentes
                # (ex: aplicação é desse mês, mas o negócio foi criado antes).
                data_aplicacao = _parse_data_campo(cf_valor(d, CF_DATA_ULTIMA_APLICACAO))
                add_brt = to_brt(d.get("add_time"))
                if data_aplicacao and add_brt and add_brt.date() == data_aplicacao:
                    total_mes_novos += 1
                else:
                    total_mes_reaplicados += 1

    return {
        "hoje": total_hoje,
        "mes": total_mes,
        "mes_novos": total_mes_novos,
        "mes_reaplicados": total_mes_reaplicados,
    }


# ---------------------------------------------------------------------------
# ABA "METAS BOARD ACADEMY" — spec-painel-metas-board-academy.md
# Sempre nos 4 pipelines: Sniper(88) + Elite(87) + Olympus(46) + Navigator(51).
# ---------------------------------------------------------------------------
PIPELINE_ID_OLYMPUS_BA = 46
PIPELINE_ID_ELITE_BA = 87
PIPELINES_METAS_BA = {"Sniper": PIPELINE_ID_SNIPER, "Elite": PIPELINE_ID_ELITE_BA,
                       "Olympus": PIPELINE_ID_OLYMPUS_BA, "Navigator": PIPELINE_ID_NAVIGATOR}

META_RECEITA_MES = 2_700_000.0
META_TKM_BLENDED = 13140.34
META_GANHOS_MES = 206
META_GANHOS_POR_PRODUTO = {"PFCC": 155, "LEAN": 17, "CES": 8, "ABP": 5, "Outros": 21}
TX_MQL_PARA_GANHO = 0.05
TX_LEAD_PARA_MQL = 0.44
META_MQL_TOTAL = 4120
META_MQL_POR_PRODUTO = {"PFCC": 3100, "LEAN": 340, "CES": 160}
META_LEADS_TOTAL = 9364
TKM_FIXO_POR_PRODUTO = {"PFCC": 15000.0, "LEAN": 8990.0, "CES": 12990.0, "ABP": 25000.0}

HISTORICO_JAN_JUN_2026 = {
    "negocios_criados": 22442,
    "negocios_ganhos": 445,
    "tx_conversao_pct": 1.98,
    "mix": [
        {"produto": "PFCC", "qtd": 334, "pct": 75.06},
        {"produto": "LEAN", "qtd": 38, "pct": 8.54},
        {"produto": "CES", "qtd": 17, "pct": 3.82},
        {"produto": "ABP", "qtd": 11, "pct": 2.47},
        {"produto": "Outros", "qtd": 45, "pct": 10.11},
    ],
}


def _produto_lead_gerado_ba(deal):
    """2.2 — Leads Gerados por produto: cascata simples via utm_campaign."""
    utm = norm(cf_valor(deal, CF_UTM_CAMPAIGN) or "")
    if "lean" in utm:
        return "LEAN"
    if "ces" in utm:
        return "CES"
    if "pfcc" in utm and "workshop" not in utm:
        return "PFCC"
    return "Outros"


def _mql_lean_ces_ba(deal):
    """2.3 — LEAN/CES: Pontuação >= 6 (utm já garantido por quem chama)."""
    try:
        pontuacao_raw = cf_valor(deal, CF_PONTUACAO)
        pontuacao = float(pontuacao_raw) if pontuacao_raw not in (None, "") else None
    except (TypeError, ValueError):
        pontuacao = None
    return pontuacao is not None and pontuacao >= PONTUACAO_MINIMA_LEAN_CES


def _produto_ganho_ba(deal):
    """2.4 — Negócios Ganhos: cascata lean, pfcc, ces, campo Produto, Outros."""
    utm = norm(cf_valor(deal, CF_UTM_CAMPAIGN) or "")
    if "lean" in utm:
        return "LEAN"
    if "pfcc" in utm:
        return "PFCC"
    if "ces" in utm:
        return "CES"
    produto_campo = norm(cf_valor(deal, CF_PRODUTO) or cf_valor(deal, CF_NOME_PRODUTO) or "")
    for produto, chave in (("LEAN", "lean"), ("PFCC", "pfcc"), ("CES", "ces"), ("ABP", "abp")):
        if chave in produto_campo:
            return produto
    return "Outros"


def _nome_base_titulo(titulo):
    """Remove o sufixo 'N / M' (aceita variações de espaço ao redor da barra) do título."""
    if not titulo:
        return ""
    m = re.search(r"\s*\d+\s*/\s*\d+\s*$", titulo)
    if m:
        return titulo[: m.start()].strip()
    return titulo.strip()


def consolidar_parcelas_ba(deals_ganhos):
    """Agrupa negócios ganhos que são parcelas do mesmo pagamento (título 'Nome N/M'):
    soma o valor, conta como 1 negócio ganho só. Produto do grupo = produto do 1º deal."""
    grupos = {}
    ordem = []
    for d in deals_ganhos:
        titulo = d.get("title") or ""
        nome_exibicao = _nome_base_titulo(titulo) or titulo
        chave = norm(nome_exibicao)
        if chave not in grupos:
            grupos[chave] = {"deals": [], "valor_total": 0.0, "nome": nome_exibicao}
            ordem.append(chave)
        grupos[chave]["deals"].append(d)
        grupos[chave]["valor_total"] += float(d.get("value") or 0)

    consolidados = []
    for chave in ordem:
        g = grupos[chave]
        consolidados.append({
            "nome": g["nome"],
            "valor": round(g["valor_total"], 2),
            "produto": _produto_ganho_ba(g["deals"][0]),
            "qtd_parcelas": len(g["deals"]),
            "ids": [dd.get("id") for dd in g["deals"]],
        })
    return consolidados


def montar_metas_board_academy(ano, mes, hoje, ontem):
    inicio_mes = dt.date(ano, mes, 1)
    fim_mes = dt.date(ano + 1, 1, 1) - dt.timedelta(days=1) if mes == 12 else dt.date(ano, mes + 1, 1) - dt.timedelta(days=1)
    desde_iso = f"{inicio_mes.isoformat()}T00:00:00Z"
    ate_iso = f"{fim_mes.isoformat()}T23:59:59Z"

    todos_deals = []
    deals_ganhos_mes = []
    vistos_gerais = set()

    for pid in PIPELINES_METAS_BA.values():
        abertos = buscar_deals_por_pipeline(pid, "open")
        ganhos = buscar_deals_por_pipeline(pid, "won", desde_iso, ate_iso)
        perdidos = buscar_deals_por_pipeline(pid, "lost", desde_iso, ate_iso)
        for d in abertos + ganhos + perdidos:
            did = d.get("id")
            if did not in vistos_gerais:
                vistos_gerais.add(did)
                todos_deals.append(d)
        for d in ganhos:
            won_brt = to_brt(d.get("won_time"))
            if won_brt and (won_brt.year, won_brt.month) == (ano, mes):
                deals_ganhos_mes.append(d)

    # ---- 2.1/2.2/2.3: universo de Leads Gerados (Data da última aplicação no mês) ----
    leads_mes = []
    for d in todos_deals:
        data_aplicacao = _parse_data_campo(cf_valor(d, CF_DATA_ULTIMA_APLICACAO))
        if data_aplicacao and (data_aplicacao.year, data_aplicacao.month) == (ano, mes):
            leads_mes.append((d, data_aplicacao))

    leads_novos = 0
    leads_reaplicacao = 0
    leads_hoje = 0
    leads_ontem = 0
    leads_por_produto = {"PFCC": 0, "LEAN": 0, "CES": 0, "Outros": 0}
    leads_hoje_por_produto = {"PFCC": 0, "LEAN": 0, "CES": 0, "Outros": 0}
    leads_ontem_por_produto = {"PFCC": 0, "LEAN": 0, "CES": 0, "Outros": 0}
    mql_por_produto = {"PFCC": 0, "LEAN": 0, "CES": 0}
    mql_hoje = 0
    mql_ontem = 0
    mql_hoje_por_produto = {"PFCC": 0, "LEAN": 0, "CES": 0}
    mql_ontem_por_produto = {"PFCC": 0, "LEAN": 0, "CES": 0}

    for d, data_aplicacao in leads_mes:
        if data_aplicacao == hoje:
            leads_hoje += 1
        if data_aplicacao == ontem:
            leads_ontem += 1

        add_brt = to_brt(d.get("add_time"))
        if add_brt and add_brt.date() == data_aplicacao:
            leads_novos += 1
        else:
            leads_reaplicacao += 1

        produto_lead = _produto_lead_gerado_ba(d)
        leads_por_produto[produto_lead] = leads_por_produto.get(produto_lead, 0) + 1
        if data_aplicacao == hoje:
            leads_hoje_por_produto[produto_lead] = leads_hoje_por_produto.get(produto_lead, 0) + 1
        if data_aplicacao == ontem:
            leads_ontem_por_produto[produto_lead] = leads_ontem_por_produto.get(produto_lead, 0) + 1

        eh_mql = False
        if produto_lead == "PFCC":
            eh_mql = _pfcc_valido(cf_valor(d, CF_UTM_CAMPAIGN)) and _cargo_eh_mql(cf_valor(d, CF_CARGO_NEGOCIO))
        elif produto_lead in ("LEAN", "CES"):
            eh_mql = _mql_lean_ces_ba(d)
        if eh_mql:
            mql_por_produto[produto_lead] = mql_por_produto.get(produto_lead, 0) + 1
            if data_aplicacao == hoje:
                mql_hoje += 1
                mql_hoje_por_produto[produto_lead] = mql_hoje_por_produto.get(produto_lead, 0) + 1
            if data_aplicacao == ontem:
                mql_ontem += 1
                mql_ontem_por_produto[produto_lead] = mql_ontem_por_produto.get(produto_lead, 0) + 1

    total_leads = len(leads_mes)
    total_mql = sum(mql_por_produto.values())

    # ---- 2.4/2.5: Ganhos consolidados (parcelas) + TKM ----
    consolidados = consolidar_parcelas_ba(deals_ganhos_mes)
    ganhos_por_produto = {}
    receita_por_produto = {}
    for c in consolidados:
        ganhos_por_produto[c["produto"]] = ganhos_por_produto.get(c["produto"], 0) + 1
        receita_por_produto[c["produto"]] = round(receita_por_produto.get(c["produto"], 0.0) + c["valor"], 2)
    total_ganhos = len(consolidados)
    receita_total = round(sum(c["valor"] for c in consolidados), 2)
    tkm_por_produto = {p: round(safe_div(receita_por_produto.get(p, 0.0), ganhos_por_produto.get(p, 0)), 2)
                        for p in ganhos_por_produto}
    tkm_blended = round(safe_div(receita_total, total_ganhos), 2)

    tx_lead_para_mql_real = round(safe_div(total_mql, total_leads) * 100, 1)
    tx_mql_para_ganho_real = round(safe_div(total_ganhos, total_mql) * 100, 1)

    return {
        "quanto_estamos_da_meta": {
            "leads": {"realizado": total_leads, "meta": META_LEADS_TOTAL,
                      "pct": round(safe_div(total_leads, META_LEADS_TOTAL) * 100, 1)},
            "mql": {"realizado": total_mql, "meta": META_MQL_TOTAL,
                    "pct": round(safe_div(total_mql, META_MQL_TOTAL) * 100, 1),
                    "por_produto": {p: {"realizado": mql_por_produto.get(p, 0), "meta": META_MQL_POR_PRODUTO.get(p, 0)}
                                     for p in ("PFCC", "LEAN", "CES")}},
            "ganhos": {"realizado": total_ganhos, "meta": META_GANHOS_MES,
                       "pct": round(safe_div(total_ganhos, META_GANHOS_MES) * 100, 1),
                       "por_produto": {p: {"realizado": ganhos_por_produto.get(p, 0), "meta": META_GANHOS_POR_PRODUTO.get(p, 0)}
                                        for p in META_GANHOS_POR_PRODUTO}},
            "receita": {"realizado": receita_total, "meta": META_RECEITA_MES,
                        "pct": round(safe_div(receita_total, META_RECEITA_MES) * 100, 1)},
            "taxas": {
                "lead_para_mql_real_pct": tx_lead_para_mql_real,
                "lead_para_mql_assumida_pct": round(TX_LEAD_PARA_MQL * 100, 1),
                "mql_para_ganho_real_pct": tx_mql_para_ganho_real,
                "mql_para_ganho_assumida_pct": round(TX_MQL_PARA_GANHO * 100, 1),
            },
        },
        "no_mes_geral": {
            "leads_total": total_leads,
            "leads_novos": leads_novos,
            "leads_reaplicacao": leads_reaplicacao,
            "leads_hoje": leads_hoje,
            "leads_ontem": leads_ontem,
            "leads_por_produto": leads_por_produto,
            "leads_hoje_por_produto": leads_hoje_por_produto,
            "leads_ontem_por_produto": leads_ontem_por_produto,
            "mql_total": total_mql,
            "mql_hoje": mql_hoje,
            "mql_ontem": mql_ontem,
            "mql_por_produto": mql_por_produto,
            "mql_hoje_por_produto": mql_hoje_por_produto,
            "mql_ontem_por_produto": mql_ontem_por_produto,
            "ganhos_total": total_ganhos,
            "ganhos_por_produto": ganhos_por_produto,
            "receita_total": receita_total,
            "receita_por_produto": receita_por_produto,
            "tkm_por_produto": tkm_por_produto,
            "tkm_blended": tkm_blended,
        },
        "historico": HISTORICO_JAN_JUN_2026,
        "metas_config": {
            "receita_mensal": META_RECEITA_MES,
            "tkm_blended_meta": META_TKM_BLENDED,
            "ganhos_mes": META_GANHOS_MES,
            "ganhos_por_produto": META_GANHOS_POR_PRODUTO,
            "tx_mql_para_negocio": TX_MQL_PARA_GANHO,
            "tx_lead_para_mql": TX_LEAD_PARA_MQL,
            "mql_total": META_MQL_TOTAL,
            "mql_por_produto": META_MQL_POR_PRODUTO,
            "leads_total": META_LEADS_TOTAL,
            "tkm_fixo_por_produto": TKM_FIXO_POR_PRODUTO,
        },
    }


def montar_resposta_metas_ba(mes_param, ano_param):
    """Rota leve — só a aba 'Metas Board Academy', carregada sob demanda (não entra
    no payload principal, é pesada: 4 pipelines x 3 status cada)."""
    hoje = hoje_brt()
    ano, mes = ano_param or hoje.year, mes_param or hoje.month
    ontem = dia_util_anterior(hoje, carregar_feriados())
    return {"metas_board_academy": montar_metas_board_academy(ano, mes, hoje, ontem)}


def montar_atingimento_historico(ano, mes, dia_inicio, dia_fim):
    """Só cálculo (não vai pro painel) — reconstrói o Atingimento (%) de cada dia,
    olhando só os negócios que já estavam ganhos até aquela data (won_time <= dia).
    Meta do mês é fixa (não muda dia a dia), só o Realizado acumulado muda."""
    colaboradores = carregar_colaboradores(mes, ano)
    users_map = pd_users()
    ganhos_mes = buscar_deals_ganhos(ano, mes, users_map)

    metas = carregar_metas(mes, ano)

    metas_mes = {s: _meta_financeira_squad(s, colaboradores, metas) for s in SQUADS_FINANCEIROS}
    metas_mes["total"] = sum(metas_mes.values())

    por_dia = []
    for dia in range(dia_inicio, dia_fim + 1):
        data_corte = dt.date(ano, mes, dia)
        realizado = {s: 0.0 for s in SQUADS_FINANCEIROS}
        for d in ganhos_mes:
            won_brt = to_brt(d.get("won_time"))
            if not won_brt or won_brt.date() > data_corte:
                continue
            squad = squad_do_deal(d, colaboradores, users_map)
            if squad in realizado:
                realizado[squad] += float(cf_valor(d, CF_MULTIPLICADOR) or 0)
        total_realizado = sum(realizado.values())
        linha = {"dia": dia}
        for s in SQUADS_FINANCEIROS:
            linha[SQUAD_DISPLAY[s]] = {
                "realizado": round(realizado[s], 2),
                "meta_mes": round(metas_mes[s], 2),
                "atingimento_pct": round(safe_div(realizado[s], metas_mes[s]) * 100, 2) if metas_mes[s] else None,
            }
        linha["Total"] = {
            "realizado": round(total_realizado, 2),
            "meta_mes": round(metas_mes["total"], 2),
            "atingimento_pct": round(safe_div(total_realizado, metas_mes["total"]) * 100, 2) if metas_mes["total"] else None,
        }
        por_dia.append(linha)

    return {"ano": ano, "mes": mes, "por_dia": por_dia}


def montar_historico_mql_ganhos_real(ano):
    """Só cálculo (não vai pro painel) — soma MQL e Ganhos por produto mês a mês
    (Jan a Jun) e calcula a taxa real MQL->Ganho e quantos MQL's por venda."""
    totais_mql = {}
    totais_ganhos = {}
    por_mes = []
    for mes in range(1, 7):
        dados = montar_metas_board_academy(ano, mes, dt.date(ano, mes, 15), dt.date(ano, mes, 14))
        g = dados["no_mes_geral"]
        por_mes.append({
            "mes": mes,
            "mql_por_produto": g["mql_por_produto"],
            "ganhos_por_produto": g["ganhos_por_produto"],
        })
        for p, v in g["mql_por_produto"].items():
            totais_mql[p] = totais_mql.get(p, 0) + v
        for p, v in g["ganhos_por_produto"].items():
            totais_ganhos[p] = totais_ganhos.get(p, 0) + v

    taxas = {}
    for p in sorted(set(list(totais_mql.keys()) + list(totais_ganhos.keys()))):
        mql = totais_mql.get(p, 0)
        ganhos = totais_ganhos.get(p, 0)
        taxa_pct = round(safe_div(ganhos, mql) * 100, 2) if mql else None
        mqls_por_venda = round(safe_div(mql, ganhos), 1) if ganhos else None
        taxas[p] = {"mql_total": mql, "ganhos_total": ganhos, "taxa_conversao_pct": taxa_pct, "mqls_por_venda": mqls_por_venda}

    return {"ano": ano, "por_mes": por_mes, "totais_mql_por_produto": totais_mql,
            "totais_ganhos_por_produto": totais_ganhos, "taxas_reais_por_produto": taxas}


def _papel_colaborador(nome_dono, colaboradores):
    """Classifica o dono do negócio como SDR/Closer/Outro, via cargo cadastrado na COLAB."""
    colab = colaboradores.get(norm(nome_dono), {})
    cargo = colab.get("cargo", "")
    if "closer" in cargo:
        return "Closer"
    if "sdr" in cargo:
        return "SDR"
    return "Outro"


MAX_NEGOCIOS_POR_COLABORADOR = 25  # evita payload gigante quando tem muito perdido no mês


def _colaboradores_do_motivo(deals, colaboradores, users_map):
    """Agrupa os negócios de um motivo por colaborador (dono), separado por papel (SDR/Closer/Outro)."""
    grupos = {"Closer": {}, "SDR": {}, "Outro": {}}
    for d in deals:
        nome_dono = owner_nome(d, users_map) or "Sem dono"
        papel = _papel_colaborador(nome_dono, colaboradores)
        grupos[papel].setdefault(nome_dono, []).append({
            "id": d.get("id"),
            "url": f"https://{PD_DOMAIN}/deal/{d.get('id')}",
        })
    resultado = {}
    for papel, pessoas in grupos.items():
        lista = [{"nome": nome, "quantidade": len(negocios), "negocios": negocios[:MAX_NEGOCIOS_POR_COLABORADOR]}
                 for nome, negocios in pessoas.items()]
        lista.sort(key=lambda x: x["quantidade"], reverse=True)
        resultado[papel] = lista
    return resultado


def _motivos_com_percentual(motivos_deals, total, colaboradores, users_map):
    """motivos_deals: {motivo: [negócio, ...]} -> lista com quantidade/% + quebra por colaborador."""
    itens = [
        {
            "motivo": motivo,
            "quantidade": len(deals),
            "percentual": round(safe_div(len(deals), total) * 100, 1),
            "colaboradores": _colaboradores_do_motivo(deals, colaboradores, users_map),
        }
        for motivo, deals in motivos_deals.items()
    ]
    itens.sort(key=lambda x: x["quantidade"], reverse=True)
    return itens


# ---- Auditoria de Perdas: quais motivos fazem sentido em cada etapa ----
# "Black List" é aceito em qualquer etapa (regra geral, não precisa repetir abaixo).
MOTIVOS_A_PARTIR_CONTATADO = {
    norm(m) for m in [
        "Sem interesse", "Parou de Responder", "Futuro 30", "Futuro 60", "Futuro 90",
        "Sem agenda / Próximo ano", "Sem condições financeiras", "Sem perfil",
        "Contato Inexistente", "Optou pela concorrência", "Oportunidade de emprego",
        "Cliente não reconheceu valor ou timing para avançar",
    ]
}
ETAPAS_A_PARTIR_CONTATADO = {
    norm(e) for e in [
        "Contatado", "Oportunidade", "Agendados", "No Show",
        "Validação de Reunião", "Validação da Reunião", "Negociação", "Inscrição em andamento",
    ]
}
REGRA_MOTIVOS_POR_ETAPA = {
    norm("Aplicação"): set(),                                    # não é pra ter perdido
    norm("Reabertura"): set(),                                    # suposição: mesma regra da Aplicação — confirmar
    norm("Etapa 1"): {norm("Duplicidade"), norm("Sem perfil")},
    norm("Etapa 2"): set(),                                       # não é pra ter perdido
    norm("Etapa 3"): {norm("Sem retorno")},
}
MOTIVO_BLACK_LIST = norm("Black List")


def _motivos_permitidos_na_etapa(etapa_norm):
    """None = etapa não mapeada (não auditamos); set() = nenhum motivo aceito nessa etapa."""
    if etapa_norm in ETAPAS_A_PARTIR_CONTATADO:
        return MOTIVOS_A_PARTIR_CONTATADO
    if etapa_norm in REGRA_MOTIVOS_POR_ETAPA:
        return REGRA_MOTIVOS_POR_ETAPA[etapa_norm]
    return None


def analisar_perdidos_periodo(data_inicio, data_fim, colaboradores, users_map, responsavel_norm=None, squad_unico=None):
    """Versão da análise 'Perdidos — Geral' com período customizado (não o mês inteiro)
    e filtro opcional por responsável (dono do negócio). squad_unico restringe a busca
    a só um funil (usado pra papéis não-admin, que só podem ver o próprio funil)."""
    desde_iso = f"{data_inicio.isoformat()}T00:00:00Z"
    ate_iso = f"{data_fim.isoformat()}T23:59:59Z"

    motivos_geral = {}
    total = 0
    squads_para_buscar = [squad_unico] if squad_unico else SQUADS_PRODUTOS
    for squad_interno in squads_para_buscar:
        pipeline_id = _pipeline_id_do_squad(squad_interno)
        if pipeline_id is None:
            continue
        perdidos = buscar_deals_por_pipeline(pipeline_id, "lost", desde_iso, ate_iso)
        for d in perdidos:
            lost_brt = to_brt(d.get("lost_time"))
            if not lost_brt:
                continue
            data_perda = lost_brt.date()
            if not (data_inicio <= data_perda <= data_fim):
                continue
            if responsavel_norm and norm(owner_nome(d, users_map)) != responsavel_norm:
                continue
            total += 1
            motivo = (d.get("lost_reason") or "").strip() or "Sem motivo"
            motivos_geral.setdefault(motivo, []).append(d)

    return {
        "total": total,
        "motivos": _motivos_com_percentual(motivos_geral, total, colaboradores, users_map),
        "periodo_inicio": data_inicio.isoformat(),
        "periodo_fim": data_fim.isoformat(),
    }


def analisar_perdidos(ano, mes, hoje, colaboradores, users_map):
    """Perdidos do dia/mês + motivo de perda (quantidade, % e quebra por colaborador SDR/Closer),
    geral e por funil (Olympus/Elite/Sniper), + Auditoria de Perdas (motivos fora da regra por etapa)."""
    inicio_mes = dt.date(ano, mes, 1)
    fim_mes = dt.date(ano + 1, 1, 1) - dt.timedelta(days=1) if mes == 12 else dt.date(ano, mes + 1, 1) - dt.timedelta(days=1)
    desde_iso = f"{inicio_mes.isoformat()}T00:00:00Z"
    ate_iso = f"{fim_mes.isoformat()}T23:59:59Z"

    por_funil = {}
    motivos_geral = {}  # motivo -> [negócio, ...]
    total_mes_geral = 0
    total_hoje_geral = 0
    auditoria_geral = {}  # etapa -> {motivo: [negócio, ...]}
    responsaveis_vistos = set()

    for squad_interno in SQUADS_PRODUTOS:
        pipeline_id = _pipeline_id_do_squad(squad_interno)
        motivos_funil = {}
        total_mes = 0
        total_hoje = 0
        auditoria_funil = {}
        if pipeline_id is not None:
            perdidos = buscar_deals_por_pipeline(pipeline_id, "lost", desde_iso, ate_iso)
            stages_map = buscar_stages_pipeline(pipeline_id)
            for d in perdidos:
                lost_brt = to_brt(d.get("lost_time"))
                if not lost_brt or (lost_brt.year, lost_brt.month) != (ano, mes):
                    continue
                total_mes += 1
                if lost_brt.date() == hoje:
                    total_hoje += 1
                responsaveis_vistos.add(owner_nome(d, users_map) or "Sem dono")
                motivo = (d.get("lost_reason") or "").strip() or "Sem motivo"
                motivos_funil.setdefault(motivo, []).append(d)
                motivos_geral.setdefault(motivo, []).append(d)

                # Auditoria: o motivo faz sentido pra etapa em que o negócio estava?
                motivo_norm = norm(motivo)
                if motivo_norm == MOTIVO_BLACK_LIST:
                    continue  # aceito em qualquer etapa
                etapa_info = stages_map.get(d.get("stage_id"))
                if not etapa_info:
                    continue
                permitidos = _motivos_permitidos_na_etapa(etapa_info["norm"])
                if permitidos is None:
                    continue  # etapa não mapeada, não auditamos
                if motivo_norm not in permitidos:
                    etapa_nome = etapa_info["nome"]
                    auditoria_funil.setdefault(etapa_nome, {}).setdefault(motivo, []).append(d)
                    auditoria_geral.setdefault(etapa_nome, {}).setdefault(motivo, []).append(d)

        por_funil[squad_interno] = {
            "total_mes": total_mes,
            "total_hoje": total_hoje,
            "motivos": _motivos_com_percentual(motivos_funil, total_mes, colaboradores, users_map),
            "auditoria": [
                {
                    "etapa": etapa,
                    "total": sum(len(v) for v in m.values()),
                    "motivos": _motivos_com_percentual(m, sum(len(v) for v in m.values()), colaboradores, users_map),
                }
                for etapa, m in sorted(auditoria_funil.items(), key=lambda kv: sum(len(v) for v in kv[1].values()), reverse=True)
            ],
        }
        total_mes_geral += total_mes
        total_hoje_geral += total_hoje

    return {
        "geral": {
            "total_mes": total_mes_geral,
            "total_hoje": total_hoje_geral,
            "motivos": _motivos_com_percentual(motivos_geral, total_mes_geral, colaboradores, users_map),
            "auditoria": [
                {
                    "etapa": etapa,
                    "total": sum(len(v) for v in m.values()),
                    "motivos": _motivos_com_percentual(m, sum(len(v) for v in m.values()), colaboradores, users_map),
                }
                for etapa, m in sorted(auditoria_geral.items(), key=lambda kv: sum(len(v) for v in kv[1].values()), reverse=True)
            ],
        },
        "por_funil": {SQUAD_DISPLAY[s]: por_funil[s] for s in SQUADS_PRODUTOS},
        "responsaveis": sorted(responsaveis_vistos),
    }


def _item_deal(d):
    return {
        "id": d.get("id"),
        "titulo": d.get("title") or f"Negócio {d.get('id')}",
        "valor": float(d.get("value") or 0),
        "url": f"https://{PD_DOMAIN}/deal/{d.get('id')}",
    }


CATEGORIAS_PRODUTO = ["abertos", "perdidos_mes", "perdidos_hoje", "ganhos_mes", "ganhos_hoje", "novos_hoje", "novos_ontem", "novos_mes", "reaplicacoes_hoje", "reaplicacoes_mes"]


def _parse_data_campo(valor):
    """Converte o valor bruto de um campo de data (ISO ou BR) num date. None se não der pra ler."""
    if not valor:
        return None
    texto = str(valor)[:10]
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return dt.datetime.strptime(texto, fmt).date()
        except ValueError:
            continue
    return None


def produtos_em_aberto_por_squad(ano, mes, hoje, ontem, retornar_detalhes=False):
    """Por produto e squad (Olympus/Elite/Sniper/Navigator), buscando direto por pipeline (funil):
    contagens de Abertos, Perdidos (mês/hoje) e Ganhos (mês/hoje) — Sniper e Navigator não têm Ganhos.
    Tudo classificado pela utm_campaign, sem depender da COLAB/cargo.
    Se retornar_detalhes=True, mantém a lista completa (com link) em vez de só o número."""
    detalhes = {s: {p: {c: [] for c in CATEGORIAS_PRODUTO} for p in PRODUTOS_TODOS}
                for s in SQUADS_PRODUTOS_ABA}

    inicio_mes = dt.date(ano, mes, 1)
    fim_mes = dt.date(ano + 1, 1, 1) - dt.timedelta(days=1) if mes == 12 else dt.date(ano, mes + 1, 1) - dt.timedelta(days=1)
    desde_iso = f"{inicio_mes.isoformat()}T00:00:00Z"
    ate_iso = f"{fim_mes.isoformat()}T23:59:59Z"

    for squad_interno in SQUADS_PRODUTOS_ABA:
        pipeline_id = _pipeline_id_do_squad(squad_interno)
        tem_ganhos = squad_interno != "sniper"
        if pipeline_id is None:
            continue

        abertos = buscar_deals_por_pipeline(pipeline_id, "open")
        for d in abertos:
            produto = classificar_produto(d) or "Não classificado"
            detalhes[squad_interno][produto]["abertos"].append(_item_deal(d))

        # limitado ao mês selecionado (update_time) — senão baixaria o histórico inteiro do funil
        ganhos = []
        if tem_ganhos:
            ganhos = buscar_deals_por_pipeline(pipeline_id, "won", desde_iso, ate_iso)
            for d in ganhos:
                won_brt = to_brt(d.get("won_time"))
                if not won_brt or (won_brt.year, won_brt.month) != (ano, mes):
                    continue
                produto = classificar_produto_ganho(d) or "Não classificado"
                item = _item_deal(d)
                detalhes[squad_interno][produto]["ganhos_mes"].append(item)
                if won_brt.date() == hoje:
                    detalhes[squad_interno][produto]["ganhos_hoje"].append(item)

        perdidos = buscar_deals_por_pipeline(pipeline_id, "lost", desde_iso, ate_iso)
        for d in perdidos:
            lost_brt = to_brt(d.get("lost_time"))
            if not lost_brt or (lost_brt.year, lost_brt.month) != (ano, mes):
                continue
            produto = classificar_produto(d) or "Não classificado"
            item = _item_deal(d)
            detalhes[squad_interno][produto]["perdidos_mes"].append(item)
            if lost_brt.date() == hoje:
                detalhes[squad_interno][produto]["perdidos_hoje"].append(item)

        # Novos Leads (hoje/ontem/mês) — só faz sentido pra Sniper e Navigator (pedido do usuário).
        # Pela data de criação do negócio ("Negócio criado em" / add_time).
        if squad_interno in ("sniper", "navigator"):
            vistos_novos = set()
            for d in abertos + ganhos + perdidos:
                did = d.get("id")
                if did in vistos_novos:
                    continue
                vistos_novos.add(did)
                produto = classificar_produto(d) or "Não classificado"
                add_brt = to_brt(d.get("add_time"))
                if add_brt and add_brt.date() == hoje:
                    detalhes[squad_interno][produto]["novos_hoje"].append(_item_deal(d))
                if add_brt and add_brt.date() == ontem:
                    detalhes[squad_interno][produto]["novos_ontem"].append(_item_deal(d))
                if add_brt and (add_brt.year, add_brt.month) == (ano, mes):
                    detalhes[squad_interno][produto]["novos_mes"].append(_item_deal(d))

                # Reaplicação: "Data da última aplicação" != "Negócio criado em" (datas diferentes)
                data_aplicacao = _parse_data_campo(cf_valor(d, CF_DATA_ULTIMA_APLICACAO))
                if data_aplicacao and add_brt and data_aplicacao != add_brt.date():
                    if data_aplicacao == hoje:
                        detalhes[squad_interno][produto]["reaplicacoes_hoje"].append(_item_deal(d))
                    if (data_aplicacao.year, data_aplicacao.month) == (ano, mes):
                        detalhes[squad_interno][produto]["reaplicacoes_mes"].append(_item_deal(d))

    if retornar_detalhes:
        return detalhes

    # números pra todo mundo, mas a lista (id + link) fica pros GANHOS e NOVOS LEADS
    resultado = {}
    for s in SQUADS_PRODUTOS_ABA:
        resultado[s] = {}
        for p in PRODUTOS_TODOS:
            resultado[s][p] = {c: len(detalhes[s][p][c]) for c in CATEGORIAS_PRODUTO}
            for chave in ("ganhos_mes", "ganhos_hoje", "novos_hoje", "novos_ontem", "novos_mes", "reaplicacoes_hoje", "reaplicacoes_mes"):
                resultado[s][p][f"{chave}_lista"] = [
                    {"id": item["id"], "url": item["url"]} for item in detalhes[s][p][chave]
                ]
    return resultado


def montar_resposta_perdidos_periodo(mes_param, ano_param, perdidos_inicio, perdidos_fim, perdidos_responsavel, papel):
    """Resposta leve — só pro filtro de período/responsável da seção 'Perdidos — Geral'.
    Não refaz o painel inteiro (evita timeout): busca só colaboradores + usuários.
    Exige login (papel != None); papéis não-admin ficam restritos ao próprio funil."""
    if papel is None:
        return {"erro": "Login necessário pra ver a aba Perdidos."}
    hoje = hoje_brt()
    ano, mes = ano_param or hoje.year, mes_param or hoje.month
    colaboradores = carregar_colaboradores(mes, ano)
    users_map = pd_users()
    data_ini = dt.date.fromisoformat(perdidos_inicio)
    data_fim_p = dt.date.fromisoformat(perdidos_fim)
    resp_norm = norm(perdidos_responsavel) if perdidos_responsavel else None
    squad_unico = None if papel == "admin" else PAPEL_PARA_SQUAD_INTERNO.get(papel)
    if papel != "admin" and squad_unico is None:
        return {"erro": "Papel de acesso inválido."}
    return {
        "perdidos_periodo": analisar_perdidos_periodo(data_ini, data_fim_p, colaboradores, users_map, resp_norm, squad_unico)
    }




def montar_painel(ano_param=None, mes_param=None, papel=None):
    hoje = hoje_brt()
    ano, mes = ano_param or hoje.year, mes_param or hoje.month
    e_mes_atual = (ano, mes) == (hoje.year, hoje.month)

    feriados = carregar_feriados()
    colaboradores = carregar_colaboradores(mes, ano)
    metas = carregar_metas(mes, ano)
    du = calcular_dias_uteis(ano, mes, feriados)
    # Se não sobra nenhum dia útil DEPOIS de hoje, mas hoje ainda é dia útil,
    # considera hoje como "1 dia" pra não zerar o Meta/Dia 100% no último dia do mês.
    if e_mes_atual:
        # dia atual conta como "1 dia" quando não sobra nenhum depois dele (último dia útil do mês)
        dias_restantes_p100 = du["restantes"] if du["restantes"] > 0 else (1 if eh_dia_util(hoje, feriados) else 0)
    else:
        # mês já fechado: não existe "dias restantes", então usa o total de dias úteis do mês
        dias_restantes_p100 = du["total"]

    users_map = pd_users()
    deals_ganhos = buscar_deals_ganhos(ano, mes, users_map)

    ontem = dia_util_anterior(hoje, feriados)
    prox_dia_util = proximo_dia_util(hoje, feriados)

    # ---- Financeiro: Olympus (mgm) — M1/M2 (CRN) ----
    squads_fin = {s: {"bruto": 0.0, "multi": 0.0, "ontem": 0.0, "hoje": 0.0, "ontem_bruto": 0.0, "hoje_bruto": 0.0} for s in SQUADS_FINANCEIROS}
    ganhos_por_sq = _ganhos_por_squad(deals_ganhos, colaboradores, users_map)
    for squad, lista in ganhos_por_sq.items():
        if squad not in squads_fin:
            continue
        squads_fin[squad]["bruto"] = _m2.bruto(lista)
        squads_fin[squad]["multi"] = _m2.multiplicador(lista)
        for chave, dia in (("ontem", ontem), ("hoje", hoje)):
            do_dia = [d for d in lista if _fuso.data_brasil(d.get("won_time")) == dia]
            squads_fin[squad][chave] = _m2.multiplicador(do_dia)
            squads_fin[squad][chave + "_bruto"] = _m2.bruto(do_dia)

    # ---- Previsto (forecast) hoje/ontem — F1/F2 (CRN), só faz sentido no mês atual ----
    if e_mes_atual:
        forecast_abertos = buscar_forecast_deals()
        pool_previsto = forecast_abertos
        previsto_hoje = valor_previsto_por_squad(pool_previsto, colaboradores, users_map, hoje)
        previsto_ontem = valor_previsto_por_squad(pool_previsto, colaboradores, users_map, ontem)
        em_aberto_hoje = valor_abertos_por_squad(pool_previsto, colaboradores, users_map, hoje)
    else:
        pool_previsto = []
        forecast_abertos = []
        previsto_hoje = {s: {"p20": 0.0, "p50": 0.0, "p70": 0.0, "media": 0.0} for s in SQUADS_FINANCEIROS}
        previsto_ontem = {s: {"p20": 0.0, "p50": 0.0, "p70": 0.0, "media": 0.0} for s in SQUADS_FINANCEIROS}
        em_aberto_hoje = {s: 0.0 for s in SQUADS_FINANCEIROS}

    def meta_squad(squad_interno):
        return _meta_financeira_squad(squad_interno, colaboradores, metas)

    # ---- Comparativo "mesmo dia do mês passado" — só no mês atual ----
    squads_mes_passado = {s: 0.0 for s in SQUADS_FINANCEIROS}
    squads_mes_passado_bruto = {s: 0.0 for s in SQUADS_FINANCEIROS}
    squads_mes_passado_atingimento = {s: 0.0 for s in SQUADS_FINANCEIROS}
    mes_passado_rotulo = None
    colaboradores_mp, metas_mp, ano_passado, mes_passado, dia_alvo_mes_passado = None, None, None, None, None
    if e_mes_atual:
        ano_passado = ano - 1 if mes == 1 else ano
        mes_passado = 12 if mes == 1 else mes - 1
        ultimo_dia_mes_passado = calendar.monthrange(ano_passado, mes_passado)[1]
        dia_alvo_mes_passado = min(hoje.day, ultimo_dia_mes_passado)
        data_corte_mes_passado = dt.date(ano_passado, mes_passado, dia_alvo_mes_passado)
        mes_passado_rotulo = data_corte_mes_passado.strftime("%d/%m")

        # Time e meta de QUEM ESTAVA no squad naquele mês (Colab/Metas do mês, R1/R2)
        colaboradores_mp = carregar_colaboradores(mes_passado, ano_passado)
        metas_mp = carregar_metas(mes_passado, ano_passado)

        def meta_squad_mp(squad_interno):
            return _meta_financeira_squad(squad_interno, colaboradores_mp, metas_mp)

        deals_ganhos_mes_passado = [
            d for d in buscar_deals_ganhos(ano_passado, mes_passado, users_map)
            if (_fuso.data_brasil(d.get("won_time")) or data_corte_mes_passado) <= data_corte_mes_passado
        ]
        for squad, lista in _ganhos_por_squad(deals_ganhos_mes_passado, colaboradores_mp, users_map).items():
            if squad in squads_mes_passado:
                squads_mes_passado[squad] = _m2.multiplicador(lista)
                squads_mes_passado_bruto[squad] = _m2.bruto(lista)
        for s in SQUADS_FINANCEIROS:
            meta_mp_s = meta_squad_mp(s)
            squads_mes_passado_atingimento[s] = round(safe_div(squads_mes_passado[s], meta_mp_s) * 100, 2) if meta_mp_s else 0.0

    # MTD (T1): meta até D-1 (dias úteis passados = A2), realizado até agora.
    ritmo_100 = safe_div(du["passados"], du["total"])

    # Prazo do gap intermediário (40%) — parâmetro do painel (não é regra da CRN)
    dias_no_mes_calendario = calendar.monthrange(ano, mes)[1]
    dia_metade = math.ceil(dias_no_mes_calendario / 2)
    prazo_gap_intermediario = dt.date(ano, mes, dia_metade)

    dias_uteis_primeira_metade = 0
    d = dt.date(ano, mes, 1)
    while d <= prazo_gap_intermediario:
        if eh_dia_util(d, feriados):
            dias_uteis_primeira_metade += 1
        d += dt.timedelta(days=1)

    if e_mes_atual:
        restantes_prazo = 0
        d = hoje + dt.timedelta(days=1)
        while d <= prazo_gap_intermediario:
            if eh_dia_util(d, feriados):
                restantes_prazo += 1
            d += dt.timedelta(days=1)
        dias_restantes_p40 = restantes_prazo if restantes_prazo > 0 else (1 if eh_dia_util(hoje, feriados) else 0)
    else:
        restantes_prazo = 0
        dias_restantes_p40 = dias_uteis_primeira_metade

    hoje_dentro_prazo = 1 if (e_mes_atual and hoje <= prazo_gap_intermediario and eh_dia_util(hoje, feriados)) else 0
    ritmo_prazo = safe_div(du["passados"], du["passados"] + hoje_dentro_prazo + restantes_prazo)

    resultado = {
        "squads": {}, "geradoEm": dt.datetime.now(dt.timezone.utc).isoformat(),
        "mes": mes, "ano": ano, "e_mes_atual": e_mes_atual,
        "crn_versao": crn.versao(),
    }

    rotulo_financeiro = dict(SQUAD_DISPLAY)

    for squad_interno in SQUADS_FINANCEIROS:
        meta_mes = meta_squad(squad_interno)
        meta_dia = safe_div(meta_mes, du["total"])
        bruto = squads_fin[squad_interno]["bruto"]
        multi = squads_fin[squad_interno]["multi"]
        onde_100 = _t1.mtd(multi, meta_mes, du["total"], du["passados"])["meta_mtd"]
        onde_40 = (PCT_GAP_INTERMEDIARIO * meta_mes) * ritmo_prazo
        gap_100 = max(0.0, meta_mes - multi)
        gap_40 = max(0.0, (PCT_GAP_INTERMEDIARIO * meta_mes) - multi)
        meta_dia_40 = safe_div(gap_40, dias_restantes_p40)
        meta_dia_100 = safe_div(gap_100, dias_restantes_p100)
        gap_100_bruto = max(0.0, meta_mes - bruto)
        meta_dia_100_bruto = safe_div(gap_100_bruto, dias_restantes_p100)

        resultado["squads"][rotulo_financeiro[squad_interno]] = {
            "meta_mes": round(meta_mes, 2),
            "meta_dia": round(meta_dia, 2),
            "realizado_bruto": round(bruto, 2),
            "realizado_multiplicador": round(multi, 2),
            "onde_deveria_100": round(onde_100, 2),
            "onde_deveria_40": round(onde_40, 2),
            "atingimento": round(safe_div(multi, meta_mes) * 100, 2),
            "gap_100": round(gap_100, 2),
            "gap_40": round(gap_40, 2),
            "meta_dia_40": round(meta_dia_40, 2),
            "meta_dia_100": round(meta_dia_100, 2),
            "gap_100_bruto": round(gap_100_bruto, 2),
            "meta_dia_100_bruto": round(meta_dia_100_bruto, 2),
            "ontem": round(squads_fin[squad_interno]["ontem"], 2),
            "hoje": round(squads_fin[squad_interno]["hoje"], 2),
            "ontem_bruto": round(squads_fin[squad_interno]["ontem_bruto"], 2),
            "hoje_bruto": round(squads_fin[squad_interno]["hoje_bruto"], 2),
            "previsto_hoje_20": round(previsto_hoje.get(squad_interno, {}).get("p20", 0.0), 2),
            "previsto_hoje_50": round(previsto_hoje.get(squad_interno, {}).get("p50", 0.0), 2),
            "previsto_hoje_70": round(previsto_hoje.get(squad_interno, {}).get("p70", 0.0), 2),
            "previsto_hoje_media": round(previsto_hoje.get(squad_interno, {}).get("media", 0.0), 2),
            "previsto_ontem_20": round(previsto_ontem.get(squad_interno, {}).get("p20", 0.0), 2),
            "previsto_ontem_50": round(previsto_ontem.get(squad_interno, {}).get("p50", 0.0), 2),
            "previsto_ontem_70": round(previsto_ontem.get(squad_interno, {}).get("p70", 0.0), 2),
            "previsto_ontem_media": round(previsto_ontem.get(squad_interno, {}).get("media", 0.0), 2),
            "em_aberto_hoje": round(em_aberto_hoje.get(squad_interno, 0.0), 2),
            "mes_passado_mesmo_dia": round(squads_mes_passado.get(squad_interno, 0.0), 2),
            "mes_passado_bruto": round(squads_mes_passado_bruto.get(squad_interno, 0.0), 2),
            "mes_passado_atingimento": squads_mes_passado_atingimento.get(squad_interno, 0.0),
        }

    total_meta_mes = sum(resultado["squads"][rotulo_financeiro[s]]["meta_mes"] for s in SQUADS_FINANCEIROS)
    total_bruto = sum(resultado["squads"][rotulo_financeiro[s]]["realizado_bruto"] for s in SQUADS_FINANCEIROS)
    total_multi = sum(resultado["squads"][rotulo_financeiro[s]]["realizado_multiplicador"] for s in SQUADS_FINANCEIROS)
    total_ontem = sum(resultado["squads"][rotulo_financeiro[s]]["ontem"] for s in SQUADS_FINANCEIROS)
    total_hoje = sum(resultado["squads"][rotulo_financeiro[s]]["hoje"] for s in SQUADS_FINANCEIROS)
    total_ontem_bruto = sum(resultado["squads"][rotulo_financeiro[s]]["ontem_bruto"] for s in SQUADS_FINANCEIROS)
    total_hoje_bruto = sum(resultado["squads"][rotulo_financeiro[s]]["hoje_bruto"] for s in SQUADS_FINANCEIROS)
    campos_previsto = ["previsto_hoje_20", "previsto_hoje_50", "previsto_hoje_70", "previsto_hoje_media",
                        "previsto_ontem_20", "previsto_ontem_50", "previsto_ontem_70", "previsto_ontem_media",
                        "em_aberto_hoje", "mes_passado_mesmo_dia", "mes_passado_bruto"]
    totais_previsto = {
        campo: sum(resultado["squads"][rotulo_financeiro[s]][campo] for s in SQUADS_FINANCEIROS)
        for campo in campos_previsto
    }
    total_gap_40 = max(0.0, (PCT_GAP_INTERMEDIARIO * total_meta_mes) - total_multi)
    resultado["squads"]["Total"] = {
        "meta_mes": round(total_meta_mes, 2),
        "meta_dia": round(safe_div(total_meta_mes, du["total"]), 2),
        "realizado_bruto": round(total_bruto, 2),
        "realizado_multiplicador": round(total_multi, 2),
        "onde_deveria_100": round(_t1.mtd(total_multi, total_meta_mes, du["total"], du["passados"])["meta_mtd"], 2),
        "onde_deveria_40": round((PCT_GAP_INTERMEDIARIO * total_meta_mes) * ritmo_prazo, 2),
        "atingimento": round(safe_div(total_multi, total_meta_mes) * 100, 2),
        "gap_100": round(max(0.0, total_meta_mes - total_multi), 2),
        "gap_40": round(total_gap_40, 2),
        "meta_dia_40": round(safe_div(total_gap_40, dias_restantes_p40), 2),
        "meta_dia_100": round(safe_div(max(0.0, total_meta_mes - total_multi), dias_restantes_p100), 2),
        "gap_100_bruto": round(max(0.0, total_meta_mes - total_bruto), 2),
        "meta_dia_100_bruto": round(safe_div(max(0.0, total_meta_mes - total_bruto), dias_restantes_p100), 2),
        "ontem": round(total_ontem, 2),
        "hoje": round(total_hoje, 2),
        "ontem_bruto": round(total_ontem_bruto, 2),
        "hoje_bruto": round(total_hoje_bruto, 2),
        **{k: round(v, 2) for k, v in totais_previsto.items()},
        "mes_passado_atingimento": round(
            safe_div(sum(squads_mes_passado.values()), sum(meta_squad_mp(s) for s in SQUADS_FINANCEIROS)) * 100, 2
        ) if e_mes_atual else 0.0,
    }
    resultado["mes_passado_rotulo"] = mes_passado_rotulo

    # ---- Sniper: reuniões — M6 (CRN) ----
    colab_lista = next(iter(colaboradores.values()))["_colab"] if colaboradores else []
    uid_por_nome = {norm(n): i for i, n in users_map.items()}
    sdr_nomes = {n for n, m in metas.items() if m.get("papel") == "sdr"}
    sdr_ids = {uid_por_nome[n] for n in sdr_nomes if n in uid_por_nome}
    lideranca_ids = {uid_por_nome[p["nome_norm"]] for p in colab_lista
                     if p["nome_norm"] in uid_por_nome and _h1.nivel(p, colab_lista) in ("head", "team_leader", "gerente")}
    lider_agenda_ids = _m6.lideranca_que_agenda(colab_lista, uid_por_nome)
    sdr_sniper_nomes = {n for n in sdr_nomes if colaboradores.get(n, {}).get("subarea") == SQUAD_SDR}
    ids_excluidos = {uid_por_nome[n] for n in EXCLUIDOS_PAINEL if n in uid_por_nome}
    ids_sniper = (({uid_por_nome[n] for n in sdr_sniper_nomes if n in uid_por_nome}) | set(lider_agenda_ids)) - ids_excluidos
    nome_por_id = {i: norm(n) for i, n in users_map.items()}

    ini_mes = dt.date(ano, mes, 1)
    fim_mes = dt.date(ano, mes, calendar.monthrange(ano, mes)[1])
    dados_m6 = _m6.carregar_periodo(ini_mes, fim_mes)
    kw_m6 = dict(dono_do_deal=dados_m6["dono_do_deal"], sdr_ids=sdr_ids, lideranca_ids=lideranca_ids,
                 lideranca_que_agenda_ids=lider_agenda_ids)
    por_resp = _m6.reunioes_validadas(dados_m6["atividades"], ini_mes, fim_mes,
                                      deals_validados=dados_m6["deals_validados"], **kw_m6)
    validadas_sniper = {r: v for r, v in por_resp.items() if r in ids_sniper}
    validadas_total = sum(len(v) for v in validadas_sniper.values())
    d_ontem_util = ontem
    validadas_hoje = sum(1 for v in validadas_sniper.values() for a in v if _fuso.atividade_dia(a) == hoje)
    validadas_dia_anterior_util = sum(1 for v in validadas_sniper.values() for a in v if _fuso.atividade_dia(a) == d_ontem_util)

    # Previsto (FR1): agendadas para o dia, feitas ou não, do SDR
    previsto_hoje_reu = previsto_ontem_reu = 0
    if e_mes_atual:
        ativ_prev = _pd.atividades_do_periodo(min(d_ontem_util, hoje), hoje, tipos=[crn.base.config.tipo_atividade("reuniao")])
        previsto_hoje_reu = sum(len(v) for r, v in _fr1.previstas(ativ_prev, hoje, ids_sniper).items())
        previsto_ontem_reu = sum(len(v) for r, v in _fr1.previstas(ativ_prev, d_ontem_util, ids_sniper).items())

    # Conferência por SDR (universo da M6 antes da validação)
    universo = _m6.reunioes_com_status(dados_m6["atividades"], ini_mes, fim_mes,
                                       status_do_deal=dados_m6["status_do_deal"], **kw_m6)
    nao_validadas_por_resp = {}
    for item in universo:
        if not item["validada"]:
            nao_validadas_por_resp[item["responsavel"]] = nao_validadas_por_resp.get(item["responsavel"], 0) + 1

    meta_reunioes = sum(metas[n]["meta_reu"] for n in sdr_sniper_nomes)
    meta_dia_reu = safe_div(meta_reunioes, du["total"])
    onde_100_reu = _t1.mtd(validadas_total, meta_reunioes, du["total"], du["passados"])["meta_mtd"]
    onde_40_reu = (PCT_GAP_INTERMEDIARIO * meta_reunioes) * ritmo_prazo
    gap_100_reu = max(0.0, meta_reunioes - validadas_total)

    # ---- Sniper: mesmo comparativo do mês passado (M6 com Colab/Metas daquele mês) ----
    validadas_mes_passado = 0
    meta_reunioes_mes_passado = 0.0
    if e_mes_atual and colaboradores_mp is not None:
        colab_mp = next(iter(colaboradores_mp.values()))["_colab"] if colaboradores_mp else []
        ini_mp = dt.date(ano_passado, mes_passado, 1)
        corte_mp = dt.date(ano_passado, mes_passado, dia_alvo_mes_passado)
        sdr_nomes_mp = {n for n, m in metas_mp.items() if m.get("papel") == "sdr"}
        sdr_ids_mp = {uid_por_nome[n] for n in sdr_nomes_mp if n in uid_por_nome}
        lid_ids_mp = {uid_por_nome[p["nome_norm"]] for p in colab_mp
                      if p["nome_norm"] in uid_por_nome and _h1.nivel(p, colab_mp) in ("head", "team_leader", "gerente")}
        lid_ag_mp = _m6.lideranca_que_agenda(colab_mp, uid_por_nome)
        sdr_sniper_mp = {n for n in sdr_nomes_mp if colaboradores_mp.get(n, {}).get("subarea") == SQUAD_SDR}
        ids_sniper_mp = ({uid_por_nome[n] for n in sdr_sniper_mp if n in uid_por_nome} | set(lid_ag_mp)) - ids_excluidos
        dados_mp = _m6.carregar_periodo(ini_mp, corte_mp)
        por_resp_mp = _m6.reunioes_validadas(dados_mp["atividades"], ini_mp, corte_mp,
                                             deals_validados=dados_mp["deals_validados"],
                                             dono_do_deal=dados_mp["dono_do_deal"], sdr_ids=sdr_ids_mp,
                                             lideranca_ids=lid_ids_mp, lideranca_que_agenda_ids=lid_ag_mp)
        validadas_mes_passado = sum(len(v) for r, v in por_resp_mp.items() if r in ids_sniper_mp)
        meta_reunioes_mes_passado = sum(metas_mp[n]["meta_reu"] for n in sdr_sniper_mp)
    atingimento_sniper_mes_passado = round(safe_div(validadas_mes_passado, meta_reunioes_mes_passado) * 100, 2) if meta_reunioes_mes_passado else 0.0

    gap_40_reu = max(0.0, (PCT_GAP_INTERMEDIARIO * meta_reunioes) - validadas_total)
    meta_dia_40_reu = safe_div(gap_40_reu, dias_restantes_p40)
    meta_dia_100_reu = safe_div(gap_100_reu, dias_restantes_p100)

    # ---- Sniper: financeiro — M4/M5 (CRN): ganhos do campo Qualificador (G7), meta só dos SDRs (R2) ----
    opcoes_q = _pd.opcoes_campo(_campo("qualificador"))
    ganhos_sdr = _m4.ganhos_por_sdr(deals_ganhos, opcoes_q)
    meta_fin_por_pessoa = {n: metas[n]["meta_fin"] for n in sdr_sniper_nomes}
    meta_fin_sniper = sum(meta_fin_por_pessoa.values())
    realizado_fin_por_qualificador = {q: _m2.multiplicador(lista) for q, lista in ganhos_sdr.items() if q in sdr_sniper_nomes}
    realizado_fin_sniper = sum(realizado_fin_por_qualificador.values())
    resultado["debug_sniper_financeiro"] = {
        "meta_por_pessoa": meta_fin_por_pessoa,
        "realizado_multi_por_qualificador": {k: round(v, 2) for k, v in realizado_fin_por_qualificador.items()},
    }
    pct_reu = _m6.pct_reunioes(validadas_total, meta_reunioes)
    pct_fin = _m3.atingimento(realizado_fin_sniper, meta_fin_sniper)
    atingimento_financeiro = pct_fin or 0.0
    atingimento_final = _m7.pct_final(pct_reu, pct_fin, ini_mes) or 0.0

    resultado["squads"]["Sniper"] = {
        "meta_financeira": round(meta_fin_sniper, 2),
        "realizado_financeiro": round(realizado_fin_sniper, 2),
        "atingimento_financeiro": round(atingimento_financeiro, 2),
        "atingimento_final": round(atingimento_final, 2),
        "meta_mes_reunioes": meta_reunioes,
        "meta_dia_reunioes": round(meta_dia_reu, 2),
        "realizado_reunioes": validadas_total,
        # Conferência por SDR (M6): contadas e reuniões fora por "Reunião Validada? = Não"
        "reunioes_por_sdr": sorted(
            [{
                "nome": (colaboradores.get(nome_por_id.get(r, ""), {}).get("nome_exibicao") or nome_por_id.get(r, str(r)).title()),
                "meta": (metas.get(nome_por_id.get(r, "")) or {}).get("meta_reu", 0),
                "contadas": len(validadas_sniper.get(r, [])),
                "deal_fora_da_lista": nao_validadas_por_resp.get(r, 0),
                "dono_igual": 0,
                "nao_concluida": 0,
            } for r in ids_sniper],
            key=lambda x: -x["contadas"]),
        "realizado_reunioes_ate_ontem": validadas_total - validadas_hoje,
        "mes_passado_reunioes": validadas_mes_passado,
        "mes_passado_atingimento": atingimento_sniper_mes_passado,
        "onde_deveria_100": round(onde_100_reu, 2),
        "onde_deveria_40": round(onde_40_reu, 2),
        "atingimento": round(safe_div(validadas_total, meta_reunioes) * 100, 2),
        "gap_100": round(gap_100_reu, 2),
        "gap_40": round(gap_40_reu, 2),
        "meta_dia_40": round(meta_dia_40_reu, 2),
        "meta_dia_100": round(meta_dia_100_reu, 2),
        "reunioes_hoje": validadas_hoje,
        "dia_anterior_reunioes": validadas_dia_anterior_util,
        "previsto_hoje": previsto_hoje_reu,
        "previsto_ontem": previsto_ontem_reu,
        "gap_hoje": max(0, previsto_hoje_reu - validadas_hoje),
        "gap_ontem": max(0, previsto_ontem_reu - validadas_dia_anterior_util),
        "novos_leads_hoje": 0,   # preenchido mais abaixo, reaproveitando o cálculo de Produtos
        "novos_leads_ontem": 0,
    }

    resultado["premissas"] = {
        "dias_uteis_total": du["total"],
        "dias_uteis_passados": du["passados"],
        "dias_uteis_restantes": du["restantes"],
        "dias_restantes_p100": dias_restantes_p100,
        "prazo_gap_intermediario": prazo_gap_intermediario.isoformat(),
        "percentual_gap_intermediario": PCT_GAP_INTERMEDIARIO,
        "proximo_dia_util": prox_dia_util.isoformat(),
        "ritmo_100_pct": round(ritmo_100 * 100, 1),
        "ritmo_40_pct": round(ritmo_prazo * 100, 1),
    }
    if e_mes_atual:
        produtos_detalhes = produtos_em_aberto_por_squad(ano, mes, hoje, ontem)
    else:
        produtos_detalhes = {
            s: {p: {**{c: 0 for c in CATEGORIAS_PRODUTO}, "ganhos_mes_lista": [], "ganhos_hoje_lista": [],
                     "novos_hoje_lista": [], "novos_ontem_lista": [], "novos_mes_lista": [],
                     "reaplicacoes_hoje_lista": [], "reaplicacoes_mes_lista": []}
                for p in PRODUTOS_TODOS}
            for s in SQUADS_PRODUTOS_ABA
        }
    resultado["produtos_em_aberto_detalhes"] = {
        SQUAD_DISPLAY[s]: produtos_detalhes[s] for s in SQUADS_PRODUTOS_ABA
    }
    # Novos Leads (hoje/ontem) do Sniper na aba Métricas = soma dos funis Sniper+Navigator
    # (únicos onde Novos Leads faz sentido) — assim os dois nunca divergem.
    novos_leads_hoje = sum(
        produtos_detalhes[s][p]["novos_hoje"] for s in ("sniper", "navigator") for p in PRODUTOS_TODOS
    )
    novos_leads_ontem = sum(
        produtos_detalhes[s][p]["novos_ontem"] for s in ("sniper", "navigator") for p in PRODUTOS_TODOS
    )
    resultado["squads"]["Sniper"]["novos_leads_hoje"] = novos_leads_hoje
    resultado["squads"]["Sniper"]["novos_leads_ontem"] = novos_leads_ontem

    total_reaplicacoes_mes = sum(
        produtos_detalhes[s][p]["reaplicacoes_mes"] for s in ("sniper", "navigator") for p in PRODUTOS_TODOS
    )
    total_reaplicacoes_hoje = sum(
        produtos_detalhes[s][p]["reaplicacoes_hoje"] for s in ("sniper", "navigator") for p in PRODUTOS_TODOS
    )
    total_aplicacoes_mes = sum(
        produtos_detalhes[s][p]["novos_mes"] for s in ("sniper", "navigator") for p in PRODUTOS_TODOS
    ) + total_reaplicacoes_mes
    resultado["reaplicacoes"] = {
        "hoje": total_reaplicacoes_hoje,
        "mes": total_reaplicacoes_mes,
        "taxa_mes_pct": round(safe_div(total_reaplicacoes_mes, total_aplicacoes_mes) * 100, 1),
    }

    if e_mes_atual:
        resultado["mql_pfcc"] = analisar_mql_pfcc(ano, mes, hoje)
    else:
        resultado["mql_pfcc"] = {"hoje": 0, "mes": 0, "mes_novos": 0, "mes_reaplicados": 0}

    if e_mes_atual:
        resultado["perdidos_analise"] = analisar_perdidos(ano, mes, hoje, colaboradores, users_map)
    else:
        resultado["perdidos_analise"] = {
            "geral": {"total_mes": 0, "total_hoje": 0, "motivos": [], "auditoria": []},
            "por_funil": {SQUAD_DISPLAY[s]: {"total_mes": 0, "total_hoje": 0, "motivos": [], "auditoria": []} for s in SQUADS_PRODUTOS},
        }
    resultado["perdidos_analise"] = filtrar_perdidos_por_papel(resultado["perdidos_analise"], papel)
    resultado["papel_usuario"] = papel

    debug_previsto_hoje_deals = {s: [] for s in SQUADS_FINANCEIROS}
    previsto_hoje_por_bucket = {s: {"20": [], "50": [], "70": []} for s in SQUADS_FINANCEIROS}
    if e_mes_atual:
        alvo_iso = hoje.isoformat()
        vistos_dbg = set()
        for d in pool_previsto:
            did = d.get("id")
            if did in vistos_dbg:
                continue
            vistos_dbg.add(did)
            if (d.get("expected_close_date") or "")[:10] != alvo_iso:
                continue
            squad = squad_do_deal(d, colaboradores, users_map)
            try:
                prob = int(float(d.get("probability")))
            except (TypeError, ValueError):
                prob = None
            debug_previsto_hoje_deals.setdefault(squad, []).append({
                "id": did,
                "titulo": d.get("title"),
                "dono": owner_nome(d, users_map),
                "status": d.get("status"),
                "probability": prob,
                "valor": float(d.get("value") or 0),
            })
            if squad in previsto_hoje_por_bucket and prob in (20, 50, 70):
                previsto_hoje_por_bucket[squad][str(prob)].append({
                    "id": did,
                    "titulo": d.get("title") or f"Negócio {did}",
                    "url": f"https://{PD_DOMAIN}/deal/{did}",
                    "valor": float(d.get("value") or 0),
                })
    resultado["debug_previsto_hoje_deals"] = {
        SQUAD_DISPLAY.get(s, s): v[:30] for s, v in debug_previsto_hoje_deals.items() if s in SQUADS_FINANCEIROS
    }

    def _monta_bucket(lista):
        return {
            "quantidade": len(lista),
            "valor_total": round(sum(n["valor"] for n in lista), 2),
            "negocios": sorted(lista, key=lambda n: n["valor"], reverse=True)[:MAX_NEGOCIOS_POR_COLABORADOR],
        }

    resultado["previsto_hoje_detalhes"] = {
        SQUAD_DISPLAY[s]: {
            "20": _monta_bucket(previsto_hoje_por_bucket[s]["20"]),
            "50": _monta_bucket(previsto_hoje_por_bucket[s]["50"]),
            "70": _monta_bucket(previsto_hoje_por_bucket[s]["70"]),
        }
        for s in SQUADS_FINANCEIROS
    }
    resultado["previsto_hoje_detalhes"]["Total"] = {
        bucket: _monta_bucket([n for s in SQUADS_FINANCEIROS for n in previsto_hoje_por_bucket[s][bucket]])
        for bucket in ("20", "50", "70")
    }


    resultado["debug_closers"] = {
        SQUAD_DISPLAY[s]: sorted([
            {
                "nome": colaboradores[nome]["nome_exibicao"],
                "cargo": colaboradores[nome]["cargo"],
                "meta_financeira": metas.get(nome, {}).get("meta_fin", 0),
            }
            for nome in colaboradores
            if colaboradores[nome]["subarea"] == s
            and (metas.get(nome) or {}).get("papel") == "closer"
        ], key=lambda x: x["nome"])
        for s in SQUADS_FINANCEIROS
    }

    resultado["debug_forecast"] = {
        "total_pool_previsto": len(pool_previsto),
        "total_forecast_abertos_todos_status": len(forecast_abertos),
    }
    resultado["debug_sniper"] = {
        "atividades_m6_no_periodo": len(dados_m6["atividades"]),
        "deals_das_reunioes": len(dados_m6["deals"]),
        "responsaveis_sniper_contados": {nome_por_id.get(r, str(r)): len(v) for r, v in sorted(validadas_sniper.items(), key=lambda kv: -len(kv[1]))},
        "reunioes_fora_por_validada_nao": {nome_por_id.get(r, str(r)): n for r, n in nao_validadas_por_resp.items() if r in ids_sniper},
    }
    # Carimbo de parâmetros (CLAUDE.md da CRN, seção 5)
    resultado["carimbo"] = {
        "crn_versao": crn.versao(),
        "eixo": "ganho: data de ganho (BRT) · reuniões: data da reunião (BRT)",
        "valor": "c/ multiplicador (principal) e bruto",
        "recorte": "com meta (Metas do mês)",
        "squads": "Olympus e Sniper",
        "atualizado": dt.datetime.now(_fuso.TZ_BR).strftime("%d/%m/%Y %H:%M") if hasattr(_fuso, "TZ_BR") else _fuso.agora_brasil().strftime("%d/%m/%Y %H:%M"),
    }
    return resultado


# ---------------------------------------------------------------------------
# HANDLER (Vercel Python runtime)
# ---------------------------------------------------------------------------

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        if _ERRO_CRN:
            self._responder_erro_crn()
            return
        try:
            from urllib.parse import urlparse, parse_qs
            query = parse_qs(urlparse(self.path).query)
            rota = query.get("route", [None])[0]

            if rota == "me":
                self._responder_me()
                return

            if rota == "crn_versao":
                # CRN: GET /crn-versao, público, só a versão das regras (nenhum dado)
                body = json.dumps({"versao": crn.versao()}).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(body)
                return

            mes_param = int(query["mes"][0]) if "mes" in query else None
            ano_param = int(query["ano"][0]) if "ano" in query else None

            if rota == "metas_ba":
                payload = montar_resposta_metas_ba(mes_param, ano_param)
                body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(body)
                return

            if rota == "historico_mql_ganhos":
                payload = montar_historico_mql_ganhos_real(ano_param or 2026)
                body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(body)
                return

            if rota == "atingimento_historico":
                dia_inicio = int(query.get("dia_inicio", [1])[0])
                dia_fim = int(query.get("dia_fim", [8])[0])
                payload = montar_atingimento_historico(ano_param or 2026, mes_param or 8, dia_inicio, dia_fim)
                body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(body)
                return
            perdidos_inicio = query.get("perdidos_inicio", [None])[0]
            perdidos_fim = query.get("perdidos_fim", [None])[0]
            perdidos_responsavel = query.get("perdidos_responsavel", [None])[0]

            sessao = sessao_da_requisicao(self.headers)
            papel = sessao["papel"] if sessao else None

            if perdidos_inicio and perdidos_fim:
                # caminho leve: só o filtro de período/responsável, sem refazer o painel inteiro
                payload = montar_resposta_perdidos_periodo(
                    mes_param, ano_param, perdidos_inicio, perdidos_fim, perdidos_responsavel, papel
                )
            else:
                payload = montar_painel(ano_param=ano_param, mes_param=mes_param, papel=papel)

            body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            self.send_response(200)
        except Exception as e:
            body = json.dumps({"erro": str(e)}, ensure_ascii=False).encode("utf-8")
            self.send_response(500)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.end_headers()

    def do_POST(self):
        if _ERRO_CRN:
            self._responder_erro_crn()
            return
        from urllib.parse import urlparse, parse_qs
        query = parse_qs(urlparse(self.path).query)
        rota = query.get("route", [None])[0]
        if rota == "login":
            self._responder_login()
            return
        body = json.dumps({"erro": "rota não encontrada"}, ensure_ascii=False).encode("utf-8")
        self.send_response(404)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def _responder_erro_crn(self):
        body = json.dumps({"erro": "CRN não carregou: " + str(_ERRO_CRN)}, ensure_ascii=False).encode("utf-8")
        self.send_response(500)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def _responder_login(self):
        try:
            tamanho = int(self.headers.get("Content-Length", 0) or 0)
            corpo_bruto = self.rfile.read(tamanho) if tamanho else b"{}"
            dados = json.loads(corpo_bruto or b"{}")
            usuario = str(dados.get("usuario", "")).strip().lower()
            senha = str(dados.get("senha", ""))
            papel_encontrado = checa_login(usuario, senha)
            if papel_encontrado:
                resp = {"token": gera_token(usuario, papel_encontrado), "usuario": usuario, "papel": papel_encontrado}
                self.send_response(200)
            else:
                resp = {"error": "Usuário ou senha inválidos"}
                self.send_response(401)
        except Exception as e:
            resp = {"error": str(e)}
            self.send_response(500)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(json.dumps(resp, ensure_ascii=False).encode("utf-8"))

    def _responder_me(self):
        sessao = sessao_da_requisicao(self.headers)
        if sessao:
            resp = {"logado": True, "usuario": sessao["usuario"], "papel": sessao["papel"]}
        else:
            resp = {"logado": False}
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(json.dumps(resp, ensure_ascii=False).encode("utf-8"))
