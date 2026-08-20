"""
Acompanhamento de Meta — Julho/2026 (Olympus/MGM, Elite, Sniper)
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
import datetime as dt
from http.server import BaseHTTPRequestHandler
from urllib.request import Request, urlopen
from urllib.parse import urlencode

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

PD_DOMAIN = "boardacademy.pipedrive.com"
PD_TOKEN = os.environ.get("PIPEDRIVE_API_TOKEN", "")

V1_BASE = f"https://{PD_DOMAIN}/api/v1"
V2_BASE = f"https://{PD_DOMAIN}/api/v2"

FILTER_DEALS = 74674          # vendas ganhas
FILTER_DEALS_RV = 1466157     # deals com reunião válida (whitelist SDR) — confirmado funcional no dashboard de referência
FILTER_ACTIVITIES = 1310451   # atividades / reuniões
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
PALAVRAS_MQL = [
    "socio", "sócio", "partner", "founder", "fundador", "proprietario", "proprietário", "propr", "owner",
    "empresario", "empresário", "entrepreneur", "presidente", "president",
    "ceo", "cmo", "cto", "coo", "cfo", "cpo", "ciso", "chro", "cro", "cio", "chief",
    "c-level", "c level", "vice-presidente", "vice", "vp",
    "conselh", "diretor", "director", "advisor", "board member",
]

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

# Mapa interno -> exibição. Chave interna "mgm" (funis Olympus + Navigator) exibe como "Olympus".
SQUAD_DISPLAY = {"mgm": "Olympus", "elite": "Elite", "sniper": "Sniper", "navigator": "Navigator"}
SQUADS_FINANCEIROS = ["mgm", "elite"]     # closers (valor em R$)
SQUAD_SDR = "sniper"                       # reuniões
NOMES_EXTRAS_SNIPER_CRU = {"Denise Mussolin"}  # contam nas reuniões do Sniper mesmo não sendo do squad

# Pessoas que nunca devem contar, mesmo que apareçam em alguma base (divergência de cadastro etc.)
EXCLUSOES_FIXAS = {"priscila ribeiro"}

# GM: vendas dela são redistribuídas por funil (pipeline) pro squad correspondente
GM_NOME_NORMALIZADO = None  # preencher com norm("Nome da GM") — ver README
FUNIL_PARA_SQUAD = {
    "elite": "elite",
    "sniper": "sniper",
    "olympus": "mgm",
    "mgm": "mgm",
    "navigator": "mgm",
}

# ⚠️ Parâmetros que NÃO vêm de nenhuma das 3 abas do Sheets (COLAB/METAS/FERIADOS):
# percentual do "gap intermediário" e a data-prazo (ex.: 40% até 16/07/2026).
# Por ora como constantes — mover pro Sheets se o time preferir editar sem deploy.
PCT_GAP_INTERMEDIARIO = 0.40

TZ_OFFSET_HORAS = 3  # Pipedrive retorna UTC; BRT = UTC - 3h (sem horário de verão)


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
    """Pipedrive timestamp (UTC) -> datetime BRT (subtrai 3h fixas)."""
    if not iso_ts:
        return None
    ts = iso_ts.replace("Z", "")
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S"):
        try:
            d = dt.datetime.strptime(ts[:19], fmt)
            return d - dt.timedelta(hours=TZ_OFFSET_HORAS)
        except ValueError:
            continue
    return None


def hoje_brt():
    """'Hoje' correto pro fuso de Brasília — NÃO usar dt.date.today() sozinho,
    que pega a data do servidor (UTC na Vercel). Entre ~21h e 23h59 BRT, o UTC já
    virou o dia seguinte, e dt.date.today() erradamente aponta pra 'amanhã',
    zerando tudo que compara com 'hoje' (Novos Leads Hoje, MQL Hoje, etc)."""
    return (dt.datetime.now(dt.timezone.utc) - dt.timedelta(hours=TZ_OFFSET_HORAS)).date()


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


def carregar_feriados():
    rows = read_csv(SHEET_FERIADOS)
    if not rows:
        return set()
    first_key = list(rows[0].keys())[0]
    feriados = set()
    for r in rows:
        raw = (r.get(first_key) or "").strip()
        if not raw:
            continue
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y"):
            try:
                feriados.add(dt.datetime.strptime(raw, fmt).date())
                break
            except ValueError:
                continue
    return feriados


def carregar_colaboradores(mes, ano):
    rows = read_csv(SHEET_COLAB)
    if not rows:
        return {}
    fields = rows[0].keys()
    col_nome = find_col(fields, exact="nome")
    col_subarea = find_col(fields, exact="subarea")
    col_cargo = find_col(fields, exact="cargo")
    col_status = find_col(fields, contains="status")
    col_head = find_col(fields, contains="head")
    col_lider = find_col(fields, contains_all=["lider", "team"])
    col_mes_ref = find_col(fields, contains_all=["mes", "ref"])
    col_ano_ref = find_col(fields, contains_all=["ano", "ref"])

    def filtra_mes(data):
        if not (col_mes_ref and col_ano_ref):
            return data
        out = [r for r in data if str(r.get(col_mes_ref, "")).strip() == str(mes)
               and str(r.get(col_ano_ref, "")).strip() == str(ano)]
        return out if out else data  # fallback silencioso (igual ao sistema grande)

    filtrado = filtra_mes(rows)

    dedup = {}
    for r in filtrado:
        nome = norm(r.get(col_nome, ""))
        if nome and nome not in dedup:
            dedup[nome] = r

    colaboradores = {}
    for nome, r in dedup.items():
        if nome in EXCLUSOES_FIXAS:
            continue
        status = norm(r.get(col_status, ""))
        if status != "ativo":
            continue
        colaboradores[nome] = {
            "nome_exibicao": r.get(col_nome, ""),
            "subarea": norm(r.get(col_subarea, "")),
            "cargo": norm(r.get(col_cargo, "")) if col_cargo else "",
            "head": norm(r.get(col_head, "")) if col_head else "",
            "lider": norm(r.get(col_lider, "")) if col_lider else "",
        }
    return colaboradores


def carregar_metas(mes, ano):
    rows = read_csv(SHEET_METAS)
    if not rows:
        return {}
    fields = rows[0].keys()
    col_ano = find_col(fields, exact="ano")
    col_mes = find_col(fields, exact="mes")
    col_nome = find_col(fields, exact="nome")
    col_reu = find_col(fields, contains_all=["reuni", "meta"])
    col_fin = find_col(fields, contains="financ")
    col_util = find_col(fields, contains="util")

    metas = {}
    for r in rows:
        if str(r.get(col_ano, "")).strip() != str(ano):
            continue
        if str(r.get(col_mes, "")).strip() != str(mes):
            continue
        nome = norm(r.get(col_nome, ""))
        if not nome:
            continue
        meta_reu = parse_num_br(r.get(col_reu, 0))
        meta_fin = parse_num_br(r.get(col_fin, 0))
        if meta_fin == 0 and meta_reu == 0:
            continue  # sem meta nenhuma no mês = pessoa some do cálculo
        papel = "closer" if meta_reu == 0 else "sdr"
        metas[nome] = {
            "meta_reu": meta_reu,
            "meta_fin": meta_fin,
            "dias_uteis_override": parse_num_br(r.get(col_util, 0)) if col_util else 0,
            "papel": papel,
        }
    return metas


# ---------------------------------------------------------------------------
# DIAS ÚTEIS
# ---------------------------------------------------------------------------

def eh_dia_util(d, feriados):
    return d.weekday() < 5 and d not in feriados


def dias_uteis_mes(ano, mes, feriados):
    primeiro = dt.date(ano, mes, 1)
    if mes == 12:
        ultimo = dt.date(ano, 12, 31)
    else:
        ultimo = dt.date(ano, mes + 1, 1) - dt.timedelta(days=1)
    d, total = primeiro, 0
    while d <= ultimo:
        if eh_dia_util(d, feriados):
            total += 1
        d += dt.timedelta(days=1)
    return total, ultimo


def calcular_dias_uteis(ano, mes, feriados, hoje=None):
    hoje = hoje or hoje_brt()
    total, ultimo_dia = dias_uteis_mes(ano, mes, feriados)

    if (ano, mes) < (hoje.year, hoje.month):
        return {"total": total, "passados": total, "restantes": 0}

    limite = min(hoje.day, ultimo_dia.day) if (ano, mes) == (hoje.year, hoje.month) else 0
    d, passados = dt.date(ano, mes, 1), 0
    while d.day <= limite:
        if eh_dia_util(d, feriados):
            passados += 1
        if d == ultimo_dia:
            break
        d += dt.timedelta(days=1)
    passados = max(passados, 1)  # piso de 1 (nunca zero)

    restantes = 0
    d = dt.date(ano, mes, limite + 1) if limite < ultimo_dia.day else None
    while d and d <= ultimo_dia:
        if eh_dia_util(d, feriados):
            restantes += 1
        d += dt.timedelta(days=1)

    return {"total": total, "passados": passados, "restantes": restantes}


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

def pd_v1_paginado(path, filter_id, extra_params=None):
    itens, start = [], 0
    while True:
        params = {"api_token": PD_TOKEN, "filter_id": filter_id, "limit": 500, "start": start}
        params.update(extra_params or {})
        data = http_get_json(f"{V1_BASE}{path}?{urlencode(params)}")
        chunk = data.get("data") or []
        itens.extend(chunk)
        pag = (data.get("additional_data") or {}).get("pagination") or {}
        if not pag.get("more_items_in_collection"):
            break
        start = pag.get("next_start", start + 500)
    return itens


def pd_v2_paginado(path, filter_id, extra_params=None):
    itens, cursor = [], None
    while True:
        params = {"filter_id": filter_id, "limit": 500}
        if cursor:
            params["cursor"] = cursor
        params.update(extra_params or {})
        headers = {"x-api-token": PD_TOKEN}
        data = http_get_json(f"{V2_BASE}{path}?{urlencode(params)}", headers=headers)
        itens.extend(data.get("data") or [])
        cursor = (data.get("additional_data") or {}).get("next_cursor")
        if not cursor:
            break
    return itens


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


def buscar_deals_ganhos(ano, mes, users_map):
    """FILTER_DEALS, status=won, sort won_time desc — filtra por mês (BRT) com corte na paginação."""
    itens = []
    start = 0
    alvo = f"{ano:04d}-{mes:02d}"
    while True:
        params = {
            "api_token": PD_TOKEN, "filter_id": FILTER_DEALS, "status": "won",
            "sort": "won_time DESC", "limit": 500, "start": start,
        }
        data = http_get_json(f"{V1_BASE}/deals?{urlencode(params)}")
        chunk = data.get("data") or []
        parou = False
        for deal in chunk:
            won_brt = to_brt(deal.get("won_time"))
            if not won_brt:
                continue
            chave_mes = won_brt.strftime("%Y-%m")
            if chave_mes < alvo:
                parou = True
                break
            if chave_mes == alvo:
                itens.append(deal)
        if parou:
            break
        pag = (data.get("additional_data") or {}).get("pagination") or {}
        if not pag.get("more_items_in_collection"):
            break
        start = pag.get("next_start", start + 500)
    return itens


def _conta_como_sniper(nome_norm, colaboradores):
    if nome_norm in {norm(n) for n in NOMES_EXTRAS_SNIPER_CRU}:
        return True
    return colaboradores.get(nome_norm, {}).get("subarea") == SQUAD_SDR


def squad_do_deal(deal, colaboradores, users_map):
    """Retorna squad interno (mgm/elite/sniper/...) via dono normalizado, com exceção da GM.
    Pros squads financeiros (Olympus/Elite), exige Closer, Head ou Gerente no cargo."""
    nome_dono = norm(owner_nome(deal, users_map))
    if GM_NOME_NORMALIZADO and nome_dono == GM_NOME_NORMALIZADO:
        funil = norm((deal.get("pipeline_name") or deal.get("pipeline_id") or ""))
        return FUNIL_PARA_SQUAD.get(funil)
    colaborador = colaboradores.get(nome_dono)
    if not colaborador:
        return None
    subarea = colaborador["subarea"]
    if subarea.startswith("lic"):
        return None
    cargo = colaborador.get("cargo", "")
    if subarea in SQUADS_FINANCEIROS and not any(termo in cargo for termo in ("closer", "head", "gerente")):
        return None
    return subarea


def teste_activities_sem_filtro():
    """Chama /v1/activities sem filter_id nenhum, só pra confirmar se o endpoint/token funciona."""
    params = {"api_token": PD_TOKEN, "limit": 5, "start": 0}
    data = http_get_json(f"{V1_BASE}/activities?{urlencode(params)}")
    itens = data.get("data") or []
    amostra = list(itens[0].keys()) if itens else []
    return {"total_sem_filtro": len(itens), "chaves_exemplo": amostra}


def buscar_activities(ano, mes):
    """v2, paginação por cursor — confirmado funcional (o filtro só precisava estar compartilhado/token com acesso)."""
    alvo = f"{ano:04d}-{mes:02d}"
    todas = pd_v2_paginado("/activities", FILTER_ACTIVITIES)
    filtradas = [a for a in todas if (a.get("due_date") or "")[:7] == alvo]
    return filtradas, len(todas)


def campo_owner_id(obj):
    """Extrai o id do dono, cobrindo o formato v1 (user_id, às vezes dict) e v2 (owner_id, int)."""
    v = obj.get("user_id", obj.get("owner_id"))
    if isinstance(v, dict):
        return v.get("value") or v.get("id")
    return v


def extrair_owner_id(deal):
    return campo_owner_id(deal)


def montar_deals_rv_owner_map(deals_rv):
    """Mapa deal_id -> owner_id, só dos deals que estão na whitelist FILTER_DEALS_RV."""
    return {d.get("id"): extrair_owner_id(d) for d in deals_rv}


def reuniao_valida_sdr(atividade, deals_rv_owner_map):
    concluida = atividade.get("done") is True or atividade.get("status") == "done"
    if not concluida:
        return False
    responsavel = campo_owner_id(atividade)
    deal_id = atividade.get("deal_id")
    if deal_id:
        # deal precisa estar na whitelist (senão não dá pra saber o dono, e a regra exige RV)
        if deal_id not in deals_rv_owner_map:
            return False
        dono_deal = deals_rv_owner_map[deal_id]
        if dono_deal == responsavel:
            return False
    return True


# ---------------------------------------------------------------------------
# MONTAGEM DO PAINEL
# ---------------------------------------------------------------------------

def buscar_forecast_deals():
    """Pipeline aberto (FILTER_FORECAST) — v1. Pede status=open explicitamente:
    sem isso, a API retorna TODOS os status que batem com o filtro (won/lost inclusive)."""
    return pd_v1_paginado("/deals", FILTER_FORECAST, extra_params={"status": "open"})


def buscar_probabilidade_por_stage():
    """Mapa stage_id -> deal_probability da etapa (probabilidade "original" configurada
    no funil). Usado pra negócios ganhos/perdidos, cujo campo probability o próprio
    Pipedrive costuma sobrescrever (geralmente pra 100% ou 0%) na hora do fechamento."""
    data = http_get_json(f"{V1_BASE}/stages?{urlencode({'api_token': PD_TOKEN})}")
    mapa = {}
    for s in (data.get("data") or []):
        prob = s.get("deal_probability")
        if prob is not None:
            mapa[s.get("id")] = prob
    return mapa


def valor_previsto_por_squad(pool, colaboradores, users_map, data_alvo):
    """Por squad financeiro: soma bruta separada por bucket de probabilidade
    (20/50/70, comparação exata) + a média ponderada (bruto x probabilidade).
    Só considera negócios com status=open — a API do Pipedrive não filtra isso
    sozinha (retorna todos os status do filtro por padrão)."""
    buckets = {s: {20: 0.0, 50: 0.0, 70: 0.0} for s in SQUADS_FINANCEIROS}
    vistos = set()
    alvo_iso = data_alvo.isoformat()
    for d in pool:
        did = d.get("id")
        if did in vistos:
            continue
        vistos.add(did)
        if d.get("status") != "open":
            continue
        if (d.get("expected_close_date") or "")[:10] != alvo_iso:
            continue
        try:
            prob = int(float(d.get("probability")))
        except (TypeError, ValueError):
            continue
        if prob not in (20, 50, 70):
            continue
        squad = squad_do_deal(d, colaboradores, users_map)
        if squad in buckets:
            buckets[squad][prob] += float(d.get("value") or 0)

    resultado = {}
    for squad, b in buckets.items():
        media = b[20] * 0.20 + b[50] * 0.50 + b[70] * 0.70
        resultado[squad] = {"p20": b[20], "p50": b[50], "p70": b[70], "media": media}
    return resultado


def valor_abertos_por_squad(pool, colaboradores, users_map, data_alvo):
    """Soma o valor bruto dos negócios ainda em aberto (status=open) com expected_close_date == data_alvo, por squad."""
    soma = {s: 0.0 for s in SQUADS_FINANCEIROS}
    vistos = set()
    alvo_iso = data_alvo.isoformat()
    for d in pool:
        did = d.get("id")
        if did in vistos:
            continue
        vistos.add(did)
        if d.get("status") != "open":
            continue
        if (d.get("expected_close_date") or "")[:10] != alvo_iso:
            continue
        squad = squad_do_deal(d, colaboradores, users_map)
        if squad in soma:
            soma[squad] += float(d.get("value") or 0)
    return soma


def valor_abertos_por_squad(pool, colaboradores, users_map, data_alvo):
    """Soma o valor bruto dos negócios ainda em aberto (status=open) com
    expected_close_date == data_alvo, por squad."""
    soma = {s: 0.0 for s in SQUADS_FINANCEIROS}
    vistos = set()
    alvo_iso = data_alvo.isoformat()
    for d in pool:
        did = d.get("id")
        if did in vistos:
            continue
        vistos.add(did)
        if d.get("status") != "open":
            continue
        if (d.get("expected_close_date") or "")[:10] != alvo_iso:
            continue
        squad = squad_do_deal(d, colaboradores, users_map)
        if squad in soma:
            soma[squad] += float(d.get("value") or 0)
    return soma


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


SQUAD_NOME_PIPELINE = {"mgm": "olympus", "elite": "elite"}  # nome do funil no Pipedrive
PIPELINE_ID_SNIPER = 88  # ID direto, sem precisar buscar por nome
PIPELINE_ID_NAVIGATOR = 51  # ID direto — só aparece na aba Produtos, não em Perdidos
SQUADS_PRODUTOS = ["mgm", "elite", "sniper"]  # squads usados em Perdidos (Olympus/Elite/Sniper)
SQUADS_PRODUTOS_ABA = SQUADS_PRODUTOS + ["navigator"]  # squads que aparecem na aba Produtos - Detalhamento


def _pipeline_id_do_squad(squad_interno):
    if squad_interno == "sniper":
        return PIPELINE_ID_SNIPER
    if squad_interno == "navigator":
        return PIPELINE_ID_NAVIGATOR
    return buscar_pipeline_id_por_nome(SQUAD_NOME_PIPELINE[squad_interno])


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
            if norm(cf_valor(d, CF_UTM_SOURCE)) == "org":
                continue
            cargo = cf_valor(d, CF_CARGO_NEGOCIO)
            if not (cargo and str(cargo).strip()):
                continue
            cargo_norm = norm(cargo)
            if not any(norm(palavra) in cargo_norm for palavra in PALAVRAS_MQL):
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

    # ---- Financeiro: Olympus (mgm) e Elite ----
    squads_fin = {s: {"bruto": 0.0, "multi": 0.0, "ontem": 0.0, "hoje": 0.0, "ontem_bruto": 0.0, "hoje_bruto": 0.0} for s in SQUADS_FINANCEIROS}
    for deal in deals_ganhos:
        squad = squad_do_deal(deal, colaboradores, users_map)
        if squad not in squads_fin:
            continue
        bruto = float(deal.get("value") or 0)
        multi = float(cf_valor(deal, CF_MULTIPLICADOR) or 0)
        squads_fin[squad]["bruto"] += bruto
        squads_fin[squad]["multi"] += multi
        won_brt = to_brt(deal.get("won_time"))
        if won_brt and won_brt.date() == ontem:
            squads_fin[squad]["ontem"] += multi
            squads_fin[squad]["ontem_bruto"] += bruto
        if won_brt and won_brt.date() == hoje:
            squads_fin[squad]["hoje"] += multi
            squads_fin[squad]["hoje_bruto"] += bruto

    # ---- Previsto (forecast) hoje/ontem — só faz sentido no mês atual ----
    # Só negócios em aberto (revertido: não inclui mais ganhos/perdidos).
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
        return sum(m["meta_fin"] for nome, m in metas.items()
                   if colaboradores.get(nome, {}).get("subarea") == squad_interno
                   and any(termo in colaboradores.get(nome, {}).get("cargo", "") for termo in ("closer", "head", "gerente")))

    ritmo_100 = safe_div(du["passados"], du["total"])

    # Prazo do gap intermediário (40%) = sempre a metade do mês selecionado,
    # arredondando pra cima (31 dias -> dia 16; 30 -> dia 15; 29 -> dia 15; 28 -> dia 14).
    dias_no_mes_calendario = calendar.monthrange(ano, mes)[1]
    dia_metade = math.ceil(dias_no_mes_calendario / 2)
    prazo_gap_intermediario = dt.date(ano, mes, dia_metade)

    # Dias úteis da 1ª metade do mês (usado como base pro Meta/Dia 40% em mês fechado)
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
        # mês fechado: não existe "dias restantes até o prazo" — usa a 1ª metade do mês inteira
        restantes_prazo = 0
        dias_restantes_p40 = dias_uteis_primeira_metade

    ritmo_prazo = safe_div(du["passados"], du["passados"] + restantes_prazo)

    resultado = {
        "squads": {}, "geradoEm": dt.datetime.now(dt.timezone.utc).isoformat(),
        "mes": mes, "ano": ano, "e_mes_atual": e_mes_atual,
    }

    for squad_interno in SQUADS_FINANCEIROS:
        meta_mes = meta_squad(squad_interno)
        meta_dia = safe_div(meta_mes, du["total"])
        bruto = squads_fin[squad_interno]["bruto"]
        multi = squads_fin[squad_interno]["multi"]
        onde_100 = meta_mes * ritmo_100
        onde_40 = (PCT_GAP_INTERMEDIARIO * meta_mes) * ritmo_prazo
        gap_100 = max(0.0, meta_mes - multi)
        gap_40 = max(0.0, (PCT_GAP_INTERMEDIARIO * meta_mes) - multi)
        meta_dia_40 = safe_div(gap_40, dias_restantes_p40)
        meta_dia_100 = safe_div(gap_100, dias_restantes_p100)  # gap / dias úteis restantes até o fim do mês (hoje conta se for o último)
        gap_100_bruto = max(0.0, meta_mes - bruto)
        meta_dia_100_bruto = safe_div(gap_100_bruto, dias_restantes_p100)

        resultado["squads"][SQUAD_DISPLAY[squad_interno]] = {
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
        }

    total_meta_mes = sum(resultado["squads"][SQUAD_DISPLAY[s]]["meta_mes"] for s in SQUADS_FINANCEIROS)
    total_bruto = sum(resultado["squads"][SQUAD_DISPLAY[s]]["realizado_bruto"] for s in SQUADS_FINANCEIROS)
    total_multi = sum(resultado["squads"][SQUAD_DISPLAY[s]]["realizado_multiplicador"] for s in SQUADS_FINANCEIROS)
    total_ontem = sum(resultado["squads"][SQUAD_DISPLAY[s]]["ontem"] for s in SQUADS_FINANCEIROS)
    total_hoje = sum(resultado["squads"][SQUAD_DISPLAY[s]]["hoje"] for s in SQUADS_FINANCEIROS)
    total_ontem_bruto = sum(resultado["squads"][SQUAD_DISPLAY[s]]["ontem_bruto"] for s in SQUADS_FINANCEIROS)
    total_hoje_bruto = sum(resultado["squads"][SQUAD_DISPLAY[s]]["hoje_bruto"] for s in SQUADS_FINANCEIROS)
    campos_previsto = ["previsto_hoje_20", "previsto_hoje_50", "previsto_hoje_70", "previsto_hoje_media",
                        "previsto_ontem_20", "previsto_ontem_50", "previsto_ontem_70", "previsto_ontem_media",
                        "em_aberto_hoje"]
    totais_previsto = {
        campo: sum(resultado["squads"][SQUAD_DISPLAY[s]][campo] for s in SQUADS_FINANCEIROS)
        for campo in campos_previsto
    }
    total_gap_40 = max(0.0, (PCT_GAP_INTERMEDIARIO * total_meta_mes) - total_multi)
    resultado["squads"]["Total"] = {
        "meta_mes": round(total_meta_mes, 2),
        "meta_dia": round(safe_div(total_meta_mes, du["total"]), 2),
        "realizado_bruto": round(total_bruto, 2),
        "realizado_multiplicador": round(total_multi, 2),
        "onde_deveria_100": round(total_meta_mes * ritmo_100, 2),
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
    }

    # ---- Sniper: reuniões ----
    activities, total_bruto_activities = buscar_activities(ano, mes)
    deals_rv = pd_v1_paginado("/deals", FILTER_DEALS_RV, extra_params={"status": "all_not_deleted"})
    deals_rv_owner_map = montar_deals_rv_owner_map(deals_rv)

    validadas_total = 0
    validadas_hoje = 0
    validadas_ontem = 0
    validadas_dia_anterior_util = 0
    previsto_hoje_reu = 0
    previsto_ontem_reu = 0
    d_ontem_util = dia_util_anterior(hoje, feriados)

    dbg_concluidas = 0
    dbg_com_deal = 0
    dbg_passou_valida_sdr = 0
    dbg_subarea_sniper = 0
    nomes_sem_match = set()

    for a in activities:
        if a.get("done") is True or a.get("status") == "done":
            dbg_concluidas += 1
        if a.get("deal_id"):
            dbg_com_deal += 1

        # "Previsto" = agendada pra aquele dia (qualquer status), independente de validação —
        # olha só o dono da atividade e a subarea dele, sem exigir Reunião Válida.
        nome_resp_previsto = norm(users_map.get(campo_owner_id(a), ""))
        if _conta_como_sniper(nome_resp_previsto, colaboradores):
            due_previsto = a.get("due_date")
            if due_previsto == hoje.isoformat():
                previsto_hoje_reu += 1
            if due_previsto == d_ontem_util.isoformat():
                previsto_ontem_reu += 1

        if not reuniao_valida_sdr(a, deals_rv_owner_map):
            continue
        dbg_passou_valida_sdr += 1
        # escopo só sniper: precisaria mapear owner->squad; aqui assume-se filtro já traz o universo certo
        nome_resp = norm(users_map.get(campo_owner_id(a), ""))
        if not _conta_como_sniper(nome_resp, colaboradores):
            if nome_resp:
                nomes_sem_match.add(nome_resp)
            continue
        dbg_subarea_sniper += 1
        validadas_total += 1
        due = a.get("due_date")
        if due == hoje.isoformat():
            validadas_hoje += 1
        if due == d_ontem_util.isoformat():
            validadas_dia_anterior_util += 1

    meta_reunioes = sum(m["meta_reu"] for nome, m in metas.items()
                         if _conta_como_sniper(nome, colaboradores))
    meta_dia_reu = safe_div(meta_reunioes, du["total"])
    onde_100_reu = meta_reunioes * ritmo_100
    onde_40_reu = (PCT_GAP_INTERMEDIARIO * meta_reunioes) * ritmo_prazo
    gap_100_reu = max(0.0, meta_reunioes - validadas_total)
    gap_40_reu = max(0.0, (PCT_GAP_INTERMEDIARIO * meta_reunioes) - validadas_total)
    meta_dia_40_reu = safe_div(gap_40_reu, dias_restantes_p40)
    meta_dia_100_reu = safe_div(gap_100_reu, dias_restantes_p100)

    resultado["squads"]["Sniper"] = {
        "meta_mes_reunioes": meta_reunioes,
        "meta_dia_reunioes": round(meta_dia_reu, 2),
        "realizado_reunioes": validadas_total,
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
    resultado["debug_previsto_hoje_deals"] = {
        SQUAD_DISPLAY.get(s, s): v[:30] for s, v in debug_previsto_hoje_deals.items() if s in SQUADS_FINANCEIROS
    }

    nome_teste = norm("Denise Mussolin")
    colab_teste = colaboradores.get(nome_teste)
    deals_dela = [
        {"id": d.get("id"), "titulo": d.get("title"), "valor": d.get("value"),
         "dono_no_deal": owner_nome(d, users_map), "squad_atribuido": squad_do_deal(d, colaboradores, users_map)}
        for d in deals_ganhos if norm(owner_nome(d, users_map)) == nome_teste
    ]
    resultado["debug_denise"] = {
        "esta_na_colab": colab_teste is not None,
        "dados_colab": colab_teste,
        "qtd_deals_ganhos_dela_no_mes": len(deals_dela),
        "deals_dela": deals_dela[:10],
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
            and any(termo in colaboradores[nome].get("cargo", "") for termo in ("closer", "head", "gerente"))
        ], key=lambda x: x["nome"])
        for s in SQUADS_FINANCEIROS
    }

    resultado["debug_forecast"] = {
        "total_pool_previsto": len(pool_previsto),
        "total_forecast_abertos_todos_status": len(forecast_abertos),
    }
    resultado["debug_sniper"] = {
        "teste_sem_filtro": teste_activities_sem_filtro(),
        "total_activities_bruto_sem_filtro_mes": total_bruto_activities,
        "total_activities_no_mes": len(activities),
        "concluidas": dbg_concluidas,
        "com_deal_id": dbg_com_deal,
        "deals_rv_total": len(deals_rv),
        "passou_validacao_sdr": dbg_passou_valida_sdr,
        "com_subarea_sniper": dbg_subarea_sniper,
        "nomes_sem_match_no_colab": sorted(nomes_sem_match)[:15],
    }
    return resultado


# ---------------------------------------------------------------------------
# HANDLER (Vercel Python runtime)
# ---------------------------------------------------------------------------

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        try:
            from urllib.parse import urlparse, parse_qs
            query = parse_qs(urlparse(self.path).query)
            rota = query.get("route", [None])[0]

            if rota == "me":
                self._responder_me()
                return

            mes_param = int(query["mes"][0]) if "mes" in query else None
            ano_param = int(query["ano"][0]) if "ano" in query else None
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
