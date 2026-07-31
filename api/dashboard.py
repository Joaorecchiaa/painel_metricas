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

SHEET_COLAB = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSvwO3Ag2f2cbkVgR1pJZp6fANQcbualGKlAG50fmOljuEGKZ1gJBbSAjRdO3SomXUEVQOWnTvlfHRd/pub?gid=1782440078&single=true&output=csv"
SHEET_METAS = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSvwO3Ag2f2cbkVgR1pJZp6fANQcbualGKlAG50fmOljuEGKZ1gJBbSAjRdO3SomXUEVQOWnTvlfHRd/pub?gid=0&single=true&output=csv"
SHEET_FERIADOS = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSvwO3Ag2f2cbkVgR1pJZp6fANQcbualGKlAG50fmOljuEGKZ1gJBbSAjRdO3SomXUEVQOWnTvlfHRd/pub?gid=1010928978&single=true&output=csv"

# Mapa interno -> exibição. Chave interna "mgm" (funis Olympus + Navigator) exibe como "Olympus".
SQUAD_DISPLAY = {"mgm": "Olympus", "elite": "Elite", "sniper": "Sniper"}
SQUADS_FINANCEIROS = ["mgm", "elite"]     # closers (valor em R$)
SQUAD_SDR = "sniper"                       # reuniões

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
PRAZO_GAP_INTERMEDIARIO = dt.date(2026, 7, 16)

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


from urllib.error import HTTPError, URLError

def http_get_json(url, headers=None):
    req = Request(url, headers=headers or {})
    try:
        with urlopen(req, timeout=25) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except HTTPError as e:
        detalhe = e.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"Erro HTTP {e.code} em {url}: {detalhe[:600]}") from e
    except URLError as e:
        raise RuntimeError(f"Falha de conexão em {url}: {e.reason}") from e


def http_get_text(url):
    req = Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urlopen(req, timeout=25) as resp:
            return resp.read().decode("utf-8")
    except HTTPError as e:
        detalhe = e.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"Erro HTTP {e.code} em {url}: {detalhe[:600]}") from e
    except URLError as e:
        raise RuntimeError(f"Falha de conexão em {url}: {e.reason}") from e


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
    hoje = hoje or dt.date.today()
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


def cf_valor(deal, hash_):
    v = deal.get(hash_)
    if isinstance(v, dict):
        return v.get("value")
    return v


def owner_nome(deal, users_map):
    owner = deal.get("user_id")
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


def squad_do_deal(deal, colaboradores, users_map):
    """Retorna squad interno (mgm/elite/sniper/...) via dono normalizado, com exceção da GM."""
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
    """Pipeline aberto (FILTER_FORECAST) — v1."""
    return pd_v1_paginado("/deals", FILTER_FORECAST)


def valor_previsto_por_squad(pool, colaboradores, users_map, data_alvo):
    """Soma (valor bruto x probabilidade) dos negócios com expected_close_date == data_alvo
    e probability em {20, 50, 70} (comparação exata), por squad financeiro."""
    soma = {s: 0.0 for s in SQUADS_FINANCEIROS}
    vistos = set()
    alvo_iso = data_alvo.isoformat()
    for d in pool:
        did = d.get("id")
        if did in vistos:
            continue
        vistos.add(did)
        if (d.get("expected_close_date") or "")[:10] != alvo_iso:
            continue
        try:
            prob = int(float(d.get("probability")))
        except (TypeError, ValueError):
            continue
        if prob not in (20, 50, 70):
            continue
        squad = squad_do_deal(d, colaboradores, users_map)
        if squad in soma:
            soma[squad] += float(d.get("value") or 0) * (prob / 100)
    return soma


def montar_painel():
    hoje = dt.date.today()
    ano, mes = hoje.year, hoje.month

    feriados = carregar_feriados()
    colaboradores = carregar_colaboradores(mes, ano)
    metas = carregar_metas(mes, ano)
    du = calcular_dias_uteis(ano, mes, feriados)
    # Se não sobra nenhum dia útil DEPOIS de hoje, mas hoje ainda é dia útil,
    # considera hoje como "1 dia" pra não zerar o Meta/Dia 100% no último dia do mês.
    dias_restantes_p100 = du["restantes"] if du["restantes"] > 0 else (1 if eh_dia_util(hoje, feriados) else 0)

    users_map = pd_users()
    deals_ganhos = buscar_deals_ganhos(ano, mes, users_map)

    ontem = dia_util_anterior(hoje, feriados)
    prox_dia_util = proximo_dia_util(hoje, feriados)

    # ---- Financeiro: Olympus (mgm) e Elite ----
    squads_fin = {s: {"bruto": 0.0, "multi": 0.0, "ontem": 0.0, "hoje": 0.0} for s in SQUADS_FINANCEIROS}
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
        if won_brt and won_brt.date() == hoje:
            squads_fin[squad]["hoje"] += multi

    # ---- Previsto (forecast) hoje/ontem ----
    # Sem snapshot histórico: "previsto para ontem" olha o pipeline aberto de hoje
    # + os negócios já ganhos este mês (cobre quem fechou); negócios perdidos com
    # aquela data prevista não entram aqui (limitação conhecida, sem persistência).
    forecast_abertos = buscar_forecast_deals()
    pool_previsto = forecast_abertos + deals_ganhos
    previsto_hoje = valor_previsto_por_squad(pool_previsto, colaboradores, users_map, hoje)
    previsto_ontem = valor_previsto_por_squad(pool_previsto, colaboradores, users_map, ontem)

    def meta_squad(squad_interno):
        return sum(m["meta_fin"] for nome, m in metas.items()
                   if colaboradores.get(nome, {}).get("subarea") == squad_interno)

    ritmo_100 = safe_div(du["passados"], du["total"])
    dias_ate_prazo = max((PRAZO_GAP_INTERMEDIARIO - hoje).days, 0) if PRAZO_GAP_INTERMEDIARIO >= hoje else 0
    # dias úteis restantes até o prazo dos 40% (aprox. via feriados)
    restantes_prazo = 0
    d = hoje + dt.timedelta(days=1)
    while d <= PRAZO_GAP_INTERMEDIARIO:
        if eh_dia_util(d, feriados):
            restantes_prazo += 1
        d += dt.timedelta(days=1)
    ritmo_prazo = safe_div(du["passados"], du["passados"] + restantes_prazo)

    resultado = {"squads": {}, "geradoEm": dt.datetime.now().isoformat()}

    for squad_interno in SQUADS_FINANCEIROS:
        meta_mes = meta_squad(squad_interno)
        meta_dia = safe_div(meta_mes, du["total"])
        bruto = squads_fin[squad_interno]["bruto"]
        multi = squads_fin[squad_interno]["multi"]
        onde_100 = meta_mes * ritmo_100
        onde_40 = (PCT_GAP_INTERMEDIARIO * meta_mes) * ritmo_prazo
        gap_100 = max(0.0, meta_mes - multi)
        gap_40 = max(0.0, (PCT_GAP_INTERMEDIARIO * meta_mes) - multi)
        meta_dia_40 = safe_div(gap_40, restantes_prazo)
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
            "previsto_hoje": round(previsto_hoje.get(squad_interno, 0.0), 2),
            "previsto_ontem": round(previsto_ontem.get(squad_interno, 0.0), 2),
        }

    total_meta_mes = sum(resultado["squads"][SQUAD_DISPLAY[s]]["meta_mes"] for s in SQUADS_FINANCEIROS)
    total_bruto = sum(resultado["squads"][SQUAD_DISPLAY[s]]["realizado_bruto"] for s in SQUADS_FINANCEIROS)
    total_multi = sum(resultado["squads"][SQUAD_DISPLAY[s]]["realizado_multiplicador"] for s in SQUADS_FINANCEIROS)
    total_ontem = sum(resultado["squads"][SQUAD_DISPLAY[s]]["ontem"] for s in SQUADS_FINANCEIROS)
    total_hoje = sum(resultado["squads"][SQUAD_DISPLAY[s]]["hoje"] for s in SQUADS_FINANCEIROS)
    total_previsto_hoje = sum(resultado["squads"][SQUAD_DISPLAY[s]]["previsto_hoje"] for s in SQUADS_FINANCEIROS)
    total_previsto_ontem = sum(resultado["squads"][SQUAD_DISPLAY[s]]["previsto_ontem"] for s in SQUADS_FINANCEIROS)
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
        "meta_dia_40": round(safe_div(total_gap_40, restantes_prazo), 2),
        "meta_dia_100": round(safe_div(max(0.0, total_meta_mes - total_multi), dias_restantes_p100), 2),
        "gap_100_bruto": round(max(0.0, total_meta_mes - total_bruto), 2),
        "meta_dia_100_bruto": round(safe_div(max(0.0, total_meta_mes - total_bruto), dias_restantes_p100), 2),
        "ontem": round(total_ontem, 2),
        "hoje": round(total_hoje, 2),
        "previsto_hoje": round(total_previsto_hoje, 2),
        "previsto_ontem": round(total_previsto_ontem, 2),
    }

    # ---- Sniper: reuniões ----
    activities, total_bruto_activities = buscar_activities(ano, mes)
    deals_rv = pd_v1_paginado("/deals", FILTER_DEALS_RV, extra_params={"status": "all_not_deleted"})
    deals_rv_owner_map = montar_deals_rv_owner_map(deals_rv)

    validadas_total = 0
    validadas_hoje = 0
    validadas_ontem = 0
    validadas_dia_anterior_util = 0
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
        if not reuniao_valida_sdr(a, deals_rv_owner_map):
            continue
        dbg_passou_valida_sdr += 1
        # escopo só sniper: precisaria mapear owner->squad; aqui assume-se filtro já traz o universo certo
        nome_resp = norm(users_map.get(campo_owner_id(a), ""))
        if colaboradores.get(nome_resp, {}).get("subarea") != SQUAD_SDR:
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
                         if colaboradores.get(nome, {}).get("subarea") == SQUAD_SDR)
    meta_dia_reu = safe_div(meta_reunioes, du["total"])
    onde_100_reu = meta_reunioes * ritmo_100
    onde_40_reu = (PCT_GAP_INTERMEDIARIO * meta_reunioes) * ritmo_prazo
    gap_100_reu = max(0.0, meta_reunioes - validadas_total)
    gap_40_reu = max(0.0, (PCT_GAP_INTERMEDIARIO * meta_reunioes) - validadas_total)
    meta_dia_40_reu = safe_div(gap_40_reu, restantes_prazo)
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
    }

    resultado["premissas"] = {
        "dias_uteis_total": du["total"],
        "dias_uteis_passados": du["passados"],
        "dias_uteis_restantes": du["restantes"],
        "dias_restantes_p100": dias_restantes_p100,
        "prazo_gap_intermediario": PRAZO_GAP_INTERMEDIARIO.isoformat(),
        "percentual_gap_intermediario": PCT_GAP_INTERMEDIARIO,
        "proximo_dia_util": prox_dia_util.isoformat(),
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
            payload = montar_painel()
            body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            self.send_response(200)
        except Exception as e:
            body = json.dumps({"erro": str(e)}, ensure_ascii=False).encode("utf-8")
            self.send_response(500)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)
