# Acompanhamento de Meta — Julho 2026

Painel de acompanhamento de meta (Olympus/MGM, Elite, Sniper), com dados
buscados direto do Pipedrive e do Google Sheets a cada carregamento
(polling automático a cada 10 min + botão manual). Sem cache.

## Deploy

1. Suba esta pasta para um repositório novo no GitHub.
2. Importe o repositório na Vercel.
3. Configure a variável de ambiente `PIPEDRIVE_API_TOKEN` no projeto da Vercel
   (Settings → Environment Variables) — **nunca** coloque o token no código.
4. Deploy. O front acessa `/api/dashboard` (mesma origem, sem CORS).

## Pontos em aberto / decisões pendentes

1. **`GM_NOME_NORMALIZADO`** (topo de `api/dashboard.py`): preencher com o
   nome normalizado da Gerente Geral pra ativar a regra de redistribuição
   de vendas dela por funil. Deixei `None` por enquanto (regra desligada).

2. **`PCT_GAP_INTERMEDIARIO` (40%) e `PRAZO_GAP_INTERMEDIARIO` (16/07/2026)**:
   esses dois parâmetros não vêm de nenhuma das 3 abas do Sheets (COLAB,
   METAS, FERIADOS) — são específicos desta planilha. Por ora estão como
   constantes no código; se quiser editar sem precisar de deploy, dá pra
   mover pra uma 4ª aba do mesmo Sheets.

3. **`/api/dashboard.py::buscar_activities`**: a paginação por cursor (v2)
   está simplificada — validar com o volume real de atividades do mês se
   precisa de mais de uma página por chamada.

4. **Mapeamento pipeline → nome de funil normalizado** (`squad_do_deal`):
   usa `deal.pipeline_name` quando vier na resposta; se a API não trouxer,
   precisa buscar em `/v1/pipelines` e mapear por `pipeline_id`.

## Onde cada coisa mora

- `api/dashboard.py` — toda a lógica (Sheets, Pipedrive, cálculos, JSON).
- `public/index.html` — front-end (fetch + polling + botão).
- `vercel.json` — roteamento da função Python + página estática.
