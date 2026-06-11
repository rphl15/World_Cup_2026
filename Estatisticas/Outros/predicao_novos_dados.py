"""
=============================================================================
ATUALIZAÇÃO DE DADOS — media_fifa.xlsx / fifa_48_selecoes.xlsx
=============================================================================
Coleta os ÚLTIMOS N JOGOS (default 10) de cada uma das 48 seleções da Copa
do Mundo 2026 via API pública do Sofascore, extrai as ~37 métricas usadas
no projeto e regenera:

  - fifa_48_selecoes_atualizado.xlsx  → 1 aba por seleção (jogo a jogo + Média)
  - media_fifa_atualizado.xlsx        → Planilha1 (médias) + Planilha2 (correlação)

EXECUÇÃO: rode este script LOCALMENTE (na sua máquina). O ambiente do Claude
não consegue acessar api.sofascore.com (rede restrita / conteúdo via JS),
por isso este script foi feito para você rodar no seu computador.

DEPENDÊNCIAS:
    pip install requests pandas openpyxl tqdm

OBSERVAÇÕES IMPORTANTES:
  - O Sofascore às vezes responde 403 para requisições sem headers de
    navegador → por isso usamos headers + retries + backoff.
  - Os resultados intermediários (raw JSON) são salvos em ./cache_sofascore/
    para você não precisar refazer tudo se o script cair no meio.
  - Caso a API continue bloqueando, descomente o bloco PLAYWRIGHT_FALLBACK
    no final do arquivo (precisa `pip install playwright` e
    `playwright install chromium`).
  - O mapeamento de métricas (DICIONARIO_METRICAS) foi feito com base na
    estrutura padrão do endpoint /event/{id}/statistics do Sofascore.
    Alguns nomes podem variar um pouco — ajuste o dicionário se notar
    métricas vazias/zeradas em excesso após rodar.
=============================================================================
"""

# =============================================================================
# CONFIGURAÇÕES
# =============================================================================
N_JOGOS              = 10            # últimos N jogos por seleção
PASTA_CACHE          = './cache_sofascore'
ARQUIVO_SELECOES_OUT = './fifa_48_selecoes_atualizado.xlsx'
ARQUIVO_MEDIA_OUT    = './media_fifa_atualizado.xlsx'
TIMEOUT              = 15
PAUSA_ENTRE_REQS     = 1.2           # segundos, para evitar 429/403
MAX_TENTATIVAS       = 4

# =============================================================================
# IMPORTS
# =============================================================================
import os
import json
import time
import unicodedata
import requests
import pandas as pd
import numpy as np
from datetime import datetime

os.makedirs(PASTA_CACHE, exist_ok=True)

HEADERS = {
    'User-Agent': ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                    'AppleWebKit/537.36 (KHTML, like Gecko) '
                    'Chrome/124.0.0.0 Safari/537.36'),
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'pt-BR,pt;q=0.9,en;q=0.8',
    'Referer': 'https://www.sofascore.com/',
    'Origin': 'https://www.sofascore.com',
}

BASE = 'https://api.sofascore.com/api/v1'

# =============================================================================
# 1. LISTA DAS 48 SELEÇÕES (nomes EXATAMENTE como usados nos arquivos atuais)
#    + termo de busca em inglês para a API do Sofascore
# =============================================================================
SELECOES = {
    'Alemanha':              'Germany',
    'Argentina':             'Argentina',
    'Argélia':               'Algeria',
    'Arábia Saudita':        'Saudi Arabia',
    'Austrália':             'Australia',
    'Brasil':                'Brazil',
    'Bélgica':               'Belgium',
    'Bósnia e Herzegovina':  'Bosnia and Herzegovina',
    'Cabo Verde':            'Cape Verde',
    'Canadá':                'Canada',
    'Catar':                 'Qatar',
    'Colômbia':              'Colombia',
    'Coreia do Sul':         'South Korea',
    'Costa do Marfim':       "Ivory Coast",
    'Croácia':               'Croatia',
    'Curaçao':               'Curacao',
    'EUA':                   'USA',
    'Egito':                 'Egypt',
    'Equador':               'Ecuador',
    'Escócia':               'Scotland',
    'Espanha':               'Spain',
    'França':                'France',
    'Gana':                  'Ghana',
    'Haiti':                 'Haiti',
    'Holanda':               'Netherlands',
    'Inglaterra':            'England',
    'Iraque':                'Iraq',
    'Irã':                   'Iran',
    'Japão':                 'Japan',
    'Jordânia':              'Jordan',
    'Marrocos':              'Morocco',
    'México':                'Mexico',
    'Noruega':               'Norway',
    'Nova Zelândia':         'New Zealand',
    'Panamá':                'Panama',
    'Paraguai':              'Paraguay',
    'Portugal':              'Portugal',
    'RD Congo':              'DR Congo',
    'República Tcheca':      'Czechia',
    'Senegal':               'Senegal',
    'Suécia':                'Sweden',
    'Suíça':                 'Switzerland',
    'Tunísia':               'Tunisia',
    'Turquia':               'Turkey',
    'Uruguai':               'Uruguay',
    'Uzbequistão':           'Uzbekistan',
    'África do Sul':         'South Africa',
    'Áustria':               'Austria',
}

# =============================================================================
# 2. AS 37 MÉTRICAS — mapeamento (grupo, nome do stat no Sofascore)
#    `home_away` indica se pegamos sempre o valor do TIME (não do adversário)
# =============================================================================
# Estrutura de /event/{id}/statistics:
#   { "statistics": [ { "period": "ALL", "groups": [
#         { "groupName": "...", "statisticsItems": [
#               {"name": "...", "home": "..", "away": ".."} ] } ] } ] }
#
# Para cada métrica abaixo: (groupName, statName). Usamos correspondência
# por substring (case-insensitive) para tolerar pequenas variações.

MAPA_METRICAS = {
    'Posse de Bola':                       ('Possession', 'Ball possession'),
    'Total de chutes':                     ('Shots', 'Total shots'),
    'Chutes no gol':                       ('Shots', 'Shots on target'),
    'Chutes para fora':                    ('Shots', 'Shots off target'),
    'Chutes bloqueados':                   ('Shots', 'Blocked shots'),
    'Chutes dentro da área':               ('Shots', 'Shots inside box'),
    'Chutes fora da área':                 ('Shots', 'Shots outside box'),
    'Trave':                               ('Shots', 'Hit woodwork'),
    'Gols esperados (xG)':                 ('Expected', 'Expected goals'),
    'xG sofridos (xGC)':                   ('Expected', 'Expected goals'),  # do adversário
    'Gols esperados em chutes no alvo':    ('Expected', 'xG on target'),
    'Grandes chances perdidas':            ('Attack', 'Big chances missed'),
    'Chances perigosas criadas':           ('Attack', 'Big chances'),
    'Assistência esperada (xA)':           ('Expected', 'Expected assists'),
    'Passes decisivos':                    ('Passes', 'Key passes'),
    'Passes completos':                    ('Passes', 'Accurate passes'),
    'Passes no campo adversário':          ('Passes', 'Passes in opposition half'),
    'Passes no próprio campo':             ('Passes', 'Passes in own half'),
    'Passes no último terço':              ('Passes', 'Final third entries'),
    'Passes para trás':                    ('Passes', 'Long balls'),  # ajustar se houver "backward passes"
    'Laterais cobrados':                   ('Match overview', 'Throw-ins'),
    'Escanteios':                          ('Match overview', 'Corner kicks'),
    'Tiros de meta':                       ('Match overview', 'Goal kicks'),
    'Impedimentos':                        ('Match overview', 'Offsides'),
    'Faltas':                              ('Duels', 'Fouls'),
    'Faltas sofridas':                     ('Duels', 'Fouls'),  # adversário
    'Cartões amarelos':                    ('Match overview', 'Yellow cards'),
    'Cartões vermelhos':                   ('Match overview', 'Red cards'),
    'Duelos ganhos':                       ('Duels', 'Duels won'),
    'Driblado':                            ('Duels', 'Dribbled past'),
    'Interceptações':                      ('Defending', 'Interceptions'),
    'Perigo afastado':                     ('Defending', 'Clearances'),
    'Posse de bola perdida':               ('Duels', 'Possession lost'),
    'Posse recuperada no terço final':     ('Defending', 'Recoveries'),
    'Defesas do goleiro':                  ('Goalkeeping', 'Goalkeeper saves'),
    'Gol sofrido devido a erro individual':('Goalkeeping', 'Errors leading to goal'),
    'Ataque':                              ('GRAPH', 'attackPosition'),  # ver get_attack_value()
}

# =============================================================================
# 3. FUNÇÕES DE REQUISIÇÃO COM RETRY/BACKOFF
# =============================================================================

def get_json(url, tentativas=MAX_TENTATIVAS):
    """GET com retry/backoff e headers de navegador."""
    for i in range(tentativas):
        try:
            r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
            if r.status_code == 200:
                return r.json()
            elif r.status_code in (403, 429):
                wait = (i + 1) * 3
                print(f"   ⚠ HTTP {r.status_code} em {url} — aguardando {wait}s (tentativa {i+1}/{tentativas})")
                time.sleep(wait)
            else:
                print(f"   ⚠ HTTP {r.status_code} em {url}")
                return None
        except requests.RequestException as e:
            print(f"   ⚠ Erro de conexão: {e} — tentativa {i+1}/{tentativas}")
            time.sleep((i + 1) * 2)
    return None


def cache_path(nome):
    return os.path.join(PASTA_CACHE, nome)


def get_cached_or_fetch(cache_file, url):
    path = cache_path(cache_file)
    if os.path.exists(path):
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    data = get_json(url)
    time.sleep(PAUSA_ENTRE_REQS)
    if data is not None:
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f)
    return data


# =============================================================================
# 4. BUSCAR ID DO TIME NA API
# =============================================================================

def buscar_team_id(nome_busca):
    """Busca o ID da seleção nacional no Sofascore."""
    slug = unicodedata.normalize('NFKD', nome_busca).encode('ascii', 'ignore').decode().replace(' ', '%20')
    url = f"{BASE}/search/all?q={slug}"
    data = get_cached_or_fetch(f"search_{nome_busca.replace(' ', '_')}.json", url)
    if not data:
        return None
    for item in data.get('results', []):
        entity = item.get('entity', {})
        if item.get('type') == 'team' and entity.get('sport', {}).get('slug') == 'football':
            # Preferir seleções nacionais (sem clube / país == nome)
            if entity.get('national') or entity.get('name', '').lower() == nome_busca.lower():
                return entity['id']
    # fallback: primeiro time de futebol encontrado
    for item in data.get('results', []):
        entity = item.get('entity', {})
        if item.get('type') == 'team' and entity.get('sport', {}).get('slug') == 'football':
            return entity['id']
    return None


# =============================================================================
# 5. BUSCAR ÚLTIMOS N JOGOS DO TIME
# =============================================================================

def buscar_ultimos_jogos(team_id, n=N_JOGOS):
    """Retorna lista de eventos (jogos) finalizados, mais recentes primeiro."""
    eventos = []
    pagina = 0
    while len(eventos) < n and pagina < 3:  # cada página ~ até 20 jogos
        url = f"{BASE}/team/{team_id}/events/last/{pagina}"
        data = get_cached_or_fetch(f"events_{team_id}_p{pagina}.json", url)
        if not data or 'events' not in data:
            break
        finalizados = [e for e in data['events'] if e.get('status', {}).get('type') == 'finished']
        eventos.extend(finalizados)
        if not data.get('hasNextPage', False):
            break
        pagina += 1
    # mais recentes primeiro
    eventos = sorted(eventos, key=lambda e: e.get('startTimestamp', 0), reverse=True)
    return eventos[:n]


# =============================================================================
# 6. ESTATÍSTICAS DE UM JOGO
# =============================================================================

def buscar_estatisticas(event_id):
    url = f"{BASE}/event/{event_id}/statistics"
    return get_cached_or_fetch(f"stats_{event_id}.json", url)


def buscar_grafico_ataque(event_id):
    """Endpoint de gráfico de momento de jogo (usado para 'Ataque')."""
    url = f"{BASE}/event/{event_id}/graph"
    return get_cached_or_fetch(f"graph_{event_id}.json", url)


def get_attack_value(graph_data, is_home):
    """
    Soma a 'pressão ofensiva' do time ao longo do jogo a partir do gráfico
    de momento (valores positivos = pressão do mandante, negativos = visitante).
    Retorna um valor agregado comparável à coluna 'Ataque' original.
    """
    if not graph_data or 'graphPoints' not in graph_data:
        return np.nan
    pontos = [p.get('value', 0) for p in graph_data['graphPoints']]
    if not pontos:
        return np.nan
    if is_home:
        soma = sum(p for p in pontos if p > 0)
    else:
        soma = sum(-p for p in pontos if p < 0)
    return round(soma, 1)


# =============================================================================
# 7. EXTRAIR AS 37 MÉTRICAS DE UM JOGO PARA O TIME ALVO
# =============================================================================

def extrair_item(stats_json, group_substr, stat_substr, lado):
    """Procura um item de estatística por substring de grupo/nome e retorna o valor numérico do `lado` (home/away)."""
    if not stats_json or 'statistics' not in stats_json:
        return np.nan
    periodo = next((p for p in stats_json['statistics'] if p.get('period') == 'ALL'), None)
    if not periodo:
        periodo = stats_json['statistics'][0] if stats_json['statistics'] else None
    if not periodo:
        return np.nan
    for grupo in periodo.get('groups', []):
        if group_substr.lower() not in grupo.get('groupName', '').lower():
            continue
        for item in grupo.get('statisticsItems', []):
            if stat_substr.lower() in item.get('name', '').lower():
                val = item.get(f'{lado}Value', item.get(lado))
                if val is None:
                    continue
                if isinstance(val, str):
                    val = val.replace('%', '').strip()
                    try:
                        val = float(val)
                    except ValueError:
                        continue
                return float(val)
    return np.nan


def extrair_metricas_jogo(event, stats_json, graph_json, eh_mandante):
    """Extrai as 37 métricas para o time alvo nesse jogo."""
    lado = 'home' if eh_mandante else 'away'
    lado_oponente = 'away' if eh_mandante else 'home'
    resultado = {}

    for metrica, (grupo, stat) in MAPA_METRICAS.items():
        if grupo == 'GRAPH':
            resultado[metrica] = get_attack_value(graph_json, eh_mandante)
            continue
        if metrica == 'xG sofridos (xGC)':
            # xG do ADVERSÁRIO
            resultado[metrica] = extrair_item(stats_json, grupo, stat, lado_oponente)
        elif metrica == 'Faltas sofridas':
            # faltas cometidas pelo ADVERSÁRIO = faltas sofridas pelo time alvo
            resultado[metrica] = extrair_item(stats_json, grupo, stat, lado_oponente)
        elif metrica == 'Driblado':
            # "Dribbled past" do time alvo (quantas vezes foi driblado)
            resultado[metrica] = extrair_item(stats_json, grupo, stat, lado)
        else:
            resultado[metrica] = extrair_item(stats_json, grupo, stat, lado)

    return resultado


# =============================================================================
# 8. NOME DO ADVERSÁRIO (para coluna da planilha)
# =============================================================================

def nome_adversario(event, eh_mandante, mapa_pt):
    """Retorna o nome do adversário em português, se mapeável, senão em inglês."""
    adv = event['awayTeam']['name'] if eh_mandante else event['homeTeam']['name']
    inverso = {v.lower(): k for k, v in mapa_pt.items()}
    return inverso.get(adv.lower(), adv)


# =============================================================================
# 9. LOOP PRINCIPAL — COLETA DE TODAS AS SELEÇÕES
# =============================================================================

def coletar_selecao(nome_pt, nome_busca, mapa_pt):
    print(f"\n>>> {nome_pt} ({nome_busca})")
    team_id = buscar_team_id(nome_busca)
    if not team_id:
        print(f"   ✗ ID não encontrado para {nome_pt}")
        return None
    print(f"   ✓ team_id = {team_id}")

    jogos = buscar_ultimos_jogos(team_id, N_JOGOS)
    print(f"   ✓ {len(jogos)} jogos finalizados encontrados")

    colunas = {}
    for ev in jogos:
        eh_mandante = ev['homeTeam']['id'] == team_id
        adv = nome_adversario(ev, eh_mandante, mapa_pt)
        data_jogo = datetime.utcfromtimestamp(ev['startTimestamp']).strftime('%Y-%m-%d')
        col_name = f"{adv} ({data_jogo})"

        stats = buscar_estatisticas(ev['id'])
        graph = buscar_grafico_ataque(ev['id'])
        metricas = extrair_metricas_jogo(ev, stats, graph, eh_mandante)
        colunas[col_name] = metricas

    if not colunas:
        return None

    df = pd.DataFrame(colunas)
    df = df.reindex(list(MAPA_METRICAS.keys()))  # ordem fixa das 37 métricas
    df['Média'] = df.mean(axis=1, skipna=True).round(3)
    df.insert(0, 'Selecao', nome_pt)
    df.index.name = 'Metrica'
    df = df.reset_index()
    return df


def main():
    print("=" * 70)
    print(f"COLETANDO ÚLTIMOS {N_JOGOS} JOGOS DE {len(SELECOES)} SELEÇÕES")
    print("=" * 70)

    abas = {}
    medias = {}

    for nome_pt, nome_en in SELECOES.items():
        try:
            df_team = coletar_selecao(nome_pt, nome_en, SELECOES)
        except Exception as e:
            print(f"   ✗ ERRO em {nome_pt}: {e}")
            df_team = None

        if df_team is not None:
            abas[nome_pt] = df_team
            medias[nome_pt] = df_team.set_index('Metrica')['Média']
        else:
            print(f"   ⚠ {nome_pt} ficou sem dados — preenchendo com NaN")
            medias[nome_pt] = pd.Series({m: np.nan for m in MAPA_METRICAS.keys()})

    # ---- Salvar fifa_48_selecoes_atualizado.xlsx ----
    print("\n--- Salvando", ARQUIVO_SELECOES_OUT, "---")
    with pd.ExcelWriter(ARQUIVO_SELECOES_OUT, engine='openpyxl') as writer:
        for nome_pt, df_team in abas.items():
            sheet_name = f"{nome_pt}.xlsx"[:31]
            df_team.to_excel(writer, sheet_name=sheet_name, index=False)

    # ---- Salvar media_fifa_atualizado.xlsx ----
    print("--- Salvando", ARQUIVO_MEDIA_OUT, "---")
    df_media = pd.DataFrame(medias)  # index = métricas, colunas = seleções
    df_media.index.name = 'Paises'

    # Continente (mantém estrutura original — preencher manualmente se necessário)
    continentes_row = pd.Series({c: '' for c in df_media.columns}, name='Continente')
    df_media_full = pd.concat([continentes_row.to_frame().T, df_media])

    df_corr = df_media.T.corr().round(3)
    df_corr.index = [c.lower().replace('(', '').replace(')', '') for c in df_corr.index]
    df_corr.columns = df_corr.index

    with pd.ExcelWriter(ARQUIVO_MEDIA_OUT, engine='openpyxl') as writer:
        df_media_full.to_excel(writer, sheet_name='Planilha1')
        df_corr.to_excel(writer, sheet_name='Planilha2')

    print("\n✅ CONCLUÍDO!")
    print(f"   {ARQUIVO_SELECOES_OUT}")
    print(f"   {ARQUIVO_MEDIA_OUT}")
    print(f"   Cache salvo em: {PASTA_CACHE}/ (apague para forçar nova coleta)")


if __name__ == '__main__':
    main()


# =============================================================================
# PLAYWRIGHT FALLBACK (descomente se a API continuar retornando 403)
# =============================================================================
from playwright.sync_api import sync_playwright

def get_json_playwright(url):
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(extra_http_headers=HEADERS)
        page.goto(url)
        content = page.text_content('pre') or page.content()
        browser.close()
        try:
            return json.loads(content)
        except json.JSONDecodeError:
            return None

# Substitua, dentro de get_json(), a chamada requests por get_json_playwright(url)
# caso continue tomando 403 mesmo com headers + retries.