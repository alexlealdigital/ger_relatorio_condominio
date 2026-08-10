"""
Parser do formato Condomínio21 / BalAnual.rpt — Gestor Financeiro v3 Web
Detecta e processa o export real do sistema do cliente:

  Rubrica/Despesa | Previsão AAAA Aprovada | 01/AAAA ... 12/AAAA | Total |
  Realizado AAAA Média | Previsão AAAA+1 | Índice

Valores das colunas "Previsão Aprovada", "Média" e "Previsão AAAA+1" são MENSAIS.
"Total" é o realizado ANUAL. Rodapé traz "Previsão de Reajuste" (índice global),
"Rateio das Despesas / NNN Unidades" e "Fundo de Reserva - BC X%".
"""

import re
import numpy as np
import pandas as pd

MESES_PT = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun',
            'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']

# Categorização automática das rubricas (avaliada em ordem; primeira que casar vence)
CATEGORIAS = [
    ('Pessoal e Encargos', [
        'salário', 'salario', 'funcionário', 'funcionario', 'inss', 'fgts',
        'décimo', 'decimo', '13º', 'férias', 'ferias', 'rescis', 'vale',
        'encargo', 'folha',
    ]),
    ('Serviços Contratados', [
        'monitoramento', 'portaria', 'vigil', 'recolhimento', 'coleta',
        'serviço de fiscaliza', 'servico de fiscaliza', 'limpeza e conserva',
        'jardinagem', 'dedetiza', 'desentupi',
    ]),
    ('Administração e Jurídico', [
        'prolabore', 'pró-labore', 'pro-labore', 'honorário', 'honorario',
        'contábe', 'contabe', 'advocat', 'advogad', 'administra', 'cartór',
        'cartor', 'correio', 'assessoria', 'acordo', 'judicial',
    ]),
    ('Consumo', [
        'energia', 'água', 'agua', 'gás', 'glp', 'telefone', 'internet',
        'esgoto', 'saneamento',
    ]),
    ('Manutenção e Materiais', [
        'material', 'manuten', 'reforma', 'reparo', 'pintura', 'elétric',
        'eletric', 'hidraul', 'hidrául', 'ferragen', 'ferragista',
        'esquadria', 'calha', 'rufo', 'placa', 'laudo', 'projeto',
        'terraplanag', 'vidraç', 'vidrac', 'serralhe', 'gesse', 'paviment',
        'aluguel de equipamento', 'recarga', 'extintor', 'bomba', 'portão',
        'portao', 'equipamento',
    ]),
    ('Bancárias e Tributos', [
        'bancár', 'bancar', 'tarifa', 'liquida', 'boleto', 'issqn', 'iptu',
        'ipva', 'imposto', 'tributo', 'taxa munic', 'juros', 'multa',
        'darf', 'gps', 'cofins', 'pis', 'csll', 'irrf',
    ]),
]
CATEGORIA_OUTROS = 'Outras Despesas'

_RE_MES = re.compile(r'^(0[1-9]|1[0-2])/(\d{4})$')


def _norm(v) -> str:
    return str(v).strip() if v is not None and str(v) != 'nan' else ''


def _num(v):
    try:
        f = float(v)
        return None if np.isnan(f) else f
    except (TypeError, ValueError):
        return None


def _localizar_header(df: pd.DataFrame):
    """
    Encontra o header (contém 'rubrica' + colunas mensais MM/AAAA).
    Retorna (indice_da_linha, None) se o header for uma linha de dados,
    ou (None, lista_do_header) se ele já estiver em df.columns
    (caso do load_data legado, que promove a 1ª linha a columns).
    Retorna None se o formato não for BalAnual.
    """
    # Header promovido a columns?
    cols = [_norm(c) for c in df.columns.tolist()]
    meses = sum(1 for c in cols if _RE_MES.match(c))
    if meses >= 3:
        return (None, cols)

    # Header como linha de dados
    for i in range(min(12, len(df))):
        linha = [_norm(v) for v in df.iloc[i].tolist()]
        meses = sum(1 for c in linha if _RE_MES.match(c))
        if meses >= 3:
            return (i, linha)
    return None


def _eh_col_rubrica(c: str) -> bool:
    cl = c.lower().strip()
    return 'rubrica' in cl or cl in ('despesas', 'despesa', 'despesas:')


def detectar_balanual(df_raw: pd.DataFrame) -> bool:
    """True se a planilha for um export no formato Condomínio21/BalAnual."""
    if df_raw is None or df_raw.empty:
        return False
    return _localizar_header(df_raw) is not None


def _categorizar(nome: str) -> str:
    n = nome.lower()
    for categoria, palavras in CATEGORIAS:
        if any(p in n for p in palavras):
            return categoria
    return CATEGORIA_OUTROS


def processar_balanual(df_raw: pd.DataFrame,
                       num_unidades: int = None,
                       taxa_atual: float = None) -> dict:
    """
    Processa o export BalAnual e devolve o dicionário analysis_results
    com chaves compatíveis com o legado + bloco 'balanual' com os dados ricos.
    Parâmetros opcionais sobrepõem o que for detectado na planilha.
    """
    loc = _localizar_header(df_raw)
    if loc is None:
        return None
    hdr, header = loc
    if header is None:
        header = [_norm(v) for v in df_raw.iloc[hdr].tolist()]
    primeira_linha_dados = 0 if hdr is None else hdr + 1

    # ── Mapeamento de colunas ────────────────────────────────────────────
    col_rubrica = col_prev_apr = col_total = col_media = None
    col_prev_next = col_indice = None
    cols_meses = []   # [(idx, 'MM/AAAA')]
    ano_atual = ano_proximo = None

    for j, h in enumerate(header):
        hl = h.lower()
        if col_rubrica is None and _eh_col_rubrica(h):
            col_rubrica = j
        elif _RE_MES.match(h):
            cols_meses.append((j, h))
        elif 'previs' in hl and 'aprovada' in hl:
            col_prev_apr = j
        elif hl == 'total':
            col_total = j
        elif 'média' in hl or 'media' in hl or 'realizado' in hl:
            col_media = j
        elif 'índice' in hl or 'indice' in hl:
            col_indice = j
        elif 'previs' in hl:
            m = re.search(r'(20\d{2})', h)
            if m and (ano_atual is None or int(m.group(1)) != ano_atual):
                col_prev_next = j
                ano_proximo = int(m.group(1))

    # Ano do exercício = ano mais frequente entre as colunas mensais
    # (o export pode cruzar o ano: 02/2025..01/2026)
    if cols_meses:
        anos = [int(r.split('/')[1]) for _, r in cols_meses]
        ano_atual = max(set(anos), key=anos.count)
    if col_rubrica is None:
        # fallback: primeira coluna que não é mensal nem numérica de header
        indices_meses = {j for j, _ in cols_meses}
        for j in range(len(header)):
            if j not in indices_meses:
                col_rubrica = j
                break
    if ano_atual and not ano_proximo:
        ano_proximo = ano_atual + 1
    cols_meses.sort(key=lambda t: (int(t[1].split('/')[1]),
                                   int(t[1].split('/')[0])))

    # ── Itens (até a linha TOTAL) ────────────────────────────────────────
    itens, linha_total = [], None
    for i in range(primeira_linha_dados, len(df_raw)):
        nome = _norm(df_raw.iloc[i, col_rubrica])
        if not nome:
            continue
        if nome.strip().upper().rstrip(':') == 'TOTAL':
            linha_total = i
            break
        total_anual = _num(df_raw.iloc[i, col_total]) if col_total is not None else None
        if total_anual is None and cols_meses:
            soma_meses = [_num(df_raw.iloc[i, j]) for j, _ in cols_meses]
            soma_validos = [v for v in soma_meses if v is not None]
            if soma_validos:
                total_anual = sum(soma_validos)
        media = _num(df_raw.iloc[i, col_media]) if col_media is not None else None
        if media is None and total_anual is not None:
            media = total_anual / 12.0
        item = {
            'nome': nome,
            'prev_aprovada_mensal': _num(df_raw.iloc[i, col_prev_apr]) if col_prev_apr is not None else None,
            'total_anual': total_anual,
            'media_mensal': media,
            'prev_next_mensal': _num(df_raw.iloc[i, col_prev_next]) if col_prev_next is not None else None,
            'indice': _num(df_raw.iloc[i, col_indice]) if col_indice is not None else None,
            'categoria': _categorizar(nome),
        }
        itens.append(item)

    if not itens:
        return None

    # ── Totais (linha TOTAL da planilha; fallback: soma dos itens) ──────
    def _tot(col, campo):
        if linha_total is not None and col is not None:
            v = _num(df_raw.iloc[linha_total, col])
            if v is not None:
                return v
        vals = [it[campo] for it in itens if it[campo] is not None]
        return sum(vals) if vals else None

    prev_aprovada_mensal = _tot(col_prev_apr, 'prev_aprovada_mensal')
    realizado_total_anual = _tot(col_total, 'total_anual')
    realizado_medio_mensal = _tot(col_media, 'media_mensal')
    prev_next_mensal = _tot(col_prev_next, 'prev_next_mensal')

    serie_mensal = []
    for j, rotulo in cols_meses:
        v = _num(df_raw.iloc[linha_total, j]) if linha_total is not None else None
        if v is None:
            v = sum(_num(df_raw.iloc[i, j]) or 0.0
                    for i in range(primeira_linha_dados, linha_total or len(df_raw)))
        mm, yyyy = rotulo.split('/')
        mes_idx = int(mm) - 1
        rot = MESES_PT[mes_idx]
        if ano_atual and int(yyyy) != ano_atual:
            rot = f"{rot}/{yyyy[2:]}"
        serie_mensal.append({'rotulo': rot, 'valor': v})

    # ── Rodapé: índice global, unidades, fundo de reserva ───────────────
    rodape_ini = (linha_total + 1) if linha_total is not None else primeira_linha_dados
    texto_rodape = ' '.join(
        _norm(v) for i in range(rodape_ini, len(df_raw))
        for v in df_raw.iloc[i].tolist() if _norm(v)
    )

    indice_global = None
    if linha_total is not None:
        for i in range(rodape_ini, len(df_raw)):
            linha = [_norm(v) for v in df_raw.iloc[i].tolist()]
            for j, c in enumerate(linha):
                if 'reajuste' in c.lower():
                    for v in df_raw.iloc[i].tolist():
                        f = _num(v)
                        if f is not None and 0.5 < f < 2.0:
                            indice_global = f
                            break
            if indice_global:
                break

    unidades_detectadas = None
    m = re.search(r'/\s*(\d{1,5})\s*unidade', texto_rodape, re.IGNORECASE)
    if m:
        unidades_detectadas = int(m.group(1))

    fundo_pct = 5.0
    m = re.search(r'fundo\s+de\s+reserva[^%]{0,30}?(\d{1,2}(?:[.,]\d+)?)\s*%',
                  texto_rodape, re.IGNORECASE)
    if m:
        fundo_pct = float(m.group(1).replace(',', '.'))

    garantidora_pct = None
    m = re.search(r'garantidora[^%\d]{0,20}(\d{1,2}(?:[.,]\d+)?)\s*%',
                  texto_rodape, re.IGNORECASE)
    if m:
        garantidora_pct = float(m.group(1).replace(',', '.'))

    # Valores prontos do bloco "COMPOSIÇÃO DE RATEIO" (quando existem,
    # são mais confiáveis do que recalcular)
    def _valor_da_linha_rodape(*termos):
        for i in range(rodape_ini, len(df_raw)):
            rotulos = ' '.join(_norm(v) for v in df_raw.iloc[i].tolist()
                               if _norm(v) and _num(v) is None).upper()
            if all(t in rotulos for t in termos):
                for v in df_raw.iloc[i].tolist():
                    f = _num(v)
                    if f is not None and f > 0:
                        return f
        return None

    fundo_footer = _valor_da_linha_rodape('FUNDO DE RESERVA')
    garantidora_footer = _valor_da_linha_rodape('GARANTIDORA')
    rateado_footer = _valor_da_linha_rodape('TOTAL', 'DESPESAS', 'SER')

    def _valor_brl(regex):
        m = re.search(regex, texto_rodape, re.IGNORECASE)
        if not m:
            return None
        try:
            return float(m.group(1).replace('.', '').replace(',', '.'))
        except ValueError:
            return None

    # Tabela "COMPOSIÇÃO DA TAXA" por fração ideal (quando existir)
    fracoes = []
    hdr_fracao = None
    for i in range(rodape_ini, len(df_raw)):
        linha_txt = ' '.join(_norm(v) for v in df_raw.iloc[i].tolist()).upper()
        if 'FRAÇÃO IDEAL' in linha_txt and ('TAXA' in linha_txt or 'TOTAL' in linha_txt):
            hdr_fracao = i
            break
    if hdr_fracao is not None:
        for i in range(hdr_fracao + 1, min(hdr_fracao + 12, len(df_raw))):
            celulas = df_raw.iloc[i].tolist()
            nome_fr = _norm(celulas[0]) if len(celulas) else ''
            nums = [f for f in (_num(v) for v in celulas) if f is not None and f > 0]
            if not nome_fr or not nums:
                if fracoes:
                    break
                continue
            if 'REAJUSTE' in nome_fr.upper() or 'TAXA MÉDIA' in nome_fr.upper():
                break
            fracoes.append({'nome': nome_fr.rstrip(';,. '),
                            'total': max(nums)})

    # "TAXA MÉDIA ATUAL: R$ 304,88 - TAXA MÉDIA PROVISIONADA R$ 361,09"
    taxa_media_atual = _valor_brl(
        r'taxa\s+m[eé]dia\s+atual[:\s]*r?\$?\s*([\d.]+,\d{2})')
    taxa_media_provisionada = _valor_brl(
        r'provisionada[:\s]*r?\$?\s*([\d.]+,\d{2})')

    unidades = num_unidades or unidades_detectadas

    # ── Previsão do próximo ano e taxa ideal ─────────────────────────────
    if prev_next_mensal is None:
        base = realizado_medio_mensal or prev_aprovada_mensal
        prev_next_mensal = base * (indice_global or 1.0) if base else None

    def _plausivel(footer_val, estimativa):
        """Aceita valor do rodapé só se coerente com a estimativa calculada
        (evita capturar valores por-unidade de blocos de composição da taxa)."""
        if footer_val is None:
            return None
        if estimativa is None:
            return footer_val
        return footer_val if 0.5 * estimativa <= footer_val <= 2.0 * estimativa else None

    est_fundo = prev_next_mensal * fundo_pct / 100.0 if prev_next_mensal else None
    fundo_mensal = _plausivel(fundo_footer, est_fundo)
    if fundo_mensal is None and est_fundo:
        fundo_mensal = est_fundo

    est_gar = None
    if garantidora_pct and prev_next_mensal:
        est_gar = ((prev_next_mensal + (fundo_mensal or 0.0)) *
                   garantidora_pct / 100.0)
    garantidora_mensal = _plausivel(garantidora_footer, est_gar)
    if garantidora_mensal is None and est_gar:
        garantidora_mensal = est_gar

    est_rateado = None
    if prev_next_mensal:
        est_rateado = (prev_next_mensal + (fundo_mensal or 0.0) +
                       (garantidora_mensal or 0.0))
    total_rateado_mensal = _plausivel(rateado_footer, est_rateado)
    if total_rateado_mensal is None and est_rateado:
        total_rateado_mensal = est_rateado

    # Taxa ideal: rateio/unidades quando há unidades; senão, a taxa média
    # provisionada informada pela própria planilha (condomínios com rateio
    # por fração ideal, sem nº único de unidades)
    taxa_ideal_mensal = None
    if total_rateado_mensal and unidades:
        taxa_ideal_mensal = total_rateado_mensal / unidades
    elif taxa_media_provisionada:
        taxa_ideal_mensal = taxa_media_provisionada

    # Taxa aprovada do ano atual, derivada da própria planilha:
    # previsão aprovada mensal × (1 + fundo%) ÷ unidades — a mesma fórmula
    # aplicada ao orçamento aprovado na assembleia anterior.
    taxa_aprovada_atual = None
    if prev_aprovada_mensal and unidades:
        taxa_aprovada_atual = (prev_aprovada_mensal *
                               (1 + fundo_pct / 100.0)) / unidades

    # Baseline do reajuste, por prioridade:
    # 1) informada pelo operador  2) taxa média atual da planilha
    # 3) derivada da previsão aprovada
    taxa_atual_origem = None
    if taxa_atual:
        taxa_atual_origem = 'informada'
    elif taxa_media_atual:
        taxa_atual = taxa_media_atual
        taxa_atual_origem = 'planilha'
    elif taxa_aprovada_atual:
        taxa_atual = taxa_aprovada_atual
        taxa_atual_origem = 'aprovada'

    reajuste_pct = None
    if taxa_ideal_mensal and taxa_atual:
        reajuste_pct = (taxa_ideal_mensal / taxa_atual - 1.0) * 100.0

    desvio_pct = None
    if realizado_medio_mensal and prev_aprovada_mensal:
        desvio_pct = (realizado_medio_mensal / prev_aprovada_mensal - 1.0) * 100.0

    # ── Categorias (para donut e grupos legado) ──────────────────────────
    cat_tot = {}
    for it in itens:
        v = it['media_mensal'] or 0.0
        cat_tot[it['categoria']] = cat_tot.get(it['categoria'], 0.0) + v
    total_cat = sum(cat_tot.values()) or 1.0
    ordem = [c for c, _ in CATEGORIAS] + [CATEGORIA_OUTROS]
    categorias = [
        {'nome': c, 'mensal': cat_tot[c], 'pct': cat_tot[c] / total_cat * 100.0}
        for c in ordem if cat_tot.get(c)
    ]
    categorias.sort(key=lambda x: -x['mensal'])

    top_itens = sorted(
        (it for it in itens if it['media_mensal']),
        key=lambda x: -x['media_mensal']
    )

    # Grupos no formato legado (mantém /api/analisar e o frontend funcionando)
    grupos_legado = [{
        'numero': n + 1,
        'nome': c['nome'],
        'total': c['mensal'] * 12.0,
        'percentual': c['pct'],
        'itens': [
            {'nome': it['nome'], 'total': (it['media_mensal'] or 0.0) * 12.0}
            for it in itens if it['categoria'] == c['nome']
        ],
    } for n, c in enumerate(categorias)]

    anual = lambda v: v * 12.0 if v is not None else None

    return {
        # ── chaves compatíveis com o legado ──
        'formato': 'balanual',
        'nome_condominio': None,
        'data_assembleia': None,
        'ano_atual': ano_atual,
        'ano_proximo': ano_proximo,
        'num_unidades': unidades,
        'grupos': grupos_legado,
        'total_despesas': anual(prev_next_mensal),
        'fundo_reserva': anual(fundo_mensal),
        'fundo_reserva_pct': fundo_pct,
        'garantidora': anual(garantidora_mensal),
        'garantidora_pct': garantidora_pct,
        'total_rateado': anual(total_rateado_mensal),
        'taxa_ideal_mensal': taxa_ideal_mensal,
        'taxa_atual': taxa_atual,
        'reajuste_pct': reajuste_pct,
        # ── bloco rico do formato BalAnual ──
        'balanual': {
            'taxa_atual_origem': taxa_atual_origem,
            'taxa_aprovada_atual': taxa_aprovada_atual,
            'taxa_media_atual': taxa_media_atual,
            'taxa_media_provisionada': taxa_media_provisionada,
            'garantidora_mensal': garantidora_mensal,
            'garantidora_pct': garantidora_pct,
            'prev_aprovada_mensal': prev_aprovada_mensal,
            'realizado_medio_mensal': realizado_medio_mensal,
            'realizado_total_anual': realizado_total_anual,
            'desvio_pct': desvio_pct,
            'prev_next_mensal': prev_next_mensal,
            'indice_global': indice_global,
            'fundo_mensal': fundo_mensal,
            'total_rateado_mensal': total_rateado_mensal,
            'serie_mensal': serie_mensal,
            'categorias': categorias,
            'top_itens': [
                {'nome': it['nome'], 'media_mensal': it['media_mensal'],
                 'categoria': it['categoria'], 'indice': it['indice'],
                 'prev_next_mensal': it['prev_next_mensal']}
                for it in top_itens
            ],
            'fracoes': fracoes,
            'num_itens': len(itens),
            'unidades_detectadas': unidades_detectadas,
        },
    }
