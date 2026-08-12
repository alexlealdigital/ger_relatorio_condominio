"""
Gerador de PDF v1.0 — Gestor Financeiro de Condomínios
Desenha o relatório em PDF nativo (reportlab), espelhando o deck aprovado:
mesma ordem de páginas, fonte serifada, caixa alta, logo à esquerda,
capa com círculo pêssego e logo à direita.

Página 16:9 (960×540 pt) para projeção — idêntica à proporção do PPTX.
Interface: PDFGenerator(analyzer, output_path, extras=None).generate()
"""

import io
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor
from reportlab.lib.utils import ImageReader

# ─── Paleta ──────────────────────────────────────────────────────────────
NAVY   = HexColor('#0A1F44')
BLUE   = HexColor('#1A56DB')
ICE    = HexColor('#EBF0FB')
ICE2   = HexColor('#F5F7FC')
PEACH  = HexColor('#FBE3CC')
GREEN  = HexColor('#0E9F6E')
RED    = HexColor('#E8503A')
TEXT   = HexColor('#1B263B')
MUTED  = HexColor('#5A6B8C')
WHITE  = HexColor('#FFFFFF')
MPL_SEQ = ['#1A56DB', '#6E9BFF', '#0E9F6E', '#F5A623', '#8A63D2', '#E8503A',
           '#2FA8C9', '#B0B9CC']

PW, PH = 960.0, 540.0          # 16:9 em pontos (13,33" × 7,5")
IN = 72.0                       # pontos por polegada
SERIF, SERIF_B = 'Times-Roman', 'Times-Bold'


def _brl(v, dec=2):
    if v is None:
        return '—'
    s = f'{v:,.{dec}f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
    return f'R$ {s}'


def _pct(v, dec=1, sinal=False):
    if v is None:
        return '—'
    s = f'{v:+.{dec}f}' if sinal else f'{v:.{dec}f}'
    return s.replace('.', ',') + '%'


class PDFGenerator:
    def __init__(self, analyzer, output_path, extras=None):
        self.r = analyzer.analysis_results
        self.b = self.r.get('balanual') or {}
        self.extras = extras or {}
        self.c = canvas.Canvas(output_path, pagesize=(PW, PH))
        self._logo_reader = None
        path = self.extras.get('logo_path')
        if path:
            try:
                self._logo_reader = ImageReader(path)
            except Exception:
                pass

    # ── primitivas ──────────────────────────────────────────────────────
    def _t(self, x_in, y_in, texto, size=14, color=TEXT, bold=False,
           align='left', max_w_in=None):
        """Texto em coordenadas de polegadas a partir do topo. CAIXA ALTA."""
        texto = str(texto).upper()
        fonte = SERIF_B if bold else SERIF
        self.c.setFont(fonte, size)
        self.c.setFillColor(color)
        x, y = x_in * IN, PH - y_in * IN - size
        if max_w_in:
            while texto and self.c.stringWidth(texto, fonte, size) > max_w_in * IN:
                texto = texto[:-1]
        if align == 'right':
            self.c.drawRightString(x, y, texto)
        elif align == 'center':
            self.c.drawCentredString(x, y, texto)
        else:
            self.c.drawString(x, y, texto)

    def _wrap(self, x_in, y_in, w_in, texto, size=13, color=TEXT, bold=False,
              leading=1.45):
        """Parágrafo com quebra automática. Retorna o y (in) após o texto."""
        fonte = SERIF_B if bold else SERIF
        self.c.setFont(fonte, size)
        self.c.setFillColor(color)
        palavras = str(texto).upper().split()
        linhas, atual = [], ''
        for p in palavras:
            teste = f'{atual} {p}'.strip()
            if self.c.stringWidth(teste, fonte, size) <= w_in * IN:
                atual = teste
            else:
                linhas.append(atual)
                atual = p
        if atual:
            linhas.append(atual)
        y = y_in
        for linha in linhas:
            self.c.drawString(x_in * IN, PH - y * IN - size, linha)
            y += size * leading / IN
        return y

    def _card(self, x_in, y_in, w_in, h_in, fill=ICE2, radius=8):
        self.c.setFillColor(fill)
        self.c.roundRect(x_in * IN, PH - (y_in + h_in) * IN, w_in * IN,
                         h_in * IN, radius, stroke=0, fill=1)

    def _chart(self, fig, x_in, y_in, w_in):
        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=170, transparent=True,
                    bbox_inches='tight')
        plt.close(fig)
        buf.seek(0)
        img = ImageReader(buf)
        iw, ih = img.getSize()
        h_in = w_in * ih / iw
        self.c.drawImage(img, x_in * IN, PH - (y_in + h_in) * IN,
                         w_in * IN, h_in * IN, mask='auto')
        return h_in

    def _logo(self, x_in, y_in, alt_in, direita=False):
        if not self._logo_reader:
            return 0
        iw, ih = self._logo_reader.getSize()
        w_in = alt_in * iw / ih
        x = (PW - (x_in + w_in) * IN) if direita else x_in * IN
        self.c.drawImage(self._logo_reader, x, PH - (y_in + alt_in) * IN,
                         w_in * IN, alt_in * IN, mask='auto')
        return w_in

    def _titulo(self, texto, sub=None):
        self._logo(0.6, 0.38, 0.8, direita=True)
        self._t(0.6, 0.45, texto, size=26, color=NAVY, bold=True)
        if sub:
            self._t(0.6, 0.98, sub, size=12, color=MUTED)

    def _rodape(self):
        nome = self.r.get('nome_condominio') or 'PREVISÃO ORÇAMENTÁRIA'
        self._t(0.6, 7.18, nome, size=8, color=HexColor('#B6C0D4'))

    def _nova(self):
        self.c.showPage()

    @staticmethod
    def _mpl_base(ax):
        for sp in ('top', 'right', 'left'):
            ax.spines[sp].set_visible(False)
        ax.spines['bottom'].set_color('#D5DCEA')
        ax.tick_params(colors='#5A6B8C', labelsize=11, length=0)
        ax.yaxis.grid(True, color='#E4E9F4', linewidth=0.9)
        ax.set_axisbelow(True)

    # ── páginas ─────────────────────────────────────────────────────────
    def _pg_capa(self):
        self.c.setFillColor(PEACH)
        self.c.circle(13.3 * IN, PH - 1.2 * IN, 3.1 * IN, stroke=0, fill=1)
        self._t(0.9, 0.85, 'ASSEMBLEIA GERAL ORDINÁRIA', size=15, color=NAVY,
                bold=True)
        data = self.r.get('data_assembleia')
        if data:
            self._t(0.9, 1.32, f'DATA: {data}', size=11.5, color=TEXT, bold=True)
        self._logo(0.75, 0.62, 1.0, direita=True)
        ano = self.r.get('ano_proximo') or ''
        self._t(0.9, 2.75, 'PREVISÃO', size=46, color=NAVY, bold=True)
        self._t(0.9, 3.55, f'ORÇAMENTÁRIA {ano}', size=46, color=NAVY, bold=True)
        nome = self.r.get('nome_condominio') or 'CONDOMÍNIO'
        self._t(0.9, 5.2, nome, size=18, color=TEXT, bold=True)
        self._nova()

    def _pg_edital(self):
        texto = (self.extras.get('edital_texto') or '').strip()
        if not texto:
            return
        paragrafos = [p.strip() for p in texto.split('\n') if p.strip()]
        blocos, atual = [], ''
        for p in paragrafos:
            if atual and len(atual) + len(p) > 2100:
                blocos.append(atual)
                atual = p
            else:
                atual = f'{atual}\n{p}' if atual else p
        if atual:
            blocos.append(atual)
        for i, bloco in enumerate(blocos[:3]):
            suf = f' ({i + 1}/{len(blocos)})' if len(blocos) > 1 else ''
            self._titulo(f'EDITAL DE CONVOCAÇÃO{suf}')
            y = 1.7
            for par in bloco.split('\n'):
                y = self._wrap(0.9, y, 11.5, par, size=13.5) + 0.12
            self._rodape()
            self._nova()

    def _pg_detalhamento(self):
        itens = self.b.get('top_itens') or []
        if not itens:
            return
        ano = self.r.get('ano_atual') or ''
        por_pagina = 22
        paginas = [itens[i:i + por_pagina]
                   for i in range(0, len(itens), por_pagina)]
        for pg, bloco in enumerate(paginas, 1):
            self._titulo('DETALHAMENTO DAS RUBRICAS',
                         f'Despesa média mensal de {ano} por rubrica '
                         f'(página {pg} de {len(paginas)})')
            col_w, metade = 5.95, (len(bloco) + 1) // 2
            for cidx, coluna in enumerate((bloco[:metade], bloco[metade:])):
                x = 0.6 + cidx * (col_w + 0.25)
                y = 1.65
                for i, it in enumerate(coluna):
                    if i % 2 == 0:
                        self._card(x, y - 0.05, col_w, 0.4, fill=ICE2, radius=4)
                    self._t(x + 0.15, y + 0.04, it['nome'][:40], size=10,
                            color=TEXT, max_w_in=4.0)
                    self._t(x + col_w - 0.15, y + 0.04,
                            _brl(it['media_mensal'], 0), size=10, color=MUTED,
                            align='right')
                    y += 0.46
            self._rodape()
            self._nova()

    def _pg_top(self):
        top = (self.b.get('top_itens') or [])[:8]
        if not top:
            return
        ano = self.r.get('ano_atual') or ''
        self._titulo('MAIORES DESPESAS DO CONDOMÍNIO',
                     f'As 8 rubricas de maior peso na média mensal de {ano}')
        nomes = [t['nome'][:36].upper() for t in reversed(top)]
        vals = [t['media_mensal'] / 1000.0 for t in reversed(top)]
        fig, ax = plt.subplots(figsize=(11.4, 4.5))
        barras = ax.barh(nomes, vals, color='#1A56DB', height=0.62, zorder=3)
        barras[-1].set_color('#0A1F44')
        for sp in ax.spines.values():
            sp.set_visible(False)
        ax.tick_params(colors='#1B263B', labelsize=11, length=0)
        ax.xaxis.grid(True, color='#E4E9F4', linewidth=0.9)
        ax.set_axisbelow(True)
        ax.set_xlabel('R$ MIL / MÊS', fontsize=10, color='#5A6B8C')
        for barra, t in zip(barras, reversed(top)):
            idx = t.get('indice')
            var = ((idx - 1) * 100 if idx and 0.5 < idx < 2.0 else
                   ((t['prev_next_mensal'] / t['media_mensal'] - 1) * 100
                    if t.get('prev_next_mensal') and t.get('media_mensal')
                    else None))
            et = _brl(t['media_mensal'], 0)
            if var is not None:
                et += f'  ·  {_pct(var, 1, True)}'
            ax.text(barra.get_width() + max(vals) * 0.012,
                    barra.get_y() + barra.get_height() / 2, et,
                    va='center', fontsize=9, color='#5A6B8C')
        fig.tight_layout()
        self._chart(fig, 0.55, 1.5, 12.2)
        self._rodape()
        self._nova()

    def _pg_evolucao(self):
        serie = self.b.get('serie_mensal') or []
        if not serie:
            return
        ano = self.r.get('ano_atual') or ''
        prev = self.b.get('prev_aprovada_mensal')
        sub = ('Barras: realizado no mês · Linha tracejada: previsão aprovada'
               if prev else 'Despesa realizada em cada mês do exercício')
        self._titulo(f'EVOLUÇÃO MENSAL DAS DESPESAS — {ano}', sub)
        rotulos = [p['rotulo'].upper() for p in serie]
        valores = [p['valor'] / 1000.0 for p in serie]
        fig, ax = plt.subplots(figsize=(11.4, 4.2))
        cores = ['#E8503A' if (prev and v * 1000 > prev) else '#1A56DB'
                 for v in valores]
        barras = ax.bar(rotulos, valores, color=cores, width=0.62, zorder=3)
        vmax = max(valores) if valores else 1
        for barra, p in zip(barras, serie):
            ax.text(barra.get_x() + barra.get_width() / 2,
                    barra.get_height() + vmax * 0.022,
                    f"{p['valor'] / 1000:,.1f}".replace(',', 'X')
                        .replace('.', ',').replace('X', '.'),
                    ha='center', va='bottom', fontsize=9,
                    color='#1B263B', fontweight='bold')
        ax.set_ylim(0, vmax * 1.14)
        if prev:
            ax.axhline(prev / 1000.0, color='#0A1F44', linewidth=1.6,
                       linestyle=(0, (6, 4)), zorder=4)
        self._mpl_base(ax)
        ax.set_ylabel('R$ MIL / MÊS', fontsize=10, color='#5A6B8C')
        fig.tight_layout()
        self._chart(fig, 0.55, 1.5, 12.2)
        self._rodape()
        self._nova()

    def _pg_categorias(self):
        cats = self.b.get('categorias') or []
        if not cats:
            return
        self._titulo('PARA ONDE VAI O DINHEIRO',
                     'Composição da despesa média mensal por categoria')
        total = sum(cat['mensal'] for cat in cats)
        fig, ax = plt.subplots(figsize=(4.9, 4.9))
        ax.pie([cat['mensal'] for cat in cats], colors=MPL_SEQ[:len(cats)],
               startangle=90, counterclock=False,
               wedgeprops={'width': 0.34, 'edgecolor': 'white', 'linewidth': 2})
        ax.text(0, 0.08, _brl(total, 0), ha='center', va='center',
                fontsize=15, fontweight='bold', color='#0A1F44')
        ax.text(0, -0.15, 'MÉDIA / MÊS', ha='center', va='center',
                fontsize=9, color='#5A6B8C')
        self._chart(fig, 0.7, 1.6, 4.6)
        y = 1.85
        for i, cat in enumerate(cats[:7]):
            self.c.setFillColor(HexColor(MPL_SEQ[i]))
            self.c.circle(6.05 * IN, PH - (y + 0.1) * IN, 6, stroke=0, fill=1)
            self._t(6.35, y, cat['nome'], size=13, color=TEXT, bold=True)
            self._t(12.7, y, f"{_brl(cat['mensal'], 0)}  ·  {_pct(cat['pct'])}",
                    size=12, color=MUTED, align='right')
            y += 0.62
        self._rodape()
        self._nova()

    def _kpi(self, x, y, w, h, rotulo, valor, sub=None, cor=TEXT, fill=ICE2):
        self._card(x, y, w, h, fill=fill)
        self._t(x + 0.22, y + 0.18, rotulo, size=10, color=MUTED, bold=True)
        self._t(x + 0.22, y + 0.5, valor, size=22, color=cor, bold=True)
        if sub:
            self._t(x + 0.22, y + h - 0.4, sub, size=9, color=MUTED,
                    max_w_in=w - 0.44)

    def _pg_panorama(self):
        b = self.b
        if not b:
            return
        ano = self.r.get('ano_atual') or ''
        tem_apr = b.get('prev_aprovada_mensal') is not None
        sub = ('Execução orçamentária do exercício — previsto × realizado'
               if tem_apr else 'Execução do exercício e projeção para o próximo ano')
        self._titulo(f'PANORAMA {ano}', sub)
        y, h, w, gap, x0 = 1.8, 1.7, 3.87, 0.28, 0.6
        if tem_apr:
            desvio = b.get('desvio_pct')
            self._kpi(x0, y, w, h, 'PREVISÃO APROVADA / MÊS',
                      _brl(b.get('prev_aprovada_mensal'), 2),
                      f'Orçamento aprovado para {ano}')
            self._kpi(x0 + w + gap, y, w, h, 'REALIZADO MÉDIO / MÊS',
                      _brl(b.get('realizado_medio_mensal'), 2),
                      f'Total no ano: {_brl(b.get("realizado_total_anual"), 2)}')
            self._kpi(x0 + 2 * (w + gap), y, w, h, 'DESVIO SOBRE O APROVADO',
                      _pct(desvio, 1, True),
                      'Realizado acima da previsão' if (desvio or 0) > 0
                      else 'Realizado dentro da previsão',
                      cor=RED if (desvio or 0) > 0 else GREEN, fill=ICE)
            msg = (f'As despesas reais superaram a previsão aprovada em '
                   f'{_pct(desvio)} na média mensal. Manter a taxa atual '
                   f'significa operar abaixo do custo real do condomínio.'
                   if desvio and desvio > 0 else
                   'As despesas reais se mantiveram dentro da previsão '
                   'aprovada no exercício.')
        else:
            reaj = self.r.get('reajuste_pct')
            self._kpi(x0, y, w, h, 'REALIZADO MÉDIO / MÊS',
                      _brl(b.get('realizado_medio_mensal'), 2),
                      f'Total no ano: {_brl(b.get("realizado_total_anual"), 2)}')
            self._kpi(x0 + w + gap, y, w, h,
                      f'PREVISÃO {self.r.get("ano_proximo") or ""} / MÊS',
                      _brl(b.get('prev_next_mensal'), 2),
                      'Despesas projetadas para o próximo exercício')
            self._kpi(x0 + 2 * (w + gap), y, w, h, 'REAJUSTE PROJETADO',
                      _pct(reaj, 2, True),
                      'Sobre o total provisionado no período anterior',
                      cor=RED if (reaj or 0) > 0 else GREEN, fill=ICE)
            msg = (f'A despesa média mensal apurada em {ano} é a base da '
                   f'previsão orçamentária do próximo exercício e do valor '
                   f'da nova taxa condominial.')
        self._card(0.6, 3.95, 12.13, 1.35, fill=NAVY)
        self._wrap(1.0, 4.25, 11.3, msg, size=14, color=WHITE, bold=True)
        self._rodape()
        self._nova()

    def _pg_previsao(self):
        b = self.b
        ano = self.r.get('ano_proximo') or ''
        self._titulo(f'PREVISÃO ORÇAMENTÁRIA {ano}',
                     'Composição do valor mensal a ratear entre as unidades')
        linhas = [(f'DESPESAS PREVISTAS {ano}', b.get('prev_next_mensal')),
                  (f'FUNDO DE RESERVA ({_pct(self.r.get("fundo_reserva_pct"), 0)})',
                   b.get('fundo_mensal'))]
        if b.get('garantidora_mensal'):
            linhas.append((f'GARANTIDORA ({_pct(b.get("garantidora_pct"), 1)})',
                           b['garantidora_mensal']))
        y = 1.8
        for rot, val in linhas:
            self._card(0.6, y, 7.6, 0.72, fill=ICE2)
            self._t(0.95, y + 0.22, rot, size=13, color=TEXT, bold=True)
            self._t(7.85, y + 0.2, _brl(val, 2), size=15, color=NAVY,
                    bold=True, align='right')
            y += 0.88
        self._card(0.6, y + 0.06, 7.6, 0.85, fill=NAVY)
        self._t(0.95, y + 0.3, 'TOTAL A RATEAR / MÊS', size=12,
                color=HexColor('#6E9BFF'), bold=True)
        self._t(7.85, y + 0.24, _brl(b.get('total_rateado_mensal'), 2),
                size=17, color=WHITE, bold=True, align='right')

        uni = self.r.get('num_unidades')
        fracoes = b.get('fracoes') or []
        por_fr = not uni and b.get('taxa_media_provisionada')
        self._card(8.6, 1.8, 4.13, 4.6 if fracoes else 3.6, fill=ICE)
        self._t(8.95, 2.1, 'RATEIO POR FRAÇÃO IDEAL' if por_fr
                else 'RATEIO POR UNIDADE', size=11, color=MUTED, bold=True)
        self._t(8.95, 2.5, 'CONFORME BLOCOS' if por_fr
                else f'{uni or "—"} UNIDADES', size=15, color=TEXT, bold=True)
        self._t(8.95, 3.2, 'TAXA MÉDIA RESULTANTE' if por_fr
                else 'TAXA CONDOMINIAL RESULTANTE', size=11, color=MUTED)
        self._t(8.95, 3.55, _brl(self.r.get('taxa_ideal_mensal'), 2),
                size=26, color=BLUE, bold=True)
        if fracoes:
            yf = 4.5
            for fr in fracoes[:4]:
                self._t(8.95, yf, fr['nome'][:22], size=8.5, color=MUTED,
                        max_w_in=2.4)
                self._t(12.5, yf, _brl(fr['total'], 2), size=9, color=NAVY,
                        bold=True, align='right')
                yf += 0.4
        self._rodape()
        self._nova()

    def _pg_aprovacao(self):
        taxa = self.r.get('taxa_ideal_mensal')
        ano = self.r.get('ano_proximo') or ''
        fracoes = (self.b.get('fracoes') or [])
        self.c.setFillColor(NAVY)
        self.c.rect(0, PH - 0.16 * IN, PW, 0.16 * IN, stroke=0, fill=1)
        self._logo(0.6, 0.55, 1.15, direita=True)
        self._t(0.9, 1.95, 'PROPOSTA PARA APROVAÇÃO EM ASSEMBLEIA', size=14,
                color=BLUE, bold=True)
        if taxa:
            rot = 'TAXA MÉDIA CONDOMINIAL' if fracoes else 'TAXA CONDOMINIAL'
            self._t(0.9, 2.5, f'{rot} {ano}:', size=27, color=NAVY, bold=True)
            self._t(7.2, 2.42, _brl(taxa, 2), size=36, color=BLUE, bold=True)
            self._t(10.2, 2.72, '/ UNIDADE / MÊS', size=13, color=MUTED)
        if fracoes:
            self._t(0.9, 3.6, 'VALORES POR FRAÇÃO IDEAL', size=11,
                    color=MUTED, bold=True)
            y = 3.95
            for fr in fracoes[:5]:
                self._card(0.9, y, 11.5, 0.52, fill=ICE2, radius=6)
                self._t(1.2, y + 0.13, fr['nome'][:68], size=11, color=TEXT,
                        max_w_in=8.5)
                self._t(12.1, y + 0.1, _brl(fr['total'], 2), size=12.5,
                        color=NAVY, bold=True, align='right')
                y += 0.62
        nome = self.r.get('nome_condominio')
        data = self.r.get('data_assembleia')
        rod = ' · '.join(v for v in
                         [nome, f'Assembleia em {data}' if data else None] if v)
        if rod:
            self._t(0.9, 6.85, rod, size=11, color=MUTED)
        self._nova()

    def _pg_proposta(self):
        taxa_nova = self.r.get('taxa_ideal_mensal')
        if not taxa_nova:
            return
        taxa_atual = self.r.get('taxa_atual')
        reajuste = self.r.get('reajuste_pct')
        self._titulo('PROPOSTA DE REAJUSTE DA TAXA CONDOMINIAL',
                     'Valor necessário para equilibrar as contas do próximo exercício')
        if taxa_atual:
            og = self.b.get('taxa_atual_origem')
            rot = ('TAXA MÉDIA ATUAL' if og == 'planilha' else
                   (f'TAXA APROVADA {self.r.get("ano_atual") or ""}'
                    if og == 'aprovada' else 'TAXA ATUAL'))
            self._card(0.9, 2.1, 4.7, 2.4, fill=ICE2)
            self._t(1.25, 2.5, rot, size=11.5, color=MUTED, bold=True)
            self._t(1.25, 3.0, _brl(taxa_atual, 2), size=32, color=MUTED,
                    bold=True)
            self._t(1.25, 3.95, 'POR UNIDADE / MÊS', size=10.5, color=MUTED)
            # seta
            self.c.setFillColor(BLUE)
            ax0, ay = 5.95 * IN, PH - 3.35 * IN
            self.c.rect(ax0, ay - 8, 0.9 * IN, 16, stroke=0, fill=1)
            p = self.c.beginPath()
            p.moveTo(ax0 + 0.9 * IN, ay - 20)
            p.lineTo(ax0 + 1.25 * IN, ay)
            p.lineTo(ax0 + 0.9 * IN, ay + 20)
            p.close()
            self.c.drawPath(p, stroke=0, fill=1)
            self._card(7.45, 1.9, 5.0, 2.85, fill=NAVY)
            self._t(7.8, 2.28, 'TAXA PROPOSTA', size=11.5,
                    color=HexColor('#6E9BFF'), bold=True)
            self._t(7.8, 2.8, _brl(taxa_nova, 2), size=38, color=WHITE,
                    bold=True)
            self._t(7.8, 3.95, 'POR UNIDADE / MÊS', size=10.5,
                    color=HexColor('#AFC2E8'))
            if reajuste is not None:
                self._card(7.8, 4.3, 2.6, 0.46, fill=WHITE, radius=12)
                self._t(9.1, 4.4, f'REAJUSTE DE {_pct(abs(reajuste))}',
                        size=11.5, color=RED if reajuste > 0 else GREEN,
                        bold=True, align='center')
                dif = taxa_nova - taxa_atual
                self._t(0.9, 5.35,
                        f'Diferença de {_brl(abs(dif), 2)} por unidade/mês '
                        f'para cobrir o custo real de operação.',
                        size=13, color=TEXT)
        else:
            self._card(3.4, 2.1, 6.5, 2.8, fill=NAVY)
            self._t(6.65, 2.5, 'TAXA CONDOMINIAL PROPOSTA', size=12,
                    color=HexColor('#6E9BFF'), bold=True, align='center')
            self._t(6.65, 3.1, _brl(taxa_nova, 2), size=44, color=WHITE,
                    bold=True, align='center')
            self._t(6.65, 4.3, 'POR UNIDADE / MÊS', size=11,
                    color=HexColor('#AFC2E8'), align='center')
        self._rodape()
        self._nova()

    # ── geração ─────────────────────────────────────────────────────────
    def generate(self):
        self._pg_capa()
        self._pg_edital()
        self._pg_detalhamento()
        self._pg_top()
        self._pg_evolucao()
        self._pg_categorias()
        self._pg_panorama()
        self._pg_previsao()
        self._pg_aprovacao()
        self._pg_proposta()
        self.c.save()
