"""
Analisador Financeiro Inteligente de Condomínios v3.0
Lógica de engenharia reversa baseada nas apresentações reais do cliente.

Lógica extraída das apresentações:
- Despesas organizadas em GRUPOS numerados (1. Encargos Bancários, 2. Consumo, etc.)
- Cada grupo tem itens com: Total Anual, Valor Por Apto/Unidade, % do total
- Composição do Rateio: Total Despesas + Fundo de Reserva (5%) + Garantidora (3%)
- Taxa Ideal = Total Rateado / N_Unidades / 12 meses
- Reajuste = ((Taxa Nova / Taxa Atual) - 1) * 100
"""

import pandas as pd
import numpy as np
import re
import os
from typing import Dict, List, Tuple, Optional


class CondominiumFinancialAnalyzer:
    """
    Analisador financeiro inteligente.
    Detecta automaticamente a estrutura de qualquer planilha de previsão orçamentária.
    """

    def __init__(self, excel_path: str):
        self.excel_path = excel_path
        self.df_raw = None
        self.df_processed = None
        self.monthly_revenue_2025 = None
        self.monthly_revenue_2026 = None
        self.analysis_results = {}
        self._num_unidades = None
        self._taxa_atual = None
        self._ano_atual = None
        self._ano_proximo = None
        self.load_data()

    # ─────────────────────────────────────────────────────────────────────────
    # CARREGAMENTO
    # ─────────────────────────────────────────────────────────────────────────
    def load_data(self) -> bool:
        try:
            ext = os.path.splitext(self.excel_path)[1].lower()
            if ext == '.xls':
                self.df_raw = pd.read_excel(self.excel_path, sheet_name=0,
                                             engine='xlrd', header=None)
            else:
                self.df_raw = pd.read_excel(self.excel_path, sheet_name=0,
                                             engine='openpyxl', header=None)
            # Tenta com header=0 se parecer ter cabeçalho
            # Detecta automaticamente
            self._detect_header()
            print(f"✓ Dados carregados: {len(self.df_raw)} linhas, {len(self.df_raw.columns)} colunas")
            return True
        except Exception as e:
            print(f"✗ Erro ao carregar: {e}")
            return False

    def _detect_header(self):
        """Tenta recarregar com header se a primeira linha parecer cabeçalho."""
        try:
            ext = os.path.splitext(self.excel_path)[1].lower()
            engine = 'xlrd' if ext == '.xls' else 'openpyxl'
            df_h = pd.read_excel(self.excel_path, sheet_name=0,
                                  engine=engine, header=0)
            # Verifica se alguma coluna contém "Rubrica" ou "Despesa"
            cols = [str(c).lower() for c in df_h.columns]
            if any('rubrica' in c or 'despesa' in c or 'item' in c for c in cols):
                self.df_raw = df_h
                self._has_header = True
                return
        except Exception:
            pass
        self._has_header = False

    # ─────────────────────────────────────────────────────────────────────────
    # DETECÇÃO INTELIGENTE DE ANOS
    # ─────────────────────────────────────────────────────────────────────────
    def _detect_years(self):
        """Detecta os anos presentes na planilha."""
        years = set()
        df_str = self.df_raw.astype(str)
        for col in df_str.columns:
            for val in df_str[col]:
                s = str(val)
                if s in ('nan', 'None', ''): continue
                matches = re.findall(r'\b(20\d{2})\b', s)
                years.update(int(m) for m in matches)
        # Detecta por colunas com padrão MM/AAAA
        for col in df_str.columns:
            for val in df_str[col]:
                s = str(val)
                m = re.search(r'\d{2}/(20\d{2})', s)
                if m:
                    years.add(int(m.group(1)))
        years = sorted(years)
        if len(years) >= 2:
            self._ano_atual   = years[-2]
            self._ano_proximo = years[-1]
        elif len(years) == 1:
            self._ano_atual   = years[0]
            self._ano_proximo = years[0] + 1
        else:
            import datetime
            y = datetime.datetime.now().year
            self._ano_atual   = y - 1
            self._ano_proximo = y
        print(f"✓ Anos detectados: {self._ano_atual} (atual) → {self._ano_proximo} (próximo)")

    # ─────────────────────────────────────────────────────────────────────────
    # DETECÇÃO DO NÚMERO DE UNIDADES E TAXA ATUAL
    # ─────────────────────────────────────────────────────────────────────────
    def _detect_units_and_current_rate(self):
        """
        Detecta número de unidades e taxa atual de qualquer planilha.
        Procura padrões como:
        - "202 unidades", "/ 202 Unidades", "246 Unidades"
        - "Taxa atual: R$ 508,54", "Valor atual R$ 508,54"
        - "Rateio das Despesas / 202 Unidades"
        """
        df_str = self.df_raw.astype(str)
        all_text = ' '.join(str(x) for x in df_str.values.flatten())

        # Número de unidades
        patterns_units = [
            r'/\s*(\d{2,4})\s*[Uu]nidades',
            r'(\d{2,4})\s*[Uu]nidades',
            r'(\d{2,4})\s*[Aa]ptos',
            r'(\d{2,4})\s*[Aa]partamentos',
        ]
        for pat in patterns_units:
            m = re.search(pat, all_text)
            if m:
                self._num_unidades = int(m.group(1))
                break
        if not self._num_unidades:
            self._num_unidades = 202  # fallback do Portal das Dunas

        # Taxa atual
        patterns_taxa = [
            r'[Vv]alor\s+[Aa]tual[,\s]+de\s+R\$\s*([\d.,]+)',
            r'[Tt]axa\s+[Aa]tual[:\s]+R\$\s*([\d.,]+)',
            r'[Rr]ela[çc][aã]o\s+ao\s+[Vv]alor\s+[Aa]tual.*?R\$\s*([\d.,]+)',
            r'[Tt]axa\s+[Vv]igente[:\s]+R\$\s*([\d.,]+)',
        ]
        for pat in patterns_taxa:
            m = re.search(pat, all_text)
            if m:
                self._taxa_atual = self._parse_currency(m.group(1))
                break

        print(f"✓ Unidades: {self._num_unidades} | Taxa atual: R$ {self._taxa_atual or 'não encontrada'}")

    # ─────────────────────────────────────────────────────────────────────────
    # DETECÇÃO DOS GRUPOS DE DESPESA
    # ─────────────────────────────────────────────────────────────────────────
    def _detect_expense_groups(self, df_str, col_rubrica, col_total, col_por_apto):
        """
        Detecta grupos de despesa no estilo do cliente:
        "1. Encargos Bancários", "2. Consumo", etc.
        Retorna lista de grupos com seus itens.
        """
        grupos = []
        grupo_atual = None

        # Padrão de grupo: número + ponto + texto
        pat_grupo = re.compile(r'^(\d{1,2})\.\s+(.+)$')
        # Padrão de subitem: número.número + texto
        pat_item  = re.compile(r'^(\d{1,2}\.\d{1,2})\s*[-–]\s*(.+)$')
        # Padrão TOTAL do grupo
        pat_total = re.compile(r'total.*item\s*\d', re.IGNORECASE)

        for idx, row in df_str.iterrows():
            rubrica = str(row[col_rubrica]).strip() if col_rubrica is not None else ''
            if rubrica in ('nan', '', 'None'):
                continue

            # Verifica se é cabeçalho de grupo
            m_grupo = pat_grupo.match(rubrica)
            if m_grupo and not pat_item.match(rubrica):
                if grupo_atual:
                    grupos.append(grupo_atual)
                grupo_atual = {
                    'numero':   int(m_grupo.group(1)),
                    'nome':     m_grupo.group(2).strip(),
                    'itens':    [],
                    'total':    0.0,
                    'percentual': None,
                }
                continue

            # Linha de total do grupo
            if pat_total.match(rubrica) and grupo_atual:
                val = self._safe_float(row.get(col_total, 0))
                if val:
                    grupo_atual['total'] = val
                continue

            # Subitem do grupo
            if grupo_atual and (pat_item.match(rubrica) or rubrica):
                total_val    = self._safe_float(row.get(col_total, 0))
                por_apto_val = self._safe_float(row.get(col_por_apto, 0))
                if total_val and total_val > 0:
                    grupo_atual['itens'].append({
                        'nome':     rubrica,
                        'total':    total_val,
                        'por_apto': por_apto_val or 0.0,
                    })

        if grupo_atual:
            grupos.append(grupo_atual)

        # Recalcula totais de grupos sem total capturado
        for g in grupos:
            if g['total'] == 0 and g['itens']:
                g['total'] = sum(i['total'] for i in g['itens'])

        return grupos

    # ─────────────────────────────────────────────────────────────────────────
    # COLUNA FINDER INTELIGENTE
    # ─────────────────────────────────────────────────────────────────────────
    def _find_column(self, df, keywords: List[str], exclude: List[str] = None) -> Optional[int]:
        """Encontra coluna por palavras-chave no cabeçalho ou nas primeiras linhas."""
        exclude = exclude or []
        df_str = df.astype(str)

        # Verifica cabeçalhos
        for col in df_str.columns:
            col_str = str(col).lower()
            if any(kw.lower() in col_str for kw in keywords):
                if not any(ex.lower() in col_str for ex in exclude):
                    return col

        # Verifica primeiras 5 linhas
        for row_idx in range(min(5, len(df_str))):
            for col in df_str.columns:
                val = str(df_str.iloc[row_idx][col]).lower()
                if any(kw.lower() in val for kw in keywords):
                    if not any(ex.lower() in val for ex in exclude):
                        # Retorna coluna seguinte (valor fica na col à direita)
                        cols = list(df_str.columns)
                        idx  = cols.index(col)
                        if idx + 1 < len(cols):
                            return cols[idx + 1]

        return None

    # ─────────────────────────────────────────────────────────────────────────
    # DETECÇÃO DO FUNDO DE RESERVA E GARANTIDORA
    # ─────────────────────────────────────────────────────────────────────────
    def _detect_reserve_fund(self, all_text: str):
        """Detecta percentuais de Fundo de Reserva e Garantidora."""
        fundo_pct = 5.0   # padrão
        garantidora_pct = 3.0  # padrão

        m = re.search(r'[Ff]undo\s+de\s+[Rr]eserva\s*\(?[Bb][Cc]?\s*(\d+(?:[.,]\d+)?)\s*%', all_text)
        if m:
            fundo_pct = float(m.group(1).replace(',', '.'))

        m = re.search(r'[Gg]arantidora\s*\(?[Bb][Cc]?\s*(\d+(?:[.,]\d+)?)\s*%', all_text)
        if m:
            garantidora_pct = float(m.group(1).replace(',', '.'))

        return fundo_pct, garantidora_pct

    # ─────────────────────────────────────────────────────────────────────────
    # UTILITÁRIOS
    # ─────────────────────────────────────────────────────────────────────────
    def _parse_currency(self, s: str) -> float:
        if not s:
            return 0.0
        s = str(s).replace('R$', '').replace(' ', '').strip()
        # Formato BR: 1.234,56
        if re.match(r'^\d{1,3}(\.\d{3})*,\d{2}$', s):
            s = s.replace('.', '').replace(',', '.')
        elif ',' in s and '.' in s:
            s = s.replace('.', '').replace(',', '.')
        elif ',' in s:
            s = s.replace(',', '.')
        try:
            return float(s)
        except Exception:
            return 0.0

    def _safe_float(self, val) -> float:
        if val is None or str(val).strip() in ('', 'nan', 'None', '-'):
            return 0.0
        if isinstance(val, (int, float)):
            return float(val)
        return self._parse_currency(str(val))

    # ─────────────────────────────────────────────────────────────────────────
    # PIPELINE PRINCIPAL
    # ─────────────────────────────────────────────────────────────────────────
    def extract_monthly_revenue(self) -> bool:
        """Compatibilidade com a interface da GUI."""
        return True

    def process_data(self) -> bool:
        """
        Processamento inteligente:
        0. Se for export Condomínio21 (BalAnual), usa o parser dedicado
        1. Detecta anos
        2. Detecta unidades e taxa atual
        3. Encontra colunas relevantes
        4. Extrai grupos de despesa
        5. Calcula rateio, taxa ideal e reajuste
        """
        try:
            # ── Formato real do cliente: export Condomínio21/BalAnual ──
            # Procura o formato na aba carregada e, se necessário, nas demais
            # abas do arquivo (modelos com aba de instruções, capas etc.)
            from balanual_parser import detectar_balanual, processar_balanual
            candidatos = [self.df_raw]
            try:
                import pandas as _pd
                xl = _pd.ExcelFile(self.excel_path)
                for aba in xl.sheet_names:
                    df_aba = _pd.read_excel(self.excel_path, sheet_name=aba,
                                            header=None)
                    candidatos.append(df_aba)
            except Exception:
                pass
            for df_cand in candidatos:
                if not detectar_balanual(df_cand):
                    continue
                resultado = processar_balanual(
                    df_cand,
                    num_unidades=self._num_unidades,
                    taxa_atual=self._taxa_atual,
                )
                if resultado:
                    self.analysis_results = resultado
                    print(f"✓ Formato BalAnual detectado: "
                          f"{resultado['balanual']['num_itens']} rubricas, "
                          f"{resultado['ano_atual']} → {resultado['ano_proximo']}")
                    return True

            self._detect_years()
            self._detect_units_and_current_rate()

            df = self.df_raw.copy()
            df_str = df.astype(str)
            all_text = ' '.join(str(x) for x in df_str.values.flatten())

            # ── Detecta coluna de Rubrica ────────────────────────────────────
            col_rubrica = self._find_column(
                df, ['rubrica', 'despesa', 'item', 'descri'],
                exclude=['previsto', 'realizado', 'total', 'media', 'previsao']
            )
            if col_rubrica is None:
                # Usa primeira coluna não-numérica
                for col in df.columns:
                    sample = df[col].dropna().head(10)
                    if sample.dtype == object or any(isinstance(v, str) for v in sample):
                        col_rubrica = col
                        break
            if col_rubrica is None:
                col_rubrica = df.columns[0]

            # ── Detecta coluna de Total ──────────────────────────────────────
            # Procura por "Previsão XXXX" ou "Total" ou "Previsão Aprovada"
            col_total_prev_atual = self._find_col_by_year(df, self._ano_atual, 'previsto')
            col_total_realizado  = self._find_col_by_year(df, self._ano_atual, 'realizado')
            col_total_2026       = self._find_col_by_year(df, self._ano_proximo, None)

            # Colunas mensais (MM/AAAA)
            col_meses = self._find_monthly_columns(df)

            # ── Coluna de "Por Apto" ─────────────────────────────────────────
            col_por_apto = self._find_column(
                df, ['por apto', 'por unidade', 'unit', 'apto', 'unid'],
                exclude=['total', 'rubrica']
            )

            # Salva cols detectadas para o gerador
            self._cols = {
                'rubrica':           col_rubrica,
                'total_prev_atual':  col_total_prev_atual,
                'total_realizado':   col_total_realizado,
                'total_2026':        col_total_2026,
                'por_apto':          col_por_apto,
                'meses':             col_meses,
            }

            # ── Extrai grupos estilo cliente ─────────────────────────────────
            grupos = self._detect_expense_groups(
                df_str, col_rubrica,
                col_total_prev_atual or col_total_2026 or df.columns[1],
                col_por_apto
            )

            # Se não encontrou grupos numerados, constrói a partir das linhas
            if not grupos:
                grupos = self._build_groups_fallback(df, df_str, col_rubrica,
                                                      col_total_prev_atual,
                                                      col_total_2026, col_por_apto)

            print(f"✓ {len(grupos)} grupos de despesa detectados")
            for g in grupos:
                print(f"   {g['numero']}. {g['nome']}: R$ {g['total']:,.2f} ({len(g['itens'])} itens)")

            # ── Fundo de Reserva e Garantidora ──────────────────────────────
            fundo_pct, garantidora_pct = self._detect_reserve_fund(all_text)

            # ── Calcula totais ───────────────────────────────────────────────
            total_despesas = sum(g['total'] for g in grupos)
            fundo_reserva  = total_despesas * (fundo_pct / 100)
            garantidora    = total_despesas * (garantidora_pct / 100)
            total_rateado  = total_despesas + fundo_reserva + garantidora

            # Taxa ideal mensal por unidade
            taxa_ideal_mensal = total_rateado / self._num_unidades / 12

            # Taxa atual (se não detectada, usa a taxa anterior calculada)
            if not self._taxa_atual:
                # Tenta detectar de outra forma
                self._taxa_atual = self._detect_taxa_atual_from_data(
                    df, df_str, col_rubrica,
                    col_total_prev_atual, col_por_apto
                )

            # Reajuste
            if self._taxa_atual and self._taxa_atual > 0:
                reajuste_pct = ((taxa_ideal_mensal / self._taxa_atual) - 1) * 100
            else:
                reajuste_pct = None

            # ── Percentual por grupo ─────────────────────────────────────────
            for g in grupos:
                g['percentual'] = (g['total'] / total_despesas * 100) if total_despesas > 0 else 0

            # ── Dados mensais (se disponíveis) ───────────────────────────────
            dados_mensais = self._extract_monthly_data(df, col_rubrica, col_meses)

            # ── Salva resultados ─────────────────────────────────────────────
            self.df_processed = df
            self.analysis_results = {
                # Metadados
                'ano_atual':             self._ano_atual,
                'ano_proximo':           self._ano_proximo,
                'num_unidades':          self._num_unidades,
                'taxa_atual':            self._taxa_atual,

                # Grupos de despesa (estrutura do cliente)
                'grupos':                grupos,

                # Totais
                'total_despesas':        total_despesas,
                'fundo_reserva':         fundo_reserva,
                'fundo_reserva_pct':     fundo_pct,
                'garantidora':           garantidora,
                'garantidora_pct':       garantidora_pct,
                'total_rateado':         total_rateado,

                # Taxa
                'taxa_ideal_mensal':     taxa_ideal_mensal,
                'taxa_fundo_reserva':    fundo_reserva / self._num_unidades / 12,
                'taxa_total':            taxa_ideal_mensal + (fundo_reserva / self._num_unidades / 12),
                'reajuste_pct':          reajuste_pct,

                # Dados mensais (para gráfico de tendência)
                'dados_mensais':         dados_mensais,

                # DataFrame para análises extras
                'df_processado':         df,

                # Compatibilidade com gerador antigo
                'receita_anual_2025':    total_rateado,
                'receita_anual_2026':    total_rateado,
                'saldo_realizado_2025':  0,
                'saldo_previsto_2026':   0,
                'viavel_2025':           True,
                'viavel_2026':           reajuste_pct is None or reajuste_pct <= 15,
                'necessidade_aumento':   reajuste_pct is not None and reajuste_pct > 0,
                'percentual_aumento_sugerido': reajuste_pct or 0,
            }

            print(f"✓ Análise concluída")
            print(f"  Total Despesas:    R$ {total_despesas:,.2f}")
            print(f"  Total Rateado:     R$ {total_rateado:,.2f}")
            print(f"  Taxa Ideal/mês:    R$ {taxa_ideal_mensal:,.2f}")
            if reajuste_pct is not None:
                print(f"  Reajuste Sugerido: {reajuste_pct:.2f}%")
            return True

        except Exception as e:
            import traceback
            print(f"✗ Erro no processamento: {e}")
            traceback.print_exc()
            return False

    def analyze_finances(self) -> Dict:
        return self.analysis_results

    # ─────────────────────────────────────────────────────────────────────────
    # AUXILIARES DE DETECÇÃO
    # ─────────────────────────────────────────────────────────────────────────
    def _find_col_by_year(self, df, year: int, tipo: Optional[str]):
        """Encontra coluna relacionada a um ano e tipo (previsto/realizado/None)."""
        df_str = df.astype(str)
        year_str = str(year)

        priority_cols = []
        for col in df_str.columns:
            col_s = str(col).lower()
            if year_str in col_s:
                if tipo is None:
                    return col
                if tipo == 'previsto' and ('prev' in col_s or 'aprov' in col_s):
                    return col
                if tipo == 'realizado' and ('real' in col_s or 'total' in col_s):
                    return col
                priority_cols.append(col)

        # Busca nas primeiras linhas
        for row_idx in range(min(5, len(df_str))):
            for col in df_str.columns:
                val = str(df_str.iloc[row_idx][col]).lower()
                if year_str in val:
                    if tipo is None:
                        return col
                    if tipo == 'previsto' and ('prev' in val or 'aprov' in val):
                        return col
                    if tipo == 'realizado' and ('real' in val or 'total' in val):
                        return col

        return priority_cols[0] if priority_cols else None

    def _find_monthly_columns(self, df) -> List:
        """Detecta colunas mensais (MM/AAAA)."""
        result = []
        df_str = df.astype(str)
        pat = re.compile(r'\d{2}/20\d{2}')
        for col in df_str.columns:
            if pat.match(str(col)):
                result.append(col)
        if not result:
            for col in df_str.columns:
                for row_idx in range(min(5, len(df_str))):
                    if pat.match(str(df_str.iloc[row_idx][col])):
                        result.append(col)
                        break
        return result

    def _build_groups_fallback(self, df, df_str, col_rubrica,
                                col_total, col_total_2026, col_por_apto):
        """
        Fallback: se não encontrou grupos numerados,
        agrupa todas as linhas em um único grupo "Despesas".
        """
        col_val = col_total or col_total_2026 or df.columns[1]
        itens = []
        for idx, row in df_str.iterrows():
            rubrica = str(row[col_rubrica]).strip()
            if rubrica in ('nan', '', 'None'):
                continue
            if any(kw in rubrica.upper() for kw in
                   ['TOTAL', 'RATEIO', 'FUNDO', 'GARANTID', 'TAXA']):
                continue
            val = self._safe_float(row.get(col_val, 0))
            por_apto = self._safe_float(row.get(col_por_apto, 0)) if col_por_apto else 0
            if val > 0:
                itens.append({'nome': rubrica, 'total': val, 'por_apto': por_apto})

        if not itens:
            return []

        return [{
            'numero':     1,
            'nome':       'Despesas',
            'itens':      itens,
            'total':      sum(i['total'] for i in itens),
            'percentual': 100.0,
        }]

    def _detect_taxa_atual_from_data(self, df, df_str, col_rubrica,
                                      col_total, col_por_apto) -> Optional[float]:
        """Tenta inferir taxa atual de outras formas."""
        # Procura por "valor atual", "taxa vigente", etc. em qualquer célula
        all_text = ' '.join(str(x) for x in df_str.values.flatten())
        patterns = [
            r'R\$\s*([\d.,]+)\s*(?:atual|vigente|cobra)',
            r'(?:atual|vigente)\s*R\$\s*([\d.,]+)',
            r'taxa.*?R\$\s*([\d.,]+)',
        ]
        for pat in patterns:
            m = re.search(pat, all_text, re.IGNORECASE)
            if m:
                val = self._parse_currency(m.group(1))
                if 100 < val < 5000:  # range razoável para taxa condominial
                    return val
        return None

    def _extract_monthly_data(self, df, col_rubrica, col_meses) -> Dict:
        """Extrai dados mensais se disponíveis."""
        if not col_meses:
            return {}
        try:
            result = {}
            for idx, row in df.iterrows():
                rubrica = str(row[col_rubrica]).strip()
                if rubrica in ('nan', '', 'None'):
                    continue
                vals = [self._safe_float(row.get(c, 0)) for c in col_meses]
                if any(v > 0 for v in vals):
                    result[rubrica] = vals
            return result
        except Exception:
            return {}

    def print_summary(self):
        r = self.analysis_results
        if not r:
            return
        print("\n" + "="*60)
        print(f"PREVISÃO ORÇAMENTÁRIA {r['ano_proximo']}")
        print("="*60)
        for g in r.get('grupos', []):
            print(f"\n{g['numero']}. {g['nome']}: R$ {g['total']:,.2f} ({g['percentual']:.1f}%)")
            for item in g['itens']:
                print(f"   {item['nome']}: R$ {item['total']:,.2f}")
        print(f"\nTotal Despesas:   R$ {r['total_despesas']:,.2f}")
        print(f"Fundo de Reserva: R$ {r['fundo_reserva']:,.2f} ({r['fundo_reserva_pct']}%)")
        print(f"Garantidora:      R$ {r['garantidora']:,.2f} ({r['garantidora_pct']}%)")
        print(f"Total Rateado:    R$ {r['total_rateado']:,.2f}")
        print(f"\nTaxa Ideal/mês:   R$ {r['taxa_ideal_mensal']:,.2f}")
        if r.get('taxa_atual'):
            print(f"Taxa Atual:       R$ {r['taxa_atual']:,.2f}")
        if r.get('reajuste_pct') is not None:
            print(f"Reajuste:         {r['reajuste_pct']:.2f}%")
