"""
Gestor Financeiro de Condomínios — API Web v2.0 (Fase 2: Supabase)
Backend Flask para deploy no Render.

Endpoints:
  GET  /api/health      → ping (aquecimento do free tier) — público
  POST /api/analisar    → recebe planilha Excel, devolve JSON da análise — autenticado
  POST /api/gerar-pptx  → recebe planilha, devolve .pptx e grava histórico — autenticado

Autenticação:
  Header "Authorization: Bearer <access_token do Supabase Auth>".
  O token é validado contra o endpoint /auth/v1/user do Supabase — funciona
  tanto com chaves legadas (JWT HS256) quanto com o formato novo
  (sb_publishable_...), sem guardar segredos de assinatura no backend.

Variáveis de ambiente:
  SUPABASE_URL       ex.: https://xxxx.supabase.co
  SUPABASE_ANON_KEY  chave publishable/anon do MESMO projeto
  CORS_ORIGINS       ex.: https://seusite.netlify.app
  (Se SUPABASE_URL não estiver definida, a API roda em MODO ABERTO —
   apenas para desenvolvimento local. Em produção, configure sempre.)
"""

import os
import re
import tempfile
import unicodedata
from functools import wraps
from io import BytesIO
from datetime import datetime

import numpy as np
import requests as http
from flask import Flask, request, jsonify, send_file, g
from flask_cors import CORS

from condominio_app_v2 import CondominiumFinancialAnalyzer
from powerpoint_generator_v2 import PowerPointGenerator

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10 MB

_origins = os.environ.get('CORS_ORIGINS', '*')
CORS(app, origins=_origins.split(',') if _origins != '*' else '*')

SUPABASE_URL      = (os.environ.get('SUPABASE_URL') or '').rstrip('/')
SUPABASE_ANON_KEY = os.environ.get('SUPABASE_ANON_KEY') or ''
AUTH_ATIVO        = bool(SUPABASE_URL and SUPABASE_ANON_KEY)

if not AUTH_ATIVO:
    print('⚠  SUPABASE_URL/SUPABASE_ANON_KEY não configuradas — '
          'API rodando em MODO ABERTO (apenas desenvolvimento).')

EXTENSOES_PERMITIDAS = {'.xls', '.xlsx'}


# ─────────────────────────────────────────────────────────────────────────────
# AUTENTICAÇÃO (Supabase)
# ─────────────────────────────────────────────────────────────────────────────
def requer_auth(f):
    """Valida o access_token do Supabase. Em modo aberto, deixa passar."""
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not AUTH_ATIVO:
            g.user_token = None
            g.user_id = None
            return f(*args, **kwargs)

        auth_header = request.headers.get('Authorization', '')
        if not auth_header.startswith('Bearer '):
            return jsonify({'erro': 'Não autenticado. Faça login para usar o serviço.'}), 401
        token = auth_header[7:].strip()

        try:
            resp = http.get(
                f'{SUPABASE_URL}/auth/v1/user',
                headers={'Authorization': f'Bearer {token}', 'apikey': SUPABASE_ANON_KEY},
                timeout=10,
            )
        except http.RequestException:
            return jsonify({'erro': 'Falha ao validar a sessão. Tente novamente.'}), 503

        if resp.status_code != 200:
            return jsonify({'erro': 'Sessão inválida ou expirada. Faça login novamente.'}), 401

        g.user_token = token
        g.user_id = resp.json().get('id')
        return f(*args, **kwargs)
    return wrapper


def _gravar_historico(r: dict, nome_arquivo: str, arquivo_path: str = None):
    """
    Insere o registro no histórico via PostgREST usando o token DO USUÁRIO,
    respeitando o RLS (user_id preenchido por default auth.uid() no banco).
    Falha aqui não impede a entrega do PPTX — só registra no log.
    """
    if not (AUTH_ATIVO and g.get('user_token')):
        return
    payload = {
        'nome_condominio':   r.get('nome_condominio'),
        'data_assembleia':   r.get('data_assembleia'),
        'ano_previsao':      _num(r.get('ano_proximo'), int),
        'num_unidades':      _num(r.get('num_unidades'), int),
        'total_despesas':    _num(r.get('total_despesas')),
        'total_rateado':     _num(r.get('total_rateado')),
        'taxa_ideal_mensal': _num(r.get('taxa_ideal_mensal')),
        'taxa_atual':        _num(r.get('taxa_atual')),
        'reajuste_pct':      _num(r.get('reajuste_pct')),
        'nome_arquivo':      nome_arquivo,
        'arquivo_path':      arquivo_path,
    }
    try:
        resp = http.post(
            f'{SUPABASE_URL}/rest/v1/relatorios',
            json=payload,
            headers={
                'Authorization': f'Bearer {g.user_token}',
                'apikey': SUPABASE_ANON_KEY,
                'Content-Type': 'application/json',
                'Prefer': 'return=minimal',
            },
            timeout=10,
        )
        if resp.status_code not in (200, 201):
            print(f'⚠  Histórico não gravado ({resp.status_code}): {resp.text[:200]}')
    except http.RequestException as e:
        print(f'⚠  Histórico não gravado: {e}')


# ─────────────────────────────────────────────────────────────────────────────
# AUXILIARES
# ─────────────────────────────────────────────────────────────────────────────
def _parse_valor_br(texto):
    """Converte '1.234,56' / 'R$ 1234,56' / '1234.56' em float, ou None."""
    if not texto:
        return None
    t = str(texto).replace('R$', '').strip()
    try:
        if ',' in t:
            return float(t.replace('.', '').replace(',', '.'))
        return float(t)
    except ValueError:
        return None


def _num(v, cast=float):
    """Converte numpy/None para tipo nativo ou None (para o payload do banco)."""
    if v is None:
        return None
    try:
        v = cast(v)
        return None if (isinstance(v, float) and v != v) else v
    except (TypeError, ValueError):
        return None


def _sanitize_json(obj):
    """Converte tipos numpy/pandas para tipos nativos serializáveis em JSON."""
    if isinstance(obj, dict):
        return {k: _sanitize_json(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_sanitize_json(v) for v in obj]
    if isinstance(obj, (np.integer,)):
        return int(obj)
    if isinstance(obj, (np.floating,)):
        v = float(obj)
        return None if np.isnan(v) else v
    if isinstance(obj, np.bool_):
        return bool(obj)
    if isinstance(obj, float) and obj != obj:  # NaN nativo
        return None
    return obj


def _slug_filename(nome: str) -> str:
    """Gera nome de arquivo seguro a partir do nome do condomínio."""
    if not nome:
        return "Relatorio_Condominio"
    s = unicodedata.normalize('NFKD', nome).encode('ascii', 'ignore').decode()
    s = re.sub(r'[^A-Za-z0-9]+', '_', s).strip('_')
    return f"Relatorio_{s}" if s else "Relatorio_Condominio"


def _rodar_analise(req):
    """
    Fluxo compartilhado: valida o upload, roda o analisador e injeta
    os campos opcionais (mesma lógica da GUI desktop v3).
    Retorna (analyzer, None) em sucesso ou (None, (payload, status)) em erro.
    """
    arquivo = req.files.get('planilha')
    if arquivo is None or arquivo.filename == '':
        return None, ({'erro': 'Nenhuma planilha enviada. Envie um arquivo .xls ou .xlsx no campo "planilha".'}, 400)

    ext = os.path.splitext(arquivo.filename)[1].lower()
    if ext not in EXTENSOES_PERMITIDAS:
        return None, ({'erro': f'Formato "{ext}" não suportado. Envie .xls ou .xlsx.'}, 400)

    nome = (req.form.get('nome_condominio') or '').strip()
    data = (req.form.get('data_assembleia') or '').strip()
    taxa = (req.form.get('taxa_atual') or '').strip()

    # O analisador trabalha com caminho de arquivo (detecção de engine por
    # extensão), então gravamos em arquivo temporário com o sufixo correto.
    tmp = tempfile.NamedTemporaryFile(suffix=ext, delete=False)
    try:
        arquivo.save(tmp.name)
        tmp.close()

        analyzer = CondominiumFinancialAnalyzer(tmp.name)
        if analyzer.df_raw is None:
            return None, ({'erro': 'Não foi possível ler a planilha. Verifique se o arquivo está íntegro.'}, 422)

        # Injeta dados opcionais ANTES do processamento (influenciam os cálculos)
        t = _parse_valor_br(taxa)
        if t:
            analyzer._taxa_atual = t
        unidades = req.form.get('num_unidades')
        if unidades and str(unidades).strip().isdigit():
            analyzer._num_unidades = int(str(unidades).strip())

        if not analyzer.process_data():
            return None, ({'erro': 'Falha ao processar os dados da planilha.'}, 422)

        # Sem fallback silencioso: se a planilha não for um export
        # reconhecido (BalAnual/Condomínio21), devolve erro claro em vez
        # de gerar um relatório incompleto ou com dados errados.
        if analyzer.analysis_results.get('formato') != 'balanual':
            return None, ({'erro': 'Planilha não reconhecida. Envie o export '
                                   '"BalAnual" do Condomínio21 (colunas mensais '
                                   '01/AAAA a 12/AAAA). Se esta planilha é um '
                                   'export válido de outro condomínio, envie o '
                                   'arquivo ao suporte para inclusão do formato.'},
                          422)

        # Re-injeta após process_data (que pode sobrescrever) — mesma lógica da GUI
        if nome:
            analyzer.analysis_results['nome_condominio'] = nome.upper()
        if data:
            analyzer.analysis_results['data_assembleia'] = data

        return analyzer, None
    finally:
        try:
            os.unlink(tmp.name)
        except OSError:
            pass


# ─────────────────────────────────────────────────────────────────────────────
# ROTAS
# ─────────────────────────────────────────────────────────────────────────────
@app.route('/api/health', methods=['GET'])
def health():
    return jsonify({
        'status': 'ok',
        'servico': 'gestor-condominio-api',
        'versao': '2.0',
        'auth': AUTH_ATIVO,
    })


@app.route('/api/analisar', methods=['POST'])
@requer_auth
def analisar():
    analyzer, erro = _rodar_analise(request)
    if erro:
        payload, status = erro
        return jsonify(payload), status

    r = analyzer.analysis_results
    resumo = _sanitize_json({
        'nome_condominio':   r.get('nome_condominio'),
        'data_assembleia':   r.get('data_assembleia'),
        'ano_proximo':       r.get('ano_proximo'),
        'num_unidades':      r.get('num_unidades'),
        'grupos':            r.get('grupos', []),
        'total_despesas':    r.get('total_despesas'),
        'fundo_reserva':     r.get('fundo_reserva'),
        'fundo_reserva_pct': r.get('fundo_reserva_pct'),
        'garantidora':       r.get('garantidora'),
        'garantidora_pct':   r.get('garantidora_pct'),
        'total_rateado':     r.get('total_rateado'),
        'taxa_ideal_mensal': r.get('taxa_ideal_mensal'),
        'taxa_atual':        r.get('taxa_atual'),
        'reajuste_pct':      r.get('reajuste_pct'),
        'formato':           r.get('formato', 'legado'),
        'balanual':          r.get('balanual'),
    })
    return jsonify({'ok': True, 'analise': resumo})


@app.route('/api/gerar-pptx', methods=['POST'])
@requer_auth
def gerar_pptx():
    analyzer, erro = _rodar_analise(request)
    if erro:
        payload, status = erro
        return jsonify(payload), status

    # prs.save() aceita stream: geramos o PPTX inteiramente em memória,
    # sem tocar no filesystem efêmero do Render.
    extras = {
        'inadimplencia_total': _parse_valor_br(request.form.get('inadimplencia')),
        'acordos_receber':     _parse_valor_br(request.form.get('acordos')),
    }
    arquivos_temp = []

    # Texto do edital (preferido pelo cliente — vira slide legível)
    edital_texto = (request.form.get('edital_texto') or '').strip()
    if edital_texto:
        extras['edital_texto'] = edital_texto[:8000]

    # Logo padrão embutida (Suport) — usada quando nenhuma é enviada
    logo_padrao = os.path.join(os.path.dirname(__file__), 'assets',
                               'logo_padrao.png')
    if os.path.exists(logo_padrao):
        extras['logo_path'] = logo_padrao

    # Logo opcional (PNG/JPG)
    logo = request.files.get('logo')
    if logo and logo.filename:
        ext_l = os.path.splitext(logo.filename)[1].lower()
        if ext_l in ('.png', '.jpg', '.jpeg'):
            tmp_l = tempfile.NamedTemporaryFile(suffix=ext_l, delete=False)
            logo.save(tmp_l.name); tmp_l.close()
            extras['logo_path'] = tmp_l.name
            arquivos_temp.append(tmp_l.name)

    # Edital opcional (PDF ou imagem) → vira slides após a capa
    edital = request.files.get('edital')
    if edital and edital.filename:
        ext_e = os.path.splitext(edital.filename)[1].lower()
        if ext_e in ('.png', '.jpg', '.jpeg'):
            tmp_e = tempfile.NamedTemporaryFile(suffix=ext_e, delete=False)
            edital.save(tmp_e.name); tmp_e.close()
            extras['edital_paths'] = [tmp_e.name]
            arquivos_temp.append(tmp_e.name)
        elif ext_e == '.pdf':
            try:
                import fitz  # PyMuPDF
                tmp_p = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
                edital.save(tmp_p.name); tmp_p.close()
                arquivos_temp.append(tmp_p.name)
                paginas = []
                with fitz.open(tmp_p.name) as doc:
                    for pg in doc[:3]:  # máx. 3 páginas
                        pix = pg.get_pixmap(dpi=150)
                        tmp_img = tempfile.NamedTemporaryFile(
                            suffix='.png', delete=False)
                        pix.save(tmp_img.name); tmp_img.close()
                        paginas.append(tmp_img.name)
                        arquivos_temp.append(tmp_img.name)
                extras['edital_paths'] = paginas
            except Exception as e:
                print(f'⚠  Edital PDF não processado: {e}')
    buffer = BytesIO()
    gen = PowerPointGenerator(analyzer, output_path=buffer, extras=extras)
    gen.generate()
    buffer.seek(0)

    nome_arquivo = _slug_filename(analyzer.analysis_results.get('nome_condominio', ''))
    nome_arquivo += f"_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"

    # Sobe o PPTX ao Storage para permitir download pelo histórico
    # (usa o token do usuário — RLS por pasta {user_id}/...)
    arquivo_path = None
    if AUTH_ATIVO and g.get('user_token') and g.get('user_id'):
        try:
            caminho = f"{g.user_id}/{nome_arquivo}"
            resp_up = http.post(
                f'{SUPABASE_URL}/storage/v1/object/relatorios/{caminho}',
                data=buffer.getvalue(),
                headers={
                    'Authorization': f'Bearer {g.user_token}',
                    'apikey': SUPABASE_ANON_KEY,
                    'Content-Type': ('application/vnd.openxmlformats-'
                                     'officedocument.presentationml.presentation'),
                    'x-upsert': 'true',
                },
                timeout=30,
            )
            if resp_up.status_code in (200, 201):
                arquivo_path = caminho
            else:
                print(f'⚠  Storage não gravado ({resp_up.status_code}): '
                      f'{resp_up.text[:200]}')
        except http.RequestException as e:
            print(f'⚠  Storage não gravado: {e}')

    # Grava histórico no Supabase (não bloqueia a entrega em caso de falha)
    _gravar_historico(analyzer.analysis_results, nome_arquivo, arquivo_path)

    for _t in arquivos_temp:
        try:
            os.unlink(_t)
        except OSError:
            pass

    return send_file(
        buffer,
        as_attachment=True,
        download_name=nome_arquivo,
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )


@app.errorhandler(413)
def arquivo_grande(_):
    return jsonify({'erro': 'Arquivo maior que 10 MB. Envie uma planilha menor.'}), 413


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=True)
