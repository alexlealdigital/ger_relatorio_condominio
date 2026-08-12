/* ═══════════════════════════════════════════════════════════════════════
   Gestor Financeiro de Condomínios — Web · app.js (Fase 2: Supabase)
   ═══════════════════════════════════════════════════════════════════════

   CONFIGURE AQUI antes do deploy:

   1. API_URL       → URL do backend no Render
                      Ex.: 'https://gestor-condominio-api.onrender.com'
                      Dev local: 'http://localhost:5000'

   2. SUPABASE_URL  → Project URL do projeto Supabase dedicado
                      (Dashboard → Settings → API)

   3. SUPABASE_KEY  → chave publishable (sb_publishable_...) do MESMO projeto.
                      É pública por design — o RLS protege os dados.
*/
const API_URL      = 'https://gestor-condominio-api.onrender.com';
const SUPABASE_URL = 'https://scjiztniwqmdpzjppigm.supabase.co';
const SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InNjaml6dG5pd3FtZHB6anBwaWdtIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODMwMTA4NDYsImV4cCI6MjA5ODU4Njg0Nn0.kXIi6iBxyJ7YVI9LjUHGU9Zjct9MtnNwtQMgJJ94cVM';

/* ─── Supabase client ────────────────────────────────────────────────── */
const sb = window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY);
let sessao = null;

/* ─── Elementos ──────────────────────────────────────────────────────── */
const $ = (id) => document.getElementById(id);

const telaLogin    = $('telaLogin');
const telaApp      = $('telaApp');
const btnEntrar    = $('btnEntrar');
const btnSair      = $('btnSair');
const fbLogin      = $('feedbackLogin');

const dropzone     = $('dropzone');
const fileInput    = $('fileInput');
const fileNameEl   = $('fileName');
const btnAnalisar  = $('btnAnalisar');
const btnGerar     = $('btnGerar');
const fbAnalise    = $('feedbackAnalise');
const fbPpt        = $('feedbackPpt');
const secResultado = $('resultado');
const secHistorico = $('secHistorico');

let arquivoSelecionado = null;

/* ─── Formatação ─────────────────────────────────────────────────────── */
const fmtBRL = (v) =>
  (v === null || v === undefined)
    ? '—'
    : Number(v).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });

const fmtPct = (v) =>
  (v === null || v === undefined) ? '—' : `${Number(v).toFixed(2).replace('.', ',')}%`;

const fmtData = (iso) =>
  new Date(iso).toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit', year: 'numeric' });

/* ─── Sessão / troca de telas ────────────────────────────────────────── */
function aplicarSessao(s) {
  sessao = s;
  const logado = !!s;
  telaLogin.hidden = logado;
  telaApp.hidden   = !logado;
  btnSair.hidden   = !logado;
  if (logado) carregarHistorico();
}

sb.auth.getSession().then(({ data }) => aplicarSessao(data.session));
sb.auth.onAuthStateChange((_evento, s) => aplicarSessao(s));

/* ─── Troca de vistas do card de acesso ──────────────────────────── */
const vistas = ['vistaLogin', 'vistaCadastro', 'vistaEsqueci', 'vistaNovaSenha'];
function mostrarVista(id) {
  vistas.forEach((v) => { $(v).hidden = v !== id; });
  fbLogin.textContent = '';
}
$('linkCriarConta').addEventListener('click', (e) => { e.preventDefault(); mostrarVista('vistaCadastro'); });
$('linkEsqueci').addEventListener('click', (e) => { e.preventDefault(); mostrarVista('vistaEsqueci'); });
document.querySelectorAll('.linkVoltar').forEach((l) =>
  l.addEventListener('click', (e) => { e.preventDefault(); mostrarVista('vistaLogin'); }));

/* ─── Criar conta (exige chave corporativa) ──────────────────────── */
$('btnCadastrar').addEventListener('click', async () => {
  const email = $('cadEmail').value.trim();
  const senha = $('cadSenha').value;
  const chave = $('cadChave').value.trim().toUpperCase();
  if (!email || senha.length < 8 || !chave) {
    fbLogin.className = 'feedback err';
    fbLogin.textContent = 'Preencha e-mail, senha (mín. 8 caracteres) e a chave corporativa.';
    return;
  }
  $('btnCadastrar').disabled = true;
  fbLogin.className = 'feedback wait';
  fbLogin.textContent = 'Criando conta…';
  const { error } = await sb.auth.signUp({
    email, password: senha,
    options: { data: { chave_corporativa: chave } },
  });
  $('btnCadastrar').disabled = false;
  if (error) {
    fbLogin.className = 'feedback err';
    fbLogin.textContent = /chave/i.test(error.message)
      ? 'Chave de acesso corporativa inválida ou expirada.'
      : (/already/i.test(error.message)
          ? 'Este e-mail já possui conta. Use "Esqueci a senha" se necessário.'
          : 'Não foi possível criar a conta. Verifique os dados.');
  } else {
    fbLogin.className = 'feedback ok';
    fbLogin.textContent = 'Conta criada com sucesso!';
  }
});

/* ─── Esqueci a senha ────────────────────────────────────────────── */
$('btnRecuperar').addEventListener('click', async () => {
  const email = $('recEmail').value.trim();
  if (!email) {
    fbLogin.className = 'feedback err';
    fbLogin.textContent = 'Informe o e-mail.';
    return;
  }
  $('btnRecuperar').disabled = true;
  fbLogin.className = 'feedback wait';
  fbLogin.textContent = 'Enviando link…';
  const { error } = await sb.auth.resetPasswordForEmail(email, {
    redirectTo: window.location.origin + window.location.pathname,
  });
  $('btnRecuperar').disabled = false;
  fbLogin.className = error ? 'feedback err' : 'feedback ok';
  fbLogin.textContent = error
    ? 'Não foi possível enviar. Tente novamente em alguns minutos.'
    : 'Link enviado! Confira sua caixa de entrada (e o spam).';
});

/* ─── Nova senha (após clicar no link do e-mail) ─────────────────── */
sb.auth.onAuthStateChange((evento) => {
  if (evento === 'PASSWORD_RECOVERY') {
    telaLogin.hidden = false;
    telaApp.hidden = true;
    mostrarVista('vistaNovaSenha');
  }
});
$('btnSalvarSenha').addEventListener('click', async () => {
  const senha = $('novaSenha').value;
  if (senha.length < 8) {
    fbLogin.className = 'feedback err';
    fbLogin.textContent = 'A nova senha precisa de pelo menos 8 caracteres.';
    return;
  }
  $('btnSalvarSenha').disabled = true;
  const { error } = await sb.auth.updateUser({ password: senha });
  $('btnSalvarSenha').disabled = false;
  if (error) {
    fbLogin.className = 'feedback err';
    fbLogin.textContent = 'Não foi possível salvar. Peça um novo link de recuperação.';
  } else {
    fbLogin.className = 'feedback ok';
    fbLogin.textContent = 'Senha atualizada!';
    mostrarVista('vistaLogin');
  }
});

btnEntrar.addEventListener('click', async () => {
  const email = $('loginEmail').value.trim();
  const senha = $('loginSenha').value;
  if (!email || !senha) {
    fbLogin.className = 'feedback err';
    fbLogin.textContent = 'Informe e-mail e senha.';
    return;
  }
  btnEntrar.disabled = true;
  fbLogin.className = 'feedback wait';
  fbLogin.textContent = 'Entrando…';
  const { error } = await sb.auth.signInWithPassword({ email, password: senha });
  btnEntrar.disabled = false;
  if (error) {
    fbLogin.className = 'feedback err';
    fbLogin.textContent = 'E-mail ou senha incorretos.';
  } else {
    fbLogin.textContent = '';
  }
});

$('loginSenha').addEventListener('keydown', (e) => {
  if (e.key === 'Enter') btnEntrar.click();
});

btnSair.addEventListener('click', async () => {
  await sb.auth.signOut();
  secResultado.hidden = true;
  fbAnalise.textContent = '';
  fbPpt.textContent = '';
  $('loginEmail').value = '';
  $('loginSenha').value = '';
  fbLogin.textContent = '';
});

/* Headers autenticados para a API Flask */
function headersAuth() {
  return sessao ? { Authorization: `Bearer ${sessao.access_token}` } : {};
}

/* ─── Status do servidor (aquecimento do free tier) ──────────────────── */
async function pingServidor() {
  const dot = $('serverDot');
  const label = $('serverLabel');
  try {
    const inicio = Date.now();
    const res = await fetch(`${API_URL}/api/health`, { cache: 'no-store' });
    if (!res.ok) throw new Error();
    const ms = Date.now() - inicio;
    dot.className = 'dot online';
    label.textContent = ms > 5000 ? 'Servidor online (acordou agora)' : 'Servidor online';
  } catch {
    dot.className = 'dot offline';
    label.textContent = 'Servidor indisponível';
  }
}
pingServidor();

/* ─── Seleção de arquivo ─────────────────────────────────────────────── */
function definirArquivo(file) {
  if (!file) return;
  const ext = file.name.slice(file.name.lastIndexOf('.')).toLowerCase();
  if (ext !== '.xls' && ext !== '.xlsx') {
    fbAnalise.className = 'feedback err';
    fbAnalise.textContent = 'Formato não suportado. Envie um arquivo .xls ou .xlsx.';
    return;
  }
  if (file.size > 10 * 1024 * 1024) {
    fbAnalise.className = 'feedback err';
    fbAnalise.textContent = 'Arquivo maior que 10 MB.';
    return;
  }
  arquivoSelecionado = file;
  fileNameEl.textContent = file.name;
  btnAnalisar.disabled = false;
  fbAnalise.textContent = '';
  secResultado.hidden = true;
}

dropzone.addEventListener('click', () => fileInput.click());
dropzone.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInput.click(); }
});
fileInput.addEventListener('change', () => definirArquivo(fileInput.files[0]));

['dragover', 'dragenter'].forEach((ev) =>
  dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.add('dragover'); }));
['dragleave', 'drop'].forEach((ev) =>
  dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.remove('dragover'); }));
dropzone.addEventListener('drop', (e) => definirArquivo(e.dataTransfer.files[0]));

/* ─── Montagem do FormData (compartilhado) ───────────────────────────── */
function montarFormData() {
  const fd = new FormData();
  fd.append('planilha', arquivoSelecionado);
  const nome = $('nomeCondominio').value.trim();
  const data = $('dataAssembleia').value.trim();
  if (nome) fd.append('nome_condominio', nome);
  if (data && data !== 'dd/mm/aaaa') fd.append('data_assembleia', data);
  const logo = $('logoInput').files[0];
  if (logo) fd.append('logo', logo);
  const editalTexto = $('editalTexto').value.trim();
  if (editalTexto) fd.append('edital_texto', editalTexto);
  return fd;
}

async function extrairErro(res) {
  try {
    const j = await res.json();
    return j.erro || `Erro ${res.status}`;
  } catch {
    return `Erro ${res.status} ao comunicar com o servidor.`;
  }
}

/* ─── Passo 2 → Analisar ─────────────────────────────────────────────── */
btnAnalisar.addEventListener('click', async () => {
  if (!arquivoSelecionado) return;
  btnAnalisar.disabled = true;
  fbAnalise.className = 'feedback wait';
  fbAnalise.textContent = 'Analisando planilha… Se o servidor estiver dormindo (plano gratuito), pode levar até 1 minuto.';

  try {
    const res = await fetch(`${API_URL}/api/analisar`, {
      method: 'POST',
      headers: headersAuth(),
      body: montarFormData(),
    });
    if (!res.ok) throw new Error(await extrairErro(res));
    const { analise } = await res.json();
    renderResultado(analise);
    fbAnalise.className = 'feedback ok';
    fbAnalise.textContent = 'Análise concluída com sucesso.';
    secResultado.hidden = false;
    secResultado.scrollIntoView({ behavior: 'smooth', block: 'start' });
  } catch (err) {
    fbAnalise.className = 'feedback err';
    fbAnalise.textContent = err.message || 'Falha na análise.';
  } finally {
    btnAnalisar.disabled = false;
  }
});

/* ─── Renderização do resultado ──────────────────────────────────────── */
function renderResultado(a) {
  $('tituloResultado').textContent =
    a.nome_condominio ? `Análise — ${a.nome_condominio}` : 'Análise concluída';
  $('subtituloResultado').textContent =
    `Previsão orçamentária ${a.ano_proximo ?? ''}` +
    (a.data_assembleia ? ` · Assembleia em ${a.data_assembleia}` : '');

  const b = a.balanual;
  const stats = b ? [
    ['Unidades', a.num_unidades ?? '—'],
    ['Previsto aprovado / mês', fmtBRL(b.prev_aprovada_mensal)],
    ['Realizado médio / mês', fmtBRL(b.realizado_medio_mensal)],
    ['Desvio sobre o aprovado', b.desvio_pct != null ? `${b.desvio_pct >= 0 ? '+' : ''}${fmtPct(b.desvio_pct)}` : '—'],
    [`Previsão ${a.ano_proximo ?? ''} / mês`, fmtBRL(b.prev_next_mensal)],
    [`Fundo de reserva (${a.fundo_reserva_pct ?? '—'}%)`, fmtBRL(b.fundo_mensal)],
  ] : [
    ['Unidades', a.num_unidades ?? '—'],
    ['Total de despesas', fmtBRL(a.total_despesas)],
    [`Fundo de reserva (${a.fundo_reserva_pct ?? '—'}%)`, fmtBRL(a.fundo_reserva)],
    ['Total rateado', fmtBRL(a.total_rateado)],
  ];
  if (!b && a.garantidora) stats.push([`Garantidora (${a.garantidora_pct ?? '—'}%)`, fmtBRL(a.garantidora)]);

  $('stats').innerHTML = stats.map(([label, value]) => `
    <div class="stat">
      <div class="label">${label}</div>
      <div class="value">${value}</div>
    </div>`).join('');

  // Taxa ideal + badge de reajuste
  let badge = '';
  if (a.reajuste_pct !== null && a.reajuste_pct !== undefined) {
    const sobe = a.reajuste_pct >= 0;
    badge = `<span class="badge ${sobe ? 'alta' : 'baixa'}">
      ${sobe ? '▲' : '▼'} Reajuste de ${fmtPct(Math.abs(a.reajuste_pct))}
      ${a.taxa_atual ? ` sobre ${fmtBRL(a.taxa_atual)}` : ''}
    </span>`;
  }
  $('taxaCard').innerHTML = `
    <div>
      <div class="taxa-label">Taxa condominial ideal / mês</div>
      <div class="taxa-valor">${fmtBRL(a.taxa_ideal_mensal)}</div>
    </div>
    ${badge}`;

  // Grupos com barra de percentual
  const grupos = a.grupos || [];
  $('grupos').innerHTML = grupos.length
    ? grupos.map((g) => `
      <div class="grupo">
        <div class="grupo-linha">
          <span class="grupo-nome">${g.numero}. ${g.nome}</span>
          <span class="grupo-valor">${fmtBRL(g.total)} · ${fmtPct(g.percentual)}</span>
        </div>
        <div class="grupo-bar"><span style="width:${Math.min(100, Math.max(2, g.percentual || 0))}%"></span></div>
      </div>`).join('')
    : '<p class="hint">Nenhum grupo numerado detectado na planilha.</p>';
}

/* ─── Passo 3 → Gerar PPTX ───────────────────────────────────────────── */
async function gerarDocumento(rota, btn) {

  if (!arquivoSelecionado) return;
  btn.disabled = true;
  fbPpt.className = 'feedback wait';
  fbPpt.textContent = 'Gerando apresentação… Isso leva alguns segundos.';

  try {
    const res = await fetch(`${API_URL}${rota}`, {
      method: 'POST',
      headers: headersAuth(),
      body: montarFormData(),
    });
    if (!res.ok) throw new Error(await extrairErro(res));

    // Nome do arquivo vindo do header Content-Disposition (com fallback)
    const cd = res.headers.get('Content-Disposition') || '';
    const match = cd.match(/filename\*?=(?:UTF-8'')?"?([^";]+)"?/i);
    const nomeArquivo = match ? decodeURIComponent(match[1]) : 'Relatorio_Condominio';

    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = nomeArquivo;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);

    fbPpt.className = 'feedback ok';
    fbPpt.textContent = `Apresentação gerada: ${nomeArquivo}`;
    carregarHistorico();
  } catch (err) {
    fbPpt.className = 'feedback err';
    fbPpt.textContent = err.message || 'Falha ao gerar a apresentação.';
  } finally {
    btn.disabled = false;
  }
}

btnGerar.addEventListener('click', () => gerarDocumento('/api/gerar-pptx', btnGerar));
$('btnGerarPdf').addEventListener('click', () => gerarDocumento('/api/gerar-pdf', $('btnGerarPdf')));

/* ─── Histórico (leitura direta do Supabase — RLS protege) ───────────── */
async function carregarHistorico() {
  const { data, error } = await sb
    .from('relatorios')
    .select('nome_condominio, data_assembleia, ano_previsao, taxa_ideal_mensal, reajuste_pct, nome_arquivo, arquivo_path, criado_em')
    .order('criado_em', { ascending: false })
    .limit(15);

  if (error || !data || data.length === 0) {
    secHistorico.hidden = true;
    return;
  }

  secHistorico.hidden = false;
  $('historico').innerHTML = data.map((r) => `
    <div class="hist-item">
      <div class="hist-main">
        <span class="hist-nome">${r.nome_condominio || 'Condomínio'}</span>
        <span class="hist-meta">
          ${r.ano_previsao ? `Previsão ${r.ano_previsao}` : ''}
          ${r.data_assembleia ? ` · Assembleia ${r.data_assembleia}` : ''}
          · Gerado em ${fmtData(r.criado_em)}
        </span>
      </div>
      <div class="hist-valores">
        <span class="hist-taxa">${fmtBRL(r.taxa_ideal_mensal)}/mês</span>
        ${r.reajuste_pct !== null && r.reajuste_pct !== undefined
          ? `<span class="hist-reajuste ${r.reajuste_pct >= 0 ? 'alta' : 'baixa'}">${r.reajuste_pct >= 0 ? '▲' : '▼'} ${fmtPct(Math.abs(r.reajuste_pct))}</span>`
          : ''}
        ${r.arquivo_path
          ? `<button class="btn-baixar" data-path="${r.arquivo_path}">Baixar</button>`
          : ''}
      </div>
    </div>`).join('');

  document.querySelectorAll('.btn-baixar').forEach((btn) => {
    btn.addEventListener('click', async () => {
      btn.disabled = true;
      const original = btn.textContent;
      btn.textContent = '…';
      try {
        const { data, error } = await sb.storage
          .from('relatorios')
          .createSignedUrl(btn.dataset.path, 3600);
        if (error || !data?.signedUrl) throw error;
        window.open(data.signedUrl, '_blank');
      } catch {
        btn.textContent = 'Erro';
        setTimeout(() => { btn.textContent = original; btn.disabled = false; }, 2000);
        return;
      }
      btn.textContent = original;
      btn.disabled = false;
    });
  });
}
