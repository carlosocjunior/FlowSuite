// ==========================================
// FLOWSUITE - shared.js
// Funções comuns usadas pelo FlowSuite, admin.html, CalibrationFlow,
// TrainingFlow e ReportFlow. Um lugar só pra manter — corrigir um bug
// aqui corrige nas 5 ferramentas ao mesmo tempo, sem precisar editar
// cada index.html separadamente.
//
// Hospedado no repositório do FlowSuite; as outras ferramentas carregam
// via URL absoluta:
// <script src="https://carlosocjunior.github.io/FlowSuite/shared.js"></script>
//
// IMPORTANTE: esse arquivo precisa ser carregado ANTES do script principal
// de cada página (coloque a tag <script src="...shared.js"> no <head>,
// antes do <script> com o código da ferramenta).
// ==========================================

// ==========================================
// PROTEÇÃO CONTRA XSS — qualquer texto digitado pelo usuário passa por
// aqui antes de virar HTML na tela, pra não deixar alguém injetar
// <script> ou código malicioso.
// ==========================================
function escapeHtml(texto) {
    if (texto === null || texto === undefined) return '';
    return String(texto)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}

// ==========================================
// LEITURA DE PARÂMETROS DA URL (token e empresa "ver como", vindos do
// FlowSuite quando abre uma ferramenta a partir de um card)
// ==========================================
function getTokenDaURL() {
    const params = new URLSearchParams(window.location.search);
    return params.get('fs_token');
}

function getEmpresaVisualizandoDaURL() {
    const params = new URLSearchParams(window.location.search);
    return params.get('fs_empresa') || 'todas';
}

// ==========================================
// PERMISSÃO — cada ferramenta declara sua própria variável
// `permissaoUsuario` (definida depois de validar o token no FlowSuite).
// Essas funções usam essa variável do escopo da página que as chamou;
// funciona porque scripts carregados na mesma página compartilham o
// mesmo escopo de nível superior, mesmo vindo de arquivos diferentes.
// Aceitam um valor explícito também, se preferir passar direto.
// ==========================================
function podeEditar(nivelExplicito) {
    const nivel = nivelExplicito !== undefined
        ? nivelExplicito
        : (typeof permissaoUsuario !== 'undefined' ? permissaoUsuario : null);
    return nivel === 'admin' || nivel === 'editor' || nivel === 'dev' || nivel === 'master' || nivel === 'supervisor';
}

function podeExcluir(nivelExplicito) {
    const nivel = nivelExplicito !== undefined
        ? nivelExplicito
        : (typeof permissaoUsuario !== 'undefined' ? permissaoUsuario : null);
    return nivel === 'admin' || nivel === 'dev' || nivel === 'master';
}

// ==========================================
// NÍVEL DÁ ACESSO A DASHBOARD? (supervisor, admin, master, dev)
// ==========================================
function podeVerDashboard(nivelExplicito) {
    const nivel = nivelExplicito !== undefined
        ? nivelExplicito
        : (typeof permissaoUsuario !== 'undefined' ? permissaoUsuario : null);
    return nivel === 'supervisor' || nivel === 'admin' || nivel === 'dev' || nivel === 'master';
}

// ==========================================
// TRADUÇÃO — o mecanismo é compartilhado, mas cada ferramenta continua
// com seu próprio dicionário `I18N` (os textos são diferentes em cada
// uma) e sua própria variável `idiomaAtual`.
// ==========================================
function t(chave) {
    if (typeof I18N === 'undefined' || typeof idiomaAtual === 'undefined') return chave;
    return (I18N[idiomaAtual] && I18N[idiomaAtual][chave]) || (I18N.pt && I18N.pt[chave]) || chave;
}

function tf(chave, params) {
    let texto = t(chave);
    if (params) {
        Object.keys(params).forEach(p => {
            texto = texto.split(`{${p}}`).join(params[p]);
        });
    }
    return texto;
}

function aplicarIdioma() {
    document.querySelectorAll('[data-i18n]').forEach(el => {
        el.textContent = t(el.getAttribute('data-i18n'));
    });
    document.querySelectorAll('[data-i18n-placeholder]').forEach(el => {
        el.placeholder = t(el.getAttribute('data-i18n-placeholder'));
    });
    document.querySelectorAll('.seletor-idioma-btn').forEach(btn => {
        if (typeof idiomaAtual !== 'undefined') {
            btn.classList.toggle('ativo', btn.dataset.lang === idiomaAtual);
        }
    });
}

// ==========================================
// CACHE LOCAL POR EMPRESA — cada ferramenta usa isso pra montar chaves de
// localStorage que não se misturam entre empresas diferentes no mesmo
// navegador (ex: 'equipamentos_PRO' vs 'equipamentos_TST').
// Depende de `usuarioFlowsuite` (definido depois de validar o token).
// ==========================================
function chaveCache(nomeBase) {
    const empresa = (typeof usuarioFlowsuite !== 'undefined' && usuarioFlowsuite && usuarioFlowsuite.empresaVisualizando)
        || 'default';
    return `${nomeBase}_${empresa}`;
}

// ==========================================
// REVALIDAÇÃO PERIÓDICA DE SESSÃO — cada ferramenta define sua própria
// `validarAcessoFlowsuite()` e `bloquearAcesso()`; essa função só cuida
// do "de quanto em quanto tempo checar de novo".
// ==========================================
function iniciarRevalidacaoPeriodicaSessao(intervaloMs) {
    const intervalo = intervaloMs || 15000;
    setInterval(async () => {
        if (typeof validarAcessoFlowsuite !== 'function') return;
        const aindaValido = await validarAcessoFlowsuite();
        if (!aindaValido && typeof bloquearAcesso === 'function') {
            bloquearAcesso();
        }
    }, intervalo);
}

// ==========================================
// CARREGAMENTO SOB DEMANDA DE BIBLIOTECAS PESADAS (ExcelJS, Chart.js,
// html2canvas...) — só baixa quando a ferramenta realmente precisa,
// em vez de carregar sempre no início da página.
// ==========================================
const _scriptsCarregados = {};

function carregarScript(url) {
    if (_scriptsCarregados[url]) return _scriptsCarregados[url];
    _scriptsCarregados[url] = new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = url;
        script.onload = resolve;
        script.onerror = reject;
        document.head.appendChild(script);
    });
    return _scriptsCarregados[url];
}

async function garantirExcelJS() {
    if (typeof ExcelJS === 'undefined') {
        await carregarScript('https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js');
    }
    if (typeof saveAs === 'undefined') {
        await carregarScript('https://cdn.jsdelivr.net/npm/file-saver@2.0.5/dist/FileSaver.min.js');
    }
}

async function garantirChartJS() {
    if (typeof Chart === 'undefined') {
        await carregarScript('https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js');
    }
}
