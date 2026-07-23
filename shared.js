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

// ==========================================
// CHATBOT DE DÚVIDAS — widget flutuante que aparece em QUALQUER
// ferramenta que carregue este shared.js (FlowSuite, ReportFlow,
// CalibrationFlow, TrainingFlow, admin). Manda as perguntas pro Worker
// (ação 'chatDuvida'), que fala com a API da Anthropic usando uma
// chave guardada como secret (nunca exposta aqui no front).
// ==========================================
const CHAT_FLOWSUITE_API_URL = 'https://flowsuite-plus-api.flowsuite-plus.workers.dev';

// Histórico da conversa fica só em memória (some ao recarregar a
// página) — de propósito, pra não guardar conteúdo de chat em
// sessionStorage/localStorage sem necessidade.
let _chatFlowSuiteHistorico = [];
let _chatFlowSuiteAberto = false;
let _chatFlowSuiteCarregando = false;

function _chatObterToken() {
    const tokenSessao = sessionStorage.getItem('flowsuite_token');
    if (tokenSessao) return tokenSessao;
    if (typeof getTokenDaURL === 'function') {
        const tokenUrl = getTokenDaURL();
        if (tokenUrl) return tokenUrl;
    }
    return null;
}

function _chatModuloAtual() {
    if (typeof MODULO_ATUAL !== 'undefined' && MODULO_ATUAL) return MODULO_ATUAL;
    return 'hub';
}

function _chatInjetarWidget() {
    if (document.getElementById('chatFlowSuiteBotao')) return; // já injetado

    const botao = document.createElement('button');
    botao.id = 'chatFlowSuiteBotao';
    botao.title = 'Dúvidas sobre o FlowSuite';
    botao.innerHTML = '💬';
    botao.style.cssText = `
        position: fixed; bottom: 22px; right: 22px; width: 56px; height: 56px;
        border-radius: 50%; background: #007bff; color: white; border: none;
        font-size: 24px; box-shadow: 0 6px 18px rgba(0,0,0,0.25); cursor: pointer;
        z-index: 99998; display: flex; align-items: center; justify-content: center;
        transition: transform 0.2s;
    `;
    botao.onmouseenter = () => botao.style.transform = 'scale(1.08)';
    botao.onmouseleave = () => botao.style.transform = 'scale(1)';
    botao.addEventListener('click', _chatToggleJanela);

    const painel = document.createElement('div');
    painel.id = 'chatFlowSuitePainel';
    painel.style.cssText = `
        position: fixed; bottom: 90px; right: 22px; width: 340px; max-width: calc(100vw - 44px);
        height: 460px; max-height: calc(100vh - 120px); background: white; border-radius: 14px;
        box-shadow: 0 12px 40px rgba(0,0,0,0.30); display: none; flex-direction: column;
        overflow: hidden; z-index: 99999; font-family: 'Poppins', Arial, sans-serif;
    `;
    painel.innerHTML = `
        <div style="background:#007bff;color:white;padding:14px 16px;display:flex;justify-content:space-between;align-items:center;">
            <strong style="font-size:14px;">Dúvidas sobre o FlowSuite</strong>
            <span id="chatFlowSuiteFechar" style="cursor:pointer;font-size:18px;line-height:1;">×</span>
        </div>
        <div id="chatFlowSuiteMensagens" style="flex:1;overflow-y:auto;padding:14px;background:#f7f8fa;display:flex;flex-direction:column;gap:10px;font-size:13px;"></div>
        <div style="padding:10px;border-top:1px solid #eee;display:flex;gap:8px;">
            <input id="chatFlowSuiteInput" type="text" placeholder="Digite sua dúvida..." style="flex:1;padding:10px 12px;border:1px solid #ddd;border-radius:20px;font-size:13px;outline:none;">
            <button id="chatFlowSuiteEnviar" style="background:#007bff;color:white;border:none;width:38px;height:38px;border-radius:50%;cursor:pointer;font-size:16px;flex-shrink:0;">➤</button>
        </div>
    `;

    document.body.appendChild(botao);
    document.body.appendChild(painel);

    document.getElementById('chatFlowSuiteFechar').addEventListener('click', _chatToggleJanela);
    document.getElementById('chatFlowSuiteEnviar').addEventListener('click', _chatEnviarPergunta);
    document.getElementById('chatFlowSuiteInput').addEventListener('keydown', function(e) {
        if (e.key === 'Enter') _chatEnviarPergunta();
    });
}

function _chatToggleJanela() {
    _chatFlowSuiteAberto = !_chatFlowSuiteAberto;
    const painel = document.getElementById('chatFlowSuitePainel');
    if (!painel) return;
    painel.style.display = _chatFlowSuiteAberto ? 'flex' : 'none';

    if (_chatFlowSuiteAberto && _chatFlowSuiteHistorico.length === 0) {
        _chatAdicionarMensagemTela('assistant', 'Oi! Sou o assistente do FlowSuite. Pode perguntar como usar qualquer parte do sistema (ReportFlow, CalibrationFlow, TrainingFlow ou o painel principal).');
    }
    if (_chatFlowSuiteAberto) {
        const input = document.getElementById('chatFlowSuiteInput');
        if (input) input.focus();
    }
}

function _chatAdicionarMensagemTela(role, texto) {
    const container = document.getElementById('chatFlowSuiteMensagens');
    if (!container) return;
    const bolha = document.createElement('div');
    const ehUsuario = role === 'user';
    bolha.style.cssText = `
        align-self: ${ehUsuario ? 'flex-end' : 'flex-start'};
        background: ${ehUsuario ? '#007bff' : '#e9ecef'};
        color: ${ehUsuario ? 'white' : '#333'};
        padding: 8px 12px; border-radius: 14px; max-width: 85%; white-space: pre-wrap; line-height: 1.4;
    `;
    bolha.textContent = texto;
    container.appendChild(bolha);
    container.scrollTop = container.scrollHeight;
    return bolha;
}

async function _chatEnviarPergunta() {
    if (_chatFlowSuiteCarregando) return;

    const input = document.getElementById('chatFlowSuiteInput');
    const pergunta = input ? input.value.trim() : '';
    if (!pergunta) return;

    const token = _chatObterToken();
    if (!token) {
        _chatAdicionarMensagemTela('assistant', 'Você precisa estar logado no FlowSuite pra usar o assistente. Faça login e tente de novo.');
        return;
    }

    input.value = '';
    _chatAdicionarMensagemTela('user', pergunta);
    const bolhaCarregando = _chatAdicionarMensagemTela('assistant', '...');
    _chatFlowSuiteCarregando = true;

    try {
        const resposta = await fetch(CHAT_FLOWSUITE_API_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({
                acao: 'chatDuvida',
                token: token,
                pergunta: pergunta,
                ferramenta: _chatModuloAtual(),
                historico: _chatFlowSuiteHistorico
            })
        });
        const resultado = await resposta.json();

        if (resultado.sucesso) {
            bolhaCarregando.textContent = resultado.resposta;
            _chatFlowSuiteHistorico.push({ role: 'user', content: pergunta });
            _chatFlowSuiteHistorico.push({ role: 'assistant', content: resultado.resposta });
        } else {
            bolhaCarregando.textContent = 'Não consegui responder agora. Tenta de novo em instantes.';
        }
    } catch (e) {
        bolhaCarregando.textContent = 'Não consegui me conectar. Verifique sua internet e tente de novo.';
    } finally {
        _chatFlowSuiteCarregando = false;
    }
}

// Injeta o widget assim que a página (e o DOM) estiver pronta.
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', _chatInjetarWidget);
} else {
    _chatInjetarWidget();
}
