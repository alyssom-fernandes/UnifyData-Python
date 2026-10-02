"""
UnifyData — unifydata.py
Tela do fechamento em Streamlit. As regras ficam em nucleo.py.
"""
import io
from datetime import date
from html import escape
from pathlib import Path

import pandas as pd
import streamlit as st

from nucleo import (
    brl, consolidar, divergencia_repetido, empresa_do_arquivo, encontrar_divergencias, fixar_linha, formatar_cpf,
    gerar_csv, gerar_excel, lista_por_extenso, nome_exibicao, normalizar, ordenar, parse_hr_files, parse_linx_files,
    pelo_nome, renomear, so_digitos,
)

PASTA = Path(__file__).parent
EXEMPLOS = PASTA / 'exemplos'
OUTRA_EMPRESA = 'Outra empresa…'

# ──────────────────────────────────────────────
#  Configuração da Página
# ──────────────────────────────────────────────
st.set_page_config(page_title="UnifyData · Fechamento de consumo", page_icon=str(PASTA / 'assets' / 'favicon.png'), layout="centered")

st.markdown("""
<style>
    /* Cores que acompanham o tema ativo do Streamlit (ele declara color-scheme em .stApp) */
    .stApp {
        --ud-texto-2: light-dark(#52525b, #a9a9b3);
        --ud-texto-3: light-dark(#6b6b75, #8a8a94);
        --ud-borda: light-dark(#e5e5e9, #2a2a30);
        --ud-borda-2: light-dark(#d1d1d8, #3a3a42);
        --ud-sup: light-dark(#ffffff, #1a1a1d);
        --ud-acento: light-dark(#0d6e66, #3fcfb9);
        --ud-acento-fundo: light-dark(rgba(13, 110, 102, .09), rgba(63, 207, 185, .12));
        --ud-atencao: light-dark(#9a5800, #f2b544);
        --ud-atencao-fundo: light-dark(#fcf1dd, rgba(242, 181, 68, .13));
        --ud-erro: light-dark(#c4302b, #ff8a84);
        --ud-erro-fundo: light-dark(#fdeeed, rgba(255, 138, 132, .10));
        --ud-azul: light-dark(#3651d4, #8ea2ff);
        --ud-azul-fundo: light-dark(rgba(54, 81, 212, .09), rgba(142, 162, 255, .14));
        --ud-sobre-tinta: light-dark(#ffffff, #141416);
        --ud-fundo: light-dark(#f6f6f7, #141416);
        --pf-afn: light-dark(#6b1f2a, #c44a5a);
        --pf-sys: light-dark(#505050, rgba(255, 255, 255, .32));
    }
    [data-testid="stDecoration"] { display: none; }
    /* O topo tem a cor do fundo: o conteúdo rolado some por baixo dele em vez de passar sob os botões */
    [data-testid="stHeader"] { background: var(--ud-fundo); }
    [data-testid="stMainBlockContainer"] { max-width: 820px; padding-top: 4rem; padding-bottom: 2rem; }
    /* Texto do botão principal com o contraste certo nos dois temas */
    [data-testid^="stBaseButton-primary"]:not(:disabled),
    [data-testid^="stBaseButton-primary"]:not(:disabled) p,
    [data-testid^="stBaseButton-primary"]:not(:disabled) [data-testid="stIconMaterial"] { color: var(--ud-sobre-tinta) !important; }
    /* Envio de arquivos em português */
    [data-testid="stFileUploaderDropzone"] button [data-testid="stMarkdownContainer"] p { font-size: 0; }
    [data-testid="stFileUploaderDropzone"] button [data-testid="stMarkdownContainer"] p::after { content: "Escolher arquivos"; font-size: 15px; }
    [data-testid="stFileUploaderDropzoneInstructions"] span { font-size: 0; }
    [data-testid="stFileUploaderDropzoneInstructions"] span::after { content: "ou arraste os arquivos para cá"; font-size: 13.5px; }
    @media (hover: none) and (pointer: coarse) { [data-testid="stFileUploaderDropzoneInstructions"] { display: none; } }

    .ud-mono { font-family: 'JetBrains Mono', ui-monospace, monospace; }
    .ud-titulo { font-size: 40px; font-weight: 650; letter-spacing: -.035em; line-height: 1.04; margin: 0; }
    .ud-titulo span { color: var(--ud-texto-3); }
    .ud-sub { color: var(--ud-texto-2); font-size: 16px; line-height: 1.55; margin: 14px 0 6px; max-width: 62ch; }
    .ud-rotulo { font: 700 10.5px/1 'JetBrains Mono', monospace; letter-spacing: 1.4px; text-transform: uppercase; color: var(--ud-texto-3); margin: 2px 0 8px; }
    .ud-etapa { font-size: 20px; font-weight: 650; letter-spacing: -.02em; line-height: 1.25; display: flex; align-items: center; gap: 10px; }
    .ud-contador { font: 700 12.5px/1.2 'JetBrains Mono', monospace; padding: 2px 8px; border-radius: 999px; background: var(--ud-atencao-fundo); color: var(--ud-atencao); }
    .ud-texto { color: var(--ud-texto-2); font-size: 14.5px; line-height: 1.55; margin: 6px 0 2px; max-width: 66ch; }
    .ud-faixa { display: flex; flex-wrap: wrap; justify-content: center; gap: 4px 10px; padding: 9px 14px; border-radius: 10px;
                background: var(--ud-acento-fundo); color: var(--ud-texto-2); font-size: 13.5px; text-align: center; }
    .ud-faixa b { color: light-dark(#18181b, #ededef); font-weight: 600; }
    .ud-cartao-topo { display: flex; justify-content: space-between; align-items: flex-start; gap: 16px; }
    .ud-selo { display: inline-flex; align-items: center; height: 22px; padding: 0 8px; border-radius: 999px; font-size: 12px; font-weight: 650; }
    .ud-selo.laranja { background: var(--ud-atencao-fundo); color: var(--ud-atencao); }
    .ud-selo.vermelho { background: var(--ud-erro-fundo); color: var(--ud-erro); }
    .ud-selo.azul { background: var(--ud-azul-fundo); color: var(--ud-azul); }
    .ud-nome { font-size: 15.5px; font-weight: 600; letter-spacing: -.01em; margin: 8px 0 0; overflow-wrap: anywhere; }
    .ud-meta { color: var(--ud-texto-2); font-size: 13.5px; margin: 2px 0 8px; }
    .ud-valor { font-size: 15.5px; font-weight: 650; text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
    .ud-numeros { display: grid; grid-template-columns: 1fr 1fr 1.4fr; border: 1px solid var(--ud-borda); border-radius: 12px; overflow: hidden; margin: 16px 0 4px; }
    .ud-numeros > div { padding: 14px 18px; min-width: 0; }
    .ud-numeros > div + div { border-left: 1px solid var(--ud-borda); }
    .ud-numeros span { display: block; font-size: 13px; color: var(--ud-texto-2); }
    .ud-numeros b { display: block; margin-top: 2px; font-size: 24px; font-weight: 650; letter-spacing: -.02em; font-variant-numeric: tabular-nums; white-space: nowrap; }
    .ud-numeros .destaque { background: var(--ud-acento-fundo); }
    .ud-numeros .destaque span { color: var(--ud-acento); font-weight: 600; }
    .ud-numeros .destaque b { font-size: 28px; }
    .ud-tabela { width: 100%; border-collapse: collapse; font-size: 14px; margin: 6px 0; border: 0 !important; }
    .ud-tabela th, .ud-tabela td { border: 0 !important; background: none !important; }
    .ud-tabela th { text-align: left; font-weight: 600; font-size: 12.5px; color: var(--ud-texto-2); padding: 8px 10px; border-bottom: 1px solid var(--ud-borda-2) !important; white-space: nowrap; }
    .ud-tabela td { padding: 10px; border-bottom: 1px solid var(--ud-borda) !important; }
    .ud-tabela th:first-child, .ud-tabela td:first-child { padding-left: 0; }
    .ud-tabela th:last-child, .ud-tabela td:last-child { padding-right: 0; }
    .ud-tabela .num { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
    .ud-tabela th.origem { white-space: normal; }
    .ud-tabela.varias .origem { display: none; }
    .ud-rolagem { overflow-x: auto; }
    .ud-tabela tfoot td { font-weight: 650; border-bottom: 0 !important; border-top: 1px solid var(--ud-borda-2) !important; }
    .ud-tabela tbody tr:last-child td { border-bottom: 0 !important; }
    .ud-privacidade { text-align: center; color: var(--ud-texto-3); font-size: 13px; margin-top: 1.75rem; }

    /* Lateral: marca, andamento das etapas e assinatura AFN Systems */
    [data-testid="stSidebar"] [data-testid="stElementContainer"]:has(.ud-lateral-rodape) { position: static; }
    [data-testid="stSidebarUserContent"] { padding-bottom: 72px; }
    .ud-marca { display: flex; align-items: center; gap: 10px; font-size: 16px; font-weight: 650; letter-spacing: -.01em; margin: -8px 0 26px; }
    .ud-marca svg { width: 26px; height: 26px; }
    .ud-marca .a, .ud-marca .b { fill: none; stroke-width: 2.6; }
    .ud-marca .a { stroke: var(--ud-texto-2); }
    .ud-marca .b { stroke: var(--ud-acento); }
    .ud-marca .c { fill: var(--ud-acento); }
    /* A lista ocupa a largura toda da lateral: o destaque do passo atual não muda de tamanho entre as etapas */
    [data-testid="stSidebar"] [data-testid="stElementContainer"]:has(.ud-passos),
    [data-testid="stSidebar"] [data-testid="stMarkdown"]:has(.ud-passos),
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"]:has(.ud-passos) { width: 100% !important; }
    .ud-passos { list-style: none; margin: 0; padding: 0; display: flex; flex-direction: column; gap: 2px; width: 100%; }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] ol.ud-passos { display: flex !important; width: 100% !important; margin: 0 !important; padding: 0 !important; }
    .ud-passos li { position: relative; display: grid; grid-template-columns: 28px minmax(0, 1fr); gap: 12px; align-items: center; padding: 9px 10px; border-radius: 10px; margin: 0; }
    .ud-passos li.atual { background: light-dark(#f0f0f2, #222226); }
    .ud-passos li.atual::before { content: ''; position: absolute; left: -12px; top: 10px; bottom: 10px; width: 2px; border-radius: 2px; background: var(--ud-acento); }
    .ud-passos .num { width: 28px; height: 28px; display: grid; place-items: center; border-radius: 50%; border: 1px solid var(--ud-borda-2);
                      font: 700 11.5px/1 'JetBrains Mono', monospace; color: var(--ud-texto-2); }
    .ud-passos li.feito .num { border-color: transparent; background: var(--ud-acento-fundo); color: var(--ud-acento); }
    .ud-passos li.atencao .num { border-color: transparent; background: var(--ud-atencao-fundo); color: var(--ud-atencao); }
    .ud-passos li.atual:not(.feito) .num { border-color: var(--ud-acento); color: var(--ud-acento); }
    .ud-passos b { display: block; font-size: 14px; font-weight: 600; }
    .ud-passos li.bloqueado b { color: var(--ud-texto-3); }
    .ud-passos small { display: block; margin-top: 1px; font-size: 12.5px; color: var(--ud-texto-3); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .ud-lateral-rodape { position: absolute; left: 0; right: 0; bottom: 0; padding: 0 1.5rem; background: inherit; }
    .page-footer { display: flex; align-items: center; justify-content: center; gap: 10px; padding: 14px 0 18px; border-top: 1px solid var(--ud-borda); }
    .page-footer--centro { justify-content: center; margin-top: 2.5rem; }
    .pf-marca { display: flex; align-items: baseline; }
    .pf-afn, .pf-sys, .pf-pipe, .pf-info { font-family: 'JetBrains Mono', monospace; font-weight: 700; }
    .pf-afn { font-size: 10px; letter-spacing: 2px; color: var(--pf-afn); -webkit-text-stroke: .3px color-mix(in srgb, var(--pf-afn) 40%, transparent); }
    .pf-sys { font-size: 10px; letter-spacing: 2px; color: var(--pf-sys); }
    .pf-gap { width: 4px; display: inline-block; }
    .pf-pipe { font-size: 13px; color: var(--pf-sys); }
    .pf-info { font-size: 11px; letter-spacing: .3px; color: var(--pf-sys); }

    @media (max-width: 640px) {
        /* No celular o conteúdo passa por baixo dos botões do topo: a barra ganha fundo, como na versão Web */
        [data-testid="stHeader"] { background: var(--ud-fundo); border-bottom: 1px solid var(--ud-borda); }
        /* Colunas vazias (só servem de espaço no computador) não viram buracos quando as colunas empilham */
        [data-testid="stColumn"]:not(:has([data-testid="stElementContainer"])) { display: none; }
        [data-testid="stMainBlockContainer"] { padding-left: 1rem; padding-right: 1rem; padding-top: 4.5rem; }
        .ud-titulo { font-size: 30px; }
        .ud-sub { font-size: 15px; }
        .ud-tabela .origem { display: none; }
        .ud-numeros { grid-template-columns: 1fr 1fr; }
        .ud-numeros .destaque { grid-column: 1 / -1; border-left: 0 !important; border-top: 1px solid var(--ud-borda); }
        .ud-valor { text-align: left; }
    }
    @media (min-width: 641px) { .page-footer--centro { display: none; } }
</style>
""", unsafe_allow_html=True)

MARCA = ('<svg viewBox="0 0 32 32" aria-hidden="true"><rect class="a" x="4.2" y="4.2" width="15.6" height="15.6" rx="4.2"/>'
         '<rect class="b" x="12.2" y="12.2" width="15.6" height="15.6" rx="4.2"/>'
         '<path class="c" d="M12.2 16.4a4.2 4.2 0 0 1 4.2-4.2h3.4v3.4a4.2 4.2 0 0 1-4.2 4.2h-3.4z"/></svg>')
RODAPE_AFN = ('<div class="pf-marca"><span class="pf-afn">AFN</span><span class="pf-gap"></span><span class="pf-sys">SYSTEMS</span></div>'
              '<span class="pf-pipe" aria-hidden="true">|</span><span class="pf-info">UnifyData</span>')
CHECK = ('<svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.4" '
         'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 12.5 10 17.5 19 7"/></svg>')


# ──────────────────────────────────────────────
#  Estado da Sessão
# ──────────────────────────────────────────────
defaults = {
    'processing_done': False,
    'hr_df': None,
    'linx_df': None,
    'divergences': [],
    'origens': [],
    'periodos': [],
    'avisos': [],
    'ignorados': [],
    'ligados_auto': 0,
    'erros_div': {},
    'escolhas': {},
    'demo': False,
    'erro_envio': None,
    'rodada': 0,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v if not isinstance(v, (list, dict)) else type(v)()


class ArquivoDeExemplo(io.BytesIO):
    """Arquivo da pasta exemplos/ com a mesma interface dos enviados pela tela (name e getvalue)."""
    def __init__(self, caminho):
        super().__init__(caminho.read_bytes())
        self.name = caminho.name


def processar(hr_files, linx_files):
    hr_df, avisos_rh = parse_hr_files(hr_files)
    linx_df, avisos_linx, periodos = parse_linx_files(linx_files)
    if hr_df.empty:
        st.session_state.erro_envio = ('Não consegui tirar funcionários das planilhas do RH. Confira se elas têm as colunas **Nome** e **CPF**.', avisos_rh)
        return
    if linx_df.empty:
        st.session_state.erro_envio = ('Nenhum consumo encontrado. Confira se os arquivos são o relatório *Pendências por responsável* do Linx, exportado em .csv.', avisos_linx)
        return
    divergencias, ligados = encontrar_divergencias(hr_df, linx_df)
    st.session_state.update(
        hr_df=hr_df, linx_df=linx_df, origens=list(dict.fromkeys(linx_df['Origem'])), periodos=periodos,
        divergences=divergencias, avisos=avisos_rh + avisos_linx, ligados_auto=ligados, erro_envio=None, processing_done=True,
    )


def iniciar_demonstracao():
    """Processa os arquivos fictícios da pasta exemplos/; nada é gravado em lugar nenhum."""
    reiniciar()
    st.session_state.demo = True
    processar([ArquivoDeExemplo(p) for p in sorted(EXEMPLOS.glob('Empregados *.xlsx'))],
              [ArquivoDeExemplo(EXEMPLOS / n) for n in ('LOJA.csv', 'RESTAURANTE.csv')])


def reiniciar():
    rodada = st.session_state.get('rodada', 0) + 1
    for k in list(st.session_state.keys()):
        del st.session_state[k]
    for k, v in defaults.items():
        st.session_state[k] = v if not isinstance(v, (list, dict)) else type(v)()
    st.session_state.rodada = rodada


def sair_da_demonstracao():
    reiniciar()
    st.query_params.clear()


def novo_fechamento():
    if st.session_state.demo:
        sair_da_demonstracao()
    else:
        reiniciar()


# Link direto para a demonstração: ?demo=1
if st.query_params.get('demo') and not st.session_state.demo and not st.session_state.processing_done:
    iniciar_demonstracao()


# ──────────────────────────────────────────────
#  Ações das divergências (callbacks: rodam antes do próximo desenho da tela)
# ──────────────────────────────────────────────
def _div(div_id):
    # Clique duplo: o segundo evento chega quando a divergência já saiu da lista
    return next((d for d in st.session_state.divergences if d['id'] == div_id), None)


def _retirar(div_id):
    st.session_state.divergences = [d for d in st.session_state.divergences if d['id'] != div_id]
    st.session_state.erros_div.pop(div_id, None)


def _empresa_escolhida(div_id):
    empresa = st.session_state.get(f'emp_{div_id}')
    if empresa == OUTRA_EMPRESA:
        empresa = st.session_state.get(f'emp_nova_{div_id}', '')
    return ' '.join(str(empresa or '').split())


def _limpar_erro(div_id):
    st.session_state.erros_div.pop(div_id, None)


def _atualizar_repetido(nome_rh):
    """Nome que se repete no RH vira cartão para escolher a linha; se o cartão já existe, o valor dele é refeito."""
    nova = divergencia_repetido(st.session_state.hr_df, st.session_state.linx_df, nome_rh, st.session_state.escolhas)
    if not nova:
        return
    existente = next((d for d in st.session_state.divergences if d['id'] == nova['id']), None)
    if existente:
        existente.update(Valor=nova['Valor'], Origem=nova['Origem'], Opcoes=nova['Opcoes'])
    else:
        st.session_state.divergences.append(nova)


def _vincular_ao_rh(nome_linx, nome_rh, linha_rh=None):
    """
    Passa o consumo do nome do Linx para a pessoa do RH. Com a linha já definida (pelo CPF), o consumo fica preso a ela;
    sem a linha, segue pelo nome, e um nome que se repete no RH vira cartão para escolher de quem descontar.
    """
    if linha_rh is not None:
        fixar_linha(st.session_state.linx_df, nome_linx, linha_rh)
        return
    renomear(st.session_state.linx_df, nome_linx, nome_rh)
    _atualizar_repetido(nome_rh)


def aprovar(div_id):
    d = _div(div_id)
    if not d:
        return
    _retirar(div_id)
    _vincular_ao_rh(d['Linx_Nome'], d['HR_Nome'])


def ignorar(div_id):
    d = _div(div_id)
    if not d:
        return
    st.session_state.ignorados.append({'nome': d['Linx_Nome'], 'valor': d['Valor']})
    df = st.session_state.linx_df
    st.session_state.linx_df = df[~pelo_nome(df, d['Linx_Nome'])]
    _retirar(div_id)


def escolher_linha(div_id):
    d = _div(div_id)
    if not d:
        return
    escolha = st.session_state.get(f'linha_{div_id}')
    if escolha is None:
        st.session_state.erros_div[div_id] = 'Escolha de qual empresa descontar.'
        return
    # A escolha vale para o consumo que está no cartão; o que chegar depois com o mesmo nome pergunta de novo
    fixar_linha(st.session_state.linx_df, d['Linx_Nome'], escolha)
    _retirar(div_id)


def confirmar_cpf(div_id):
    """CPF de alguém do RH: o consumo vai para essa pessoa. CPF novo: cria o registro na empresa escolhida."""
    d = _div(div_id)
    if not d:
        return
    cpf = so_digitos(st.session_state.get(f'cpf_{div_id}', ''))
    hr = st.session_state.hr_df
    match = hr[hr['CPF'] == cpf]
    empresa = _empresa_escolhida(div_id)
    if len(cpf) != 11:
        st.session_state.erros_div[div_id] = 'Digite o CPF com 11 dígitos.'
    elif not match.empty:
        _retirar(div_id)
        # Mesmo CPF em duas empresas: não escolhe sozinho, a revisão pergunta de qual descontar
        _vincular_ao_rh(d['Linx_Nome'], match.iloc[0]['Nome'], int(match.index[0]) if len(match) == 1 else None)
    elif not empresa:
        st.session_state.erros_div[div_id] = ('Digite o nome da nova empresa.' if st.session_state.get(f'emp_{div_id}') == OUTRA_EMPRESA
                                              else 'Esse CPF não está no RH. Escolha a empresa para criar o registro.')
    else:
        novo = pd.DataFrame([{'Nome': d['Linx_Nome'], 'CPF': cpf, 'Empresa': empresa}])
        st.session_state.hr_df = pd.concat([hr, novo], ignore_index=True)
        _retirar(div_id)


# ──────────────────────────────────────────────
#  Pedaços de tela
# ──────────────────────────────────────────────
def cabecalho_etapa(numero, titulo, texto=None, extra=''):
    st.markdown(
        f'<div class="ud-rotulo">Etapa {numero}</div><div class="ud-etapa">{titulo}{extra}</div>'
        + (f'<div class="ud-texto">{texto}</div>' if texto else ''),
        unsafe_allow_html=True,
    )


def cabecalho_cartao(selo, cor, nome, meta, valor):
    st.markdown(
        f'<div class="ud-cartao-topo"><div><span class="ud-selo {cor}">{selo}</span>'
        f'<div class="ud-nome">{escape(nome_exibicao(nome))}</div><div class="ud-meta">{meta}</div></div>'
        f'<div class="ud-valor">{brl(valor)}</div></div>',
        unsafe_allow_html=True,
    )


def campos_cpf_empresa(i, empresas, rotulo_cpf):
    c1, c2 = st.columns([1, 1.6])
    c1.text_input(rotulo_cpf, key=f'cpf_{i}', placeholder='000.000.000-00', max_chars=24, on_change=_limpar_erro, args=(i,))
    c2.selectbox('Empresa', empresas + [OUTRA_EMPRESA], key=f'emp_{i}', index=None, placeholder='Escolha a empresa',
                 on_change=_limpar_erro, args=(i,))
    if st.session_state.get(f'emp_{i}') == OUTRA_EMPRESA:
        st.text_input('Nome da nova empresa', key=f'emp_nova_{i}', placeholder='Ex.: Posto Centro', on_change=_limpar_erro, args=(i,))
    # CPF de alguém do RH: mostra de quem é antes de confirmar, para um erro de digitação não descontar de outra pessoa
    cpf = so_digitos(st.session_state.get(f'cpf_{i}', ''))
    if len(cpf) == 11:
        donos = st.session_state.hr_df[st.session_state.hr_df['CPF'] == cpf]
        if len(donos) == 1:
            st.caption(f'CPF de **{escape(nome_exibicao(donos.iloc[0]["Nome"]))}** · {escape(donos.iloc[0]["Empresa"])}. '
                       'O consumo vai para essa pessoa.')
        elif len(donos) > 1:
            st.caption(f'CPF de **{escape(nome_exibicao(donos.iloc[0]["Nome"]))}**, que está em {len(donos)} empresas do RH.')


# ──────────────────────────────────────────────
#  Interface: abertura
# ──────────────────────────────────────────────
if st.session_state.demo:
    st.markdown('<div class="ud-faixa"><span><b>Demonstração</b> · arquivos fictícios da pasta exemplos/, nada é salvo.</span></div>',
                unsafe_allow_html=True)
    st.button('Sair da demonstração', type='tertiary', icon=':material/close:', on_click=sair_da_demonstracao)

st.markdown("""
<div class="ud-titulo" role="heading" aria-level="1">Nomes diferentes,<br><span>a mesma pessoa.</span></div>
<div class="ud-sub">Cruza os consumos do Linx com as planilhas do RH pelo nome do funcionário, mostra só o que precisa de revisão e gera a planilha de desconto separada por empresa.</div>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────
#  Interface: Upload
# ──────────────────────────────────────────────
if not st.session_state.processing_done:
    rodada = st.session_state.rodada
    st.button('Experimentar com arquivos de exemplo', icon=':material/play_arrow:', on_click=iniciar_demonstracao)

    with st.container(border=True):
        cabecalho_etapa(1, 'Planilhas do RH', 'Uma por empresa (.xls, .xlsx ou .csv). O nome do arquivo vira o nome da empresa, '
                                              'e o arquivo precisa ter as colunas Nome e CPF.')
        hr_files = st.file_uploader('Planilhas do RH', accept_multiple_files=True, type=['xlsx', 'xls', 'csv'],
                                    label_visibility='collapsed', key=f'rh_{rodada}')
        if hr_files:
            st.caption('Empresas: ' + escape(', '.join(dict.fromkeys(empresa_do_arquivo(f.name) for f in hr_files))))

    with st.container(border=True):
        cabecalho_etapa(2, 'Consumos do Linx', 'Relatórios <i>Pendências por responsável</i> em .csv, um por loja. '
                                               'O nome do arquivo vira o nome da coluna na planilha.')
        linx_files = st.file_uploader('Relatórios do Linx', accept_multiple_files=True, type=['csv', 'txt'],
                                      label_visibility='collapsed', key=f'linx_{rodada}')
        if linx_files:
            st.caption('Lojas: ' + escape(', '.join(dict.fromkeys(Path(f.name).stem for f in linx_files))))

    pronto = bool(hr_files and linx_files)
    if st.button('Processar fechamento', type='primary', width='stretch', disabled=not pronto):
        with st.spinner('Cruzando os nomes do Linx com o RH…'):
            processar(hr_files, linx_files)
        if st.session_state.processing_done:
            st.rerun()
    elif not pronto:
        st.caption('Envie as planilhas do RH e os relatórios do Linx para liberar o processamento.')

    if st.session_state.erro_envio:
        mensagem, detalhes = st.session_state.erro_envio
        st.error(mensagem, icon=':material/error:')
        for a in detalhes:
            st.caption(a)


# ──────────────────────────────────────────────
#  Interface: Divergências
# ──────────────────────────────────────────────
if st.session_state.processing_done:
    for aviso in st.session_state.avisos:
        st.warning(aviso, icon=':material/warning:')

    divs = st.session_state.divergences
    n_ligados = st.session_state.ligados_auto
    nota_ligados = (f"{n_ligados} {'nome foi ligado' if n_ligados == 1 else 'nomes foram ligados'} automaticamente "
                    f"(diferença só de acento ou espaço).") if n_ligados else ''

    if divs:
        with st.container(border=True):
            cabecalho_etapa(3, 'Divergências', 'Nomes do Linx que não batem com uma única pessoa do RH. Resolva ou ignore cada um para liberar a planilha.',
                            extra=f'<span class="ud-contador">{len(divs)}</span>')
            if nota_ligados:
                st.caption(nota_ligados)

            empresas = sorted(st.session_state.hr_df['Empresa'].unique(), key=normalizar)

            for div in divs:
                i = div['id']
                with st.container(border=True, key=f'div_{i}'):
                    if div['Tipo'] == 'Nome Semelhante':
                        cabecalho_cartao('Nome parecido', 'laranja', div['Linx_Nome'],
                                         f'No RH: {escape(nome_exibicao(div["HR_Nome"]))} · {div["Score"]}% parecido · {escape(div["Origem"])}',
                                         div['Valor'])
                        b1, b2, _ = st.columns([1.3, 1, 1.6])
                        b1.button('Aprovar vínculo', key=f'apv_{i}', type='primary', width='stretch', on_click=aprovar, args=(i,))
                        b2.button('Ignorar', key=f'ign_{i}', width='stretch', on_click=ignorar, args=(i,))
                        with st.expander('Não é a mesma pessoa?'):
                            st.caption('Informe o CPF. Se ele estiver no RH, o consumo vai para essa pessoa; se não estiver, '
                                       'escolha a empresa para criar o registro.')
                            campos_cpf_empresa(i, empresas, 'CPF')
                            st.button('Confirmar', key=f'cpf_btn_{i}', on_click=confirmar_cpf, args=(i,))

                    elif div['Tipo'] == 'Repetido':
                        cabecalho_cartao(f'Em {len(div["Opcoes"])} linhas do RH', 'azul', div['Linx_Nome'],
                                         f'O mesmo nome aparece mais de uma vez nas planilhas do RH · {escape(div["Origem"])}',
                                         div['Valor'])
                        rotulos = dict(div['Opcoes'])
                        st.radio('De quem é esse consumo?', list(rotulos), key=f'linha_{i}', index=None,
                                 format_func=lambda k, r=rotulos: r[k], on_change=_limpar_erro, args=(i,))
                        b1, b2, _ = st.columns([1.3, 1, 1.6])
                        b1.button('Confirmar', key=f'esc_{i}', type='primary', width='stretch', on_click=escolher_linha, args=(i,))
                        b2.button('Ignorar', key=f'ign_rep_{i}', width='stretch', on_click=ignorar, args=(i,))

                    else:
                        cabecalho_cartao('Fora do RH', 'vermelho', div['Linx_Nome'],
                                         f'Não aparece em nenhuma planilha do RH · {escape(div["Origem"])}', div['Valor'])
                        campos_cpf_empresa(i, empresas, 'CPF')
                        b1, b2, _ = st.columns([1.3, 1, 1.6])
                        b1.button('Criar registro', key=f'create_nf_{i}', type='primary', width='stretch', on_click=confirmar_cpf, args=(i,))
                        b2.button('Ignorar', key=f'ign_nf_{i}', width='stretch', on_click=ignorar, args=(i,))

                    if i in st.session_state.erros_div:
                        st.error(st.session_state.erros_div[i], icon=':material/error:')

    # ──────────────────────────────────────────────
    #  Interface: Download
    # ──────────────────────────────────────────────
    else:
        merged  = consolidar(st.session_state.hr_df, st.session_state.linx_df, st.session_state.origens, st.session_state.escolhas)
        origens = st.session_state.origens
        periodo = '; '.join(st.session_state.periodos)
        com_consumo = merged[merged['Total Desconto'] > 0]
        total = float(merged['Total Desconto'].sum())
        grupos = (merged.groupby('Empresa')
                  .agg(**{'Pessoas': ('Total Desconto', lambda s: int((s > 0).sum())), 'Total': ('Total Desconto', 'sum')},
                       **{f'o_{k}': (f'Total__{o}', 'sum') for k, o in enumerate(origens)})
                  .reset_index())
        grupos = grupos.iloc[sorted(range(len(grupos)), key=lambda k: normalizar(grupos.loc[k, 'Empresa']))]
        ign = st.session_state.ignorados
        nota_ignorados = (f'{len(ign)} {"divergência ignorada ficou" if len(ign) == 1 else "divergências ignoradas ficaram"} fora '
                          f'({brl(sum(x["valor"] for x in ign))}): {", ".join(nome_exibicao(x["nome"]) for x in ign)}.') if ign else ''

        with st.container(border=True):
            if total == 0:
                cabecalho_etapa(4, 'Nada para descontar', 'Nenhum consumo ficou ligado a funcionários do RH, então não há planilha a gerar. '
                                                          'Clique em Novo fechamento para recomeçar.')
                for nota in (nota_ligados, nota_ignorados):
                    if nota:
                        st.caption(nota)
                st.button('Novo fechamento', icon=':material/refresh:', type='tertiary', on_click=novo_fechamento)
            else:
                cabecalho_etapa(4, 'Fechamento pronto',
                                f'Consumos de {escape(lista_por_extenso(origens))}{f", vencimentos de {escape(periodo)}" if periodo else ""}, '
                                'cruzados com as planilhas do RH.')

                st.markdown(f"""
                <div class="ud-numeros">
                  <div><span>Com consumo</span><b>{len(com_consumo)}</b></div>
                  <div><span>Empresas</span><b>{len(grupos)}</b></div>
                  <div class="destaque"><span>Total a descontar</span><b>{brl(total)}</b></div>
                </div>""", unsafe_allow_html=True)

                linhas = ''.join(
                    f'<tr><td>{escape(g.Empresa)}</td><td class="num">{g.Pessoas}</td>'
                    + ''.join(f'<td class="num origem">{brl(getattr(g, f"o_{k}"))}</td>' for k in range(len(origens)))
                    + f'<td class="num">{brl(g.Total)}</td></tr>'
                    for g in grupos.itertuples())
                st.markdown(f"""
                <div class="ud-rolagem"><table class="ud-tabela{' varias' if len(origens) >= 4 else ''}">
                  <thead><tr><th>Empresa</th><th class="num">Pessoas</th>
                    {''.join(f'<th class="num origem">{escape(o)}</th>' for o in origens)}<th class="num">Total</th></tr></thead>
                  <tbody>{linhas}</tbody>
                  <tfoot><tr><td>Total</td><td class="num">{len(com_consumo)}</td>
                    {''.join(f'<td class="num origem">{brl(grupos[f"o_{k}"].sum())}</td>' for k in range(len(origens)))}
                    <td class="num">{brl(total)}</td></tr></tfoot>
                </table></div>""", unsafe_allow_html=True)

                sem_consumo = len(merged) - len(com_consumo)
                notas = [nota_ligados]
                if sem_consumo:
                    notas.append(f'{sem_consumo} {"funcionário" if sem_consumo == 1 else "funcionários"} do RH sem consumo também '
                                 f'{"entra" if sem_consumo == 1 else "entram"} na planilha, em cinza.')
                notas.append(nota_ignorados)
                for nota in filter(None, notas):
                    st.caption(nota)

                # Relatório de impressão: dentro de um iframe, o botão imprime só o relatório
                def tabela_relatorio(emp_df):
                    # Cabeçalho em dois níveis: a loja em cima, Vencido | Total embaixo (cabe mesmo com lojas de nome longo)
                    lojas = ''.join(f'<th colspan="2" class="loja">{escape(o)}</th>' for o in origens)
                    sub = '<th class="num sub ini">Vencido</th><th class="num sub">Total</th>' * len(origens)
                    corpo = ''.join(
                        f'<tr class="{"zerado" if r["Total Desconto"] == 0 else ""}"><td>{escape(r["Nome"])}</td><td>{escape(formatar_cpf(r["CPF"]))}</td>'
                        + ''.join(f'<td class="num ini">{brl(r[f"Vencido__{o}"])}</td><td class="num">{brl(r[f"Total__{o}"])}</td>' for o in origens)
                        + f'<td class="num ini">{brl(r["Total Desconto"])}</td></tr>'
                        for _, r in ordenar(emp_df).iterrows())
                    rodape = ''.join(f'<td class="num ini">{brl(emp_df[f"Vencido__{o}"].sum())}</td><td class="num">{brl(emp_df[f"Total__{o}"].sum())}</td>' for o in origens)
                    return (f'<table><thead><tr><th rowspan="2" class="pessoa">Funcionário</th><th rowspan="2" class="cpf">CPF</th>{lojas}'
                            f'<th rowspan="2" class="num ini">Total desconto</th></tr><tr>{sub}</tr></thead>'
                            f'<tbody>{corpo}</tbody><tfoot><tr><td colspan="2">Total da empresa</td>{rodape}'
                            f'<td class="num ini">{brl(emp_df["Total Desconto"].sum())}</td></tr></tfoot></table>')

                hoje = date.today().strftime('%d/%m/%Y')
                relatorio = (
                    f'<div class="cab"><div><h1>Fechamento de consumo</h1><p>{escape(" · ".join(origens))}'
                    f'{f" · Vencimentos de {escape(periodo)}" if periodo else ""}</p></div>'
                    f'<div class="meta">UnifyData<br>{hoje}</div></div>'
                    f'<div class="resumo"><div><span>Com consumo</span><b>{len(com_consumo)}</b></div>'
                    f'<div><span>Empresas</span><b>{len(grupos)}</b></div><div><span>Total a descontar</span><b>{brl(total)}</b></div></div>'
                    + ''.join(f'<section><h2>{escape(emp)}</h2>{tabela_relatorio(merged[merged["Empresa"] == emp])}</section>'
                              for emp in grupos['Empresa'])
                    + f'<div class="rodape"><span>Gerado pelo UnifyData em {hoje}. {escape(nota_ignorados) or "Sem divergências ignoradas."}</span>'
                      f'<span class="assin"><b>AFN</b> SYSTEMS</span></div>')

                paisagem = '@page { size: A4 landscape; } th.pessoa { width: auto; }' if len(origens) >= 3 else ''
                botao_imprimir = f"""<!doctype html><html><head><meta charset="utf-8">
<link href="https://fonts.googleapis.com/css2?family=Instrument+Sans:wght@400;500;600;700&family=JetBrains+Mono:wght@700&display=swap" rel="stylesheet">
<style>
  :root {{ --fundo: #f6f6f7; --botao: #ffffff; --borda: #d1d1d8; --texto: #18181b; --destaque: #18181b; }}
  html, body {{ margin: 0; background: var(--fundo); font-family: 'Instrument Sans', system-ui, sans-serif; }}
  button {{ width: 100%; height: 38px; display: inline-flex; align-items: center; justify-content: center; gap: 8px;
           border: 1px solid var(--borda); border-radius: 10px; background: var(--botao); color: var(--texto);
           font: 400 15px 'Instrument Sans', system-ui, sans-serif; cursor: pointer; }}
  button:hover {{ border-color: var(--destaque); color: var(--destaque); }}
  button:focus-visible {{ outline: 2px solid var(--destaque); outline-offset: -2px; }}
  button svg {{ width: 17px; height: 17px; fill: none; stroke: currentColor; stroke-width: 1.75; stroke-linecap: round; stroke-linejoin: round; }}
  .relatorio {{ display: none; }}
  @media print {{
    @page {{ size: A4; margin: 14mm 12mm;
             @bottom-right {{ content: counter(page) " / " counter(pages); font: 7.5pt 'JetBrains Mono', monospace; color: #888; }} }} {paisagem}
    html, body {{ background: #fff; }}
    button {{ display: none; }}
    .relatorio {{ display: block; font-size: 9.5pt; line-height: 1.4; color: #111; }}
  }}
  .cab {{ display: flex; justify-content: space-between; align-items: flex-end; padding-bottom: 8pt; border-bottom: 1.5pt solid #111; margin-bottom: 12pt; }}
  h1 {{ font-size: 15pt; margin: 0; }} .cab p {{ margin: 2pt 0 0; color: #444; }} .meta {{ text-align: right; color: #444; }}
  .resumo {{ display: flex; gap: 22pt; margin-bottom: 14pt; }} .resumo span {{ display: block; color: #555; font-size: 8.5pt; }} .resumo b {{ font-size: 12pt; }}
  /* A empresa que cabe numa página não se divide; o total aparece uma vez só, no fim da tabela */
  section {{ margin-bottom: 14pt; break-inside: avoid; }} h2 {{ font-size: 11pt; margin: 0 0 5pt; break-after: avoid; }} tr {{ break-inside: avoid; }}
  table {{ width: 100%; border-collapse: collapse; }} tfoot {{ display: table-row-group; }}
  th {{ background: #58111a; color: #fff; text-align: left; vertical-align: bottom; padding: 4pt 5pt; font-size: 8.5pt; font-weight: 600; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
  th.loja {{ text-align: center; padding-bottom: 2pt; border-bottom: .5pt solid rgba(255,255,255,.35); border-left: .5pt solid rgba(255,255,255,.2); }}
  th.sub {{ padding-top: 2pt; font-size: 8pt; font-weight: 500; }} th.pessoa {{ width: 30%; }} th.cpf {{ width: 27mm; }}
  th.ini {{ border-left: .5pt solid rgba(255,255,255,.2); }} td.ini {{ border-left: .5pt solid #e2e2e2; }}
  td {{ padding: 3.5pt 5pt; border-bottom: .5pt solid #ccc; }} .num {{ text-align: right; white-space: nowrap; }}
  td:first-child {{ min-width: 56mm; }} td:nth-child(2), td.num {{ white-space: nowrap; }}
  tfoot td {{ font-weight: 700; border-bottom: 0; border-top: 1pt solid #111; }} .zerado td {{ color: #888; }}
  .rodape {{ display: flex; justify-content: space-between; gap: 12pt; margin-top: 16pt; color: #666; font-size: 8pt; }}
  .assin {{ font-family: 'JetBrains Mono', monospace; font-weight: 700; letter-spacing: 1.5px; color: #505050; white-space: nowrap; }} .assin b {{ color: #6b1f2a; }}
</style></head><body>
<button type="button" onclick="window.print()"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M7 8V3.5h10V8M7 17H5a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><path d="M7 14h10v6.5H7z"/></svg>Imprimir ou PDF</button>
<div class="relatorio">{relatorio}</div>
<script>
  // Copia as cores do botão CSV ao lado: acompanha o tema mesmo quando ele troca sem recarregar a página
  function copiarTema() {{
    try {{
      const doc = window.parent.document;
      const vizinho = doc.querySelector('[data-testid="stBaseButton-secondary"]');
      const app = doc.querySelector('.stApp');
      if (!vizinho || !app || vizinho.matches(':hover, :focus-visible')) return;
      const v = getComputedStyle(vizinho), r = document.documentElement.style;
      r.setProperty('--fundo', getComputedStyle(app).backgroundColor);
      r.setProperty('--botao', v.backgroundColor);
      r.setProperty('--borda', v.borderTopColor);
      r.setProperty('--texto', v.color);
      const primario = doc.querySelector('[data-testid="stBaseButton-primary"]');
      if (primario) r.setProperty('--destaque', getComputedStyle(primario).backgroundColor);
    }} catch (e) {{}}
  }}
  copiarTema();
  setInterval(copiarTema, 800);
</script>
</body></html>"""

                st.markdown('<div style="height: 6px"></div>', unsafe_allow_html=True)
                sufixo = date.today().isoformat()
                a1, a2, a3 = st.columns([1.3, 0.8, 1.3])
                a1.download_button(
                    label='Baixar Excel', data=gerar_excel(merged, origens), file_name=f'Fechamento_UnifyData_{sufixo}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    type='primary', icon=':material/download:', width='stretch'
                )
                a2.download_button(
                    label='CSV', data=gerar_csv(merged, origens), file_name=f'Fechamento_UnifyData_{sufixo}.csv',
                    mime='text/csv', icon=':material/table:', width='stretch'
                )
                with a3:
                    st.iframe(botao_imprimir, height=38)

                st.button('Novo fechamento', icon=':material/refresh:', type='tertiary', on_click=novo_fechamento)

st.markdown('<div class="ud-privacidade">Os arquivos são lidos só nesta sessão e não ficam gravados em lugar nenhum.</div>'
            f'<footer class="page-footer page-footer--centro" aria-label="AFN Systems, UnifyData">{RODAPE_AFN}</footer>',
            unsafe_allow_html=True)


# ──────────────────────────────────────────────
#  Lateral: andamento das etapas (desenhada por último, com o estado já atualizado)
# ──────────────────────────────────────────────
def passo(numero, titulo, resumo, estado, atual):
    classes = ' '.join(filter(None, [estado, 'atual' if atual else '']))
    marcador = CHECK if estado == 'feito' else numero
    return f'<li class="{classes}"><span class="num">{marcador}</span><span><b>{titulo}</b><small>{escape(resumo)}</small></span></li>'


ss = st.session_state
rodada = ss.rodada
feito = ss.processing_done
n_rh = len(ss.get(f'rh_{rodada}') or [])
n_linx = len(ss.get(f'linx_{rodada}') or [])
if feito:
    n_pessoas, n_empresas = len(ss.hr_df), ss.hr_df['Empresa'].nunique()
    resumo_rh = (f"{n_pessoas} {'pessoa' if n_pessoas == 1 else 'pessoas'} · "
                 f"{n_empresas} {'empresa' if n_empresas == 1 else 'empresas'}")
    resumo_linx = f"{len(ss.origens)} {'relatório' if len(ss.origens) == 1 else 'relatórios'}"
    n_div = len(ss.divergences)
    if n_div:
        resumo_planilha = 'Depois das divergências'
    else:
        total_lateral = float(consolidar(ss.hr_df, ss.linx_df, ss.origens, ss.escolhas)['Total Desconto'].sum())
        resumo_planilha = brl(total_lateral) if total_lateral else 'Nada para descontar'
else:
    resumo_rh = f"{n_rh} {'planilha' if n_rh == 1 else 'planilhas'}" if n_rh else 'Nenhuma planilha'
    resumo_linx = f"{n_linx} {'relatório' if n_linx == 1 else 'relatórios'}" if n_linx else 'Nenhum relatório'
    n_div = 0
atual = 'rh' if not feito and not n_rh else 'linx' if not feito else 'div' if n_div else 'planilha'

with st.sidebar:
    st.markdown(
        f'<div class="ud-marca">{MARCA}<span>UnifyData</span></div>'
        '<div class="ud-rotulo">Fechamento</div>'
        '<ol class="ud-passos">'
        + passo(1, 'Planilhas do RH', resumo_rh, 'feito' if feito else '', atual == 'rh')
        + passo(2, 'Consumos do Linx', resumo_linx, 'feito' if feito else '', atual == 'linx')
        + passo(3, 'Divergências', (f'{n_div} a revisar' if n_div else 'Nenhuma pendente') if feito else 'Depois de processar',
                ('atencao' if n_div else 'feito') if feito else 'bloqueado', atual == 'div')
        + passo(4, 'Planilha', resumo_planilha if feito else 'Depois de processar',
                ('feito' if not n_div else 'bloqueado') if feito else 'bloqueado', atual == 'planilha')
        + '</ol>',
        unsafe_allow_html=True,
    )
    st.markdown(f'<div class="ud-lateral-rodape"><footer class="page-footer" aria-label="AFN Systems, UnifyData">{RODAPE_AFN}</footer></div>',
                unsafe_allow_html=True)
