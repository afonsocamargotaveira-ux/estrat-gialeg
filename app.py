import datetime as dt
from pathlib import Path

import streamlit as st

from gerador import gerar, nome_arquivo
import io

from tecad import Tesauro, ler_xlsx, montar_or, salvar_base

st.set_page_config(page_title="EstratégiaLeg", page_icon="🔎", layout="centered")

LOGO = Path(__file__).with_name("logo.svg").read_text(encoding="utf-8")

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
:root{--azul:#15803D;--marinho:#064E3B;--laranja:#16A34A;--laranja2:#4ADE80;--cinza:#F1F7F3;--cinza2:#DAEADF;--texto:#14261C;--muted:#6B7280;}
html, body, [class*="css"], .stApp {font-family:'Plus Jakarta Sans',system-ui,sans-serif;color:var(--texto);}
.stApp{background:var(--cinza);}
#MainMenu, footer, header[data-testid="stHeader"]{visibility:hidden;height:0;}
.block-container{max-width:860px;padding-top:1.2rem;padding-bottom:4rem;}

.topo{display:flex;align-items:center;gap:14px;padding:6px 2px 18px;}
.topo svg{width:54px;height:54px;filter:drop-shadow(0 6px 10px rgba(6,78,59,.25));}
.marca{font-size:2rem;font-weight:800;letter-spacing:-.03em;line-height:1;color:var(--marinho);}
.marca b{color:var(--laranja);font-weight:800;}
.sub{color:var(--muted);font-size:.85rem;margin-top:5px;font-weight:500;}

.hero{background:linear-gradient(120deg,var(--marinho) 0%,var(--azul) 70%);border-radius:20px;padding:26px 28px;color:#fff;
 position:relative;overflow:hidden;margin-bottom:22px;box-shadow:0 14px 30px -14px rgba(6,78,59,.55);}
.hero:after{content:"";position:absolute;right:-50px;top:-60px;width:210px;height:210px;border-radius:50%;
 background:radial-gradient(circle at 30% 30%,var(--laranja2),var(--laranja));opacity:.95;}
.hero h1{font-size:1.55rem;font-weight:800;margin:0 0 6px;letter-spacing:-.02em;color:#fff;max-width:75%;}
.hero p{margin:0;color:#D5F2E0;font-size:.95rem;max-width:70%;}

.sec{display:flex;align-items:center;gap:12px;margin:26px 0 10px;}
.sec .n{width:32px;height:32px;border-radius:10px;background:var(--laranja);color:#fff;font-weight:800;display:grid;place-items:center;
 box-shadow:0 6px 12px -4px rgba(22,163,74,.55);}
.sec h3{margin:0;padding:0!important;font-size:1.15rem;font-weight:800;color:var(--marinho);letter-spacing:-.01em;}
.sec small{line-height:1.3;color:var(--muted);font-weight:500;display:block;margin-top:2px;}

div[data-testid="stVerticalBlockBorderWrapper"]{background:#fff;border:1px solid var(--cinza2)!important;border-radius:18px!important;
 box-shadow:0 8px 24px -16px rgba(6,78,59,.35);}
.seq{display:inline-block;background:#E3F5EA;color:var(--azul);font-weight:700;font-size:.78rem;padding:4px 12px;border-radius:999px;margin-bottom:6px;}

label, .stTextInput label p, .stTextArea label p, .stDateInput label p, .stNumberInput label p, .stMultiSelect label p, .stFileUploader label p
{font-weight:600!important;color:var(--marinho)!important;font-size:.88rem!important;}
input, textarea, div[data-baseweb="select"]>div, div[data-baseweb="input"]>div{border-radius:12px!important;}
.stTextInput input, .stTextArea textarea, .stDateInput input, .stNumberInput input{background:#F6FBF8!important;border:1px solid var(--cinza2)!important;}
.stTextInput input:focus, .stTextArea textarea:focus{border-color:var(--azul)!important;box-shadow:0 0 0 3px rgba(21,128,61,.18)!important;}
span[data-baseweb="tag"]{background:#E3F5EA!important;color:var(--azul)!important;border-radius:8px!important;}
section[data-testid="stFileUploaderDropzone"]{background:#F6FBF8;border:1.5px dashed #95D5AF;border-radius:14px;}
section[data-testid="stFileUploaderDropzone"] button{border-radius:10px;border-color:var(--azul);color:var(--azul);}

.stButton>button[kind="primary"]{background:linear-gradient(135deg,var(--laranja2),var(--laranja));border:0;color:#fff;font-weight:700;
 border-radius:14px;padding:.7rem 1.6rem;box-shadow:0 12px 20px -10px rgba(22,163,74,.7);transition:transform .15s;}
.stButton>button[kind="primary"]:hover{transform:translateY(-1px);filter:brightness(1.04);color:#fff;}
.stButton>button[kind="primary"]:disabled{background:#CBD2DE;box-shadow:none;color:#fff;}
.stDownloadButton>button{background:linear-gradient(135deg,var(--azul),var(--marinho));color:#fff;border:0;font-weight:700;border-radius:14px;
 padding:.7rem 1.6rem;box-shadow:0 12px 20px -10px rgba(6,78,59,.7);}
.stDownloadButton>button:hover{color:#fff;filter:brightness(1.1);}
div[data-testid="stAlert"]{border-radius:14px;}
hr{border-color:var(--cinza2);}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

st.markdown(
    f"""<div class="topo">{LOGO}<div><div class="marca">Estratégia<b>Leg</b></div>
    <div class="sub">Estratégias de pesquisa de proposições no Sileg</div></div></div>
    <div class="hero"><h1>Monte sua estratégia de pesquisa em minutos</h1>
    <p>Preencha os campos e baixe o documento no modelo oficial, com a formatação original.</p></div>""",
    unsafe_allow_html=True,
)


def secao(n, titulo, legenda=""):
    leg = f"<small>{legenda}</small>" if legenda else ""
    st.markdown(f'<div class="sec"><div class="n">{n}</div><div><h3>{titulo}</h3>{leg}</div></div>',
                unsafe_allow_html=True)


@st.cache_resource(show_spinner="Carregando o Tesauro da Câmara…")
def carregar_tesauro():
    return Tesauro()


@st.cache_resource(show_spinner="Lendo o arquivo do Tesauro (pode levar alguns segundos)…")
def tesauro_do_upload(conteudo: bytes):
    return Tesauro(ler_xlsx(io.BytesIO(conteudo)))


REL = {"Termo e sinônimos (UP)": ("base", "sinonimos"), "Termos específicos (TE)": ("especificos",),
       "Termos relacionados (TR)": ("relacionados",), "Termos genéricos (TG)": ("genericos",)}
REL_PADRAO = {1: ["Termo e sinônimos (UP)", "Termos específicos (TE)", "Termos relacionados (TR)"]}


def _h(*partes):
    import hashlib
    return hashlib.md5("|".join(map(str, partes)).encode()).hexdigest()[:8]


def inserir_sugestoes(i, chave):
    sel = st.session_state.get(chave, [])
    st.session_state[f"texto{i}"] = montar_or(sel)


CAMPOS_PADRAO = ["ementa", "explicação da ementa", "indexação"]

# ---------- 1. Identificação ----------
secao(1, "Identificação")
with st.container(border=True):
    tema = st.text_input("Tema da pesquisa", key="tema", placeholder="Digite o tema da pesquisa",
                         help="Aparece no título, no cabeçalho e no rodapé do documento.")
    c1, c2 = st.columns(2)
    data = c1.date_input("Data da atualização", value=dt.date.today(), format="DD/MM/YYYY")
    protocolo = c2.text_input("Protocolo Prisma do último resultado enviado", placeholder="Ex.: 2025/000123",
                              help="Número do protocolo da planilha com o último resultado enviado.")

# ---------- 2. Sequências ----------
secao(2, "Sequências de termos", "Campo Pesquisa textual da cesta de união (OR)")
n = st.number_input("Quantas sequências de termos? (quantas vezes duplicar o campo Pesquisa textual)",
                    min_value=1, max_value=30, value=1, step=1)

# --- Tesauro: base embutida ou arquivo enviado hoje ---
with st.expander("📥 Atualizar o Tesauro (tecadTermos.xlsx)"):
    arq_tecad = st.file_uploader(
        "Envie o arquivo tecadTermos.xlsx mais recente", type=["xlsx"], key="up_tecad",
        help="Baixe em dadosabertos.camara.leg.br (Tesauro/TECAD). O arquivo enviado passa a ser usado nas "
             "sugestões desta sessão.")
    salvar = st.checkbox("Guardar como base padrão para as próximas vezes", value=False, key="salvar_tecad")

tes = None
try:
    if arq_tecad is not None:
        tes = tesauro_do_upload(arq_tecad.getvalue())
        st.caption(f"Tesauro em uso: **{arq_tecad.name}** ({len(tes.itens):,} termos).".replace(",", "."))
        if salvar and st.session_state.get("tecad_salvo") != arq_tecad.file_id:
            salvar_base(tes.itens)
            st.session_state["tecad_salvo"] = arq_tecad.file_id
            st.toast("Base padrão atualizada.")
    elif tema.strip():
        tes = carregar_tesauro()
except Exception as e:
    st.warning(f"Tesauro indisponível: {e}")

# --- bases do Tesauro ---
bases = []
if tes is not None and not tema.strip():
    st.caption("Digite o tema (seção 1) para ver as sugestões do Tesauro.")
if tes is not None and tema.strip():
    with st.container(border=True):
        achados = tes.buscar(tema.strip())
        padrao = [t for t in tes.exatos(tema.strip()) if t in achados] or achados[:1]
        bases = st.multiselect(
            "Termos do Tesauro da Câmara usados como base para as sugestões", achados, default=padrao,
            key=f"bases_{_h(tema.strip())}",
            help="Termos do Tesauro (TECAD) que correspondem ao tema. Escolha um ou mais.")
        if not achados:
            st.info("Não encontrei esse tema no Tesauro. Tente outra grafia ou digite os termos manualmente.")
rel = tes.relacoes(bases) if tes is not None and bases else None

sequencias = []
for i in range(1, int(n) + 1):
    with st.container(border=True):
        st.markdown(f'<span class="seq">{i}ª sequência</span>', unsafe_allow_html=True)
        campos = st.multiselect("Campos a marcar", CAMPOS_PADRAO, default=CAMPOS_PADRAO, key=f"campos{i}",
                                help="Aparece como: “(marcar ementa, explicação da ementa e indexação)”.")
        texto = st.text_area("Sequência de termos", key=f"texto{i}", height=110,
                             help="Termos separados por OR. Um parágrafo por linha.")
        sequencias.append({"campos": campos, "texto": texto})
        if rel:
            with st.expander("💡 Sugestões do Tesauro da Câmara", expanded=(i == 1)):
                tipos = st.multiselect("Tipos de termo", list(REL), key=f"tipos{i}",
                                       default=REL_PADRAO.get(i, []),
                                       help="UP = sinônimos, TE = específicos, TR = relacionados, TG = genéricos.")
                opcoes = []
                for t in tipos:
                    for chave in REL[t]:
                        opcoes += rel[chave]
                opcoes = list(dict.fromkeys(opcoes))
                chave_sel = f"sel{i}_{_h(tema.strip(), bases, tipos)}"
                if opcoes:
                    st.multiselect("Termos sugeridos (desmarque os que não quiser)", opcoes, default=opcoes,
                                   key=chave_sel)
                    st.button("Inserir na sequência", key=f"ins{i}", on_click=inserir_sugestoes,
                              args=(i, chave_sel))
                elif not tipos:
                    st.caption("Selecione ao menos um tipo de termo para ver as sugestões.")
                else:
                    st.caption("Nenhum termo deste tipo para os termos-base escolhidos.")

with st.container(border=True):
    imgs_estrutura = st.file_uploader("Imagem de como ficará a estrutura da pesquisa:",
                                      type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="img_estrutura")

# ---------- 3. Exclusão ----------
secao(3, "Lista de exclusão", "Proposições ativas que não tratam do tema")
with st.container(border=True):
    lista_excl = st.text_area("Proposições a excluir", placeholder="PL 1234/2020, PL 55/2021, PEC 7/2019",
                              help="Lista na horizontal, separada por vírgulas (quebras de linha também são aceitas).")

secao(4, "ATENÇÃO: NÃO EXCLUIR", "Termos que não devem levar à exclusão")
with st.container(border=True):
    nao_excl = st.text_area("Lista de termos", key="nao_excl", height=120, help="Um termo por linha.")

secao(5, "ATENÇÃO: TEMAS PARA EXCLUIR",
      "Complementa: “EXCLUIR da planilha eventuais proposições que apareçam na pesquisa e que:”")
with st.container(border=True):
    temas_excl = st.text_area("Lista de termos/temas", key="temas_excl", height=120, help="Um termo por linha.")

# ---------- Geração ----------
st.markdown("<br>", unsafe_allow_html=True)
pendencias = []
if not tema.strip():
    pendencias.append("informe o tema")
if not protocolo.strip():
    pendencias.append("informe o protocolo Prisma")
for i, s in enumerate(sequencias, 1):
    if not s["texto"].strip():
        pendencias.append(f"{i}ª sequência sem termos")
if not imgs_estrutura:
    pendencias.append("falta a imagem da estrutura da pesquisa")
if pendencias:
    st.warning("Pendências: " + "; ".join(pendencias) + ".")

b1, b2 = st.columns([1, 1])
if b1.button("Gerar documento", type="primary", disabled=not tema.strip()):
    try:
        dados = gerar(tema.strip(), data.strftime("%d/%m/%Y"), protocolo.strip(), sequencias,
                      lista_excl, nao_excl, temas_excl, [f.getvalue() for f in imgs_estrutura])
        st.session_state["docx"] = (dados, nome_arquivo(tema))
    except Exception as e:
        st.error(f"Erro ao gerar o documento: {e}")

if "docx" in st.session_state:
    dados, nome = st.session_state["docx"]
    b2.download_button("⬇ Baixar .docx", dados, file_name=nome,
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    st.success("Documento gerado.")
