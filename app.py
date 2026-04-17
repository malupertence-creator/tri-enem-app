import pandas as pd
import streamlit as st
import plotly.express as px
from io import BytesIO

ARQUIVO_BASE = "base_TRI_ENEM_FINAL_2021_2024_corrigida.xlsx"
AREAS_VALIDAS = ["CN", "CH", "LC", "MT"]
MODO_OFICIAL = "Prova oficial"
MODO_MISTO = "Simulado misto"

CORES_AREA = {
    "CN": "#22C55E",
    "CH": "#3B82F6",
    "LC": "#EC4899",
    "MT": "#F59E0B",
}

COLUNAS_RESULTADO = [
    "Nome", "Turma", "Area", "Tipo_Simulacao", "Ano", "Acertos",
    "Porcentagem_Acertos", "Nota_min", "Nota_media", "Nota_max", "Observacao"
]

st.set_page_config(
    page_title="Simulador TRI ENEM",
    page_icon="📘",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
.block-container {
    padding-top: 1.5rem;
    padding-bottom: 2rem;
    padding-left: 2rem;
    padding-right: 2rem;
    max-width: 1350px;
}

html, body, [class*="css"] {
    font-family: "Segoe UI", sans-serif;
}

.section-card {
    background: #FFFFFF;
    border: 1.5px solid #CBD5E1;
    border-radius: 20px;
    padding: 24px;
    box-shadow: 0 10px 24px rgba(15, 23, 42, 0.10);
    margin-bottom: 1.25rem;
}

.area-badge {
    display: inline-block;
    padding: 0.35rem 0.75rem;
    border-radius: 999px;
    font-size: 0.85rem;
    font-weight: 700;
    margin-right: 0.4rem;
    margin-bottom: 0.4rem;
}

.badge-cn { background: #DCFCE7; color: #166534; }
.badge-ch { background: #DBEAFE; color: #1D4ED8; }
.badge-lc { background: #FCE7F3; color: #BE185D; }
.badge-mt { background: #FEF3C7; color: #B45309; }

.info-box {
    background: #F8FAFC;
    border: 1.5px solid #CBD5E1;
    border-radius: 16px;
    padding: 16px;
    margin-bottom: 16px;
}

.profile-card {
    background: #FFFFFF;
    border: 1.5px solid #CBD5E1;
    border-radius: 18px;
    padding: 18px;
    box-shadow: 0 6px 18px rgba(15, 23, 42, 0.08);
    margin-bottom: 12px;
}

.metric-card-caption {
    color: #475569;
    font-size: 0.92rem;
    margin-top: -6px;
    margin-bottom: 10px;
}

.small-note {
    color: #64748B;
    font-size: 0.92rem;
}

.mode-pill {
    display: inline-block;
    padding: 0.35rem 0.8rem;
    border-radius: 999px;
    font-size: 0.85rem;
    font-weight: 700;
    background: #EDE9FE;
    color: #5B21B6;
    margin-bottom: 0.8rem;
}

.mode-pill-misto {
    background: #FEF3C7;
    color: #92400E;
}

.mode-pill-oficial {
    background: #DBEAFE;
    color: #1D4ED8;
}

div[data-testid="stMetric"] {
    background: #FFFFFF;
    border: 1.5px solid #C7D2FE;
    border-radius: 18px;
    padding: 16px;
    box-shadow: 0 8px 20px rgba(15, 23, 42, 0.08);
}

div[data-testid="stForm"] {
    background: #FFFFFF;
    border: 1.5px solid #CBD5E1;
    border-radius: 20px;
    padding: 22px;
    box-shadow: 0 8px 24px rgba(15, 23, 42, 0.08);
}

.stButton > button,
[data-testid="stDownloadButton"] > button {
    background: linear-gradient(90deg, #7C3AED, #6D28D9);
    color: white;
    border: none;
    border-radius: 12px;
    padding: 0.7rem 1.3rem;
    font-weight: 700;
    box-shadow: 0 8px 18px rgba(124, 58, 237, 0.25);
}

.stButton > button:hover,
[data-testid="stDownloadButton"] > button:hover {
    color: white;
    border: none;
    background: linear-gradient(90deg, #6D28D9, #5B21B6);
}

section[data-testid="stSidebar"] {
    background: #F1F5F9;
    border-right: 1.5px solid #CBD5E1;
}

div[data-testid="stDataFrame"] {
    border: 1.5px solid #CBD5E1;
    border-radius: 16px;
    overflow: hidden;
    box-shadow: 0 6px 18px rgba(15, 23, 42, 0.08);
}

hr {
    border: none;
    border-top: 1.5px solid #CBD5E1;
    margin: 1.5rem 0;
}

div[data-baseweb="input"] input {
    background-color: #FFFFFF !important;
    border: 2px solid #64748B !important;
    border-radius: 10px !important;
    color: #0F172A !important;
    font-weight: 500;
}

div[data-baseweb="select"] > div {
    background-color: #FFFFFF !important;
    border: 2px solid #64748B !important;
    border-radius: 10px !important;
}

div[data-baseweb="input"] input:focus {
    border: 2px solid #7C3AED !important;
    box-shadow: 0 0 0 2px rgba(124, 58, 237, 0.25);
}

label {
    font-weight: 600 !important;
    color: #1E293B !important;
}

.header-box {
    background: linear-gradient(135deg, #6D28D9, #7C3AED);
    border-radius: 18px;
    padding: 28px 32px;
    color: white;
    margin-bottom: 28px;
    box-shadow: 0 10px 24px rgba(0,0,0,0.15);
}

.header-title {
    font-size: 2.2rem;
    font-weight: 700;
    line-height: 1.2;
    margin-bottom: 8px;
}

.header-subtitle {
    font-size: 1rem;
    opacity: 0.92;
}
</style>
""",
    unsafe_allow_html=True,
)

if "resultados" not in st.session_state:
    st.session_state["resultados"] = pd.DataFrame(columns=COLUNAS_RESULTADO)


@st.cache_data
def carregar_base_tri(caminho_arquivo: str) -> pd.DataFrame:
    abas = pd.read_excel(caminho_arquivo, sheet_name=None)
    tri = pd.concat(abas.values(), ignore_index=True)
    tri.columns = tri.columns.str.strip()

    tri["Area"] = tri["Area"].astype(str).str.strip().str.upper()
    tri["Ano"] = pd.to_numeric(tri["Ano"], errors="coerce")
    tri["Acertos"] = pd.to_numeric(tri["Acertos"], errors="coerce")
    tri["Min"] = pd.to_numeric(tri["Min"], errors="coerce")
    tri["Media"] = pd.to_numeric(tri["Media"], errors="coerce")
    tri["Max"] = pd.to_numeric(tri["Max"], errors="coerce")

    tri = tri.dropna(subset=["Area", "Ano", "Acertos"])
    return tri


def obter_resultados() -> pd.DataFrame:
    resultados = st.session_state["resultados"].copy()
    for coluna in COLUNAS_RESULTADO:
        if coluna not in resultados.columns:
            resultados[coluna] = None
    return resultados[COLUNAS_RESULTADO].copy()


def salvar_resultado(novo: dict) -> None:
    resultados = obter_resultados()
    st.session_state["resultados"] = pd.concat(
        [resultados, pd.DataFrame([novo])],
        ignore_index=True
    )


def apagar_resultados() -> None:
    st.session_state["resultados"] = pd.DataFrame(columns=COLUNAS_RESULTADO)


def consultar_tri(tri: pd.DataFrame, area: str, ano: int, acertos: int) -> pd.DataFrame:
    return tri[
        (tri["Area"] == area)
        & (tri["Ano"] == ano)
        & (tri["Acertos"] == acertos)
    ]


def consultar_estimativa_mista(tri: pd.DataFrame, area: str, acertos: int) -> pd.DataFrame:
    return tri[
        (tri["Area"] == area)
        & (tri["Acertos"] == acertos)
    ].copy()


def consolidar_estimativa_mista(consulta: pd.DataFrame) -> dict | None:
    if consulta.empty:
        return None

    anos_base = sorted(consulta["Ano"].dropna().astype(int).unique().tolist())
    minimos = consulta["Min"].dropna()
    medias = consulta["Media"].dropna()
    maximos = consulta["Max"].dropna()

    nota_min = float(minimos.min()) if not minimos.empty else None
    nota_media = float(medias.mean()) if not medias.empty else None
    nota_max = float(maximos.max()) if not maximos.empty else None
    nota_provavel_inferior = float(medias.quantile(0.25)) if len(medias) >= 2 else nota_media
    nota_provavel_superior = float(medias.quantile(0.75)) if len(medias) >= 2 else nota_media
    mediana = float(medias.median()) if not medias.empty else nota_media

    return {
        "anos_base": anos_base,
        "nota_min": nota_min,
        "nota_media": nota_media,
        "nota_max": nota_max,
        "faixa_provavel_min": nota_provavel_inferior,
        "faixa_provavel_max": nota_provavel_superior,
        "mediana": mediana,
        "amostras": len(consulta),
    }


def normalizar_nome(nome: str) -> str:
    return nome.strip().title()


def ano_para_exibicao(valor) -> str:
    if pd.isna(valor):
        return "Misto"
    try:
        return str(int(valor))
    except Exception:
        return str(valor)


def tipo_para_exibicao(valor: str) -> str:
    if not valor:
        return MODO_OFICIAL
    return str(valor)


def ordenar_resultados(df: pd.DataFrame) -> pd.DataFrame:
    temp = df.copy()
    temp["Ano_sort"] = pd.to_numeric(temp["Ano"], errors="coerce").fillna(9999)
    temp = temp.sort_values(by=["Turma", "Nome", "Ano_sort", "Area", "Tipo_Simulacao"])
    return temp.drop(columns=["Ano_sort"])


def obter_melhor_resultado_por_aluno(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df.copy()

    temp = df.copy()
    temp["Nome"] = temp["Nome"].astype(str).str.strip()
    temp["Turma"] = temp["Turma"].astype(str).str.strip()
    temp["Area"] = temp["Area"].astype(str).str.strip().str.upper()
    temp["Ano_ordenacao"] = pd.to_numeric(temp["Ano"], errors="coerce").fillna(9999)

    melhor = (
        temp.sort_values(
            by=["Nome", "Nota_media", "Acertos", "Ano_ordenacao"],
            ascending=[True, False, False, False],
        )
        .groupby("Nome", as_index=False)
        .first()
    )

    return melhor.drop(columns=["Ano_ordenacao"])


def obter_resumo_alunos(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(
            columns=[
                "Nome", "Turma", "Qtd_Provas", "Melhor_Area", "Melhor_Modo", "Melhor_Ano",
                "Melhor_Nota", "Media_Notas", "Media_Acertos",
                "Ultimo_Ano", "Ultima_Area", "Ultimo_Modo", "Ultima_Nota",
            ]
        )

    temp = df.copy()
    temp["Nome"] = temp["Nome"].astype(str).str.strip()
    temp["Turma"] = temp["Turma"].astype(str).str.strip()
    temp["Area"] = temp["Area"].astype(str).str.strip().str.upper()
    temp["Tipo_Simulacao"] = temp["Tipo_Simulacao"].fillna(MODO_OFICIAL)
    temp["Ano_ordenacao"] = pd.to_numeric(temp["Ano"], errors="coerce").fillna(9999)

    linhas = []

    for nome, grupo in temp.groupby("Nome"):
        grupo = grupo.sort_values(by=["Ano_ordenacao", "Area", "Tipo_Simulacao"])

        melhor = grupo.sort_values(
            by=["Nota_media", "Acertos", "Ano_ordenacao"],
            ascending=[False, False, False],
        ).iloc[0]

        ultimo = grupo.sort_values(by=["Ano_ordenacao", "Area"], ascending=[False, True]).iloc[0]

        linhas.append(
            {
                "Nome": nome,
                "Turma": grupo.iloc[0]["Turma"],
                "Qtd_Provas": len(grupo),
                "Melhor_Area": melhor["Area"],
                "Melhor_Modo": melhor["Tipo_Simulacao"],
                "Melhor_Ano": ano_para_exibicao(melhor["Ano"]),
                "Melhor_Nota": round(float(melhor["Nota_media"]), 2),
                "Media_Notas": round(float(grupo["Nota_media"].mean()), 2),
                "Media_Acertos": round(float(grupo["Acertos"].mean()), 2),
                "Ultimo_Ano": ano_para_exibicao(ultimo["Ano"]),
                "Ultima_Area": ultimo["Area"],
                "Ultimo_Modo": ultimo["Tipo_Simulacao"],
                "Ultima_Nota": round(float(ultimo["Nota_media"]), 2),
            }
        )

    return pd.DataFrame(linhas).sort_values(by="Nome").reset_index(drop=True)


def gerar_resumo_por_turma(resultados: pd.DataFrame) -> pd.DataFrame:
    if resultados.empty:
        return pd.DataFrame(
            columns=[
                "Turma", "Qtd_Registros", "Qtd_Alunos",
                "Media_Melhor_Acertos", "Media_Melhor_Porcentagem", "Media_Melhor_Nota_TRI",
            ]
        )

    linhas = []

    for turma, grupo in resultados.groupby("Turma"):
        melhor = obter_melhor_resultado_por_aluno(grupo)

        linhas.append(
            {
                "Turma": turma,
                "Qtd_Registros": len(grupo),
                "Qtd_Alunos": grupo["Nome"].astype(str).str.strip().nunique(),
                "Media_Melhor_Acertos": round(float(melhor["Acertos"].mean()), 2),
                "Media_Melhor_Porcentagem": round(float(melhor["Porcentagem_Acertos"].mean()), 2),
                "Media_Melhor_Nota_TRI": round(float(melhor["Nota_media"].mean()), 2),
            }
        )

    return pd.DataFrame(linhas).sort_values(by="Turma").reset_index(drop=True)


def exibir_cabecalho() -> None:
    st.markdown(
        """
<div class="header-box">
    <div class="header-title">Simulador TRI ENEM</div>
    <div class="header-subtitle">
        Desenvolvido pela professora Maria Luiza Pertence.<br>
    </div>
</div>
""",
        unsafe_allow_html=True,
    )


def exibir_badges_areas() -> None:
    st.markdown(
        """
<div>
    <span class="area-badge badge-cn">CN</span>
    <span class="area-badge badge-ch">CH</span>
    <span class="area-badge badge-lc">LC</span>
    <span class="area-badge badge-mt">MT</span>
</div>
""",
        unsafe_allow_html=True,
    )


def excel_bytes(df: pd.DataFrame) -> BytesIO:
    buffer = BytesIO()
    df.to_excel(buffer, index=False, engine="openpyxl")
    buffer.seek(0)
    return buffer


def exibir_pill_modo(modo: str):
    classe = "mode-pill-oficial" if modo == MODO_OFICIAL else "mode-pill-misto"
    st.markdown(f'<div class="mode-pill {classe}">{modo}</div>', unsafe_allow_html=True)


try:
    tri = carregar_base_tri(ARQUIVO_BASE)
except FileNotFoundError:
    st.error(f"Arquivo base não encontrado: {ARQUIVO_BASE}")
    st.stop()
except Exception as e:
    st.error(f"Erro ao carregar a base TRI: {e}")
    st.stop()

anos_disponiveis = sorted([int(x) for x in tri["Ano"].dropna().unique()])

with st.sidebar:
    st.markdown("## 📚 Menu")
    menu = st.selectbox(
        "Escolha uma seção",
        [
            "Início",
            "Simulador",
            "Buscar aluno",
            "Resultado por turma",
            "Resumo da turma",
            "Alunos consolidados",
            "Desempenho consolidado da turma",
            "Histórico completo",
        ],
    )

    st.markdown("---")
    st.markdown("### 🎯 Áreas")
    exibir_badges_areas()

    st.markdown("### 📅 Anos")
    st.write(", ".join(map(str, anos_disponiveis)))

    st.markdown("---")
    st.markdown("### ⚙️ Sessão")

    resultados_sidebar = obter_resultados()
    st.write(f"Registros na sessão: **{len(resultados_sidebar)}**")

    if not resultados_sidebar.empty:
        st.download_button(
            label="📥 Baixar resultados",
            data=excel_bytes(resultados_sidebar),
            file_name="resultados_alunos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if st.button("🗑️ Apagar todos os resultados"):
        apagar_resultados()
        st.success("Todos os resultados da sessão foram apagados.")
        st.rerun()

if menu == "Início":
    exibir_cabecalho()
    resultados = obter_resultados()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Áreas", "4")
    c2.metric("Anos disponíveis", str(len(anos_disponiveis)))
    c3.metric("Registros salvos", str(len(resultados)))
    c4.metric("Modos", "Oficial + Misto")

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Bem-vinda ao painel")
    st.write(
        "Use o menu lateral para simular resultados, buscar alunos, analisar turmas, "
        "comparar desempenhos e exportar históricos."
    )
    st.info(
        "O modo oficial consulta um ano específico. O modo misto usa o histórico de vários anos "
        "para gerar uma estimativa em escala ENEM baseada apenas na quantidade de acertos."
    )
    st.markdown("</div>", unsafe_allow_html=True)

elif menu == "Simulador":
    exibir_cabecalho()

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Simular resultado")
    exibir_badges_areas()

    with st.form("form_simulador"):
        col1, col2 = st.columns(2)

        with col1:
            nome = st.text_input("Nome do aluno")
            turma = st.text_input("Turma")
            area = st.selectbox("Área", AREAS_VALIDAS)
            tipo_simulacao = st.radio("Modo de cálculo", [MODO_OFICIAL, MODO_MISTO], horizontal=True)

        with col2:
            if tipo_simulacao == MODO_OFICIAL:
                ano = st.selectbox("Ano da prova", anos_disponiveis)
            else:
                ano = None
                st.caption("No modo misto, o sistema ignora o ano e consolida os anos disponíveis.")

            acertos = st.number_input("Número de acertos", min_value=0, max_value=45, step=1)

        enviar = st.form_submit_button("Calcular e salvar")

    st.markdown("</div>", unsafe_allow_html=True)

    if enviar:
        nome = normalizar_nome(nome)
        turma = turma.strip()

        if not nome or not turma:
            st.warning("Preencha nome e turma.")
            st.stop()

        porcentagem = round((int(acertos) / 45) * 100, 1)

        if tipo_simulacao == MODO_OFICIAL:
            consulta = consultar_tri(tri, area, int(ano), int(acertos))

            if consulta.empty:
                areas_ano = sorted(
                    tri.loc[tri["Ano"] == int(ano), "Area"].dropna().unique().tolist()
                )
                st.error("Dados não encontrados para essa combinação.")
                st.info(f"Áreas disponíveis para {ano}: {', '.join(areas_ano)}")
                st.stop()

            linha = consulta.iloc[0]
            nota_min = None if pd.isna(linha["Min"]) else float(linha["Min"])
            nota_media = None if pd.isna(linha["Media"]) else float(linha["Media"])
            nota_max = None if pd.isna(linha["Max"]) else float(linha["Max"])
            observacao = f"Estimativa oficial baseada no ano {int(ano)}."

            st.markdown('<div class="section-card">', unsafe_allow_html=True)
            exibir_pill_modo(MODO_OFICIAL)
            st.subheader("Resultado estimado")

            c1, c2, c3 = st.columns(3)
            c1.metric("Nota mínima", "-" if nota_min is None else f"{nota_min:.1f}")
            c2.metric("Nota média", "-" if nota_media is None else f"{nota_media:.1f}")
            c3.metric("Nota máxima", "-" if nota_max is None else f"{nota_max:.1f}")

            st.markdown("<br>", unsafe_allow_html=True)

            c4, c5, c6 = st.columns(3)
            c4.metric("Área", area)
            c5.metric("Ano", str(ano))
            c6.metric("Acertos", f"{int(acertos)}/45")

            if porcentagem >= 60:
                st.success(f"Porcentagem de acertos: {porcentagem:.1f}%")
            else:
                st.warning(f"Porcentagem de acertos: {porcentagem:.1f}%")

            st.caption(observacao)
            st.markdown("</div>", unsafe_allow_html=True)

            novo = {
                "Nome": nome,
                "Turma": turma,
                "Area": area,
                "Tipo_Simulacao": MODO_OFICIAL,
                "Ano": int(ano),
                "Acertos": int(acertos),
                "Porcentagem_Acertos": porcentagem,
                "Nota_min": nota_min,
                "Nota_media": nota_media,
                "Nota_max": nota_max,
                "Observacao": observacao,
            }

        else:
            consulta_mista = consultar_estimativa_mista(tri, area, int(acertos))
            consolidado = consolidar_estimativa_mista(consulta_mista)

            if not consolidado:
                st.error("Não foi possível gerar estimativa mista para essa combinação.")
                st.stop()

            nota_min = consolidado["nota_min"]
            nota_media = consolidado["nota_media"]
            nota_max = consolidado["nota_max"]
            faixa_provavel_min = consolidado["faixa_provavel_min"]
            faixa_provavel_max = consolidado["faixa_provavel_max"]
            anos_base_txt = ", ".join(map(str, consolidado["anos_base"]))
            observacao = (
                "Estimativa mista baseada no histórico consolidado dos anos "
                f"{anos_base_txt}. Não representa a TRI oficial de uma edição específica."
            )

            st.markdown('<div class="section-card">', unsafe_allow_html=True)
            exibir_pill_modo(MODO_MISTO)
            st.subheader("Resultado estimado")

            c1, c2, c3 = st.columns(3)
            c1.metric("Nota mínima", "-" if nota_min is None else f"{nota_min:.1f}")
            c2.metric("Nota média", "-" if nota_media is None else f"{nota_media:.1f}")
            c3.metric("Nota máxima", "-" if nota_max is None else f"{nota_max:.1f}")

            st.markdown("<br>", unsafe_allow_html=True)

            c4, c5, c6 = st.columns(3)
            c4.metric("Área", area)
            c5.metric("Base histórica", anos_base_txt)
            c6.metric("Acertos", f"{int(acertos)}/45")

            st.markdown("<br>", unsafe_allow_html=True)
            c7, c8 = st.columns(2)
            c7.metric(
                "Faixa provável",
                "-" if faixa_provavel_min is None or faixa_provavel_max is None
                else f"{faixa_provavel_min:.1f} a {faixa_provavel_max:.1f}"
            )
            c8.metric("Mediana histórica", "-" if consolidado["mediana"] is None else f"{consolidado['mediana']:.1f}")

            if porcentagem >= 60:
                st.success(f"Porcentagem de acertos: {porcentagem:.1f}%")
            else:
                st.warning(f"Porcentagem de acertos: {porcentagem:.1f}%")

            st.caption(observacao)
            st.markdown("</div>", unsafe_allow_html=True)

            novo = {
                "Nome": nome,
                "Turma": turma,
                "Area": area,
                "Tipo_Simulacao": MODO_MISTO,
                "Ano": None,
                "Acertos": int(acertos),
                "Porcentagem_Acertos": porcentagem,
                "Nota_min": nota_min,
                "Nota_media": nota_media,
                "Nota_max": nota_max,
                "Observacao": observacao,
            }

        salvar_resultado(novo)
        st.success("Resultado salvo com sucesso.")

    resultados = obter_resultados()
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Últimos resultados salvos")

    if resultados.empty:
        st.info("Ainda não há resultados salvos.")
    else:
        ultimos = ordenar_resultados(resultados).tail(15).copy()
        ultimos["Ano"] = ultimos["Ano"].apply(ano_para_exibicao)
        st.dataframe(ultimos, use_container_width=True, hide_index=True)

    st.markdown("</div>", unsafe_allow_html=True)

elif menu == "Buscar aluno":
    exibir_cabecalho()
    resultados = obter_resultados()

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Buscar aluno")

    if resultados.empty:
        st.info("Ainda não há resultados salvos.")
    else:
        busca = st.text_input("Digite o nome ou parte do nome")

        if busca:
            filtro = resultados[
                resultados["Nome"].astype(str).str.strip().str.lower().str.contains(
                    busca.strip().lower(), na=False
                )
            ].copy()

            if filtro.empty:
                st.warning("Nenhum aluno encontrado.")
            else:
                alunos = filtro["Nome"].drop_duplicates().tolist()

                for aluno in alunos:
                    historico = filtro[filtro["Nome"] == aluno].copy()
                    historico["Ano_ordenacao"] = pd.to_numeric(historico["Ano"], errors="coerce").fillna(9999)
                    historico = historico.sort_values(by=["Ano_ordenacao", "Area", "Tipo_Simulacao"]).drop(columns=["Ano_ordenacao"])

                    melhor = historico.assign(Ano_ordenacao=pd.to_numeric(historico["Ano"], errors="coerce").fillna(9999)).sort_values(
                        by=["Nota_media", "Acertos", "Ano_ordenacao"],
                        ascending=[False, False, False],
                    ).iloc[0]

                    ultimo = historico.assign(Ano_ordenacao=pd.to_numeric(historico["Ano"], errors="coerce").fillna(9999)).sort_values(
                        by=["Ano_ordenacao", "Area"], ascending=[False, True]
                    ).iloc[0]

                    st.markdown(
                        f"""
<div class="profile-card">
    <div style="font-size:1.2rem; font-weight:700; color:#0F172A;">
        {aluno}
    </div>
    <div style="color:#475569; margin-top:4px;">
        Histórico individual consolidado
    </div>
</div>
""",
                        unsafe_allow_html=True,
                    )

                    c1, c2, c3 = st.columns(3)
                    c1.metric("Turma", historico.iloc[0]["Turma"])
                    c2.metric("Quantidade de provas", len(historico))
                    c3.metric("Melhor nota", f"{melhor['Nota_media']:.1f}")

                    st.write(
                        f"**Melhor resultado:** {melhor['Nota_media']:.1f} "
                        f"({melhor['Area']} - {ano_para_exibicao(melhor['Ano'])} - {tipo_para_exibicao(melhor['Tipo_Simulacao'])})"
                    )
                    st.write(
                        f"**Resultado mais recente:** {ultimo['Nota_media']:.1f} "
                        f"({ultimo['Area']} - {ano_para_exibicao(ultimo['Ano'])} - {tipo_para_exibicao(ultimo['Tipo_Simulacao'])})"
                    )

                    historico_exib = historico.copy()
                    historico_exib["Ano"] = historico_exib["Ano"].apply(ano_para_exibicao)
                    st.dataframe(
                        historico_exib[[
                            "Tipo_Simulacao", "Area", "Ano", "Acertos", "Porcentagem_Acertos", "Nota_media"
                        ]],
                        use_container_width=True,
                        hide_index=True,
                    )
                    st.markdown("---")

    st.markdown("</div>", unsafe_allow_html=True)

elif menu == "Resultado por turma":
    exibir_cabecalho()
    resultados = obter_resultados()

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Resultado por turma")

    if resultados.empty:
        st.info("Ainda não há resultados salvos.")
    else:
        turmas_disponiveis = sorted(resultados["Turma"].dropna().astype(str).unique().tolist())

        if not turmas_disponiveis:
            st.info("Ainda não há turmas cadastradas.")
        else:
            turma = st.selectbox("Selecione a turma", turmas_disponiveis)
            df_turma = resultados[resultados["Turma"] == turma].copy()

            if df_turma.empty:
                st.warning("Nenhum registro encontrado para essa turma.")
            else:
                ranking = obter_melhor_resultado_por_aluno(df_turma)
                ranking = ranking.sort_values(
                    by=["Nota_media", "Acertos", "Nome"],
                    ascending=[False, False, True],
                ).reset_index(drop=True)

                ranking.insert(0, "Posição", ranking.index + 1)
                ranking["Medalha"] = ""
                ranking["Ano"] = ranking["Ano"].apply(ano_para_exibicao)

                if len(ranking) >= 1:
                    ranking.loc[0, "Medalha"] = "🥇"
                if len(ranking) >= 2:
                    ranking.loc[1, "Medalha"] = "🥈"
                if len(ranking) >= 3:
                    ranking.loc[2, "Medalha"] = "🥉"

                st.dataframe(
                    ranking[
                        [
                            "Medalha", "Posição", "Nome", "Tipo_Simulacao", "Area", "Ano",
                            "Acertos", "Porcentagem_Acertos", "Nota_media",
                        ]
                    ],
                    use_container_width=True,
                    hide_index=True,
                )

                fig = px.bar(
                    ranking,
                    x="Nome",
                    y="Nota_media",
                    text="Nota_media",
                    color="Area",
                    color_discrete_map=CORES_AREA,
                    title=f"Melhor nota média por aluno - {turma}",
                )

                fig.update_traces(
                    texttemplate="%{text:.1f}",
                    textposition="outside",
                    marker_line_width=0,
                )

                fig.update_layout(
                    xaxis_title="Aluno",
                    yaxis_title="Nota média",
                    plot_bgcolor="white",
                    paper_bgcolor="white",
                    legend_title="Área",
                    title_x=0.02,
                )

                st.plotly_chart(fig, use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)

elif menu == "Resumo da turma":
    exibir_cabecalho()
    resultados = obter_resultados()

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Resumo da turma")

    if resultados.empty:
        st.info("Ainda não há resultados salvos.")
    else:
        turmas_disponiveis = sorted(resultados["Turma"].dropna().astype(str).unique().tolist())

        if not turmas_disponiveis:
            st.info("Ainda não há turmas cadastradas.")
        else:
            turma = st.selectbox("Selecione a turma", turmas_disponiveis)
            df_turma = resultados[resultados["Turma"] == turma].copy()

            if df_turma.empty:
                st.warning("Nenhum registro encontrado para essa turma.")
            else:
                melhor_por_aluno = obter_melhor_resultado_por_aluno(df_turma)
                alunos_unicos = df_turma["Nome"].nunique()

                st.markdown(
                    """
<div class="info-box">
    <strong>Visão geral da turma</strong><br>
    <span style="color:#475569;">
        Indicadores calculados com base no melhor resultado de cada aluno.
    </span>
</div>
""",
                    unsafe_allow_html=True,
                )

                c1, c2, c3 = st.columns(3)
                c1.metric("Quantidade de registros", len(df_turma))
                c2.metric("Quantidade de alunos", alunos_unicos)
                c3.metric("Média TRI estimada", f"{melhor_por_aluno['Nota_media'].mean():.2f}")

                c4, c5, c6 = st.columns(3)
                c4.metric("Média de acertos", f"{melhor_por_aluno['Acertos'].mean():.2f}")
                c5.metric("Média de porcentagem", f"{melhor_por_aluno['Porcentagem_Acertos'].mean():.2f}%")
                c6.metric("Maior nota média", f"{melhor_por_aluno['Nota_media'].max():.2f}")

                st.markdown("<br>", unsafe_allow_html=True)
                st.subheader("Resumo consolidado dos alunos")

                resumo = obter_resumo_alunos(df_turma).sort_values(
                    by=["Melhor_Nota", "Media_Notas", "Nome"],
                    ascending=[False, False, True],
                ).reset_index(drop=True)

                st.dataframe(resumo, use_container_width=True, hide_index=True)

    st.markdown("</div>", unsafe_allow_html=True)

elif menu == "Alunos consolidados":
    exibir_cabecalho()
    resultados = obter_resultados()

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Alunos consolidados")

    if resultados.empty:
        st.info("Ainda não há resultados salvos.")
    else:
        resumo = obter_resumo_alunos(resultados)
        st.dataframe(resumo, use_container_width=True, hide_index=True)

    st.markdown("</div>", unsafe_allow_html=True)

elif menu == "Desempenho consolidado da turma":
    exibir_cabecalho()
    resultados = obter_resultados()

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Desempenho consolidado da turma")

    if resultados.empty:
        st.info("Ainda não há resultados salvos.")
    else:
        turmas_disponiveis = sorted(resultados["Turma"].dropna().astype(str).unique().tolist())

        if not turmas_disponiveis:
            st.info("Ainda não há turmas cadastradas.")
        else:
            turma = st.selectbox("Selecione a turma", turmas_disponiveis)
            df_turma = resultados[resultados["Turma"] == turma].copy()

            if df_turma.empty:
                st.warning("Nenhum registro encontrado para essa turma.")
            else:
                resumo = obter_resumo_alunos(df_turma)

                resumo = resumo.sort_values(
                    by=["Melhor_Nota", "Media_Notas", "Nome"],
                    ascending=[False, False, True],
                ).reset_index(drop=True)

                resumo.insert(0, "Posição", resumo.index + 1)
                st.dataframe(resumo, use_container_width=True, hide_index=True)

    st.markdown("</div>", unsafe_allow_html=True)

elif menu == "Histórico completo":
    exibir_cabecalho()
    resultados = obter_resultados()

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Histórico completo")

    if resultados.empty:
        st.info("Ainda não há resultados salvos.")
    else:
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            filtro_turma = st.selectbox(
                "Filtrar por turma",
                ["Todas"] + sorted(resultados["Turma"].dropna().astype(str).unique().tolist()),
            )

        with col2:
            filtro_area = st.selectbox(
                "Filtrar por área",
                ["Todas"] + AREAS_VALIDAS,
            )

        with col3:
            filtro_modo = st.selectbox(
                "Filtrar por modo",
                ["Todos", MODO_OFICIAL, MODO_MISTO],
            )

        with col4:
            anos_hist = sorted(resultados["Ano"].dropna().astype(int).unique().tolist())
            filtro_ano = st.selectbox(
                "Filtrar por ano",
                ["Todos", "Misto"] + anos_hist,
            )

        df = resultados.copy()

        if filtro_turma != "Todas":
            df = df[df["Turma"] == filtro_turma]

        if filtro_area != "Todas":
            df = df[df["Area"] == filtro_area]

        if filtro_modo != "Todos":
            df = df[df["Tipo_Simulacao"] == filtro_modo]

        if filtro_ano == "Misto":
            df = df[df["Ano"].isna()]
        elif filtro_ano != "Todos":
            df = df[df["Ano"] == filtro_ano]

        df = ordenar_resultados(df)
        df_exib = df.copy()
        df_exib["Ano"] = df_exib["Ano"].apply(ano_para_exibicao)
        st.dataframe(df_exib, use_container_width=True, hide_index=True)

        st.download_button(
            label="📥 Baixar histórico filtrado",
            data=excel_bytes(df),
            file_name="historico_filtrado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader("Média por turma")

        resumo_turmas = gerar_resumo_por_turma(resultados)
        st.dataframe(resumo_turmas, use_container_width=True, hide_index=True)

    st.markdown("</div>", unsafe_allow_html=True)
