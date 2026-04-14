import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from pathlib import Path
import plotly.express as px
from io import BytesIO

ARQUIVO_BASE = "base_TRI_ENEM_FINAL_2021_2024_corrigida.xlsx"
ARQUIVO_SAIDA = "resultados_alunos.xlsx"
AREAS_VALIDAS = ["CN", "CH", "LC", "MT"]

CORES_AREA = {
    "CN": "#22C55E",
    "CH": "#3B82F6",
    "LC": "#EC4899",
    "MT": "#F59E0B",
}


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

.main-title {
    font-size: 2.8rem;
    font-weight: 800;
    color: #0F172A;
    margin-bottom: 0.2rem;
}

.subtitle {
    font-size: 1.05rem;
    color: #334155;
    margin-bottom: 1.5rem;
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

/* =========================
   INPUTS MAIS VISÍVEIS
========================= */

/* INPUTS STREAMLIT (mais específico) */
div[data-baseweb="input"] input {
    background-color: #FFFFFF !important;
    border: 2px solid #64748B !important;
    border-radius: 10px !important;
    color: #0F172A !important;
    font-weight: 500;
}

/* SELECT */
div[data-baseweb="select"] > div {
    background-color: #FFFFFF !important;
    border: 2px solid #64748B !important;
    border-radius: 10px !important;
}

/* FOCUS */
div[data-baseweb="input"] input:focus {
    border: 2px solid #7C3AED !important;
    box-shadow: 0 0 0 2px rgba(124, 58, 237, 0.25);
}

/* Labels */
label {
    font-weight: 600 !important;
    color: #1E293B !important;
}
</style>
""",
    unsafe_allow_html=True,
)


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

    return tri


def carregar_resultados(caminho_arquivo: str) -> pd.DataFrame:
    colunas = [
        "Nome", "Turma", "Area", "Ano", "Acertos",
        "Porcentagem_Acertos", "Nota_min", "Nota_media", "Nota_max",
    ]

    if Path(caminho_arquivo).exists():
        df = pd.read_excel(caminho_arquivo)

        for coluna in colunas:
            if coluna not in df.columns:
                df[coluna] = None

        df = df[colunas].copy()
        df["Nome"] = df["Nome"].astype(str).str.strip()
        df["Turma"] = df["Turma"].astype(str).str.strip()
        df["Area"] = df["Area"].astype(str).str.strip().str.upper()

        return df

    return pd.DataFrame(columns=colunas)


def salvar_resultado(caminho_arquivo: str, novo: dict) -> None:
    resultados = carregar_resultados(caminho_arquivo)

    if resultados.empty:
        resultados = pd.DataFrame([novo])
    else:
        resultados = pd.concat([resultados, pd.DataFrame([novo])], ignore_index=True)

    resultados.to_excel(caminho_arquivo, index=False)


def consultar_tri(tri: pd.DataFrame, area: str, ano: int, acertos: int) -> pd.DataFrame:
    return tri[
        (tri["Area"] == area)
        & (tri["Ano"] == ano)
        & (tri["Acertos"] == acertos)
    ]


def normalizar_nome(nome: str) -> str:
    return nome.strip().title()


def obter_melhor_resultado_por_aluno(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df.copy()

    temp = df.copy()
    temp["Nome"] = temp["Nome"].astype(str).str.strip()
    temp["Turma"] = temp["Turma"].astype(str).str.strip()
    temp["Area"] = temp["Area"].astype(str).str.strip().str.upper()

    melhor = (
        temp.sort_values(
            by=["Nome", "Nota_media", "Acertos", "Ano"],
            ascending=[True, False, False, False],
        )
        .groupby("Nome", as_index=False)
        .first()
    )

    return melhor


def obter_resumo_alunos(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(
            columns=[
                "Nome", "Turma", "Qtd_Provas", "Melhor_Area", "Melhor_Ano",
                "Melhor_Nota", "Media_Notas", "Media_Acertos",
                "Ultimo_Ano", "Ultima_Area", "Ultima_Nota",
            ]
        )

    temp = df.copy()
    temp["Nome"] = temp["Nome"].astype(str).str.strip()
    temp["Turma"] = temp["Turma"].astype(str).str.strip()
    temp["Area"] = temp["Area"].astype(str).str.strip().str.upper()

    linhas = []

    for nome, grupo in temp.groupby("Nome"):
        grupo = grupo.sort_values(by=["Ano", "Area"])

        melhor = grupo.sort_values(
            by=["Nota_media", "Acertos", "Ano"],
            ascending=[False, False, False],
        ).iloc[0]

        ultimo = grupo.sort_values(by=["Ano", "Area"], ascending=[False, True]).iloc[0]

        linhas.append(
            {
                "Nome": nome,
                "Turma": grupo.iloc[0]["Turma"],
                "Qtd_Provas": len(grupo),
                "Melhor_Area": melhor["Area"],
                "Melhor_Ano": int(melhor["Ano"]),
                "Melhor_Nota": round(float(melhor["Nota_media"]), 2),
                "Media_Notas": round(float(grupo["Nota_media"].mean()), 2),
                "Media_Acertos": round(float(grupo["Acertos"].mean()), 2),
                "Ultimo_Ano": int(ultimo["Ano"]),
                "Ultima_Area": ultimo["Area"],
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
    components.html(
        """
<div style="
    background: linear-gradient(135deg, #D1D5DB, #A78BFA, #7C3AED);
    padding: 30px;
    border-radius: 22px;
    color: white;
    margin-bottom: 24px;
    box-shadow: 0 12px 28px rgba(124, 58, 237, 0.25);
    font-family: 'Segoe UI', sans-serif;
">
    <div style="
        font-size: 2.4rem;
        font-weight: 800;
        margin-bottom: 0.4rem;
    ">
        📘 Simulador TRI ENEM
    </div>

    <div style="
        font-size: 1.05rem;
        opacity: 0.96;
    ">
        Acompanhe resultados, evolução dos alunos e desempenho por turma.
    </div>
</div>
""",
        height=170,
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


if menu == "Início":
    exibir_cabecalho()
    resultados = carregar_resultados(ARQUIVO_SAIDA)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Áreas", "4")
    c2.metric("Anos disponíveis", str(len(anos_disponiveis)))
    c3.metric("Registros salvos", str(len(resultados)))
    c4.metric("Modo", "TRI ENEM")

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Bem-vinda ao painel")
    st.write(
        "Use o menu lateral para simular resultados, buscar alunos, analisar turmas, "
        "comparar desempenhos e exportar históricos."
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

        with col2:
            ano = st.selectbox("Ano da prova", anos_disponiveis)
            acertos = st.number_input("Número de acertos", min_value=0, max_value=45, step=1)

        enviar = st.form_submit_button("Calcular e salvar")

    st.markdown("</div>", unsafe_allow_html=True)

    if enviar:
        nome = normalizar_nome(nome)
        turma = turma.strip()

        if not nome or not turma:
            st.warning("Preencha nome e turma.")
            st.stop()

        consulta = consultar_tri(tri, area, int(ano), int(acertos))

        if consulta.empty:
            areas_ano = sorted(
                tri.loc[tri["Ano"] == int(ano), "Area"].dropna().unique().tolist()
            )
            st.error("Dados não encontrados para essa combinação.")
            st.info(f"Áreas disponíveis para {ano}: {', '.join(areas_ano)}")
            st.stop()

        linha = consulta.iloc[0]
        porcentagem = round((int(acertos) / 45) * 100, 1)

        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Resultado estimado")

        c1, c2, c3 = st.columns(3)
        c1.metric("Nota mínima", "-" if pd.isna(linha["Min"]) else f"{linha['Min']:.1f}")
        c2.metric("Nota média", "-" if pd.isna(linha["Media"]) else f"{linha['Media']:.1f}")
        c3.metric("Nota máxima", "-" if pd.isna(linha["Max"]) else f"{linha['Max']:.1f}")

        st.markdown("<br>", unsafe_allow_html=True)

        c4, c5, c6 = st.columns(3)
        c4.metric("Área", area)
        c5.metric("Ano", str(ano))
        c6.metric("Acertos", f"{int(acertos)}/45")

        if porcentagem >= 60:
            st.success(f"Porcentagem de acertos: {porcentagem:.1f}%")
        else:
            st.warning(f"Porcentagem de acertos: {porcentagem:.1f}%")

        st.markdown("</div>", unsafe_allow_html=True)

        novo = {
            "Nome": nome,
            "Turma": turma,
            "Area": area,
            "Ano": int(ano),
            "Acertos": int(acertos),
            "Porcentagem_Acertos": porcentagem,
            "Nota_min": None if pd.isna(linha["Min"]) else float(linha["Min"]),
            "Nota_media": None if pd.isna(linha["Media"]) else float(linha["Media"]),
            "Nota_max": None if pd.isna(linha["Max"]) else float(linha["Max"]),
        }

        try:
            salvar_resultado(ARQUIVO_SAIDA, novo)
            st.success("Resultado salvo com sucesso.")
        except Exception as e:
            st.error(f"Erro ao salvar o resultado: {e}")

    resultados = carregar_resultados(ARQUIVO_SAIDA)
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Últimos resultados salvos")

    if resultados.empty:
        st.info("Ainda não há resultados salvos.")
    else:
        ultimos = resultados.sort_values(by=["Turma", "Nome", "Ano", "Area"]).tail(15)
        st.dataframe(ultimos, use_container_width=True, hide_index=True)

    st.markdown("</div>", unsafe_allow_html=True)

elif menu == "Buscar aluno":
    exibir_cabecalho()
    resultados = carregar_resultados(ARQUIVO_SAIDA)

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
                    historico = historico.sort_values(by=["Ano", "Area"])

                    melhor = historico.sort_values(
                        by=["Nota_media", "Acertos", "Ano"],
                        ascending=[False, False, False],
                    ).iloc[0]

                    ultimo = historico.sort_values(
                        by=["Ano", "Area"], ascending=[False, True]
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
                        f"({melhor['Area']} - {int(melhor['Ano'])})"
                    )
                    st.write(
                        f"**Resultado mais recente:** {ultimo['Nota_media']:.1f} "
                        f"({ultimo['Area']} - {int(ultimo['Ano'])})"
                    )

                    st.dataframe(
                        historico[["Area", "Ano", "Acertos", "Porcentagem_Acertos", "Nota_media"]],
                        use_container_width=True,
                        hide_index=True,
                    )
                    st.markdown("---")

    st.markdown("</div>", unsafe_allow_html=True)

elif menu == "Resultado por turma":
    exibir_cabecalho()
    resultados = carregar_resultados(ARQUIVO_SAIDA)

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

                if len(ranking) >= 1:
                    ranking.loc[0, "Medalha"] = "🥇"
                if len(ranking) >= 2:
                    ranking.loc[1, "Medalha"] = "🥈"
                if len(ranking) >= 3:
                    ranking.loc[2, "Medalha"] = "🥉"

                st.dataframe(
                    ranking[
                        [
                            "Medalha", "Posição", "Nome", "Area", "Ano",
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
    resultados = carregar_resultados(ARQUIVO_SAIDA)

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
    resultados = carregar_resultados(ARQUIVO_SAIDA)

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
    resultados = carregar_resultados(ARQUIVO_SAIDA)

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
    resultados = carregar_resultados(ARQUIVO_SAIDA)

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Histórico completo")

    if resultados.empty:
        st.info("Ainda não há resultados salvos.")
    else:
        col1, col2, col3 = st.columns(3)

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
            anos_hist = sorted(resultados["Ano"].dropna().astype(int).unique().tolist())
            filtro_ano = st.selectbox(
                "Filtrar por ano",
                ["Todos"] + anos_hist,
            )

        df = resultados.copy()

        if filtro_turma != "Todas":
            df = df[df["Turma"] == filtro_turma]

        if filtro_area != "Todas":
            df = df[df["Area"] == filtro_area]

        if filtro_ano != "Todos":
            df = df[df["Ano"] == filtro_ano]

        df = df.sort_values(by=["Turma", "Nome", "Ano", "Area"])
        st.dataframe(df, use_container_width=True, hide_index=True)

        buffer = BytesIO()
        df.to_excel(buffer, index=False, engine="openpyxl")
        buffer.seek(0)

        st.download_button(
            label="Baixar histórico em Excel",
            data=buffer,
            file_name="historico_filtrado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader("Média por turma")

        resumo_turmas = gerar_resumo_por_turma(resultados)
        st.dataframe(resumo_turmas, use_container_width=True, hide_index=True)

    st.markdown("</div>", unsafe_allow_html=True)