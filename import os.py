import os
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import date

# --------------------- Configuração básica ---------------------
st.set_page_config(page_title="Dashboard de RH", layout="wide")
st.title("📊 Dashboard de RH")

DEFAULT_EXCEL_PATH = "BaseFuncionarios.xlsx"
DATE_COLS = ["Data de Nascimento", "Data de Contratacao", "Data de Demissao"]

# --------------------- Funções utilitárias ---------------------
def brl(x: float) -> str:
    try:
        return f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "R$ 0,00"


def prepare_df(df: pd.DataFrame) -> pd.DataFrame:
    for c in df.select_dtypes(include="object").columns:
        df[c] = df[c].astype(str).str.strip()

    for c in DATE_COLS:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], dayfirst=True, errors="coerce")

    if "Sexo" in df.columns:
        df["Sexo"] = df["Sexo"].str.upper().replace({"MASCULINO": "M", "FEMININO": "F"})

    for col in ["Salario Base", "Impostos", "Beneficios", "VT", "VR"]:
        if col not in df.columns:
            df[col] = 0.0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    today = pd.Timestamp(date.today())

    if "Data de Nascimento" in df.columns:
        df["Idade"] = ((today - df["Data de Nascimento"]).dt.days // 365).clip(lower=0)

    if "Data de Contratacao" in df.columns:
        meses = (today.year - df["Data de Contratacao"].dt.year) * 12 + \
                (today.month - df["Data de Contratacao"].dt.month)
        df["Tempo de Casa (meses)"] = meses.clip(lower=0)

    if "Data de Demissao" in df.columns:
        df["Status"] = np.where(df["Data de Demissao"].notna(), "Desligado", "Ativo")
    else:
        df["Status"] = "Ativo"

    df["Custo Total Mensal"] = df[["Salario Base", "Impostos", "Beneficios", "VT", "VR"]].sum(axis=1)
    return df


@st.cache_data
def load_from_path(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=0, engine="openpyxl")
    return prepare_df(df)


@st.cache_data
def load_from_bytes(uploaded_bytes) -> pd.DataFrame:
    df = pd.read_excel(uploaded_bytes, sheet_name=0, engine="openpyxl")
    return prepare_df(df)

# --------------------- Sidebar: fonte de dados ---------------------
with st.sidebar:
    st.header("📂 Fonte de dados")
    up = st.file_uploader("Carregar Excel (.xlsx)", type=["xlsx"])
    caminho_manual = st.text_input("Ou caminho do Excel", value=DEFAULT_EXCEL_PATH)
    st.divider()
    if up is None:
        existe = os.path.exists(caminho_manual)
        st.write(f"Arquivo: **{caminho_manual}** — { '✅ Encontrado' if existe else '❌ Não existe'}")

# --------------------- Carregamento ---------------------
df, fonte = None, None
if up is not None:
    try:
        df = load_from_bytes(up)
        fonte = "Upload"
    except Exception as e:
        st.error(f"Erro ao ler Excel (Upload): {e}")
        st.stop()
else:
    try:
        if not os.path.exists(caminho_manual):
            st.error(f"Arquivo não encontrado em: {caminho_manual}")
            st.stop()
        df = load_from_path(caminho_manual)
        fonte = "Caminho"
    except Exception as e:
        st.error(f"Erro ao ler Excel (Caminho): {e}")
        st.stop()

st.caption(f"Dados carregados via **{fonte}** | Linhas: {len(df)} | Colunas: {len(df.columns)}")
with st.expander("Ver colunas detectadas"):
    st.write(list(df.columns))

# --------------------- Filtros ---------------------
st.sidebar.header("🔎 Filtros")

def msel(col):
    if col in df.columns:
        vals = sorted(df[col].dropna().unique())
        return st.sidebar.multiselect(col, vals)
    return []

area_sel   = msel("Área")
nivel_sel  = msel("Nível")
cargo_sel  = msel("Cargo")
sexo_sel   = msel("Sexo")
status_sel = msel("Status")
nome_busca = st.sidebar.text_input("Buscar por Nome Completo")

# Faixas
if "Idade" in df.columns and not df["Idade"].dropna().empty:
    ida_min, ida_max = int(df["Idade"].min()), int(df["Idade"].max())
    faixa_idade = st.sidebar.slider("Faixa Etária", ida_min, ida_max, (ida_min, ida_max))
else:
    faixa_idade = None

if "Salario Base" in df.columns and not df["Salario Base"].dropna().empty:
    sal_min, sal_max = float(df["Salario Base"].min()), float(df["Salario Base"].max())
    faixa_sal = st.sidebar.slider("Faixa Salarial", sal_min, sal_max, (sal_min, sal_max))
else:
    faixa_sal = None

# Aplica filtros de forma otimizada
df_f = df.copy()
filters = {
    "Área": area_sel,
    "Nível": nivel_sel,
    "Cargo": cargo_sel,
    "Sexo": sexo_sel,
    "Status": status_sel
}

for col, vals in filters.items():
    if vals:
        df_f = df_f[df_f[col].isin(vals)]

if nome_busca and "Nome Completo" in df_f.columns:
    df_f = df_f[df_f["Nome Completo"].str.contains(nome_busca, case=False, na=False)]

if faixa_idade:
    df_f = df_f[(df_f["Idade"] >= faixa_idade[0]) & (df_f["Idade"] <= faixa_idade[1])]

if faixa_sal:
    df_f = df_f[(df_f["Salario Base"] >= faixa_sal[0]) & (df_f["Salario Base"] <= faixa_sal[1])]

# --------------------- KPIs ---------------------
def k_headcount_ativo(d): return int((d["Status"] == "Ativo").sum())
def k_desligados(d): return int((d["Status"] == "Desligado").sum())
def k_idade_media(d): return float(d["Idade"].mean()) if "Idade" in d else 0.0
def k_tempo_casa(d): return float(d["Tempo de Casa (meses)"].mean()) if "Tempo de Casa (meses)" in d else 0.0
def k_turnover(d):
    ativos = (d["Status"] == "Ativo").sum()
    desligados = (d["Status"] == "Desligado").sum()
    total = len(d)
    return (desligados / total * 100) if total > 0 else 0

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Headcount Ativo", k_headcount_ativo(df_f))
c2.metric("Desligados", k_desligados(df_f))
c3.metric("Idade Média", f"{k_idade_media(df_f):.1f} anos")
c4.metric("Tempo Médio de Casa", f"{k_tempo_casa(df_f):.1f} meses")
c5.metric("Turnover", f"{k_turnover(df_f):.1f}%")

st.divider()

# --------------------- Gráficos ---------------------
colA, colB = st.columns(2)
with colA:
    if "Área" in df_f.columns:
        d = df_f.groupby("Área").size().reset_index(name="Headcount")
        fig = px.bar(d, x="Área", y="Headcount", title="Headcount por Área")
        st.plotly_chart(fig, use_container_width=True)

with colB:
    if "Cargo" in df_f.columns:
        d = df_f.groupby("Cargo", as_index=False)["Salario Base"].mean().sort_values("Salario Base", ascending=False)
        fig = px.bar(d, x="Cargo", y="Salario Base", title="Salário Médio por Cargo")
        st.plotly_chart(fig, use_container_width=True)

colC, colD = st.columns(2)
with colC:
    if "Idade" in df_f.columns:
        fig = px.histogram(df_f, x="Idade", nbins=20, title="Distribuição de Idades")
        st.plotly_chart(fig, use_container_width=True)

with colD:
    if "Sexo" in df_f.columns:
        d = df_f["Sexo"].value_counts().reset_index()
        d.columns = ["Sexo", "Contagem"]
        fig = px.pie(d, values="Contagem", names="Sexo", title="Distribuição por Sexo")
        st.plotly_chart(fig, use_container_width=True)

# Boxplot de salários
if "Salario Base" in df_f.columns and not df_f.empty:
    st.subheader("📌 Distribuição de Salários")
    fig = px.box(df_f, y="Salario Base", points="all", title="Boxplot de Salário Base")
    st.plotly_chart(fig, use_container_width=True)

st.divider()

# --------------------- Tabela e Downloads ---------------------
st.subheader("📋 Tabela (dados filtrados)")
st.dataframe(df_f, use_container_width=True)

csv_bytes = df_f.to_csv(index=False).encode("utf-8")
st.download_button("Baixar CSV filtrado", data=csv_bytes, file_name="funcionarios_filtrado.csv", mime="text/csv")

if st.toggle("Gerar Excel filtrado"):
    from io import BytesIO
    buff = BytesIO()
    with pd.ExcelWriter(buff, engine="openpyxl") as writer:
        df_f.to_excel(writer, index=False, sheet_name="Filtrado")
    st.download_button("Baixar Excel filtrado", data=buff.getvalue(), file_name="funcionarios_filtrado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
