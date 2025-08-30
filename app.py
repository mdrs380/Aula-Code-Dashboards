# Como rodar:
# 0) Crie um ambiente virtual  ->  python -m venv venv
# 1) Ative a venv  ->  .venv\Scripts\Activate.ps1   (Windows)  |  source .venv/bin/activate  (Mac/Linux)
# 2) Instale deps  ->  pip install -r requirements.txt
# 3) Rode          ->  streamlit run app.py

import os
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import date

# --------------------- Configuração básica ---------------------
st.set_page_config(
    page_title="Dashboard de RH", 
    layout="wide", 
    initial_sidebar_state="expanded",
    page_icon="📊"
)

# --------------------- Estilo e Layout ---------------------
st.markdown("## ✨ Dashboard de Recursos Humanos")
st.markdown("Bem-vindo ao painel analítico da sua equipe.")
st.markdown("---")

# Paleta de cores para os gráficos
COLOR_PALETTE = {
    "M": "#2c6fbb",  # Azul vibrante para homens
    "F": "#ff6f61",  # Rosa moderno para mulheres
    "main_color": "#2c6fbb",
    "secondary_color": "#4a90e2",
    "accent_color": "#f8c402"
}

# Se o arquivo estiver na mesma pasta do app.py, pode deixar assim.
# Ajuste para o caminho local caso esteja em outra pasta (ex.: r"C:\...\BaseFuncionarios.xlsx")
DEFAULT_EXCEL_PATH = "BaseFuncionarios.xlsx"
DATE_COLS = ["Data de Nascimento", "Data de Contratacao", "Data de Demissao"]

# --------------------- Funções utilitárias e KPIs ---------------------
def brl(x: float) -> str:
    """Formata um número como moeda brasileira (R$)."""
    try:
        return f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "R$ 0,00"

def prepare_df(df: pd.DataFrame) -> pd.DataFrame:
    """Prepara o dataframe, padronizando dados e criando colunas derivadas."""
    # Padroniza textos
    for c in df.select_dtypes(include="object").columns:
        df[c] = df[c].astype(str).str.strip()

    # Datas
    for c in DATE_COLS:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], dayfirst=True, errors="coerce")

    # Padroniza Sexo
    if "Sexo" in df.columns:
        df["Sexo"] = (
            df["Sexo"].astype(str).str.upper()
            .replace({"MASCULINO": "M", "FEMININO": "F", "F": "F", "M": "M"})
        )

    # Garante numéricos
    for col in ["Salario Base", "Impostos", "Beneficios", "VT", "VR"]:
        if col not in df.columns:
            df[col] = 0.0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    # Colunas derivadas
    today = pd.Timestamp(date.today())

    if "Data de Nascimento" in df.columns:
        df["Idade"] = ((today - df["Data de Nascimento"]).dt.days // 365).clip(lower=0)

    if "Data de Contratacao" in df.columns:
        meses = (today.year - df["Data de Contratacao"].dt.year) * 12 + \
                (today.month - df["Data de Contratacao"].month)
        df["Tempo de Casa (meses)"] = meses.clip(lower=0)

    if "Data de Demissao" in df.columns:
        df["Status"] = np.where(df["Data de Demissao"].notna(), "Desligado", "Ativo")
    else:
        df["Status"] = "Ativo"

    df["Custo Total Mensal"] = df[["Salario Base", "Impostos", "Beneficios", "VT", "VR"]].sum(axis=1)
    return df

@st.cache_data
def load_from_path(path: str) -> pd.DataFrame:
    """Carrega o dataframe de um caminho de arquivo local."""
    df = pd.read_excel(path, sheet_name=0, engine="openpyxl")
    return prepare_df(df)

@st.cache_data
def load_from_bytes(uploaded_bytes) -> pd.DataFrame:
    """Carrega o dataframe de um arquivo enviado pelo usuário."""
    df = pd.read_excel(uploaded_bytes, sheet_name=0, engine="openpyxl")
    return prepare_df(df)

def k_headcount_ativo(d):  
    return int((d["Status"] == "Ativo").sum()) if "Status" in d.columns else 0

def k_desligados(d):  
    return int((d["Status"] == "Desligado").sum()) if "Status" in d.columns else 0

def k_folha(d):
    return float(d.loc[d["Status"] == "Ativo", "Salario Base"].sum()) \
        if ("Status" in d.columns and "Salario Base" in d.columns) else 0.0

def k_custo_total(d):
    return float(d.loc[d["Status"] == "Ativo", "Custo Total Mensal"].sum()) \
        if ("Status" in d.columns and "Custo Total Mensal" in d.columns) else 0.0

def k_idade_media(d):
    return float(d["Idade"].mean()) if "Idade" in d.columns and len(d) > 0 else 0.0

def k_tempo_casa_medio(d):
    col = "Tempo de Casa (meses)"
    return float(d[col].mean()) if col in d.columns and len(d) > 0 else 0.0

def k_avaliacao_media(d):
    col = "Avaliação do Funcionário"
    return float(d[col].mean()) if col in d.columns and len(d) > 0 else 0.0

def k_avaliacao_menor_que_7(d):
    """Conta o número de funcionários com avaliação menor que 7."""
    col = "Avaliação do Funcionário"
    if col in d.columns and not d[col].isna().all():
        return int((d[col] < 7.0).sum())
    return 0

def k_aposentadoria_proxima(d):
    """Conta funcionários que se aposentam em 1 ano ou menos (idade de 60 anos)."""
    col_nascimento = "Data de Nascimento"
    if col_nascimento in d.columns:
        hoje = pd.Timestamp(date.today())
        idade_futura = hoje.year - d[col_nascimento].dt.year
        return int((idade_futura >= 59).sum())
    return 0

# --------------------- Sidebar: fonte de dados ---------------------
with st.sidebar:
    st.header("Fonte de dados")
    st.caption("Use **Upload** ou informe o caminho do arquivo .xlsx")
    up = st.file_uploader("Carregar Excel (.xlsx)", type=["xlsx"])
    caminho_manual = st.text_input("Ou caminho do Excel", value=DEFAULT_EXCEL_PATH)
    st.divider()
    if up is None:
        existe = os.path.exists(caminho_manual)
        st.write(f"Arquivo em caminho: **{caminho_manual}**")
        st.write("Existe: ", "✅ Sim" if existe else "❌ Não")

# --------------------- Carregamento com erros visíveis ---------------------
df = None
fonte = None
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
            st.info("Dica: coloque o .xlsx na mesma pasta do app.py ou ajuste o caminho acima.")
            st.stop()
        df = load_from_path(caminho_manual)
        fonte = "Caminho"
    except Exception as e:
        st.error(f"Erro ao ler Excel (Caminho): {e}")
        st.stop()

st.caption(f"Dados carregados via **{fonte}**. Linhas: {len(df)} | Colunas: {len(df.columns)}")

# Mostra colunas detectadas (ajuda no debug)
with st.expander("Ver colunas detectadas"):
    st.write(list(df.columns))
st.markdown("---")

# --------------------- Filtros ---------------------
st.sidebar.header("Filtros")

def msel(col_name: str):
    """Cria um seletor múltiplo na barra lateral para uma coluna."""
    if col_name in df.columns:
        vals = sorted([v for v in df[col_name].dropna().unique()])
        return st.sidebar.multiselect(col_name, vals)
    return []

area_sel   = msel("Área")
nivel_sel  = msel("Nível")
cargo_sel  = msel("Cargo")
sexo_sel   = msel("Sexo")
status_sel = msel("Status")
nome_busca = st.sidebar.text_input("Buscar por Nome Completo")

# Períodos
def date_bounds(series: pd.Series):
    """Retorna os limites de data de uma série."""
    s = series.dropna()
    if s.empty:
        return None
    return (s.min().date(), s.max().date())

contr_bounds = date_bounds(df["Data de Contratacao"]) if "Data de Contratacao" in df.columns else None
demis_bounds = date_bounds(df["Data de Demissao"]) if "Data de Demissao" in df.columns else None

if contr_bounds:
    d1, d2 = st.sidebar.date_input("Período de Contratação", value=contr_bounds)
else:
    d1, d2 = None, None

if demis_bounds:
    d3, d4 = st.sidebar.date_input("Período de Demissão", value=demis_bounds)
else:
    d3, d4 = None, None

# Sliders (idade e salário)
if "Idade" in df.columns and not df["Idade"].dropna().empty:
    ida_min, ida_max = int(df["Idade"].min()), int(df["Idade"].max())
    faixa_idade = st.sidebar.slider("Faixa Etária", ida_min, ida_max, (ida_min, ida_max))
else:
    faixa_idade = None

if "Salario Base" in df.columns and not df["Salario Base"].dropna().empty:
    sal_min, sal_max = float(df["Salario Base"].min()), float(df["Salario Base"].max())
    faixa_sal = st.sidebar.slider("Faixa de Salário Base", float(sal_min), float(sal_max), (float(sal_min), float(sal_max)))
else:
    faixa_sal = None

# Novo filtro: Avaliação do Funcionário
if "Avaliação do Funcionário" in df.columns:
    avaliacao_max_sel = st.sidebar.number_input(
        "Avaliação Máxima do Funcionário",
        min_value=0.0,
        max_value=10.0,
        value=10.0,
        step=0.1,
        format="%.1f"
    )
else:
    avaliacao_max_sel = 10.0


# Aplica filtros
df_f = df.copy()

def apply_in(df_, col, values):
    """Filtra um dataframe com base em uma lista de valores."""
    if values and col in df_.columns:
        return df_[df_[col].isin(values)]
    return df_

df_f = apply_in(df_f, "Área", area_sel)
df_f = apply_in(df_f, "Nível", nivel_sel)
df_f = apply_in(df_f, "Cargo", cargo_sel)
df_f = apply_in(df_f, "Sexo", sexo_sel)
df_f = apply_in(df_f, "Status", status_sel)

if nome_busca and "Nome Completo" in df_f.columns:
    df_f = df_f[df_f["Nome Completo"].str.contains(nome_busca, case=False, na=False)]

if faixa_idade and "Idade" in df_f.columns:
    df_f = df_f[(df_f["Idade"] >= faixa_idade[0]) & (df_f["Idade"] <= faixa_idade[1])]

if faixa_sal and "Salario Base" in df_f.columns:
    df_f = df_f[(df_f["Salario Base"] >= faixa_sal[0]) & (df_f["Salario Base"] <= faixa_sal[1])]

if "Avaliação do Funcionário" in df_f.columns:
    df_f = df_f[df_f["Avaliação do Funcionário"] <= avaliacao_max_sel]

if d1 and d2 and "Data de Contratacao" in df_f.columns:
    df_f = df_f[(df_f["Data de Contratacao"].isna()) |
                ((df_f["Data de Contratacao"] >= pd.to_datetime(d1)) &
                 (df_f["Data de Contratacao"] <= pd.to_datetime(d2)))]

if d3 and d4 and "Data de Demissao" in df_f.columns:
    df_f = df_f[(df_f["Data de Demissao"].isna()) |
                ((df_f["Data de Demissao"] >= pd.to_datetime(d3)) &
                 (df_f["Data de Demissao"] <= pd.to_datetime(d4)))]


# --------------------- Visualização principal ---------------------
tab_resumo, tab_graficos, tab_tabela = st.tabs(["Resumo", "Gráficos", "Tabela de Funcionários"])

with tab_resumo:
    st.subheader("Métricas Chave de RH")
    col1, col2, col3 = st.columns(3)
    col1.metric("Headcount Ativo", k_headcount_ativo(df_f))
    col2.metric("Desligados", k_desligados(df_f))
    col3.metric("Folha Salarial", brl(k_folha(df_f)))

    col4, col5, col6 = st.columns(3)
    col4.metric("Custo Total", brl(k_custo_total(df_f)))
    col5.metric("Idade Média", f"{k_idade_media(df_f):.1f} anos")
    col6.metric("Avaliação Média", f"{k_avaliacao_media(df_f):.2f}")
    
    st.metric("Próximo da Aposentadoria", k_aposentadoria_proxima(df_f))
    
    st.divider()
    
    count_abaixo_7 = k_avaliacao_menor_que_7(df_f)
    with st.expander(f"Funcionários com Avaliação < 7 ({count_abaixo_7})"):
        col_avaliacao = "Avaliação do Funcionário"
        if col_avaliacao in df_f.columns and not df_f[col_avaliacao].dropna().empty:
            df_abaixo_7 = df_f[df_f[col_avaliacao] < 7.0]
            if not df_abaixo_7.empty:
                # Pega nome, cargo, nível, área, salário e avaliação para evitar repetições
                funcionarios_abaixo_7 = df_abaixo_7[['Nome Completo', 'Nível', 'Área', 'Cargo', 'Salario Base', 'Avaliação do Funcionário']].drop_duplicates()
                
                # Formata a lista para exibição
                lista_para_exibir = [
                    f"<li style='color: #ff4b4b;'><b>- {row['Nome Completo']}</b> (Nível: {row['Nível']} | Área: {row['Área']} | Cargo: {row['Cargo']}) - Salário: {brl(row['Salario Base'])} | Avaliação: {row['Avaliação do Funcionário']:.1f}</li>"
                    for index, row in funcionarios_abaixo_7.iterrows()
                ]

                st.markdown(
                    f"""
                    <div style="border: 2px solid #ff4b4b; border-radius: 5px; padding: 10px; background-color: #ffe6e6;">
                        <p style="font-weight: bold; color: #ff4b4b;">
                            Atenção: Os seguintes funcionários possuem avaliação inferior a 7.
                        </p>
                        <hr style="border-top: 1px solid #ff4b4b; margin: 10px 0;">
                        <ul style="list-style-type: none; padding-left: 0;">
                            {''.join(lista_para_exibir)}
                        </ul>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
            else:
                st.info("Nenhum funcionário com avaliação inferior a 7 no filtro atual.")
        else:
            st.info("A coluna 'Avaliação do Funcionário' não está disponível.")

with tab_graficos:
    st.subheader("Análise Gráfica")
    colA, colB = st.columns(2)
    with colA:
        if "Área" in df_f.columns:
            d = df_f.groupby("Área").size().reset_index(name="Headcount")
            if not d.empty:
                fig = px.bar(
                    d,
                    x="Área",
                    y="Headcount",
                    title="Headcount por Área",
                    color_discrete_sequence=[COLOR_PALETTE["main_color"]]
                )
                st.plotly_chart(fig, use_container_width=True)

    with colB:
        if "Cargo" in df_f.columns and "Salario Base" in df_f.columns:
            d = df_f.groupby("Cargo", as_index=False)["Salario Base"].mean().sort_values("Salario Base", ascending=False)
            if not d.empty:
                fig = px.bar(
                    d, 
                    x="Cargo", 
                    y="Salario Base", 
                    title="Salário Médio por Cargo",
                    color_discrete_sequence=[COLOR_PALETTE["secondary_color"]]
                )
                st.plotly_chart(fig, use_container_width=True)

    colC, colD = st.columns(2)
    with colC:
        if "Idade" in df_f.columns and not df_f["Idade"].dropna().empty:
            fig = px.histogram(
                df_f, 
                x="Idade", 
                nbins=20, 
                title="Distribuição de Idade",
                color_discrete_sequence=[COLOR_PALETTE["main_color"]] # Alterado para a cor principal
            )
            st.plotly_chart(fig, use_container_width=True)

    with colD:
        if "Sexo" in df_f.columns:
            d = df_f["Sexo"].value_counts().reset_index()
            d.columns = ["Sexo", "Contagem"]
            if not d.empty:
                fig = px.pie(
                    d,
                    values="Contagem",
                    names="Sexo",
                    title="Distribuição por Sexo",
                    color="Sexo",
                    color_discrete_map={
                        "M": COLOR_PALETTE["M"],
                        "F": COLOR_PALETTE["F"]
                    }
                )
                st.plotly_chart(fig, use_container_width=True)

with tab_tabela:
    st.subheader("Tabela de Funcionários (dados filtrados)")
    st.dataframe(df_f, use_container_width=True)

    csv_bytes = df_f.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Baixar CSV filtrado",
        data=csv_bytes,
        file_name="funcionarios_filtrado.csv",
        mime="text/csv"
    )

    # Exportar Excel filtrado (opcional)
    to_excel = st.toggle("Gerar Excel filtrado para download")
    if to_excel:
        from io import BytesIO
        buff = BytesIO()
        with pd.ExcelWriter(buff, engine="openpyxl") as writer:
            df_f.to_excel(writer, index=False, sheet_name="Filtrado")
        st.download_button(
            "Baixar Excel filtrado",
            data=buff.getvalue(),
            file_name="funcionarios_filtrado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
