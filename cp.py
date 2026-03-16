from __future__ import annotations
import sys
import unicodedata
from pathlib import Path
from typing import Any, Iterable
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
# =========================================================
# FALLBACK PARA AMBIENTE SEM STREAMLIT
# =========================================================
try:
   import streamlit as st  # type: ignore
   STREAMLIT_OK = True
except ModuleNotFoundError:
   STREAMLIT_OK = False
   class _DummyCache:
       def __call__(self, func):
           return func
   class _DummyContext:
       def __enter__(self):
           return self
       def __exit__(self, exc_type, exc, tb):
           return False
       def metric(self, *args, **kwargs):
           return None
       def write(self, *args, **kwargs):
           return None
       def markdown(self, *args, **kwargs):
           return None
       def subheader(self, *args, **kwargs):
           return None
       def info(self, *args, **kwargs):
           return None
       def warning(self, *args, **kwargs):
           return None
       def error(self, *args, **kwargs):
           return None
       def dataframe(self, *args, **kwargs):
           return None
       def caption(self, *args, **kwargs):
           return None
       def pyplot(self, *args, **kwargs):
           return None
   class _DummySidebar(_DummyContext):
       def header(self, *args, **kwargs):
           return None
       def multiselect(self, label, options, default=None, **kwargs):
           return default if default is not None else []
       def date_input(self, label, value=None, **kwargs):
           return value
   class _DummyStreamlit:
       cache_data = _DummyCache()
       sidebar = _DummySidebar()
       def set_page_config(self, *args, **kwargs):
           return None
       def markdown(self, *args, **kwargs):
           return None
       def title(self, *args, **kwargs):
           print(*args)
           return None
       def caption(self, *args, **kwargs):
           return None
       def subheader(self, *args, **kwargs):
           print(*args)
           return None
       def info(self, *args, **kwargs):
           print(*args)
           return None
       def warning(self, *args, **kwargs):
           print(*args)
           return None
       def error(self, *args, **kwargs):
           print("ERRO:", *args)
           return None
       def write(self, *args, **kwargs):
           print(*args)
           return None
       def dataframe(self, *args, **kwargs):
           return None
       def pyplot(self, *args, **kwargs):
           return None
       def metric(self, *args, **kwargs):
           return None
       def columns(self, n):
           return [_DummyContext() for _ in range(n)]
       def tabs(self, names):
           return [_DummyContext() for _ in names]
       def selectbox(self, label, options, index=0, **kwargs):
           if not options:
               return None
           index = max(0, min(index, len(options) - 1))
           return options[index]
       def file_uploader(self, *args, **kwargs):
           return None
       def stop(self):
           raise SystemExit(0)
   st = _DummyStreamlit()
 
st.set_page_config(page_title="Dashboard de Produtividade em Obras", page_icon="📊", layout="wide")
st.markdown(
   """
<style>
.block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
.metric-card {
   background-color: #111827;
   color: white;
   padding: 1rem;
   border-radius: 14px;
   border: 1px solid #374151;
}
.small-note {
   font-size: 0.9rem;
   color: #6b7280;
}
</style>
""",
   unsafe_allow_html=True,
)
st.title("📊 Dashboard de Produtividade em Obras")
st.caption("Análise descritiva da base df_diarios.xlsx com foco em mão de obra e produtividade diária.")
 
# =========================================================
# FUNÇÕES AUXILIARES
# =========================================================
ARQUIVO_PRINCIPAL_CANDIDATOS: list[Path] = [
   Path(r"C:\Users\labsfiap\Downloads\df_diarios.xlsx"),
   Path(__file__).resolve().parent / "df_diarios.xlsx",
]
 
@st.cache_data
def carregar_arquivo(caminho: str) -> pd.DataFrame:
   path = Path(caminho)
   sufixo = path.suffix.lower()
   if sufixo == ".csv":
       return pd.read_csv(path)
   if sufixo in {".xlsx", ".xls"}:
       return pd.read_excel(path)
   raise ValueError(f"Formato não suportado para leitura automática: {path.name}")
 
def carregar_base_principal() -> dict[str, Any] | None:
   caminho_encontrado = next((p for p in ARQUIVO_PRINCIPAL_CANDIDATOS if p.exists()), None)
   if caminho_encontrado is None:
       return None
   df = carregar_arquivo(str(caminho_encontrado))
   if df.empty:
       raise ValueError(f"A base {caminho_encontrado} foi encontrada, mas está vazia.")
   return {"df": df, "caminho": str(caminho_encontrado)}
 
def normalizar_nome(col: Any) -> str:
   texto = str(col).strip().lower()
   texto = unicodedata.normalize("NFKD", texto)
   texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
   # Remove separadores e símbolos para facilitar match (ex.: ip_d, ip d, ip/d -> ipd).
   return "".join(ch for ch in texto if ch.isalnum())
 
def achar_coluna(colunas: Iterable[Any], candidatos: list[str]) -> str | None:
   colunas_norm = [(c, normalizar_nome(c)) for c in colunas]
   for cand in candidatos:
       cand_norm = normalizar_nome(cand)
       if not cand_norm:
           continue
       # 1) prioridade para match exato (evita confundir "Di" com "|Di|", etc.)
       for original, chave_norm in colunas_norm:
           if cand_norm == chave_norm:
               return original
       # 2) fallback por "contém" só para candidatos com 3+ chars (menos ambiguidade)
       if len(cand_norm) >= 3:
           for original, chave_norm in colunas_norm:
               if cand_norm in chave_norm:
                   return original
   return None
 
def classificar_variavel(df: pd.DataFrame, col: str) -> tuple[str, str]:
   serie = df[col]
   nome = normalizar_nome(col)
   if pd.api.types.is_datetime64_any_dtype(serie):
       return "Temporal / quantitativa discreta", "Permite organizar a produtividade ao longo do tempo."
   if "data" in nome or "periodo" in nome or "dia" in nome:
       return "Temporal / quantitativa discreta", "Representa o momento do registro e permite análises por período."
   if pd.api.types.is_numeric_dtype(serie):
       unicos = serie.nunique(dropna=True)
       if unicos <= 12 and pd.api.types.is_integer_dtype(serie):
           return "Quantitativa discreta", "Representa contagens ou valores inteiros usados na análise."
       return "Quantitativa contínua", "Representa medidas numéricas usadas nas análises estatísticas e gráficas."
   return "Qualitativa nominal", "Representa categorias para agrupamento e comparação no dashboard."
 
def resumo_estatistico(serie: pd.Series) -> dict[str, float] | None:
   serie = pd.to_numeric(serie, errors="coerce").dropna()
   if serie.empty:
       return None
   media = float(serie.mean())
   mediana = float(serie.median())
   moda = serie.mode()
   moda_valor = float(moda.iloc[0]) if not moda.empty else np.nan
   amplitude = float(serie.max() - serie.min())
   variancia = float(serie.var())
   desvio = float(serie.std())
   cv = float((desvio / media) * 100) if media != 0 else np.nan
   return {
       "Média": media,
       "Mediana": mediana,
       "Moda": moda_valor,
       "Mínimo": float(serie.min()),
       "Máximo": float(serie.max()),
       "Amplitude": amplitude,
       "Variância": variancia,
       "Desvio-padrão": desvio,
       "Coef. de variação (%)": cv,
       "Qtd. registros": float(len(serie)),
   }
 
def detectar_outliers_iqr(serie: pd.Series) -> pd.Series:
   serie = pd.to_numeric(serie, errors="coerce").dropna()
   if serie.empty:
       return pd.Series(dtype=float)
   q1 = serie.quantile(0.25)
   q3 = serie.quantile(0.75)
   iqr = q3 - q1
   lim_inf = q1 - 1.5 * iqr
   lim_sup = q3 + 1.5 * iqr
   return serie[(serie < lim_inf) | (serie > lim_sup)]
 
def grafico_histograma(serie: pd.Series, titulo: str) -> None:
   fig, ax = plt.subplots(figsize=(8, 4))
   ax.hist(pd.to_numeric(serie, errors="coerce").dropna(), bins=20)
   ax.set_title(titulo)
   ax.set_xlabel("Valor")
   ax.set_ylabel("Frequência")
   st.pyplot(fig)
   plt.close(fig)
 
def grafico_densidade(serie: pd.Series, titulo: str) -> None:
   valores = pd.to_numeric(serie, errors="coerce").dropna()
   if len(valores) < 3:
       st.info("Sem dados suficientes para gráfico de densidade.")
       return
   densidade, bins = np.histogram(valores, bins=20, density=True)
   centros = (bins[:-1] + bins[1:]) / 2
   fig, ax = plt.subplots(figsize=(8, 4))
   ax.plot(centros, densidade, linewidth=2)
   ax.fill_between(centros, densidade, alpha=0.25)
   ax.set_title(titulo)
   ax.set_xlabel("Valor")
   ax.set_ylabel("Densidade aproximada")
   st.pyplot(fig)
   plt.close(fig)
 
def grafico_boxplot_por_grupo(df: pd.DataFrame, col_grupo: str, col_valor: str, titulo: str) -> None:
   grupos = df[[col_grupo, col_valor]].dropna().copy()
   if grupos.empty:
       st.info("Sem dados suficientes para boxplot.")
       return
   categorias = grupos[col_grupo].astype(str).unique().tolist()
   dados = []
   categorias_validas = []
   for cat in categorias:
       serie = pd.to_numeric(grupos.loc[grupos[col_grupo].astype(str) == cat, col_valor], errors="coerce").dropna()
       if len(serie) > 0:
           dados.append(serie)
           categorias_validas.append(cat)
   if not dados:
       st.info("Sem dados suficientes para boxplot.")
       return
   fig, ax = plt.subplots(figsize=(10, 4.8))
   ax.boxplot(dados, tick_labels=categorias_validas, vert=True)
   ax.set_title(titulo)
   ax.set_xlabel(col_grupo)
   ax.set_ylabel(col_valor)
   plt.xticks(rotation=45, ha="right")
   st.pyplot(fig)
   plt.close(fig)
 
def grafico_barras_top(df: pd.DataFrame, col: str, titulo: str, top_n: int = 10) -> None:
   contagem = df[col].astype(str).value_counts().head(top_n)
   if contagem.empty:
       st.info("Sem dados suficientes para gráfico de barras.")
       return
   fig, ax = plt.subplots(figsize=(9, 4.5))
   contagem.plot(kind="bar", ax=ax)
   ax.set_title(titulo)
   ax.set_xlabel(col)
   ax.set_ylabel("Quantidade de registros")
   plt.xticks(rotation=45, ha="right")
   st.pyplot(fig)
   plt.close(fig)
 
def grafico_linha_tempo(df: pd.DataFrame, col_data: str, col_valor: str) -> None:
   base = df[[col_data, col_valor]].dropna().copy()
   if base.empty:
       st.info("Sem dados suficientes para série temporal.")
       return
   base[col_data] = pd.to_datetime(base[col_data], errors="coerce")
   base = base.dropna().sort_values(col_data)
   if base.empty:
       st.info("Sem dados suficientes para série temporal.")
       return
   serie = base.groupby(col_data)[col_valor].mean().reset_index()
   fig, ax = plt.subplots(figsize=(10, 4.5))
   ax.plot(serie[col_data], serie[col_valor])
   ax.set_title(f"Evolução média de {col_valor} no tempo")
   ax.set_xlabel("Data")
   ax.set_ylabel("Média")
   plt.xticks(rotation=45, ha="right")
   st.pyplot(fig)
   plt.close(fig)
 
def tabela_grupo(df: pd.DataFrame, grupo: str, valor: str) -> pd.DataFrame:
   base = df[[grupo, valor]].dropna().copy()
   if base.empty:
       return pd.DataFrame()
   base[valor] = pd.to_numeric(base[valor], errors="coerce")
   base = base.dropna()
   if base.empty:
       return pd.DataFrame()
   resumo = base.groupby(grupo)[valor].agg(["count", "mean", "median", "std", "min", "max"])
   resumo["amplitude"] = resumo["max"] - resumo["min"]
   resumo["cv_%"] = np.where(resumo["mean"] != 0, (resumo["std"] / resumo["mean"]) * 100, np.nan)
   resumo = resumo.sort_values("mean", ascending=False)
   return resumo.reset_index()
 
def resposta_comparacao_grupos(df: pd.DataFrame, col_grupo: str, col_valor: str) -> str:
   tabela = tabela_grupo(df, col_grupo, col_valor)
   if tabela.empty:
       return f"Não há dados suficientes para comparar os grupos de {col_grupo}."
   maior = tabela.iloc[0]
   menor = tabela.sort_values("mean", ascending=True).iloc[0]
   heterogeneidade = float(tabela["cv_%"].mean()) if "cv_%" in tabela.columns else np.nan
   texto_heterogeneidade = (
       "heterogeneidade alta" if pd.notna(heterogeneidade) and heterogeneidade >= 30 else "heterogeneidade moderada/baixa"
   )
   return (
       f"Há diferença entre os grupos de {col_grupo}. "
       f"Maior produtividade média: {maior[col_grupo]} ({maior['mean']:.2f}). "
       f"Menor produtividade média: {menor[col_grupo]} ({menor['mean']:.2f}). "
       f"O comportamento geral indica {texto_heterogeneidade}."
   )
 
def interpretar_media_mediana(media: float, mediana: float) -> str:
   if pd.isna(media) or pd.isna(mediana):
       return "Sem dados suficientes para interpretação."
   diferenca = abs(media - mediana)
   base = abs(mediana) if mediana != 0 else 1
   rel = diferenca / base
   if rel < 0.1:
       return "Média e mediana estão próximas, sugerindo distribuição mais equilibrada e menor influência de valores extremos."
   if media > mediana:
       return "A média está acima da mediana, sugerindo possível presença de valores altos extremos puxando a distribuição para cima."
   return "A média está abaixo da mediana, sugerindo possível presença de valores baixos extremos puxando a distribuição para baixo."
 
def preparar_base(df: pd.DataFrame) -> dict[str, Any]:
   col_ipd = achar_coluna(
       df.columns,
       [
           "ip_d",
           "ip d",
           "ip/d",
           "ipd",
           "i/d",
           "di",
           "indice de produtividade diario",
           "indice produtividade diaria",
           "indice_produtividade_diaria",
           "produtividade diaria",
       ],
   )
   col_tipo_insumo = achar_coluna(df.columns, ["tipo_insumo", "tipo de insumo", "insumo", "categoria insumo"])
   col_obra = achar_coluna(df.columns, ["obra", "id_obra", "nome_obra"])
   col_bloco = achar_coluna(df.columns, ["bloco"])
   col_servico = achar_coluna(df.columns, ["servico", "serviço", "descricao", "descrição", "atividade", "etapa", "frente"])
   col_data = achar_coluna(
       df.columns,
       [
           "data",
           "periodo",
           "período",
           "dia",
           "created",
           "created_at",
           "timestamp",
           "datetime",
       ],
   )
   col_horas = achar_coluna(df.columns, ["horas", "hh", "horas empregadas"])
   col_qtd = achar_coluna(df.columns, ["quantidade", "qtd", "quant_produzida", "quantidade produzida"])
   if col_ipd is None:
       raise ValueError(f"Não encontrei a coluna de produtividade (ex.: ip_d ou Di). Colunas encontradas: {list(df.columns)}")
   for c in [col_ipd, col_horas, col_qtd]:
       if c is not None:
           df[c] = pd.to_numeric(df[c], errors="coerce")
   if col_data is not None:
       df[col_data] = pd.to_datetime(df[col_data], errors="coerce")
   if col_tipo_insumo is not None:
       serie_tipo = df[col_tipo_insumo].astype(str).str.lower().str.strip()
       mascara_mao = (
           serie_tipo.str.contains("mao de obra", na=False)
           | serie_tipo.str.contains("mão de obra", na=False)
           | serie_tipo.str.contains("mao", na=False)
       )
       df_mao = df[mascara_mao].copy()
   else:
       df_mao = df.copy()
   return {
       "df": df,
       "df_mao": df_mao,
       "col_ipd": col_ipd,
       "col_tipo_insumo": col_tipo_insumo,
       "col_obra": col_obra,
       "col_bloco": col_bloco,
       "col_servico": col_servico,
       "col_data": col_data,
       "col_horas": col_horas,
       "col_qtd": col_qtd,
   }
 
def render_dashboard() -> None:
   try:
       base_local = carregar_base_principal()
   except Exception as e:
       st.error(f"Erro ao ler df_diarios.xlsx: {e}")
       st.stop()
   if base_local is None:
       st.error(
           "Não encontrei o arquivo df_diarios.xlsx para leitura automática. "
           "Verifique a existência do arquivo em:\n"
           "- C:/Users/canav/Downloads/df_diarios.xlsx"
       )
       st.stop()
   df = base_local["df"].copy()
   st.caption(f"Base ativa: df_diarios.xlsx — {base_local['caminho']}")
   try:
       base_info = preparar_base(df)
   except Exception as e:
       st.error(str(e))
       st.stop()
   df_mao = base_info["df_mao"]
   col_ipd = base_info["col_ipd"]
   col_obra = base_info["col_obra"]
   col_bloco = base_info["col_bloco"]
   col_servico = base_info["col_servico"]
   col_data = base_info["col_data"]
   col_tipo_insumo = base_info["col_tipo_insumo"]
   st.sidebar.header("Filtros")
   base_filtro = df_mao.copy()
   if col_obra is not None:
       obras = sorted(base_filtro[col_obra].dropna().astype(str).unique().tolist())
       obras_sel = st.sidebar.multiselect("Obra", obras, default=obras)
       if obras_sel:
           base_filtro = base_filtro[base_filtro[col_obra].astype(str).isin(obras_sel)]
   if col_bloco is not None:
       blocos = sorted(base_filtro[col_bloco].dropna().astype(str).unique().tolist())
       blocos_sel = st.sidebar.multiselect("Bloco", blocos, default=blocos)
       if blocos_sel:
           base_filtro = base_filtro[base_filtro[col_bloco].astype(str).isin(blocos_sel)]
   if col_servico is not None:
       servicos = sorted(base_filtro[col_servico].dropna().astype(str).unique().tolist())
       servicos_sel = st.sidebar.multiselect("Serviço / descrição", servicos, default=servicos)
       if servicos_sel:
           base_filtro = base_filtro[base_filtro[col_servico].astype(str).isin(servicos_sel)]
   if col_data is not None and base_filtro[col_data].notna().any():
       data_min = base_filtro[col_data].min().date()
       data_max = base_filtro[col_data].max().date()
       periodo = st.sidebar.date_input("Período", value=(data_min, data_max), min_value=data_min, max_value=data_max)
       if isinstance(periodo, tuple) and len(periodo) == 2:
           ini, fim = periodo
           base_filtro = base_filtro[(base_filtro[col_data].dt.date >= ini) & (base_filtro[col_data].dt.date <= fim)]
   if col_tipo_insumo is not None:
       st.warning(
           "A análise foi filtrada para mão de obra porque comparar produtividade misturando materiais, equipamentos e pessoas pode gerar interpretações distorcidas. "
           "Ao isolar mão de obra, a análise fica mais coerente para gestão, planejamento e orçamento."
       )
   else:
       st.info(
           "A planilha não possui coluna de tipo de insumo. A análise foi feita com todos os registros disponíveis."
       )
   if base_filtro.empty:
       st.error("Os filtros aplicados não retornaram registros.")
       st.stop()
   tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
       "Visão geral",
       "Variáveis",
       "Estatística descritiva",
       "Comparações",
       "Perguntas orientadoras",
       "Video extra",
   ])
   with tab1:
       st.subheader("Resumo da base filtrada")
       c1, c2, c3, c4 = st.columns(4)
       c1.metric("Registros", f"{len(base_filtro):,}".replace(",", "."))
       c2.metric("Colunas", len(base_filtro.columns))
       if col_obra is not None:
           c3.metric("Obras", base_filtro[col_obra].nunique())
       if col_servico is not None:
           c4.metric("Serviços/descrições", base_filtro[col_servico].nunique())
       stats_ipd = resumo_estatistico(base_filtro[col_ipd])
       if stats_ipd:
           a1, a2, a3, a4 = st.columns(4)
           a1.metric(f"Média de {col_ipd}", f"{stats_ipd['Média']:.2f}")
           a2.metric(f"Mediana de {col_ipd}", f"{stats_ipd['Mediana']:.2f}")
           a3.metric("Desvio-padrão", f"{stats_ipd['Desvio-padrão']:.2f}")
           a4.metric("CV (%)", f"{stats_ipd['Coef. de variação (%)']:.2f}")
       st.markdown("### Prévia da base (após filtro de mão de obra)")
       st.dataframe(base_filtro.head(20), use_container_width=True)
       st.markdown(f"### Distribuição de {col_ipd}")
       grafico_histograma(base_filtro[col_ipd], f"Histograma de {col_ipd}")
       grafico_densidade(base_filtro[col_ipd], f"Densidade aproximada de {col_ipd}")
       outliers = detectar_outliers_iqr(base_filtro[col_ipd])
       st.info(f"Quantidade de outliers identificados em {col_ipd} pelo critério do IQR: {len(outliers)}")
   with tab2:
       st.subheader("Reconhecimento e classificação das variáveis")
       linhas = []
       for col in base_filtro.columns:
           tipo, papel = classificar_variavel(base_filtro, col)
           linhas.append({"Coluna": col, "Classificação": tipo, "Papel na análise": papel})
       df_vars = pd.DataFrame(linhas)
       st.dataframe(df_vars, use_container_width=True)
       st.markdown("### Comentário geral")
       st.write(
           "As variáveis qualitativas, como obra, bloco, serviço e tipo de insumo, são usadas para segmentar e comparar os grupos. "
           "As variáveis quantitativas, como ip_d, horas e quantidade produzida, permitem calcular médias, medianas, dispersão e apoiar decisões operacionais."
       )
   with tab3:
       st.subheader("Medidas de posição central e dispersão")
       numericas = [c for c in base_filtro.columns if pd.api.types.is_numeric_dtype(base_filtro[c])]
       if not numericas:
           st.info("Não há variáveis numéricas suficientes para análise estatística.")
       else:
           idx = numericas.index(col_ipd) if col_ipd in numericas else 0
           variavel_num = st.selectbox("Escolha uma variável numérica", numericas, index=idx)
           estat = resumo_estatistico(base_filtro[variavel_num])
           if estat:
               b1, b2, b3 = st.columns(3)
               b1.metric("Média", f"{estat['Média']:.2f}")
               b2.metric("Mediana", f"{estat['Mediana']:.2f}")
               b3.metric("Moda", f"{estat['Moda']:.2f}" if pd.notna(estat['Moda']) else "N/A")
               b4, b5, b6 = st.columns(3)
               b4.metric("Amplitude", f"{estat['Amplitude']:.2f}")
               b5.metric("Variância", f"{estat['Variância']:.2f}")
               b6.metric("Desvio-padrão", f"{estat['Desvio-padrão']:.2f}")
               st.metric("Coeficiente de variação (%)", f"{estat['Coef. de variação (%)']:.2f}")
               st.markdown("### Interpretação")
               st.write(
                   f"- A média de **{variavel_num}** mostra o valor médio observado na base filtrada.\n"
                   f"- A mediana mostra o valor central e costuma ser mais resistente a extremos.\n"
                   f"- A moda mostra o valor mais frequente, quando existir repetição relevante.\n"
                   f"- A amplitude, variância e desvio-padrão ajudam a medir a variabilidade.\n"
                   f"- O coeficiente de variação mostra a dispersão relativa: quanto maior, menor a previsibilidade."
               )
               st.info(interpretar_media_mediana(estat["Média"], estat["Mediana"]))
               grafico_histograma(base_filtro[variavel_num], f"Distribuição de {variavel_num}")
   with tab4:
       st.subheader("Comparações entre grupos")
       opcoes_grupo = [c for c in [col_obra, col_bloco, col_servico] if c is not None]
       if not opcoes_grupo:
           st.info("Não foram detectadas colunas categóricas principais para comparação.")
       else:
           grupo = st.selectbox("Escolha a coluna de comparação", opcoes_grupo)
           tabela = tabela_grupo(base_filtro, grupo, col_ipd)
           if tabela.empty:
               st.info("Não há dados suficientes para comparar os grupos selecionados.")
           else:
               st.dataframe(tabela, use_container_width=True)
               grafico_boxplot_por_grupo(base_filtro, grupo, col_ipd, f"Boxplot de {col_ipd} por {grupo}")
               grafico_barras_top(base_filtro, grupo, f"Quantidade de registros por {grupo}", top_n=15)
               mais_prod = tabela.iloc[0]
               menos_prod = tabela.sort_values("mean", ascending=True).iloc[0]
               mais_estavel = tabela.sort_values("cv_%", ascending=True).iloc[0]
               menos_estavel = tabela.sort_values("cv_%", ascending=False).iloc[0]
               st.markdown("### Destaques automáticos")
               st.write(f"**Maior produtividade média:** {mais_prod[grupo]} (média = {mais_prod['mean']:.2f})")
               st.write(f"**Menor produtividade média:** {menos_prod[grupo]} (média = {menos_prod['mean']:.2f})")
               st.write(f"**Grupo mais previsível:** {mais_estavel[grupo]} (CV = {mais_estavel['cv_%']:.2f}%)")
               st.write(f"**Grupo menos previsível:** {menos_estavel[grupo]} (CV = {menos_estavel['cv_%']:.2f}%)")
   with tab5:
       st.subheader("Perguntas orientadoras respondidas com o dashboard")
       st.markdown("### Leitura exploratória e formulação de perguntas")
       estat_ipd = resumo_estatistico(base_filtro[col_ipd])
       outliers_ipd = detectar_outliers_iqr(base_filtro[col_ipd])
       q1_ipd = float(pd.to_numeric(base_filtro[col_ipd], errors="coerce").quantile(0.25))
       q3_ipd = float(pd.to_numeric(base_filtro[col_ipd], errors="coerce").quantile(0.75))
       if col_obra is not None:
           tab_obras = tabela_grupo(base_filtro, col_obra, col_ipd)
           if not tab_obras.empty:
               obra_mais = tab_obras.iloc[0]
               obra_menos = tab_obras.sort_values("mean", ascending=True).iloc[0]
               resp_obra = (
                   f"{obra_mais[col_obra]} tem a maior média ({obra_mais['mean']:.2f}) e "
                   f"{obra_menos[col_obra]} a menor ({obra_menos['mean']:.2f})."
               )
           else:
               resp_obra = "Não há dados suficientes por obra para comparação."
       else:
           resp_obra = "A base não possui coluna de obra."
       col_grupo_b = col_bloco if col_bloco is not None else col_servico
       if col_grupo_b is not None:
           tab_grupo_b = tabela_grupo(base_filtro, col_grupo_b, col_ipd)
           if not tab_grupo_b.empty:
               grp_mais = tab_grupo_b.iloc[0]
               grp_menos = tab_grupo_b.sort_values("mean", ascending=True).iloc[0]
               resp_grupo = (
                   f"No agrupamento por {col_grupo_b}, o maior desempenho é {grp_mais[col_grupo_b]} "
                   f"({grp_mais['mean']:.2f}) e o menor é {grp_menos[col_grupo_b]} ({grp_menos['mean']:.2f})."
               )
           else:
               resp_grupo = f"Não há dados suficientes para comparar {col_grupo_b}."
       else:
           resp_grupo = "A base não possui coluna de bloco ou serviço."
       if estat_ipd is not None and pd.notna(estat_ipd["Coef. de variação (%)"]):
           resp_previsibilidade = (
               f"CV geral de {col_ipd} = {estat_ipd['Coef. de variação (%)']:.2f}%. "
               "Quanto menor o CV, mais previsível é o desempenho."
           )
       else:
           resp_previsibilidade = "Não há dados suficientes para medir previsibilidade geral."
       resp_outliers = f"Foram identificados {len(outliers_ipd)} outliers em {col_ipd} pelo critério do IQR."
       if col_grupo_b is not None and "tab_grupo_b" in locals() and not tab_grupo_b.empty:
           grupo_critico = tab_grupo_b.sort_values("cv_%", ascending=False).iloc[0]
           resp_prioridade = (
               f"O grupo {grupo_critico[col_grupo_b]} merece atenção prioritária, "
               f"pois apresentou maior variabilidade (CV = {grupo_critico['cv_%']:.2f}%)."
           )
       else:
           resp_prioridade = "Não foi possível apontar um grupo prioritário com os dados atuais."
       resp_orcamento = (
           f"Uma faixa de referência para planejamento está entre Q1={q1_ipd:.2f} e Q3={q3_ipd:.2f} de {col_ipd}, "
           "com ajustes conforme obra e serviço."
       )
       st.markdown("**Pergunta + resposta (resumo):**")
       st.markdown("- **Qual obra tem melhor e pior produtividade?**")
       st.write(f"Resposta: {resp_obra}")
       st.markdown("- **Onde o desempenho varia mais (bloco/serviço)?**")
       st.write(f"Resposta: {resp_grupo}")
       st.markdown("- **A produtividade está previsível?**")
       st.write(f"Resposta: {resp_previsibilidade}")
       st.markdown("- **Há situações atípicas (outliers)?**")
       st.write(f"Resposta: {resp_outliers}")
       st.markdown("- **Quais grupos priorizar para melhoria?**")
       st.write(f"Resposta: {resp_prioridade}")
       st.markdown("- **Que faixa usar como referência de planejamento/orçamento?**")
       st.write(f"Resposta: {resp_orcamento}")
       st.markdown("### A) Existe diferença de produtividade entre obras?")
       if col_obra is None:
           st.write("A base não possui coluna de obra para essa comparação.")
       else:
           st.write(resposta_comparacao_grupos(base_filtro, col_obra, col_ipd))
       st.markdown("### B) Existe diferença entre blocos ou serviços?")
       col_grupo_b = col_bloco if col_bloco is not None else col_servico
       if col_grupo_b is None:
           st.write("A base não possui coluna de bloco ou serviço para essa comparação.")
       else:
           st.write(resposta_comparacao_grupos(base_filtro, col_grupo_b, col_ipd))
       st.markdown("### C) A média e a mediana estão próximas?")
       if estat_ipd is None:
           st.write("Não há dados numéricos suficientes para avaliar média e mediana.")
       else:
           st.write(
               f"Média = {estat_ipd['Média']:.2f} | Mediana = {estat_ipd['Mediana']:.2f}. "
               f"{interpretar_media_mediana(estat_ipd['Média'], estat_ipd['Mediana'])}"
           )
       st.markdown("### D) Qual grupo tem produtividade mais previsível?")
       col_grupo_d = col_obra or col_bloco or col_servico
       if col_grupo_d is None:
           st.write("Não há coluna categórica principal para identificar previsibilidade por grupo.")
       else:
           tab_prev = tabela_grupo(base_filtro, col_grupo_d, col_ipd)
           if tab_prev.empty:
               st.write("Não há dados suficientes para avaliar previsibilidade por grupo.")
           else:
               mais_prev = tab_prev.sort_values("cv_%", ascending=True).iloc[0]
               menos_prev = tab_prev.sort_values("cv_%", ascending=False).iloc[0]
               st.write(
                   f"Grupo mais previsível: {mais_prev[col_grupo_d]} (CV = {mais_prev['cv_%']:.2f}%). "
                   f"Grupo menos previsível: {menos_prev[col_grupo_d]} (CV = {menos_prev['cv_%']:.2f}%)."
               )
   with tab6:
       st.subheader("Video extra")
       st.markdown("### 1) Dúvidas")
       st.caption("Apenas perguntas para orientar o estudo, sem respostas nesta seção.")
       st.markdown(
           "- Como os coeficientes são definidos nas tabelas de orçamento?\n"
           "- Esses coeficientes são sempre os mesmos para todas as obras?\n"
           "- Fatores como experiência da equipe ou condições da obra podem alterar esses valores?\n"
           "- Como verificar se o coeficiente usado no orçamento está próximo da realidade?"
       )
       st.markdown("### 2) Pontos principais que compreendi")
       st.markdown(
           "- O coeficiente representa a relação entre recurso utilizado e unidade de serviço.\n"
           "- Ele aparece nas tabelas de composição de preços usadas no orçamento de obras.\n"
           "- Esse valor ajuda a estimar tempo, custo e quantidade de mão de obra necessária.\n"
           "- O coeficiente transforma dados de produtividade em informação útil para planejamento."
       )
       st.markdown("### 3) Explicação do coeficiente (com minhas palavras)")
       st.write(
           "O coeficiente é um valor que indica quanto de recurso é necessário para produzir uma unidade de serviço. "
           "Nas tabelas de orçamento, ele serve para estimar tempo de execução, quantidade de trabalhadores e custo das atividades. "
           "Ao multiplicar o coeficiente pela quantidade prevista do serviço, torna-se possível estimar os recursos necessários para executar a obra."
       )
   st.markdown("---")
   st.caption("Projeto preparado para apresentação em sala. Ajuste os nomes das colunas se sua planilha usar outro padrão.")
 
# =========================================================
# TESTES LOCAIS
# =========================================================
def _criar_df_teste() -> pd.DataFrame:
   return pd.DataFrame(
       {
           "obra": ["A", "A", "B", "B"],
           "bloco": ["1", "1", "2", "2"],
           "servico": ["Alvenaria", "Pintura", "Alvenaria", "Pintura"],
           "tipo_insumo": ["MÃO DE OBRA", "MÃO DE OBRA", "MATERIAL", "MÃO DE OBRA"],
           "data": pd.to_datetime(["2026-03-01", "2026-03-02", "2026-03-03", "2026-03-04"]),
           "horas": [8, 7, 6, 5],
           "quantidade": [10, 12, 14, 16],
           "ip_d": [1.2, 1.4, 2.5, 1.1],
       }
   )
 
def run_tests() -> None:
   df = _criar_df_teste()
   assert normalizar_nome("Descrição") == "descricao"
   assert achar_coluna(df.columns, ["ip_d"]) == "ip_d"
   estat = resumo_estatistico(df["ip_d"])
   assert estat is not None
   assert round(estat["Média"], 2) == 1.55
   assert round(estat["Mediana"], 2) == 1.30
   assert round(estat["Amplitude"], 2) == 1.40
   outliers = detectar_outliers_iqr(pd.Series([1, 1, 1, 1, 10]))
   assert len(outliers) == 1
   assert float(outliers.iloc[0]) == 10.0
   base_info = preparar_base(df.copy())
   assert base_info["col_ipd"] == "ip_d"
   assert len(base_info["df_mao"]) == 3
   tab = tabela_grupo(base_info["df_mao"], "obra", "ip_d")
   assert not tab.empty
   assert set(tab.columns) >= {"obra", "count", "mean", "median", "std", "min", "max", "amplitude", "cv_%"}
   texto = interpretar_media_mediana(10, 10)
   assert "próximas" in texto or "proximas" in normalizar_nome(texto)
   print("Todos os testes passaram.")
 
if __name__ == "__main__":
   if STREAMLIT_OK:
       render_dashboard()
   else:
       print("Streamlit não está instalado neste ambiente.")
       print("Para rodar a interface, instale com: pip install streamlit")
       print("Executando testes locais...")
       run_tests()