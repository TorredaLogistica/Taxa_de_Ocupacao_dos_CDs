import os
from pathlib import Path
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Taxa de Ocupação dos CDs", page_icon="📦", layout="wide", initial_sidebar_state="expanded")

DIRETORIO_APP = Path(__file__).resolve().parent
ARQUIVO_PADRAO = DIRETORIO_APP / "Ocupação dos CDs.xlsx"
ABA_PADRAO = "Ocupação"
MESES_PT = {1:"janeiro",2:"fevereiro",3:"março",4:"abril",5:"maio",6:"junho",7:"julho",8:"agosto",9:"setembro",10:"outubro",11:"novembro",12:"dezembro"}
CORES_EMPRESA = {"NET":"#1E90FF","EMBRATEL":"#17239B","OMR":"#F26C2F","CLARO TV":"#7A0C96","CLARO MÓVEL":"#E044A7","CLARO MOVEL":"#E044A7","NET PROJETOS":"#50307F","TELMEX":"#2FA84F","CLARO FIXO":"#00A6B2","NEXTEL":"#7F7F7F","REYC":"#F2C811"}
CORES_PADRAO = ["#1E90FF", "#17239B", "#F26C2F", "#7A0C96", "#E044A7", "#50307F", "#2FA84F", "#00A6B2", "#F2C811", "#7F7F7F", "#8B5CF6", "#14B8A6"]

st.markdown("""
<style>
.block-container {padding-top: 2.8rem !important; padding-bottom: 1rem; max-width: 100%;}
.main-title {font-size:30px;font-weight:900;color:#111827;line-height:1.65;margin:0 0 6px 0;padding:14px 0 4px 0;overflow:visible !important;white-space:normal !important;}
.sub-title {font-size:14px;color:#64748b;margin-bottom:18px;}
.metric-card {
    background:linear-gradient(135deg,#eff6ff 0%,#ffffff 100%);
    border:1px solid #bfdbfe;
    border-radius:16px;
    padding:18px 14px;
    box-shadow:0 3px 12px rgba(37,99,235,.10);
    height:158px;
    min-height:158px;
    max-height:158px;
    display:flex;
    flex-direction:column;
    justify-content:center;
    align-items:center;
    text-align:center;
    overflow:hidden;
}
.metric-value {
    font-size:clamp(24px,2vw,34px);
    font-weight:900;
    color:#0f172a;
    line-height:1.05;
    white-space:nowrap;
    text-align:center;
    width:100%;
}
.metric-name {
    font-size:13px;
    font-weight:900;
    color:#0f172a;
    line-height:1.25;
    width:100%; max-width:100%; text-align:center;
    display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical;
    overflow:hidden; word-break:normal;
    margin-top:8px;
}
.metric-label {
    font-size:13px;
    color:#0f172a;
    margin-top:8px;
    font-weight:900;
    line-height:1.25;
    text-align:center;
    width:100%;
}
.metric-label-small {
    font-size:11px;
    color:#0f172a;
    margin-top:8px;
    font-weight:900;
    line-height:1.20;
    text-align:center;
    width:100%;
}
.section-title {font-size:19px;font-weight:900;color:#111827;margin-top:18px;margin-bottom:10px;line-height:1.45;}
.chart-title-fix {font-size:20px;font-weight:900;color:#111827;margin-top:8px;margin-bottom:0px;line-height:1.45;}
.legend-card {border:1px solid #e5e7eb;border-radius:14px;padding:10px 12px;background:#fff;box-shadow:0 2px 8px rgba(15,23,42,.04);max-height:235px;overflow-y:auto;margin-top:6px;}
.legend-row {display:flex;align-items:center;justify-content:space-between;gap:8px;border-bottom:1px solid #f1f5f9;padding:6px 0;font-size:13px;}
.legend-row:last-child {border-bottom:none;}
.legend-left {display:flex;align-items:center;gap:8px;min-width:0;}
.legend-dot {width:11px;height:11px;border-radius:50%;flex:0 0 auto;}
.legend-name {font-weight:800;color:#334155;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.legend-pct {font-weight:900;color:#0f172a;white-space:nowrap;}
[data-testid="stDataFrame"] {border:1px solid #e5e7eb;border-radius:10px;overflow:hidden;}
[data-testid="stVerticalBlock"] {gap:0.75rem;}
</style>
""", unsafe_allow_html=True)

def normalizar_percentual(s):
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce")
    return (s.astype(str).str.replace("%","",regex=False).str.replace(",",".",regex=False).str.replace("-","",regex=False).replace({"":np.nan,"nan":np.nan,"None":np.nan}).pipe(pd.to_numeric, errors="coerce").apply(lambda x: x/100 if pd.notna(x) and x > 1 else x))

def fmt_num(v):
    return "0" if pd.isna(v) else f"{v:,.0f}".replace(",", ".")

def fmt_mil(v):
    if pd.isna(v):
        return "0 Mil"
    return f"{v/1000:,.1f} Mil".replace(",","X").replace(".",",").replace("X",".")

def fmt_pct(v):
    return "0,00%" if pd.isna(v) else f"{v*100:,.2f}%".replace(",","X").replace(".",",").replace("X",".")

def preparar_base(df):
    df=df.copy(); df.columns=[str(c).strip() for c in df.columns]
    df["MÊS_ANO_REFERÊNCIA"]=pd.to_datetime(df["MÊS_ANO_REFERÊNCIA"], errors="coerce")
    df=df.dropna(subset=["MÊS_ANO_REFERÊNCIA"])
    for c in ["Ocupação Armaz. PPP","Capacidade Armaz. PPP","Ocupação Armaz. Blocado","Capacidade Armaz. Blocado","Ocupação Total","Capacidade Total","Ocupação em PPP - Sem consumo 2 anos","Ocupação em PPP - Obsoletos Massivo","Ocupação em PPP - Descarte","Ocupação em PPP - Provisão","Ocupação em PPP - Celular","Ocupação em PPP - Total"]:
        if c in df.columns: df[c]=pd.to_numeric(df[c], errors="coerce").fillna(0)
    for c in ["% Ocupação PPP","% Ocupação Blocado","% Ocupação Total","% Ocupação sem consumo CD"]:
        if c in df.columns: df[c]=normalizar_percentual(df[c])
    df["Ano"]=df["MÊS_ANO_REFERÊNCIA"].dt.year
    df["MesNum"]=df["MÊS_ANO_REFERÊNCIA"].dt.month
    df["Mês"]=df["MesNum"].map(MESES_PT)
    df["Ano_Mes"]=df["MÊS_ANO_REFERÊNCIA"].dt.to_period("M").astype(str)
    return df

@st.cache_data(show_spinner=False)
def carregar_excel(caminho, aba):
    return preparar_base(pd.read_excel(caminho, sheet_name=aba, engine="openpyxl"))

def filtrar(df, col, sel):
    return df[df[col].isin(sel)] if sel else df

def tabela_ocupacao(df_base, grupo):
    t=df_base.groupby(grupo, dropna=False).agg(**{"Cap. Blocado":("Capacidade Armaz. Blocado","sum"),"Armaz. Blocado":("Ocupação Armaz. Blocado","sum"),"Cap. PPP":("Capacidade Armaz. PPP","sum"),"Armaz. PPP":("Ocupação Armaz. PPP","sum"),"Cap. Total":("Capacidade Total","sum"),"Ocup. Total":("Ocupação Total","sum")}).reset_index()
    t["% Ocupação"]=np.where(t["Cap. Total"]>0,t["Ocup. Total"]/t["Cap. Total"],0)
    return t.sort_values("% Ocupação", ascending=False)

def add_total(t, grupo):
    total=pd.DataFrame([{grupo:"Total","Cap. Blocado":t["Cap. Blocado"].sum(),"Armaz. Blocado":t["Armaz. Blocado"].sum(),"Cap. PPP":t["Cap. PPP"].sum(),"Armaz. PPP":t["Armaz. PPP"].sum(),"Cap. Total":t["Cap. Total"].sum(),"Ocup. Total":t["Ocup. Total"].sum()}])
    total["% Ocupação"]=np.where(total["Cap. Total"]>0,total["Ocup. Total"]/total["Cap. Total"],0)
    return pd.concat([t,total], ignore_index=True)

def formatar_tabela(t, grupo, detalhada=True):
    t2=t.copy()
    for c in ["Cap. Blocado","Armaz. Blocado","Cap. PPP","Armaz. PPP","Cap. Total","Ocup. Total"]:
        if c in t2.columns: t2[c]=t2[c].apply(fmt_num)
    t2["% Ocupação"]=t2["% Ocupação"].apply(fmt_pct)
    return t2[[grupo,"Cap. Blocado","Armaz. Blocado","Cap. PPP","Armaz. PPP","Cap. Total","Ocup. Total","% Ocupação"]] if detalhada else t2[[grupo,"% Ocupação"]]

def destacar_total(styler_df, grupo):
    def estilo_linha(row):
        if str(row.get(grupo, "")).strip().lower() == "total":
            return ["background-color: #dbeafe; color: #0f172a; font-weight: 900; border-top: 2px solid #2563eb;" for _ in row]
        return ["" for _ in row]
    return styler_df.style.apply(estilo_linha, axis=1)

def mostrar_tabela_total(df_tab, grupo, altura):
    st.dataframe(destacar_total(df_tab, grupo), use_container_width=True, hide_index=True, height=altura)

def dados_percentual(df_base, grupo):
    t=df_base.groupby(grupo, dropna=False)["Ocupação Total"].sum().reset_index()
    t=t[t["Ocupação Total"]>0].sort_values("Ocupação Total", ascending=False)
    total=t["Ocupação Total"].sum(); t["%"] = np.where(total>0, t["Ocupação Total"]/total, 0)
    return t

def cor_item(nome, idx, grupo):
    return CORES_EMPRESA.get(str(nome), CORES_PADRAO[idx % len(CORES_PADRAO)]) if grupo=="Empresa" else CORES_PADRAO[idx % len(CORES_PADRAO)]

def html_legenda(df_base, grupo):
    t=dados_percentual(df_base, grupo).reset_index(drop=True)
    rows=[]
    for i,row in t.iterrows():
        nome=str(row[grupo]); cor=cor_item(nome,i,grupo)
        rows.append(f'<div class="legend-row"><div class="legend-left"><span class="legend-dot" style="background:{cor};"></span><span class="legend-name">{nome}</span></div><span class="legend-pct">{fmt_pct(row["%"])}</span></div>')
    return '<div class="legend-card">' + ''.join(rows) + '</div>'

def grafico_rosca(df_base, grupo, titulo):
    t=dados_percentual(df_base, grupo)
    if t.empty: return go.Figure()

    # Percentual formatado com 2 casas decimais e vírgula decimal.
    # Exemplo: 44,18% em vez de 44,2%.
    t["% Label"] = t["%"].apply(fmt_pct)

    color_map={str(row[grupo]): cor_item(str(row[grupo]), i, grupo) for i,row in t.reset_index(drop=True).iterrows()}
    fig=px.pie(t, names=grupo, values="Ocupação Total", hole=.58, color=grupo, color_discrete_map=color_map, title=titulo, custom_data=["% Label"])
    fig.update_traces(textposition="inside", textinfo="text", texttemplate="%{customdata[0]}", textfont=dict(size=14,color="white",family="Arial Black"), insidetextorientation="radial")
    fig.update_layout(height=350, margin=dict(l=10,r=10,t=52,b=8), showlegend=False, title=dict(font=dict(size=15), x=.02, xanchor="left"), uniformtext_minsize=9, uniformtext_mode="hide")
    return fig

def grafico_barras_linha(df_base):
    d=df_base.copy(); g=d.groupby(["Ano","MesNum","Mês"]).agg({"Capacidade Total":"sum","Ocupação Total":"sum"}).reset_index().sort_values(["Ano","MesNum"])
    g["Eixo"]=g["Mês"].str.capitalize()+"<br>"+g["Ano"].astype(str)
    g["Ocupação Total %"]=np.where(g["Capacidade Total"]>0,g["Ocupação Total"]/g["Capacidade Total"],0)
    fig=go.Figure()
    fig.add_bar(x=g["Eixo"], y=g["Capacidade Total"], name="Capacidade Total", marker_color="#1E90FF", offsetgroup=1)
    fig.add_bar(x=g["Eixo"], y=g["Ocupação Total"], name="Ocupação Total", marker_color="#17239B", offsetgroup=2)
    fig.add_trace(go.Scatter(x=g["Eixo"], y=g["Ocupação Total %"]*100, mode="lines+markers", name="Ocupação Total %", yaxis="y2", line=dict(color="#F26C2F", width=3.2), marker=dict(size=8)))
    max_vol=max(g["Capacidade Total"].max(),g["Ocupação Total"].max()) if not g.empty else 0
    y2_min=max(0,(g["Ocupação Total %"].min()*100)-8) if not g.empty else 0; y2_max=min(100,(g["Ocupação Total %"].max()*100)+8) if not g.empty else 100
    if y2_max-y2_min<15: y2_min=max(0,y2_min-5); y2_max=min(100,y2_max+5)
    fig.update_layout(barmode="group",height=610,title_text="",margin=dict(l=60,r=76,t=82,b=60),yaxis=dict(title="Volume",range=[0,max_vol*1.22 if max_vol else 1],tickformat=",.0f",gridcolor="#e5e7eb"),yaxis2=dict(title="% Ocupação",overlaying="y",side="right",range=[y2_min,y2_max],ticksuffix="%",showgrid=False),xaxis=dict(title="",tickfont=dict(size=11),automargin=True),legend=dict(orientation="h",yanchor="bottom",y=1.05,xanchor="left",x=.01,font=dict(size=12),title_text=""),plot_bgcolor="white",paper_bgcolor="white")
    for _,row in g.iterrows():
        x=row["Eixo"]; cap=row["Capacidade Total"]; occ=row["Ocupação Total"]; pct=row["Ocupação Total %"]*100
        fig.add_annotation(x=x,y=cap,xref="x",yref="y",text=fmt_num(cap),showarrow=False,xshift=-28,yshift=-16,font=dict(size=11,color="#475569",family="Arial Black"),bgcolor="#dbeafe",bordercolor="#dbeafe",borderpad=4,opacity=.96)
        fig.add_annotation(x=x,y=occ,xref="x",yref="y",text=fmt_num(occ),showarrow=False,xshift=28,yshift=-16,font=dict(size=11,color="#475569",family="Arial Black"),bgcolor="#c7d2fe",bordercolor="#c7d2fe",borderpad=4,opacity=.96)
        fig.add_annotation(x=x,y=pct,xref="x",yref="y2",text=f"<b>{fmt_pct(pct/100)}</b>",showarrow=False,yshift=18,font=dict(size=15,color="#475569",family="Arial Black"),bgcolor="#e5e7eb",bordercolor="#e5e7eb",borderpad=4,opacity=.98)
    return fig

def maior_ocupacao(df_base, grupo):
    t=tabela_ocupacao(df_base, grupo)
    t=t[(t["Cap. Total"] > 0) & (t[grupo].notna())]
    if t.empty: return "-", 0
    r=t.iloc[0]
    return str(r[grupo]), float(r["% Ocupação"])

st.markdown('<div class="main-title">📦 Taxa de Ocupação dos CDs</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Indicador de capacidade, ocupação total e taxa de ocupação por CD, UF, Empresa e Operador Logístico.</div>', unsafe_allow_html=True)
try:
    if not os.path.exists(ARQUIVO_PADRAO): st.error(f"Arquivo '{ARQUIVO_PADRAO}' não encontrado. Coloque a base na mesma pasta do app."); st.stop()
    df=carregar_excel(ARQUIVO_PADRAO, ABA_PADRAO)
except Exception as e:
    st.error(f"Erro ao carregar a base: {e}"); st.stop()

with st.sidebar:
    st.header("🔎 Filtros principais")
    data_ref=df["MÊS_ANO_REFERÊNCIA"].max(); ano_padrao=int(data_ref.year) if pd.notna(data_ref) else None; mes_padrao=MESES_PT.get(int(data_ref.month)) if pd.notna(data_ref) else None
    anos=sorted(df["Ano"].dropna().unique().tolist())
    ano_sel=st.multiselect("Ano", anos, default=[ano_padrao] if ano_padrao in anos else (anos[-1:] if anos else []))
    df_f=filtrar(df,"Ano",ano_sel)
    meses_lista=df_f[["MesNum","Mês"]].drop_duplicates().sort_values("MesNum")["Mês"].tolist(); mes_default=[mes_padrao] if mes_padrao in meses_lista else (meses_lista[-1:] if meses_lista else [])
    mes_sel=st.multiselect("Mês", meses_lista, default=mes_default)
    if mes_sel: df_f=df_f[df_f["Mês"].isin(mes_sel)]
    for col,label in [("Empresa","Empresa"),("UF","UF"),("Unidade","Unidade"),("Operador Logístico","Operador Logístico")]:
        sel=st.multiselect(label, sorted(df_f[col].dropna().unique().tolist())); df_f=filtrar(df_f,col,sel)
if df_f.empty: st.warning("Nenhum dado encontrado para os filtros selecionados."); st.stop()

cap_total=df_f["Capacidade Total"].sum(); occ_total=df_f["Ocupação Total"].sum(); pct_total=occ_total/cap_total if cap_total else 0
maior_unidade, pct_maior_unidade = maior_ocupacao(df_f, "Unidade")
maior_uf, pct_maior_uf = maior_ocupacao(df_f, "UF")

k1,k2,k3,k4,k5=st.columns([1,1,1,1,1], gap="large")
with k1: st.markdown(f'<div class="metric-card"><div class="metric-value">{fmt_mil(occ_total)}</div><div class="metric-label">Ocupação Total</div></div>', unsafe_allow_html=True)
with k2: st.markdown(f'<div class="metric-card"><div class="metric-value">{fmt_mil(cap_total)}</div><div class="metric-label">Capacidade Total</div></div>', unsafe_allow_html=True)
with k3: st.markdown(f'<div class="metric-card"><div class="metric-value">{fmt_pct(pct_total)}</div><div class="metric-label">% Ocupação Total</div></div>', unsafe_allow_html=True)
with k4: st.markdown(f'<div class="metric-card"><div class="metric-value">{fmt_pct(pct_maior_unidade)}</div><div class="metric-name" title="{maior_unidade}">{maior_unidade}</div><div class="metric-label-small">Unidade com maior ocupação</div></div>', unsafe_allow_html=True)
with k5: st.markdown(f'<div class="metric-card"><div class="metric-value">{fmt_pct(pct_maior_uf)}</div><div class="metric-name" title="{maior_uf}">{maior_uf}</div><div class="metric-label-small">UF com maior ocupação</div></div>', unsafe_allow_html=True)

st.markdown('<div class="section-title">📋 Tabelas</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">Relação por Unidade</div>', unsafe_allow_html=True)
mostrar_tabela_total(formatar_tabela(add_total(tabela_ocupacao(df_f,"Unidade"),"Unidade"),"Unidade",True), "Unidade", 370)
uf_col,op_col=st.columns(2,gap="large")
with uf_col:
    st.markdown('<div class="section-title">UF</div>', unsafe_allow_html=True)
    mostrar_tabela_total(formatar_tabela(add_total(tabela_ocupacao(df_f,"UF"),"UF"),"UF",False), "UF", 300)
with op_col:
    st.markdown('<div class="section-title">Operador Logístico</div>', unsafe_allow_html=True)
    mostrar_tabela_total(formatar_tabela(add_total(tabela_ocupacao(df_f,"Operador Logístico"),"Operador Logístico"),"Operador Logístico",False), "Operador Logístico", 300)

st.markdown('<div class="section-title">🍩 Gráficos de Ocupação - Total</div>', unsafe_allow_html=True)
g1,g2,g3=st.columns(3,gap="large")
with g1:
    st.plotly_chart(grafico_rosca(df_f,"Operador Logístico","Ocupação por Operador"), use_container_width=True, config={"displayModeBar":False})
    st.markdown(html_legenda(df_f,"Operador Logístico"), unsafe_allow_html=True)
with g2:
    st.plotly_chart(grafico_rosca(df_f,"Empresa","Ocupação por Empresa"), use_container_width=True, config={"displayModeBar":False})
    st.markdown(html_legenda(df_f,"Empresa"), unsafe_allow_html=True)
with g3:
    st.plotly_chart(grafico_rosca(df_f,"UF","Ocupação por UF"), use_container_width=True, config={"displayModeBar":False})
    st.markdown(html_legenda(df_f,"UF"), unsafe_allow_html=True)

st.divider(); st.markdown('<div class="section-title">📊 Evolução de Capacidade x Ocupação Total</div>', unsafe_allow_html=True)
st.caption("Este gráfico possui filtros próprios e independentes dos filtros principais da lateral.")
with st.expander("🔎 Filtros independentes do gráfico de evolução", expanded=True):
    st.markdown("**Ano/Mês do gráfico**")
    anos_g=sorted(df["Ano"].dropna().unique().tolist())
    anos_g_sel=st.multiselect("Selecione o(s) ano(s)", anos_g, default=[ano_padrao] if ano_padrao in anos_g else (anos_g[-1:] if anos_g else []), key="ano_mes_g")
    meses_por_ano={}
    if anos_g_sel:
        cols_anos=st.columns(min(len(anos_g_sel),3))
        for idx,ano in enumerate(anos_g_sel):
            meses_ano=df[df["Ano"]==ano][["MesNum","Mês"]].drop_duplicates().sort_values("MesNum")["Mês"].tolist()
            meses_default=meses_ano if ano==ano_padrao else []
            with cols_anos[idx % len(cols_anos)]: meses_por_ano[ano]=st.multiselect(f"Meses de {ano}", meses_ano, default=meses_default, key=f"meses_g_{ano}")
    df_g_parts=[df[(df["Ano"]==ano)&(df["Mês"].isin(meses))] for ano,meses in meses_por_ano.items() if meses]
    df_g=pd.concat(df_g_parts, ignore_index=True) if df_g_parts else df.iloc[0:0].copy()
    f3,f4,f5,f6=st.columns(4)
    for col,label,key,container in [("Empresa","Empresa do gráfico","empresa_g",f3),("UF","UF do gráfico","uf_g",f4),("Unidade","Unidade do gráfico","unidade_g",f5),("Operador Logístico","Operador do gráfico","operador_g",f6)]:
        with container: sel=st.multiselect(label, sorted(df_g[col].dropna().unique().tolist()), key=key)
        df_g=filtrar(df_g,col,sel)
if df_g.empty: st.warning("Nenhum dado encontrado para os filtros independentes do gráfico.")
else:
    st.markdown('<div class="chart-title-fix">Evolução de Capacidade x Ocupação Total</div>', unsafe_allow_html=True)
    st.plotly_chart(grafico_barras_linha(df_g), use_container_width=True, config={"displayModeBar":False})

with st.expander("⬇️ Baixar dados filtrados"):
    st.download_button("Baixar base filtrada em CSV", data=df_f.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig"), file_name="ocupacao_cds_filtrado.csv", mime="text/csv")
st.caption("Desenvolvido em Python/Streamlit | Indicador Taxa de Ocupação dos CDs")
