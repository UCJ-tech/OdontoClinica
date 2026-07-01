

import io
import base64
from datetime import date, timedelta

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# Importa as funções de negócio do script principal
from analise_dentistas import (
    buscar_estimates,
    processar_dados,
    gerar_excel,
    gerar_basic_auth,
    BASE_URL,
    ENDPOINT,
)

# ── Paleta UCJ ────────────────────────────────────────────────────────────────
BORDE      = "#871F1B"
BORDE_ESC  = "#5C1F1D"
BORDE_CLR  = "#FAEDEC"
BORDE_MCL  = "#F5D4D3"
VERDE_OK   = "#1A7A1A"
AMARELO    = "#7A5C00"
CINZA      = "#6C757D"
CINZA_ESC  = "#2C2C2C"

# ── Configuração da página ────────────────────────────────────────────────────
st.set_page_config(
    page_title="UCJ — Faturamento por Dentista",
    page_icon="",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS customizado UCJ ───────────────────────────────────────────────────────
st.markdown(f"""
<style>
  /* Fundo e fonte geral */
  html, body, [data-testid="stAppViewContainer"] {{
      background-color: #FAFAFA;
      font-family: Arial, sans-serif;
  }}

  /* Sidebar */
  [data-testid="stSidebar"] {{
      background-color: {BORDE_ESC} !important;
  }}
  [data-testid="stSidebar"] * {{
      color: #FFFFFF !important;
  }}
  [data-testid="stSidebar"] .stDateInput label,
  [data-testid="stSidebar"] .stTextInput label,
  [data-testid="stSidebar"] .stSelectbox label {{
      color: {BORDE_CLR} !important;
      font-weight: 600;
      font-size: 0.82rem;
      letter-spacing: 0.05em;
      text-transform: uppercase;
  }}
  [data-testid="stSidebar"] input {{
      background: rgba(255,255,255,0.1) !important;
      color: #fff !important;
      border: 1px solid rgba(255,255,255,0.3) !important;
      border-radius: 6px !important;
  }}
  [data-testid="stSidebar"] hr {{
      border-color: rgba(255,255,255,0.2) !important;
  }}

  /* Botão primário */
  .stButton > button {{
      background-color: {BORDE} !important;
      color: #ffffff !important;
      border: none !important;
      border-radius: 6px !important;
      font-weight: 700 !important;
      font-size: 0.9rem !important;
      padding: 0.55rem 1.4rem !important;
      width: 100%;
      transition: background 0.2s;
  }}
  .stButton > button:hover {{
      background-color: {BORDE_ESC} !important;
  }}

  /* Download button */
  .stDownloadButton > button {{
      background-color: #FFFFFF !important;
      color: {BORDE} !important;
      border: 2px solid {BORDE} !important;
      border-radius: 6px !important;
      font-weight: 700 !important;
      width: 100%;
  }}
  .stDownloadButton > button:hover {{
      background-color: {BORDE_CLR} !important;
  }}

  /* Cartões de métricas */
  [data-testid="stMetric"] {{
      background: #FFFFFF;
      border-left: 5px solid {BORDE};
      border-radius: 8px;
      padding: 16px 20px !important;
      box-shadow: 0 1px 4px rgba(0,0,0,0.07);
  }}
  [data-testid="stMetricLabel"] {{
      font-size: 0.72rem !important;
      font-weight: 700 !important;
      text-transform: uppercase;
      letter-spacing: 0.06em;
      color: {BORDE_ESC} !important;
  }}
  [data-testid="stMetricValue"] {{
      font-size: 1.6rem !important;
      font-weight: 800 !important;
      color: {BORDE} !important;
  }}

  /* Tabs */
  .stTabs [data-baseweb="tab-list"] {{
      border-bottom: 2px solid {BORDE_MCL};
      gap: 4px;
  }}
  .stTabs [data-baseweb="tab"] {{
      font-weight: 600;
      font-size: 0.85rem;
      color: {CINZA} !important;
      border-radius: 6px 6px 0 0;
      padding: 8px 18px;
  }}
  .stTabs [aria-selected="true"] {{
      background-color: {BORDE_CLR} !important;
      color: {BORDE} !important;
      border-top: 3px solid {BORDE} !important;
  }}

  /* Cabeçalho topo */
  .ucj-header {{
      background: {BORDE};
      color: white;
      padding: 18px 28px;
      border-radius: 10px;
      margin-bottom: 24px;
      display: flex;
      align-items: center;
      justify-content: space-between;
  }}
  .ucj-header h1 {{
      margin: 0;
      font-size: 1.4rem;
      font-weight: 800;
      letter-spacing: 0.02em;
  }}
  .ucj-header span {{
      font-size: 0.82rem;
      opacity: 0.8;
  }}

  /* Semáforo badges */
  .badge-manter  {{ background:#E8F5E9; color:#1A7A1A; padding:3px 10px; border-radius:12px; font-weight:700; font-size:0.78rem; }}
  .badge-avaliar {{ background:#FFF9E6; color:#7A5C00; padding:3px 10px; border-radius:12px; font-weight:700; font-size:0.78rem; }}
  .badge-revisar {{ background:#FDECEA; color:#871F1B; padding:3px 10px; border-radius:12px; font-weight:700; font-size:0.78rem; }}

  /* Tabela de ranking */
  .rank-table {{ width:100%; border-collapse:collapse; font-size:0.85rem; }}
  .rank-table th {{
      background:{BORDE}; color:#fff; padding:10px 14px;
      text-align:left; font-size:0.78rem; text-transform:uppercase; letter-spacing:0.05em;
  }}
  .rank-table td {{
      padding:10px 14px; border-bottom:1px solid #F0E4E3;
      vertical-align:middle; color:{CINZA_ESC};
  }}
  .rank-table tr:nth-child(even) td {{ background:#FBF5F5; color:{CINZA_ESC}; }}
  .rank-table tr:hover td {{ background:#F5D4D3; color:{CINZA_ESC}; }}

  /* Seção título */
  .section-title {{
      font-size:0.75rem; font-weight:700; text-transform:uppercase;
      letter-spacing:0.08em; color:{BORDE_ESC}; margin:24px 0 12px;
      border-left:4px solid {BORDE}; padding-left:10px;
  }}

  /* Status chips */
  .chip-approved {{ background:#E8F5E9; color:#1A7A1A; padding:2px 10px; border-radius:10px; font-size:0.78rem; font-weight:600; }}
  .chip-open     {{ background:#FFF9E6; color:#7A5C00; padding:2px 10px; border-radius:10px; font-size:0.78rem; font-weight:600; }}
  .chip-rejected {{ background:#FDECEA; color:#871F1B; padding:2px 10px; border-radius:10px; font-size:0.78rem; font-weight:600; }}

  div[data-testid="stHorizontalBlock"] {{ gap: 16px; }}
  .stAlert {{ border-radius: 8px !important; }}
</style>
""", unsafe_allow_html=True)


# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt_brl(v):
    return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_brl0(v):
    return f"R$ {int(v):,}".replace(",", ".")

def semaforo(rank, n):
    t1, t2 = max(1, n // 3), max(2, 2 * n // 3)
    if rank < t1:
        return '<span class="badge-manter">MANTER</span>'
    elif rank < t2:
        return '<span class="badge-avaliar">AVALIAR</span>'
    else:
        return '<span class="badge-revisar">REVISAR</span>'

def chip_status(s):
    m = {"APPROVED": "approved", "OPEN": "open", "REJECTED": "rejected"}
    k = m.get(s, "open")
    label = {"APPROVED": "Aprovado", "OPEN": "Em Aberto", "REJECTED": "Rejeitado"}.get(s, s)
    return f'<span class="chip-{k}">{label}</span>'

def gerar_excel_bytes(df_det, df_orc, df_exec, df_stat, d_ini, d_fim):
    """Gera o Excel em memória e retorna bytes para download."""
    import tempfile, os
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        path = tmp.name
    import analise_dentistas as ad
    ad.OUTPUT_FILE = path
    gerar_excel(df_det, df_orc, df_exec, df_stat, d_ini, d_fim)
    with open(path, "rb") as f:
        data = f.read()
    os.unlink(path)
    return data


# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
        <div style='text-align:center; padding:16px 0 8px'>
            <div style='font-size:2rem; font-weight:900; letter-spacing:0.04em;
                        color:#FAEDEC; line-height:1'>UCJ</div>
            <div style='font-size:0.72rem; color:rgba(255,255,255,0.65);
                        margin-top:2px; letter-spacing:0.1em'>UFMG CONSULTORIA JUNIOR</div>
        </div>
        <hr>
    """, unsafe_allow_html=True)

    st.markdown("**CREDENCIAIS DA API**")

    subscriber_id = st.text_input("Subscriber ID", placeholder="Ex: 12345")
    clinic_id     = st.text_input("Clinic ID", placeholder="Deixe vazio se não usar")
    username      = st.text_input("Usuario", placeholder="seu@email.com")
    password      = st.text_input("Senha", type="password", placeholder="••••••••")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown("**PERIODO DE ANALISE**")

    hoje   = date.today()
    d_ini  = st.date_input("Data inicial", value=hoje.replace(day=1),
                            max_value=hoje, format="DD/MM/YYYY")
    d_fim  = st.date_input("Data final",   value=hoje,
                            max_value=hoje, format="DD/MM/YYYY")

    st.markdown("<hr>", unsafe_allow_html=True)
    buscar = st.button("Buscar dados")

    st.markdown("""
        <div style='position:fixed; bottom:20px; font-size:0.7rem;
                    color:rgba(255,255,255,0.4); text-align:center; width:220px'>
            UCJ · Dados via API Clinicorp
        </div>
    """, unsafe_allow_html=True)


# ── CABEÇALHO PRINCIPAL ───────────────────────────────────────────────────────
st.markdown(f"""
<div class="ucj-header">
  <div>
    <h1>Analise de Faturamento por Dentista</h1>
    <span>Decisao estrategica de manutencao de profissionais</span>
  </div>
  <div style='text-align:right; opacity:0.7; font-size:0.8rem'>
    UFMG Consultoria Junior<br>via API Clinicorp
  </div>
</div>
""", unsafe_allow_html=True)


# ── ESTADO DA SESSÃO ──────────────────────────────────────────────────────────
if "dados" not in st.session_state:
    st.session_state.dados = None

# ── AÇÃO: BUSCAR ──────────────────────────────────────────────────────────────
if buscar:
    erros = []
    if not subscriber_id: erros.append("Subscriber ID")
    if not username:       erros.append("Usuario")
    if not password:       erros.append("Senha")
    if d_ini > d_fim:      erros.append("Data inicial deve ser anterior à data final")

    if erros:
        st.error("Preencha os campos obrigatorios: " + " · ".join(erros))
    else:
        import analise_dentistas as ad
        ad.SUBSCRIBER_ID = subscriber_id
        ad.CLINIC_ID     = clinic_id
        ad.USERNAME      = username
        ad.PASSWORD      = password

        with st.spinner("Consultando API Clinicorp..."):
            try:
                registros = buscar_estimates(
                    d_ini.strftime("%Y-%m-%d"),
                    d_fim.strftime("%Y-%m-%d"),
                )
                df_det, df_orc, df_exec, df_stat = processar_dados(registros)
                st.session_state.dados = {
                    "df_det":  df_det,
                    "df_orc":  df_orc,
                    "df_exec": df_exec,
                    "df_stat": df_stat,
                    "d_ini":   d_ini.strftime("%Y-%m-%d"),
                    "d_fim":   d_fim.strftime("%Y-%m-%d"),
                }
                st.success(f"{len(registros)} orcamentos carregados com sucesso.")
            except SystemExit:
                st.error("Erro ao consultar a API. Verifique as credenciais e tente novamente.")
            except Exception as e:
                st.error(f"Erro inesperado: {e}")


# ── DASHBOARD ─────────────────────────────────────────────────────────────────
if st.session_state.dados:
    d      = st.session_state.dados
    df_det = d["df_det"]
    df_orc = d["df_orc"]
    df_stat= d["df_stat"]
    d_ini  = d["d_ini"]
    d_fim  = d["d_fim"]

    # Pré-processar
    df_aprov = df_det[df_det["Status_Orcamento"] == "APPROVED"]
    aprov_por = (
        df_aprov.groupby("Profissional_Executante")
        .agg(Fat_Aprovado=("Valor_Final","sum"),
             Proc_Aprovado=("Procedimento","count"))
        .reset_index()
        .rename(columns={"Profissional_Executante":"Profissional"})
    )

    df_rank = df_orc.drop(columns=["Visao"], errors="ignore").copy()
    df_rank = df_rank.merge(aprov_por, on="Profissional", how="left")
    df_rank["Fat_Aprovado"]   = df_rank["Fat_Aprovado"].fillna(0).astype(int)
    df_rank["Taxa_Aprov_Pct"] = (
        df_rank["Fat_Aprovado"] / df_rank["Faturamento_Total"].replace(0,1) * 100
    ).round(1)
    df_rank = df_rank.sort_values("Faturamento_Total", ascending=False).reset_index(drop=True)

    total_fat   = int(df_rank["Faturamento_Total"].sum())
    total_aprov = int(df_rank["Fat_Aprovado"].sum())
    total_proc  = int(df_rank["Qtd_Procedimentos"].sum())
    n_dent      = len(df_rank)
    ticket      = round(total_fat / max(total_proc, 1), 2)
    n           = len(df_rank)

    # ── KPIs ─────────────────────────────────────────────────────────────────
    st.markdown('<p class="section-title">Resumo do periodo</p>', unsafe_allow_html=True)
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Faturamento Total",  fmt_brl0(total_fat))
    c2.metric("Faturamento Aprovado", fmt_brl0(total_aprov))
    c3.metric("Procedimentos",      str(total_proc))
    c4.metric("Dentistas Ativos",   str(n_dent))
    c5.metric("Ticket Medio",       fmt_brl(ticket))

    st.markdown("<br>", unsafe_allow_html=True)

    # ── TABS ─────────────────────────────────────────────────────────────────
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "Ranking de Dentistas",
        "Analise de Procedimentos",
        "Graficos",
        "Status dos Orcamentos",
        "Dados Detalhados",
    ])

    # ── TAB 1: RANKING ────────────────────────────────────────────────────────
    with tab1:
        st.markdown('<p class="section-title">Ranking por faturamento — recomendacao de manutencao</p>',
                    unsafe_allow_html=True)

        linhas_html = ""
        for i, row in df_rank.iterrows():
            badge  = semaforo(i, n)
            td = f'padding:10px 14px;border-bottom:1px solid #F0E4E3;vertical-align:middle;color:{CINZA_ESC};'
            bg = "background:#FBF5F5;" if i % 2 != 0 else ""
            linhas_html += f"""
            <tr style='{bg}'>
              <td style='{td}font-weight:700;color:{BORDE};text-align:center'>{i+1}</td>
              <td style='{td}font-weight:600'>{row['Profissional'].title()}</td>
              <td style='{td}'>{badge}</td>
              <td style='{td}font-weight:700;color:{BORDE_ESC}'>{fmt_brl0(row['Faturamento_Total'])}</td>
              <td style='{td}color:{"#1A7A1A" if row["Fat_Aprovado"]>0 else CINZA}'>{fmt_brl0(row['Fat_Aprovado'])}</td>
              <td style='{td}text-align:center'>{row['Taxa_Aprov_Pct']:.1f}%</td>
              <td style='{td}text-align:center'>{int(row['Qtd_Procedimentos'])}</td>
              <td style='{td}text-align:center'>{int(row['Qtd_Tratamentos'])}</td>
              <td style='{td}text-align:center'>{fmt_brl(row['Ticket_Medio_Procedimento'])}</td>
            </tr>"""

        st.markdown(f"""
        <table class="rank-table">
          <thead>
            <tr>
              <th>#</th><th>Dentista</th><th>Recomendacao</th>
              <th>Fat. Total</th><th>Fat. Aprovado</th>
              <th>Taxa Aprov.</th><th>Procedimentos</th>
              <th>Tratamentos</th><th>Ticket Medio</th>
            </tr>
          </thead>
          <tbody>{linhas_html}</tbody>
        </table>
        <br>
        <div style='font-size:0.75rem;color:{CINZA};padding:8px;
                    background:{BORDE_CLR};border-radius:6px;'>
          <b>MANTER</b> = top 1/3 por faturamento &nbsp;|&nbsp;
          <b>AVALIAR</b> = faixa media &nbsp;|&nbsp;
          <b>REVISAR</b> = bottom 1/3 — analisar custo x beneficio
        </div>
        """, unsafe_allow_html=True)

    # ── TAB 2: ANÁLISE DE PROCEDIMENTOS ──────────────────────────────────────
    with tab2:
        df_atv = df_det[df_det["Deletado"] == False].copy()

        # Pivot procedimento x status
        pv = (
            df_atv.groupby(["Procedimento", "Status_Orcamento"])
            .agg(Qtd=("Valor_Final","count"), Valor=("Valor_Final","sum"))
            .reset_index()
        )
        pivot = pv.pivot_table(
            index="Procedimento", columns="Status_Orcamento",
            values=["Qtd","Valor"], fill_value=0,
        )
        pivot.columns = [f"{stat}_{m}" for m, stat in pivot.columns]
        pivot = pivot.reset_index()
        for col in ["APPROVED_Qtd","APPROVED_Valor","OPEN_Qtd","OPEN_Valor",
                    "REJECTED_Qtd","REJECTED_Valor"]:
            if col not in pivot.columns:
                pivot[col] = 0

        # Dentistas por procedimento
        dent_por_proc = (
            df_atv.groupby("Procedimento")["Profissional_Executante"]
            .agg(N_Dentistas="nunique",
                 Dentistas=lambda x: ", ".join(sorted(x.dropna().unique())))
            .reset_index()
        )
        ticket_proc = (
            df_atv[df_atv["Status_Orcamento"] == "APPROVED"]
            .groupby("Procedimento")["Valor_Final"]
            .mean().round(2).reset_index()
            .rename(columns={"Valor_Final":"Ticket_Medio_Aprovado"})
        )
        df_proc = (
            pivot
            .merge(dent_por_proc, on="Procedimento", how="left")
            .merge(ticket_proc,   on="Procedimento", how="left")
        )
        df_proc["Ticket_Medio_Aprovado"] = df_proc["Ticket_Medio_Aprovado"].fillna(0)
        df_proc["Total_Orcado"] = (
            df_proc["APPROVED_Qtd"] + df_proc["OPEN_Qtd"] + df_proc["REJECTED_Qtd"]
        ).astype(int)
        df_proc = df_proc.sort_values("APPROVED_Valor", ascending=False).reset_index(drop=True)

        # ── KPIs de procedimento ──────────────────────────────────────────────
        n_proc_unicos = len(df_proc)
        proc_criticos = int((df_proc["N_Dentistas"] == 1).sum())
        proc_top_fat  = df_proc.iloc[0]["Procedimento"] if len(df_proc) else "—"
        proc_mais_qtd = df_proc.sort_values("Total_Orcado", ascending=False).iloc[0]["Procedimento"] if len(df_proc) else "—"

        st.markdown('<p class="section-title">Visao geral dos procedimentos</p>', unsafe_allow_html=True)
        kc1, kc2, kc3, kc4 = st.columns(4)
        kc1.metric("Tipos de Procedimento", str(n_proc_unicos))
        kc2.metric("Procedimentos Criticos", str(proc_criticos),
                   help="Executados por apenas 1 dentista na clinica")
        kc3.metric("Maior Faturamento Aprov.", proc_top_fat[:30] + ("..." if len(proc_top_fat) > 30 else ""))
        kc4.metric("Mais Orcado no Periodo",  proc_mais_qtd[:30] + ("..." if len(proc_mais_qtd) > 30 else ""))

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Tabela principal de procedimentos ─────────────────────────────────
        st.markdown('<p class="section-title">Procedimentos — volume, valor e criticidade</p>',
                    unsafe_allow_html=True)

        linhas_proc = ""
        for i, row in df_proc.iterrows():
            critico = row["N_Dentistas"] == 1
            bg_row  = f'background:{BORDE_CLR};' if critico else ("background:#FBF5F5;" if i % 2 != 0 else "background:#FFFFFF;")
            tx      = BORDE_ESC if critico else CINZA_ESC
            td      = f'padding:10px 14px;border-bottom:1px solid #F0E4E3;vertical-align:middle;color:{tx};'
            critico_badge = (
                f'<span style="background:{BORDE_CLR};color:{BORDE_ESC};padding:2px 8px;'
                f'border:1.5px solid {BORDE};border-radius:10px;font-size:0.72rem;font-weight:700;">CRITICO</span>'
                if critico else ""
            )
            taxa_aprov = round(row["APPROVED_Qtd"] / max(row["Total_Orcado"], 1) * 100, 0)
            cor_barra = VERDE_OK if taxa_aprov >= 60 else (AMARELO if taxa_aprov >= 30 else BORDE)
            barra = (
                f'<div style="background:#eee;border-radius:4px;height:8px;width:100%;margin-top:3px">'
                f'<div style="background:{cor_barra};width:{taxa_aprov}%;height:8px;border-radius:4px"></div>'
                f'</div>'
                f'<span style="font-size:0.75rem;color:{cor_barra}">{int(taxa_aprov)}%</span>'
            )
            linhas_proc += f"""
            <tr style='{bg_row}'>
              <td style='{td}font-weight:600;max-width:240px;word-wrap:break-word'>
                {row['Procedimento']} {critico_badge}
              </td>
              <td style='{td}text-align:center'>{int(row['Total_Orcado'])}</td>
              <td style='{td}text-align:center;color:{VERDE_OK};font-weight:600'>{int(row['APPROVED_Qtd'])}</td>
              <td style='{td}text-align:center;color:{AMARELO}'>{int(row['OPEN_Qtd'])}</td>
              <td style='{td}text-align:center;color:{BORDE}'>{int(row['REJECTED_Qtd'])}</td>
              <td style='{td}'>{barra}</td>
              <td style='{td}text-align:right;font-weight:700;color:{BORDE_ESC}'>{fmt_brl0(row['APPROVED_Valor'])}</td>
              <td style='{td}text-align:right'>{fmt_brl(row['Ticket_Medio_Aprovado'])}</td>
              <td style='{td}text-align:center'>{int(row['N_Dentistas'])}</td>
              <td style='{td}font-size:0.75rem;max-width:200px;word-wrap:break-word'>{row['Dentistas']}</td>
            </tr>"""

        st.markdown(f"""
        <table class="rank-table">
          <thead><tr>
            <th>Procedimento</th>
            <th style='text-align:center'>Total Orc.</th>
            <th style='text-align:center'>Aprovado</th>
            <th style='text-align:center'>Em Aberto</th>
            <th style='text-align:center'>Rejeitado</th>
            <th>Taxa Aprov.</th>
            <th style='text-align:right'>Valor Aprov.</th>
            <th style='text-align:right'>Ticket Medio</th>
            <th style='text-align:center'>Dentistas</th>
            <th>Executantes</th>
          </tr></thead>
          <tbody>{linhas_proc}</tbody>
        </table>
        <br>
        <div style='font-size:0.75rem;color:{CINZA};padding:8px;
                    background:{BORDE_CLR};border-radius:6px;'>
          Linhas em rosa = procedimento <b>CRITICO</b>: apenas 1 dentista na clinica executa esse procedimento.
          A saida desse profissional impacta diretamente a oferta do servico.
        </div>
        """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Cruzamento procedimento x dentista ────────────────────────────────
        st.markdown('<p class="section-title">Quem faz o que — cruzamento procedimento x dentista</p>',
                    unsafe_allow_html=True)

        dentistas_unicos = sorted(df_atv["Profissional_Executante"].dropna().unique())
        filtro_dent_proc = st.selectbox(
            "Filtrar por dentista",
            ["Todos"] + dentistas_unicos,
            key="filtro_proc_dent"
        )

        df_cross = (
            df_atv.groupby(["Procedimento","Profissional_Executante","Status_Orcamento"])
            .agg(Qtd=("Valor_Final","count"), Valor=("Valor_Final","sum"))
            .reset_index()
        )
        pv2 = df_cross.pivot_table(
            index=["Procedimento","Profissional_Executante"],
            columns="Status_Orcamento",
            values=["Qtd","Valor"], fill_value=0
        )
        pv2.columns = [f"{stat}_{m}" for m, stat in pv2.columns]
        pv2 = pv2.reset_index()
        for col in ["APPROVED_Qtd","APPROVED_Valor","OPEN_Qtd","REJECTED_Qtd"]:
            if col not in pv2.columns:
                pv2[col] = 0
        pv2["Total"] = (pv2["APPROVED_Qtd"] + pv2["OPEN_Qtd"] + pv2.get("REJECTED_Qtd", 0)).astype(int)
        pv2 = pv2.sort_values(["Procedimento","APPROVED_Valor"], ascending=[True,False])

        if filtro_dent_proc != "Todos":
            pv2 = pv2[pv2["Profissional_Executante"] == filtro_dent_proc]

        linhas_cross = ""
        prev_proc = None
        for i, (_, row) in enumerate(pv2.iterrows()):
            proc_label = row["Procedimento"] if row["Procedimento"] != prev_proc else ""
            prev_proc  = row["Procedimento"]
            bg_proc    = f'background:#F9F3F2;' if proc_label else f'background:#FFFFFF;'
            td         = f'padding:10px 14px;border-bottom:1px solid #F0E4E3;vertical-align:middle;color:{CINZA_ESC};'
            linhas_cross += f"""
            <tr style='{bg_proc}'>
              <td style='{td}font-weight:{"600" if proc_label else "400"};
                         color:{BORDE_ESC if proc_label else CINZA_ESC}'>{proc_label}</td>
              <td style='{td}'>{row['Profissional_Executante'].title() if pd.notna(row['Profissional_Executante']) else ''}</td>
              <td style='{td}text-align:center'>{int(row['Total'])}</td>
              <td style='{td}text-align:center;color:{VERDE_OK};font-weight:600'>{int(row['APPROVED_Qtd'])}</td>
              <td style='{td}text-align:center;color:{AMARELO}'>{int(row['OPEN_Qtd'])}</td>
              <td style='{td}text-align:center;color:{BORDE}'>{int(row.get("REJECTED_Qtd",0))}</td>
              <td style='{td}text-align:right;font-weight:700;color:{BORDE_ESC}'>{fmt_brl0(row['APPROVED_Valor'])}</td>
            </tr>"""

        st.markdown(f"""
        <table class="rank-table">
          <thead><tr>
            <th>Procedimento</th><th>Dentista</th>
            <th style='text-align:center'>Total</th>
            <th style='text-align:center'>Aprovado</th>
            <th style='text-align:center'>Em Aberto</th>
            <th style='text-align:center'>Rejeitado</th>
            <th style='text-align:right'>Valor Aprovado</th>
          </tr></thead>
          <tbody>{linhas_cross}</tbody>
        </table>
        """, unsafe_allow_html=True)

    # ── TAB 3: GRAFICOS ───────────────────────────────────────────────────────
    with tab3:
        col_a, col_b = st.columns(2)

        with col_a:
            st.markdown('<p class="section-title">Faturamento total por dentista</p>',
                        unsafe_allow_html=True)
            df_plot = df_rank[["Profissional","Faturamento_Total","Fat_Aprovado"]].copy()
            df_plot["Profissional"] = df_plot["Profissional"].str.title()
            df_plot = df_plot.sort_values("Faturamento_Total")

            fig = go.Figure()
            fig.add_bar(
                name="Aprovado",
                y=df_plot["Profissional"],
                x=df_plot["Fat_Aprovado"],
                orientation="h",
                marker_color=VERDE_OK,
                marker_line_width=0,
            )
            fig.add_bar(
                name="Em aberto / Rejeitado",
                y=df_plot["Profissional"],
                x=df_plot["Faturamento_Total"] - df_plot["Fat_Aprovado"],
                orientation="h",
                marker_color=BORDE_MCL,
                marker_line_width=0,
            )
            fig.update_layout(
                barmode="stack",
                plot_bgcolor="#FAFAFA",
                paper_bgcolor="#FAFAFA",
                height=max(280, 42 * n_dent),
                margin=dict(l=0, r=16, t=8, b=8),
                legend=dict(orientation="h", y=-0.12, font=dict(size=11)),
                xaxis=dict(tickprefix="R$ ", showgrid=True,
                           gridcolor="#F0E4E3", tickfont=dict(size=10)),
                yaxis=dict(tickfont=dict(size=10)),
                font=dict(family="Arial"),
            )
            st.plotly_chart(fig, use_container_width=True)

        with col_b:
            st.markdown('<p class="section-title">Ticket medio por dentista</p>',
                        unsafe_allow_html=True)
            df_tick = df_rank[["Profissional","Ticket_Medio_Procedimento"]].copy()
            df_tick["Profissional"] = df_tick["Profissional"].str.title()
            df_tick = df_tick.sort_values("Ticket_Medio_Procedimento")

            fig2 = px.bar(
                df_tick,
                x="Ticket_Medio_Procedimento",
                y="Profissional",
                orientation="h",
                color="Ticket_Medio_Procedimento",
                color_continuous_scale=[[0,"#F5D4D3"],[0.5,"#9B2520"],[1,"#5C1F1D"]],
            )
            fig2.update_traces(marker_line_width=0)
            fig2.update_layout(
                plot_bgcolor="#FAFAFA",
                paper_bgcolor="#FAFAFA",
                height=max(280, 42 * n_dent),
                margin=dict(l=0, r=16, t=8, b=8),
                coloraxis_showscale=False,
                xaxis=dict(tickprefix="R$ ", showgrid=True,
                           gridcolor="#F0E4E3", tickfont=dict(size=10)),
                yaxis=dict(tickfont=dict(size=10)),
                font=dict(family="Arial"),
            )
            st.plotly_chart(fig2, use_container_width=True)

        st.markdown('<p class="section-title">Volume de procedimentos por dentista</p>',
                    unsafe_allow_html=True)
        df_proc = df_rank[["Profissional","Qtd_Procedimentos","Qtd_Tratamentos"]].copy()
        df_proc["Profissional"] = df_proc["Profissional"].str.title()

        fig3 = go.Figure()
        fig3.add_bar(name="Procedimentos", x=df_proc["Profissional"],
                     y=df_proc["Qtd_Procedimentos"], marker_color=BORDE)
        fig3.add_bar(name="Tratamentos",   x=df_proc["Profissional"],
                     y=df_proc["Qtd_Tratamentos"],   marker_color=BORDE_MCL)
        fig3.update_layout(
            barmode="group",
            plot_bgcolor="#FAFAFA",
            paper_bgcolor="#FAFAFA",
            height=320,
            margin=dict(l=0, r=16, t=8, b=8),
            legend=dict(orientation="h", y=1.08),
            yaxis=dict(showgrid=True, gridcolor="#F0E4E3"),
            font=dict(family="Arial"),
        )
        st.plotly_chart(fig3, use_container_width=True)

        # Top procedimentos
        st.markdown('<p class="section-title">Procedimentos mais frequentes (todos os dentistas)</p>',
                    unsafe_allow_html=True)
        top_proc = (
            df_det[df_det["Deletado"] == False]["Procedimento"]
            .value_counts()
            .head(10)
            .reset_index()
        )
        top_proc.columns = ["Procedimento","Quantidade"]
        fig4 = px.bar(top_proc.sort_values("Quantidade"),
                      x="Quantidade", y="Procedimento", orientation="h",
                      color_discrete_sequence=[BORDE])
        fig4.update_layout(
            plot_bgcolor="#FAFAFA", paper_bgcolor="#FAFAFA",
            height=320, margin=dict(l=0, r=16, t=8, b=8),
            yaxis=dict(tickfont=dict(size=10)),
            font=dict(family="Arial"),
        )
        st.plotly_chart(fig4, use_container_width=True)

    # ── TAB 4: STATUS ─────────────────────────────────────────────────────────
    with tab4:
        st.markdown('<p class="section-title">Distribuicao por status de orcamento</p>',
                    unsafe_allow_html=True)

        col_p, col_t = st.columns([1, 1])

        with col_p:
            stat_colors_plot = {
                "APPROVED": VERDE_OK,
                "OPEN":     "#E8A020",
                "REJECTED": BORDE,
            }
            cores = [stat_colors_plot.get(s, CINZA) for s in df_stat["Status_Orcamento"]]
            labels_pt = {"APPROVED":"Aprovado","OPEN":"Em Aberto","REJECTED":"Rejeitado"}
            labels_show = [labels_pt.get(s,s) for s in df_stat["Status_Orcamento"]]

            fig5 = go.Figure(go.Pie(
                labels=labels_show,
                values=df_stat["Valor_Total"],
                marker=dict(colors=cores, line=dict(color="#fff", width=2)),
                hole=0.45,
                textinfo="label+percent",
                textfont=dict(size=12, family="Arial"),
            ))
            fig5.update_layout(
                paper_bgcolor="#FAFAFA",
                height=300,
                margin=dict(l=0, r=0, t=8, b=8),
                showlegend=False,
                font=dict(family="Arial"),
            )
            st.plotly_chart(fig5, use_container_width=True)

        with col_t:
            total_val = int(df_stat["Valor_Total"].sum())
            for _, srow in df_stat.iterrows():
                status   = srow["Status_Orcamento"]
                label_pt = labels_pt.get(status, status)
                pct      = round(srow["Valor_Total"] / max(total_val,1) * 100, 1)
                cor_bg   = {"APPROVED":"#E8F5E9","OPEN":"#FFF9E6","REJECTED":"#FDECEA"}.get(status,"#fff")
                cor_tx   = stat_colors_plot.get(status, CINZA)
                st.markdown(f"""
                <div style='background:{cor_bg};border-left:5px solid {cor_tx};
                            border-radius:8px;padding:14px 18px;margin-bottom:10px'>
                  <div style='font-weight:700;font-size:0.8rem;
                              text-transform:uppercase;letter-spacing:0.06em;
                              color:{cor_tx}'>{label_pt}</div>
                  <div style='font-size:1.5rem;font-weight:800;color:{cor_tx}'>
                    {fmt_brl0(srow["Valor_Total"])}
                  </div>
                  <div style='font-size:0.8rem;color:#6C757D'>
                    {int(srow["Qtd_Procedimentos"])} procedimentos &nbsp;·&nbsp; {pct}% do total
                  </div>
                </div>
                """, unsafe_allow_html=True)

    # ── TAB 5: DETALHADO ─────────────────────────────────────────────────────
    with tab5:
        st.markdown('<p class="section-title">Filtros</p>', unsafe_allow_html=True)

        fc1, fc2, fc3 = st.columns(3)
        dentistas_lista = ["Todos"] + sorted(
            df_det["Profissional_Executante"].dropna().unique().tolist()
        )
        status_lista = ["Todos"] + sorted(df_det["Status_Orcamento"].dropna().unique().tolist())

        filtro_dent   = fc1.selectbox("Dentista",         dentistas_lista)
        filtro_status = fc2.selectbox("Status",           status_lista)
        filtro_proc   = fc3.text_input("Procedimento (busca livre)", "")

        df_filtrado = df_det.copy()
        if filtro_dent   != "Todos":
            df_filtrado = df_filtrado[df_filtrado["Profissional_Executante"] == filtro_dent]
        if filtro_status != "Todos":
            df_filtrado = df_filtrado[df_filtrado["Status_Orcamento"] == filtro_status]
        if filtro_proc:
            df_filtrado = df_filtrado[
                df_filtrado["Procedimento"].fillna("").str.contains(filtro_proc, case=False)
            ]

        st.markdown(f'<p class="section-title">{len(df_filtrado)} procedimentos encontrados</p>',
                    unsafe_allow_html=True)

        # Exibir com chips de status
        df_exibir = df_filtrado[[
            "Data","Paciente","Status_Orcamento",
            "Profissional_Executante","Procedimento",
            "Dente","Valor_Tabela","Valor_Final","Deletado"
        ]].copy()
        df_exibir["Data"] = pd.to_datetime(df_exibir["Data"], errors="coerce").dt.strftime("%d/%m/%Y")
        df_exibir = df_exibir.rename(columns={
            "Data":"Data","Paciente":"Paciente","Status_Orcamento":"Status",
            "Profissional_Executante":"Dentista","Procedimento":"Procedimento",
            "Dente":"Dente","Valor_Tabela":"Vlr. Tabela","Valor_Final":"Vlr. Final",
            "Deletado":"Deletado"
        })

        st.dataframe(
            df_exibir,
            use_container_width=True,
            height=400,
            column_config={
                "Vlr. Tabela": st.column_config.NumberColumn(format="R$ %.2f"),
                "Vlr. Final":  st.column_config.NumberColumn(format="R$ %.2f"),
                "Status": st.column_config.TextColumn(),
            }
        )

    # ── DOWNLOAD ──────────────────────────────────────────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<p class="section-title">Exportar planilha</p>', unsafe_allow_html=True)

    dl1, dl2, dl3 = st.columns([2, 1, 3])
    with dl1:
        with st.spinner("Preparando Excel..."):
            xlsx_bytes = gerar_excel_bytes(
                d["df_det"], d["df_orc"], d["df_exec"], d["df_stat"],
                d["d_ini"], d["d_fim"]
            )
        st.download_button(
            label="Baixar planilha Excel (UCJ)",
            data=xlsx_bytes,
            file_name=f"UCJ_faturamento_{d['d_ini']}_a_{d['d_fim']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

else:
    # Estado inicial — sem dados
    st.markdown(f"""
    <div style='text-align:center; padding:60px 20px; color:{CINZA}'>
        <div style='font-size:3rem; margin-bottom:16px; color:{BORDE_MCL}'>&#9679;</div>
        <div style='font-size:1.1rem; font-weight:600; color:{BORDE_ESC}; margin-bottom:8px'>
            Nenhum dado carregado
        </div>
        <div style='font-size:0.88rem'>
            Preencha as credenciais e o periodo na barra lateral<br>
            e clique em <b>Buscar dados</b> para iniciar a analise.
        </div>
    </div>
    """, unsafe_allow_html=True)
