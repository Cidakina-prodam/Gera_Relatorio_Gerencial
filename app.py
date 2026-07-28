"""
Conferência GDS-1 — Horas, Receita, Colaboradores e Despesa
=============================================================
App de apoio para conferir rapidamente os dados do Portal GDS-1 sem precisar
montar o relatório completo em Word. Aplica a correção de horas por núcleo
(prioriza o detalhamento por núcleo sobre o campo-resumo do contrato).

Como rodar localmente:
    pip install -r requirements.txt
    streamlit run app.py

Como publicar (gratuito):
    1. Suba esta pasta para um repositório no GitHub.
    2. Acesse https://share.streamlit.io, conecte sua conta GitHub.
    3. Escolha o repositório e o arquivo app.py. Deploy.
"""

import io
import json
from collections import Counter, defaultdict

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Conferência GDS-1", page_icon="📊", layout="wide")

# ----------------------------------------------------------------------
# Helpers de cálculo (mesma lógica corrigida usada no relatório em Word)
# ----------------------------------------------------------------------

def num(v):
    if v in ("", None):
        return 0.0
    return float(v)


def horas_contrato_mes(lancamentos_mes: dict, contrato_id: str):
    """Prioriza o detalhamento por núcleo ('{id}_{NUCLEO}') sobre o campo-resumo ('{id}')."""
    prefixo = f"{contrato_id}_"
    splits = {k: v for k, v in lancamentos_mes.items() if k.startswith(prefixo)}
    if splits:
        posH = sum(num(v.get("posH")) for v in splits.values())
        gfpH = sum(num(v.get("gfpH")) for v in splits.values())
    else:
        bare = lancamentos_mes.get(contrato_id, {})
        posH = num(bare.get("posH"))
        gfpH = num(bare.get("gfpH"))
    return posH, gfpH


def nucleo_do_split(chave: str) -> str:
    return chave.split("_", 1)[1]


@st.cache_data(show_spinner=False)
def processar_portal(backup_bytes: bytes):
    backup = json.loads(backup_bytes)
    contratos = {str(c["id"]): c for c in backup["contratos"]}
    meses = sorted(m for m in backup["lancamentos"].keys())

    linhas = []  # detalhe por contrato/mes
    for mes in meses:
        lan = backup["lancamentos"][mes].get("contratos", {})
        for cid, c in contratos.items():
            posH, gfpH = horas_contrato_mes(lan, cid)
            if posH == 0 and gfpH == 0:
                continue
            vlrHora = c.get("vlrHora", 0)
            linhas.append(
                dict(
                    mes=mes,
                    contrato_id=cid,
                    secretaria=c.get("secretaria"),
                    contrato=c.get("contrato"),
                    nucleo=c.get("nucleo"),
                    status=c.get("status"),
                    vlrHora=vlrHora,
                    posH=posH,
                    gfpH=gfpH,
                    posV=posH * vlrHora,
                    gfpV=gfpH * vlrHora,
                )
            )
    df = pd.DataFrame(linhas)

    # núcleo mensal usando o núcleo REAL de cada split (não o núcleo primário do contrato)
    nucleo_rows = []
    for mes in meses:
        lan = backup["lancamentos"][mes].get("contratos", {})
        tmp = defaultdict(float)
        for cid, c in contratos.items():
            splits = {k: v for k, v in lan.items() if k.startswith(cid + "_")}
            bare = lan.get(cid, {})
            if splits:
                for k, v in splits.items():
                    tmp[nucleo_do_split(k)] += num(v.get("posH"))
            else:
                tmp[c["nucleo"]] += num(bare.get("posH"))
        for n, h in tmp.items():
            if h:
                nucleo_rows.append(dict(mes=mes, nucleo=n, posH=h))
    df_nucleo = pd.DataFrame(nucleo_rows)

    return backup, contratos, meses, df, df_nucleo


@st.cache_data(show_spinner=False)
def processar_colaboradores(colab_bytes: bytes):
    d = json.loads(colab_bytes)
    colabs = d["colaboradores"]
    ativos = [c for c in colabs if not c.get("dataSaida")]
    return colabs, ativos


@st.cache_data(show_spinner=False)
def processar_consumo_csv(csv_bytes: bytes):
    df = pd.read_csv(io.BytesIO(csv_bytes), sep=";", encoding="utf-8-sig")
    return df


def fmt_h(v):
    return f"{v:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_money(v):
    return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# ----------------------------------------------------------------------
# Interface
# ----------------------------------------------------------------------

st.title("📊 Conferência GDS-1")
st.caption(
    "Sobe os arquivos exportados do Portal GDS-1 (e opcionalmente Colaboradores / OS / Consumo) "
    "e confere na hora as horas, receita, equipe e despesa — já com a correção de horas por núcleo aplicada."
)

with st.sidebar:
    st.header("Arquivos")
    f_portal = st.file_uploader("backup_portal_gds_*.json", type="json", key="portal")
    f_colab = st.file_uploader("backup_colaboradores_*.json (opcional)", type="json", key="colab")
    f_consumo = st.file_uploader("consumo_os_prodam_*.csv (opcional)", type="csv", key="consumo")
    st.markdown("---")
    st.caption(
        "Correção aplicada: horas são somadas a partir do detalhamento por núcleo "
        "(`contrato_NUCLEO`) sempre que existir, em vez do campo-resumo do contrato "
        "— que pode ficar desatualizado."
    )

if not f_portal:
    st.info("⬅️ Suba pelo menos o `backup_portal_gds_*.json` na barra lateral para começar.")
    st.stop()

backup, contratos, meses, df, df_nucleo = processar_portal(f_portal.getvalue())

meses_label = {m: pd.to_datetime(m + "-01").strftime("%b/%y").capitalize() for m in meses}

# ---- KPIs ----
tot_posH = df["posH"].sum()
tot_gfpH = df["gfpH"].sum()
tot_posV = df["posV"].sum()
tot_gfpV = df["gfpV"].sum()

c1, c2, c3, c4 = st.columns(4)
c1.metric("Horas GDS-1 (total)", fmt_h(tot_posH))
c2.metric("Receita GDS-1", fmt_money(tot_posV))
c3.metric("Horas GFP (total)", fmt_h(tot_gfpH))
c4.metric("Receita GFP", fmt_money(tot_gfpV))

st.markdown("---")

# ---- Resumo mensal ----
st.subheader("Resumo mensal — GDS-1 x GFP")
resumo = (
    df.groupby("mes")[["posH", "gfpH", "posV", "gfpV"]]
    .sum()
    .reindex(meses)
    .fillna(0)
)
resumo_fmt = pd.DataFrame(
    {
        "Mês": [meses_label[m] for m in resumo.index],
        "Horas GDS-1": resumo["posH"].map(fmt_h),
        "Receita GDS-1": resumo["posV"].map(fmt_money),
        "Horas GFP": resumo["gfpH"].map(fmt_h),
        "Receita GFP": resumo["gfpV"].map(fmt_money),
    }
)
st.dataframe(resumo_fmt, hide_index=True, use_container_width=True)

# ---- Núcleo mensal ----
st.subheader("Horas GDS-1 por núcleo")
if not df_nucleo.empty:
    piv = df_nucleo.pivot_table(index="nucleo", columns="mes", values="posH", aggfunc="sum").reindex(
        columns=meses
    ).fillna(0)
    piv["Total"] = piv.sum(axis=1)
    piv_fmt = piv.copy()
    for c in piv_fmt.columns:
        piv_fmt[c] = piv_fmt[c].map(fmt_h)
    piv_fmt.columns = [meses_label.get(c, c) for c in piv.columns]
    st.dataframe(piv_fmt, use_container_width=True)
else:
    st.caption("Sem dados de núcleo para o período.")

# ---- Saldo por contrato ----
st.subheader("Saldo de horas por contrato")
saldo_rows = []
for cid, c in contratos.items():
    acum = df.loc[df["contrato_id"] == cid, "posH"].sum()
    horasTotal = c.get("horasTotal", 0)
    if horasTotal == 0 and acum == 0:
        continue
    saldo = horasTotal - acum
    pct = acum / horasTotal * 100 if horasTotal else 0
    saldo_rows.append(
        dict(
            Contrato=f"{c.get('secretaria')} — {c.get('contrato')}",
            Status=c.get("status"),
            **{"Horas Totais": fmt_h(horasTotal), "Acumulado": fmt_h(acum), "Saldo": fmt_h(saldo), "% Executado": f"{pct:.1f}%".replace(".", ",")}
        )
    )
st.dataframe(pd.DataFrame(saldo_rows), hide_index=True, use_container_width=True)
st.caption(
    "⚠️ Contratos renovados/recriados no Portal aparecem como linhas separadas "
    "(um id 'Vencido' antigo e um id 'Ativo' novo). Some as duas linhas para o total real do contrato."
)

st.markdown("---")

# ---- Colaboradores ----
if f_colab:
    colabs, ativos = processar_colaboradores(f_colab.getvalue())
    st.subheader("Colaboradores")
    c1, c2, c3 = st.columns(3)
    c1.metric("Total ativos", len(ativos))
    c2.metric("Prodam", sum(1 for c in ativos if c.get("tipo") == "Prodam"))
    c3.metric("Terceiros", sum(1 for c in ativos if c.get("tipo") == "Terceiro"))

    por_nucleo = Counter(c.get("nucleo") for c in ativos)
    st.dataframe(
        pd.DataFrame(sorted(por_nucleo.items()), columns=["Núcleo", "Colaboradores"]),
        hide_index=True,
    )

    st.markdown("---")

# ---- Despesa terceiros (consumo CSV) ----
if f_consumo:
    st.subheader("Consumo / OS (bruto)")
    df_consumo = processar_consumo_csv(f_consumo.getvalue())
    st.dataframe(df_consumo, use_container_width=True)
    st.caption("Exibido conforme o CSV enviado — cálculo de despesa por fornecedor ainda não automatizado aqui.")

st.markdown("---")

# ---- Download dos dados calculados ----
st.subheader("Exportar")
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as writer:
    resumo_fmt.to_excel(writer, sheet_name="Resumo mensal", index=False)
    if not df_nucleo.empty:
        piv_fmt.to_excel(writer, sheet_name="Por núcleo")
    pd.DataFrame(saldo_rows).to_excel(writer, sheet_name="Saldo por contrato", index=False)
    df.to_excel(writer, sheet_name="Detalhe bruto", index=False)
st.download_button(
    "⬇️ Baixar Excel com os dados corrigidos",
    data=buf.getvalue(),
    file_name="conferencia_gds1.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
