import io
from datetime import date

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st


st.set_page_config(
    page_title="3-Statement Financial Model | The Mountain Path Academy",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    :root { --navy:#082B4C; --blue:#0B5E8E; --gold:#D6A43B; --cream:#F7F4EC; }
    .stApp { background: linear-gradient(180deg,#ffffff 0%,#f7fafc 100%); }
    h1,h2,h3 { color:var(--navy); }
    [data-testid="stSidebar"] { background:var(--navy); }
    [data-testid="stSidebar"] * { color:#fff; }
    [data-testid="stSidebar"] input { color:#111 !important; }
    [data-testid="stSidebar"] [data-baseweb="select"] * { color:#111 !important; }
    .hero {padding:1.35rem 1.6rem;border-radius:16px;background:linear-gradient(120deg,#082B4C,#0B5E8E);color:white;margin-bottom:1rem;box-shadow:0 8px 24px #082b4c22}
    .hero h1 {color:white;margin:0;font-size:2rem}.hero p{margin:.35rem 0 0;color:#e8f4fa}
    .tag {display:inline-block;background:#D6A43B;color:#082B4C;padding:.25rem .7rem;border-radius:20px;font-weight:700;margin-top:.7rem}
    .note {background:#eef7fb;border-left:5px solid #0B5E8E;padding:.8rem 1rem;border-radius:8px;margin:.5rem 0 1rem}
    .check-ok {background:#e8f7ee;border:1px solid #4caf70;padding:.7rem 1rem;border-radius:8px;color:#185b31;font-weight:700}
    .check-bad {background:#fff0f0;border:1px solid #dc5a5a;padding:.7rem 1rem;border-radius:8px;color:#8d2020;font-weight:700}
    div[data-testid="stMetric"] {background:white;border:1px solid #dce6ec;padding:.8rem;border-radius:12px;box-shadow:0 3px 10px #082b4c0d}
    </style>
    """,
    unsafe_allow_html=True,
)


def money(x):
    return f"₹{x:,.1f} Cr"


def statement_table(rows, years, percentage_rows=None):
    percentage_rows = percentage_rows or []
    df = pd.DataFrame(rows, index=years).T
    styled = df.style.format("{:,.1f}")
    bold = [r for r in df.index if r.startswith(("Gross", "EBIT", "Profit", "Total", "Net Cash", "Closing"))]
    styled = styled.set_properties(subset=pd.IndexSlice[bold, :], **{"font-weight": "bold", "background-color": "#eaf3f8"})
    return df, styled


def excel_download(statements):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book
        title_fmt = wb.add_format({"bold": True, "font_size": 16, "font_color": "white", "bg_color": "#082B4C", "align": "center"})
        header_fmt = wb.add_format({"bold": True, "font_color": "white", "bg_color": "#0B5E8E", "align": "center", "border": 1})
        input_fmt = wb.add_format({"font_color": "#0000FF", "bg_color": "#FFF7DD", "num_format": "0.0", "border": 1})
        input_pct_fmt = wb.add_format({"font_color": "#0000FF", "bg_color": "#FFF7DD", "num_format": "0.0%", "border": 1})
        formula_fmt = wb.add_format({"font_color": "#008000", "num_format": "#,##0.0;[Red](#,##0.0);-", "border": 1})
        total_fmt = wb.add_format({"bold": True, "bg_color": "#EAF3F8", "num_format": "#,##0.0;[Red](#,##0.0);-", "top": 1, "bottom": 1})
        label_fmt = wb.add_format({"border": 1})
        section_fmt = wb.add_format({"bold": True, "font_color": "#082B4C", "bg_color": "#D9EAF2"})
        note_fmt = wb.add_format({"italic": True, "font_color": "#555555"})
        check_fmt = wb.add_format({"bold": True, "num_format": "0.00", "border": 1})

        # Assumptions sheet: all editable model drivers are in blue.
        ws = wb.add_worksheet("Assumptions")
        writer.sheets["Assumptions"] = ws
        ws.merge_range("A1:D1", f"{company} — Dynamic Model Assumptions", title_fmt)
        ws.write("A2", "Assumption", header_fmt)
        for col, year in enumerate(years, 1):
            ws.write(1, col, year, header_fmt)
        assumption_rows = [
            ("Prior-year revenue", [base_revenue, base_revenue, base_revenue], False),
            ("Revenue growth", growth, True), ("COGS / Revenue", cogs_pct, True),
            ("Operating expenses / Revenue", opex_pct, True), ("Tax rate", [tax_rate] * 3, True),
            ("Receivable days (DSO)", dso, False), ("Inventory days (DIO)", dio, False),
            ("Payable days (DPO)", dpo, False), ("Other current assets / Revenue", [oca_pct] * 3, True),
            ("Other current liabilities / Revenue", [ocl_pct] * 3, True), ("Capex / Revenue", capex_pct, True),
            ("Depreciation / opening PPE", [dep_pct] * 3, True),
            ("Interest rate on average debt", [interest_rate] * 3, True),
            ("New borrowing / (repayment)", debt_change, False), ("Dividend payout / PAT", [dividend_pct] * 3, True),
        ]
        for r, (label, vals, is_pct) in enumerate(assumption_rows, 2):
            ws.write(r, 0, label, label_fmt)
            for c, value in enumerate(vals, 1):
                ws.write_number(r, c, value, input_pct_fmt if is_pct else input_fmt)
        ws.write("A19", "Opening Balance-Sheet Inputs", section_fmt)
        opening_items = [
            ("Opening cash", opening_cash), ("Opening debt", opening_debt), ("Opening net PPE", opening_ppe),
            ("Opening shareholders' equity", opening_equity), ("Opening receivables", opening_ar),
            ("Opening inventory", opening_inventory), ("Opening payables", opening_ap),
            ("Opening other current assets", opening_oca), ("Opening other current liabilities", opening_ocl),
        ]
        for r, (label, value) in enumerate(opening_items, 19):
            ws.write(r, 0, label, label_fmt)
            ws.write_number(r, 1, value, input_fmt)
        ws.write("A30", "Blue font = editable input | Green font = linked formula", note_fmt)
        ws.set_column("A:A", 38); ws.set_column("B:D", 17); ws.freeze_panes(2, 1)

        def create_statement(sheet_name, title, labels, formulas, cached_values, total_labels):
            sh = wb.add_worksheet(sheet_name)
            writer.sheets[sheet_name] = sh
            sh.merge_range("A1:D1", title, title_fmt)
            sh.write("A2", "Particulars", header_fmt)
            for col, year in enumerate(years, 1):
                sh.write(1, col, year, header_fmt)
            for r, label in enumerate(labels, 2):
                sh.write(r, 0, label, total_fmt if label in total_labels else label_fmt)
                for c in range(3):
                    fmt = total_fmt if label in total_labels else formula_fmt
                    sh.write_formula(r, c + 1, formulas[r - 2][c], fmt, float(cached_values[r - 2][c]))
            sh.set_column("A:A", 38); sh.set_column("B:D", 17); sh.freeze_panes(2, 1)
            return sh

        # Income Statement formulas.
        is_labels = list(income_rows.keys())
        is_values = [income_rows[x] for x in is_labels]
        is_formulas = []
        for i, col in enumerate(["B", "C", "D"]):
            prev_rev = "Assumptions!$B$3" if i == 0 else f"{'BCD'[i-1]}3"
            opening_debt_ref = "Assumptions!$B$21" if i == 0 else f"'Balance Sheet'!{'BCD'[i-1]}11"
            opening_ppe_ref = "Assumptions!$B$22" if i == 0 else f"'Balance Sheet'!{'BCD'[i-1]}7"
            is_formulas.append([
                f"={prev_rev}*(1+Assumptions!{col}4)", f"=-{col}3*Assumptions!{col}5",
                f"=SUM({col}3:{col}4)", f"=-{col}3*Assumptions!{col}6", f"=SUM({col}5:{col}6)",
                f"=-{opening_ppe_ref}*Assumptions!{col}14",
                f"=SUM({col}7:{col}8)", f"=-AVERAGE({opening_debt_ref},'Balance Sheet'!{col}11)*Assumptions!{col}15",
                f"=SUM({col}9:{col}10)", f"=-MAX(0,{col}11*Assumptions!{col}7)", f"=SUM({col}11:{col}12)",
            ])
        is_formulas = [list(x) for x in zip(*is_formulas)]
        create_statement("Income Statement", f"{company} — Forecast Income Statement", is_labels, is_formulas, is_values,
                         {"Gross Profit", "EBITDA", "EBIT", "Profit Before Tax", "Profit After Tax"})

        # Balance Sheet formulas.
        bs_labels = list(bs_rows.keys())
        bs_values = [bs_rows[x] for x in bs_labels]
        bs_formulas = []
        for i, col in enumerate(["B", "C", "D"]):
            prev_ppe = "Assumptions!$B$22" if i == 0 else f"{'BCD'[i-1]}7"
            prev_debt = "Assumptions!$B$21" if i == 0 else f"{'BCD'[i-1]}11"
            prev_equity = "Assumptions!$B$23" if i == 0 else f"{'BCD'[i-1]}12"
            bs_formulas.append([
                f"='Cash Flow'!{col}18", f"='Income Statement'!{col}3*Assumptions!{col}8/365",
                f"=-'Income Statement'!{col}4*Assumptions!{col}9/365", f"='Income Statement'!{col}3*Assumptions!{col}11",
                f"={prev_ppe}+'Income Statement'!{col}3*Assumptions!{col}13+'Income Statement'!{col}8",
                f"=SUM({col}3:{col}7)", f"=-'Income Statement'!{col}4*Assumptions!{col}10/365",
                f"='Income Statement'!{col}3*Assumptions!{col}12", f"=MAX(0,{prev_debt}+Assumptions!{col}16)",
                f"={prev_equity}+'Income Statement'!{col}13-MAX(0,'Income Statement'!{col}13*Assumptions!{col}17)",
                f"=SUM({col}9:{col}12)", f"={col}8-{col}13",
            ])
        bs_formulas = [list(x) for x in zip(*bs_formulas)]
        bs_sheet = create_statement("Balance Sheet", f"{company} — Forecast Balance Sheet", bs_labels, bs_formulas, bs_values,
                                    {"Total Assets", "Total Liabilities & Equity", "Balance Check"})
        bs_sheet.conditional_format("B14:D14", {"type": "cell", "criteria": "!=", "value": 0, "format": wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})})

        # Cash Flow formulas.
        cf_labels = list(cf_rows.keys())
        cf_values = [cf_rows[x] for x in cf_labels]
        cf_formulas = []
        for i, col in enumerate(["B", "C", "D"]):
            pcol = "BCD"[i - 1] if i > 0 else None
            open_ar_ref = "Assumptions!$B$24" if i == 0 else f"'Balance Sheet'!{pcol}4"
            open_inv_ref = "Assumptions!$B$25" if i == 0 else f"'Balance Sheet'!{pcol}5"
            open_oca_ref = "Assumptions!$B$27" if i == 0 else f"'Balance Sheet'!{pcol}6"
            open_ap_ref = "Assumptions!$B$26" if i == 0 else f"'Balance Sheet'!{pcol}9"
            open_ocl_ref = "Assumptions!$B$28" if i == 0 else f"'Balance Sheet'!{pcol}10"
            open_debt_ref = "Assumptions!$B$21" if i == 0 else f"'Balance Sheet'!{pcol}11"
            open_cash_ref = "Assumptions!$B$20" if i == 0 else f"{pcol}18"
            cf_formulas.append([
                f"='Income Statement'!{col}13", f"=-'Income Statement'!{col}8",
                f"=-('Balance Sheet'!{col}4-{open_ar_ref})", f"=-('Balance Sheet'!{col}5-{open_inv_ref})",
                f"=-('Balance Sheet'!{col}6-{open_oca_ref})", f"='Balance Sheet'!{col}9-{open_ap_ref}",
                f"='Balance Sheet'!{col}10-{open_ocl_ref}", f"=SUM({col}3:{col}9)",
                f"=-'Income Statement'!{col}3*Assumptions!{col}13", f"={col}11",
                f"='Balance Sheet'!{col}11-{open_debt_ref}", f"=-MAX(0,'Income Statement'!{col}13*Assumptions!{col}17)",
                f"=SUM({col}13:{col}14)", f"=SUM({col}10,{col}12,{col}15)", f"={open_cash_ref}", f"=SUM({col}16:{col}17)",
            ])
        cf_formulas = [list(x) for x in zip(*cf_formulas)]
        create_statement("Cash Flow", f"{company} — Forecast Cash Flow Statement", cf_labels, cf_formulas, cf_values,
                         {"Net Cash from Operating Activities", "Net Cash from Investing Activities", "Net Cash from Financing Activities", "Net Change in Cash", "Closing Cash"})

        ws.activate()
    return output.getvalue()


st.sidebar.markdown("## ⛰️ Mountain Path Academy")
st.sidebar.caption("Dynamic 3-Statement Model")
st.sidebar.markdown("---")
st.sidebar.info("Change the assumptions below, then click **Apply Inputs & Refresh Model**.")

assumptions_form = st.sidebar.form("financial_model_assumptions")

with assumptions_form.expander("1 · Model setup", expanded=True):
    company = st.text_input("Company name", "MPA Manufacturing Ltd.")
    start_year = st.number_input("First forecast year", 2024, 2035, date.today().year, 1)
    unit = st.selectbox("Display unit", ["₹ Crore", "₹ Million"], index=0)
    opening_cash = st.number_input("Opening cash", 0.0, 10000.0, 40.0, 5.0)
    opening_debt = st.number_input("Opening debt", 0.0, 10000.0, 150.0, 5.0)
    opening_ppe = st.number_input("Opening net PPE", 0.0, 10000.0, 320.0, 10.0)
    opening_equity = st.number_input("Opening shareholders' equity", 0.0, 10000.0, 320.0, 10.0)

years = [f"FY{int(start_year)+i}" for i in range(3)]

with assumptions_form.expander("2 · Revenue & operating assumptions", expanded=True):
    base_revenue = st.number_input("Prior-year revenue", 1.0, 100000.0, 1000.0, 25.0)
    growth = [st.slider(f"Revenue growth — {y}", -20.0, 50.0, v, 0.5) / 100 for y, v in zip(years, [10.0, 12.0, 11.0])]
    cogs_pct = [st.slider(f"COGS / Revenue — {y}", 20.0, 90.0, 60.0, 0.5) / 100 for y in years]
    opex_pct = [st.slider(f"Operating expenses / Revenue — {y}", 5.0, 50.0, 18.0, 0.5) / 100 for y in years]
    tax_rate = st.slider("Corporate tax rate", 0.0, 50.0, 25.0, 0.5) / 100

with assumptions_form.expander("3 · Working capital assumptions"):
    dso = [st.number_input(f"Receivable days — {y}", 0, 180, 45, 1) for y in years]
    dio = [st.number_input(f"Inventory days — {y}", 0, 365, 60, 1) for y in years]
    dpo = [st.number_input(f"Payable days — {y}", 0, 180, 40, 1) for y in years]
    oca_pct = st.slider("Other current assets / Revenue", 0.0, 20.0, 4.0, 0.5) / 100
    ocl_pct = st.slider("Other current liabilities / Revenue", 0.0, 20.0, 5.0, 0.5) / 100

with assumptions_form.expander("4 · Investment, financing & payout"):
    capex_pct = [st.slider(f"Capex / Revenue — {y}", 0.0, 30.0, 8.0, 0.5) / 100 for y in years]
    dep_pct = st.slider("Depreciation / opening PPE", 0.0, 30.0, 10.0, 0.5) / 100
    interest_rate = st.slider("Interest rate on average debt", 0.0, 25.0, 8.0, 0.25) / 100
    debt_change = [st.number_input(f"New borrowing / (repayment) — {y}", -500.0, 500.0, v, 5.0) for y, v in zip(years, [20.0, -10.0, -20.0])]
    dividend_pct = st.slider("Dividend payout / PAT", 0.0, 100.0, 25.0, 1.0) / 100

with assumptions_form.expander("5 · Opening working-capital balances"):
    opening_ar = st.number_input("Opening receivables", 0.0, 10000.0, 120.0, 5.0)
    opening_inventory = st.number_input("Opening inventory", 0.0, 10000.0, 100.0, 5.0)
    opening_ap = st.number_input("Opening payables", 0.0, 10000.0, 80.0, 5.0)
    opening_oca = st.number_input("Opening other current assets", 0.0, 10000.0, 40.0, 5.0)
    opening_ocl = st.number_input("Opening other current liabilities", 0.0, 10000.0, 70.0, 5.0)

submitted = assumptions_form.form_submit_button(
    "🔄 Apply Inputs & Refresh Model",
    type="primary",
    use_container_width=True,
)

if submitted:
    st.toast("Financial statements refreshed using the revised assumptions.", icon="✅")

# Model engine
revenue=[]; cogs=[]; gross=[]; opex=[]; ebitda=[]; dep=[]; ebit=[]; debt=[]; interest=[]; pbt=[]; tax=[]; pat=[]; dividends=[]
ar=[]; inventory=[]; oca=[]; ap=[]; ocl=[]; capex=[]; ppe=[]; equity=[]; cash=[]
cfo=[]; cfi=[]; cff=[]; net_cash_change=[]; checks=[]

prev_rev=base_revenue; prev_ppe=opening_ppe; prev_debt=opening_debt; prev_equity=opening_equity; prev_cash=opening_cash
prev_ar=opening_ar; prev_inv=opening_inventory; prev_oca=opening_oca; prev_ap=opening_ap; prev_ocl=opening_ocl

for i in range(3):
    rev=prev_rev*(1+growth[i]); cg=rev*cogs_pct[i]; gp=rev-cg; op=rev*opex_pct[i]; eda=gp-op
    dv=prev_ppe*dep_pct; ei=eda-dv
    end_debt=max(0.0,prev_debt+debt_change[i]); intr=(prev_debt+end_debt)/2*interest_rate
    before_tax=ei-intr; tx=max(0.0,before_tax*tax_rate); profit=before_tax-tx; div=max(0.0,profit*dividend_pct)
    end_ar=rev*dso[i]/365; end_inv=cg*dio[i]/365; end_oca=rev*oca_pct
    end_ap=cg*dpo[i]/365; end_ocl=rev*ocl_pct
    cx=rev*capex_pct[i]; end_ppe=prev_ppe+cx-dv; end_equity=prev_equity+profit-div
    operating_cf=profit+dv-(end_ar-prev_ar)-(end_inv-prev_inv)-(end_oca-prev_oca)+(end_ap-prev_ap)+(end_ocl-prev_ocl)
    investing_cf=-cx; financing_cf=(end_debt-prev_debt)-div; change=operating_cf+investing_cf+financing_cf; end_cash=prev_cash+change
    liabilities_equity=end_ap+end_ocl+end_debt+end_equity
    total_assets=end_cash+end_ar+end_inv+end_oca+end_ppe
    check=total_assets-liabilities_equity
    revenue.append(rev); cogs.append(cg); gross.append(gp); opex.append(op); ebitda.append(eda); dep.append(dv); ebit.append(ei)
    debt.append(end_debt); interest.append(intr); pbt.append(before_tax); tax.append(tx); pat.append(profit); dividends.append(div)
    ar.append(end_ar); inventory.append(end_inv); oca.append(end_oca); ap.append(end_ap); ocl.append(end_ocl); capex.append(cx); ppe.append(end_ppe); equity.append(end_equity); cash.append(end_cash)
    cfo.append(operating_cf); cfi.append(investing_cf); cff.append(financing_cf); net_cash_change.append(change); checks.append(check)
    prev_rev=rev; prev_ppe=end_ppe; prev_debt=end_debt; prev_equity=end_equity; prev_cash=end_cash
    prev_ar=end_ar; prev_inv=end_inv; prev_oca=end_oca; prev_ap=end_ap; prev_ocl=end_ocl

income_rows={"Revenue":revenue,"Cost of goods sold":[-x for x in cogs],"Gross Profit":gross,"Operating expenses":[-x for x in opex],"EBITDA":ebitda,"Depreciation":[-x for x in dep],"EBIT":ebit,"Finance cost":[-x for x in interest],"Profit Before Tax":pbt,"Income tax":[-x for x in tax],"Profit After Tax":pat}
bs_rows={"Cash & cash equivalents":cash,"Trade receivables":ar,"Inventory":inventory,"Other current assets":oca,"Net property, plant & equipment":ppe,"Total Assets":[cash[i]+ar[i]+inventory[i]+oca[i]+ppe[i] for i in range(3)],"Trade payables":ap,"Other current liabilities":ocl,"Borrowings":debt,"Shareholders' equity":equity,"Total Liabilities & Equity":[ap[i]+ocl[i]+debt[i]+equity[i] for i in range(3)],"Balance Check":checks}
cf_rows={"Profit After Tax":pat,"Depreciation":dep,"Change in receivables":[-(ar[i]-(opening_ar if i==0 else ar[i-1])) for i in range(3)],"Change in inventory":[-(inventory[i]-(opening_inventory if i==0 else inventory[i-1])) for i in range(3)],"Change in other current assets":[-(oca[i]-(opening_oca if i==0 else oca[i-1])) for i in range(3)],"Change in payables":[ap[i]-(opening_ap if i==0 else ap[i-1]) for i in range(3)],"Change in other current liabilities":[ocl[i]-(opening_ocl if i==0 else ocl[i-1]) for i in range(3)],"Net Cash from Operating Activities":cfo,"Capital expenditure":cfi,"Net Cash from Investing Activities":cfi,"Net borrowing / (repayment)":[debt[i]-(opening_debt if i==0 else debt[i-1]) for i in range(3)],"Dividends paid":[-x for x in dividends],"Net Cash from Financing Activities":cff,"Net Change in Cash":net_cash_change,"Opening Cash":[opening_cash]+cash[:-1],"Closing Cash":cash}

income_df, income_style=statement_table(income_rows,years)
bs_df, bs_style=statement_table(bs_rows,years)
cf_df, cf_style=statement_table(cf_rows,years)
assumptions_df=pd.DataFrame({"Assumption":["Revenue growth","COGS / Revenue","Opex / Revenue","DSO","DIO","DPO","Capex / Revenue"],**{years[i]:[growth[i],cogs_pct[i],opex_pct[i],dso[i],dio[i],dpo[i],capex_pct[i]] for i in range(3)}}).set_index("Assumption")

st.markdown(f'<div class="hero"><h1>Interlinked Financial Statements</h1><p>{company} · Three-Year Dynamic Forecast · {unit}</p><span class="tag">The Mountain Path Academy</span></div>',unsafe_allow_html=True)

if min(cash) < 0:
    st.warning("The model produces a negative cash balance. Discuss the funding gap with students and revise borrowing or operating assumptions.")

cols=st.columns(4)
cols[0].metric(f"Revenue · {years[-1]}",money(revenue[-1]),f"{growth[-1]:.1%}")
cols[1].metric(f"EBITDA · {years[-1]}",money(ebitda[-1]),f"{ebitda[-1]/revenue[-1]:.1%} margin")
cols[2].metric(f"PAT · {years[-1]}",money(pat[-1]),f"{pat[-1]/revenue[-1]:.1%} margin")
cols[3].metric(f"Closing Cash · {years[-1]}",money(cash[-1]),f"Debt: {money(debt[-1])}")

tabs=st.tabs(["Executive Dashboard","Income Statement","Balance Sheet","Cash Flow","Assumptions & Drivers","Teaching Guide"])

with tabs[0]:
    st.markdown('<div class="note"><b>Model flow:</b> Revenue & margins → profit → working capital and PPE → cash flow → closing cash on the balance sheet.</div>',unsafe_allow_html=True)
    c1,c2=st.columns(2)
    with c1:
        fig=go.Figure()
        fig.add_trace(go.Bar(name="Revenue",x=years,y=revenue,marker_color="#0B5E8E"))
        fig.add_trace(go.Bar(name="EBITDA",x=years,y=ebitda,marker_color="#D6A43B"))
        fig.add_trace(go.Bar(name="PAT",x=years,y=pat,marker_color="#4F8A5B"))
        fig.update_layout(title="Profitability Trend",barmode="group",height=390,legend_orientation="h",margin=dict(t=55,b=20))
        st.plotly_chart(fig,use_container_width=True)
    with c2:
        fig=go.Figure()
        fig.add_trace(go.Scatter(name="Cash",x=years,y=cash,mode="lines+markers",line=dict(color="#D6A43B",width=4)))
        fig.add_trace(go.Scatter(name="Debt",x=years,y=debt,mode="lines+markers",line=dict(color="#9F2B2B",width=4)))
        fig.add_trace(go.Scatter(name="Equity",x=years,y=equity,mode="lines+markers",line=dict(color="#0B5E8E",width=4)))
        fig.update_layout(title="Funding & Capital Position",height=390,legend_orientation="h",margin=dict(t=55,b=20))
        st.plotly_chart(fig,use_container_width=True)
    st.subheader("Balance-sheet integrity")
    if max(abs(x) for x in checks) < 0.01:
        st.markdown('<div class="check-ok">✓ Balance sheet balances in every forecast year.</div>',unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="check-bad">⚠ Maximum imbalance: {max(abs(x) for x in checks):,.2f}</div>',unsafe_allow_html=True)

with tabs[1]:
    st.subheader("Forecast Income Statement")
    st.caption(f"Amounts in {unit}. Negative values are expenses.")
    st.dataframe(income_style,use_container_width=True,height=485)
with tabs[2]:
    st.subheader("Forecast Balance Sheet")
    st.dataframe(bs_style,use_container_width=True,height=500)
with tabs[3]:
    st.subheader("Forecast Cash Flow Statement")
    st.dataframe(cf_style,use_container_width=True,height=650)
with tabs[4]:
    st.subheader("Key forecast drivers")
    display_assumptions=assumptions_df.copy()
    for row in ["Revenue growth","COGS / Revenue","Opex / Revenue","Capex / Revenue"]:
        display_assumptions.loc[row]=display_assumptions.loc[row]*100
    st.dataframe(display_assumptions.style.format("{:,.1f}"),use_container_width=True)
    st.info("Percentages are displayed as percentage points; DSO, DIO and DPO are displayed in days.")
    st.markdown("**Core formulas used**")
    st.latex(r"Revenue_t = Revenue_{t-1}\times(1+Growth_t)")
    st.latex(r"Receivables_t = Revenue_t\times DSO_t/365")
    st.latex(r"Inventory_t = COGS_t\times DIO_t/365")
    st.latex(r"PPE_t = PPE_{t-1}+Capex_t-Depreciation_t")
    st.latex(r"Cash_t = Cash_{t-1}+CFO_t+CFI_t+CFF_t")
    st.latex(r"Equity_t = Equity_{t-1}+PAT_t-Dividends_t")
with tabs[5]:
    st.subheader("Suggested classroom demonstration")
    st.markdown("""
1. Increase revenue growth and observe the effect on profit, receivables, inventory and cash.
2. Increase DSO to demonstrate how profit can rise while operating cash flow deteriorates.
3. Increase inventory days and explain the working-capital cash lock-up.
4. Increase capex and trace depreciation, PPE, investing cash flow and closing cash.
5. Change borrowing to fund a cash deficit; observe finance cost and debt on the balance sheet.
6. Increase dividend payout and trace the effect on financing cash flow, cash and equity.
7. Confirm that closing cash from the cash-flow statement equals cash on the balance sheet.
    """)
    st.warning("This is an educational forecasting model, not investment, accounting, tax or audit advice.")

st.markdown("---")
c1,c2=st.columns([1,2])
with c1:
    workbook=excel_download({"Income Statement":income_df,"Balance Sheet":bs_df,"Cash Flow":cf_df,"Assumptions":assumptions_df})
    st.download_button("⬇ Download model output as Excel",workbook,file_name=f"{company.replace(' ','_')}_3_statement_model.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
with c2:
    st.caption("Designed for financial-modelling education by The Mountain Path Academy · themountainpathacademy.com")
