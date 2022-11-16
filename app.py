# pylint: disable=missing-module-docstring
# pylint: disable=missing-class-docstring
# pylint: disable=missing-function-docstring
import streamlit as st
import pandas as pd
import plotly.express as px


st.set_page_config(page_title="Sales Dashboard",
                   page_icon=":bar_chart:",
                   layout="wide"
                   )

uploaded_file = st.file_uploader('Choose a XLSX file', type='xlsx')
# ---- READ EXCEL ----


@st.cache
def get_data_from_excel():
    df = pd.read_excel(
        # io='SI,RI,PMGand RCS FY23 sales report.xlsx',
        uploaded_file,
        engine='openpyxl',
        sheet_name='Report 1',
        skiprows=1,
        usecols='P, Q, T, AD, AE, AI',
    )
    return df


if uploaded_file:
    df = get_data_from_excel()

    df = df[df["Sales Rep Name"].str.contains(
        "RADI, UMMAIR|ABU QAZA, AHMED|@MOH TENDER")]

    df.rename(columns={'Net Trade Sales in TAR @ AOP FX': 'Total',
                       'Sales Order PO Number': 'PO Number'}, inplace=True)

    df = df[df['PO Number'].str.contains('Sample') == False]
    df = df[df['PO Number'].str.contains('sample') == False]
    df['PO Number'] = df['PO Number'].str.split(' ').str[0]
    df['PO Number'] = df['PO Number'].str.split('-').str[0]

    df.dropna(subset=["Total", "PO Number"], inplace=True)

    # st.dataframe(df)

    # # ---- SIDEBAR ----
    st.sidebar.header("Please Filter Here:")

    qtr = st.sidebar.multiselect(
        "Select Quarter:",
        options=df['Fiscal Qtr'].unique(),
        default=df['Fiscal Qtr'].unique()
    )

    if qtr is not True:
        sales_rep = df['Sales Rep Name'].where(
            df['Fiscal Qtr'].isin(qtr)).dropna()

        sales_rep = st.sidebar.multiselect(
            "Select Sales Rep:",
            options=sales_rep.unique(),
            default=sales_rep.unique()
        )
        if sales_rep is not True:
            po = df['PO Number'].where(
                df['Sales Rep Name'].isin(sales_rep)).dropna()
            po = df['PO Number'].where(
                df['Fiscal Qtr'].isin(qtr)).dropna()

            po = st.sidebar.multiselect(
                "Select PO Number:",
                options=po.unique(),
                default=po.unique()
            )

            st.sidebar.header("Risk Assessment:")

            if po is not True:
                invoice = df['Invoice Number'].where(
                    df['PO Number'].isin(po)).dropna()
                invoice = st.sidebar.multiselect(
                    "Select Invoice Number:",
                    options=invoice.unique(),
                    # default=invoice.unique()
                )

    df_selection = df.query(
        "`Fiscal Qtr` == @qtr & `Sales Rep Name` == @sales_rep & `PO Number` == @po"
    )
    df_selection_inv = df.query(
        "`Fiscal Qtr` == @qtr & `Sales Rep Name` == @sales_rep & `Invoice Number` == @invoice & `PO Number` == @po"
    )

    # ---- MAINPAGE ----
    st.title(":bar_chart: Sales Dashboard")
    st.markdown("##")

    # TOP KPI's
    total_sales = int(df_selection["Total"].sum())
    total_invoice = int(df_selection_inv["Total"].sum())
    total_risk = int(total_sales - total_invoice)

    cpad1, col, pad2 = st.columns((10, 10, 10))
    with col:
        st.subheader("Total Sales:")
        st.subheader(f"US $ {total_sales:,}")
    st.markdown("##")

    cpad1, col, pad2 = st.columns((10, 10, 10))
    with col:
        st.subheader("Risk: ")
        st.subheader(f" US $ {total_invoice:,}")
    st.markdown("##")

    cpad1, col, pad2 = st.columns((10, 10, 10))
    with col:
        st.subheader("Total: ")
        st.subheader(f" US $ {total_risk:,}")

        # f'<h1 style="color:#33ff33;"> US $ {total_invoice: ,}</h1>', unsafe_allow_html=True)

    st.markdown("""---""")

    # SALES REP [BAR CHART]
    sales_by_rep = df_selection.groupby(by=["Sales Rep Name"]).sum()[["Total"]]
    fig_sales = px.bar(
        sales_by_rep,
        x=sales_by_rep.index, text_auto=True,
        y="Total",
        title="<b>Sales by Sales Rep Name</b>",
        color_discrete_sequence=["#0083B8"] * len(sales_by_rep),
        template="plotly_white",
    )
    fig_sales.update_layout(
        xaxis=dict(tickmode="linear"),
        plot_bgcolor="rgba(0,0,0,0)",
        yaxis=(dict(showgrid=False)),
    )
    st.plotly_chart(fig_sales, use_container_width=True, color=sales_by_rep)

    # st.dataframe(df_selection)

# ---- HIDE STREAMLIT STYLE ----
# hide_st_style = """
#             <style>
#             #MainMenu {visibility: hidden;}
#             footer {visibility: hidden;}
#             header {visibility: hidden;}
#             </style>
#             """
# st.markdown(hide_st_style, unsafe_allow_html=True)
