from babel.numbers import format_currency
import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Sales Dashboard",
                   page_icon=":bar_chart:",
                   layout="wide"
                   )

st.session_state.df = None
uploaded_file = st.file_uploader('Choose a XLSX file', type='xlsx')

# ---- READ EXCEL ----
@st.experimental_memo
def get_data_from_excel():
    """Reads data from excel file and return DataFrame """
    df = pd.read_excel(
        uploaded_file,
        engine='openpyxl',
        sheet_name='Report 1',
        skiprows=1,usecols='P, Q, T, AD, AE, AI'
        )
    return df

if st.session_state.df is None:
    if uploaded_file:
        df = get_data_from_excel()
        st.session_state.df = df

        df = df[df["Sales Rep Name"].str.contains(
        "@MOH TENDER|RADI, UMMAIR|HARARAH, AHMED|ABOELENAN, SHAFEIK|IBRAHIM, AHMED|"
        "ALZAHRANI, SAEED|MAREY, MOSTAFA|AL-HAID, RAZAN WALID|KHATER, AHMED|ABU QAZA, AHMED|"
        "TURKISTANI, ADNAN|ALQAHTANI, ABDULLAH|UNASSIGNED"
        )]

        df = df.rename(columns={'Net Trade Sales in TAR @ AOP FX': 'Total',
                                'Sales Order PO Number': 'PO Number',
                                'GFS Std Level 9 Desc': 'Group'})

        df['PO Number'] = df['PO Number'].str.extract(r'^(\d+)', expand=False)

        df.dropna(subset=["Total", "PO Number"], inplace=True)
        df = df[df["Total"] != 0.000]

        # ---- SIDEBAR ----
        st.sidebar.info("When filtering, please make sure to CLEAR \n- SELECT ALL")
        st.sidebar.header("Please Filter Here:")

        def filter_data(data, column, options):
            """
            Filter the dataframe using the given column and options.
            :param data: The DataFrame to be filtered
            :param column: The column on which to filter the data
            :param options: The options to include in the filtered data
            :returns: The filtered DataFrame
            """
            if options is not None:
                data = data[data[column].isin(options)]
            return data

        filters = {}

        # Select Qaurter filter
        options = st.sidebar.multiselect(
            "Select Quarter:",
            options=['Select All'] + list(df['Fiscal Qtr'].unique()),
            default=['Select All']
            )
        if 'Select All' in options:
            options = df['Fiscal Qtr'].unique()
        else:
            options = list(set(options).intersection(set(df['Fiscal Qtr'].unique())))
        filters['Fiscal Qtr'] = options
        df = filter_data(df, 'Fiscal Qtr', options)

        # Select Sales Rep filter
        options = st.sidebar.multiselect(
            "Select Sales Rep:",
            options=['Select All'] + list(df['Sales Rep Name'].unique()),
            default=['Select All']
            )
        if 'Select All' in options:
            options = df['Sales Rep Name'].unique()
        else:
            options = list(set(options).intersection(set(df['Sales Rep Name'].unique())))
        filters['Sales Rep Name'] = options
        df = filter_data(df, 'Sales Rep Name', options)

        # Select PO Number filter
        options = st.sidebar.multiselect(
            "Select PO Number:",
            options=['Select All'] + list(df['PO Number'].unique()),
            default=['Select All']
        )
        if 'Select All' in options:
            options = df['PO Number'].unique()
        else:
            options = list(set(options).intersection(set(df['PO Number'].unique())))
        filters['PO Number'] = options
        df = filter_data(df, 'PO Number', options)

        # ---- MAINPAGE ----
        st.title(":bar_chart: Sales Dashboard")
        st.markdown("##")

        # TOP KPI's
        total_sales = round(df['Total'].sum(), 2)

        cpad1, col, pad2 = st.columns((10, 10, 10))
        with col:
            st.subheader("Total Sales:")
            st.subheader(f"US $ {total_sales:,}")
        st.markdown("##")

        st.markdown("""---""")

        # CHARTS
        sales_by_rep = df.groupby(by=["Sales Rep Name"], group_keys=False).sum()[["Total"]]

        # Sales Rep Bar Chart
        # Show text outside the bar in USD
        sales_by_rep["formatted_text"] = (sales_by_rep["Total"]
                                    .apply(lambda x: format_currency(x, 'USD',
                                    locale='en_US',
                                    currency_digits=True)))
        sales_by_rep["hover_data"] = (df.groupby("Sales Rep Name")["PO Number"]
                                .unique()
                                .apply(lambda x: '<br>'.join(['PO Number: ' + i for i in x])))

        fig_sales = px.bar(
            sales_by_rep,
            x="Total",
            y=sales_by_rep.index,
            text='formatted_text',
            text_auto=False,
            hover_data=['hover_data'],
            title="<b>Sales by Sales Rep</b>",
            color_discrete_sequence=["#0e72b5"] * len(sales_by_rep),
            template="plotly_white",
            orientation='h'
        )
        fig_sales.update_traces(textposition='outside', hovertemplate= '%{customdata[0]}')

        fig_sales.update_layout(
            xaxis=(dict(showgrid=False)),
            yaxis=dict(tickmode="linear"),
            plot_bgcolor="rgba(0,0,0,0)",
            xaxis_title="Total Sales",
            yaxis_title="Sales Rep",
            margin=dict(
            l=30,
            r=30,
            b=50,
            t=50,
            pad=10
        ),
        )
        st.plotly_chart(fig_sales, use_container_width=True, color=sales_by_rep)

        # Quarter Pie Chart
        fig = px.pie(
            df,
            values='Total',
            names='Fiscal Qtr',
            title='<b>Total Sales by Quarter (%)</b>',
            color_discrete_sequence=px.colors.diverging.RdYlBu_r,
            hole=0.3
        )
        fig.update_traces(textposition='inside', textinfo='percent+label')
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("""---""")

        st.dataframe(df, use_container_width=True)
else:
    df = st.session_state.df
