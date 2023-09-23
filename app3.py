"""Module providing a function printing python version."""
from babel.numbers import format_currency
import streamlit as st
import pandas as pd
import plotly.express as px

# Set Streamlit page configuration
st.set_page_config(page_title="Sales Dashboard",
                   page_icon=":bar_chart:",
                   layout="wide"
                   )

# Allow the user to upload an XLSX file
uploaded_file = st.file_uploader('Choose a XLSX file', type='xlsx')

# Function to read data from the uploaded Excel file
@st.cache_data
def get_data_from_excel():
    """Reads data from excel file and return DataFrame """
    df = pd.read_excel(
        uploaded_file,
        engine='openpyxl',
        sheet_name='Report 1',
        skiprows=1,usecols='G, Q, R, U, V, AF, AJ'
        )
    return df

if uploaded_file:
    df = get_data_from_excel()

    # Filter the DataFrame
    df = df[df["Sales Rep Name"].str.contains(
    "RADI, UMMAIR|UNASSIGNED"
    )]

    df = df.rename(columns={'Net Trade Sales in TAR @ AOP FX': 'Total',
                            'Sales Order PO Number': 'PO Number'})

    df = df[~df["Ship To Name"].str.contains(
    "Zahran Operations & Maintanance Co.|NUPCO Dammam -Ryl Commision Store|"
    "NUPCO Jeddah DC MOI-Security Forces|NUPCO Jeddah MOE -KAUH|"
    "Prince Sultan Military Medical City|Al Marjan medical center company|"
    "NUPCO Qassim DC - MOH Store|Care Medical Center-Riyadh|"
    "Najran University Hospital"
    )]

    df['PO Number'] = df['PO Number'].str.extract(r'^(\d+)', expand=False)

    df.dropna(subset=["Total", "PO Number"], inplace=True)
    df = df[df["Total"] != 0.000]

    df['Invoice Number'] = df['Invoice Number'].astype(str) #remove commas

    # ---- SIDEBAR ----
    st.sidebar.info("When filtering, please make sure to CLEAR \n- SELECT ALL")
    st.sidebar.header("Please Filter Here:")

    # Function to filter the DataFrame based on user-selected options
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
    quarter_options = st.sidebar.multiselect(
        "Select Quarter:",
        options=['Select All'] + list(df['Fiscal Qtr'].unique()),
        default=['Select All']
        )
    if 'Select All' in quarter_options:
        quarter_options = df['Fiscal Qtr'].unique()
    else:
        quarter_options = list(set(quarter_options).intersection(set(df['Fiscal Qtr'].unique())))
    filters['Fiscal Qtr'] = quarter_options
    df = filter_data(df, 'Fiscal Qtr', quarter_options)

    # Select Sales Rep filter
    rep_options = st.sidebar.multiselect(
        "Select Sales Rep:",
        options=['Select All'] + list(df['Sales Rep Name'].unique()),
        default=['Select All']
        )
    if 'Select All' in rep_options:
        rep_options = df['Sales Rep Name'].unique()
    else:
        rep_options = list(set(rep_options).intersection(set(df['Sales Rep Name'].unique())))
    filters['Sales Rep Name'] = rep_options
    df = filter_data(df, 'Sales Rep Name', rep_options)

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

    # Select Ship To Name filter
    options = st.sidebar.multiselect(
        "Select Ship-To:",
        options=['Select All'] + list(df['Ship To Name'].unique()),
        default=['Select All']
    )
    if 'Select All' in options:
        options = df['Ship To Name'].unique()
    else:
        options = list(set(options).intersection(set(df['Ship To Name'].unique())))
    filters['Ship To Name'] = options
    df = filter_data(df, 'Ship To Name', options)

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

    # Display the filtered DataFrame
    st.dataframe(df, use_container_width=True)
