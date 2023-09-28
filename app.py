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
cache_key = "data_" + str(uploaded_file)


@st.cache_data(experimental_allow_widgets=True)
def get_data_from_excel(file):
    """Reads data from excel file and return DataFrame """
    df = pd.read_excel(
        file,
        engine='openpyxl',
        sheet_name='Report 1',
        skiprows=1, usecols='G, Q, R, U, V, AF, AJ'
    )
    return df


if uploaded_file:
    df = get_data_from_excel(uploaded_file)

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
        "NUPCO Qassim DC - MOH Store|Care Medical Center-Riyadh|NUPCO KAEC DC MODA - KSAFH Tabuk|"
        "Najran University Hospital|King Khaled Hospital - Majmaah|Ministry of Health Bisha"
    )]

    df['PO Number'] = df['PO Number'].str.extract(r'^(\d+)', expand=False)

    df.dropna(subset=["Total", "PO Number"], inplace=True)
    df = df[df["Total"] != 0.000]

    df['Invoice Number'] = df['Invoice Number'].astype(str)  # remove commas

    # if st.button("Clear Cache"):
    #     st.rerun()

    # ---- SIDEBAR ----
    st.sidebar.info(
        "When filtering, please make sure to CLEAR: \n- Select All")
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
        quarter_options = list(set(quarter_options).intersection(
            set(df['Fiscal Qtr'].unique())))
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
        rep_options = list(set(rep_options).intersection(
            set(df['Sales Rep Name'].unique())))
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
        options = list(set(options).intersection(
            set(df['PO Number'].unique())))
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
        options = list(set(options).intersection(
            set(df['Ship To Name'].unique())))
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

    # Sales Rep Charts
    bar1, bar2, bar3 = st.columns((20, 10, 20))
    with bar1:
        sales_by_rep = df.groupby(
            by=["Sales Rep Name"], group_keys=False).sum()[["Total"]]

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
            y="Total",
            x=sales_by_rep.index,
            text='formatted_text',
            text_auto=False,
            hover_data=['hover_data'],
            title="<b>Sales by Sales Rep</b>",
            color_discrete_sequence=["#0e72b5"] * len(sales_by_rep),
            template="plotly_white",
            orientation='v'
        )
        fig_sales.update_traces(textposition='outside',
                                hovertemplate='%{customdata[0]}')

        fig_sales.update_layout(
            xaxis=(dict(showgrid=False)),
            yaxis=dict(tickmode="auto"),
            plot_bgcolor="rgba(0,0,0,0)",
            xaxis_title="Sales Rep",
            yaxis_title="Total Sales",
            margin=dict(
                l=30,
                r=30,
                b=50,
                t=50,
                pad=10
            ),
        )
        st.plotly_chart(fig_sales, use_container_width=True,
                        color=sales_by_rep)

    with bar3:
        fig = px.pie(
            df,
            values='Total',
            names='Sales Rep Name',
            # title='<b>Total Sales by Quarter (%)</b>',
            color_discrete_sequence=px.colors.diverging.RdYlBu_r,
            hole=0.3
        )
        fig.update_traces(textposition='outside', textinfo='percent+label')
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("""---""")

    # Quarter Charts
    bar4, bar5, bar6 = st.columns((20, 10, 20))
    with bar4:

        sales_by_qtr = df.groupby(
            by=["Fiscal Qtr"], group_keys=False).sum()[["Total"]]

        sales_by_qtr["formatted_text"] = (sales_by_qtr["Total"]
                                          .apply(lambda x: format_currency(x, 'USD',
                                                                           locale='en_US',
                                                                           currency_digits=True)))
        sales_by_qtr["hover_data"] = (df.groupby("Fiscal Qtr")["PO Number"]
                                      .unique()
                                      .apply(lambda x: '<br>'.join(['PO Number: ' + i for i in x])))

        fig_qtr = px.bar(
            sales_by_qtr,
            y="Total",
            x=sales_by_qtr.index,
            text='formatted_text',
            text_auto=False,
            hover_data=['hover_data'],
            title="<b>Sales by Quarter</b>",
            color_discrete_sequence=["#0e72b5"] * len(sales_by_qtr),
            template="plotly_white",
            orientation='v'
        )
        fig_qtr.update_traces(textposition='outside',
                              hovertemplate='%{customdata[0]}')

        fig_qtr.update_layout(
            xaxis=(dict(showgrid=False)),
            yaxis=dict(tickmode="auto"),
            plot_bgcolor="rgba(0,0,0,0)",
            xaxis_title="Quarter",
            yaxis_title="Total Sales",
            margin=dict(
                l=30,
                r=30,
                b=50,
                t=50,
                pad=10
            ),
        )
        st.plotly_chart(fig_qtr, use_container_width=True,
                        color=sales_by_qtr)

    with bar6:
        fig = px.pie(
            df,
            values='Total',
            names='Fiscal Qtr',
            # title='<b>Total Sales by Quarter (%)</b>',
            color_discrete_sequence=px.colors.diverging.RdYlBu_r,
            hole=0.3
        )
        fig.update_traces(textposition='inside',
                          textinfo='percent+label')
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("""---""")

    # CFN Charts
    sales_by_cfn = df.groupby(
        by=["CFN Id"], group_keys=False).sum()[["Total"]]

    top_10_sales = sales_by_cfn.nlargest(10, 'Total')
    top_10_sales["formatted_text"] = top_10_sales["Total"].apply(
        lambda x: format_currency(x, 'USD', locale='en_US', currency_digits=True))

    # ALL CFNs
    # sales_by_cfn["formatted_text"] = (sales_by_cfn["Total"]
    #                                   .apply(lambda x: format_currency(x, 'USD',
    #                                                                    locale='en_US',
    #                                                                    currency_digits=True)))
    fig_CFN = px.bar(
        # sales_by_cfn,
        top_10_sales,
        y="Total",
        # x=sales_by_cfn.index,
        x=top_10_sales.index,
        text='formatted_text',
        text_auto=False,
        title="<b>Top 10 Sales by CFN</b>",
        color_discrete_sequence=["#0e72b5"],
        template="plotly_white",
        orientation='v'
    )
    fig_CFN.update_traces(textposition='outside', hovertemplate=None)

    fig_CFN.update_layout(
        yaxis=dict(tickmode="auto"),
        xaxis={'categoryorder': 'total descending'},
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis_title="CFN",
        yaxis_title="Total Sales",
        margin=dict(
            l=30,
            r=30,
            b=50,
            t=50,
            pad=10
        ),
    )
    st.plotly_chart(fig_CFN, use_container_width=True)

    # HIDE LESS THAN THRESHOLD
    THRESHOLD_PERCENTAGE = 3
    total_percentage = df['Total'].sum()
    threshold_value = total_percentage * (THRESHOLD_PERCENTAGE / 100)
    df_grouped = df.copy()
    df_grouped.loc[df_grouped['Total'] < threshold_value, 'CFN Id'] = 'Others'
    df_grouped['Percentage'] = df_grouped['Total'] / total_percentage * 100
    df_grouped = df_grouped[df_grouped['CFN Id'] != 'Others'].round(2)

    fig_2 = px.pie(
        df_grouped,
        values='Percentage',
        names='CFN Id',
        color_discrete_sequence=px.colors.diverging.RdYlBu_r,
        hole=0.4,
        hover_data=['Total']
    )

    fig_2.update_traces(title_text='<b>Higher than 3%</b>', title_position='middle center',
                        textposition='inside',
                        texttemplate='%{label}<br>%{value:.2f}%')
    st.plotly_chart(fig_2, use_container_width=True)

    st.markdown("""---""")

    # Display the filtered DataFrame
    # st.dataframe(df, use_container_width=True)
