"""Module providing a function printing python version."""
from babel.numbers import format_currency
import streamlit as st
import pandas as pd
import plotly.express as px

# Set Streamlit page configuration
st.set_page_config(page_title="Sales Dashboard",
                   page_icon=":bar_chart:",
                   layout="centered"
                   )

# Allow the user to upload an XLSX file
uploaded_file = st.file_uploader('Choose a XLSX file', type='xlsx')

# Function to read data from the uploaded Excel file
CACHE_KEY = "data_" + str(uploaded_file)


@st.cache_data(ttl=3600)
def get_data_from_excel(file):
    """Reads data from excel file and return DataFrame """
    df_int = pd.read_excel(
        file,
        engine='openpyxl',
        sheet_name='Report 1',
        skiprows=1, usecols='G, I, L, Q, R, U, V, AF, AG, AJ'
    )
    return df_int


if uploaded_file:
    df = get_data_from_excel(uploaded_file)

    # Filter the DataFrame
    df = df.rename(columns={'Net Trade Sales in TAR @ AOP FX': 'Total',
                            'Sales Order PO Number': 'PO Number',
                            'Net Trade Sales Qty in Base UOM': 'Quantity'})

    df = df[df["Ship To Name"].str.contains(
        "Qateef Central Hospital|Ministry of Health Dammam|Dammam Central Hospital|"
        "DAMMAM MEDICAL COMPLEX MOH|Ministry of Health Al-Ahsa|"
        "NUPCO Dammam DC - MOH Ahsa|Prince Saud Bin Jalawy Hospital|"
        "Ministry of Health Hafr Al-Batin|NUPCO Dammam DC -  MOH"
    )]

    df = df[df["Sales Force Id"].str.contains(
        "SA801|SA802|SA806"
    )]

    # Remove everything but numbers
    # df['PO Number'] = df['PO Number'].str.extract(r'^(\d+)', expand=False)

    po_mapping = {'4300000911-2021100147001': '4300000911',
                  '4600033909 - 2021100209001': '4600033909'}
    df['PO Number'] = df['PO Number'].replace(po_mapping)

    df.dropna(subset=["Total", "PO Number"], inplace=True)
    df = df[df["Total"] != 0.000]

    df['Invoice Number'] = df['Invoice Number'].astype(str)  # remove commas

    cot_mapping = {'SA801': 'EBD',
                   'SA802': 'EMID',
                   'SA806': 'GYN'}
    df['Sales Force Id'] = df['Sales Force Id'].replace(cot_mapping)

    mpg_mapping = {'J5': 'ACC', 'L6': 'ES', 'M2': 'HW', 'M4': 'HI', 'N6': 'OS',
                   'P3': 'ST', 'Q3': 'VS', 'R9': 'TrueClear', 'U4': 'Skin Stapler'}
    df['MPG Id'] = df['MPG Id'].replace(mpg_mapping)

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
        options=list(df['Fiscal Qtr'].unique()),
        default=list(df['Fiscal Qtr'].unique())
    )
    filters['Fiscal Qtr'] = quarter_options
    df = filter_data(df, 'Fiscal Qtr', quarter_options)

    # Select Sales Rep filter
    rep_options = st.sidebar.multiselect(
        "Select Sales Rep:",
        options=list(df['Sales Rep Name'].unique()),
        default=list(df['Sales Rep Name'].unique())
    )
    filters['Sales Rep Name'] = rep_options
    df = filter_data(df, 'Sales Rep Name', rep_options)

    # Select PO Number filter
    po_options = st.sidebar.multiselect(
        "Select PO Number:",
        options=['Select All'] + list(df['PO Number'].unique()),
        default=['Select All']
    )
    if 'Select All' in po_options:
        po_options = df['PO Number'].unique()
    else:
        po_options = list(set(po_options).intersection(
            set(df['PO Number'].unique())))
    filters['PO Number'] = po_options
    df = filter_data(df, 'PO Number', po_options)

    # Select Ship To Name filter
    ship_options = st.sidebar.multiselect(
        "Select Ship-To:",
        options=['Select All'] + list(df['Ship To Name'].unique()),
        default=['Select All']
    )
    if 'Select All' in ship_options:
        ship_options = df['Ship To Name'].unique()
    else:
        ship_options = list(set(ship_options).intersection(
            set(df['Ship To Name'].unique())))
    filters['Ship To Name'] = ship_options
    df = filter_data(df, 'Ship To Name', ship_options)

    # ---- MAINPAGE ----
    st.title(":bar_chart: Sales Dashboard")
    st.markdown("##")

    # TOP KPI's
    total_sales = round(df['Total'].sum(), 2)

    cpad1, col, pad2 = st.columns((10, 10, 10))
    with col:
        st.subheader("Total Sales:")
        st.subheader(f"US ${total_sales:,}")
    st.markdown("##")

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

    # Sales by ACCOUNTs Bar Chart
    ship_to = df.groupby(
        by=["Ship To Name"], group_keys=False).sum()[["Total"]]

    ship_to_large = ship_to.nlargest(10, 'Total')

    ship_to_large["formatted_text"] = (ship_to_large["Total"]
                                       .apply(lambda x: format_currency(x, 'USD',
                                                                        locale='en_US',
                                                                        currency_digits=True)))
    ship_to_large["hover_data"] = (df.groupby("Ship To Name")["PO Number"]
                                   .unique()
                                   .apply(lambda x: '<br>'.join(['PO Number: ' + i for i in x])))

    fig_acc = px.bar(
        ship_to_large,
        x="Total",
        y=ship_to_large.index,
        text='formatted_text',
        text_auto=False,
        hover_data=['hover_data'],
        title="<b>Sales by Accounts</b>",
        color_discrete_sequence=["#0e72b5"] * len(ship_to_large),
        template="plotly_white",
        orientation='h'
    )
    fig_acc.update_traces(textposition='auto',
                          hovertemplate='%{customdata[0]}')

    fig_acc.update_layout(
        xaxis=(dict(showgrid=False)),
        # yaxis={'categoryorder': 'total descending'},
        yaxis=dict(tickmode="auto"),
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis_title="Total Sales",
        yaxis_title="Account",
        margin=dict(
            l=30,
            r=30,
            b=20,
            t=30,
            pad=10
        ),
    )

    st.plotly_chart(fig_acc, use_container_width=True,
                    color=ship_to)

    st.markdown("""---""")

    # Sales by COT Bar & Pie Chart
    bar7, bar8, bar9 = st.columns((20, 10, 20))

    with bar7:

        cot = df.groupby(
            by=["Sales Force Id"], group_keys=False).sum()[["Total"]]

        cot["formatted_text"] = (cot["Total"]
                                 .apply(lambda x: format_currency(x, 'USD',
                                                                  locale='en_US',
                                                                  currency_digits=True)))

        fig_cot = px.bar(
            cot,
            y="Total",
            x=cot.index,
            text='formatted_text',
            text_auto=False,
            # hover_data=['hover_data'],
            title="<b>Sales by COT</b>",
            color_discrete_sequence=["#0e72b5"] * len(cot),
            template="plotly_white",
            orientation='v'
        )
        fig_cot.update_traces(textposition='outside',
                              hovertemplate=None)

        fig_cot.update_layout(
            xaxis=(dict(showgrid=False)),
            yaxis=dict(tickmode="auto"),
            plot_bgcolor="rgba(0,0,0,0)",
            xaxis_title="COT",
            yaxis_title="Total Sales",
            margin=dict(
                l=30,
                r=30,
                b=50,
                t=50,
                pad=10
            ),
        )
        st.plotly_chart(fig_cot, use_container_width=True,
                        color=cot)

    with bar9:
        fig = px.pie(
            df,
            values='Total',
            names='Sales Force Id',
            # title='<b>Total Sales by COT (%)</b>',
            color_discrete_sequence=px.colors.diverging.RdYlBu_r,
            hole=0.3
        )
        fig.update_traces(textposition='inside',
                          textinfo='percent+label')
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("""---""")

    # Top 10 Sales by CFN Bar chart
    total_sales = df["Total"].sum()

    # Group by CFN and calculate the sum of "Total" for each CFN
    sales_by_cfn = df.groupby(by=["CFN Id"], group_keys=False).sum()[["Total"]]

    # Sort the data by "Total" in descending order and take the top 10
    top_10_sales = sales_by_cfn.nlargest(10, 'Total')

    # Calculate the percentage of each CFN's total sales as a percentage of the total sum
    top_10_sales["Percentage"] = (top_10_sales["Total"] / total_sales) * 100

    # Add a formatted text column for display
    top_10_sales["formatted_text"] = top_10_sales["Total"].apply(
        lambda x: format_currency(x, 'USD', locale='en_US', currency_digits=True))

    # Create the hovertext with formatted text and percentage
    hover_text = top_10_sales.apply(
        lambda row: f"{row['formatted_text']}<br>({row['Percentage']:.2f}%)", axis=1)

    fig_CFN = px.bar(
        top_10_sales,
        y="Total",
        x=top_10_sales.index,
        text=hover_text,
        text_auto=False,
        title="<b>Top 10 Sales by CFN</b>",
        color_discrete_sequence=["#0e72b5"],
        template="plotly_white",
        orientation='v'
        # height=700
    )

    fig_CFN.update_traces(textposition='outside',
                          hovertemplate='%{text}', cliponaxis=False)

    fig_CFN.update_layout(
        # xaxis=dict(tickmode="auto"),
        xaxis=(dict(showgrid=False)),
        # yaxis={'categoryorder': 'total descending'},
        yaxis=dict(tickmode="auto"),
        xaxis_title="CFN",
        yaxis_title="Total Sales",
        margin=dict(
            l=60,
            r=30,
            b=50,
            t=50,
            pad=10
        ),
    )

    fig.update_yaxes(scaleratio=10)

    st.plotly_chart(fig_CFN, use_container_width=True)

    st.markdown("""---""")

    # Total by MPG ID
    st.text("Total Sales by MPG:")

    # Group by 'MPG Id', sum, and then reset the index to make 'MPG Id' a column
    mpg_total = df.groupby(by=["MPG Id"]).sum().reset_index()[
        ["MPG Id", "Total"]]

    # Calculate the overall total
    overall_total = mpg_total['Total'].sum()

    # Add a percentage column
    mpg_total['Percentage'] = (mpg_total['Total'] / overall_total) * 100

    # Format the 'Total' column to include commas for thousands
    mpg_total['Total'] = mpg_total['Total'].apply(lambda x: f"$ {x:,.2f}")

    # Format the 'Percentage' column to show as a percentage with 2 decimal places
    mpg_total['Percentage'] = mpg_total['Percentage'].apply(
        lambda x: f"{x:.2f}%")

    TABLE_WIDTH = "100%"

    # Centering the table and adjusting its width with CSS
    st.write(
        f"""
        <style>
        .my-table {{
            margin: 0 auto;
            text-align: center;
            width: {TABLE_WIDTH};
            color: WhiteSmoke;
        }}
        .my-table th {{
            text-align: center;
            color: Peru;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )

    st.write(
        mpg_total.to_html(classes=["my-table"], index=False),
        unsafe_allow_html=True
    )

    st.markdown("""---""")

    st.text("Total Sales:")
    st.text(f"US ${total_sales:,.2f}")
    # quantity_sum_by_cfn = df.groupby(
    #     "CFN Id")["Quantity"].agg('sum').reset_index()

    # CFN/Qty/Total HTML Table
    quantity_sum_by_cfn = df.groupby('CFN Id').agg(
        {'Quantity': 'sum', 'Total': 'sum'}).reset_index()

    quantity_sum_by_cfn = quantity_sum_by_cfn[quantity_sum_by_cfn["Quantity"] != 0.000]

    quantity_sum_by_cfn = quantity_sum_by_cfn.sort_values(
        by='Total', ascending=False)

    quantity_sum_by_cfn['%'] = (
        (quantity_sum_by_cfn['Total'] / total_sales) * 100).round(2)

    quantity_sum_by_cfn['Total'] = quantity_sum_by_cfn['Total'].apply(
        lambda x: f"$ {x:,.2f}")

    quantity_sum_by_cfn['Quantity'] = quantity_sum_by_cfn['Quantity'].astype(
        int)

    quantity_sum_by_cfn['Quantity'] = quantity_sum_by_cfn['Quantity'].apply(
        lambda x: f"{x:,}")

    quantity_sum_by_cfn = quantity_sum_by_cfn.rename(columns={'CFN Id': 'CFN',
                                                              'Quantity': 'Total Quantity',
                                                              'Total': 'Total Sales'})

    TABLE_WIDTH = "100%"

    # Centering the table and adjusting its width with CSS
    st.write(
        f"""
        <style>
        .my-table {{
            margin: 0 auto;
            text-align: center;
            width: {TABLE_WIDTH};
            color: WhiteSmoke;
        }}
        .my-table th {{
            text-align: center;
            color: Peru;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )
    st.write(
        pd.DataFrame(quantity_sum_by_cfn[['CFN', 'Total Quantity', 'Total Sales', '%']]).to_html(
            classes=["my-table"], index=False),
        unsafe_allow_html=True
    )

    # Display the filtered DataFrame
    # st.dataframe(df, use_container_width=True)

    # sorted_df = quantity_sum_by_cfn.sort_values(by='Total', ascending=False)
    # st.dataframe(sorted_df, hide_index=True,
    #              use_container_width=True)
