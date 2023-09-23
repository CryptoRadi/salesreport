    
    
 #how to add checkbox to enable/disable filter   
    
    filter_enabled = st.sidebar.checkbox("Enable PO Number filter", value=True)
    if filter_enabled:
        options = st.sidebar.multiselect(
            "Select PO Number:",
            options=df['PO Number'].unique(),
            default=df['PO Number'].unique()
        )
        options = list(set(options).intersection(set(df['PO Number'].unique())))
        filters['PO Number'] = options
        df = filter_data(df, 'PO Number', options)