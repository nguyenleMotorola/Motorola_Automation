import pandas as pd
import streamlit as st
from st_aggrid import AgGrid
from st_aggrid.grid_options_builder import GridOptionsBuilder
from io import BytesIO
import altair as alt
import plotly.graph_objects as go
from streamlit_metrics import metric

st.set_page_config(
    page_title="Avigilon Yield",
    layout="wide"
)

# cache the dataframe to memory so we don't have to read it over and over again
@st.cache(suppress_st_warning=True)
def get_data(file_name): 
    data = pd.read_csv(file_name)

    rty = data.iat[0,3] # Rolled Throughput Yield
    report_date = data.iat[0,4] # get report date
    sections = ["Forge Set-Parameters", "Forge Burn-in", "Forge Inspection", "Forge Eyeball Lens-Tuning", "Forge Lens-Tuning", "Forge Base-Programming"]

    total_calc_indices = data.loc[data['ReportTitle'].str.contains("textbox", case=False)].index.tolist() # indices to locate total calculations, index + 1 is the row containing the values
    yield_table_indices = data.loc[data['ReportTitle'].str.contains("report_group", case=False)].index.tolist() # indices to locate yield data, index + 1 is the row containing the values
    failure_table_indices = data.loc[data['ReportTitle'].str.contains("failure", case=False)].index.tolist() # indices to locate failure data, index + 1 is the row containing the values
    stop_indices = total_calc_indices + yield_table_indices + failure_table_indices # indices to locate each tables where they start and end 
    stop_indices.sort()

    yield_dict = dict.fromkeys(sections) 

    # # Find Total Calculations Tables 
    # # 1) get Total Calculations for each section, it's the top 4 values (Yield, Total, Passed, Failed)
    # # Search for lower case cell value = "textbox" in the first column, the next row are the total calculation values 
    # # Get all row indices that the first column contains "textbox", case insensitive
    # #
    for i, section in enumerate(sections):
        section_total = data.iloc[[total_calc_indices[i] + 1]].dropna(axis='columns', how='all') # get row index + 1, this row contains the total calculations values
        if "report" not in section_total.iat[0,0].lower(): # if total calculations row is not blank
            if section != "Forge Base-Programming":
                section_total.columns = ['Yield', 'Total', 'Passed', 'Failed']
            else:
                section_total.columns = ['Sudo-Yield', 'Total', 'Passed', 'Failed']
        else:
            section_total = pd.DataFrame() # empty dataframe if no data
        yield_dict[section] = [section_total]    
    

    # # Find Yield Tables 
    # # Get Yield Table for each section, it's the table below the top 4 Values
    # #
    for i, section in enumerate(sections):
        start_index = yield_table_indices[i] + 1
        # print(stop_indices.index(yield_table_indices[i] ))
        end_index = stop_indices[stop_indices.index(yield_table_indices[i] ) + 1] # get the next beginning of the next table as the end index for this table
        section_yield = data.iloc[start_index:end_index].dropna(axis='columns', how='all') # get row index + 1, this row contains the total calculations values
        if len(section_yield.columns) == 5:
            section_yield.columns = ['Model', 'Total', 'Passed', 'Failed', 'Yield']
        elif len(section_yield.columns) == 6:
            section_yield.columns = ['Model', 'Convert', 'Total', 'Passed', 'Failed', 'Yield']
        yield_dict[section].append(section_yield)

    # # Find Failure Tables #
    failure_dict = {
        "failuremode3": "Set-Parameters Failures",
        "failuremode": "Burn-in Failures",
        "failuremode2":"Inspection Failures",
        "failuremode4": "Eyeball Lens-Tuning Failures",
        "failuremode1": "Lens-Tuning Failures",
        "failuremode5": "Base-Programming Failures",
    }

    for i, section in enumerate(sections):
        start_index = failure_table_indices[i] + 1
        table_header = data.iat[failure_table_indices[i],0]
        
        if table_header.lower() == "failuremode5":
            section_failures = data.iloc[start_index:].dropna(axis='columns', how='all')
        else:
            end_index = stop_indices[stop_indices.index(failure_table_indices[i]) + 1] # get the next beginning of the next table as the end index for this table
            section_failures = data.iloc[start_index:end_index].dropna(axis='columns', how='all') # get row index + 1, this row contains the total calculations values
        # assign columns
        header = failure_dict[table_header.lower()]
        if len(section_failures.columns) == 3: # only 3 columns (failures, model, qty)
            section_failures.columns = [header, 'Model', 'Qty']
        elif len(section_failures.columns) == 4: # 4 columns (failures, model, qty)
            section_failures.columns = [header, 'Model', 'Convert', 'Qty']
        yield_dict[section].append(section_failures)
    return yield_dict, rty, report_date

def generate_error_paretos(error_df):
    if len(error_df) > 0:
        model_list = error_df['Model'].unique()
        col1, col2 = st.columns(2) 
        paretos = []
        for model in model_list:
            
            df = error_df.loc[error_df['Model'] == model]
            df['Qty'] = df['Qty'].astype(int)
            df['percent'] = (df['Qty'] / (df['Qty'].sum())).transform('{:,.1%}'.format)
            df['cumpercentage'] = (df['Qty'].cumsum())/(df['Qty'].sum())
            df["Failure Code"] = df.iloc[:, 0].str.split(":").str[0] # extract codes
            df["Description"] = df.iloc[:, 0]
            df["Label"] = df["Failure Code"].map(str) + " " + df['Qty'].map(str) + ' ' + df['percent'].map(str)
            df = df.sort_values(by=['Qty'], ascending=False)

            sort_order = df['Qty'].tolist()
            base = alt.Chart(df).encode(
                x = alt.X("combined:N", sort=sort_order, axis=alt.Axis(labelAngle=360), title="Error Code / Quantity / Percentage"), # O stands for a discrete ordered quantity
            ).properties (
                title=model
            ).transform_calculate(
                combined = "split(datum.Label, ' ')" 
            )

            # Create the bars with length encoded along the Y axis
            bars = base.mark_bar(size = 25, color="#67B7D1").encode(
                y = alt.Y("Qty:Q", title="Count", axis=alt.Axis(labels=False)),
                tooltip=['Description']
            )

            # Create the line chart with length encoded along the Y axis
            line = base.mark_line(strokeWidth = 1.5, color = "#cb4154"
            ).encode(
                y = alt.Y("cumpercentage:Q", title="Cumulative Percentage", axis=alt.Axis(format=".1%")),
            )

            # Mark the percentage values on the line with Circle marks
            points = base.mark_circle(strokeWidth = 3, color = "#cb4154"
            ).encode(
                y = alt.Y('cumpercentage:Q', axis=None)
            )

            # Mark the Circle marks with the value text
            point_text = points.mark_text(
                align='left',
                baseline='middle',
                dx = -10, 
                dy = -10,
                size = 12
            ).encode(
                y= alt.Y('cumpercentage:Q', axis=None),
                # we'll use the percentage as the text
                text=alt.Text('cumpercentage:Q', format="0.0%"),
                color= alt.value("#cb4154")
            )
            # Layer all the elements together 
            pareto = (bars + line + points + point_text).resolve_scale(
                y = 'independent'
            )
            paretos.append(pareto)

        chunked_list = list()
        chunk_size = int(len(model_list) / 2) + (len(model_list) % 2 > 0)
        for i in range(0, len(paretos), chunk_size):
            chunked_list.append(paretos[i:i+chunk_size])
        with col1:
            for pareto in chunked_list[0]:
                st.altair_chart(pareto, use_container_width=True)
        if len(chunked_list) > 1:
            with col2:
                for pareto in chunked_list[1]:
                    st.altair_chart(pareto, use_container_width=True)


# ------ Main page ------
st.title("🔥 Avigilon Dashboard")
st.markdown("##")

st.sidebar.header('Drag and drop your Excel File here')
uploaded_AV_file = st.sidebar.file_uploader('Choose a csv file', type='csv', key=2)

if uploaded_AV_file:
    yield_dict, rty, report_date = get_data(uploaded_AV_file)
    metric("Rolled Throughput Yield", rty)
    st.markdown(f"<h3 style='text-align: center;'>{report_date}</h3>", unsafe_allow_html=True)

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')

    for yield_name, yield_values in yield_dict.items():
        with st.expander(yield_name):
            yield_title = yield_name.partition(' ')[2]# remove word Forge from title
            if yield_values:
                failure_table = None
                for value_index, value in enumerate(yield_values):
                    if value_index == 0: # total calculations
                        st.markdown(f"<h3 style='text-align: center;'>Total ({yield_title})</h3>", unsafe_allow_html=True)
                        total_table = go.Figure(data=go.Table(
                                                            header=dict(values=list(value.columns), fill_color='#90ee90'), 
                                                            cells=dict(values=value.values.flatten().tolist(), fill_color='#E5ECF6')
                                                            )
                                                )
                        total_table.update_layout(margin=dict(l=1,r=1,b=1,t=1), autosize=False,
                                                                        height=50)
                        total_table.update_traces(cells_font=dict(size = 15))

                        st.plotly_chart(total_table, use_container_width=True)
                        if len(yield_values) == 3: # create columns
                            cols = st.columns([1, 1.5])
                        elif len(yield_values) < 3:
                            cols = st.columns(1) 
                    elif value_index == 1:
                        with cols[0]: # yield table
                            st.markdown(f"<h3 style='text-align: center;'>Yield Table ({yield_title})</h3>", unsafe_allow_html=True)
                            yield_table = value
                            yield_table.to_excel(writer, sheet_name=yield_name, index=False) # write to excel
                            gb_yield = GridOptionsBuilder.from_dataframe(yield_table)
                            gb_yield.configure_pagination()
                            gb_yield.configure_side_bar()
                            gb_yield.configure_default_column(groupable=True, value=True, enableRowGroup=True, editable=True)
                            # gb_yield.configure_column('Model', rowGroup=True)
                            gridOptions = gb_yield.build()
                            AgGrid(yield_table, gridOptions=gridOptions, enable_enterprise_modules=True)
                    elif value_index == 2: # failure table
                        with cols[1]:
                            st.markdown(f"<h3 style='text-align: center;'>Error Code Table ({yield_title})</h3>", unsafe_allow_html=True)
                            failure_table = value
                            if len(failure_table) > 0:                                                             
                                failure_table.to_excel(writer, sheet_name=failure_table.columns[0], index=False)
                                gb_failure = GridOptionsBuilder.from_dataframe(failure_table)
                                gb_failure.configure_pagination()
                                gb_failure.configure_side_bar()
                                gb_failure.configure_default_column(groupable=True, value=True, enableRowGroup=True, editable=True)
                                gb_failure.configure_column('Model', rowGroup=True)
                                gridOptions = gb_failure.build()
                                AgGrid(failure_table, gridOptions=gridOptions, enable_enterprise_modules=True)

                        # Create Pareto chart for error codes https://medium.com/analytics-vidhya/creating-a-dual-axis-pareto-chart-in-altair-e3673107dd14
                        st.markdown(f"<h3 style='text-align: center;'>Error Code Pareto Charts ({yield_title})</h3>", unsafe_allow_html=True)
                        generate_error_paretos(failure_table)
    # save and export to excel
    writer.save()
    excel_data = output.getvalue()
    # Close the Pandas Excel writer and output the Excel file when user click a button.
    st.download_button(label="Export data to Excel", file_name='Avigilon Report Data.xlsx', data=excel_data)

        