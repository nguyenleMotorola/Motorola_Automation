from email.policy import default
import pandas as pd
import streamlit as st
import altair as alt
from st_aggrid import AgGrid
from st_aggrid.grid_options_builder import GridOptionsBuilder
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go
st.set_page_config(
    page_title="Avigilon Error by Month",
    layout="wide"
)

# cache the dataframe to memory so we don't have to read it over and over again
@st.cache(allow_output_mutation=True)
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

    error_dict = dict.fromkeys(sections) 
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
        error_dict[section] = section_failures
    return error_dict, report_date

@st.cache(allow_output_mutation=True)
def extract_month_from_date(report_date):
    d = report_date.split(",")[1]
    month = d.split(" ")[1]
    return month
@st.cache(allow_output_mutation=True)
def generate_error_paretos(error_df, topN):
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
            df = df.sort_values(by=['Qty'], ascending=False)[0:st.session_state["topN"]] # grab the top N 

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
            # Create the bars with length encoded along the Y axis
            bars1 = base.mark_bar(size = 25, color="#67B7D1").encode(
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
            pareto = (bars + bars1 + line + points + point_text).resolve_scale(
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


# ------ Home page ------
if __name__ == '__main__':
    st.title("💡 Avigilon Error by Month")
    st.markdown("##")

    st.sidebar.header('Drag and drop your Excel Files here')
    uploaded_files = st.sidebar.file_uploader('Select Monthly Avigilon Files', type='csv', key=3, accept_multiple_files=True)


    if uploaded_files:
        st.session_state["topN"] = st.sidebar.slider("Select the number of top results for each month:", 0, 10, 5)
        # Create a Pandas Excel writer using XlsxWriter as the engine.
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        model_list = []
        monthly_error_by_model_dict = {}
        error_dict_by_month = {}
        month_list = []

        for file in uploaded_files:
            error_dict, report_date = get_data(file)
            month = extract_month_from_date(report_date) # get month from date string 
            error_dict_by_month[month] = error_dict["Forge Inspection"]
            error_dict_by_month[month]['Month'] = month 
            error_dict_by_month[month]['Qty'] = error_dict_by_month[month]['Qty'].astype(int)
            model_list.extend(error_dict_by_month[month]['Model'].unique())
            month_list.append(month)

        top10_per_model = {}
        model_list = [*set(model_list)]
        # Model Sections
        section_string = ""

        for model in model_list:
            model_with_hyphen = ""
            if ' ' in model:
                model_with_hyphen = model.replace(" ", "-").lower()
            else:
                model_with_hyphen = model.lower()
            section_string += f'''
                                - [{model}](#{model_with_hyphen})
                                '''
        st.sidebar.markdown(section_string, unsafe_allow_html=True)

        for i, model in enumerate(model_list):
            # st.write(str(i) + model)
            for z, month in enumerate(month_list):
                # st.write(str(z) + month)
                if model in error_dict_by_month[month]['Model'].values: # the model is in this month 
                    # grab the top 10 error for this model in this month
                    model_df = error_dict_by_month[month][error_dict_by_month[month]['Model'] == model]
                    if model in top10_per_model:
                        top10_per_model[model] = pd.concat([top10_per_model[model], model_df[0:st.session_state["topN"]]], ignore_index=True)
                    else:
                        top10_per_model[model] = model_df[0:st.session_state["topN"]]
            top10_per_model[model]["Failure Code"] = top10_per_model[model].iloc[:, 0].str.split(":").str[0] # Canceled by operator, just grab "Canceled"

        for model in model_list:
            failure_code_list = top10_per_model[model]["Failure Code"].unique()
            chart_count = 0
            bar_charts = list()
            st.header(model)
            with st.expander("Expand to see data for this model"):
                st.dataframe(top10_per_model[model])
            for failure_code in failure_code_list:
                chart_count += 1
                df = top10_per_model[model][top10_per_model[model]["Failure Code"] == failure_code]
                bar_chart = alt.Chart(df).mark_bar().encode(
                    alt.Column('Failure Code', title=" "),
                    alt.X('Month', title=" ", sort='-y'),
                    alt.Y('Qty', axis = alt.Axis(grid=False), sort="ascending"),
                    alt.Color('Month'),
                    tooltip = [alt.Tooltip('Inspection Failures'),
                    alt.Tooltip('Qty'),
                    alt.Tooltip('Month')]
                )
                if chart_count == 1:
                    chart1 = bar_chart
                elif chart_count == 2:
                    chart2 = bar_chart
                elif chart_count == 3:
                    chart3 = bar_chart
                elif chart_count == 4:
                    chart4 = bar_chart
                elif chart_count == 5:
                    chart5 = bar_chart
                elif chart_count == 6:
                    chart6 = bar_chart
                elif chart_count == 7:
                    chart7 = bar_chart
                elif chart_count == 8:
                    chart8 = bar_chart
                elif chart_count == 9:
                    chart9 = bar_chart
                elif chart_count == 10:
                    chart10 = bar_chart
                elif chart_count == 11:
                    chart11 = bar_chart
                elif chart_count == 12:
                    chart12 = bar_chart
                elif chart_count == 13:
                    chart13 = bar_chart
                elif chart_count == 14:
                    chart14 = bar_chart
                elif chart_count == 15:
                    chart15 = bar_chart
                elif chart_count == 16:
                    chart16 = bar_chart
                elif chart_count == 17:
                    chart17 = bar_chart
                elif chart_count == 18:
                    chart18 = bar_chart
                elif chart_count == 19:
                    chart19 = bar_chart
                elif chart_count == 20:
                    chart20 = bar_chart
                elif chart_count == 21:
                    chart21 = bar_chart
                elif chart_count == 22:
                    chart22 = bar_chart
                elif chart_count == 23:
                    chart23 = bar_chart
                elif chart_count == 24:
                    chart24 = bar_chart
                elif chart_count == 25:
                    chart25 = bar_chart
                elif chart_count == 26:
                    chart26 = bar_chart
                elif chart_count == 27:
                    chart27 = bar_chart
                elif chart_count == 28:
                    chart28 = bar_chart
                elif chart_count == 29:
                    chart29 = bar_chart
                elif chart_count == 30:
                    chart30 = bar_chart

            if chart_count == 1:
                (chart1)
            elif chart_count == 2:
                (chart1 | chart2)
            elif chart_count == 3:
                (chart1 | chart2 | chart3)
            elif chart_count == 4:
                (chart1 | chart2 | chart3 | chart4)
            elif chart_count == 5:
                (chart1 | chart2 | chart3 | chart4 | chart5)
            elif chart_count == 6:
                (chart1 | chart2 | chart3 | chart4 | chart5) & (chart6)
            elif chart_count == 7:
                (chart1 | chart2 | chart3 | chart4 | chart5) & (chart6 | chart7)
            elif chart_count == 8:
                (chart1 | chart2 | chart3 | chart4 | chart5) & (chart6 | chart7 | chart8)
            elif chart_count == 9:
                (chart1 | chart2 | chart3 | chart4 | chart5) & (chart6 | chart7 | chart8 | chart9)
            elif chart_count == 10:
                (chart1 | chart2 | chart3 | chart4 | chart5) & (chart6 | chart7 | chart8 | chart9 | chart10)
            elif chart_count == 11:
                (chart1 | chart2 | chart3 | chart4 | chart5) & (chart6 | chart7 | chart8 | chart9 | chart10) & (chart11)
            elif chart_count == 12:
                (chart1 | chart2 | chart3 | chart4 | chart5) & (chart6 | chart7 | chart8 | chart9 | chart10) & (chart11 | chart12)
            elif chart_count == 13:
                (chart1 | chart2 | chart3 | chart4 | chart5) & (chart6 | chart7 | chart8 | chart9 | chart10) & (chart11 | chart12 | chart13)
            elif chart_count == 14:
                (chart1 | chart2 | chart3 | chart4 | chart5) & (chart6 | chart7 | chart8 | chart9 | chart10) & (chart11 | chart12 | chart13 | chart14)
            elif chart_count == 15:
                (chart1 | chart2 | chart3 | chart4 | chart5) & (chart6 | chart7 | chart8 | chart9 | chart10) & (chart11 | chart12 | chart13 | chart14 | chart15)
            elif chart_count == 16:
                (chart1 | chart2 | chart3 | chart4 | chart5) & (chart6 | chart7 | chart8 | chart9 | chart10) & (chart11 | chart12 | chart13 | chart14 | chart15) \
                & (chart16)
            elif chart_count == 17:
                (chart1 | chart2 | chart3 | chart4 | chart5) & (chart6 | chart7 | chart8 | chart9 | chart10) & (chart11 | chart12 | chart13 | chart14 | chart15) \
                & (chart16 | chart17)
            elif chart_count == 18:
                (chart1 | chart2 | chart3 | chart4 | chart5) & (chart6 | chart7 | chart8 | chart9 | chart10) & (chart11 | chart12 | chart13 | chart14 | chart15) \
                & (chart16 | chart17 | chart18)
            elif chart_count == 19:
                (chart1 | chart2 | chart3 | chart4 | chart5) & (chart6 | chart7 | chart8 | chart9 | chart10) & (chart11 | chart12 | chart13 | chart14 | chart15) \
                & (chart16 | chart17 | chart18 | chart19)
            elif chart_count == 20:
                (chart1 | chart2 | chart3 | chart4 | chart5) & (chart6 | chart7 | chart8 | chart9 | chart10) & (chart11 | chart12 | chart13 | chart14 | chart15) \
                & (chart16 | chart17 | chart18 | chart19 | chart20)
            elif chart_count == 21:
                (chart1 | chart2 | chart3 | chart4 | chart5) & (chart6 | chart7 | chart8 | chart9 | chart10) & (chart11 | chart12 | chart13 | chart14 | chart15) \
                & (chart16 | chart17 | chart18 | chart19 | chart20) & (chart21)
            elif chart_count == 22:
                (chart1 | chart2 | chart3 | chart4 | chart5) & (chart6 | chart7 | chart8 | chart9 | chart10) & (chart11 | chart12 | chart13 | chart14 | chart15) \
                & (chart16 | chart17 | chart18 | chart19 | chart20) & (chart21 | chart22)
            elif chart_count == 23:
                (chart1 | chart2 | chart3 | chart4 | chart5) & (chart6 | chart7 | chart8 | chart9 | chart10) & (chart11 | chart12 | chart13 | chart14 | chart15) \
                & (chart16 | chart17 | chart18 | chart19 | chart20) & (chart21 | chart22 | chart23)
            elif chart_count == 24:
                (chart1 | chart2 | chart3 | chart4 | chart5) & (chart6 | chart7 | chart8 | chart9 | chart10) & (chart11 | chart12 | chart13 | chart14 | chart15) \
                & (chart16 | chart17 | chart18 | chart19 | chart20) & (chart21 | chart22 | chart23 | chart24)
            elif chart_count == 25:
                (chart1 | chart2 | chart3 | chart4 | chart5) & (chart6 | chart7 | chart8 | chart9 | chart10) & (chart11 | chart12 | chart13 | chart14 | chart15) \
                & (chart16 | chart17 | chart18 | chart19 | chart20) & (chart21 | chart22 | chart23 | chart24 | chart25)
            elif chart_count == 26:
                (chart1 | chart2 | chart3 | chart4 | chart5) & (chart6 | chart7 | chart8 | chart9 | chart10) & (chart11 | chart12 | chart13 | chart14 | chart15) \
                & (chart16 | chart17 | chart18 | chart19 | chart20) & (chart21 | chart22 | chart23 | chart24 | chart25) & (chart26)
            elif chart_count == 27:
                (chart1 | chart2 | chart3 | chart4 | chart5) & (chart6 | chart7 | chart8 | chart9 | chart10) & (chart11 | chart12 | chart13 | chart14 | chart15) \
                & (chart16 | chart17 | chart18 | chart19 | chart20) & (chart21 | chart22 | chart23 | chart24 | chart25) & (chart26 | chart27)
            elif chart_count == 28:
                (chart1 | chart2 | chart3 | chart4 | chart5) & (chart6 | chart7 | chart8 | chart9 | chart10) & (chart11 | chart12 | chart13 | chart14 | chart15) \
                & (chart16 | chart17 | chart18 | chart19 | chart20) & (chart21 | chart22 | chart23 | chart24 | chart25) & (chart26 | chart27 | chart28)
            elif chart_count == 29:
                (chart1 | chart2 | chart3 | chart4 | chart5) & (chart6 | chart7 | chart8 | chart9 | chart10) & (chart11 | chart12 | chart13 | chart14 | chart15) \
                & (chart16 | chart17 | chart18 | chart19 | chart20) & (chart21 | chart22 | chart23 | chart24 | chart25) & (chart26 | chart27 | chart28 | chart29)
            elif chart_count == 30:
                (chart1 | chart2 | chart3 | chart4 | chart5) & (chart6 | chart7 | chart8 | chart9 | chart10) & (chart11 | chart12 | chart13 | chart14 | chart15) \
                & (chart16 | chart17 | chart18 | chart19 | chart20) & (chart21 | chart22 | chart23 | chart24 | chart25) & (chart26 | chart27 | chart28 | chart29 | chart30)
            