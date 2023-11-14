import pandas as pd
from st_aggrid import AgGrid, GridUpdateMode
from st_aggrid.grid_options_builder import GridOptionsBuilder
import streamlit as st
import collections
from st_aggrid.shared import JsCode
from datetime import datetime
import plotly.graph_objs as go
from plotly import tools

st.set_page_config(
    page_title="WatchGuard Yield",
    layout="wide"
)

# ----- Helper function -----
def prod(val) :    
    res = 1        
    for ele in val:        
        res *= ele        
    return res

# cache the dataframe to memory so we don't have to read it over and over again
# Create a dictionary whose key is the sheet name and values are the codes in that sheet name
@st.cache()
def get_data_from_excel(file_name):
    df = pd.read_excel(io=file_name)
    df = df[['Test Name', 'Product Category', 'Product Test Type', 'Product Tested', 'Product Tested Description', 'Qty Failed', 'Qty Passed', 'Test Category']] # extract only the relevant columns

    df["Product Tested Description"] = df["Product Tested Description"].str.upper() # convert all codes to upper case
    # remove any rows that has "PCBA" in Product Tested and "DWG" in Product Tested Description
    df = df[df["Product Tested"].str.contains("PCBA") == False]
    df = df[df["Product Tested Description"].str.contains("DWG") == False]
    # delete WGA00300 (discontinued), ignore all WGP
    df = df[df["Product Tested Description"].str.contains("WGA00300") == False]
    df = df[df["Product Tested Description"].str.contains("WGP") == False]

    #remove all codes with -CS at the end
    df = df[df["Product Tested Description"].str.contains("-CS") == False]

    # replace _ with -
    df["Product Tested Description"] = df["Product Tested Description"].str.replace('_', '-')

    family_desc_sheet = collections.defaultdict(list) # create an empty dict
    prod_test_desc = df["Product Tested Description"].unique()
    for desc in prod_test_desc:
        temp_desc = desc.split("-")[0] 
        if temp_desc in ["WGA00356", "WGA00508"]:
            family_desc_sheet["HiFi Mic"].append(desc)

        elif temp_desc in ["WGA00751", "WGA00750"]:
            family_desc_sheet["HiFi Mic 2"].append(desc)

        elif temp_desc in ["WGA00359"]:
            family_desc_sheet["HiFi Base"].append(desc)

        elif temp_desc in ["WGA00370", "WGA00526"]:
            family_desc_sheet["4RE Display"].append(desc)

        elif temp_desc in ["WGA00383", "WGA00391", "WGA00607"]:
            family_desc_sheet["4RE POE"].append(desc)

        elif temp_desc in ["WGA00480", "WGA00615"]:
            family_desc_sheet["4RE DVRS"].append(desc)

        elif temp_desc in ["WGA00437", "WGA00496", "WGA00498", "WGA00500", "WGA00543", "WGA00485", "WGA00522"]:
            family_desc_sheet["4RE Camera"].append(desc)

        elif temp_desc in ["WGA00450"]:
            family_desc_sheet["DV1"].append(desc)
        
        elif temp_desc in ["WGA00451"]:
            family_desc_sheet["DV1 MOD"].append(desc)

        elif temp_desc in ["WGA00461", "WGA00468"]:
            family_desc_sheet["DV1 CAM"].append(desc)

        elif temp_desc in ["WGA00675", "WGA00682", "WGA00684", "WGA00690", "WGA00691", "WGA00700"]:
            family_desc_sheet["M500"].append(desc)

        elif temp_desc in ["WGA00548"]:
            family_desc_sheet["VISTA"].append(desc)

        elif temp_desc in ["WGA00520", "WGA00552", "WGA00600", "WGA00600"]:
            family_desc_sheet["VISTA HD"].append(desc)

        elif temp_desc in ["WGA00555"]:
            family_desc_sheet["VISTA TS"].append(desc)

        elif temp_desc in ["WGA00574"]:
            family_desc_sheet["VISTA POE"].append(desc)

        elif temp_desc in ["WGA00576", "WGA00583", "WGA00584"]: 
            family_desc_sheet["VISTA XLT"].append(desc)
        
        elif temp_desc in ["WGA00578", "WGA00582", "WGA00XXX"]: # ask if this is always the case
            family_desc_sheet["VISTA XLT Head Cam"].append(desc)

        elif temp_desc in ["WGA00586", "WGA00537"]:
            family_desc_sheet["VISTA WiFi Charge Base"].append(desc)

        elif temp_desc in ["WGA00608"]:
            family_desc_sheet["VISTA USB Charge Base"].append(desc)
        
        elif temp_desc in ["WGA00625", "WGA00627"]:
            family_desc_sheet["V300"].append(desc)

        elif temp_desc in ["WGA00635", "WGA00640"]:
            family_desc_sheet["V300 Docks"].append(desc)

        elif temp_desc in ["WGA00650"]:
            family_desc_sheet["V300 TS2"].append(desc)

    return df, family_desc_sheet





if __name__ == '__main__':
    # ------ USER INTERFACE ------
    # ------ Home page ------
    st.title(":bar_chart: WatchGuard Yield Dashboard")
    st.markdown("##")

    st.sidebar.header('Drag and drop your Excel File here')
    uploaded_file = st.sidebar.file_uploader('Choose a XLSX file', type='xlsx', key=1)

    if uploaded_file:
        st.session_state['df'], family_desc_sheet = get_data_from_excel(uploaded_file)

        # --- Sidebar ---
        sheet_names =   ["Overall",\
                        "HiFi Mic", "HiFi Base", \
                        "4RE Display", "4RE POE", "4RE DVRS", "4RE Camera", \
                        "DV1", "DV1 MOD", "DV1 CAM",\
                        "M500",\
                        "VISTA", "VISTA HD", "VISTA TS", "VISTA POE", "VISTA XLT", "VISTA XLT Head Cam", "VISTA WiFi Charge Base", "VISTA USB Charge Base",\
                        "V300", "V300 BWC", "V300 Docks", "V300 TS2"]
        st.session_state['selected_family_desc'] = st.sidebar.selectbox('Select Data:', options = ["Overall"] + sorted(list(family_desc_sheet.keys())))
        # --- Main Page ---
        rty_dict = dict.fromkeys(family_desc_sheet.keys()) # get list of families

        # ============== OVERALL PAGE ============== """
        if st.session_state['selected_family_desc'] == "Overall":  
            st.session_state['yield_dict'] = {}
            final_average = 0
            final_length = 0 # number of families that has final values
            run_in_average = 0
            run_in_length = 0 # number of families that has run-in values
            st.markdown(f"<h1 style='text-align: center;'>Overall Rolled Throughput Yield</h1>", unsafe_allow_html=True)          
            for family_key in family_desc_sheet.keys(): # for each family, calculate their RTY and yield
                list_of_codes = family_desc_sheet.get(family_key)
                if list_of_codes:
                    full_data = st.session_state['df'].loc[st.session_state['df']['Product Tested Description'].isin(family_desc_sheet.get(family_key))]

                    # calculate yield and its RTY
                    yield_by_prod = full_data.groupby(by=['Product Test Type', 'Product Tested', 'Product Tested Description']).sum().reset_index()
                    yield_by_prod['Yield'] = yield_by_prod["Qty Passed"] / (yield_by_prod["Qty Passed"] + yield_by_prod["Qty Failed"])
                    # drop any yield row with value 0 before calculating the RTY
                    yield_rows = yield_by_prod['Yield'].to_frame() # convert series to dataframe
                    yield_rows = yield_rows.loc[~(yield_rows==0).all(axis=1)] # drop rows with 0
                    yield_rows = yield_rows.squeeze() # convert dataframe to series
                    rty = yield_rows.prod() # calculate RTY
                    rty_dict[family_key] = rty

                    yield_by_prod['Yield'] = yield_by_prod['Yield'].astype(float).map("{:.1%}".format)

                    # calculate yield by product test type
                    yield_by_prod_test_type = full_data.groupby(by=['Product Test Type']).sum().reset_index()
                    yield_by_prod_test_type['Yield'] = yield_by_prod_test_type["Qty Passed"] / (yield_by_prod_test_type["Qty Passed"] + yield_by_prod_test_type["Qty Failed"])
                    yield_by_prod_test_type['Yield'] = yield_by_prod_test_type['Yield'].astype(float).map("{:.1%}".format)
                    
                    yield_by_prod_and_test_type = []
                    yield_by_prod_and_test_type.append(yield_by_prod)
                    yield_by_prod_and_test_type.append(yield_by_prod_test_type)
                    st.session_state['yield_dict'][family_key] = yield_by_prod_and_test_type


                    curr_final = yield_by_prod_test_type[yield_by_prod_test_type['Product Test Type'] == "Final"]['Yield']
                    if not curr_final.empty:
                        curr_final_float = float(curr_final.values[0].strip('%')) / 100
                        final_average += curr_final_float
                        final_length += 1

                    curr_run_in = yield_by_prod_test_type[yield_by_prod_test_type['Product Test Type'] == "Run-In"]['Yield']
                    if not curr_run_in.empty:
                        curr_run_in_float = float(curr_run_in.values[0].strip('%')) / 100
                        run_in_average += curr_run_in_float
                        run_in_length += 1
            
            # FINAL AVG is the avg of all final values by product test type (the second yield table)
            # RUN IN AVG is the avg of all run in values by product test type (the second yield table)
            final_average = final_average / final_length
            run_in_average = run_in_average / run_in_length
            overall_rty = "{:.1%}".format(final_average * run_in_average) # OVERALL RTY = FINAL AVG * RUN_IN AVERAGE
            
            # overall_rty = 1
            # for i in rty_dict:
            #     overall_rty = overall_rty*rty_dict[i]
            # overall_rty = "{:.1%}".format(overall_rty)
            st.markdown(f"<h2 style='text-align: center;'>{overall_rty}</h2>", unsafe_allow_html=True)

            # st.metric("", "{:.1%}".format(overall_rty))

            st.markdown(f"<h1 style='text-align: center;'>RTY by family</h1>", unsafe_allow_html=True)
            col1, col2, col3 = st.columns(3)
            num_per_col = int(len(rty_dict) / 3) + (len(rty_dict) % 3 > 0)

            i = 0
            rty_dict = dict(sorted(rty_dict.items()))
            for key, value in rty_dict.items():
                value = "{:.1%}".format(value)
                if i < num_per_col: # 0-3 
                    with col1:
                        st.metric(key, value)
                        i = i + 1
                elif (i >= num_per_col) and (i < num_per_col*2): # 4-6
                    with col2:
                        st.metric(key, value)
                        i = i + 1
                else:
                    with col3:
                        st.metric(key, value)
                        i = i + 1      
            # NOTE: Create a top 10 chart for RTY 
            
            
        else:
            # ============== INDIVIDUAL FAMILY PAGE ============== """     
            full_data = st.session_state['df'].loc[st.session_state['df']['Product Tested Description'].isin(family_desc_sheet.get(st.session_state['selected_family_desc']))]
            
            dates_with_time = full_data['Test Name'].unique()
            dates = []
            for date in dates_with_time:
                dt = datetime.strptime(date.split(" ")[0], '%m/%d/%Y').date()
                dates.append(dt)
            dates.sort() # sort dates from past to present
            begin_date = dates[0] # get begin and end dates in a list
            end_date = dates[-1] 
            begin_date_str = begin_date.strftime('%m/%d/%Y') # begin date
            end_date_str = end_date.strftime('%m/%d/%Y') # end date
            date_title = begin_date_str + "  - " + end_date_str

            st.markdown(f"<h1 style='text-align: center;'>{st.session_state['selected_family_desc']}</h1>", unsafe_allow_html=True)
            st.markdown(f"<h2 style='text-align: center;'>{date_title}</h2>", unsafe_allow_html=True)

            # --- CSS --- https://discuss.streamlit.io/t/ag-grid-component-with-input-support/8108/138?page=7
            yield_color = JsCode(
                """
                    function(params) {
                        var yield = parseFloat(params.data.Yield) / 100.0;
                        if (yield <= 0.92) {
                            return { 'backgroundColor': '#ff0000' }
                        }
                        if (yield > 0.92) {
                            return { 'backgroundColor': '#99ff99' }
                        }
                    };
                """
            )

            # Display Yield by Product Test Type, Product Tested, and Product Tested Description
            col1, col2 = st.columns((1.5,1))
            with col1: 

                st.subheader("Yield")
                yield_by_prod = st.session_state['yield_dict'][st.session_state['selected_family_desc']][0]
                # yield_by_prod = full_data.groupby(by=['Product Test Type', 'Product Tested', 'Product Tested Description']).sum().reset_index()
                # yield_by_prod['Yield'] = yield_by_prod["Qty Passed"] / (yield_by_prod["Qty Passed"] + yield_by_prod["Qty Failed"])
                # yield_by_prod['Yield'] = yield_by_prod['Yield'].astype(float).map("{:.1%}".format)
                gb_yield_by_prod = GridOptionsBuilder.from_dataframe(yield_by_prod)
                gb_yield_by_prod.configure_default_column(value=True, editable=True)
                gb_yield_by_prod.configure_column("Yield", cellStyle=yield_color)
                gb_yield_by_prod.configure_column("Product Tested", hide=True)
                # sel_mode = st.radio('Selection Type', options = ['single', 'multiple'])
                gb_yield_by_prod.configure_selection(selection_mode='multiple', use_checkbox=True)
                gb_yield_by_prod.configure_column("Product Test Type", headerCheckboxSelection = True) # allow users to select rows with checkboxs
                gridOptions = gb_yield_by_prod.build()
                yield_by_prod_table = AgGrid(yield_by_prod, gridOptions=gridOptions, update_mode = GridUpdateMode.SELECTION_CHANGED, enable_enterprise_modules=True, fit_columns_on_grid_load = True, allow_unsafe_jscode=True)
                try: 
                    sel_row = pd.DataFrame(yield_by_prod_table["selected_rows"])
                    sel_row['Yield'] = sel_row['Yield'].str.rstrip("%").astype(float)/100 # convert percentages to floats to graph y axis values 
                    final_rows = sel_row[sel_row['Product Test Type'] == 'Final']
                    integration_rows = sel_row[sel_row['Product Test Type'] == 'Integration']
                    run_in_rows = sel_row[sel_row['Product Test Type'] == 'Run-In']
                    final_trace = go.Bar(
                        x = final_rows['Product Tested Description'],
                        y = final_rows['Yield'], 
                        text = final_rows['Yield'].apply(lambda x: '{0:.1f}%'.format(x*100)),
                        textposition = 'auto',
                        hovertemplate = '%{x}: %{text}',
                        name='Final'
                    )
                    integration_trace = go.Bar(
                        x = integration_rows['Product Tested Description'],
                        y = integration_rows['Yield'],
                        text = integration_rows['Yield'].apply(lambda x: '{0:.1f}%'.format(x*100)),
                        textposition = 'auto',
                        hovertemplate = '%{x}: %{text}',
                        name='Integration'
                    )
                    run_in_trace = go.Bar(
                        x = run_in_rows['Product Tested Description'],
                        y = run_in_rows['Yield'],
                        text = run_in_rows['Yield'].apply(lambda x: '{0:.1f}%'.format(x*100)),
                        textposition = 'auto',
                        hovertemplate = '%{x}: %{text}',
                        name='Run-in'            
                    )
                    fig = tools.make_subplots(rows=1, cols=3,
                            shared_xaxes=True, shared_yaxes=True,
                            vertical_spacing=0.001,
                            subplot_titles = ('Final', 'Integration', 'Run-In'))
                    fig.append_trace(final_trace, row=1, col=1) 
                    fig.append_trace(integration_trace, row=1, col=2)
                    fig.append_trace(run_in_trace, row=1, col=3)
                    fig.update_layout(margin=dict(l=0,r=0,b=0,t=0)) # remove white margin
                    # fig.update_yaxes(tickformat = ",.0%")
                    # fig.update_yaxes(hoverformat = ",.0%")
                    fig.update_layout(showlegend=True)
                    fig.update_layout(xaxis={'categoryorder':'total ascending'})
                    st.plotly_chart(fig)
                except:
                    st.warning('Please select rows by checking their checkboxes!' )

            with col2:
                # Display Yield by Product Test Type
                st.subheader("Yield by Product Test Type")
                yield_by_prod_test_type = st.session_state['yield_dict'][st.session_state['selected_family_desc']][1]
                gb_yield_by_prod_test_type = GridOptionsBuilder.from_dataframe(yield_by_prod_test_type)
                gb_yield_by_prod_test_type.configure_default_column(value=True, editable=True)
                gb_yield_by_prod_test_type.configure_column("Yield", cellStyle=yield_color)
                gridOptions = gb_yield_by_prod_test_type.build()
                AgGrid(yield_by_prod_test_type, gridOptions=gridOptions, enable_enterprise_modules=True, fit_columns_on_grid_load = True, allow_unsafe_jscode=True)



            full_data_expand = st.expander("Expand for Full Data")
            with full_data_expand:
                st.subheader("Full Data")
                gb_data = GridOptionsBuilder.from_dataframe(full_data)
                # gb_data.configure_pagination(paginationPageSize=20, paginationAutoPageSize=False)
                gb_data.configure_default_column(value=True, editable=True, groupable=True)
                gridOptions = gb_data.build()
                AgGrid(full_data, gridOptions=gridOptions, enable_enterprise_modules=True, allow_unsafe_jscode=True)

                defined_codes = pd.Series(list(family_desc_sheet.values())).explode().tolist()
                undefined_df = st.session_state['df'][st.session_state['df']["Product Tested Description"].isin(defined_codes) == False] # extract rows with undefined codes

            with st.expander("Data with undefined codes"):
                st.write(undefined_df)
