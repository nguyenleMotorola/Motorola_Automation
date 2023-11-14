
import pandas as pd
import streamlit as st
import httplib2
import os
import xml.etree.ElementTree as ET
from googleapiclient import discovery
from google.oauth2 import service_account
import xlsxwriter
from io import BytesIO

st.set_page_config(
    page_title="Inventory Discrepancy",
    layout="wide"
)

# cache the dataframe to memory so we don't have to read it over and over again
@st.cache(suppress_st_warning=True)
def get_data(file_name): 
    data = pd.read_excel(file_name)
    return data

# ------ Main page ------
st.title("🔥 Inventory Discrepancy Analysis")
st.markdown("##")

st.sidebar.header('Drag and drop your 3PL and SAP Snapshot Files here')
uploaded_files = st.sidebar.file_uploader('', type='xlsx', key=5, accept_multiple_files=True)


if uploaded_files:

    # ==== CREATE EXCEL RESULT FILE ====
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine = 'xlsxwriter')
    
    # =================== GET DATA, SAP, 3PL, NON-INVENTORY ===================
    for file in uploaded_files:
        if "SAP" in file.name:   # read in SAP data file
            SAP_data = get_data(file) 
            pre_SAP_data = SAP_data
            pre_SAP_data.to_excel(writer, sheet_name='SAP Data', index=False)
            SAP_data.columns = SAP_data.columns.str.replace('[#,@,&,:]','', regex=True) # remove special characters from column names
            SAP_data.columns = SAP_data.columns.str.lower()                  # lowercase column names
        elif "3PL" in file.name: # read in 3PL data file
            PL_data = get_data(file) 
            pre_PL_data = PL_data
            pre_PL_data.to_excel(writer, sheet_name='3PL Data', index=False)
            PL_data.columns = PL_data.columns.str.replace('[#,@,&,:]','', regex=True) # remove special characters from column names
            PL_data.columns = PL_data.columns.str.lower()                  # lowercase column names
        # elif "Non-Inventory" in file.name:
        #     non_inventory_selector = {'SU': 'storage_unit'}
        #     non_inventory_data = pd.read_excel(file, sheet_name="OFFSITE Non-Inventory")
        #     non_inventory_su = non_inventory_data.rename(columns=non_inventory_selector)[[*non_inventory_selector.values()]]
    try:
        non_inventory_data = []
        scopes = ["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/spreadsheets"]
        secret_file = os.path.join(os.getcwd(), 'client_secret.json')

        SPREADSHEET_ID = '16OzdGvKSDwbkG5F_fZVxFjB4fdZuoD1krGdU_iMZErU'
        RANGE = "'OFFSITE Non-Inventory'!A2:A" 
        credentials = service_account.Credentials.from_service_account_file(secret_file, scopes=scopes)
        service = discovery.build('sheets', 'v4', credentials=credentials)

        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE).execute()
        non_inventory_data = result.get('values', [])

        if not non_inventory_data:
            print('No data found')
        else:
            non_inventory_data = pd.DataFrame(non_inventory_data, columns = ['storage_unit'])
            # print(type(non_inventory_data['storage_unit']))
            # non_inventory_data = [j for sub in non_inventory_data for j in sub]
            # for row in non_inventory_data:
            #     print(row[0])

        # ============ WRITE TO GOOGLE SHEETS ============
        # SAP_Warehouse_results = []
        # data = {
        #     'values' : SAP_Warehouse_results # in 2d array
        # }
        # sheet.values().update(spreadsheetId=spreadsheet_id, body=data, range=range_name, valueInputOption='USER_ENTERED').execute() 

    except OSError as e:
        print(e)

    # =========================== ANALYSIS ============================
    # Compare SAP storage unit codes against 3P column C (HAWB)
    # SAP Material (column C) against 3PL SKU (Column B)
    # SAP Available Stock (Column E) against 3PL Moveable Unit Label (Column P)
    SAP_selector = {'storage unit':'storage_unit', 'material':'material', 'available stock':'available_stock'} # select columns from SAP data and rename
    SAP_df = SAP_data.rename(columns=SAP_selector)[[*SAP_selector.values()]]

    PL_selector = {'hawb':'storage_unit', 'sku':'material', 'movable unit label':'available_stock'} # select columns from 3Pl data and rename so that it matches with the corresponding columns in SAP
    PL_df = PL_data.rename(columns=PL_selector)[[*PL_selector.values()]]
    PL_codes = PL_df['storage_unit'].apply(str) # get 3PL storage unit codes
    discrep_df = SAP_df.merge(PL_df, indicator=True, how='outer')

    # ===== COMPARE BY STORAGE UNIT CODES ONLY =====
    header = ["Both", "SAP Only", ""]
    st.subheader("Matching Storage Unit Codes") 
    matching_codes_compare_df = SAP_df['storage_unit'].to_frame().merge(PL_df['storage_unit'].to_frame(), indicator=True, how='outer')
    matching_codes_df = matching_codes_compare_df.loc[lambda x:x['_merge'] == 'both'] 
    st.write(matching_codes_df.astype('object'))
    st.write(len(matching_codes_df))

    st.subheader("SAP Only Codes") 
    sap_only_codes_df = matching_codes_compare_df.loc[lambda x:x['_merge'] == 'left_only'] 
    st.write(sap_only_codes_df.astype('object'))
    st.write(len(sap_only_codes_df))

    st.subheader("3PL Only Codes") 
    pl_only_codes_df = matching_codes_compare_df.loc[lambda x:x['_merge'] == 'right_only'] 
    st.write(pl_only_codes_df.astype('object'))
    st.write(len(pl_only_codes_df))

    st.subheader("Matching Storage Unit Codes and Material Number") 
    matching_codes_material_compare_df = SAP_df[['storage_unit', 'material']].merge(PL_df[['storage_unit', 'material']], indicator=True, how='outer')
    matching_codes_material_compare_df = matching_codes_material_compare_df.loc[lambda x:x['_merge'] == 'both'] 
    st.write(matching_codes_material_compare_df.astype('object'))
    st.write(len(matching_codes_material_compare_df))

    st.subheader("SAP Only Code and Material Number") 
    sap_only_codes_material_df = matching_codes_material_compare_df.loc[lambda x:x['_merge'] == 'left_only'] 
    st.write(sap_only_codes_material_df.astype('object'))
    st.write(len(sap_only_codes_material_df))

    st.subheader("3PL Only Code and Material Number") 
    pl_only_codes_material_df = matching_codes_material_compare_df.loc[lambda x:x['_merge'] == 'right_only'] 
    st.write(pl_only_codes_material_df.astype('object'))
    st.write(len(pl_only_codes_material_df))

    st.header("SAP full data")
    st.write(SAP_df)
    st.write(len(SAP_df))

    st.header("3PL full data")
    st.write(PL_df.astype('object'))
    st.write(len(PL_df))

    st.subheader("Matching Storage Unit Codes and Material Number and Quantity") 
    both_df = discrep_df.loc[lambda x: x['_merge'] == 'both']
    st.write(both_df.astype('object'))
    st.write(len(both_df))

    st.subheader("SAP Only Storage Unit Codes and Material Number and Quantity") 
    sap_only_df = discrep_df.loc[lambda x: x['_merge'] == 'left_only']
    st.write(sap_only_df.astype('object'))
    st.write(len(sap_only_df))
    sap_only_df = sap_only_df.drop(columns=['_merge'])
    sap_only_df.to_excel(writer, sheet_name='SAP ONLY', index=False)

    st.subheader("3PL Only Storage Unit Codes and Material Number and Quantity") 
    pl_only_df = discrep_df.loc[lambda x: x['_merge'] == 'right_only']
    st.write(pl_only_df.astype('object'))
    st.write(len(pl_only_df))
    pl_only_df = pl_only_df.drop(columns=['_merge'])
    pl_only_df.to_excel(writer, sheet_name='3PL ONLY', index=False)

    st.subheader("Discrepancies between 3PL and SAP") # show the differences in available stock quantity between both data 
    merge_selector = {'storage_unit':'storage_unit', 'material':'material', 'available_stock_x':'sap_available_stock', 'available_stock_y':'3pl_available_stock'}
    df_merged = sap_only_df.merge(pl_only_df, on=['storage_unit', 'material'])
    df_merged = df_merged.rename(columns=merge_selector)[[*merge_selector.values()]]
    st.write(df_merged.astype('object'))
    st.write(len(df_merged))
    df_merged.to_excel(writer, sheet_name='3PL vs SAP', index=False)

    st.subheader("3PL FIXED")
    st.write("3PL moveable unit label is replaced with the corresponding SAP available stock values")
    pl_fixed = PL_data.merge(SAP_df, left_on=['sku', 'hawb'], right_on=['material', 'storage_unit'], how='left') 
    pl_fixed = pl_fixed.drop(columns = ['movable unit label', 'storage_unit', 'material'])
    pl_fixed = pl_fixed.rename(columns={"available_stock": "moveable unit label FIXED"})
    pl_fixed.to_excel(writer, sheet_name="3PL FIXED", index=False)
    st.write(pl_fixed.astype('object'))


    # st.subheader("Non Inventory Data (Storage Unit Codes only)") # list not dataframe
    # # non_inventory_data = pd.Series( (v[0] for v in non_inventory_data), name="storage_unit").to_frame()
    # st.write(non_inventory_data)

    st.subheader("Non Inventory Codes that are not in 3PL")
    df = non_inventory_data.merge(PL_codes, indicator=True, how='outer')
    non_inventory_codes_only = df.loc[lambda x:x['_merge'] == 'left_only']
    non_inventory_codes_only = non_inventory_codes_only.drop(columns=['_merge'])
    st.write(non_inventory_codes_only.astype('object'))
    st.write(len(non_inventory_codes_only))
    non_inventory_codes_only.to_excel(writer, sheet_name='NI Codes not in 3PL', index=False)

    st.subheader("3PL Codes not in Non Inventory")
    df2 = df.loc[lambda x:x['_merge'] == 'right_only']
    df2 = df2.drop(columns=['_merge'])
    st.write(df2.astype('object'))
    df2.to_excel(writer, sheet_name='3PL Codes not in NI', index=False)

    writer.save()
    excel_data = output.getvalue()
    st.download_button(label="Export data to Excel", file_name='Inventory_Discrepancy.xlsx', data=excel_data)

    