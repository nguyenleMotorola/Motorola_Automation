import cx_Oracle
import pandas as pd
import streamlit as st
from io import BytesIO


st.set_page_config(
    page_title="Display Bin Audit",
    layout="wide"
)

# cache the dataframe to memory so we don't have to read it over and over again
# Create a dictionary whose key is the sheet name and values are the codes in that sheet name
# @st.cache()
# def get_data_from_excel(file_name):
def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'}) 
    worksheet.set_column('A:A', None, format1)  
    writer.save()
    processed_data = output.getvalue()
    return processed_data

if __name__ == '__main__':
    st.title(":bar_chart: Display Bin Audit")
    st.markdown("##")


    # ------ USER INTERFACE ------
    # ------ Home page ------
    conStr = 'valordfmprd/oracl3@il01dbpn3:1521/VALORORA'
    conn = cx_Oracle.connect(conStr)
    cur = conn.cursor()

    query = """SELECT WHEN, PERSON, PART, QTY, BIN, STATUS FROM INVENTORY_BIN_DATA"""

    cur.execute(query)
    bin_audit_list = cur.fetchall()
    bin_audit_df = pd.DataFrame(bin_audit_list)
    bin_audit_df.columns = ['Date', 'Person', 'Part', 'QTY', 'Bin', 'Status']    
    st.dataframe(bin_audit_df)
    # do something like fetch, insert etc.
    cur.close()
    conn.close()

    df_xlsx = to_excel(bin_audit_df)
    st.download_button(label='📥 Export to Excel',
                                    data=df_xlsx ,
                                    file_name= 'Bin_Audit_Data.xlsx')

 