import streamlit as st
import pandas as pd
import time
import datetime
import os
from tkinter import filedialog
import tkinter as tk
import numpy as np
from openpyxl import load_workbook

today=datetime.datetime.today().strftime("%d.%m.%Y")

st.title('HUAWEI')
st.markdown('[Huawei Sharpoint](https://arrowelectronics.sharepoint.com/:f:/r/sites/EMEAProductCatalog/Shared%20Documents/A%20-%20Price%20Catalog/D%20-%20H/HUAWEI?csf=1&web=1&e=mkQgMK/) ')

st.markdown('------------------------------------------------------------------')

col1, col2 = st.columns(2)
with col2:
    data=st.date_input("Pricelist start date", value=datetime.date.today())
    t=str(data.strftime("%d.%m.%Y"))

with col1:
    id=st.selectbox('Your name - ID',('Karol Kwiecień - A86227','Katarzyna Czyż - A86361', 'Emil Twardowski - A93176', 'Chistian Gay - A60276',
                                      'Paweł Czaja - A89264', 'Hanna Źródlewska - 132693', 'Karmen Bautembach - 136182'),index=None,placeholder="Select your name")

    if id is not None:
        # st.write(id[-6:])
        number_id = id[-6:]


# selecting pricelist and folder directory
st.markdown('------------------------------------------------------------------')

def select_file():
    root = tk.Tk()
    root.withdraw()
    folder_path = root.filename = tk.filedialog.askopenfilename(master=root, filetypes=[("Excel files", "*.xlsx")])
    root.destroy()
    return folder_path

selected_folder_path = st.session_state.get("folder_path", None)

folder_select_button = st.button("Select Pricelist file",type='primary')

if folder_select_button:
    selected_folder_path = select_file()
    path = os.path.dirname(selected_folder_path)

    st.session_state.folder_path = selected_folder_path
    st.session_state.path =path

if 'folder_path' not in st.session_state:
    st.write('No file selected')
else:
    st.write(st.session_state['folder_path'])
    st.success('Pricelist file added')
    # st.write(st.session_state['path'])

st.markdown('------------------------------------------------------------------')
# ------------------------------------  TAR  -----------------------------------
st.markdown('### :small_blue_diamond:  TAR file ###')

btn1 = st.button(' Create TAR', type='primary')
if btn1:
    if 'folder_path' not in st.session_state:
        st.error('No pricelist selected')

    else:
        if id is not None:
            number_id = id[-6:]

            # st.write(st.session_state['folder_path'])
            url = st.session_state['folder_path']

            with st.spinner('Prepering file...'):

                pricelist = pd.read_excel(url)
                TAR = pricelist.filter(["PartNumber", 'List Price\n (EUR)', 'Authorization Discount Off','Authorization Unit Price\n(EUR FOB HongKong)'])
            # # #zmiana nazw kolumn
                TAR.rename(columns={'PartNumber': 'SKU', 'List Price\n (EUR)': 'Public Price','Authorization Discount Off': 'StdRebate','Authorization Unit Price\n(EUR FOB HongKong)': 'Channel Price'}, inplace=True)
            # # #dodawanie nowych kolumn i uzupełnianie
                TAR.insert(loc=1, column='AccountCode', value="ALL")
                TAR.insert(loc=2, column='Currency', value="EUR")
                TAR.insert(loc=3, column='Quantity', value="1")
                TAR['Valid From'] = t
                TAR['Valid To'] = ""
                TAR['SearchAgain'] = "YES"
                TAR['UnitID'] = "PCS0DEC"
                TAR['LegalEntity'] = "101-120-141-311-501-502-550-560-580-582-700"
                TAR['StdRebate'] = TAR['StdRebate'] * 100
                TAR['Item Group'] = 'HUAWEI'


                # TAR2=st.dataframe(TAR.set_index("SKU"))
                TAR = TAR[['SKU', 'AccountCode', 'LegalEntity', 'Currency', 'Quantity', 'Public Price', 'StdRebate','Channel Price', 'Valid From', 'Valid To', 'SearchAgain', 'UnitID', 'Item Group']]

                # # Set up SKU number as index
                TAR.set_index('SKU', inplace=True)

                # usuwanie SKu z zerowymi cenami i uzupelnianie brakujacych wartość
                TAR = TAR[TAR['Public Price'] != 0]
                TAR.loc[TAR['Channel Price'] == ' ', 'Channel Price'] = TAR['Public Price']
                TAR['StdRebate'] = TAR['StdRebate'].replace(np.nan, 0)

                TAR['Channel Price'] = TAR['Channel Price'].astype('float64')


                TAR.to_excel(st.session_state['path'] + "/HUAWEI-"+number_id+"-TAR-" + today + "-SKU.xlsx", sheet_name='TAR', startrow=1)

                wbook = load_workbook(st.session_state['path'] + "/HUAWEI-"+number_id+"-TAR-" + today + "-SKU.xlsx")
                sheet = wbook.active
                sheet['A1'] = 'iAssetSync'
                sheet['B1'] = 'No'
                wbook.save(st.session_state['path'] + "/HUAWEI-"+number_id+"-TAR-" + today + "-SKU.xlsx")

                st.write(TAR)

                # st.write(st.session_state['folder_path'])
                # st.write(st.session_state['path'])
                st.success('TAR saved in pricelist file directory')
        else:
            st.error("Select your name -ID")

st.markdown('----------------------------------------------------------------------------------------------------')
# --------------------------------     UPD     -------------------------------------------------------------------
st.markdown('### 	:small_blue_diamond: UPD file ###')
with st.expander(' Data for UPD'):
# subgrups = st.file_uploader(label='Subgrups files',accept_multiple_files=True)

    ekstrakt=st.file_uploader(label=':small_blue_diamond: Extract file', accept_multiple_files=False, type=["csv"])
    if ekstrakt is not None:
        st.success('Extract file added')
    # subgroups=st.file_uploader(label=':small_blue_diamond: Subgroups', accept_multiple_files=True, type=["xlsx"])
    # pricelist = pd.read_csv(ekstrakt)
    # if pricelist is not None:
    #     pl=pd.read_excel(pricelist)
    #     with open(pricelist.name,"wb")as f:
    #         f.write(pricelist.getbuffer())
    #     st.success('Pricelist file added')

btn2 = st.button('Create UPD', type='primary')
if btn2:
        if 'folder_path' not in st.session_state:
            st.error('No pricelist selected')
        else:
            if id is not None:
                number_id = id[-6:]

                # st.write(st.session_state['folder_path'])
                url = st.session_state['folder_path']

                if ekstrakt is not None:
                    with st.status("Prepering UPD", expanded=True) as status:

                        st.write("Comparing Pricelist files with extract")
                        pricelist = pd.read_excel(url)
                        df2 = pd.read_csv(ekstrakt, delimiter=';', usecols=['ItemID','ItemDescription','ItemGroupIDSub1','ItemGroupIDSub2','ItemGroupIDSub3','ItemGroupIDSub4','ItemGroupIDSub5','CustomerEDI','ORIGCOUNTRYREGIONID'])

                        # #usuwanie znakow specjalnych w SKU (regex-[r]-regular expresion) tutaj usuwane puste znaki space, tab, enter [regex= \s]
                        st.write("Removing special characters from description")
                        pricelist['PartNumber'] = pricelist['PartNumber'].str.replace(r'\s', "", regex=True)

                        # # usuwanie przecinków z description
                        pricelist['Description'] = pricelist['Description'].astype(str)
                        pricelist['Description'] = pricelist['Description'].str.replace(',', " ")
                        pricelist['Description'] = pricelist['Description'].str.replace(r'[|]', ' ', regex=True)
                        pricelist['Description'] = pricelist['Description'].str.replace('  ', " ")

                        UPD = pricelist.filter(["PartNumber", "Description", "Software and Hardware Attributes", "Pack Weight\n (kg) ", "Pack Dimension\n (D*W*H mm) ","Net Dimension\n (D*W*H mm) ", "Discount Category", "Product Line", "Product Family", "Sub Product Family"])

                        # #zmiana nazw kolumn
                        UPD.rename(columns={'PartNumber': 'SKU'}, inplace=True)

                        # ----  wymiary ------
                        st.write("Extracting weight and dimension")

                        #   GROSS
                        UPD["Pack Dimension\n (D*W*H mm) "] = UPD["Pack Dimension\n (D*W*H mm) "].str.extract(
                            r'([0-9]{1,5}[*][0-9]{1,5}[*][0-9]{1,5})', expand=True)
                        UPD["Pack Dimension\n (D*W*H mm) "] = UPD["Pack Dimension\n (D*W*H mm) "].str.replace(r'[*]{2}', '*', regex=True)
                        UPD[['Gross width', 'Gross Height', 'Gross Depth']] = UPD["Pack Dimension\n (D*W*H mm) "].str.split(r'[*]', expand=True)

                        UPD['Gross width'] = UPD['Gross width'].fillna(0)
                        UPD['Gross Height'] = UPD['Gross Height'].fillna(0)
                        UPD['Gross Depth'] = UPD['Gross Depth'].fillna(0)
                        UPD = UPD.astype({'Gross width': 'int', 'Gross Height': 'int', 'Gross Depth': 'int'})

                        UPD['Gross width'] = UPD['Gross width'] / 1000
                        UPD['Gross Height'] = UPD['Gross Height'] / 1000
                        UPD['Gross Depth'] = UPD['Gross Depth'] / 1000

                        #  NET
                        UPD["Net Dimension\n (D*W*H mm) "] = UPD["Net Dimension\n (D*W*H mm) "].str.extract(r'([0-9]{1,5}[*][0-9]{1,5}[*][0-9]{1,5})',
                                                                                                            expand=True)
                        UPD["Net Dimension\n (D*W*H mm) "] = UPD["Net Dimension\n (D*W*H mm) "].str.replace(r'[*]{2}', '*', regex=True)
                        UPD[['Net width', 'Net Height', 'Net Depth']] = UPD["Net Dimension\n (D*W*H mm) "].str.split(r'[*]', expand=True)

                        UPD['Net width'] = UPD['Net width'].fillna(0)
                        UPD['Net Height'] = UPD['Net Height'].fillna(0)
                        UPD['Net Depth'] = UPD['Net Depth'].fillna(0)
                        UPD = UPD.astype({'Net width': 'int', 'Net Height': 'int', 'Net Depth': 'int'})

                        UPD['Net width'] = UPD['Net width'] / 1000
                        UPD['Net Height'] = UPD['Net Height'] / 1000
                        UPD['Net Depth'] = UPD['Net Depth'] / 1000

                        UPD.loc[(UPD['Net width'] == 0) & (UPD['Gross width'] != 0), 'Net width'] = UPD['Gross width']
                        UPD.loc[(UPD['Net Height'] == 0) & (UPD['Gross Height'] != 0), 'Net Height'] = UPD['Gross Height']
                        UPD.loc[(UPD['Net Depth'] == 0) & (UPD['Gross Depth'] != 0), 'Net Depth'] = UPD['Gross Depth']

                        UPD.loc[(UPD['Gross width'] == 0) & (UPD['Net width'] != 0), 'Gross width'] = UPD['Net width']
                        UPD.loc[(UPD['Gross Height'] == 0) & (UPD['Net Height'] != 0), 'Gross Height'] = UPD['Net Height']
                        UPD.loc[(UPD['Gross Depth'] == 0) & (UPD['Net Depth'] != 0), 'Gross Depth'] = UPD['Net Depth']

                        UPD["Pack Weight\n (kg) "] = UPD["Pack Weight\n (kg) "].fillna(0)
                        UPD.loc[(UPD["Pack Weight\n (kg) "] != 0) | ((UPD['Gross width'] != 0) & (UPD['Gross Height'] != 0) & (UPD['Gross Depth'] != 0)) | ((UPD['Net width'] != 0) & (UPD['Net Height'] != 0) & (UPD['Net Depth'] != 0)), 'W&D'] = 'HARD'

                        st.write("Categorize Activity 1,2,3")
                        # # ------------  Activity 1 ------------------------------------------------
                        # # HARD
                        UPD.loc[(UPD['W&D'] == 'HARD'), 'Activity 1'] = 'HARD'
                        UPD.loc[(UPD["Software and Hardware Attributes"] == 'Hardware'), 'Activity 1'] = 'HARD'
                        UPD.loc[(UPD["Discount Category"] == 'Hardware'), 'Activity 1'] = 'HARD'

                        UPD.loc[(UPD['Description'].str.contains('support', case=False) & (UPD['W&D'] == 'HARD')), 'Activity 1'] = 'HARD +Support'
                        UPD.loc[(UPD['Description'].str.contains('Power Suply', case=False)) & (~UPD['Description'].str.contains('with', case=False) & (UPD['W&D'] == 'HARD')), 'Activity 1'] = 'HARD Power Suply'

                        # # SOFT
                        UPD.loc[(UPD['Software and Hardware Attributes'] == 'Self-developed software') | (
                        (UPD['Software and Hardware Attributes'] == 'Software Annuity')), 'Activity 1'] = 'SOFT Licences'
                        UPD.loc[(UPD['Description'].str.contains('upgrade', case=False) & (UPD["Activity 1"] == 'SOFT Licences')), 'Activity 1'] = 'SOFT Upgrade'
                        UPD.loc[(UPD['Description'].str.contains('Subscription', case=False) & ( UPD["Activity 1"] == 'SOFT Licences')), 'Activity 1'] = 'SOFT Subscription'
                        UPD.loc[(UPD['Description'].str.contains('Subscription', case=False) & (UPD['Description'].str.contains('Support', case=False)) & (UPD["Activity 1"] == 'SOFT Subscription')), 'Activity 1'] = 'SOFT Subscript +Support'
                        UPD.loc[UPD['Discount Category'].str.contains("License", case=False, na=False), 'Activity 1'] = 'SOFT Licences'

                        # UPD.loc[(UPD['Description'].str.contains('subscript',case=False)&(UPD["Activity 1"]=='SOFT Licences')),
                        # 'Activity 1']='SOFT Subscription'
                        # UPD.loc[(UPD['Description'].str.contains('secur',case=False)&(UPD["Activity 1"]=='SOFT Licences')),'Activity 1']='SOFT
                        # SECURE'

                        # # SERVICE
                        UPD.loc[UPD["Software and Hardware Attributes"] == 'Service', 'Activity 1'] = 'SERVICE'
                        UPD.loc[((UPD['Discount Category'] == 'Outsourcing') & (UPD['W&D'] != 'HARD') & (UPD['Description'].str.contains('service', case=False))), 'Activity 1'] = 'SERVICE'
                        UPD.loc[(UPD['Description'].str.contains('Data visualization service', case=False)), 'Activity 1'] = 'SERVICE'
                        UPD.loc[(UPD['Description'].str.contains('Security Service', case=False) & (UPD['Description'].str.contains('Yearly', case=False))), 'Activity 1'] = 'SERVICE'
                        UPD.loc[UPD['Description'].str.match(r'(^Security Service)', case=False, na=False), ['Activity 1']] = 'SERVICE'
                        UPD.loc[UPD['Activity 1'].isnull(), 'Activity 1'] = 'SERVICE'

                        # UPD.loc[(UPD["Method of Delivery"]=='Electronic') & ((UPD['Description'].str.contains('service',case=False)) | (UPD[
                        # 'Description'].str.contains('support',case=False))),'Activity 1']='SERVICE'

                        # # ------------     Activity 2    --------------

                        UPD.loc[UPD['Activity 1'] == 'HARD', 'Activity 2'] = 'OTHERS'
                        UPD.loc[((UPD['Activity 1'] == 'HARD') & (UPD['Description'].str.contains('server', case=False))), 'Activity 2'] = 'SERVER'
                        UPD.loc[((UPD['Activity 1'] == 'HARD') & (UPD['Description'].str.contains('port', case=False))), 'Activity 2'] = 'SERVER'
                        UPD.loc[((UPD['Activity 1'] == 'HARD') & (UPD['Description'].str.contains('unit', case=False))), 'Activity 2'] = 'SERVER'
                        UPD.loc[((UPD['Activity 1'] == 'HARD') & (UPD['Description'].str.contains('SSD', case=False))), 'Activity 2'] = 'STORAGE'
                        UPD.loc[((UPD['Activity 1'] == 'HARD') & (UPD['Description'].str.contains('HDD', case=False))), 'Activity 2'] = 'STORAGE'
                        UPD.loc[UPD['Activity 1'] == 'HARD +Support', 'Activity 2'] = 'OTHERS'
                        UPD.loc[UPD['Activity 1'] == 'SOFT Licences', 'Activity 2'] = 'OTHERS'
                        UPD.loc[UPD['Activity 1'] == 'SOFT Upgrade', 'Activity 2'] = 'OTHERS'
                        UPD.loc[UPD['Activity 1'] == 'SOFT SECURE', 'Activity 2'] = 'SECURITY'
                        UPD.loc[UPD['Activity 1'] == 'SOFT Subscription', 'Activity 2'] = 'OTHERS'
                        UPD.loc[UPD['Activity 1'] == 'SOFT Subscript +Support', 'Activity 2'] = 'OTHERS'
                        UPD.loc[UPD['Activity 1'] == 'SERVICE', 'Activity 2'] = 'MAINT'
                        UPD.loc[(UPD['Activity 1'] == 'SERVICE') & (UPD['Discount Category'].str.contains('Training', case=False)) & (
                            UPD['Product Family'].str.contains('Training', case=False)), 'Activity 2'] = 'TRAINING'

                        # # UPD.loc[(UPD['Activity 1']=='SOFT Subscription')&((UPD['Description'].str.contains('secur',case=False))'Activity
                        # 2']='OTHERS'
                        # UPD.loc[(UPD['Description'].str.contains('secur',case=False)&(UPD["Activity 1"]=='SOFT Subscription')),
                        # 'Activity 2']='SECURITY'
                        # UPD.loc[UPD['Activity 1']=='SERVICE','Activity 2']='OTHERS'

                        # # ------------     Activity 3    ---------------

                        UPD.loc[(UPD["Description"].str.contains('renewal', case=False)) & (UPD['Activity 1'] != 'HARD'), 'Activity 3'] = 'RENEWAL'
                        UPD.loc[UPD['Activity 3'] != 'RENEWAL', 'Activity 3'] = 'INITIAL'

                        # #dodawanie nowych kolumn i uzupełnianie
                        UPD.insert(loc=0, column='Item Group', value="HUAWEI")
                        UPD.insert(loc=2, column="Vendor SKU", value="")
                        UPD.insert(loc=3, column="Item Type", value='Item')
                        UPD['Inventory Model Group'] = 'FIFOARW03'
                        UPD['Life Cycle'] = 'Online'
                        UPD['Stock Management'] = 'BACK TO BACK'
                        UPD['Finance Project Category'] = UPD['Item Group']
                        UPD['ItemPrimaryVendId'] = ''
                        UPD['Volume'] = ''
                        UPD['Legacy Id'] = ''
                        UPD['Customer EDI'] = 'YES'
                        UPD['List Price UpDate'] = 'YES'
                        UPD['Dual Use'] = 'YES'
                        UPD['Virtual Item'] = 'NO'
                        UPD['Arrow Brand'] = 'HUA'
                        UPD['Purchase Delivery Time'] = ''
                        UPD['Sales Delivery Time'] = ''
                        UPD['Production Type'] = 'NONE'
                        UPD['Unit point'] = ''
                        UPD['Warranty'] = ''
                        UPD['Renewal term'] = ''
                        UPD['Origin'] = 'CHN'
                        UPD.rename(columns={"Pack Weight\n (kg) ": "Weight"}, inplace=True)
                        UPD['Tare Weight'] = UPD['Weight'] * 0.2152
                        UPD['Finance Activity'] = UPD['Activity 1']
                        UPD['Special marker']=''
                        UPD['Special VAT Code']=''


                        st.write("Adding subgroup SUBGRUPY 1 according files")
                        # -----   przypisywanie SUBGRUPY 1
                        sub1 = pd.read_excel(st.session_state['path'] + "/subgrup1.xlsx", usecols=['Sub group 1', 'Description'])
                        sub1.rename(columns={'Description': 'Description_sub1', 'Sub group 1': 'ItemGroupIDSub1'}, inplace=True)
                        UPD = pd.merge(UPD, sub1, left_on='Software and Hardware Attributes', right_on='Description_sub1', how='left')
                        UPD.loc[UPD['Software and Hardware Attributes'] == ' ', 'ItemGroupIDSub1'] = '####'

                        st.write("Adding subgroup SUBGRUPY 2 according files")
                        # ------   przypisywanie SUBGRUPY 2
                        sub2 = pd.read_excel(st.session_state['path'] + "/subgrup2.xlsx", usecols=['Sub group 2', 'Description'])
                        sub2.rename(columns={'Description': 'Description_sub2', 'Sub group 2': 'ItemGroupIDSub2'}, inplace=True)
                        # #usuwanie znakow specjalnych na końcu i na początku
                        UPD['Product Family'] = UPD['Product Family'].str.strip()
                        UPD = pd.merge(UPD, sub2, left_on='Product Family', right_on='Description_sub2', how='left')
                        UPD.loc[UPD['Product Family'] == ' ', 'ItemGroupIDSub2'] = '####'

                        st.write("Adding subgroup SUBGRUPY 3 according files")
                        # ------   przypisywanie SUBGRUPY 3
                        sub3 = pd.read_excel(st.session_state['path'] + "/subgrup3.xlsx", usecols=['Sub group 3', 'Description'])
                        sub3.rename(columns={'Description': 'Description_sub3', 'Sub group 3': 'ItemGroupIDSub3'}, inplace=True)
                        # #usuwanie znakow specjalnych na końcu i na początku
                        UPD['Product Line'] = UPD['Product Line'].str.strip()
                        UPD = pd.merge(UPD, sub3, left_on='Product Line', right_on='Description_sub3', how='left')
                        UPD.loc[UPD['Product Line'] == ' ', 'ItemGroupIDSub3'] = '####'
                        UPD.loc[UPD['Product Line'] == '', 'ItemGroupIDSub3'] = '####'

                        st.write("Adding subgroup SUBGRUPY 4 according files")
                        # ------   przypisywanie SUBGRUPY 4
                        sub4 = pd.read_excel(st.session_state['path'] + "/subgrup4.xlsx", usecols=['Sub group 4', 'Description'])
                        sub4.rename(columns={'Description': 'Description_sub4', 'Sub group 4': 'ItemGroupIDSub4'}, inplace=True)
                        sub4['Description_sub4'] = sub4['Description_sub4'].str.lower()
                        # #usuwanie znakow specjalnych na końcu i na początku
                        UPD['Discount Category'] = UPD['Discount Category'].str.strip().str.lower()
                        UPD = pd.merge(UPD, sub4, left_on='Discount Category', right_on='Description_sub4', how='left')
                        UPD.loc[UPD['Discount Category'] == ' ', 'ItemGroupIDSub4'] = '####'
                        UPD.loc[UPD['Discount Category'].isnull(), 'ItemGroupIDSub4'] = '####'

                        st.write("Adding subgroup SUBGRUPY 5 according files")
                        # ------   przypisywanie SUBGRUPY 5
                        sub5 = pd.read_excel(st.session_state['path'] + "/subgrup5.xlsx", usecols=['Sub group 5', 'Description'])
                        sub5.rename(columns={'Description': 'Description_sub5', 'Sub group 5': 'ItemGroupIDSub5'}, inplace=True)
                        # #usuwanie znakow specjalnych na końcu i na początku
                        UPD['Sub Product Family'] = UPD['Sub Product Family'].str.strip()
                        UPD = pd.merge(UPD, sub5, left_on='Sub Product Family', right_on='Description_sub5', how='left')
                        UPD.loc[UPD['Sub Product Family'] == ' ', 'ItemGroupIDSub5'] = '####'
                        UPD.loc[UPD['Sub Product Family'] == '', 'ItemGroupIDSub5'] = '####'
                        UPD.loc[UPD['Sub Product Family'].isnull(), 'ItemGroupIDSub5'] = '####'

                        # pobiera dane GvN z App.py do gv1
                        gn1 = st.session_state['gvn'][['merge', 'Intrastat Code', 'Gross/Net Classification', 'Gross/Net', 'SUBBRAND']]

                        UPD['merge'] = UPD['Activity 1'] + UPD['Activity 2'] + UPD['Activity 3']

                        UPD2 = UPD.merge(gn1, on="merge", how='left')

                        st.write("Adding Special VAT Code")
                        ###  Special VAT Code
                        UPD2.loc[(UPD2['Activity 1'] == 'SERVICE') & (UPD2['Activity 2'] == 'MAINT') & (UPD2['Description'].str.contains('HARD' or 'warranty', case=False)), 'Special VAT Code'] = 'SPVAT0001'

                        # ### Dimension Group
                        UPD2["Dimension Group"] = "PHYSICAL"
                        UPD2.loc[UPD2["Intrastat Code"] == "00000000", "Dimension Group"] = "STDBATCH3"

                        #### Serial Number Group
                        UPD2['Serial Number Group'] = "SN-AECS"
                        serial = {'00000000': "SN-AECS", '85444210': "SN-AECS", '85044030': "SN-AECS", }
                        UPD2['Serial Number Group'] = UPD2["Intrastat Code"].map(serial)

                        UPD2.loc[UPD2['Intrastat Code'] == '00000000', 'Virtual Item'] = 'YES'
                        UPD2['Activity 1'] = UPD2['Activity 1'].str.extract(r'(HARD|SOFT|SERVICE)')
                        UPD2['Finance Activity'] = UPD2['Finance Activity'].str.extract(r'(HARD|SOFT|SERVICE)')

                        st.write("Adding Special marker")
                        ###  Special marker
                        UPD2.loc[((UPD2['Activity 1'] == 'SERVICE') & (UPD2['Activity 2'] == 'TRAINING')) & (~UPD2['Description'].str.contains('|'.join(['fee', 'Travel costs', 'Travel expenses', 'Conference', 'event', 'training materials', 'books']),case=False)), 'Special marker'] = "GTU_12"
                        UPD2.loc[(UPD2['Activity 1'] == 'HARD'), 'Special marker'] = "MPP_GTU_06"

                        # #porzadkowanie kolumn
                        UPD2 = UPD2[['Item Group', 'SKU', 'Vendor SKU', 'Item Type', 'Intrastat Code', 'Dimension Group', 'Serial Number Group','Inventory Model Group', 'Life Cycle', 'Activity 1', 'Activity 2', 'Activity 3', 'Stock Management', 'ItemGroupIDSub1', 'ItemGroupIDSub2','ItemGroupIDSub3', 'ItemGroupIDSub4','ItemGroupIDSub5', 'Description', 'ItemPrimaryVendId', 'Weight', 'Tare Weight', 'Gross width', 'Gross Height','Gross Depth', 'Net width', 'Net Height', 'Net Depth', 'Volume', 'Legacy Id', 'Finance Project Category','Finance Activity', 'Customer EDI', 'List Price UpDate', 'Dual Use', 'Virtual Item', 'Arrow Brand', 'Gross/Net','Gross/Net Classification','Purchase Delivery Time', 'Sales Delivery Time', 'Production Type', 'Unit point', 'Warranty', 'SUBBRAND','Renewal term', 'Special VAT Code', 'Origin', 'Special marker']]

                        roznica = UPD2.loc[~UPD2['SKU'].isin(df2['ItemID'])]


                    status.update(label="UPD complete!", state="complete", expanded=False)

                    if not roznica.empty:

                        row_count = str(len(roznica))
                        st.write(':point_right:  '+ row_count + " new items to create")
                        roznica.to_excel(st.session_state['path'] + "/HUAWEI-" + number_id + "-UPD-" + today + "-SKU.xlsx", sheet_name='UPD',startrow=1, index=False)
                        st.success('UPD saved in pricelist file directory')
                    else:
                        st.info('No new items to create')
                else:
                    st.error("Extract file NOT selected")
            else:
                st.error("Select your name -ID")

st.markdown('----------------------------------------------------------------------------------------------------')
# --------------------------------     AMD     -------------------------------------------------------------------
st.markdown('### :small_blue_diamond:  AMD file ###')
st.write('EDI, COO, subgroups, description, online - offline correction')

btn1 = st.button(' Create AMD', type='primary')
