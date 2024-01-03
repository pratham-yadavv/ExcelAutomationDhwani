import openpyxl as xl
import pandas as pd
import numpy as np
from datetime import datetime


_3A_path = r'C:/Pratham/python/ExcelAutomation/_3A_rabi.xlsx'
_3B_path = r'C:/Pratham/python/ExcelAutomation/_3B_rabi.xlsx'
_3C_path = r'C:/Pratham/python/ExcelAutomation/_3C_rabi.xlsx'
_4A_path = r'C:/Pratham/python/ExcelAutomation/_4A_rabi.xlsx'
_4B_path = r'C:/Pratham/python/ExcelAutomation/_4B_rabi.xlsx'
#Taking dates from user
start_week_date = input('Enter week start date in DD-MM-YYY format : ')
end_week_date = input('Enter week end date in DD-MM-YYYY format : ')
start_season_date = input('Enter season start date in DD-MM-YYYY format : ')



#StartProgram
def convert_date(date_obj):
    date_str=str(date_obj)
    try:
        # Try parsing the date using the first format ("%d-%m-%Y %H:%M:%S")
        date_a = datetime.strptime(date_str, "%d-%m-%Y %H:%M:%S")
    except ValueError:
        try:
            # If parsing fails, try the second format ("%d/%m/%Y %H:%M")
            date_a = datetime.strptime(date_str, "%Y-%d-%m %H:%M:%S")
        except ValueError:
            return None
    return date_a.strftime('%d-%m-%Y')

def splitter(count_obj) :
    count = len(count_obj.split("||"))
    return count



DMS = r'C:/Pratham/python/ExcelAutomation/Rabi 23-24 Demo Plot & DMS user report & Program dashboard template.xlsx'
workbook = xl.load_workbook(DMS)
sheet = workbook['2.DMS-UserSummary']
sheet2 = workbook['3.Program Dashboard']
sheet3 = workbook['1.DemoPlot-Summary']



list1 = [_3A_path,_3B_path,_3C_path,_4A_path,_4B_path]
list2 = ['CIPT II','Srijan','Viksat','Pani','Pradan','SSP','SIED']
test_data = ['CIPT IITest','SRIJAN Test','VIKSAT Test','PANI Test','PRADAN Test','SSP Test','SIED Test']

list_season_3A_partner = []
list_week_3A_partner = []
list_season_3B_partner = []
list_week_3B_partner = []
list_season_3C_partner = []
list_week_3C_partner = []



for i in list1:
    df = pd.read_excel(i) #Load Data
    df['Server Synced At']=df['Server Synced At'].astype(str)
    df['Server Synced At'] = pd.to_datetime(df['Server Synced At'].apply(convert_date),format='%d-%m-%Y')
    df['Server Synced At'] = df['Server Synced At'].dt.to_pydatetime()
    
    #converting_dateOBj_to_Datetime
    start_date_week = pd.to_datetime(start_week_date,format='%d-%m-%Y')
    end_date_week = pd.to_datetime(end_week_date,format='%d-%m-%Y')
    start_date_season = pd.to_datetime(start_season_date,format='%d-%m-%Y')

    #WEEK_VALUE
    week_df = df[(df['Server Synced At']>=start_date_week)&(df['Server Synced At']<=end_date_week)]
    #SEASON_VALUE
    season_df = df[(df['Server Synced At']>=start_date_season)&(df['Server Synced At']<=end_date_week)]
    # season_value = len(season_df)
    
    if i == _3A_path:
        _3A_list_week = []
        _3A_list_season=[]

        
        season_df = season_df.drop(season_df[season_df['Surveyor Name'].isin(test_data)].index)

        for x in list2:
            count_df_3A_week = week_df[week_df['partner']== x]
            df_3aW = count_df_3A_week[['Surveyor Id']]
            list_week_3A_partner.append(df_3aW)

            count_df_3A_season = season_df[season_df['partner']== x]
            df_3aS = count_df_3A_season[['Surveyor Id']]
            list_season_3A_partner.append(df_3aS)

        for x in list2:
            count_PARTNER_week = week_df[week_df['partner']== x].shape[0]
            _3A_list_week.append(count_PARTNER_week)

        for x in list2:
            count_PARTNER_season = season_df[season_df['partner']== x].shape[0]
            _3A_list_season.append(count_PARTNER_season)
        

        # cell_01 = sheet['C10']
        # cell_02 = sheet2['C15'] 
        # value_cipt = cell_01.value
        # cell_01.value = (int(value_cipt) + int(_3A_list_week[0])) 
        # cell_02.value = (int(value_cipt) + int(_3A_list_week[0]))

        # cell_03 = sheet['D10']
        # cell_04 = sheet2['D15'] 
        # value_srijan = cell_03.value
        # cell_03.value = (int(value_srijan) + int(_3A_list_week[1])) 
        # cell_04.value = (int(value_srijan) + int(_3A_list_week[1]))

        # cell_05 = sheet['E10']
        # cell_06 = sheet2['E15'] 
        # value_viksat = cell_05.value
        # cell_05.value = (int(value_viksat) + int(_3A_list_week[2]))
        # cell_06.value = (int(value_viksat) + int(_3A_list_week[2]))

        # cell_07 = sheet['F10']
        # cell_08 = sheet2['F15'] 
        # value_pani = cell_07.value
        # cell_07.value = (int(value_pani) + int(_3A_list_week[3])) 
        # cell_08.value = (int(value_pani) + int(_3A_list_week[3]))

        # cell_09 = sheet['G10']
        # cell_010 = sheet2['G15'] 
        # value_pradan = cell_09.value
        # cell_09.value = (int(value_pradan) + int(_3A_list_week[4])) 
        # cell_010.value = (int(value_pradan) + int(_3A_list_week[4]))

        # cell_011 = sheet['H10']
        # cell_012 = sheet2['H15'] 
        # value_ssp = cell_011.value
        # cell_011.value = (int(value_ssp) + int(_3A_list_week[5])) 
        # cell_012.value = (int(value_ssp) + int(_3A_list_week[5]))


            
        
        cell_1 = sheet['C8']
        cell_1.value = _3A_list_week[0]
        cell_2 = sheet['D8']
        cell_2.value = _3A_list_week[1]
        cell_3 = sheet['E8']
        cell_3.value = _3A_list_week[2]
        cell_4 = sheet['F8']
        cell_4.value = _3A_list_week[3]
        cell_5 = sheet['G8']
        cell_5.value = _3A_list_week[4]
        cell_6 = sheet['H8']
        cell_6.value = _3A_list_week[5]
        cell_6I = sheet['I8']
        cell_6I.value = _3A_list_week[6]

        cell_7 = sheet['C9']
        cell_7.value = _3A_list_season[0]
        cell_8 = sheet['D9']
        cell_8.value = _3A_list_season[1]
        cell_9 = sheet['E9']
        cell_9.value = _3A_list_season[2]
        cell_10 = sheet['F9']
        cell_10.value = _3A_list_season[3]
        cell_11 = sheet['G9']
        cell_11.value = _3A_list_season[4]
        cell_12 = sheet['H9']
        cell_12.value = _3A_list_season[5]
        cell_12I = sheet['I9']
        cell_12I.value = _3A_list_season[6]

        cell_92 = sheet2['C13']              #3.PROGRAM DASHBOARD
        cell_92.value = _3A_list_season[0]
        cell_93 = sheet2['D13']
        cell_93.value = _3A_list_season[1]
        cell_94 = sheet2['E13']
        cell_94.value = _3A_list_season[2]
        cell_95 = sheet2['F13']
        cell_95.value = _3A_list_season[3]
        cell_96 = sheet2['G13']
        cell_96.value = _3A_list_season[4]
        cell_97 = sheet2['H13']
        cell_97.value = _3A_list_season[5]
        cell_97I = sheet2['I13']
        cell_97I.value = _3A_list_season[6]

        print("3A run successfully")
        # +++++++++++++++++++++++++++++++++++++++++++++++++++++
    elif i == _3B_path:
        _3B_list_week =[]
        _3B_list_season=[]
        _3B_list_week_unique =[]
        _3B_list_season_unique=[]
        _3B_list_season_crop =[]
        _3B_partner_sum = []

        
        season_df = season_df.drop(season_df[season_df['Surveyor Name'].isin(test_data)].index)
        
        for x in list2:
            count_df_3B_week = week_df[week_df['partner']== x]
            df_3bW = count_df_3B_week[['Surveyor Id']]
            list_week_3B_partner.append(df_3bW)

            count_df_3B_season = season_df[season_df['partner']== x]
            df_3bS = count_df_3B_season[['Surveyor Id']]
            list_season_3B_partner.append(df_3bS)


        for x in list2:
            count_PARTNER_week = week_df[week_df['partner']== x].shape[0]
            _3B_list_week.append(count_PARTNER_week)


        for x in list2:
            count_PARTNER_season = season_df[season_df['partner']== x].shape[0]
            _3B_list_season.append(count_PARTNER_season)

        for x in list2:
            PARTNER_week = week_df[week_df['partner']== x]
            _unique_week_value_PARTNER = PARTNER_week['Generate Unique ID for Farmer'].nunique()
            _3B_list_week_unique.append(_unique_week_value_PARTNER)

        for x in list2:
            PARTNER_season = season_df[season_df['partner']== x]
            _unique_season_value_PARTNER = PARTNER_season['Generate Unique ID for Farmer'].nunique()
            _3B_list_season_unique.append(_unique_season_value_PARTNER)
            total_sum = PARTNER_season['In how much area are you adopting new and improved farming practices for this Crop?'].sum()
            
            if x == 'CIPT II':
                total_sum = total_sum*0.4
                _3B_partner_sum.append(total_sum)
            elif x == 'Srijan':
                total_sum = total_sum*0.16
                _3B_partner_sum.append(total_sum)
            elif x == 'Viksat':
                total_sum = total_sum*0.01
                _3B_partner_sum.append(total_sum)
            elif x=='Pani':
                total_sum = total_sum*0.08
                _3B_partner_sum.append(total_sum)
            elif x == 'Pradan':
                total_sum = total_sum*0.13
                _3B_partner_sum.append(total_sum)
            elif x == 'SSP':
                total_sum = total_sum*0.4
                _3B_partner_sum.append(total_sum)
            elif x == 'SIED':
                total_sum = total_sum*0.4
                _3B_partner_sum.append(total_sum)

  

        for x in list2:
            count_PARTNER_season = season_df[(season_df['partner']== x)&(season_df['Is this Crop Card - Programme Plot? ']=='Yes')].shape[0]
            _3B_list_season_crop.append(count_PARTNER_season)

        
        #data filling for unique week values
        cell_13 = sheet['C12']
        cell_13.value = _3B_list_week_unique[0]
        cell_14 = sheet['D12']
        cell_14.value = _3B_list_week_unique[1]
        cell_15 = sheet['E12']
        cell_15.value = _3B_list_week_unique[2]
        cell_16 = sheet['F12']
        cell_16.value = _3B_list_week_unique[3]
        cell_17 = sheet['G12']
        cell_17.value = _3B_list_week_unique[4]
        cell_18 = sheet['H12']
        cell_18.value = _3B_list_week_unique[5]
        cell_18I = sheet['I12']
        cell_18I.value = _3B_list_week_unique[6]

        #data filling for unique season values

        cell_19 = sheet['C13']
        cell_19.value = _3B_list_season_unique[0]
        cell_20 = sheet['D13']
        cell_20.value = _3B_list_season_unique[1]
        cell_21 = sheet['E13']
        cell_21.value = _3B_list_season_unique[2]
        cell_22 = sheet['F13']
        cell_22.value = _3B_list_season_unique[3]
        cell_23 = sheet['G13']
        cell_23.value = _3B_list_season_unique[4]
        cell_24 = sheet['H13']
        cell_24.value = _3B_list_season_unique[5]
        cell_24I = sheet['I13']
        cell_24I.value = _3B_list_season_unique[6]

        
        #data filling for week values
        
        cell_25 = sheet['C14']
        cell_25.value = _3B_list_week[0]
        cell_26 = sheet['D14']
        cell_26.value = _3B_list_week[1]
        cell_27 = sheet['E14']
        cell_27.value = _3B_list_week[2]
        cell_28 = sheet['F14']
        cell_28.value = _3B_list_week[3]
        cell_29 = sheet['G14']
        cell_29.value = _3B_list_week[4]
        cell_30 = sheet['H14']
        cell_30.value = _3B_list_week[5]
        cell_30I = sheet['I14']
        cell_30I.value = _3B_list_week[6]
        
        #data filling for season value

        cell_31 = sheet['C15']
        cell_31.value = _3B_list_season[0]
        cell_32 = sheet['D15']
        cell_32.value = _3B_list_season[1]
        cell_33 = sheet['E15']
        cell_33.value = _3B_list_season[2]
        cell_34 = sheet['F15']
        cell_34.value = _3B_list_season[3]
        cell_35 = sheet['G15']
        cell_35.value = _3B_list_season[4]
        cell_36 = sheet['H15']
        cell_36.value = _3B_list_season[5]
        cell_36I = sheet['I15']
        cell_36I.value = _3B_list_season[6]


        #data filling for season crop - Yes
        cell_37 = sheet['C16']
        cell_37.value = _3B_list_season_crop[0]
        cell_38 = sheet['D16']
        cell_38.value = _3B_list_season_crop[1]
        cell_39 = sheet['E16']
        cell_39.value = _3B_list_season_crop[2]
        cell_40 = sheet['F16']
        cell_40.value = _3B_list_season_crop[3]
        cell_41 = sheet['G16']
        cell_41.value = _3B_list_season_crop[4]
        cell_42 = sheet['H16']
        cell_42.value = _3B_list_season_crop[5]
        cell_42I = sheet['I16']
        cell_42I.value = _3B_list_season_crop[6]


        cell_98 = sheet2['C18']               #3.PROGRAM DASHBOARD
        cell_98.value = _3B_list_season_unique[0]
        cell_99 = sheet2['D18']
        cell_99.value = _3B_list_season_unique[1]
        cell_100 = sheet2['E18']
        cell_100.value = _3B_list_season_unique[2]
        cell_101 = sheet2['F18']
        cell_101.value = _3B_list_season_unique[3]
        cell_102 = sheet2['G18']
        cell_102.value = _3B_list_season_unique[4]
        cell_103 = sheet2['H18']
        cell_103.value = _3B_list_season_unique[5]
        cell_103I = sheet2['I18']
        cell_103I.value = _3B_list_season_unique[6]


        cell_104 = sheet2['C20']                 #3.PROGRAM DASHBOARD
        cell_104.value = _3B_list_season[0]
        cell_105 = sheet2['D20']
        cell_105.value = _3B_list_season[1]
        cell_106 = sheet2['E20']
        cell_106.value = _3B_list_season[2]
        cell_107 = sheet2['F20']
        cell_107.value = _3B_list_season[3]
        cell_108 = sheet2['G20']
        cell_108.value = _3B_list_season[4]
        cell_109 = sheet2['H20']
        cell_109.value = _3B_list_season[5]
        cell_109I = sheet2['I20']
        cell_109I.value = _3B_list_season[6]


        cell_110 = sheet2['C21']                 #3.PROGRAM DASHBOARD
        cell_110.value = _3B_partner_sum[0]
        cell_111 = sheet2['D21']
        cell_111.value = _3B_partner_sum[1]
        cell_112 = sheet2['E21']
        cell_112.value = _3B_partner_sum[2]
        cell_113 = sheet2['F21']
        cell_113.value = _3B_partner_sum[3]
        cell_114 = sheet2['G21']
        cell_114.value = _3B_partner_sum[4]
        cell_115 = sheet2['H21']
        cell_115.value = _3B_partner_sum[5]
        cell_115I = sheet2['I21']
        cell_115I.value = _3B_partner_sum[6]



        print("3B run successfully")

    elif i == _3C_path:
        _3C_week = []
        _3C_PT_week=[]
        _3C_CT_week =[]
        _3C_season =[]
        _3C_PT_season = []
        _3C_CT_season = []
        _3C_unique_PT = []
        _3C_unique_CT = []

        
        season_df = season_df.drop(season_df[season_df['Surveyor Name'].isin(test_data)].index)

        for x in list2:
            count_df_3C_week = week_df[week_df['partner']== x]
            df_3cW = count_df_3C_week[['Surveyor Id']]
            list_week_3C_partner.append(df_3cW)

            count_df_3C_season = season_df[season_df['partner']== x]
            df_3cS = count_df_3C_season[['Surveyor Id']]
            list_season_3C_partner.append(df_3cS)

        for x in list2:
            count_PARTNER_week = week_df[week_df['partner']== x].shape[0]
            _3C_week.append(count_PARTNER_week)

        for x in list2:
            count_PARTNER_week_PT = week_df[(week_df['partner']== x)&(week_df['Select Plot Category   ']=='Programme Plot')].shape[0]
            _3C_PT_week.append(count_PARTNER_week_PT)

        for x in list2:
            count_PARTNER_week_CT = week_df[(week_df['partner']== x)&(week_df['Select Plot Category   ']=='Control Plot')].shape[0]
            _3C_CT_week.append(count_PARTNER_week_CT)

        for x in list2:
            count_PARTNER_season = season_df[season_df['partner']== x].shape[0]
            _3C_season.append(count_PARTNER_season)

        for x in list2:
            count_PARTNER_season_PT = season_df[(season_df['partner']== x)&(season_df['Select Plot Category   ']=='Programme Plot')].shape[0]
            _3C_PT_season.append(count_PARTNER_season_PT)

            #3.PROGRAM DASHBOARD
            df_PT_ = season_df[(season_df['partner']== x)&(season_df['Select Plot Category   ']=='Programme Plot')]
            count_df_PT = df_PT_['Programme Farmer Unique ID'].nunique()
            _3C_unique_PT.append(count_df_PT)
        for x in list2:
            count_PARTNER_season_CT = season_df[(season_df['partner']== x)&(season_df['Select Plot Category   ']=='Control Plot')].shape[0]
            _3C_CT_season.append(count_PARTNER_season_CT)

            #3.PROGRAM DASHBOARD
            df_CT_ = season_df[(season_df['partner']== x)&(season_df['Select Plot Category   ']=='Control Plot')]
            count_df_CT = df_CT_['Control Farmer Unique ID'].nunique()
            _3C_unique_CT.append(count_df_CT)

        #data filling in week  
        cell_43 = sheet['C22']
        cell_43.value = _3C_week[0]
        cell_44 = sheet['D22']
        cell_44.value = _3C_week[1]
        cell_45 = sheet['E22']
        cell_45.value = _3C_week[2]
        cell_46 = sheet['F22']
        cell_46.value = _3C_week[3]
        cell_47 = sheet['G22']
        cell_47.value = _3C_week[4]
        cell_48 = sheet['H22']
        cell_48.value = _3C_week[5]
        cell_48I = sheet['I22']
        cell_48I.value = _3C_week[6]

        #data filling in PT week
        cell_49 = sheet['C23']
        cell_49.value = _3C_PT_week[0]
        cell_50 = sheet['D23']
        cell_50.value = _3C_PT_week[1]
        cell_51 = sheet['E23']
        cell_51.value = _3C_PT_week[2]
        cell_52 = sheet['F23']
        cell_52.value = _3C_PT_week[3]
        cell_53 = sheet['G23']
        cell_53.value = _3C_PT_week[4]
        cell_54 = sheet['H23']
        cell_54.value = _3C_PT_week[5]
        cell_54I = sheet['I23']
        cell_54I.value = _3C_PT_week[6]

        #data filing in CT week
        cell_55 = sheet['C24']
        cell_55.value = _3C_CT_week[0]
        cell_56 = sheet['D24']
        cell_56.value = _3C_CT_week[1]
        cell_57 = sheet['E24']
        cell_57.value = _3C_CT_week[2]
        cell_58 = sheet['F24']
        cell_58.value = _3C_CT_week[3]
        cell_59 = sheet['G24']
        cell_59.value = _3C_CT_week[4]
        cell_60 = sheet['H24']
        cell_60.value = _3C_CT_week[5]
        cell_60I = sheet['I24']
        cell_60I.value = _3C_CT_week[6]

        #data filling in season 
        cell_61 = sheet['C25']
        cell_61.value = _3C_season[0]
        cell_62 = sheet['D25']
        cell_62.value = _3C_season[1]
        cell_63 = sheet['E25']
        cell_63.value = _3C_season[2]
        cell_64 = sheet['F25']
        cell_64.value = _3C_season[3]
        cell_65 = sheet['G25']
        cell_65.value = _3C_season[4]
        cell_66 = sheet['H25']
        cell_66.value = _3C_season[5]
        cell_66I = sheet['I25']
        cell_66I.value = _3C_season[6]

        #data filling in PT season
        cell_67 = sheet['C26']
        cell_67.value = _3C_PT_season[0]
        cell_68 = sheet['D26']
        cell_68.value = _3C_PT_season[1]
        cell_69 = sheet['E26']
        cell_69.value = _3C_PT_season[2]
        cell_70 = sheet['F26']
        cell_70.value = _3C_PT_season[3]
        cell_71 = sheet['G26']
        cell_71.value = _3C_PT_season[4]
        cell_72 = sheet['H26']
        cell_72.value = _3C_PT_season[5]
        cell_72I = sheet['I26']
        cell_72I.value = _3C_PT_season[6]

        #data filing in CT Season
        cell_73 = sheet['C28']
        cell_73.value = _3C_CT_season[0]
        cell_74 = sheet['D28']
        cell_74.value = _3C_CT_season[1]
        cell_75 = sheet['E28']
        cell_75.value = _3C_CT_season[2]
        cell_76 = sheet['F28']
        cell_76.value = _3C_CT_season[3]
        cell_78 = sheet['G28']
        cell_78.value = _3C_CT_season[4]
        cell_79 = sheet['H28']
        cell_79.value = _3C_CT_season[5]
        cell_79I = sheet['I28']
        cell_79I.value = _3C_CT_season[6]


        cell_116 = sheet2['C26']             #3.PROGRAM DASHBOARD
        cell_116.value = _3C_PT_season[0]
        cell_117 = sheet2['D26']
        cell_117.value = _3C_PT_season[1]
        cell_118 = sheet2['E26']
        cell_118.value = _3C_PT_season[2]
        cell_119 = sheet2['F26']
        cell_119.value = _3C_PT_season[3]
        cell_120 = sheet2['G26']
        cell_120.value = _3C_PT_season[4]
        cell_121 = sheet2['H26']
        cell_121.value = _3C_PT_season[5]
        cell_121I = sheet2['I26']
        cell_121I.value = _3C_PT_season[6]

        #data filing in CT Season            #3.PROGRAM DASHBOARD
        cell_122 = sheet2['C28']
        cell_122.value = _3C_CT_season[0]
        cell_123 = sheet2['D28']
        cell_123.value = _3C_CT_season[1]
        cell_124 = sheet2['E28']
        cell_124.value = _3C_CT_season[2]
        cell_125 = sheet2['F28']
        cell_125.value = _3C_CT_season[3]
        cell_126 = sheet2['G28']
        cell_126.value = _3C_CT_season[4]
        cell_127 = sheet2['H28']
        cell_127.value = _3C_CT_season[5]
        cell_127I = sheet2['I28']
        cell_127I.value = _3C_CT_season[6]


        cell_128 = sheet2['C25']             #3.PROGRAM DASHBOARD
        cell_128.value = _3C_unique_PT[0]
        cell_129 = sheet2['D25']
        cell_129.value = _3C_unique_PT[1]
        cell_130 = sheet2['E25']
        cell_130.value = _3C_unique_PT[2]
        cell_131 = sheet2['F25']
        cell_131.value = _3C_unique_PT[3]
        cell_132 = sheet2['G25']
        cell_132.value = _3C_unique_PT[4]
        cell_133 = sheet2['H25']
        cell_133.value = _3C_unique_PT[5]
        cell_133I = sheet2['I25']
        cell_133I.value = _3C_unique_PT[6]

        cell_134 = sheet2['C27']             #3.PROGRAM DASHBOARD
        cell_134.value = _3C_unique_CT[0]
        cell_135 = sheet2['D27']
        cell_135.value = _3C_unique_CT[1]
        cell_136 = sheet2['E27']
        cell_136.value = _3C_unique_CT[2]
        cell_137 = sheet2['F27']
        cell_137.value = _3C_unique_CT[3]
        cell_138 = sheet2['G27']
        cell_138.value = _3C_unique_CT[4]
        cell_139 = sheet2['H27']
        cell_139.value = _3C_unique_CT[5]
        cell_139I = sheet2['I27']
        cell_139I.value = _3C_unique_CT[6]

        print("3C run successfully")

    elif i == _4A_path:
        _4A_season =[]
        _4A_unique_season = []
        list_week_4A_partner = []
        list_4A_sum = []

        
        season_df = season_df.drop(season_df[season_df['Surveyor Name'].isin(test_data)].index)


        for x in list2:
            count_PARTNER_season_4A = season_df[season_df['partner']== x].shape[0]
            _4A_season.append(count_PARTNER_season_4A)
        
        for x in list2:
            PARTNER_season_4A_unique = season_df[season_df['partner']== x]
            _unique_season_value_PARTNER = PARTNER_season_4A_unique['Select unique Serial No. for Demo Plot'].nunique()
            _4A_unique_season.append(_unique_season_value_PARTNER)

        for x in list2:
            PARTNER_season_4A_sum = season_df[season_df['partner']== x]
            value_4A = 0
            

            for i in range(35,48,2):
                
                df_sum_4A=PARTNER_season_4A_sum.iloc[:,i].to_frame()
                if i == 35:
                    count_1 = df_sum_4A.dropna(subset=['If yes select the planned activities'])
                    count_1 = count_1['If yes select the planned activities'].apply(splitter).sum()
                    value_4A = value_4A + count_1
                elif i == 37:
                    count_2 = df_sum_4A.dropna(subset=['If yes select the planned activities.1'])
                    count_2 = count_2['If yes select the planned activities.1'].apply(splitter).sum()
                    value_4A = value_4A + count_2
            
                elif i == 39:
                    count_3 = df_sum_4A.dropna(subset=['If yes select the planned activities.2'])
                    count_3 = count_3['If yes select the planned activities.2'].apply(splitter).sum()
                    value_4A = value_4A + count_3
                
                elif i == 41:
                    count_4 = df_sum_4A.dropna(subset=['If yes select the planned activities.3'])
                    count_4 = count_4['If yes select the planned activities.3'].apply(splitter).sum()
                    value_4A = value_4A + count_4
                    
                    
                elif i == 43:
                    count_5 = df_sum_4A.dropna(subset=['If yes select the planned activities.4'])
                    count_5 = count_5['If yes select the planned activities.4'].apply(splitter).sum()
                    value_4A = value_4A + count_5
            
                elif i == 45:
                    count_6 = df_sum_4A.dropna(subset=['If yes select the planned activities.5'])
                    count_6 = count_6['If yes select the planned activities.5'].apply(splitter).sum()
                    value_4A = value_4A + count_6
                    
                elif i == 47:
                    count_7 = df_sum_4A.dropna(subset=['If yes select the planned activities.6'])
                    count_7 = count_7['If yes select the planned activities.6'].apply(splitter).sum()
                    value_4A = value_4A + count_7
            list_4A_sum.append(value_4A)
        

    
        for x in list2:
            count_df_4A_week = week_df[week_df['partner']== x]
            df_3cW = count_df_4A_week[['Surveyor Id']]
            list_week_4A_partner.append(df_3cW)

        

        cell_140 = sheet3['C7']             
        cell_140.value = _4A_unique_season[0]
        cell_141 = sheet3['D7']
        cell_141.value = _4A_unique_season[1]
        cell_142 = sheet3['E7']
        cell_142.value = _4A_unique_season[2]
        cell_143 = sheet3['F7']
        cell_143.value = _4A_unique_season[3]
        cell_145 = sheet3['G7']
        cell_145.value = _4A_unique_season[4]
        cell_146 = sheet3['H7']
        cell_146.value = _4A_unique_season[5]
        cell_146I = sheet3['I7']
        cell_146I.value = _4A_unique_season[6]

        cell_147 = sheet3['C10']             
        cell_147.value = _4A_season[0]
        cell_148 = sheet3['D10']
        cell_148.value = _4A_season[1]
        cell_149 = sheet3['E10']
        cell_149.value = _4A_season[2]
        cell_150 = sheet3['F10']
        cell_150.value = _4A_season[3]
        cell_151 = sheet3['G10']
        cell_151.value = _4A_season[4]
        cell_152 = sheet3['H10']
        cell_152.value = _4A_season[5]
        cell_152I = sheet3['I10']
        cell_152I.value = _4A_season[6]


        cell_153 = sheet3['C12']             
        cell_153.value = list_4A_sum[0]
        cell_154 = sheet3['D12']
        cell_154.value = list_4A_sum[1]
        cell_155 = sheet3['E12']
        cell_155.value = list_4A_sum[2]
        cell_156 = sheet3['F12']
        cell_156.value = list_4A_sum[3]
        cell_157 = sheet3['G12']
        cell_157.value = list_4A_sum[4]
        cell_158 = sheet3['H12']
        cell_158.value = list_4A_sum[5]
        cell_158I = sheet3['I12']
        cell_158I.value = list_4A_sum[6]


        print("4A run successfully")
        

    elif i== _4B_path:
        _4B_week =[]
        _4B_unique_week = []
        _4B_season =[]
        _4B_unique_season = []
        _4B_farmerSUM_week = []
        _4B_farmerSUM_season = []
        list_4B_sum_week =[]
        list_4B_sum_season =[]
        list_week_4B_partner=[]

        
        season_df = season_df.drop(season_df[season_df['Surveyor Name'].isin(test_data)].index)

        for x in list2:
            count_PARTNER_season_4B = season_df[season_df['partner']== x].shape[0]
            _4B_season.append(count_PARTNER_season_4B)
        
        for x in list2:
            PARTNER_season_4B_unique = season_df[season_df['partner']== x]
            _unique_season_value_PARTNER = PARTNER_season_4B_unique['Select Demo Plot for tracking '].nunique()
            _4B_unique_season.append(_unique_season_value_PARTNER)

        for x in list2:
            count_PARTNER_week_4B = week_df[week_df['partner']== x].shape[0]
            _4B_week.append(count_PARTNER_week_4B)
        
        for x in list2:
            PARTNER_week_4B_unique = week_df[week_df['partner']== x]
            _unique_week_value_PARTNER = PARTNER_week_4B_unique['Select Demo Plot for tracking '].nunique()
            _4B_unique_week.append(_unique_week_value_PARTNER)

        for x in list2:
            sum_farmer_week = week_df[week_df['partner']== x]
            sum_farmer_week_value = sum_farmer_week['Number of farmers who attended the Demo'].sum()
            _4B_farmerSUM_week.append(sum_farmer_week_value)

        for x in list2:
            sum_farmer_season = season_df[season_df['partner']== x]
            sum_farmer_season_value = sum_farmer_season['Number of farmers who attended the Demo'].sum()
            _4B_farmerSUM_season.append(sum_farmer_season_value)

        
        for x in list2:
            PARTNER_week_4B_sum = week_df[week_df['partner']== x]
            value_4A = 0
            

            for i in range(45,52):
                
                df_sum_4B=PARTNER_week_4B_sum.iloc[:,i].to_frame()
                if i == 45:
                    count_1 = df_sum_4B.dropna(subset=['Land Preparation: Select demonstration activities'])
                    count_1 = count_1['Land Preparation: Select demonstration activities'].apply(splitter).sum()
                    value_4A = value_4A + count_1
                elif i == 46:
                    count_2 = df_sum_4B.dropna(subset=['Seed Treatment and Sowing: Select demonstration activities'])
                    count_2 = count_2['Seed Treatment and Sowing: Select demonstration activities'].apply(splitter).sum()
                    value_4A = value_4A + count_2
            
                elif i == 47:
                    count_3 = df_sum_4B.dropna(subset=['Soil Health and Nutrition: Select demonstration activities'])
                    count_3 = count_3['Soil Health and Nutrition: Select demonstration activities'].apply(splitter).sum()
                    value_4A = value_4A + count_3
                
                elif i == 48:
                    count_4 = df_sum_4B.dropna(subset=['Plant Growth: Select demonstration activities'])
                    count_4 = count_4['Plant Growth: Select demonstration activities'].apply(splitter).sum()
                    value_4A = value_4A + count_4
                    
                    
                elif i == 49:
                    count_5 = df_sum_4B.dropna(subset=['Pest and Weed Management: Select demonstration activities'])
                    count_5 = count_5['Pest and Weed Management: Select demonstration activities'].apply(splitter).sum()
                    value_4A = value_4A + count_5
            
                elif i == 50:
                    count_6 = df_sum_4B.dropna(subset=['Irrigation: Select demonstration activities'])
                    count_6 = count_6['Irrigation: Select demonstration activities'].apply(splitter).sum()
                    value_4A = value_4A + count_6
                    
                elif i == 51:
                    count_7 = df_sum_4B.dropna(subset=['Post Harvesting: Select demonstration activities'])
                    count_7 = count_7['Post Harvesting: Select demonstration activities'].apply(splitter).sum()
                    value_4A = value_4A + count_7
            list_4B_sum_week.append(value_4A)



        for x in list2:
            PARTNER_week_4B_sum = season_df[season_df['partner']== x]
            value_4A = 0
            

            for i in range(45,52):
                
                df_sum_4B=PARTNER_week_4B_sum.iloc[:,i].to_frame()
                if i == 45:
                    count_8 = df_sum_4B.dropna(subset=['Land Preparation: Select demonstration activities'])
                    count_8 = count_8['Land Preparation: Select demonstration activities'].apply(splitter).sum()
                    value_4A = value_4A + count_8
                elif i == 46:
                    count_9 = df_sum_4B.dropna(subset=['Seed Treatment and Sowing: Select demonstration activities'])
                    count_9 = count_9['Seed Treatment and Sowing: Select demonstration activities'].apply(splitter).sum()
                    value_4A = value_4A + count_9
            
                elif i == 47:
                    count_10 = df_sum_4B.dropna(subset=['Soil Health and Nutrition: Select demonstration activities'])
                    count_10 = count_10['Soil Health and Nutrition: Select demonstration activities'].apply(splitter).sum()
                    value_4A = value_4A + count_10
                
                elif i == 48:
                    count_11 = df_sum_4B.dropna(subset=['Plant Growth: Select demonstration activities'])
                    count_11 = count_11['Plant Growth: Select demonstration activities'].apply(splitter).sum()
                    value_4A = value_4A + count_11
                    
                    
                elif i == 49:
                    count_12 = df_sum_4B.dropna(subset=['Pest and Weed Management: Select demonstration activities'])
                    count_12 = count_12['Pest and Weed Management: Select demonstration activities'].apply(splitter).sum()
                    value_4A = value_4A + count_12
            
                elif i == 50:
                    count_13 = df_sum_4B.dropna(subset=['Irrigation: Select demonstration activities'])
                    count_13 = count_13['Irrigation: Select demonstration activities'].apply(splitter).sum()
                    value_4A = value_4A + count_13
                    
                elif i == 51:
                    count_14 = df_sum_4B.dropna(subset=['Post Harvesting: Select demonstration activities'])
                    count_14 = count_14['Post Harvesting: Select demonstration activities'].apply(splitter).sum()
                    value_4A = value_4A + count_13
            list_4B_sum_season.append(value_4A)


        
        for x in list2:
            count_df_4B_week = week_df[week_df['partner']== x]  
            df_3cW = count_df_4B_week[['Surveyor Id']]
            list_week_4B_partner.append(df_3cW)

        
        cell_159 = sheet3['C14']             
        cell_159.value = _4B_unique_week[0]
        cell_160 = sheet3['D14']
        cell_160.value = _4B_unique_week[1]
        cell_161 = sheet3['E14']
        cell_161.value = _4B_unique_week[2]
        cell_162 = sheet3['F14']
        cell_162.value = _4B_unique_week[3]
        cell_163 = sheet3['G14']
        cell_163.value = _4B_unique_week[4]
        cell_164 = sheet3['H14']
        cell_164.value = _4B_unique_week[5]
        cell_164I = sheet3['I14']
        cell_164I.value = _4B_unique_week[6]


        cell_165 = sheet3['C15']             
        cell_165.value = _4B_week[0]
        cell_166 = sheet3['D15']
        cell_166.value = _4B_week[1]
        cell_167 = sheet3['E15']
        cell_167.value = _4B_week[2]
        cell_168 = sheet3['F15']
        cell_168.value = _4B_week[3]
        cell_169 = sheet3['G15']
        cell_169.value = _4B_week[4]
        cell_170 = sheet3['H15']
        cell_170.value = _4B_week[5]
        cell_170I = sheet3['I15']
        cell_170I.value = _4B_week[6]


        cell_171 = sheet3['C16']             
        cell_171.value = list_4B_sum_week[0]
        cell_172 = sheet3['D16']
        cell_172.value = list_4B_sum_week[1]
        cell_173 = sheet3['E16']
        cell_173.value = list_4B_sum_week[2]
        cell_174 = sheet3['F16']
        cell_174.value = list_4B_sum_week[3]
        cell_175 = sheet3['G16']
        cell_175.value = list_4B_sum_week[4]
        cell_176 = sheet3['H16']
        cell_176.value = list_4B_sum_week[5]
        cell_176I = sheet3['I16']
        cell_176I.value = list_4B_sum_week[6]


        cell_177 = sheet3['C17']             
        cell_177.value = _4B_farmerSUM_week[0]
        cell_178 = sheet3['D17']
        cell_178.value = _4B_farmerSUM_week[1]
        cell_179 = sheet3['E17']
        cell_179.value = _4B_farmerSUM_week[2]
        cell_180 = sheet3['F17']
        cell_180.value = _4B_farmerSUM_week[3]
        cell_181 = sheet3['G17']
        cell_181.value = _4B_farmerSUM_week[4]
        cell_182 = sheet3['H17']
        cell_182.value = _4B_farmerSUM_week[5]
        cell_182I = sheet3['I17']
        cell_182I.value = _4B_farmerSUM_week[6]


        cell_183 = sheet3['C18']             
        cell_183.value = _4B_unique_season[0]
        cell_184 = sheet3['D18']
        cell_184.value = _4B_unique_season[1]
        cell_185 = sheet3['E18']
        cell_185.value = _4B_unique_season[2]
        cell_186 = sheet3['F18']
        cell_186.value = _4B_unique_season[3]
        cell_187 = sheet3['G18']
        cell_187.value = _4B_unique_season[4]
        cell_188 = sheet3['H18']
        cell_188.value = _4B_unique_season[5]
        cell_188I = sheet3['I18']
        cell_188I.value = _4B_unique_season[6]


        cell_189 = sheet3['C19']             
        cell_189.value = _4B_season[0]
        cell_190 = sheet3['D19']
        cell_190.value = _4B_season[1]
        cell_191 = sheet3['E19']
        cell_191.value = _4B_season[2]
        cell_192 = sheet3['F19']
        cell_192.value = _4B_season[3]
        cell_193 = sheet3['G19']
        cell_193.value = _4B_season[4]
        cell_194 = sheet3['H19']
        cell_194.value = _4B_season[5]
        cell_194I = sheet3['I19']
        cell_194I.value = _4B_season[6]



        cell_195 = sheet3['C20']             
        cell_195.value = list_4B_sum_season[0]
        cell_196 = sheet3['D20']
        cell_196.value = list_4B_sum_season[1]
        cell_197 = sheet3['E20']
        cell_197.value = list_4B_sum_season[2]
        cell_198 = sheet3['F20']
        cell_198.value = list_4B_sum_season[3]
        cell_199 = sheet3['G20']
        cell_199.value = list_4B_sum_season[4]
        cell_200 = sheet3['H20']
        cell_200.value = list_4B_sum_season[5]
        cell_200I = sheet3['I20']
        cell_200I.value = list_4B_sum_season[6]




        cell_201 = sheet3['C22']             
        cell_201.value = _4B_farmerSUM_season[0]
        cell_202 = sheet3['D22']
        cell_202.value = _4B_farmerSUM_season[1]
        cell_203 = sheet3['E22']
        cell_203.value = _4B_farmerSUM_season[2]
        cell_204 = sheet3['F22']
        cell_204.value = _4B_farmerSUM_season[3]
        cell_205 = sheet3['G22']
        cell_205.value = _4B_farmerSUM_season[4]
        cell_206 = sheet3['H22']
        cell_206.value = _4B_farmerSUM_season[5]
        cell_206I = sheet3['I22']
        cell_206I.value = _4B_farmerSUM_season[6]

        print("4B run successfully")

unique_value_week_3A3B3C = []
unique_value_season_3A3B3C = []
unique_value_week_4A4B = [] 
for i in range(0,7):
    df_week_C3A3B3C = pd.concat([list_week_3A_partner[i],list_week_3B_partner[i],list_week_3C_partner[i]])
    a = df_week_C3A3B3C['Surveyor Id'].nunique()
    unique_value_week_3A3B3C.append(a)

for i in range(0,7):
    df_season_C = pd.concat([list_season_3A_partner[i],list_season_3B_partner[i],list_season_3C_partner[i]])
    a = df_season_C['Surveyor Id'].nunique()
    unique_value_season_3A3B3C.append(a)


for i in range(0,7):
    df_week_C4A4B = pd.concat([list_week_4A_partner[i],list_week_4B_partner[i]])
    a = df_week_C4A4B['Surveyor Id'].nunique()
    unique_value_week_4A4B.append(a)



# print(unique_value_week)
# print(unique_value_season)
cell_80 = sheet['C4']
cell_80.value = unique_value_week_3A3B3C[0]
cell_81 = sheet['D4']
cell_81.value = unique_value_week_3A3B3C[1]
cell_82 = sheet['E4']
cell_82.value = unique_value_week_3A3B3C[2]
cell_83 = sheet['F4']
cell_83.value = unique_value_week_3A3B3C[3]
cell_84 = sheet['G4']
cell_84.value = unique_value_week_3A3B3C[4]
cell_85 = sheet['H4']
cell_85.value = unique_value_week_3A3B3C[5]
cell_85I = sheet['I4']
cell_85I.value = unique_value_week_3A3B3C[6]

cell_86 = sheet['C6']
cell_86.value = unique_value_season_3A3B3C[0]
cell_87 = sheet['D6']
cell_87.value = unique_value_season_3A3B3C[1]
cell_88 = sheet['E6']
cell_88.value = unique_value_season_3A3B3C[2]
cell_89 = sheet['F6']
cell_89.value = unique_value_season_3A3B3C[3]
cell_90 = sheet['G6']
cell_90.value = unique_value_season_3A3B3C[4]
cell_91 = sheet['H6']
cell_91.value = unique_value_season_3A3B3C[5]
cell_91I = sheet['I6']
cell_91I.value = unique_value_season_3A3B3C[6]


cell_207 = sheet3['C4']
cell_207.value = unique_value_week_4A4B[0]
cell_208 = sheet3['D4']
cell_208.value = unique_value_week_4A4B[1]
cell_209 = sheet3['E4']
cell_209.value = unique_value_week_4A4B[2]
cell_210 = sheet3['F4']
cell_210.value = unique_value_week_4A4B[3]
cell_211 = sheet3['G4']
cell_211.value = unique_value_week_4A4B[4]
cell_212 = sheet3['H4']
cell_212.value = unique_value_week_4A4B[5]
cell_212I = sheet3['I4']
cell_212I.value = unique_value_week_4A4B[6]


workbook.save(DMS)
print("Hurray ! Program run Successfully")