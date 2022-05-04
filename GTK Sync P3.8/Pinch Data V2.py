from tkinter import filedialog
import pandas as pd
import os
import tkinter.filedialog
import openpyxl
from os import walk
import json
import math
import matplotlib.pyplot as plt


Test_Number = 1

df33 = pd.DataFrame()

# File dialog to select folder containing excel files to import
full_path = tkinter.filedialog.askdirectory(initialdir='.')
os.chdir(full_path)

f1 = []
f2 = []
for (dirpath, dirnames, filenames) in walk(full_path):
    f1.extend(filenames)
    f2.extend(dirnames)

for folders in f2:
    if folders != 'Plots' or folders != 'position 2':
        f3 = []
        for (dirpath2, dirnames2, filenames2) in walk(full_path+'/'+folders):
            f3.extend(filenames2)
        df1 = pd.DataFrame()
        df2 = pd.DataFrame()
        for files in f3:
            if files.endswith(".csv"):
                test_path = full_path+'/'+folders+'/'+files
                print(test_path)
                df1 = pd.read_csv(test_path)
            if files.endswith('data.xlsx'):
                test_path1 = full_path+'/'+folders+'/'+files
                print(test_path1)
                df2 = pd.read_excel(test_path1, sheet_name='force-displacement CSV data')
            if files.endswith('.json'):
                json_file = files

        seconds_list = []
        print(folders+' '+files)
        for tick in df1['rtc_ticks']:
            second_ticks = tick / 125000
            seconds_list.append(second_ticks)

        rtc_delta = []
        rtc_ET = []
        previous_value = seconds_list[0]
        for second in seconds_list:

            second_delta = second - previous_value
            rtc_delta.append(second_delta)
            ET = second_delta + previous_value - seconds_list[0]
            rtc_ET.append(ET)
            previous_value = second


        df1.insert(22, 'rtc[s]', seconds_list)
        df1.insert(23, 'rtc delta', rtc_delta)
        df1.insert(24, 'rtc ET[s]', rtc_ET)

        adc_2 = df1['pinch_adc']
        force_1 = df2[df2.columns[0]]
        adc_5 = []
        for adc_1 in adc_2:
            adc_5.append(adc_1)
        i = 0
        for adc in adc_5:
            if adc_5[i] < adc_5[i+1] and adc_5[i+1] < adc_5[i+2] and adc_5[i+2] < adc_5[i+3] and adc_5[i+3] < adc_5[i+4] and adc_5[i+4] < adc_5[i+5] and adc_5[i+5] < adc_5[i+6] and adc_5[i+6] < adc_5[i+7]:
                break
            i = i + 1


        reset_force = []
        for force_value in force_1:
            reset_force.append(force_value)

        k = 0
        for force in force_1:
            if force_1[k] < force_1[k+1] and force_1[k+1] < force_1[k+2] and force_1[k+2] < force_1[k+3] and force_1[k+3] < force_1[k+4] and force_1[k+4] < force_1[k+5] and force_1[k+5] < force_1[k+6] and force_1[k+6] < force_1[k+7]:
                break
            k = k + 1

        force_2 = reset_force[k:]

        fig1 = plt.figure()                                                           #creates plot to select points for initial travel

        min_press = min(df1['pinch_adc'])
        adclist = []
        for adc in df1['pinch_adc']:
            adc_1000 = (adc/1000)-min_press/1000
            adclist.append(1*adc_1000)

        adclist3 = []
        for adc2 in df1['pinch_adc']:
            adclist3.append(adc2)

        adc_5_len = len(adc_5)
        adclist2 = adclist[i:]
        adclist4 = adclist3[i:]


        x1 = df1['rtc ET[s]']
        x1_len = len(x1)
        x11 = x1[i:]

        reset_time_list = []
        for time in x11:
            reset_time = time - x1[i]
            reset_time_list.append(reset_time)

        x2 = df2[df2.columns[1]]
        x2_len = len(x2)

        x22 = x2[k:]

        reset_time_list2 = []
        for time2 in x22:
            reset_time2 = time2 - x22[k]
            reset_time_list2.append(reset_time2)


        files_2 = os.listdir(full_path+'/'+folders)


        f = open(full_path + '/' + folders + '/' + json_file)
        file_info = json.load(f)

        Unit_SN = file_info['serial_number']


        if not os.path.exists(full_path+'/Plots'):
            os.makedirs(full_path+'/Plots')
        if not os.path.exists(full_path+'/Plots/'+Unit_SN+'.png'):

            plt.plot(reset_time_list2,force_2, label='Load Cell')
            plt.plot(reset_time_list,adclist2, label='Normalised ADC')

            plt.legend(title='KPI Sources')


            plt.savefig(full_path+'/Plots/'+Unit_SN+'.png',dpi=500)
            plt.close()

        # plt.show()
        #
        # print('debug')
        #
        #
        #
        #
        # plt.close()

        df33 = pd.DataFrame()
        force_4 = df2.iloc[:,0]

        force_4_len = len(force_4)
        force_5 = force_4[k:]

        Fox_Force_5 = []
        for jj in force_5:
            Fox_Force_5.append(jj)

        force_5.index -= (k)

        df33['ADC Clock'] = pd.Series(reset_time_list)
        df33['Pinch ADC'] = pd.Series(adclist4)
        df33['Force Clock'] = pd.Series(reset_time_list2)
        df33['Force'] = pd.Series(Fox_Force_5)


        df33.to_excel(full_path+'/'+folders+'/'+Unit_SN+' results.xlsx', index=False)
print('Done!')








