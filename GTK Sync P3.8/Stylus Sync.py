from tkinter import filedialog

import numpy as np
import pandas as pd
import os
import tkinter.filedialog
import openpyxl
from os import walk
import json
import math
import matplotlib.pyplot as plt
from datetime import datetime


Test_Number = 1
mainloop = 0

df33 = pd.DataFrame()

# File dialog to select folder containing excel files to import
full_path = tkinter.filedialog.askdirectory(initialdir='C://')
os.chdir(full_path)

f1 = []
f2 = []
for (dirpath, dirnames, filenames) in walk(full_path):
    f1.extend(filenames)
    f2.extend(dirnames)
    mainloop = mainloop + 1
    if mainloop == 2:
        break


    for folders in f2:
        if folders != 'Plots' or folders != 'Results Package':
            f3 = []
            for (dirpath2, dirnames2, filenames2) in walk(full_path+'/'+folders):
                f3.extend(dirnames2)
            # df1 = pd.DataFrame()
            # df2 = pd.DataFrame()
            print(folders)
            csv = 0
            excel = 0
            report = 0
            MD = 0
            UP = 0
            # for folders2 in f3:
            #     if folders2 != 'Plots' or folders2 != 'Results Package':
            #         if csv == 1 and excel == 1 and report == 1 and MD == 1 and UP == 1:
            #             break
            #
            #         if folders2 == 'position 1' or 'position 2':
            #             if folders2 == 'position 1':
            #                 MD = 1
            #             if folders2 == 'position 2':
            #                 UP = 1

            f4 = []
            for (dirpath3, dirnames3, filenames3) in walk(full_path+'/'+folders):
                f4.extend(filenames3)
            print('debug')
            df1 = pd.DataFrame()
            df2 = pd.DataFrame()

            for files in f4:
                if files.endswith(".csv"):
                    if files.startswith('230'):
                        test_path = full_path+'/'+folders+'/'+files
                        print(test_path)
                        df1 = pd.read_csv(test_path)
                        csv = 1
                if files.endswith('time.csv'):
                    test_path1 = full_path+'/'+folders+'/'+files
                    print(test_path1)
                    df2 = pd.read_csv(test_path1)
                    excel = 1
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

            df3 = df2.loc[df2['TimeStamp'] != 'TimeStamp']
            df3.dropna(axis=0, how='all')








            adc_2 = df1['stylus']
            force_1 = df3[df2.columns[2]]
            Z_pos = df2[df2.columns[1]]


            df4 = df2.loc[df2['TimeStamp'] == 'TimeStamp']

            index_Force = df4.index.tolist()


            # for time in df2['TimeStamp']:
            #     if
            df2_len = len(df2)

            d = {}
            index_l = len(index_Force)
            index_Force.append(df2_len)
            l = 0
            for index in index_Force:
                if l == 0:
                    d[index] = pd.DataFrame()
                    d[index] = df2.iloc[:(index)]
                elif 0 < l < index_l:
                    d[index] = pd.DataFrame()
                    d[index] = df2.iloc[(index_Force[l-1]+1):index]
                else:
                    d[index] = pd.DataFrame()
                    d[index] = df2.iloc[(index_Force[l-1]+1):index]
                l = l + 1


            df5 = df1.loc[df1['stylus'] == 0]
            df6 = pd.DataFrame()

            previous_time = df5['rtc ET[s]'][0]
            time_list = []
            o = 0
            for time in df5['rtc ET[s]']:
                if time - previous_time > 4:
                    time_list.append(time)
                    df6 = df6.append(df5.iloc[o-1])
                previous_time = time
                o = o + 1

            index_adc = df6.index.tolist()

            index_adc_2 = []
            for values in index_adc:
                values2 = values - 1
                index_adc_2.append(values2)

            j = 0
            for force_sample in d:
                if j == 0:
                    adc_start_index = index_adc[j]

                    force_df = d[force_sample]

                    last_force = 100
                    r = 0
                    for force in df3['Pressure(N)']:
                        if float(force) > last_force:
                            break
                        last_force = float(force)
                        r = r + 1



                    str_time_list = []

                    for time_stamp in df3['TimeStamp']:
                        # split_time = time_stamp.split(" ")
                        # time = split_time[1]

                        if len(str(time_stamp)) > 0:

                            if len(time_stamp) > 15:
                                date_time = datetime.strptime(time_stamp, '%Y-%m-%d %H:%M:%S.%f')
                                str_time_list.append(date_time)
                            else:
                                date_time = datetime.strptime(time_stamp, '%M:%S.%f')
                                str_time_list.append(date_time)


                    g = 0
                    time_d_list = []
                    # time_total = 0
                    elapsed_list = []
                    for time_2 in str_time_list:
                        if g >= 1:
                            prev_value_3 = str_time_list[g-1]
                            time_d = time_2 - prev_value_3
                            time_d_list.append(time_d)
                            if g == 52:
                                print('debug')
                            if g == 1:
                                time_total = time_d
                                elapsed_list.append(time_total)
                            else:
                                time_total = time_d + time_total
                                prev_value = time_d
                                elapsed_list.append(time_total)
                        g = g + 1

                    time_0 = time_d_list[0] - time_d_list[0]

                    elapsed_list.insert(0, time_0)

                    df3.insert(3, 'date-time', elapsed_list)

                    second_list_2 = []
                    for value_3 in elapsed_list:
                        value4 = str(value_3)
                        split_list = value4.split(':')
                        seconds = split_list[-1]
                        minute = split_list[-2]
                        total = float(minute)*60 + float(seconds)
                        # print("debug")
                        second_list_2.append(float(total))

                    df3.insert(4, 'rtc ET[s]', second_list_2)
                    force_start_row = df3.iloc[r]


                    adc_start_line = df1.iloc[adc_start_index]

                    adju_time_l = []
                    for time_3 in df3['rtc ET[s]']:
                        adju_time = time_3 + (adc_start_line['rtc ET[s]']-force_start_row['rtc ET[s]'])
                        adju_time_l.append(adju_time)

                    df3.insert(5, 'Sync rtc ET[s]', adju_time_l)

                    files_2 = os.listdir(full_path+'/'+folders)


                    f = open(full_path+'/'+folders+'/'+ json_file)
                    file_info = json.load(f)

                    Unit_SN = file_info['serial_number']

                    min_adc = min(df1['stylus_raw'])
                    normal_adc = []
                    for value_4 in df1['stylus_raw']:
                        new_value = (value_4/3000)-(min_adc/3000)
                        normal_adc.append(new_value)


                    force_list = []
                    for force in df3['Pressure(N)']:
                        force_list.append(float(force))

                    z_list = []
                    for z in df3['Position(mm)']:
                        z_list.append(float(z))


                    if not os.path.exists(full_path+'/Results Package'):
                        os.makedirs(full_path+'/Results Package')
                    if not os.path.exists(full_path+'/Results Package/Plots'):
                        os.makedirs(full_path+'/Results Package/Plots')
                    if not os.path.exists(full_path+'/Results Package/Plots'+Unit_SN+'.png'):

                        # full_adjusted_time = []
                        # for time_value in df2['rtc ET[s]']:
                        #     adju_time2 = time_value + (adc_start_line['rtc ET[s]']-force_start_row['rtc ET[s]'])
                        #     full_adjusted_time.append(adju_time2)
                        #
                        # force_list_3 = []
                        # for force_3 in df2['Pressure(N)']:
                        #     force_list_3.append(float(force_3))


                        fig1 = plt.figure()
                        # plt.plot(adju_time_l,force_df['Pressure(N)'], label='Load Cell')
                        plt.plot(df1['rtc ET[s]'],normal_adc, label='Normalised ADC')
                        plt.plot(adju_time_l,force_list, label='Load Cell')

                        # plt.plot(fixed_time_list,Z_Pos, label='Z_Position')

                        plt.legend(title='KPI Sources')
                        # plt.show()


                        print('debug')



                        plt.savefig(full_path+'/Results Package/Plots/'+Unit_SN+'.png',dpi=500)
                        plt.close()

                        df33['ADC Clock'] = pd.Series(df1['rtc ET[s]'])
                        df33['Pinch ADC'] = pd.Series(df1['stylus_raw'])
                        df33['Force Clock'] = pd.Series(adju_time_l)
                        df33['Force'] = pd.Series(force_list)
                        df33['Z_Position'] = pd.Series(z_list)

                        df33.to_excel(full_path+'/Results Package/'+Unit_SN+'results.xlsx', index=False)





                    print('debug')
                j = j + 1














