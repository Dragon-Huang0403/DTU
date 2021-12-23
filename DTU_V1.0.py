# -*- coding: utf-8 -*-

# Created on Wed Sep 15 16:48:48 2021

# @author: 11006747

# create exe step
# 1. creaet virtual envirconment in avaconda
# 2. install used libary
# 3. pip install PyQt5 --user #fix bug   #in cmd
# 4. pyinstaller -D xxxx.pu in cmd
# 5. copy folder(aardvark_py) from 'C:\Users\11006747\Anaconda3\envs\DTU\Lib\site-packages'
#     to the folder ('\dist\DTU_V1') createed by pyinstaller 
# or do this to repleace step4 & %
# pyinstaller -D DTU_V1.0.py --hiddenimport aardvark_py --collect-all aardvark_py --clean -w

# referecne:
# 1. https://www.pythonheidong.com/blog/article/564906/23b64289805d98dc77c0/
# 2. https://medium.com/pyladies-taiwan/python-%E5%B0%87python%E6%89%93%E5%8C%85%E6%88%90exe%E6%AA%94-32a4bacbe351
# 3. https://blog.csdn.net/moxiao1995071310/article/details/116932406

from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import time, csv, json
import threading as th
from aardvark_py import *
import multiprocessing as mp
import os, math
import win32api, win32net
import pandas as pd
from shareplum import Site, Office365
from shareplum.site import Version

import matplotlib.pyplot as plt
from matplotlib.animation import FuncAnimation
import matplotlib.animation as ani

DIMM_zone = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
DIMM_address_order = ['0x3c', '0x3e', '0x38', '0x3a', '0x34', '0x36', '0x30', '0x32']
DIMM_address_two_power_level = ['0x3e','0x3a','0x36','0x32']
BUS_TIMEOUT = 100  # ms
port    = 0
bitrate = 100 #kHz
length = 26 
item = 0
power_limit = 25
execute_time_gap = 10 #ms
monitor_time_gap = 1000 #ms
config = 0
Stop_testing =True
file_name= 'DDR5_TTV_Utility'
path = os.path.join(os.path.expandvars("%userprofile%"),file_name)
default_setting_filename = os.path.join(path,'Default_setting.json')
temp_data_csv = os.path.join(path,"temp_data.csv")

SHAREPOINT_URL = 'https://wistron.sharepoint.com/'
SHAREPOINT_SITE = r'https://wistron.sharepoint.com/sites/EKT/Labs/'
SHAREPOINT_DOC = 'For Autolog Testing'
Autolog_filename = 'Autolog.xlsm'
Autolog_file = '\\'.join([SHAREPOINT_DOC, Autolog_filename])
Autolog_local_path = os.path.join(path,Autolog_filename)
usage_log_filename = 'DDR5_TTV_Utility_Usage_Record.xlsx'
usage_log_sharepoint_path = '\\'.join([SHAREPOINT_DOC, usage_log_filename])
usage_log_local_path = os.path.join(path,usage_log_filename)
Data_temp_filename = 'Data_temp.csv'
Data_temp_local_path = os.path.join(path,Data_temp_filename)
Plot_data_filename = 'Plot_data.csv'
Plot_data_local_path = os.path.join(path,Plot_data_filename)
plot_label = ['x']+DIMM_address_order

data_label = ['Time', 'DIMM Zone', 'Address', 'Raw data',
              'Voltage', 'Tsensor0', 'Tsensor1', 'Tsensor2',
              'Tsensor3', 'Tsensor4', 'Tsensor5', 'Tsensor6',
              'Tsensor Max.', 'Power']
action = ''
how_many_zones =''
monitor_update_times = 1000

def initial_output_data():
    global output_data, data_temp, monitor_x
    monitor_x =0
    initial_data_temp_csv()
    initial_plot_data()
    data_temp = dict()
    output_data = dict()

    for i in DIMM_zone:
        if i in DIMM_zone[::2]:
            DIMM_address_order_temp = DIMM_address_order
        else:
            DIMM_address_order_temp = DIMM_address_order[::-1]
        for j in DIMM_address_order_temp:
            output_data[i,j] = initial_data_dict(i,j)

def initial_data_dict(DIMM_z,Addr):
    data_dict = {
                'Time' : '',
                'DIMM Zone': DIMM_z,
                'Address': Addr,
                'Raw data': '',
                'Voltage': '',
                'Tsensor0': '',
                'Tsensor1': '',
                'Tsensor2': '',
                'Tsensor3': '',
                'Tsensor4': '', 
                'Tsensor5': '',
                'Tsensor6': '',
                'Tsensor Max.': '' ,
                'Power' : ''
                    }
    return data_dict

def wistron_sharepoint_login(USERNAME,PASSWORD):
    try:
        global folder
        authcookie = Office365(SHAREPOINT_URL, username=USERNAME, password=PASSWORD).GetCookies()
        site = Site(SHAREPOINT_SITE, version=Version.v365, authcookie=authcookie)
        folder = site.Folder('')
        try:
            autolog_data = folder.get_file(Autolog_file)
            with open(Autolog_local_path, 'wb') as f:
                f.write(autolog_data)
                f.close()
        except: 
            pass
        return True
    except:
        return False
    
def get_user_name():
    global User
    user_info = win32net.NetUserGetInfo(win32net.NetGetAnyDCName(), win32api.GetUserName(), 2)
    User = user_info["full_name"]
    return User

def save_usage_log_data():
    current_time = time.localtime()
    Date = time.strftime("%Y/%m/%d %H:%M:%S", current_time)
    Month = time.strftime("%Y/%m", current_time)
    usage_log_data = {
            "User" : User,
            "Date" : Date,
            "Action" : action,
            "How many DIMM zones" : how_many_zones,
            "Auto mode status" : auto_mode_status,
            "Persistent storage status" : persistent_storage_status,
            "Two levels status" : two_levels_status,
            "Project Name" : Project_name,
            "Month" : Month
        }
    df = pd.read_excel(usage_log_local_path)
    df_combine = df.append(usage_log_data,ignore_index=True)
    df_combine.to_excel(usage_log_local_path,index=False)

def update_usage_data_to_sharepoint():
    try:
        global folder
        try:
            usage_log_data = folder.get_file(usage_log_sharepoint_path)
        except:
            authcookie = Office365(SHAREPOINT_URL, username=Email, password=Password).GetCookies()
            site = Site(SHAREPOINT_SITE, version=Version.v365, authcookie=authcookie)
            folder = site.Folder('')
            usage_log_data = folder.get_file(usage_log_sharepoint_path)
            
        with open(usage_log_local_path, 'wb') as f:
            f.write(usage_log_data)
            f.close()
        
        save_usage_log_data()
        with open (usage_log_local_path,'rb')as f:
            usage_log_data = f.read()
            f.close()
        folder.upload_file(usage_log_data, usage_log_sharepoint_path)

        return True
    except:
        return False

def save_default():
    default_setting = {
            "item_name" : item_name,
            "function_set" : function_set,
            "Need_power" : Need_power,
            "DIMM_zone_list" : DIMM_zone_list,
            "DIMM_address_list" : DIMM_address_list,
            "auto_mode_status" : auto_mode_status,
            "persistent_storage_status" : persistent_storage_status,
            "two_levels_status" : two_levels_status,
            "Email" : Email,
            "Password" : Password
            }
    with open(default_setting_filename, 'w') as f:
        json.dump(default_setting, f)
        f.close()
        
def get_default():
    global item_name, function_set, Need_power, DIMM_zone_list, DIMM_address_list
    global auto_mode_status, persistent_storage_status, two_levels_status
    global Email, Password
    try:
        with open(default_setting_filename, 'r') as f:
            default_setting = json.load(f)
            f.close()
    except:
        default_setting = {
            "item_name" : "DDR5_TTV_Data_1",
            "function_set" : "read",
            "Need_power" : ["1","1"],
            "DIMM_zone_list" : DIMM_zone,
            "DIMM_address_list" : DIMM_address_order,
            "auto_mode_status" : "0",
            "persistent_storage_status" : "0",
            "two_levels_status" : "0",
            "Email" : "",
            "Password" : ""
            }
    item_name = default_setting['item_name']
    function_set = default_setting['function_set']
    Need_power = default_setting['Need_power']
    DIMM_zone_list = ['','','','','','','','']
    for i in range(len(DIMM_zone)):
        if DIMM_zone[i]in default_setting['DIMM_zone_list']:
            DIMM_zone_list[i] = DIMM_zone[i]
                
    DIMM_address_list = ['','','','','','','','']
    for i in range(len(DIMM_address_order)):
        if DIMM_address_order[i]in default_setting['DIMM_address_list']:
            DIMM_address_list[i] = DIMM_address_order[i]
    auto_mode_status = default_setting['auto_mode_status']
    persistent_storage_status = default_setting['persistent_storage_status']
    two_levels_status = default_setting['two_levels_status']
    Email = default_setting['Email']
    Password = default_setting['Password']

# def plot_test():
#         global Stop_testing
        
#         plt.style.use('fivethirtyeight')
#         #time.sleep(2)
#         pause = False
#         while len(pd.read_csv(Plot_data_local_path)) <2:
#             pass
#         #index = count()
#         def onClick(event):
#             print(123)
        
#         def animate(i):
#             plot_all_data = pd.read_csv(Plot_data_local_path)
#             #x_vals.append(next(index))

#             x_vals = (plot_all_data[plot_label[0]])
#             y_vals1 = (plot_all_data[plot_label[1]])
#             y_vals2 = (plot_all_data[plot_label[2]])
#             y_vals3 = (plot_all_data[plot_label[3]])
#             y_vals4 = (plot_all_data[plot_label[4]])
#             y_vals5 = (plot_all_data[plot_label[5]])
#             y_vals6 = (plot_all_data[plot_label[6]])
#             y_vals7 = (plot_all_data[plot_label[7]])
#             y_vals8 = (plot_all_data[plot_label[8]])
            
            
#             plt.cla()
#             if y_vals1.iloc[-1]:
#                 plt.plot(x_vals,y_vals1,label=plot_label[1],marker="o")
#             if y_vals2.iloc[-1]:
#                 plt.plot(x_vals,y_vals2,label=plot_label[2],marker="o")
#             if y_vals3.iloc[-1]:
#                 plt.plot(x_vals,y_vals3,label=plot_label[3],marker="o")
#             if y_vals4.iloc[-1]:
#                 plt.plot(x_vals,y_vals4,label=plot_label[4],marker="o")
#             if y_vals5.iloc[-1]:
#                 plt.plot(x_vals,y_vals5,label=plot_label[5],marker="o")
#             if y_vals6.iloc[-1]:
#                 plt.plot(x_vals,y_vals6,label=plot_label[6],marker="o")
#             if y_vals7.iloc[-1]:
#                 plt.plot(x_vals,y_vals7,label=plot_label[7],marker="o")
#             if y_vals8.iloc[-1]:
#                 plt.plot(x_vals,y_vals8,label=plot_label[8],marker="o")
            
#             plt.legend(title="Tsensor Max.",loc='upper left',fancybox=True)
#             '''
#             if len(x_vals) >= 20:
#                 x_vals.pop(0)
#                 y_vals1.pop(0)
#                 y_vals2.pop(0)
#                 y_vals3.pop(0)
#                 y_vals4.pop(0)
#                 y_vals5.pop(0)
#                 y_vals6.pop(0)
#                 y_vals7.pop(0)
#                 y_vals8.pop(0)
#             '''
#             plt.xlabel('Time (sec.)')
#             plt.ylabel('Temperautre (C)')
#             plt.tight_layout()
#             #print(plt.fignum_exists())
        
#         #plt.canvas.mpl_connect('button_press_event', onClick)
#         ani = FuncAnimation(plt.gcf(),animate,interval=1000)
        
#         plt.tight_layout()
#         plt.show()
#         print(123)
#         Stop_testing =True

def monitor_plot():

    def animate(i):
        plot_all_data = pd.read_csv(Plot_data_local_path)
        x_vals = (plot_all_data[plot_label[0]])
        y_vals1 = (plot_all_data[plot_label[1]])
        y_vals2 = (plot_all_data[plot_label[2]])
        y_vals3 = (plot_all_data[plot_label[3]])
        y_vals4 = (plot_all_data[plot_label[4]])
        y_vals5 = (plot_all_data[plot_label[5]])
        y_vals6 = (plot_all_data[plot_label[6]])
        y_vals7 = (plot_all_data[plot_label[7]])
        y_vals8 = (plot_all_data[plot_label[8]])
        
        plt.cla()
        if y_vals1.iloc[-1]:
            ax.plot(x_vals,y_vals1,label=plot_label[1],marker="o")
        if y_vals2.iloc[-1]:
            ax.plot(x_vals,y_vals2,label=plot_label[2],marker="o")
        if y_vals3.iloc[-1]:
            ax.plot(x_vals,y_vals3,label=plot_label[3],marker="o")
        if y_vals4.iloc[-1]:
            ax.plot(x_vals,y_vals4,label=plot_label[4],marker="o")
        if y_vals5.iloc[-1]:
            ax.plot(x_vals,y_vals5,label=plot_label[5],marker="o")
        if y_vals6.iloc[-1]:
            ax.plot(x_vals,y_vals6,label=plot_label[6],marker="o")
        if y_vals7.iloc[-1]:
            ax.plot(x_vals,y_vals7,label=plot_label[7],marker="o")
        if y_vals8.iloc[-1]:
            ax.plot(x_vals,y_vals8,label=plot_label[8],marker="o")

        ax.legend(title="Tsensor Max.",loc='upper left',fancybox=True)

        ax.set_xlabel('Time (sec.)')
        ax.set_ylabel('Temperautre (C)')
        fig.tight_layout()
    
    def on_press(event):
        if event.key.isspace():
            if anim.running:
                anim.event_source.stop()
            else:
                anim.event_source.start()
            anim.running ^= True
        else:
            pass
    
    while len(pd.read_csv(Plot_data_local_path)) <2:
            pass
    
    plt.style.use('fivethirtyeight')
    fig, ax = plt.subplots()
    fig.canvas.mpl_connect('key_press_event', on_press)
    anim = ani.FuncAnimation(fig, animate, #frames=self.update_time,
                             interval=monitor_time_gap, repeat=False)
    anim.running = True
    plt.show()

def temp_transfer (TMPL, TMPH):
    TMPL_Bin = format(TMPL,'#010b')[2:4]
    TMPH_Bin = format(TMPH,'#010b')[2:]
    TMP = int('0b' + TMPH_Bin + TMPL_Bin, 2)
    return '%.2f' %((2.56 / 1024 * TMP - 0.75) * 100)

def voltage_transfer (Voltage_raw):   
    return Voltage_raw * 5.64 / 100

def get_PWM (Voltage, Power):
    PWM = (Power - 0.393) / (0.0098 * Voltage + 0.0015)
    PWM = math.ceil(PWM)
    if PWM <= 0:
        return 0
    else:
        return PWM

def get_power(Voltage, PWM):
    return PWM * (0.0098 * Voltage + 0.0015) + 0.393

def data_conversion(data):
    power = 0
    temperature = []
    data_result = []
    raw_data = ''
    
    for i in range(len(data)): # create hex raw data
        raw_data += "%02x " %(int(data[i]))
    voltage = voltage_transfer(data[20])
    for i in range(4):
        power += get_power(voltage, data[i]) / 4
    for i in range(6, 20, 2):
        temperature.append(temp_transfer(data[i], data[i+1]))
    temp_max = max(temperature)
    
    data_result.append('%.2f' % voltage)
    data_result += temperature
    data_result.append(temp_max)
    data_result.append('%.2f' % power)
    
    return raw_data.strip(), data_result

def check_Aardvark():
    (num, ports, unique_ids) = aa_find_devices_ext(16, 16)
    if num > 0:
        #print("%d device(s) found:" % num)
        # Print the information on each device
        for i in range(num):
            port      = ports[i]
            unique_id = unique_ids[i]
            # Determine if the device is in-use
            inuse = "(avail)"
            if (port & AA_PORT_NOT_FREE):
                inuse = "(in-use)"
                port  = port & ~AA_PORT_NOT_FREE
            # Display device port number, in-use status, and serial number
            #print("    port = %d   %s  (%04d-%06d)" %(port, inuse, unique_id // 1000000, unique_id % 1000000))
            if inuse == "(avail)":
                return 1 #avail
            #print('Please reconnect Aardvark')
            return 2 # please reconnect aardvark
            #messagebox.showinfo("123","456")
    #else:
        #print("No devices found.")
    return 0 # no devices found

def Aardvark_setup_connect(port, BUS_TIMEOUT, bitrate):
    handle = aa_open(port)
    aa_configure(handle,  AA_CONFIG_SPI_I2C)
    aa_i2c_pullup(handle, AA_I2C_PULLUP_NONE)
    aa_target_power(handle, AA_TARGET_POWER_NONE)
    bitrate = aa_i2c_bitrate(handle, bitrate)
    bus_timeout = aa_i2c_bus_timeout(handle, BUS_TIMEOUT)
    return handle
    
def Aardvark_disconnect(handle):
    aa_close(handle)

def i2c_read(handle, addr, item, length):
    aa_i2c_write(handle, addr, AA_I2C_NO_FLAGS, array('B', [item & 0xff]))
    (count, data_read) = aa_i2c_read(handle, addr, AA_I2C_NO_FLAGS, length)
    return (count,data_read)

def i2c_write(handle, addr, item, PWM, fix_power):
    if fix_power == '1':
        data_temp_i2c = [item & 0xff]+[PWM] *4 +[0]
        for i in range(6,24):
            data_temp_i2c.append(i)
        data_temp_i2c.append(1)
        data_out = array('B', data_temp_i2c)
    else:
        data_out = array('B', [item & 0xff]+[PWM] *4)
    aa_i2c_write(handle, addr, AA_I2C_NO_FLAGS, data_out)

def to_active_ttv_test_board(handle):
    aa_i2c_write(handle, 0x77, AA_I2C_NO_FLAGS, array('B', [ 4]))
    (count, data_read) = aa_i2c_read(handle, 0x77, AA_I2C_NO_FLAGS, 1)
    return data_read

# def test():
#     #print(123)
#     pass

def initial_data_temp_csv():
    with open(Data_temp_local_path, 'w', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=data_label)
        writer.writeheader()
    with open(Plot_data_local_path, 'w', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=plot_label)
        writer.writeheader()

def save_temp_data(item):
    global plot_data
    addr_temp = data_temp[item]['Address']
    if plot_data[addr_temp] ==0:
        plot_data[addr_temp] = data_temp[item]['Tsensor Max.']
    else:
        global monitor_x
        monitor_x +=1
        with open(Plot_data_local_path, 'a', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=plot_label)
            writer.writerow(plot_data)
            f.close()
        initial_plot_data()
    with open(Data_temp_local_path, 'a', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=data_label)
        writer.writerow(data_temp[item])
        f.close()
        
def initial_plot_data():
    global plot_data
    plot_data = dict()
    plot_data[plot_label[0]] = monitor_x
    for i in DIMM_address_order:
        plot_data[i] = 0

class 	DDR5_DIMM_TTV_Toolkit:

    def __init__(self, master):
        self.master = master
        self.frame = ttk.Frame(self.master, borderwidth = 5,relief = 'groove')
        self.master.geometry('984x667+300+100')
        self.master.title("DTU (DDR5 TTV Utility) V1.0")
        self.master.resizable(FALSE,FALSE)
        self.frame.grid()
        self.frame_top = ttk.Frame(self.frame,padding="10 0 0 10")
        self.frame_top.grid(column=0,row=0,sticky="W",columnspan=1)
        
        '''Function block'''
        self.frame_function = ttk.Frame(self.frame_top,padding="10 10 0 0")
        self.frame_function.grid(column=0,row=0,sticky="W")
        self.function = StringVar() 
        self.function.set(function_set)
        self.read = Radiobutton(self.frame_function, text='Get Data', variable=self.function
                                    , command=self.check_status, value='read', font = ("Times",14))
        self.read.grid(column=0, row=0,sticky="W",pady=5)
        self.write = Radiobutton(self.frame_function, text='Set Power', variable=self.function
                                     , command=self.check_status, value='write', font = ("Times",14))
        self.write.grid(column=0, row=1,sticky="W")
        self.vcmd = (self.master.register(self.power_validate), '%P')
        self.Need_power_entry = Entry(self.frame_function,validate='key'
                                          , validatecommand=self.vcmd,width=7
                                          ,justify='center')
        self.Need_power_w = ttk.Label(self.frame_function,
                                      text=("W     Max. Power = "+str(power_limit)+" W")
                                      ,font = ("Times",14))
        self.Need_power_entry.grid(column=1,row=1,sticky="W",padx=5)
        self.Need_power_w.grid(column=4,row=1,sticky="W")
        self.Need_power_entry.insert(0,Need_power[0])
        
        self.Need_power_and = ttk.Label(self.frame_function,text="and",font = ("Times",14))
        self.Need_power_and.grid(column=2,row=1)
        self.Need_power_entry_2 = Entry(self.frame_function,validate='key',
                                        background ='gray60'
                                          , validatecommand=self.vcmd,width=7
                                          ,justify='center')
        self.Need_power_entry_2.grid(column=3,row=1,sticky="W",padx=5)
        self.Need_power_entry_2.insert(0,Need_power[1])
        
        '''Power setting options'''
        self.power_setting = ttk.Frame(self.frame_top,padding="30 0 0 0")
        self.power_setting.grid(column=0,row=1,sticky="W")
        ttk.Label(self.power_setting).grid(column=0,row=0,pady=15)
        self.power_setting_lable_frame = ttk.LabelFrame(self.power_setting,text="Power Setting Additional Options")
        self.power_setting_lable_frame.grid(column=0,row=0)
        self.fix_power_var = StringVar()
        self.fix_power_CV = Checkbutton(self.power_setting_lable_frame,text="Persistent Storage",
                                variable=self.fix_power_var,onvalue='1',
                                offvalue='0',font=("Times",12))
        self.fix_power_CV.grid(column = 0, row = 0,sticky="W",padx = '0 20')
        self.fix_power_var.set('0')
        
        self.two_power_level_var = StringVar()
        self.two_power_level_CB = Checkbutton(self.power_setting_lable_frame,text="Two Levels", command = self.check_two_power_level,
                                variable=self.two_power_level_var,onvalue='1',
                                offvalue='0',font=("Times",12))
        self.two_power_level_CB.grid(column = 1, row = 0,sticky="W",padx='0 0')
        self.two_power_level_var.set(two_levels_status)
        
        '''DIMM block'''
        self.frame_DIMM = ttk.Frame(self.frame_top,padding="10 0 0 0")
        self.frame_DIMM.grid(column=0,row=2,sticky="E W")
        ttk.Label(self.frame_DIMM,text="DIMM Zones",font=("Times",12)).grid(column=0,row=0,sticky="W",pady=5)
        self.DIMM_zone =[]
        self.DIMM_zone_CB =[]

        for i in range(len(DIMM_zone)):
            self.DIMM_zone.append(StringVar())
            self.DIMM_zone_CB.append(Checkbutton(self.frame_DIMM,text=DIMM_zone[i],
                                variable=self.DIMM_zone[i],onvalue=DIMM_zone[i],
                                offvalue='',font=("Times",12)))
            self.DIMM_zone_CB[i].grid(column = i+1, row = 0,sticky="W")
            self.DIMM_zone[i].set(DIMM_zone_list[i])
            self.DIMM_zone_CB[i].grid_remove()

        self.DIMM_zone_R = StringVar() 
        self.DIMM_zone_RB =[]
        for i in range(len(DIMM_zone)):
            self.DIMM_zone_RB.append(Radiobutton(self.frame_DIMM,text=DIMM_zone[i],
                                variable=self.DIMM_zone_R,value=DIMM_zone[i],
                               font=("Times",12)))
            self.DIMM_zone_RB[i].grid(column = i+1, row = 0,sticky="W")
            self.DIMM_zone_RB[i].grid_remove()
        self.DIMM_zone_R.set('A')          
        
        ttk.Label(self.frame_DIMM,text="DIMM Address",font=("Times",12)).grid(column=0,row=1,sticky="W")
        self.DIMM_address =[]
        self.DIMM_address_CB =[]
        for i in range(len(DIMM_address_order)):
            self.DIMM_address.append(StringVar())
            self.DIMM_address_CB.append(Checkbutton(self.frame_DIMM, text=DIMM_address_order[i]
                                        ,variable=self.DIMM_address[i],onvalue=DIMM_address_order[i]
                                          ,offvalue='',font=("Times",12)))
            self.DIMM_address_CB[i].grid(column=1+i, row=1,sticky="W")
            self.DIMM_address[i].set(DIMM_address_list[i])
                
        '''Button block'''
        self.frame_button = ttk.Frame(self.frame,padding = "0 20 15 10")
        self.frame_button.grid(column=1,row=0,sticky="N E")
        self.execute_BT_name = StringVar()
        self.execute_BT_name.set('Start (F5)')
        self.execute_BT = Button(self.frame_button,width=12,font=("Times",13),
                                 textvariable=self.execute_BT_name,command=self.check_execute,pady=10,bg='Burlywood')
        self.execute_BT.grid(column=1,row=0)
        self.master.bind("<F5>", lambda event:self.check_execute())
        
        self.stop_BT = Button(self.frame_button,width=12,font=("Times",13),text="Stop",command=self.Stop,pady=5)
        self.stop_BT.grid(column=1,row=1)
        self.stop_BT['state'] = DISABLED
        
        self.skip_BT = Button(self.frame_button,width=12,font=("Times",13),text="Skip (Esc)",command=self.Skip,pady=5)
        self.skip_BT.grid(column=1,row=2)
        self.master.bind("<Escape>", lambda event:self.Skip())
        self.monitor_BT = Button(self.frame_button,width=12,font=("Times",13),text="Monitor",command=self.check_monitor,pady=5)
        self.monitor_BT.grid(column=1,row=3)
        self.skip_BT['state'] = DISABLED
        
        self.auto_mode_var = StringVar()
        Radiobutton(self.frame_button, text='Basic Mode', variable=self.auto_mode_var
                                    , command=self.check_status, value='0',
                                    font = ("Times",14)).grid(column = 0, row = 0,
                                                              sticky="W",padx='0 40')
        Radiobutton(self.frame_button, text='Advanced Mode', variable=self.auto_mode_var
                                    , command=self.check_status, value='1',
                                    font = ("Times",14)).grid(column = 0, row = 1,
                                                              sticky="W N",padx='0 40')
        self.auto_mode_var.set(auto_mode_status)
        
        self.project_name = StringVar()
        Label(self.frame_button,anchor = "w",font = ("Times",12),text='Project Name :').grid(column=0,row=2,sticky="W S")
        #Entry(self.frame_button,width=20,justify='center',textvariable=self.project_name)#.grid(column=0,row=3,sticky="W N")
        self.project_name_combobox = ttk.Combobox(self.frame_button, textvariable=self.project_name)
        self.project_name_combobox.grid(column=0,row=3,sticky="W N")
        
        #TEMPTEMPTEMP
        #Button(self.frame_button,width=12,font=("Times",13),text="Temp",command=self.new_window).grid(column=0,row=3)
        
        '''Status
        self.status_text= StringVar()
        self.status_text.set("Getting data from DIMM zone A")
        Label(self.frame,textvariable=self.status_text,bg = "yellow",font=("Times",14,"bold")).grid(column=0,row=1,columnspan = 2)
        '''
        
        '''Treeview block'''
        self.frame_tree = ttk.Frame(self.frame,borderwidth = 3,relief="groove")
        self.frame_tree.grid(column=0,row=2,columnspan=2,sticky = "W E")
        self.tree = ttk.Treeview(self.frame_tree,height = 18)
        self.tree.grid(column=0,row=0,columnspan=2,sticky = "W E")
        self.tree.column("#0", width=130,anchor='center')
        
        #self.tree['show'] = 'headings'
        self.tree_columns = ['dimm_zone','address', 'voltage','t0','t1','t2','t3','t4','t5','t6','tmax','power']
        self.tree_columns_name = ['DIMM Zone','Address', 'Voltage','Tsensor 0','Tsensor 1','Tsensor 2','Tsensor 3','Tsensor 4','Tsensor 5','Tsensor 6','Tsensor Max','Power']
        self.tree['columns'] = self.tree_columns
        for i in range(len(self.tree_columns)):
            if i == 0:
                self.tree_width = 85
            elif i ==1:
                self.tree_width = 60
            elif i ==2:
                self.tree_width = 62
            elif i ==10:
                self.tree_width = 85
            elif i ==11:
                self.tree_width = 71
            else:
                self.tree_width = 65
            self.tree.column(self.tree_columns[i],width=self.tree_width, anchor='center')
            self.tree.heading(self.tree_columns[i], text=self.tree_columns_name[i])
        #self.test_list = ["A",'0x30','11.96', '32.75', '51', '58', '44.25', '44.5', '44.75', '56.25', '58', '11.07']
        
        '''
        for i in range(3):
            self.test_list[0] = chr(65+i)
            if i % 2 == 0:
                self.tree_config ="config_1"
            else:
                self.tree_config ="config_2"
            for j in range(8):
                self.test_list[1] = DIMM_address_order[j]
                self.tree.insert("","end",j+i*8, text=time.strftime('%x %X') 
                                 , value = self.test_list, tags=(self.tree_config))
        '''
        self.tree.tag_configure("config_2",background ='wheat')
        self.tree.tag_configure("config_3",background ='yellow',foreground='red')
        #print(self.tree.get_children())
        
        self.scrollerbar = ttk.Scrollbar(self.frame_tree, orient=VERTICAL, command=self.tree.yview)
        self.scrollerbar.grid(column=2,row=0,sticky="N S E W")
        self.tree['yscrollcommand'] = self.scrollerbar.set
        
        self.enable_TTB_BT = Button(self.frame_tree,text="Enable TTV test Board",command=self.enable_ttv_test_board,font=("Times",13))
        self.enable_TTB_BT.grid(column=0,row=1,sticky="W")
        
        self.frame_tree_button = ttk.Frame(self.frame_tree)
        self.frame_tree_button.grid(column=1,row=1,columnspan=2,sticky="E")
        self.clear_BT = Button(self.frame_tree_button,text="Clear data", command = self.clear_log,font=("Times",13),width =15)
        self.clear_BT.grid(column=0,row=0,sticky="E")
        self.save_BT = Button(self.frame_tree_button,text="Save to File (Ctrl+S)",command=self.Save,font=("Times",13))
        self.save_BT.grid(column=2,row=0,sticky="E")
        self.master.bind("<Control-s>", lambda event:self.Save())
        self.master.bind("<Control-S>", lambda event:self.Save())
        
        '''Bottom block'''
        self.frame_bottom = ttk.Frame(self.frame,borderwidth = 3,relief="groove")
        self.frame_bottom.grid(column=0,row=3,columnspan=2,sticky="W E")
        ttk.Label(self.frame_bottom,text="Produced by Dragon Huang").grid(column=0,row=0,sticky="W")        
        
        self.check_status()
        self.master.focus_get()
        self.check_idendity()

    def mesgbox_info(self,text):
        messagebox.showinfo("MessageBox",text)
    
    def mesgbox_warning(self,text):
        messagebox.showwarning("Warning Message", text)
        
    def get_projecet_list(self):
        try:
            self.project_dataframe = pd.read_excel(Autolog_local_path,sheet_name = 'ProjectList')
            project_dataframe_list = self.project_dataframe['PROJECT_NAME'].tolist()
            project_list = list(set(project_dataframe_list[3:]))
            project_list_sorted = project_dataframe_list[0:3] + sorted(project_list, key=lambda x: str(x))
            self.project_name_combobox['value'] = project_list_sorted
        except:
            pass
    
    def check_idendity(self):
        if wistron_sharepoint_login(Email, Password):
            self.get_projecet_list()
        else:
            new_window(self.frame,self.master)    
    
    def check_status(self):
        global auto_mode_status,Need_power,function_set
        global persistent_storage_status
        self.check_DIMM_zone()
        self.check_DIMM_address()
        auto_mode_status = self.auto_mode_var.get()
        function_set = self.function.get()
        persistent_storage_status = '0'
        self.check_two_power_level()
        
        if self.Need_power_entry.get() =="":
            self.Need_power_entry.insert(0,0)

        if self.two_power_level_var.get() == "1":
            if self.Need_power_entry_2.get() =="":
                self.Need_power_entry_2.insert(0,0)
        
        self.power_setting_lable_frame.grid_remove()
        if auto_mode_status == '1':
            for i in range(len(DIMM_zone)):
                self.DIMM_zone_CB[i].grid()
                self.DIMM_zone_RB[i].grid_remove()
            if function_set == 'write':
                self.power_setting_lable_frame.grid()
                persistent_storage_status = self.fix_power_var.get()
                Need_power[1] = self.Need_power_entry_2.get()
                
        elif self.function.get() == 'write':
            for i in range(len(DIMM_zone)):
                self.DIMM_zone_CB[i].grid_remove()
                self.DIMM_zone_RB[i].grid_remove()
        else:
            for i in range(len(DIMM_zone)):
                self.DIMM_zone_CB[i].grid_remove()
                self.DIMM_zone_RB[i].grid()
        if self.function.get() == 'write':
            Need_power[0] = self.Need_power_entry.get()
            self.Need_power_entry.grid()
            self.Need_power_w.grid()
        else:
            self.Need_power_entry.grid_remove()
            self.Need_power_w.grid_remove()
    
    def check_execute(self):
        global Project_name
        Project_name = self.project_name.get()
        if Project_name:
            aardvark_status = check_Aardvark()
            if aardvark_status == 1:
                global action, how_many_zones
                self.check_status()
                action = function_set
                how_many_zones = len(DIMM_zone_list)
                th.Thread(target=self.Execute).start()
            elif Stop_testing ==False:
                pass
            elif aardvark_status ==0:
                self.mesgbox_warning("No Aardvark Found")
            else:
                self.mesgbox_warning("Aardvark is in use, please replug Aardvark")
        else:
            self.mesgbox_warning("Please choose or enter your project name.")
    
    def Execute(self):

        handle = Aardvark_setup_connect(port, BUS_TIMEOUT, bitrate)
        global Keep_going, Stop_testing
        global config, data_temp,output_data
        Stop_testing =False
        self.execute_BT['state'] = DISABLED
        self.monitor_BT['state'] = DISABLED
        self.save_BT['state'] = DISABLED
        self.clear_BT['state'] = DISABLED
        self.enable_TTB_BT['state'] = DISABLED
        self.stop_BT['state'] = NORMAL
        self.skip_BT['state'] = NORMAL
        
        
        

        for DIMM_Z in DIMM_zone_list:
            if Stop_testing == True:
                    break
            if config % 2 == 0:
                self.tree_config ="config_1"
            else:
                self.tree_config ="config_2"
            config += 1

            if DIMM_Z in DIMM_zone[::2]:  #odd DIMM Zone
                DIMM_addr_list_temp = DIMM_address_list
            else:
                DIMM_addr_list_temp = DIMM_address_list[::-1]

            for addr in DIMM_addr_list_temp:
                if Stop_testing == True:
                    break
                Keep_going =True
                
                times = 0
                data_temp_len = len(data_temp)
                current_time = time.strftime('%x %X')
                data_temp[data_temp_len] = initial_data_dict(DIMM_Z,addr)
                
                addr = int(addr,16)
                a= self.tree.insert("","end",data_temp_len,text = current_time,
                                 value=(data_temp[data_temp_len] ['DIMM Zone'],
                                        data_temp[data_temp_len] ['Address']),tag=(self.tree_config))
                self.tree.see(a)
                
                while Keep_going:
                    (count, data_read) = i2c_read(handle, addr, item, length)
                    times+=1
                    self.master.after(execute_time_gap)
                    if count == length:
                        if data_read[24] == 180 and data_read[25] == 190:
                            break
                        else:
                            try:
                                if times_wierd > 5:
                                    break
                                elif data_read == b:
                                    times_wierd +=0
                                else:
                                    b = data_read
                                    times_wierd = 0
                            except:
                                b = data_read
                                times_wierd = 0
                    if times >= 5:

                        self.tree.item(data_temp_len,
                        value=(data_temp[data_temp_len] ['DIMM Zone'],
                                        data_temp[data_temp_len] ['Address'],
                                        'Read',times,"times"),tag='config_3')

                if count == length:
                    (raw_data, data_result) = data_conversion(data_read)
                    
                if function_set == "write":
                    if two_levels_status == "1" and hex(addr) in DIMM_address_two_power_level:
                        Need_power_temp = Need_power[1]
                    else:
                        Need_power_temp = Need_power[0]
                            
                    while Keep_going:
                        PWM = get_PWM(float(data_result[0]),float(Need_power_temp))
                        i2c_write(handle, addr, item, PWM,persistent_storage_status)
                        
                        while Keep_going:
                            self.master.after(execute_time_gap)
                            (count, data_read) = i2c_read(handle, addr, item, length)
                            times+=1
                            if count == length: 
                                if data_read[24] == 180 and data_read[25] == 190:
                                    break
                                else:
                                    try:
                                        if data_read == b:
                                            break
                                    except:
                                        b = data_read
                                    
                        if PWM == data_read[0] ==data_read[1] ==data_read[2] ==data_read[3]:
                            break
                    if count == length:
                        (raw_data, data_result) = data_conversion(data_read)
                if count == length:
                    data_temp[data_temp_len]['Time'] = current_time
                    data_temp[data_temp_len]['Raw data'] = raw_data
                    data_temp[data_temp_len]['Voltage'] = data_result[0]
                    data_temp[data_temp_len]['Tsensor0'] = data_result[1]
                    data_temp[data_temp_len]['Tsensor1'] = data_result[2]
                    data_temp[data_temp_len]['Tsensor2'] = data_result[3]
                    data_temp[data_temp_len]['Tsensor3'] = data_result[4]
                    data_temp[data_temp_len]['Tsensor4'] = data_result[5]
                    data_temp[data_temp_len]['Tsensor5'] = data_result[6]
                    data_temp[data_temp_len]['Tsensor6'] = data_result[7]
                    data_temp[data_temp_len]['Tsensor Max.'] = data_result[8]
                    data_temp[data_temp_len]['Power'] = data_result[9]
                    
                    if function_set == 'read':
                        output_data[DIMM_Z,hex(addr)] = data_temp[data_temp_len]
                    
                    self.tree.item(data_temp_len,text = current_time,
                                 value=(
                                        data_temp[data_temp_len] ['DIMM Zone'],
                                        data_temp[data_temp_len] ['Address'],
                                        data_temp[data_temp_len] ['Voltage'],
                                        data_temp[data_temp_len] ['Tsensor0'],
                                        data_temp[data_temp_len] ['Tsensor1'],
                                        data_temp[data_temp_len] ['Tsensor2'],
                                        data_temp[data_temp_len] ['Tsensor3'],
                                        data_temp[data_temp_len] ['Tsensor4'],
                                        data_temp[data_temp_len] ['Tsensor5'],
                                        data_temp[data_temp_len] ['Tsensor6'],
                                        data_temp[data_temp_len] ['Tsensor Max.'],
                                        data_temp[data_temp_len] ['Power']
                                     ),tag=(self.tree_config))
                    
                    save_temp_data(data_temp_len)
            
            try:
                if Stop_testing == False:
                    self.mesgbox_info("Please switch to next DIMM zone " +
                                        DIMM_zone_list[DIMM_zone_list.index(DIMM_Z)+1])
            except:
                pass
        Stop_testing =True
        self.execute_BT['state'] = NORMAL
        self.monitor_BT['state'] = NORMAL
        self.save_BT['state'] = NORMAL
        self.clear_BT['state'] = NORMAL
        self.enable_TTB_BT['state'] = NORMAL
        self.stop_BT['state'] = DISABLED
        self.skip_BT['state'] = DISABLED
        Aardvark_disconnect(handle)
        save_default()
        update_usage_data_to_sharepoint()
        
    def Stop(self):
        global Stop_testing, Keep_going
        Stop_testing = True
        Keep_going = False
        
    def Skip(self):
        global Keep_going
        Keep_going = False
        #print(self.tree.get_children())
        #print("%03x " %(len(self.tree.get_children())))
        #print(self.frame.winfo_reqwidth())
        #print(self.frame.winfo_reqheight())
        
    def enable_ttv_test_board(self):
        aardvark_status = check_Aardvark()
        if aardvark_status == 1:
            handle = Aardvark_setup_connect(port, BUS_TIMEOUT, bitrate)
            data_read = to_active_ttv_test_board(handle)
            Aardvark_disconnect(handle)
            if data_read:
                self.mesgbox_info("DDR5 TTV test board enabled")
            else:
                 self.mesgbox_warning("No DDR5 TTV test board found")
        elif Stop_testing ==False:
            pass
        elif aardvark_status ==0:
            self.mesgbox_warning("No Aardvark Found")
        else:
            self.mesgbox_warning("Aardvark is in use, please replug Aardvark")
    
    def Save(self):
        global item_name,action
        action = 'Save'
        filename=filedialog.asksaveasfilename(initialfile = item_name,defaultextension=".csv")
        n = 0
        if filename: 
            item_name_temp = filename[filename.rfind('/')+1:filename.rfind('.csv')]
            if item_name_temp[:14] == 'DDR5_TTV_Data_':
                item_name = item_name_temp
                for i in range(len(item_name)-1,0,-1):
                        try:
                            int(item_name[i])
                            n +=1
                        except:
                            break
                try:
                    item_name = item_name[:-n]+ str(int(item_name[-n:])+1)
                except:
                    pass
                save_default()
                update_usage_data_to_sharepoint()
            
            try:
                with open(filename, 'w', newline='') as f:
                        writer = csv.DictWriter(f, fieldnames=data_label)
                        writer.writeheader()
                        for i in output_data:
                            writer.writerow(output_data[i])
                        f.close()
            except:
                self.mesgbox_warning("Can't save because the file is opening.")
            
    def check_DIMM_zone(self):
        global DIMM_zone_list
        DIMM_zone_list = []
        if self.auto_mode_var.get() == '1':
            for DIMM_n in range(len(DIMM_zone)):
                DIMM_zone_list.append(self.DIMM_zone[DIMM_n].get())
            a = DIMM_zone_list.count('')
            for i in range(a):
                DIMM_zone_list.remove('')
        elif self.function.get() == 'write':
            DIMM_zone_list.append(' ')
        else:
            DIMM_zone_list.append(self.DIMM_zone_R.get())
        return DIMM_zone_list
            
    def check_DIMM_address(self):
        global DIMM_address_list
        self.DIMM_address_get =[]
        for i in range(len(DIMM_address_order)):
            self.DIMM_address_get.append(self.DIMM_address[i].get())
        DIMM_address_list = self.DIMM_address_get
        a = self.DIMM_address_get.count('')
        for i in range(a):
            self.DIMM_address_get.remove('')
        return self.DIMM_address_get
        
    def check_two_power_level(self):
        global two_levels_status
        two_levels_status = '0'
        if auto_mode_status == '1' and function_set == 'write' and self.two_power_level_var.get() == "1":
            self.Need_power_and.grid()
            self.Need_power_entry_2.grid()
            two_levels_status = self.two_power_level_var.get()
            #print(self.master.cget('bg'))
            for i in range(len(DIMM_address_order)):
                if DIMM_address_order[i] in DIMM_address_two_power_level:
                    self.DIMM_address_CB[i].config(background ='gray60')
        else:
            self.Need_power_and.grid_remove()
            self.Need_power_entry_2.grid_remove()
            for i in range(len(DIMM_address_order)):
                if DIMM_address_order[i] in DIMM_address_two_power_level:
                    self.DIMM_address_CB[i].config(bg='SystemButtonFace')
    
    def clear_log(self):
        if Stop_testing == True:
            initial_output_data()
            for i in self.tree.get_children():
                self.tree.delete(i)
    
    def power_validate(self, input):
        #print(self.Need_power_entry.get())
        try:
            if float(input) <= power_limit:
                return True
            else:
                return False
        except:
            if input == "":
                return True
            else:
                return False
    
    def check_monitor(self):
        global Project_name
        Project_name = self.project_name.get()
        if Project_name:
            aardvark_status = check_Aardvark()
            if aardvark_status == 1:
                global action
                action = 'Monitor'
                th.Thread(target=self.Monitor).start()
                #mp.Process(target=plot_test).start()
                mp.Process(target=monitor_plot).start()
            elif Stop_testing ==False:
                pass
            elif aardvark_status ==0:
                 self.mesgbox_warning("No Aardvark Found")
            else:
                self.mesgbox_warning("Aardvark is in use, please replug Aardvark")
        else:
            self.mesgbox_warning("Please choose or enter your project name.")
    
    def Monitor(self):
        global Stop_testing, config
        global data_temp,output_data,how_many_zones
        Stop_testing =False
        self.execute_BT['state'] = DISABLED
        self.skip_BT['state'] = DISABLED
        self.save_BT['state'] = DISABLED
        self.clear_BT['state'] = DISABLED
        self.enable_TTB_BT['state'] = DISABLED
        self.monitor_BT['state'] = DISABLED
        self.stop_BT['state'] = NORMAL
        handle = Aardvark_setup_connect(port, BUS_TIMEOUT, bitrate)
        monitor_times= 0
        how_many_zones = monitor_update_times
        
        while Stop_testing is False:
            self.master.after(monitor_time_gap)
            #time.sleep(monitor_time_gap/1000)
            monitor_times +=1
            if config % 2 == 0:
                self.tree_config ="config_1"
            else:
                self.tree_config ="config_2"
            config += 1
            
            for addr in self.check_DIMM_address():
                if Stop_testing == True:
                    break
                times =0
                #data_local = []
                data_temp_len = len(data_temp)
                current_time = time.strftime('%x %X')
                data_temp[data_temp_len] = initial_data_dict('',addr)
                
                addr = int(addr,16)
                a= self.tree.insert("","end",data_temp_len,text = current_time,
                                 value=(data_temp[data_temp_len] ['DIMM Zone'],
                                        data_temp[data_temp_len] ['Address']),tag=(self.tree_config))
                #self.tree.see(a)
                while times <= 5:
                    #print(Keep_going)
                    (count, data_read) = i2c_read(handle, addr, item, length)
                    times+=1
                    #print("execute1", times, " times")
                    #print(count,data_read)
                    self.master.after(execute_time_gap)
                    if count == length:
                        if data_read[24] == 180 and data_read[25] == 190:
                            break
                        else:
                            try:
                                if data_read == b:
                                    break
                            except:
                                b = data_read
                if count == length:
                    (raw_data, data_result) = data_conversion(data_read)
                    
                    data_temp[data_temp_len]['Time'] = current_time
                    data_temp[data_temp_len]['Raw data'] = raw_data
                    data_temp[data_temp_len]['Voltage'] = data_result[0]
                    data_temp[data_temp_len]['Tsensor0'] = data_result[1]
                    data_temp[data_temp_len]['Tsensor1'] = data_result[2]
                    data_temp[data_temp_len]['Tsensor2'] = data_result[3]
                    data_temp[data_temp_len]['Tsensor3'] = data_result[4]
                    data_temp[data_temp_len]['Tsensor4'] = data_result[5]
                    data_temp[data_temp_len]['Tsensor5'] = data_result[6]
                    data_temp[data_temp_len]['Tsensor6'] = data_result[7]
                    data_temp[data_temp_len]['Tsensor Max.'] = data_result[8]
                    data_temp[data_temp_len]['Power'] = data_result[9]
                    
                    self.tree.item(data_temp_len,text = current_time,
                                 value=(
                                        data_temp[data_temp_len] ['DIMM Zone'],
                                        data_temp[data_temp_len] ['Address'],
                                        data_temp[data_temp_len] ['Voltage'],
                                        data_temp[data_temp_len] ['Tsensor0'],
                                        data_temp[data_temp_len] ['Tsensor1'],
                                        data_temp[data_temp_len] ['Tsensor2'],
                                        data_temp[data_temp_len] ['Tsensor3'],
                                        data_temp[data_temp_len] ['Tsensor4'],
                                        data_temp[data_temp_len] ['Tsensor5'],
                                        data_temp[data_temp_len] ['Tsensor6'],
                                        data_temp[data_temp_len] ['Tsensor Max.'],
                                        data_temp[data_temp_len] ['Power']
                                     ),tag=(self.tree_config))
                    
                    save_temp_data(data_temp_len)
                self.tree.see(a)
            
            if monitor_times >= monitor_update_times:
                monitor_times =0
                print(monitor_times)
                update_usage_data_to_sharepoint()
        how_many_zones = monitor_times
        update_usage_data_to_sharepoint()
        Aardvark_disconnect(handle)
        self.execute_BT['state'] = NORMAL
        self.monitor_BT['state'] = NORMAL
        self.save_BT['state'] = NORMAL
        self.clear_BT['state'] = NORMAL
        self.enable_TTB_BT['state'] = NORMAL
        self.skip_BT['state'] = DISABLED
        self.stop_BT['state'] = DISABLED
    
class new_window:
    def __init__(self,frame,master):
        self.master = master
        self.frame = frame
        self.frame.state(['disabled']) 
        self.TP = Toplevel(self.frame)
        #self.TP.attributes('-topmost', True)
        self.TP.geometry('380x150+700+300')
        self.TP.transient(self.frame)
        self.TP.title('Identity Verification')
        #self.TP.config(bg="white")
        self.TP.grid()
        self.TP.grab_set()
        self.TP.resizable(FALSE,FALSE)
        
        self.email_account = StringVar()
        self.password = StringVar()
        ttk.Label(self.TP,text="E-mail",font = ("Times",14)).grid(column=0,row=0,pady=15)
        Entry(self.TP, justify='center',width =20,textvariable=self.email_account).grid(column=1,row=0,columnspan=2)
        ttk.Label(self.TP,text="@wistron.com",font = ("Times",14)).grid(column=3,row=0)
        self.email_account.set(Email[:Email.find('@')])
        
        ttk.Label(self.TP,text="Password",font = ("Times",14)).grid(column=0,row=1,padx=15)
        Entry(self.TP, justify='center',width =20,textvariable=self.password, show='*').grid(column=1,row=1,columnspan=2)
        Button(self.TP,text="Log in",font=("Times",13),command = self.log_in_button,
               width =15).grid(column=0,row=3,columnspan=2,sticky="N S",pady = 15)
        self.TP.bind("<Return>", lambda event:self.log_in_button())
        
        self.login_status = Label(self.TP,text="Please Try Again",fg='red',font = ("Times",14))
        self.login_status.grid(column=2,row=3,columnspan=2)
        self.login_status.grid_remove()
        self.TP.protocol("WM_DELETE_WINDOW",self.del_win)
        self.frame.grid_remove()
        '''
        self.label = ttk.Label(self.TP,text="This is going to vanish!!!!")
        self.TP.protocol("WM_DELETE_WINDOW",self.del_win)
        self.BT = ttk.Button(self.TP,text= "OK",command = self.del_win)
        self.BT.grid()
        self.BT.focus_set()
        '''
       
        #self.plot()
    def log_in_button(self):
        email = self.email_account.get()+'@wistron.com'
        password = self.password.get()
        if wistron_sharepoint_login(email,password):
            global Email, Password
            Email = email
            Password = password
            save_default()
            DDR5_DIMM_TTV_Toolkit.get_projecet_list(Toolkit)
            #print(Email,Password)
            self.TP.destroy()
            self.frame.grid()
            
        else:
            self.login_status.grid()
    
    def del_win(self):
        self.master.destroy()

    # def plot(self):
  
    #     # the figure that will contain the plot
    #     self.fig = Figure(figsize = (5, 5),
    #                  dpi = 100)
      
    #     # list of squares
    #     y = [i**2 for i in range(101)]
      
    #     # adding the subplot
    #     plot1 = self.fig.add_subplot(111)
      
    #     # plotting the graph
    #     plot1.plot(y)
      
    #     # creating the Tkinter canvas
    #     # containing the Matplotlib figure
    #     self.canvas = FigureCanvasTkAgg(self.fig,master = self.TP)  
    #     self.canvas.draw()
      
    #     # placing the canvas on the Tkinter window
    #     self.canvas.get_tk_widget().pack()
      
    #     # creating the Matplotlib toolbar
    #     self.toolbar = NavigationToolbar2Tk(self.canvas,
    #                                    self.TP)
    #     self.toolbar.update()
      
    #     # placing the toolbar on the Tkinter window
    #     self.canvas.get_tk_widget().pack()


'''main script'''

if __name__ == '__main__':
	mp.freeze_support()


if __name__ == "__main__":
    if not os.path.isdir(path):
        os.mkdir(path)
        first_use = True
    else:
        first_use = False
        
    root = Tk()
    s = ttk.Style()
    
    if root.getvar('tk_patchLevel')=='8.6.9': #and OS_Name=='nt':
        def fixed_map(option):
            # Fix for setting text colour for Tkinter 8.6.9
            # From: https://core.tcl.tk/tk/info/509cafafae
            #
            # Returns the style map for 'option' with any styles starting with
            # ('!disabled', '!selected', ...) filtered out.
            #
            # style.map() returns an empty list for missing options, so this
            # should be future-safe.
            return [elm for elm in s.map('Treeview', query_opt=option) if elm[:2] != ('!disabled', '!selected')]
        s.map('Treeview', foreground=fixed_map('foreground'), background=fixed_map('background'))   
    
    initial_output_data()
    get_user_name()
    get_default()
    Toolkit = DDR5_DIMM_TTV_Toolkit(root)
    root.mainloop()
