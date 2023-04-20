import os
os.system("pip install -r requirements.txt")
import xlrd
import numpy as np 
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog

# Funciton to read data from input file and produce time and RPM plots
def run_function(file_path):

    # Give the location of the file 
    loc = file_path

    # To open Workbook 
    wb = xlrd.open_workbook(loc) 
    sheet = wb.sheet_by_index(0) 

    # Skip the first row 
    row = 1

    # Get the number of columns
    cols = sheet.ncols

    # Create arrays for each column
    arrays = []
    for i in range(cols):
        arrays.append([])

    # Iterate through each row and store values in each array
    while(row < sheet.nrows):
        for i in range(cols):
            arrays[i].append(sheet.cell_value(row, i))
        row += 1

    # Designate corresponding arrays, array name corresponds to column header
    SeqID = arrays[0]
    TimeStamp = arrays[1]
    Channel = arrays[2]
    Direction = arrays[3]
    FrameID = arrays[4]
    FrameType = arrays[5]
    FrameFormat = arrays[6]
    Length = arrays[7]
    FrameData = arrays[8]
    Message = arrays[9]

    # Get the length of each array
    lengths = []
    for array in arrays:
        lengths.append(len(array))

    # Generate decimal value of CAN_ID hex value so that can be matched up with table in TJ40 CAN Manual
    CAN_ID_dec = np.array([int(x, 16) for x in FrameID])

    # Preallocate array for each parameter and its time stamp in TJ40 CAN Manual table
    berModule_Status = []
    berModule_Status_time = []

    ECU_Mode = []
    ECU_Mode_time = []

    RC_pulseWidth = []
    RC_pulseWidth_time = []

    func_Warnings = []
    func_Warnings_time = []

    func_Errors = []
    func_Errors_time = []

    HW_Errors = []
    HW_Errors_time = []

    glowPlug_Current = []
    glowPlug_Current_time = []

    fuelPump_Current = []
    fuelPump_Current_time = []

    BEM_status = []
    BEM_status_time = []

    engine_Speed = []
    engine_Speed_time = []

    ECU_Temp = []
    ECU_Temp_time = []

    semiCond_Temp = []
    semiCond_Temp_time = []

    intake_Temp = []
    intake_Temp_time = []

    exhaust_Temp = []
    exhaust_Temp_time = []

    converter_currentOut = []
    converter_currentOut_time = []

    engine_nominalSpeed = []
    engine_nominalSpeed_time = []

    engine_reqSpeed = []
    engine_reqSpeed_time = []

    fuelPump_Speed = []
    fuelPump_Speed_time = []

    discreteInputs_State = []
    discreteInputs_State_time = []

    generator_rectifiedVoltage = []
    generator_rectifiedVoltage_time = []

    supply_Voltage = []
    supply_Voltage_time = []

    controlLever_Position = []
    controlLever_Position_time = []

    # Initialize initial time variable
    time_initial = 0

    # Sort through data and add each parameter to appropriate array
    for i in range(len(CAN_ID_dec)):
        # Isolate data portion and remove spaces
        msg = FrameData[i]
        data = msg[15:]
        data = data.replace(" ", "")
        
        # Go from system time format to time in seconds starting at zero
        time = TimeStamp[i]
        time = time[:8] + time[9:]
        time = float(time)
        if i == 0:
            time_initial = time
        time = time-time_initial

        # Identify parameter and assign to array
        if CAN_ID_dec[i] == 1800:
            berModule_Status.append(data)
            berModule_Status_time.append(time)
        elif CAN_ID_dec[i] == 1804:
            ECU_Mode.append(data)
            ECU_Mode_time.append(time)
        elif CAN_ID_dec[i] == 1808:
            RC_pulseWidth.append(data)
            RC_pulseWidth_time.append(time)
        elif CAN_ID_dec[i] == 1812:
            func_Warnings.append(data)
            func_Warnings_time.append(time)
        elif CAN_ID_dec[i] == 1816:
            func_Errors.append(data)
            func_Errors_time.append(time)
        elif CAN_ID_dec[i] == 1820:
            HW_Errors.append(data)
            HW_Errors_time.append(time)
        elif CAN_ID_dec[i] == 1824:
            glowPlug_Current.append(data)
            glowPlug_Current_time.append(time)
        elif CAN_ID_dec[i] == 1828:
            fuelPump_Current.append(data)
            fuelPump_Current_time.append(time)
        elif CAN_ID_dec[i] == 1832:
            BEM_status.append(data)
            BEM_status_time.append(time)
        elif CAN_ID_dec[i] == 1836:
            engine_Speed.append(data)
            engine_Speed_time.append(time)
        elif CAN_ID_dec[i] == 1840:
            ECU_Temp.append(data)
            ECU_Temp_time.append(time)
        elif CAN_ID_dec[i] == 1844:
            semiCond_Temp.append(data)
            semiCond_Temp_time.append(time)
        elif CAN_ID_dec[i] == 1848:
            intake_Temp.append(data)
            intake_Temp_time.append(time)
        elif CAN_ID_dec[i] == 1852:
            exhaust_Temp.append(data)
            exhaust_Temp_time.append(time)
        elif CAN_ID_dec[i] == 1856:
            converter_currentOut.append(data)
            converter_currentOut_time.append(time)
        elif CAN_ID_dec[i] == 1860:
            engine_nominalSpeed.append(data)
            engine_nominalSpeed_time.append(time)
        elif CAN_ID_dec[i] == 1864:
            engine_reqSpeed.append(data)
            engine_reqSpeed_time.append(time)
        elif CAN_ID_dec[i] == 1868:
            fuelPump_Speed.append(data)
            fuelPump_Speed_time.append(time)
        elif CAN_ID_dec[i] == 1872:
            discreteInputs_State.append(data)
            discreteInputs_State_time.append(time)
        elif CAN_ID_dec[i] == 1876:
            generator_rectifiedVoltage.append(data)
            generator_rectifiedVoltage_time.append(time)
        elif CAN_ID_dec[i] == 1880:
            supply_Voltage.append(data)
            supply_Voltage_time.append(time)
        elif CAN_ID_dec[i] == 1884:
            controlLever_Position.append(data)
            controlLever_Position_time.append(time)

    # Generate decimal values of all hex data (short,long,etc data types only)
    RC_pulseWidth_dec = np.array([int(x, 16) for x in RC_pulseWidth])
    fuelPump_Current_dec = np.array([int(x, 16) for x in fuelPump_Current])
    engine_Speed_dec = np.array([int(x, 16) for x in engine_Speed])
    ECU_Temp_dec =np.array([int(x, 16) for x in ECU_Temp])
    semiCond_Temp_dec = np.array([int(x, 16) for x in semiCond_Temp])
    exhaust_Temp_dec = np.array([int(x, 16) for x in exhaust_Temp])
    engine_reqSpeed_dec = np.array([int(x, 16) for x in engine_reqSpeed])
    fuelPump_Speed_dec = np.array([int(x, 16) for x in fuelPump_Speed])
    discreteInputs_State_dec = np.array([int(x, 16) for x in discreteInputs_State])

    # Generate deciam value of a hex data (float data type)
    glowPlug_Current_dec = float2dec(glowPlug_Current)
    intake_Temp_dec = float2dec(intake_Temp)
    converter_currentOut_dec = float2dec(converter_currentOut)
    engine_nominalSpeed_dec = float2dec(engine_nominalSpeed)
    generator_rectifiedVoltage_dec = float2dec(generator_rectifiedVoltage)
    supply_Voltage_dec = float2dec(supply_Voltage)
    controlLever_Position_dec = float2dec(controlLever_Position)

    # Create time figure and axes
    fig1, axs1 = plt.subplots(2, 3, figsize=(16, 8))

    # EGT v Time plot
    ax0 = axs1[0, 0]
    ax0.plot(exhaust_Temp_time, exhaust_Temp_dec)
    ax0.set_title('Exhaust Temperature')
    ax0.set_xlabel('Time (s)')
    ax0.set_ylabel('Temperature (deg C)')

    # Engine Speed v Time plot
    ax1 = axs1[0, 1]
    ax1.plot(engine_Speed_time, engine_Speed_dec)
    ax1.set_title('Engine Speed')
    ax1.set_xlabel('Time (s)')
    ax1.set_ylabel('Speed (rpm)')

    # Rectified voltage v Time plot
    ax2 = axs1[0, 2]
    ax2.plot(generator_rectifiedVoltage_time, generator_rectifiedVoltage_dec)
    ax2.set_title('Generator Rectified Voltage')
    ax2.set_xlabel('Time (s)')
    ax2.set_ylabel('Voltage (V)')

    # ECU Temp v Time plot
    ax3 = axs1[1, 0]
    ax3.plot(ECU_Temp_time, ECU_Temp_dec)
    ax3.set_title('ECU Temperature')
    ax3.set_xlabel('Time (s)')
    ax3.set_ylabel('Temperature (deg C)')

    # Fuel Pump Speed v Time plot
    ax4 = axs1[1, 1]
    ax4.plot(fuelPump_Speed_time, fuelPump_Speed_dec)
    ax4.set_title('Fuel Pump Speed')
    ax4.set_xlabel('Time (s)')
    ax4.set_ylabel('Speed (rpm)')

    # Supply Voltage v Time plot
    ax5 = axs1[1, 2]
    ax5.plot(supply_Voltage_time, supply_Voltage_dec)
    ax5.set_title('Supply Voltage')
    ax5.set_xlabel('Time (s)')
    ax5.set_ylabel('Voltage (V)')

    # Create RPM figure and axes
    fig2, axs2 = plt.subplots(1, 3, figsize=(16, 8))

    # EGT v Engine Speed plot
    ax7 = axs2[0]
    ax7.plot(engine_Speed_dec, exhaust_Temp_dec)
    ax7.set_title('EGT')
    ax7.set_xlabel('Engine RPM')
    ax7.set_ylabel('Temp (deg C)')

    # Rectified voltage data trasnmitted less frequently so prepare a new engine speed array...
    # that lines up with frequency of rectified voltage
    freq = int(len(engine_Speed_dec)/len(generator_rectifiedVoltage_dec))
    new_engine_Speed_dec = engine_Speed_dec[::freq]
    dif = len(generator_rectifiedVoltage_dec)-len(new_engine_Speed_dec)

    # Rectified Voltage v Engine Speed plot
    ax8 = axs2[1]
    ax8.plot(new_engine_Speed_dec[:dif], generator_rectifiedVoltage_dec)
    ax8.set_title('Generated Voltage')
    ax8.set_xlabel('Engine RPM')
    ax8.set_ylabel('Voltage (V)')

    # Fuel Pump Speed v Engine Speed plot
    ax9 = axs2[2]
    ax9.plot(engine_Speed_dec, fuelPump_Speed_dec)
    ax9.set_title('Fuel Pump')
    ax9.set_xlabel('Engine RPM')
    ax9.set_ylabel('Fuel Pump RPM (RPM)')

    # Display Figures
    plt.show()


# Function to convert an IEEE 754 Single Point FLoat hex value to its decimal value
def float2dec(hex_arr): 
    # HEX TO DECIMAL
    dec_arr = []
    for i in range(len(hex_arr)): 
        hex_value = hex_arr[i]

        #Using binary notation 
        bin_value = bin(int(hex_value, 16))[2:]

        # Prepend 0's to make sure the binary string has correct length
        bin_value = bin_value.zfill(32)

        # Extract mantissa, sign and exponent from binary string
        mantissa = bin_value[9:32]
        sign = int(bin_value[0],2)
        exponent = int(bin_value[1:9],2)

        # Calculate the true exponent (with offset)
        exponent_offset = 127
        true_exponent = exponent - exponent_offset

        # Calculate the mantissa
        mantissa = 1 + int(mantissa, 2) / (2 ** 23)

        # Calculate final result
        result = (-1) ** sign * 2 ** true_exponent * mantissa

        dec_arr.append(result)
    return dec_arr 

# Function to browse for a file in file explorer
def browse_function(entry_field):
    # Open file explorer dialog box
    file_path = filedialog.askopenfilename()

    # Fill text entry with the selected file
    entry_field.delete(0, tk.END)
    entry_field.insert(0, file_path)

# create the window 
window = tk.Tk()

# set title 
window.title('CAN2USB Data GUI')

# set size 
window.geometry('650x200')

# create label at the top 
top_label = tk.Label(window, text='CAN Data Postprocessing', font=("Helvetica", 20, 'bold'))
top_label.pack(padx=20, pady=20)

# create label 
label = tk.Label(window, text='Enter File Path:', font=("Helvetica", 16))
label.pack(side=tk.LEFT, padx=20, pady=30)

# create entry field 
file_path_entry = tk.Entry(window, width=50)
file_path_entry.pack(side=tk.LEFT, padx=20, pady=20)

# create file browse button 
browse_button = tk.Button(window, text='Browse', bg='orange', command=lambda: browse_function(file_path_entry))
browse_button.pack(side=tk.LEFT, padx=0, pady=20)

# create run button 
run_button = tk.Button(window, text='Run', bg='green', command=lambda: run_function(file_path_entry.get()))
run_button.pack(padx=10, pady=20, expand=True, anchor=tk.CENTER)

# start window 
window.mainloop()


