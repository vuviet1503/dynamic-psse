import os, sys
import pandas as pd
import numpy as np
import psse35



def excel(file_path):
    # Open thi Excel file
    xls = pd.ExcelFile(file_path)
    # Parse the Excel file
    bus = pd.read_excel(xls, 'BUS', skiprows= 1)
    line = pd.read_excel(xls, 'LINE', skiprows= 1)
    
    return bus, line

def psse_init(bus, line):
    # Import the PSSE module
    import psspy

    # Initialize PSSE
    ierr = psspy.psseinit(50)
    ierr = psspy.newcase_2(basemva = 100, basefreq = 50)

    # Add buses
    for i in range(bus.shape[0]):
        # Code
        if np.isnan(bus['CODE'][i]):
            bus.loc[i, 'CODE']= 1
        
        ierr = psspy.bus_data_4(ibus = int(bus['NO'][i]), 0, intgar1 = int(bus['CODE'][i]), realar1 = float(bus['kV'][i]), name = str(bus['NAME'][i]))

    return ierr



if __name__ == '__main__': 
    file_path = r'C:\Users\Admin\Desktop\code\SF2024\SF_2041_inputs\Input7bus.xlsx'
    bus, line = excel(file_path)
    psse_init(bus, line)
    # print(bus.shape[1])

