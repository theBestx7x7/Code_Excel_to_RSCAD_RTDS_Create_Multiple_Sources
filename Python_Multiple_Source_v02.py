# -*- coding: utf-8 -*-
"""
Created on Tue Dec 10 00:29:06 2024

@author: Franz Guzman
"""

import pandas as pd
import rtds.rscadfx
import time

# Path to access the Excel file
excel_path = "C:/One Drive Franz/OneDrive/1. ESTUDIOS - RODRIGO/1. UNICAMP - M. INGENIERÍA ELÉCTRICA/0. Cursos y Capacitaciones/0. Curso_RTDS_2024_UNICAMP/1. Aula 01/Check_List_RTDS_2.xlsx"  # Replace with the path to your file

# Read a specific sheet from the Excel file
one_sheet = pd.read_excel(excel_path, sheet_name='GERADORES') 

# Start and end points for the table data
init_row, init_col = (7, 13)  # Starting point of the data table
end_row, end_col = (one_sheet.shape[0], one_sheet.shape[1]-3)  # End point of the data table
len_row = end_row - init_row
len_col = end_col - init_col
print("Dimension of one_sheet:", one_sheet.shape)  # Dimensions of rows and columns in the sheet
print("Starting point: ", init_row, " - ", init_col, "; Ending point: ", end_row, " - ", end_col)

# Read only specific columns from the sheet (columns 13 to 34)
input_cols = [i for i in range(init_col, end_col)]  # Range from column 13 to 34
print(input_cols)

# Load the data file again using the selected columns
df = pd.read_excel(excel_path, sheet_name='GERADORES', header=init_row, usecols=input_cols)  # Load data to read specific columns
print("Column names: ", df.columns)
print("New dimension of df: ", df.shape)

# =============================================================================
# Column names for the SOURCE component required to upload data in RSCAD
col_name_SRC = [
    "Name", "ZSeq", "ZType", "R1s", "R1p", "L1p", "R0p", "L0p",
    "Es", "F0", "Ph", "Imon", "IAnam", "IBnam", "ICnam",
    "srcBrk", "swdnm", "Pmon", "Qmon", "Pnam", "Qnam"
] 
# =============================================================================

# Open a connection to RSCAD FX from the script.
# Any code executed within the scope of the following statement
# Will be run while connected to RSCAD FX.

with rtds.rscadfx.remote_connection() as app:

    # Open the case file
    case_id1 = app.open_case(r"C:\Users\Franz Guzman\OneDrive\2. TRABAJO - RODRIGO\Archivos - Auxiliares\Documents\RSCAD\RTDS_USER_FX\fileman\6. Sexta_Aula\Teste_2_Python.rtfx")
    Teste_2_Python = case_id1.draft.get_subpage("SS #1")
    
    # Reference coordinates of the element
    elements_in_line = 0  # Counter for the number of elements (sources) in a row
    x_position, y_position = 176, 100  # Initial position to create a Source in RSCAD
    delta_x, delta_y = 192, 100  # Spacing between elements (sources)

    # Load data row by row
    for data in range(df.shape[0]): 
        row = df.iloc[data]  # Retrieve data for each row during the iteration
        print(f"Source {data}:")
        i = 0  # Counter for the number of columns (parameters for sources)
 
        # Create Source (i_data) in RSCAD
        lf_rtds_sharc_sld_SRC_id = Teste_2_Python.insert_component("lf_rtds_sharc_sld_SRC", x_position, y_position)  # Create source at specified coordinates
        SRC_id = lf_rtds_sharc_sld_SRC_id.unique_id  # Identify the element's unique ID       
        lf_rtds_sharc_sld_SRC_id = case_id1.get_object(SRC_id)  # Select the component by ID
        
        # Adjust spacing between elements
        x_position = x_position + delta_x 
        elements_in_line += 1 
        
        if elements_in_line == 8:  # Ensure a maximum of 8 elements per row
            y_position = y_position + delta_y # Ensure the height elements per row
            elements_in_line = 0
            x_position = 176

        for col_name, value in row.items(): # Get data from rows, which are saved in 'col_name' and 'value'
            col_name_str = str(col_name).strip()  # Convert to string and remove spaces
            col_name_SRC_str = str(col_name_SRC[i]).strip()  # Convert to string and remove spaces
            value_str = str(value).strip()  # Convert to string and remove spaces

            if col_name_str == col_name_SRC_str:  # Compare column names in RSCAD and Excel
                print(f"  {col_name_SRC_str}: {value_str}" , ' -> ok ')
                # Set Source parameters in RSCAD using data from Excel
                lf_rtds_sharc_sld_SRC_id.set_parameter(col_name_SRC_str, value_str)  # Assign the value to the parameter
            else:
                print(f"The column name {col_name} is different from {col_name_SRC[i]}")
                print(f"Please change the column name in Excel from: {col_name} to -> {col_name_SRC[i]}")
            i += 1
    
    print(".....Data for SOURCES has been loaded successfully.....")
