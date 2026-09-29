import sys
import os
import CC_Method_Analysis_v_0_1 as CC
from docx import Document 
from docx.shared import Inches, Cm, Pt
import pandas as pd


#______________Test and Error information____________________________

Test_Analyte = input("Enter name of the measurand: ").strip()
Test_Unit = input("What is the unit of measure? ").strip()

Output_file_name = 'MethodComparison.docx'

Error_level_cut_off = 3 # cut-off between absolute and percentage error
error1 = 20 # absolute error limit below the cut-off
error2 = 15 # percentage error limit above the cut-off
#___________________________________________________________

def get_input_file_name():
    """ Ask the user to input the name of the input file and check if it exists. """

    print(f"Checking directory: {os.getcwd()}")
    xlsx_files = [f for f in os.listdir() if f.lower().endswith('.xlsx')]

    if xlsx_files:
        print(f"Found the following .xlsx files: {', '.join(xlsx_files)}")
    else:
        print("No .xlsx files found in the current directory.")
        sys.exit(1)

    while True:
        Input_file_name = input("\nPlease enter the name of the input file (e.g., Input_file.xlsx): ").strip()
        #remove quotes if user drags file into terminal
        Input_file_name = Input_file_name.strip('"').strip("'")

        if not Input_file_name:
            print("Please enter a valid file name.")
            continue

        # auto-add extension if not provided
        if not Input_file_name.lower().endswith('.xlsx'):
            Input_file_name += '.xlsx'

        if os.path.isfile(Input_file_name):
            print(f"File '{Input_file_name}' found.")
            return Input_file_name
        else:
            print(f"File '{Input_file_name}' does not exist. Please try again.")

# Get and validate input file
Input_file_name = get_input_file_name().strip()


#open a blank document
document = Document()
document.save(Output_file_name)

'''
Method comparison begins here; loading data into dataframe, df_MC
'''

print("Recreate Figures folder...", flush=True)
CC.manage_figures_folder() # delete and recreate the Figures folder

print("Reading Data...", flush = True)
df_MC = pd.read_excel(Input_file_name, 
                      sheet_name='Accuracy', 
                      usecols='B:G', 
                      skiprows=range(0, 16))


print("Generating Method Comparison Report...")
CC.MC_output(analyte=Test_Analyte,
            document = Document(Output_file_name), 
             x = df_MC.iloc[:,3], 
             y = df_MC.iloc[:,4], 
             z = None,
             z2 = None,
             Unit = Test_Unit, 
             Error_level_cut_off = Error_level_cut_off, 
             error1 = error1, 
             error2 = error2)