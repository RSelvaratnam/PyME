import sys
import os
import CC_Method_Analysis_v_0_1 as CC


#______________Test and Error information____________________________

Test_Analyte = 'Nt-proBNP'
Test_Unit = 'ng/L'
Input_file_name = 'Input_file.xlsx'
Output_file_name = 'MethodComparison.docx'

Error_level_cut_off = None # cut-off between absolute and percentage error
error1 = None # absolute error limit below the cut-off
error2 = 15 # percentage error limit above the cut-off
#___________________________________________________________

from docx import Document 
from docx.shared import Inches, Cm, Pt
import pandas as pd

#open a blank document
document = Document()

#adjust page margins
sections = document.sections
for section in sections:
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)

document.save(Output_file_name)

'''
Method comparison begins here; loading data into dataframe, df_MC
'''

print("Recreate Figures folder...", flush=True)
CC.manage_figures_folder() # delete and recreate the Figures folder

print("Reading Data...", flush = True)
df_MC = pd.read_excel(Input_file_name, 
                      sheet_name='Method Comparison', 
                      usecols='B:G', 
                      skiprows=range(0, 43))


print("Generating Method Comparison Report...")
CC.MC_output(analyte=Test_Analyte,
            document = Document(Output_file_name), 
             x = df_MC.iloc[:,2], 
             y = df_MC.iloc[:,3], 
             z = None,
             z2 = df_MC['Day'],
             Unit = Test_Unit, 
             Error_level_cut_off = Error_level_cut_off, 
             error1 = error1, 
             error2 = error2)