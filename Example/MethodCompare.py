import sys
import os

# Get the absolute path of the parent directory
# This automatically finds the 'My Python Modules for CC Method Comp' folder
parent_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))

# Add that directory to Python's search path
sys.path.insert(0, parent_dir)

# Now you can import your module exactly as you wanted
import CC_Method_Analysis_v_0_1 as CC



######## Error information #################
Error_level_cut_off = None # cut-off between absolute and percentage error
error1 = None # absolute error limit below the cut-off
error2 = 15 # percentage error limit above the cut-off
############################################

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

document.save('MethodComparison.docx')

'''
Method comparison begins here; loading data into dataframe, df_MC
'''
Input_file_name = 'NT-proBNP_.xlsx'
df_MC = pd.read_excel(Input_file_name, 
                      sheet_name='Method Comparison', 
                      usecols='B:G', 
                      skiprows=range(0, 43))

#if Test_Analyte == Test_Analyte:
#    Test_Analyte = Test_Analyte + " (sn: " + Test_Instrument_sn +")"


df_MC.rename(columns={'Primary Refence Method Result': f"{Reference_Analyte}", "Test Method Result":Test_Analyte}, inplace=True)

df_MC