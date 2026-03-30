import pandas as pd
import CC_Method_Analysis_v_0_1 as CC
from docx import Document 
from docx.shared import Inches, Cm, Pt
CC.manage_figures_folder() # delete and recreate the Figures folder within the current directory

#______________Test and Error information____________________________
Input_file_name = 'NT-proBNP_.xlsx'
Test_Analyte = 'A1c'
Test_Unit = '%'

Error_level_cut_off = None # cut-off between absolute and percentage error
error1 = None # absolute error limit below the cut-off
error2 = 7 # percentage error limit above the cut-off

#___________________________________________________________

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

#read the data from the excel file
df_MC = pd.read_excel(Input_file_name, 
                      sheet_name='Method Comparison', 
                      usecols='B:G', 
                      skiprows=range(0, 43))

# Call the method comparison function
CC.MC_output(analyte=Test_Analyte,
            document = Document("MethodComparison.docx"), 
             x = df_MC.iloc[:,2], 
             y = df_MC.iloc[:,3], 
             z = None,
             z2 = df_MC['Day'],
             Unit = Test_Unit, 
             Error_level_cut_off = Error_level_cut_off, 
             error1 = error1, 
             error2 = error2)