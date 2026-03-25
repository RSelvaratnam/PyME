import CC_Method_Analysis_v_0_1 as CC

Input_file_name = 'NT-proBNP_.xlsx'

######## Error information #################
Error_level_cut_off = None # cut-off between absolute and percentage error
error1 = None # absolute error limit below the cut-off
error2 = 15 # percentage error limit above the cut-off
############################################

#from docx import Document
#from docx.shared import Inches, Cm, Pt

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
