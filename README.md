# Readme
This is a Python based Method evaluation (PyMe) module for comparing two test methods. The main PyMe file consists of some crude function utilized to run method comparison studies. By no means, they are perfect but do get the task done.

Many of the functions have various dependencies. 
These include, pandas, numpy, uncertainties, scipy, statsmodels, pingouin, python-docx, PIL, statsmodels, and plotly, pygam, and few others dedicated to running these on the Windows.

## Example

The example contains an excel file of studies comparing NT-proBNP results from the Abbott Alinity i instrument versus that obtained on the Quidel Ortho Vitros 5600 method.

The example also contains *MethodCompare.py* which uses the excel file *NT-proBNB_.xlsx* to read the data in and output a word document report called *MethodComparison.docx*.

The example illustrates use of the **`MC_output'** function, which takes in the reference and test method and does regression and statistics, with the output as a report in word document format

To run the script in an appropiate python environment, do the following
1. Navigate to the directory where the .py files and input file is located.  
2. Type python MethodCompare.py

You should get a report out witin a few seconds called "MethodComparison.docx"

## Explainations of Functions
