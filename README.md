# Readme

PyME is short for Python for Method Evaluation. PyMe module consists of some crude function utilized to run method comparison studies. By no means, they are perfect but do get the task done.

Many of the functions have various dependencies. 
These include, pandas, numpy, uncertainties, scipy, statsmodels, pingouin, python-docx, PIL, statsmodels, and plotly, pygam, and few others dedicated to running these on the Windows.

## Example

### Method Comparison
The example folder contains an excel file of studies comparing NT-proBNP results from the Abbott Alinity i instrument versus that obtained on the Quidel Ortho Vitros 5600 method.

The example also contains *MethodCompare.py* which uses the excel file to read the data in and output a word document report.

The example illustrates use of the **`MC_output'** function, which takes in the reference and test method and does regression and statistics, with the output as a report in word document format

## Explainations of Functions
