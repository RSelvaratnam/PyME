import os
import csv

import numpy as np
import pandas as pd

#ploting libraries
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from PIL import Image

#word document related libraries
from docx.shared import Inches, Cm, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

from uncertainties import ufloat # required for data_summary function


def FigWhiteCrop_stream(fig_or_pil, margin=10):
    """
    Takes a plotly fig OR a PIL Image, crops white borders, 
    returns a BytesIO ready for document.add_picture()
    """
    # 1. Get a PIL Image
    if hasattr(fig_or_pil, 'write_image'): # it's a plotly fig
        buf = BytesIO()
        # to_image is faster than write_image for memory
        img_bytes = fig_or_pil.to_image(format="png", scale=2)
        buf = BytesIO(img_bytes)
        img = Image.open(buf)
    else:
        img = fig_or_pil # already PIL

    # 2. Your exact crop logic, but on the in-memory image
    img_array = np.array(img)
    
    # Handle RGBA -> use only RGB
    if img_array.ndim == 3:
        rgb = img_array[:, :, :3]
    else:
        rgb = img_array

    # mask where NOT white (any channel != 255)
    non_white_mask = np.any(rgb != 255, axis=2)
    
    if not np.any(non_white_mask):
        # all white, return original
        out = BytesIO()
        img.save(out, format="PNG")
        out.seek(0)
        return out

    rows, cols = np.where(non_white_mask)
    top, bottom = rows.min(), rows.max()
    left, right = cols.min(), cols.max()

    # Clamp to image bounds so margin doesn't go negative
    left = max(0, left - margin)
    top = max(0, top - margin)
    right = min(img.width, right + margin)
    bottom = min(img.height, bottom + margin)

    cropped_img = img.crop((left, top, right, bottom))

    # 3. Return as stream for docx
    out_stream = BytesIO()
    cropped_img.save(out_stream, format="PNG", dpi=(300,300))
    out_stream.seek(0)
    return out_stream

def Histogram_grouped(x, y):
    df = pd.DataFrame({'Reference Method': x, 'Test Method': y})
    df_long = pd.melt(df, value_vars=['Reference Method', 'Test Method'], 
                      var_name='Method', value_name='Result')

    fig = px.histogram(df_long, x='Result', color='Method',
                       histnorm='probability',
                       barmode='group',
                       marginal="box",
                       hover_data=df_long.columns)

    fig.update_layout(
        plot_bgcolor="white", paper_bgcolor="white",
        xaxis=dict(showgrid=True, gridcolor="lightgrey"),
        yaxis=dict(showgrid=True, gridcolor="lightgrey"),
        xaxis_title='Result', yaxis_title='Count (Normalized)',
        bargap=0.2,
        height = 450, width = 900,
    )
    # fig.show() # remove this for report generation

    # No more fig.write_image() / FigWhiteCrop() to disk
    return FigWhiteCrop_stream(fig, margin=10)


def sample_comparions_tableX(x, y, z, **error_info):
    """
    This function generates a table looking at the difference & % difference between the methods.
    Based on error compute if the difference is within acceptable error.

    The error can be defined as total allowable error or per IQMH terminology of allowable performance limits(APL),
    or even precision goal.

    This function is designed to take a single concentration cut-off and two error thresholds.
    One error threshold, the first error1 variable is to be DEFINED IN ABSOLUTE NUMBERS.
    The other error value, error2 is to be DEFINED IN PERCENTAGE TERMS.
    When both error threshold is provided, the concentration_cut MUST BE PROVIDED.

    When concentration cut-off is provided, the error1 term is for absolute error below the cut-off and the
    error2 term is applied above the cut-off

    x = reference method
    y = test method
    z = hue or characteristic or unique to this sample
    """
    error1 = error_info.get("error1", None)
    error2 = error_info.get("error2", None)
    Error_level_cut_off = error_info.get("Error_level_cut_off", None)
  
    if Error_level_cut_off is not None:
        Error_level_cut_off = float(Error_level_cut_off)

    # Create dataframe
    analyte_dataSummary = pd.concat([x, y, z], axis=1).reset_index(drop=True)
    analyte_dataSummary.index = range(1, len(analyte_dataSummary) + 1)
    analyte_dataSummary.index.name = 'Sample ID'
    analyte_dataSummary = analyte_dataSummary.reset_index()

    # Convert to numeric, coerce errors to NaN, because we perform numeric calculations
    analyte_dataSummary.iloc[:,1] = pd.to_numeric(analyte_dataSummary.iloc[:,1], errors='coerce') #reference method
    analyte_dataSummary.iloc[:,2] = pd.to_numeric(analyte_dataSummary.iloc[:,2], errors='coerce') #test method

    # Calculate differences and percent differences
    # Replace 0 with NaN in the denominator to prevent ZeroDivisionError
    safe_denominator = analyte_dataSummary.iloc[:,1].replace(0, float('nan'))
    
    analyte_dataSummary["Y-X"] = analyte_dataSummary.iloc[:,2] - analyte_dataSummary.iloc[:,1]
    analyte_dataSummary["%(Y-X)/X"] = (analyte_dataSummary["Y-X"] / safe_denominator) * 100
    analyte_dataSummary["Y-X"] = pd.to_numeric(analyte_dataSummary["Y-X"], errors='coerce')
    analyte_dataSummary["%(Y-X)/X"] = pd.to_numeric(analyte_dataSummary["%(Y-X)/X"], errors='coerce')
    analyte_dataSummary[["Y-X", "%(Y-X)/X"]] = analyte_dataSummary[["Y-X", "%(Y-X)/X"]].round(3)

    # Apply Pass/Fail criteria
    def check_pass_fail(row):
        # Check for NaN in x.name or Test Method (Y)
        if pd.isna(row[x.name]) or pd.isna(row[y.name]):
            return "Not assessed"

        if Error_level_cut_off is not None:
            # We removed the float cast from here because we already did it at the top of the main function!
            
            if error1 is None or error2 is None:
                raise ValueError(
                    "Please provide both 'error1' and 'error2' if and when you specify 'Error_level_cut_off' value."
                )
            
            # Check if x is numeric before comparing to the float cut-off
            if isinstance(row[x.name], (int, float)):
                if row[x.name] >= Error_level_cut_off:
                    return "**FAIL**" if abs(row["%(Y-X)/X"]) > error2 else "pass"
                else:
                    return "**FAIL**" if abs(row["Y-X"]) > error1 else "pass"
            else:
                # If it's a string (e.g., "<0.05"), it skips the >= comparison entirely
                return "Not assessed" 

        elif Error_level_cut_off is None:
            if error1 is not None and error2 is None:
                # Check if x is numeric before comparing to the float cut-off
                if isinstance(row[x.name], (int, float)):
                    return "**FAIL**" if abs(row["Y-X"]) > error1 else "pass"
                else:
                    return "Not assessed"  # If it's a string (e.g., "<0.05"), it skips the numeric comparison entirely
            elif error2 is not None and error1 is None:
                if isinstance(row[x.name], (int, float)):
                    return "**FAIL**" if abs(row["%(Y-X)/X"]) > error2 else "pass"
                else:
                    return "Not assessed"
            else:
                return "Not assessed"  # When both error1 and error2 are None or both are provided (not handled here)
                
        return "Not assessed"

    analyte_dataSummary["Pass/Fail"] = analyte_dataSummary.apply(check_pass_fail, axis=1)

    df_selected_columns = analyte_dataSummary.iloc[:,1:3] #select only reference and test method columns
    df_selected_columns["Y-X"] = analyte_dataSummary["Y-X"]
    df_selected_columns["%(Y-X)/X"] = analyte_dataSummary["%(Y-X)/X"]

    # Return sorted dataframe
    return analyte_dataSummary.sort_values(by=[x.name]), df_selected_columns
def data_summary(dataAll):
    """
    Returns a professional summary table with automatic, correct significant figures
    using the gold-standard 'uncertainties' package. The method also computes the CV
    
    Mean ± SEM is displayed perfectly formatted with proper scientific rounding.
    Can aways check the results of this with 'df.describe()' and the CV with 'df.std()/df.mean()'
    """
    # Keep only numeric columns
    df = dataAll.select_dtypes(include=[np.number])
    
    # We'll build a list of dicts → clean, readable, and flexible
    rows = []
    
    for col in df.columns:
        data = df[col].dropna()
        n = len(data)
        
        if n == 0:
            continue
        if n == 1:
            # Only one value → SEM undefined
            mean = data.iloc[0]
            sem = np.nan
            sd = np.nan
            cv = np.nan
        else:
            mean = data.mean()
            sd = data.std(ddof=1)
            sem = sd / np.sqrt(n)
            cv = round((sd / mean * 100),2) if mean != 0 else np.nan
        
        # This is the magic line — ufloat handles ALL sig fig rules automatically
        mean_with_unc = ufloat(mean, sem) if n > 1 else mean
        
        rows.append({
            "Variable": col,
            "N": int(n),
            "Mean ± SEM": mean_with_unc,          
            "Mean": mean,
            "SEM": sem,
            "SD": sd,
            "Min": data.min(),
            "Median": data.median(),
            "Max": data.max(),
            "%CV": cv
        })
    
    summary = pd.DataFrame(rows)
    
    # Reorder columns nicely
    col_order = ["Variable", "N", "Mean ± SEM", "Mean", "SEM", "SD", "Min", "Median", "Max", "%CV"]
    summary = summary[col_order]
    
    # Optional: round numeric columns that aren't already perfect
    numeric_cols = ["Mean", "SEM", "SD", "%CV", "Min", "Median", "Max"]
    summary[numeric_cols] = summary[numeric_cols].round(6)  # safe, won't hurt ufloat display
    
    return summary

def reframe_data(x, y, z=None, z2=None):
    '''
    This function takes in x, y, z, and z2 as inputs and returns a cleaned DataFrame  
    with numeric values for x and y, and drops any rows with NaN values.
    
    z and z2 are optional and can be included if provided.
    '''
    # Create a dictionary to hold the data    
    data = {x.name: pd.to_numeric(x, errors="coerce"),
            y.name: pd.to_numeric(y, errors="coerce")}
    if z is not None:
        data[z.name] = z
    if z2 is not None:
        data[z2.name] = z2
        
    return pd.DataFrame(data).dropna()

def Deming_Plot_Equal_Variance_with_Error2_PX(x, y, z=None, z2=None, **error_info):
    """
    Creates a Deming regression plot with error bands using Plotly Express.
    
    Parameters:
    - x, y: Pandas Series or array-like, numeric data for x and y axes.
    - z: Continuous variable for color mapping (optional).
    - z2: Categorical variable for symbol mapping (optional).
    - error_info: Dictionary with keys 'error1' (absolute error), 'error2' (percentage error),
                  and 'Error_level_cut_off' (concentration cut-off).

    Returns:
    - Tuple: (number of points, slope, y-intercept, Pearson R).

    """

    """
    Below is the most common type of Deming Regression with Equal Variance
    Here the Lambda ratio is simply set to 1
    """

    # Extract error parameters
    error1 = error_info.get("error1")
    error2 = error_info.get("error2")
    cutoff = error_info.get("Error_level_cut_off")
    output_path = error_info.get("output_path", "Figures/Deming_regression.png")

    # Reframe and clean data
    df = reframe_data(x, y, z, z2)

    # Extract cleaned data
    x, y = df[x.name], df[y.name]
    z = df[z.name] if z is not None else None
    z2 = df[z2.name] if z2 is not None else None
    xlen = len(x)

    if len(x) == len(y):
        ## below is the math for Deming regression with equal variance
        ##
        x_bar = np.mean(x)
        y_bar = np.mean(y)
        p = sum([(xi - x_bar) * (yi - y_bar) for xi, yi in zip(x, y)])
        u = sum((xi - x_bar) ** 2 for xi in x)
        q = sum((yi - y_bar) ** 2 for yi in y)
        # here we take the lambda ratio as (1) by assuming equal variance
        lambda_ratio = 1
        # slope
        slope = (
            (lambda_ratio * q - u)
            + (((u - lambda_ratio * q) ** 2) + 4 * lambda_ratio * (p**2)) ** 0.5
        ) / (2 * lambda_ratio * p)
        # intercept
        yintercept = y_bar - slope * x_bar
        
        # Pearson R
        pearson_r = np.corrcoef(x, y)[0, 1]

    else:
        print("X and Y Variables are note the same length")

    # Create error band x-values
    x_range = np.linspace(x.min(), x.max(), 100)
    
    # Calculate error bands
    if error1 is None and error2 is None:
        yabove = ybelow = x_range
    elif error1 is not None and error2 is None:
        yabove = x_range + error1
        ybelow = x_range - error1
    elif error2 is not None and error1 is None:
        yabove = x_range * (1 + error2 / 100)
        ybelow = x_range * (1 - error2 / 100)
    else:
        yabove = np.where(x_range <= cutoff, x_range + error1, x_range * (1 + error2 / 100))
        ybelow = np.where(x_range <= cutoff, x_range - error1, x_range * (1 - error2 / 100))

    scatter_args = {"x": x, "y": y}  # base/mandatory arguments
    if z is not None and z2 is not None:
        scatter_args.update({"color": z, "symbol": z2, "labels": {"color": z.name, "symbol": z2.name}})
    elif z is not None:
        scatter_args.update({"color": z, "labels": {"color": z.name}})
    elif z2 is not None:
        scatter_args.update({"symbol": z2, "labels": {"symbol": z2.name}})
    scatter_fig = px.scatter(**scatter_args, color_continuous_scale="Bluered_r")

    # Initialize figure
    fig = go.Figure()

    # Add error band (first layer)
    fig.add_trace(
        go.Scatter(
            x=np.concatenate([x_range, x_range[::-1]]),
            y=np.concatenate([yabove, ybelow[::-1]]),
            fill="toself",
            fillcolor="rgba(211,211,211,0.60)",  # dark gray with 15% opacity
            line=dict(color="rgba(255,255,255,0)"),
            name="Allowable Error",
            showlegend=True,
        )
    )

    # Add scatter traces, ensuring legend and color scale
    for trace in scatter_fig.data:

        if not trace.name: #If the trace doesn't have a name (e.g., when z and z2 are None)
            trace.name = "Samples" # give it a default label
            
        trace.update(showlegend=True)  # Ensure legend is shown for z2 and z, i.e. use labels above.
        fig.add_trace(trace)

    # Apply color scale to final figure if z is provided
    if z is not None:   
        fig.update_layout(coloraxis=dict(colorscale="Bluered_r"))
    
    # Add line of identity
    expansion = 0.025 * x.min()
    x_ext = [x.min() - expansion, x.max() + expansion]
    fig.add_trace(
        go.Scatter(
            x=x_ext,
            y=x_ext,
            mode="lines",
            line=dict(color="black", dash="dot", width=0.8),
            name="Line of Identity",
        )
    )

    # Add Deming regression line
    fig.add_trace(
        go.Scatter(
            x=x_ext,
            y=[slope * x_ext[0] + yintercept, slope * x_ext[1] + yintercept],
            mode="lines",
            line=dict(color="Purple", dash="solid", width=1.0),
            name="Line of Fit",
        )
    )

    # Add regression annotation
    fig.add_annotation(
        xref="paper",
        yref="paper",
        x=0.1,
        y=0.9,
        xanchor="left",
        yanchor="top",
        text=f"Regression by Deming<br>y = {slope:.2f}x{'+' if yintercept > 0 else ''}{yintercept:.2f}<br>Pearson R: {pearson_r:.3f}<br>n = {xlen}",
        showarrow=False,
        font=dict(size=12, color="black"),
    )

    # Update layout
    fig.update_layout(
        #plot_bgcolor="white",
        #paper_bgcolor="white",
        xaxis=dict(showgrid=True, gridcolor="lightgrey", title=x.name),
        yaxis=dict(showgrid=True, gridcolor="lightgrey", title=y.name),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        template="plotly_white",
        width=600,
        height=600,
        )
    return FigWhiteCrop_stream(fig, margin=10), xlen, slope, yintercept, pearson_r
    
#####################################################
# Difference plot with Median And error regions ####
########################################################

def Difference_plot_median_with_error2(x, y, z=None, z2=None, **error_info):
    """
    error1 = error in absolute terms as indicated in IQMH
    error2 = error in percentage terms as indicated in IQMH; not considered in this function and plot
    concentration_cut = the cut-off at which error1 and error2 are applicable

    if concentration_cut is not provided, but only one of the error terms is provided, 
    then error (either error1 or erro2) is calculated across all range of x. 
    """
    # Extract parameters
    error1 = error_info.get("error1", None)
    Error_level_cut_off = error_info.get("Error_level_cut_off", None)
    output_path = error_info.get("output_path", "Figures/Difference_Plot_median.png")

    # Input validation
    if not hasattr(x, 'name') or not hasattr(y, 'name'):
        raise ValueError("Inputs x and y must have a 'name' attribute (e.g., pandas Series).")
    if z is not None and not hasattr(z, 'name'):
        raise ValueError("Input z must have a 'name' attribute if provided.")
    if z2 is not None and not hasattr(z2, 'name'):
        raise ValueError("Input z2 must have a 'name' attribute if provided.")

    # Reframe and clean data
    df = reframe_data(x, y, z, z2)
    x, y = df[x.name], df[y.name]
    z = df[z.name] if z is not None else None
    z2 = df[z2.name] if z2 is not None else None
    ydiff = y - x

    # Build scatter plot arguments dynamically
    scatter_args = {
        "x": x,
        "y": ydiff,
        "labels": {
            "x": x.name,
            "y": f"{y.name} - {x.name}"
        }
    }
    if z is not None:
        scatter_args["color"] = z
        scatter_args["color_continuous_scale"] = "Bluered_r"
        scatter_args["labels"]["color"] = z.name
    if z2 is not None:
        scatter_args["symbol"] = z2
        scatter_args["labels"]["symbol"] = z2.name

    fig = px.scatter(**scatter_args)

    # Calculate expansion factor
    expansion_factor = 0.025 * (x.max() - x.min())

    # x-values to span with some expansion
    x_range=[x.min() - expansion_factor, x.max() + expansion_factor]
    
    # median bias line
    median_y = ydiff.median()
    fig.add_trace(
        go.Scatter(
            x = x_range,
            y=[median_y, median_y],
            mode="lines",
            line=dict(color="red", dash="solid", width=0.8),
            name=f"Median Bias: {round(median_y, 1)}",
        )
    )

    #plot IQR lines
    q1 = ydiff.quantile(0.25)
    q3 = ydiff.quantile(0.75)
    for q_val, label in zip([q1, q3], ["25th Pct", "75th Pct"]):
        fig.add_trace(go.Scatter(
            x=x_range, y=[q_val, q_val],
            mode="lines", line=dict(color="red", dash="dot", width=0.6),
            name=f"{label}: {round(q_val, 2)}"
        ))

    # Add agreement line or the zero line
    fig.add_trace(
        go.Scatter(
            x=x_range,
            y=[0, 0],
            mode="lines",
            line=dict(color="black", dash="dot", width=0.8),
            name="Agreement Line",
        )
    )

    # Add error band if error1 is provided
    if error1:
        x1 = Error_level_cut_off if Error_level_cut_off not in (0, None) else x.max() + expansion_factor
        fig.add_shape(
            type="rect",
            x0=x.min() - expansion_factor,
            x1=x1,
            y0=0 + error1,
            y1=0 - error1,
            fillcolor="lightgrey", 
            opacity=0.6,
            layer="below",
            line_width=0,
            name ="Allowable Error",
        )

    # Update layout
    fig.update_layout(
        plot_bgcolor="white",
        paper_bgcolor="white",
        xaxis=dict(showgrid=True, gridcolor="lightgrey", title=x.name),
        yaxis=dict(showgrid=True, gridcolor="lightgrey", title=f"{y.name} - {x.name}"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        height=600, width=600, 
    )

    return median_y, FigWhiteCrop_stream(fig, margin=10)

#####################################################
# % Difference plot with Median And Error Regions ####
########################################################

def Difference_plot_Percent_median_with_error2(x, y, z=None, z2=None, **error_info):
    """
    error1 = error in absolute terms as indicated in IQMH - Not considered in this function and plot
    error2 = error in percentage terms as indicated in IQMH
    concentration_cut = the cut-off at which error1 and error2 are applicable

    If concentration_cut is not provided, but only one of the error terms is provided, 
    then error (either error1 or erro2) is calculated across all range of x.
    """

    # Extract parameters
    error2 = error_info.get("error2", None)
    Error_level_cut_off = error_info.get("Error_level_cut_off", None)
    output_path = error_info.get("output_path", "Figures/Difference_Plot_median_per.png")

    # Input validation
    if not hasattr(x, 'name') or not hasattr(y, 'name'):
        raise ValueError("Inputs x and y must have a 'name' attribute (e.g., pandas Series).")
    if z is not None and not hasattr(z, 'name'):
        raise ValueError("Input z must have a 'name' attribute if provided.")
    if z2 is not None and not hasattr(z2, 'name'):
        raise ValueError("Input z2 must have a 'name' attribute if provided.")

    # Convert to numeric and create DataFrame
    df = reframe_data(x, y, z, z2)
    # Extract cleaned data
    x, y = df[x.name], df[y.name]
    z = df[z.name] if z is not None else None
    z2 = df[z2.name] if z2 is not None else None

    # Calculate percent difference
    # Replace 0 with NaN in x to prevent zero-division without aborting the script
    x = x.replace(0, float('nan'))
    ydiff = 100 * (y - x) / x

    # Build scatter plot arguments dynamically
    scatter_args = {
        "x": x,
        "y": ydiff,
        "labels": {
            "x": x.name,
            "y": "% Difference"
        }
    }
    if z is not None:
        scatter_args["color"] = z
        scatter_args["color_continuous_scale"] = "Bluered_r"
        scatter_args["labels"]["color"] = z.name
    if z2 is not None:
        scatter_args["symbol"] = z2
        scatter_args["labels"]["symbol"] = z2.name
    fig = px.scatter(**scatter_args)

    # Calculate expansion factor
    expansion_factor = 0.025 * (x.max() - x.min())
    x_range=[x.min() - expansion_factor, x.max() + expansion_factor]
    
    # Add median bias line
    median_y = ydiff.median()
    fig.add_trace(
        go.Scatter(
            x=x_range,
            y=[median_y, median_y],
            mode="lines",
            line=dict(color="red", dash="solid", width=0.8),
            name=f"Median Bias: {round(median_y, 1)}%",
        )
    )

    #plot IQR lines
    q1 = ydiff.quantile(0.25)
    q3 = ydiff.quantile(0.75)
    for q_val, label in zip([q1, q3], ["25th Pct", "75th Pct"]):
        fig.add_trace(go.Scatter(
            x=x_range, y=[q_val, q_val],
            mode="lines", line=dict(color="red", dash="dot", width=0.6),
            name=f"{label}: {round(q_val, 2)}%"
        ))

    # Add agreement line
    fig.add_trace(
        go.Scatter(
            x=[x.min() - expansion_factor, x.max() + expansion_factor],
            y=[0, 0],
            mode="lines",
            line=dict(color="black", dash="dot", width=0.8),
            name="Agreement Line",
        )
    )

    # Add error band if error2 is provided
    if error2:
        if Error_level_cut_off not in (0, None):
            x0 = Error_level_cut_off
            x1 = x.max() + expansion_factor
        else:
            x0 = x.min() - expansion_factor
            x1 = x.max() + expansion_factor
        fig.add_shape(
            type="rect",
            x0=x0,
            x1=x1,
            y0=0 + error2,
            y1=0 - error2,
            fillcolor="lightgrey",
            opacity=0.6,
            layer="below",
            line_width=0,
            name = "Allowable Error"
        )

    # Update layout
    fig.update_layout(
        plot_bgcolor="white",
        paper_bgcolor="white",
        xaxis=dict(showgrid=True, gridcolor="lightgrey", title=x.name),
        yaxis=dict(showgrid=True, gridcolor="lightgrey", title="% Difference"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
    )

    return median_y, FigWhiteCrop_stream(fig, margin=10)
##
# Function to write results to CSV
def write_to_csv(analyte, n, slope, y_intercept, pearson_r, median_bias, percent_bias, unit, filename="MethodSummary.csv"):
    '''
    Write the results to a CSV file.
    
    Args:
        analyte: The analyte name or identifier.
        n: Sample size.
            Obtainble from function Deming_Plot_Equal_Variance_with_Error2_PX
        slope: Slope of the regression line.
            Obtainble from function Deming_Plot_Equal_Variance_with_Error2_PX
        y_intercept: Y-intercept of the regression line.
            Obtainble from function Deming_Plot_Equal_Variance_with_Error2_PX
        pearson_r: Pearson correlation coefficient from function.
            Obtainble from function Deming_Plot_Equal_Variance_with_Error2_PX
        median_bias: Median bias value.
            Obtainble from function Difference_plot_median_with_error2
        percent_bias: Percent bias value.
            Obtainble from function Difference_plot_Percent_median_with_error2
        unit: Unit of measurement.
        filename: Name of the CSV file (default: 'results.csv').
    '''
    
    # Define the header for the CSV file
    header = ["Analyte", "N", 'Slope', 'y-Intercept', "Pearson's r", 'Median Bias', '% Bias', 'Unit']
    
    # Check if file exists to determine if we need to write the header
    file_exists = os.path.isfile(filename)
    
    # Open the file in append mode
    with open(filename, mode='a', newline='') as file:
        writer = csv.writer(file)
        
        # Write header if file doesn't exist
        if not file_exists:
            writer.writerow(header)
            
        # Write the data row
        writer.writerow([analyte, n, slope, y_intercept, pearson_r, median_bias, percent_bias, unit])

# function to create a word table from our data frame
# You have this function below else where.  Consider make this function as part of a class
def create_word_table(document, MyTable):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    
    table = document.add_table(MyTable.shape[0] + 1, MyTable.shape[1])
    table.style = document.styles["Light Shading Accent 1"] # previously set to "Light Grid Accent 1"
    table.autofit = True
    table.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # Set column widths (optional) 
    for col in table.columns:
        for cell in col.cells:
            cell.width = Cm(2.5)  # Set each column width to 3 cm
    

    # Add the header rows
    for j in range(MyTable.shape[-1]):
        cell = table.cell(0, j)
        cell.text = MyTable.columns[j]
        # Set font size for header cell
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER # center text in each cell
            for run in paragraph.runs:
                run.font.size = Pt(10)
    
    # Add the rest of the data frame
    for i in range(MyTable.shape[0]):
        for j in range(MyTable.shape[-1]):
            cell = table.cell(i + 1, j)
            cell.text = str(MyTable.values[i, j])
            # Set font size for data cell
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER # center text in each cell
                for run in paragraph.runs:
                    run.font.size = Pt(10)
    document.save("MethodComparison.docx")

def MC_output(document,analyte, x, y, z=None, z2=None, Unit=None, **error_info):    
    """
    This function is for creatining the entire Method Comparison section;
    It outputs a Table summary, comparing the difference and %difference given the error conditions
    It outputs a Summary and Table of parired Student T-test
    It outputs Histrogram and Q-Q plots for distribution analysis of each axis
    it outputs a Regression Analysis by Deming, a Difference plot, and %Difference plot
    """

    error1 = error_info.get("error1", None)
    error2 = error_info.get("error2", None)
    Error_level_cut_off = error_info.get("Error_level_cut_off", None)

    # adjust page margins
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(0.5)
        section.bottom_margin = Cm(0.5)
        section.left_margin = Cm(0.5)
        section.right_margin = Cm(0.5)
        
    # add heading
    #p = document.add_heading(f"Method Comparison Studies", level=1)
    p = document.add_heading(f" {analyte} (new) vs. {analyte} (old)", level=2)
    p = document.add_paragraph("")

    p.add_run(
        "Pass/Fail is determined based on the allowable error, which is defined as the following."
    )
    p = document.add_paragraph("")
    if Error_level_cut_off == 0 or Error_level_cut_off == None:
        p.add_run("Across all range of the measuring interval: ")
        if error2 is not None:
            p.add_run(str(error2))
            p.add_run("%")
        elif error1 is not None:
            p.add_run(" ± ")
            p.add_run(str(error1))
            p.add_run(" ")
            p.add_run(Unit)
    else:
        p.add_run("< ")
        p.add_run(str(Error_level_cut_off))
        p.add_run(": ± ")
        p.add_run(str(error1))
        p.add_run(" ")
        p.add_run(Unit)

        p = document.add_paragraph("")
        p.add_run("≥ ")
        p.add_run(str(Error_level_cut_off))
        p.add_run(": ±")
        p.add_run(str(error2))
        p.add_run("%")

    p = document.add_paragraph("")
    p.add_run("The Reference Method (X) is: ")
    p.add_run(f"{analyte} (old) - Current method").bold = True

    p = document.add_paragraph("")
    p.add_run("The Test Method (Y) is: ")
    p.add_run(f"{analyte} (new) - Test method").bold = True

    ### Distribution Analysis
    #### Histogram
    #document.add_page_break()
    document.add_heading("Distribution of Measurements:", level=3)

    print(f"Creating Histogram for {analyte} (old) and {analyte} (new)", flush=True)
    hist_stream = Histogram_grouped(x, y)
    
    document.add_picture(hist_stream, width=Inches(8))

    Table1, Table1b = sample_comparions_tableX(x, y, z, **error_info)

    # Table summary of Distribution Analysis
    Table1b_summary = data_summary(Table1b)
    create_word_table(document, Table1b_summary)

    ### Regression and Bias Analysis
    #document.add_page_break()
    document.add_heading("Regression & Bias Analysis", level=3)
    # Deming Regression assuming equal variance plot
    deming_stream, n, slope, yintercept, Pearson_r = Deming_Plot_Equal_Variance_with_Error2_PX(x, y, z, z2, **error_info)

    p = document.add_paragraph("")
    p.add_run("Regression is based on Deming method, assuming equal variance of both methods. ")
    document.add_picture(deming_stream, width=Inches(6.5))


    
    # Summary Table of the Methods
    p = document.add_paragraph("")
    p.add_run("Comparison of Paired Measurements").bold = True
    create_word_table(document, Table1)
    document.add_page_break()

    # Difference Plot
    median_bias, differencePlotStream = Difference_plot_median_with_error2(x, y, z, z2, **error_info)
    # %Difference Plot
    percent_bias, percentDiffPlotStream = Difference_plot_Percent_median_with_error2(x, y, z, z2, **error_info)


    p = document.add_paragraph("")
    p.add_run("Difference Plots").bold = True

    table = document.add_table(rows=2, cols=1)

    cell_1 = table.cell(0, 0)
    p = cell_1.paragraphs[0]
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    run = p.add_run()
    run.add_picture(differencePlotStream, width=Inches(6))
    # Add a line break
    p = cell_1.add_paragraph("")

    cell_2 = table.cell(1, 0)
    p = cell_2.paragraphs[0]
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = p.add_run()
    run.add_picture(
        percentDiffPlotStream, width=Inches(6)
    )
    
    write_to_csv(analyte, n, round(slope,4), round(yintercept,4), round(Pearson_r,4), round(median_bias,4), round(percent_bias,4), Unit)
    document.save("MethodComparison.docx")