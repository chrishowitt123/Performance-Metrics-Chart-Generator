"""
Performance Metrics Chart Generator

This script generates and saves trend charts for performance metrics, focusing on metrics 
that are either marked as 'Red' or showing significant trends. It reads data from an Excel file,
processes the metrics, and creates interactive charts using Plotly.

Dependencies:
- plotly.graph_objects: For creating interactive charts
- plotly.io: For saving and rendering charts
- pandas: For data manipulation
- os: For file and directory operations
"""

import plotly.graph_objects as go
import plotly.io as pio
import pandas as pd
import os

# Configure save directory for output charts
save_dir = r"C:\Users\chowitt\OneDrive - States of Guernsey\Desktop\charts"
os.makedirs(save_dir, exist_ok=True)  # Create directory if it doesn't exist

# Load data from Excel file and convert dates
df = pd.read_excel(r"C:\Users\chowitt\Downloads\data.xlsx")
df['Date'] = pd.to_datetime(df['Date'])  # Convert string dates to datetime objects

# Get the most recent date for each metric to identify latest status
latest_dates = df.groupby('Metric Reference')['Date'].max()

# Initialize list to store metrics requiring attention
reporting_metrics = []

# Identify metrics that are either red or showing significant trends (↗ or ↘)
for metric_ref in latest_dates.index:
    metric_data = df[df['Metric Reference'] == metric_ref]
    latest_row = metric_data[metric_data['Date'] == latest_dates[metric_ref]]
    if (latest_row['3Q Trend'].iloc[0] == '↗' or 
        latest_row['3Q Trend'].iloc[0] == '↘' or 
        latest_row['RAG Text'].iloc[0] == 'Red'):
        reporting_metrics.append(metric_ref)

def format_value(value, unit_type):
    """
    Format numerical values based on their unit type for display.
    
    Args:
        value (float/int): The numerical value to format
        unit_type (str): The type of unit (e.g., 'Percentage', 'Currency', etc.)
    
    Returns:
        str: Formatted value as a string with appropriate formatting and symbols
    
    Examples:
        >>> format_value(0.756, "Percentage")
        "76%"
        >>> format_value(1500000, "Currency millions")
        "£1.5m"
    """
    try:
        num_value = float(value)
        
        if pd.isna(num_value):
            return ""
            
        # Handle different unit types with appropriate formatting
        if unit_type == "Percentage":
            # Convert decimal to percentage (e.g., 0.756 -> 76%)
            return f"{int(round(num_value * 100)):,}%"
            
        elif unit_type == "Whole number":
            # Format with thousands separator (e.g., 1000 -> 1,000)
            return f"{int(num_value):,}"
            
        elif unit_type == "Decimal":
            # Format with thousands separator, showing decimals only if needed
            if num_value.is_integer():
                return f"{int(num_value):,}"
            return f"{num_value:,.2f}"
            
        elif unit_type in ["Thousands", "Currency thousands"]:
            # Format as whole number with thousands separator
            return f"{int(num_value):,}"
            
        elif unit_type == "Currency millions":
            # Format as millions with £ symbol (e.g., 1.5 -> £1.5m)
            if num_value % 1 == 0:
                return f"£{int(num_value):,}m"
            return f"£{num_value:,.1f}m"
            
        elif unit_type == "Currency":
            # Format with £ symbol and thousands separator
            if num_value.is_integer():
                return f"£{int(num_value):,}"
            return f"£{num_value:,.2f}"
            
        elif unit_type == "Currency small":
            # Format with £ symbol and 3 decimal places
            return f"£{num_value:,.3f}"
            
        # Default format: round to 2 decimal places with thousands separator
        return f"{round(num_value, 2):,}"
        
    except (ValueError, TypeError):
        return str(value)



def create_chart(metric_data, metric_name):
    """
    Create a Plotly figure for a given metric with appropriate styling and thresholds.
    
    Args:
        metric_data (pd.DataFrame): DataFrame containing the metric data
        metric_name (str): Name of the metric for the chart title
    
    Returns:
        go.Figure: Plotly figure object with complete chart configuration
    """
    # Data preparation and sorting
    metric_data = metric_data.sort_values('Date')
    num_quarters = len(metric_data)
    unit_type = metric_data['Units of Measure'].iloc[0]
    
    # Convert Value column to numeric format
    metric_data['Value'] = pd.to_numeric(metric_data['Value'], errors='coerce')
    
    # Special handling for Currency millions - convert values to millions
    if unit_type == "Currency millions":
        metric_data['Value'] = metric_data['Value'] / 1000000
        # Convert threshold values if they exist
        threshold_columns = ['Red Above', 'Red Below', 'Amber Above', 'Amber Below', 'Target']
        for col in threshold_columns:
            if col in metric_data.columns:
                metric_data[col] = pd.to_numeric(metric_data[col], errors='coerce') / 1000000
    
    # Initialize Plotly figure
    fig = go.Figure()
    x_positions = list(range(num_quarters))
    
    # Add threshold lines with appropriate colors and styles
    threshold_definitions = [
        ('Red Above', 'red'), 
        ('Red Below', 'red'),
        ('Amber Above', 'orange'), 
        ('Amber Below', 'orange'),
        ('Target', 'green')
    ]
    
    # Add each threshold line if it exists in the data
    for threshold, color in threshold_definitions:
        if threshold in metric_data.columns and pd.notna(metric_data[threshold].iloc[0]):
            fig.add_trace(go.Scatter(
                x=[0, num_quarters-1],
                y=[metric_data[threshold].iloc[0], metric_data[threshold].iloc[0]],
                mode='lines',
                line=dict(color=color, dash='dot', width=1),
                showlegend=False
            ))

    # Add main trend line with smooth interpolation
    fig.add_trace(go.Scatter(
        x=x_positions,
        y=metric_data['Value'],
        mode='lines',
        line=dict(
            color='#2E74B5',  # Corporate blue color
            width=3,
            shape='spline',  # Smooth line
            smoothing=1.3    # Smoothing factor
        ),
        showlegend=False
    ))

    # Add marker for the most recent data point
    fig.add_trace(go.Scatter(
        x=[x_positions[-1]],
        y=[metric_data['Value'].iloc[-1]],
        mode='markers',
        marker=dict(color='#2E74B5', size=11),
        showlegend=False
    ))

    # Add label for the most recent value
    fig.add_annotation(
        x=x_positions[-1],
        y=metric_data['Value'].iloc[-1],
        text=format_value(metric_data['Value'].iloc[-1], unit_type),
        showarrow=False,
        yshift=20,
        font=dict(size=17, color='#2E74B5', family='Calibri')
    )

    # Calculate y-axis range including thresholds and add padding
    y_values = [metric_data['Value'].min(), metric_data['Value'].max()]
    threshold_columns = ['Red Above', 'Red Below', 'Amber Above', 'Amber Below', 'Target']
    
    # Include threshold values in y-axis range calculation
    for col in threshold_columns:
        if col in metric_data.columns:
            val = pd.to_numeric(metric_data[col].iloc[0], errors='coerce')
            if pd.notna(val):
                y_values.append(val)
    
    # Filter out NA values and calculate range with padding
    valid_y_values = [y for y in y_values if pd.notna(y)]
    y_range = max(valid_y_values) - min(valid_y_values)
    y_min = min(valid_y_values) - (y_range * 0.2)  # 20% padding below
    y_max = max(valid_y_values) + (y_range * 0.3)  # 30% padding above

    # Configure y-axis tick format based on unit type
    ytick_format = ""
    if unit_type == "Percentage":
        ytick_format = ".0%"
    elif unit_type in ["Thousands", "Currency thousands", "Whole number"]:
        ytick_format = ","
    elif unit_type == "Currency millions":
        ytick_format = ",.1f"  # Number format for millions
    elif unit_type == "Currency":
        ytick_format = "£,.2f"
    elif unit_type == "Currency small":
        ytick_format = "£,.3f"
    elif unit_type == "Decimal":
        ytick_format = ",.2f"
    else:
        ytick_format = "," # Defult showing commas

    # Update chart layout with comprehensive styling
    fig.update_layout(
        title=dict(
            text=metric_name,
            font=dict(family='Calibri', size=18, color='#2E74B5'),
            y=0.95,
            yanchor='top',
            x=0.05,
            xanchor='left'
        ),
        plot_bgcolor='white',
        showlegend=False,
        width=800,
        height=400,
        margin=dict(l=80, r=60, t=60, b=90),
        xaxis=dict(
            ticktext=metric_data['Quarter'],
            tickvals=x_positions,
            tickangle=0,
            showgrid=False,
            domain=[0, 1],
            range=[-0.5, num_quarters-0.5],
            tickfont=dict(family='Calibri', color='#808080', size=17)
        ),
        yaxis=dict(
            showgrid=False,
            range=[y_min, y_max],
            tickfont=dict(family='Calibri', color='#808080', size=17),
            tickformat=ytick_format,
            ticksuffix='m' if unit_type == "Currency millions" else '',
            tickprefix='£' if unit_type == "Currency millions" else ''
        ),
        font=dict(family='Calibri', size=17)
    )

    # Add year labels below quarters
    years = metric_data['Year'].unique()
    quarters = metric_data['Quarter'].tolist()

    # Calculate year label positions
    year_positions = {}
    for i, (year, quarter) in enumerate(zip(metric_data['Year'], quarters)):
        if year not in year_positions:
            year_positions[year] = []
        year_positions[year].append(i)

    # Add year annotations
    for year in year_positions:
        positions = year_positions[year]
        middle_position = sum(positions) / len(positions)
        
        fig.add_annotation(
            text=str(int(year)),
            x=middle_position,
            y=-0.3,
            xref="x",
            yref="paper",
            showarrow=False,
            font=dict(size=17, family='Calibri', color='#808080')
        )

    return fig

# Configure Plotly to show charts in browser and save as PNG
pio.renderers.default = "png"

# Main processing loop
print(f"Total metrics to process: {len(reporting_metrics)}")

for index, metric_ref in enumerate(reporting_metrics, 1):
    try:
        print(f"\nProcessing metric {index} of {len(reporting_metrics)}")
        # Get data for current metric
        metric_data = df[df['Metric Reference'] == metric_ref].copy()
        metric_name = metric_data['Metric Name'].iloc[0]
        print(f"Creating chart for: {metric_name}")
        
        # Generate chart
        fig = create_chart(metric_data, metric_name)
        
        # Create valid filename by removing invalid characters
        valid_filename = "".join(x for x in metric_name if x.isalnum() or x in (' ', '-', '_'))
        png_path = os.path.join(save_dir, f"{valid_filename}.png")
        
        # Display chart in browser
        fig.show()
        
        # Save chart as PNG
        print(f"Saving PNG file...")
        try:
            pio.write_image(fig, png_path, engine='kaleido', width=800, height=400)
            print(f"PNG file saved successfully")
        except Exception as e:
            print(f"Error saving PNG: {str(e)}")
        
        # Clean up memory
        fig.data = []
        fig.layout = {}
        
        print(f"Successfully processed metric {index} of {len(reporting_metrics)}")
        
    except Exception as e:
        print(f"Error processing metric {index}: {str(e)}")
        continue

print("\nProcessing complete!")

# Print the last filter context row for verification
print(df.iloc[-1, 0])
