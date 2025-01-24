import pandas as pd
import plotly.graph_objects as go
import tkinter as tk
from tkinter import filedialog
import os
import base64

# Function to extract icon name based on element type (first part of the name)
def get_icon_name(element):
    # NEW: strip whitespace from the raw string
    element = element.strip()
    return "-".join(element.split("-")[:3])

# Function to clean the element name by removing _UP, _CT, _DN
def clean_element_name(element):
    return element.replace("_UP", "").replace("_CT", "").replace("_DN", "")

# Function to convert image to base64, handling SVG vs PNG
def encode_image_to_base64(image_path):
    """
    Detect file extension and return the appropriate base64 data URI.
    """
    file_extension = os.path.splitext(image_path)[1].lower()
    # Determine the correct MIME type
    if file_extension == ".svg":
        mime_type = "image/svg+xml"
    elif file_extension == ".png":
        mime_type = "image/png"
    else:
        # If other images are possible, adjust here. For now, default to PNG if unknown.
        mime_type = "image/png"

    with open(image_path, 'rb') as f:
        encoded = base64.b64encode(f.read()).decode('ascii')
    return f"data:{mime_type};base64,{encoded}"

# Function to generate symmetric y-offsets
def generate_symmetric_offsets(n, step):
    """
    Generate a list of y-offsets symmetric about y=0.

    Parameters:
    - n (int): Number of files.
    - step (float): The step size between offsets.

    Returns:
    - List[float]: A list of y-offsets.
    """
    offsets = []
    if n % 2 == 1:
        # If odd, include y=0
        offsets.append(0)
        for i in range(1, (n//2)+1):
            offsets.append(i * step)
            offsets.append(-i * step)
    else:
        # If even, start with positive offset
        for i in range(1, (n//2)+1):
            offsets.append(i * step)
            offsets.append(-i * step)
    return offsets[:n]

def main():
    # Step 1: Use tkinter to allow multiple file selection
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_paths = filedialog.askopenfilenames(
        title="Select Excel file(s)",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if not file_paths:
        print("No file selected, exiting.")
        exit()

    # Path to the folder containing the element icons
    icon_folder = filedialog.askdirectory(title="Select the folder containing element icons")

    # Step 2: Initialize plot and variables
    fig = go.Figure()
    offset_step = 5  # Adjust as needed
    n_files = len(file_paths)
    y_offsets = generate_symmetric_offsets(n_files, offset_step)
    missing_dimensions_elements = []  # Collect elements with missing dimensions
    required_icons = set()  # Store the icon names as required icon names
    element_parameters = {}  # Dictionary to store element parameters
    cryomodules = {}  # Dictionary to store cryomodules and their boundaries

    # Step 3: Iterate through each selected file and collect cryomodule boundaries
    for file_path in file_paths:
        df = pd.read_excel(file_path)

        # NEW: strip whitespace from the element column
        df.iloc[:, 1] = df.iloc[:, 1].astype(str).fillna('').str.strip()

        # Remove all _CT entries related to cryomodules
        df = df[~df.iloc[:, 1].str.contains('CM') | ~df.iloc[:, 1].str.contains('_CT')]

        # Identify cryomodule boundaries
        for idx in range(len(df)):
            element = df.iloc[idx, 1]
            location = df.iloc[idx, 4]

            # Check if the element is a cryomodule _UP or _DN entry
            if 'CM' in element and ('_UP' in element or '_DN' in element):
                cm_name = element.replace('_UP', '').replace('_DN', '')
                if cm_name not in cryomodules:
                    cryomodules[cm_name] = {'start': None, 'end': None}
                if '_UP' in element:
                    cryomodules[cm_name]['start'] = location
                elif '_DN' in element:
                    cryomodules[cm_name]['end'] = location

    # Step 4: Plot elements except cryomodules
    for file_index, file_path in enumerate(file_paths):
        df = pd.read_excel(file_path)

        # NEW: strip whitespace from the element column
        df.iloc[:, 1] = df.iloc[:, 1].astype(str).fillna('').str.strip()

        # Remove all _CT entries related to cryomodules
        df = df[~df.iloc[:, 1].str.contains('CM') | ~df.iloc[:, 1].str.contains('_CT')]

        # Assign symmetric y_offset
        y_offset = y_offsets[file_index]

        for idx in range(len(df)):
            element = df.iloc[idx, 1]
            location = df.iloc[idx, 4]

            # Skip all cryomodule entries
            if 'CM' in element:
                continue

            # Extract the base name for the icon
            icon_name = get_icon_name(element)
            required_icons.add(f"{icon_name}.svg")
            clean_name = clean_element_name(element)

            # Check if the element is a 3-line element (*_UP, *_CT, *_DN)
            if '_UP' in element:
                # Initialize variables
                up_location = location
                dn_location = None
                central_ct_location = None

                # Find the corresponding '_DN' entry
                for j in range(idx + 1, len(df)):
                    next_elem = df.iloc[j, 1]
                    if '_DN' in next_elem and clean_element_name(next_elem) == clean_name:
                        dn_location = df.iloc[j, 4]
                        break

                if dn_location is not None:
                    # Find the central '_CT' entry (without '_P1_' or '_P2_')
                    for j in range(idx + 1, len(df)):
                        next_elem = df.iloc[j, 1]
                        if '_CT' in next_elem and clean_element_name(next_elem) == clean_name:
                            if '_P1_' not in next_elem and '_P2_' not in next_elem:
                                central_ct_location = df.iloc[j, 4]
                                break

                    # If central '_CT' not found, optionally handle it
                    if central_ct_location is None:
                        print(f"Warning: No central '_CT' found for element {element} in file {file_path}. Using first '_CT' if found.")
                        for j in range(idx + 1, len(df)):
                            next_elem = df.iloc[j, 1]
                            if '_CT' in next_elem and clean_element_name(next_elem) == clean_name:
                                central_ct_location = df.iloc[j, 4]
                                break
                        if central_ct_location is None:
                            print(f"Error: No '_CT' entry found for element {element} in file {file_path}. Skipping.")
                            continue

                    # Calculate length
                    length = dn_location - up_location
                    element_height = length

                    if length == 0:
                        # Represent zero-length elements as vertical dashed lines
                        fig.add_shape(
                            type="line",
                            x0=location, x1=location,
                            y0=y_offset - 0.5, y1=y_offset + 0.5,
                            line=dict(color="Red", dash="dash"),
                        )
                        # Add invisible hover-enabled marker for zero-length element
                        fig.add_trace(go.Scatter(
                            x=[location],
                            y=[y_offset],
                            mode="markers",
                            marker=dict(size=1, color="rgba(0,0,0,0)"),
                            hoverinfo="text",
                            text=[f"{clean_name} (Zero Length at: {location:.2f} m)"],
                            showlegend=False
                        ))
                    else:
                        # Check for the SVG icon
                        icon_path = os.path.join(icon_folder, f"{icon_name}.svg")
                        if os.path.exists(icon_path):
                            fig.add_shape(
                                type="rect",
                                x0=up_location,
                                x1=dn_location,
                                y0=y_offset - 0.5,
                                y1=y_offset + 0.5,
                                fillcolor="white",
                                line=dict(color="black"),
                            )
                            # Add the SVG image
                            fig.add_layout_image(
                                source=encode_image_to_base64(icon_path),
                                x=central_ct_location,
                                y=y_offset,
                                xref="x",
                                yref="y",
                                sizex=length * 0.8,  # Adjust size as needed
                                sizey=1,  # Adjust size as needed
                                xanchor="center",
                                yanchor="middle"
                            )
                        else:
                            # If no icon found, use a simple rectangle with text
                            fig.add_shape(
                                type="rect",
                                x0=up_location,
                                x1=dn_location,
                                y0=y_offset - 0.5,
                                y1=y_offset + 0.5,
                                fillcolor="lightgray",
                                line=dict(color="black"),
                            )

                        # Add hover information
                        fig.add_trace(go.Scatter(
                            x=[(up_location + dn_location) / 2],
                            y=[y_offset],
                            mode="markers",
                            marker=dict(size=1, color="rgba(0,0,0,0)"),
                            hoverinfo="text",
                            text=[f"{clean_name}<br>Start: {up_location:.2f} m<br>End: {dn_location:.2f} m<br>Length: {length:.2f} m"],
                            showlegend=False
                        ))

            elif '_UP' in element and not any('_DN' in df.iloc[j, 1] and clean_element_name(df.iloc[j, 1]) == clean_name
                                              for j in range(idx + 1, len(df))):
                # Represent elements with only '_UP'
                fig.add_shape(
                    type="line",
                    x0=location, x1=location,
                    y0=y_offset - 0.5, y1=y_offset + 0.5,
                    line=dict(color="Blue", dash="dash"),
                )
                # Add hover information
                fig.add_trace(go.Scatter(
                    x=[location],
                    y=[y_offset],
                    mode="markers",
                    marker=dict(size=1, color="rgba(0,0,0,0)"),
                    hoverinfo="text",
                    text=[f"{clean_name} (UP only at: {location:.2f} m)"],
                    showlegend=False
                ))

            # Check if the element is a single-line element (*_CT without *_UP and *_DN)
            elif '_CT' in element and not ('_UP' in df.iloc[idx-1, 1] and '_DN' in df.iloc[idx+1, 1]):
                # Representation should be a vertical dashed line
                fig.add_shape(
                    type="line",
                    x0=location, x1=location,
                    y0=y_offset - 0.5, y1=y_offset + 0.5,
                    line=dict(color="Green", dash="dash"),
                )
                # Add hover information
                fig.add_trace(go.Scatter(
                    x=[location],
                    y=[y_offset],
                    mode="markers",
                    marker=dict(size=1, color="rgba(0,0,0,0)"),
                    hoverinfo="text",
                    text=[f"{clean_name} (CT only at: {location:.2f} m)"],
                    showlegend=False
                ))

    # Step 5: Plot cryomodules as slightly lighter rectangles
    for cm_name, boundaries in cryomodules.items():
        if boundaries['start'] is not None and boundaries['end'] is not None:
            up_location = boundaries['start']
            dn_location = boundaries['end']
            length = dn_location - up_location

            # Plot cryomodule as a light gray rectangle
            fig.add_shape(
                type="rect",
                x0=up_location,
                x1=dn_location,
                y0=-max(y_offsets) - 2,  # Extend below all elements
                y1=max(y_offsets) + 2,   # Extend above all elements
                fillcolor="rgba(200,200,200,0.3)",  # Light gray with transparency
                line=dict(color="gray", dash="dot"),
                layer="below"  # Ensure it's plotted below other elements
            )

            # Add hover information for cryomodule
            fig.add_trace(go.Scatter(
                x=[(up_location + dn_location) / 2],
                y=[0],  # Place at y=0
                mode="markers",
                marker=dict(size=1, color="rgba(0,0,0,0)"),
                hoverinfo="text",
                text=[f"{cm_name}<br>Start: {up_location:.2f} m<br>End: {dn_location:.2f} m<br>Length: {length:.2f} m"],
                showlegend=False
            ))

    # Step 6: Customize layout, adjust y-axis range
    fig.update_layout(
        title="Graphical Representation of Accelerator Lattice Elements",
        xaxis_title="Longitudinal Position (m)",
        yaxis_title="Y Offset (arbitrary units)",
        showlegend=False,
        # Set y-axis range to include all elements plus some padding
        yaxis=dict(
            range=[min(y_offsets) - 3, max(y_offsets) + 3],
            zeroline=True,
            zerolinecolor='black',
            zerolinewidth=1
        ),
        # Enable zoom
        dragmode='zoom',
        # Add modebar buttons
        modebar=dict(
            orientation='v',
            bgcolor='white',
            color='black',
            activecolor='gray'
        ),
        # Add grid
        xaxis=dict(
            showgrid=True,
            gridwidth=1,
            gridcolor='lightgray',
            zeroline=True,
            zerolinecolor='black',
            zerolinewidth=1
        ),
        # Set plot background color
        plot_bgcolor='white',
        # Set paper background color
        paper_bgcolor='white',
        # Add hover mode
        hovermode='closest'
    )

    # Show the plot
    fig.show()

    # Print any missing icons
    print("\nRequired icons:")
    for icon in sorted(required_icons):
        print(f"- {icon}")

if __name__ == "__main__":
    main()
