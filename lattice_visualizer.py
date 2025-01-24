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

    # NEW: strip whitespace from the element column (assuming column index 1 is your element name)
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
                            x0=up_location, x1=dn_location,
                            y0=y_offset - element_height / 2, y1=y_offset + element_height / 2,
                            line=dict(color="RoyalBlue"),
                            fillcolor="LightSkyBlue",
                            opacity=0  # Invisible rectangle
                        )
                        encoded_image = encode_image_to_base64(icon_path)
                        # Add the icon/image inside the rectangle
                        fig.add_layout_image(
                            dict(
                                source=encoded_image,
                                xref="x", yref="y",
                                x=(up_location + dn_location) / 2,
                                y=y_offset,
                                sizex=length,
                                sizey=length,
                                xanchor="center", yanchor="middle"
                            )
                        )
                    else:
                        # If no SVG icon, fall back to a rectangle
                        fig.add_shape(
                            type="rect",
                            x0=up_location, x1=dn_location,
                            y0=y_offset - element_height / 2, y1=y_offset + element_height / 2,
                            line=dict(color="RoyalBlue"),
                            fillcolor="LightSkyBlue",
                            opacity=0.5
                        )

                    # Add invisible hover-enabled marker with element details
                    fig.add_trace(go.Scatter(
                        x=[central_ct_location],
                        y=[y_offset],
                        mode="markers",
                        marker=dict(size=1, color="rgba(0,0,0,0)"),
                        hoverinfo="text",
                        text=[f"{clean_name} (Length: {length:.2f} m, Start: {up_location:.2f} m, End: {dn_location:.2f} m)"],
                        showlegend=False
                    ))
                    element_parameters[clean_name] = {"type": "quadrupole", "gradient": 0.0, "location": central_ct_location}

            # Handle elements with only '_UP' (no corresponding '_DN')
            elif '_UP' in element and not any('_DN' in df.iloc[j, 1] and clean_element_name(df.iloc[j, 1]) == clean_name
                                              for j in range(idx + 1, len(df))):
                # Represent elements with only '_UP'
                fig.add_shape(
                    type="line",
                    x0=location, x1=location,
                    y0=y_offset - 0.5, y1=y_offset + 0.5,
                    line=dict(color="Orange", dash="dash"),
                )
                # Add invisible hover-enabled marker
                fig.add_trace(go.Scatter(
                    x=[location],
                    y=[y_offset],
                    mode="markers",
                    marker=dict(size=1, color="rgba(0,0,0,0)"),
                    hoverinfo="text",
                    text=[f"{clean_name} (Only _UP at: {location:.2f} m)"],
                    showlegend=False
                ))

        # Check if the element is a single-line element (*_CT without *_UP and *_DN)
        elif '_CT' in element and not ('_UP' in df.iloc[idx-1, 1] and '_DN' in df.iloc[idx+1, 1]):
            # Representation should be a vertical dashed line
            fig.add_shape(
                type="line",
                x0=location, x1=location,
                y0=y_offset - 0.5, y1=y_offset + 0.5,
                line=dict(color="Green", dash="dash")
            )
            # Add invisible hover-enabled marker
            fig.add_trace(go.Scatter(
                x=[location],
                y=[y_offset],
                mode="markers",
                marker=dict(size=1, color="rgba(0,0,0,0)"),
                hoverinfo="text",
                text=[f"{clean_name} (CT position: {location:.2f} m)"],
                showlegend=False
            ))
            missing_dimensions_elements.append(f"{clean_name} at {location:.2f} m")
            element_parameters[clean_name] = {"type": "solenoid", "field": 0.0, "location": location}

        # Handle elements with only '_DN' (no corresponding '_UP')
        elif '_DN' in element and not any('_UP' in df.iloc[j, 1] and clean_element_name(df.iloc[j, 1]) == clean_name
                                          for j in range(0, idx)):
            # Represent elements with only '_DN'
            fig.add_shape(
                type="line",
                x0=location, x1=location,
                y0=y_offset - 0.5, y1=y_offset + 0.5,
                line=dict(color="Purple", dash="dash"),
            )
            # Add invisible hover-enabled marker
            fig.add_trace(go.Scatter(
                x=[location],
                y=[y_offset],
                mode="markers",
                marker=dict(size=1, color="rgba(0,0,0,0)"),
                hoverinfo="text",
                text=[f"{clean_name} (Only _DN at: {location:.2f} m)"],
                showlegend=False
            ))

    # Step 5: Plot cryomodules as slightly lighter rectangles
    for cm_name, boundaries in cryomodules.items():
        if boundaries['start'] is not None and boundaries['end'] is not None:
            up_location = boundaries['start']
            dn_location = boundaries['end']
            cm_length = dn_location - up_location
            fig.add_shape(
                type="rect",
                x0=up_location, x1=dn_location,
                y0=min(y_offsets) - 1, y1=max(y_offsets) + 1,  # Cover the entire y-range
                line=dict(color="Gray"),
                fillcolor="Gray",
                opacity=0.4
            )
            # Add invisible hover-enabled marker for cryomodules
            fig.add_trace(go.Scatter(
                x=[(up_location + dn_location) / 2],
                y=[0],  # Position at y=0
                mode="markers",
                marker=dict(size=1, color="rgba(0,0,0,0)"),
                hoverinfo="text",
                text=[f"{cm_name} (Cryomodule, Length: {cm_length:.2f} m)"],
                showlegend=False
            ))

# Step 6: Customize layout, adjust y-axis range
fig.update_layout(
    title="Graphical Representation of Accelerator Elements",
    xaxis_title="Longitudinal Position (m)",
    yaxis_title="Y Offset (arbitrary units)",
    yaxis=dict(
        range=[min(y_offsets) - 2, max(y_offsets) + 2],
        zeroline=True,
        showticklabels=True
    ),
    showlegend=False
)

# Step 7: Save plot and missing elements as HTML
plot_html = fig.to_html(full_html=False, include_plotlyjs="cdn", div_id="plotly-graph")

# Create HTML structure with tabs, scaling inputs, and missing elements
output_html = 'interactive_lattice.html'
with open(output_html, 'w', encoding='utf-8') as f:
    f.write("""
<!DOCTYPE html>
<html>
<head>
    <title>Interactive Accelerator Lattice Plot</title>
    <style>
        /* Style the tab */
        .tab {
            overflow: hidden;
            border: 1px solid #ccc;
            background-color: #f1f1f1;
        }
        /* Style the buttons inside the tab */
        .tab button {
            background-color: inherit;
            float: left;
            border: none;
            outline: none;
            cursor: pointer;
            padding: 14px 16px;
            transition: 0.3s;
        }
        /* Change background color of buttons on hover */
        .tab button:hover {
            background-color: #ddd;
        }
        /* Create an active/current tablink class */
        .tab button.active {
            background-color: #ccc;
        }
        /* Style the tab content */
        .tabcontent {
            display: none;
            padding: 6px 12px;
            border: 1px solid #ccc;
            border-top: none;
        }
    </style>
</head>
<body>

<h2>Interactive Accelerator Lattice Plot with Cryomodules</h2>

<div class="tab">
  <button class="tablinks" onclick="openTab(event, 'Plot')" id="defaultOpen">Interactive Plot</button>
  <button class="tablinks" onclick="openTab(event, 'MissingElements')">Missing Dimension Elements</button>
  <button class="tablinks" onclick="openTab(event, 'RequiredIcons')">Required Icon Names</button>
</div>

<div id="Plot" class="tabcontent">
""")
    f.write(plot_html)
    f.write("</div>\n<div id='MissingElements' class='tabcontent'><h3>Missing Dimension Elements</h3><ul>")

    # Write missing elements to HTML
    for element in missing_dimensions_elements:
        f.write(f"<li>{element}</li>")

    f.write("</ul></div>\n<div id='RequiredIcons' class='tabcontent'><h3>Required Icon Names</h3><ul>")

    # Write required icons to HTML
    for icon_name in sorted(required_icons):
        f.write(f"<li>{icon_name}</li>")

    f.write("""
</ul></div>

<script>
// Open tab functionality
function openTab(evt, tabName) {
    var i, tabcontent, tablinks;
    tabcontent = document.getElementsByClassName("tabcontent");
    for (i = 0; i < tabcontent.length; i++) {
        tabcontent[i].style.display = "none";
    }
    tablinks = document.getElementsByClassName("tablinks");
    for (i = 0; i < tablinks.length; i++) {
        tablinks[i].className = tablinks[i].className.replace(" active", "");
    }
    document.getElementById(tabName).style.display = "block";
    evt.currentTarget.className += " active";
}

// Open the default tab (the Plot tab)
document.getElementById("defaultOpen").click();
</script>

</body>
</html>
""")

print(f"Interactive HTML file with cryomodules plotted saved as: {output_html}")
