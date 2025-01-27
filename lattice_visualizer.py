import pandas as pd
import plotly.graph_objects as go
import tkinter as tk
from tkinter import filedialog
import os
import base64
import json

# If you want all rectangles to have a fixed vertical size:
FIXED_ELEMENT_HEIGHT = 2.0

def get_icon_name(element):
    """Take the first three dash-delimited parts as icon name, trimming whitespace."""
    element = element.strip()
    return "-".join(element.split("-")[:3])

def clean_element_name(element):
    """Remove _UP, _CT, _DN from element name."""
    return element.replace("_UP", "").replace("_CT", "").replace("_DN", "")

def encode_image_to_base64(image_path):
    """
    Convert a PNG or SVG file to a base64 data URI string.
    """
    file_extension = os.path.splitext(image_path)[1].lower()
    if file_extension == ".svg":
        mime_type = "image/svg+xml"
    elif file_extension == ".png":
        mime_type = "image/png"
    else:
        mime_type = "image/png"

    with open(image_path, 'rb') as f:
        encoded = base64.b64encode(f.read()).decode('ascii')
    return f"data:{mime_type};base64,{encoded}"

def generate_symmetric_offsets(n, step):
    """
    Generate a list of y-offsets, symmetric about y=0, spaced by 'step'.
    """
    offsets = []
    if n % 2 == 1:
        offsets.append(0)
        for i in range(1, (n // 2) + 1):
            offsets.append(i * step)
            offsets.append(-i * step)
    else:
        for i in range(1, (n // 2) + 1):
            offsets.append(i * step)
            offsets.append(-i * step)
    return offsets[:n]

def guess_icon_type(icon_filename):
    """
    Very naive guess of 'type' based on partial substring matches in the icon filename.
    Adjust for your actual naming conventions or logic.
    """
    name_lower = icon_filename.lower()
    if "qd" in name_lower or "qf" in name_lower or "qbtl" in name_lower or "ql" in name_lower:
        return "Quadrupole"
    elif "bpm" in name_lower:
        return "Beam Position Monitor (BPM)"
    elif "ycor" in name_lower or "xcor" in name_lower or "xycor" in name_lower:
        return "Corrector"
    elif "dpl" in name_lower or "dpll" in name_lower:
        return "Dipole"
    elif "sol" in name_lower or "solenoid" in name_lower:
        return "Solenoid"
    elif "cav" in name_lower:
        return "Cavity"
    elif "3ws" in name_lower or "xyws" in name_lower:
        return "Wire Scanner"
    elif "marker" in name_lower:
        return "Marker"
    else:
        return "Unknown"

# 1) File selection via tkinter
root = tk.Tk()
root.withdraw()
file_paths = filedialog.askopenfilenames(
    title="Select Excel file(s)",
    filetypes=[("Excel files", "*.xlsx")]
)
if not file_paths:
    print("No files selected, exiting.")
    exit()

# 2) Icon folder selection
icon_folder = filedialog.askdirectory(
    title="Select the folder containing element icons"
)

# Prepare the main figure
fig = go.Figure()
offset_step = 5
n_files = len(file_paths)
y_offsets = generate_symmetric_offsets(n_files, offset_step)

missing_dimensions_elements = []
required_icons = set()
cryomodules = {}

# We'll track min/max x-limits for building the mini-map
global_min_x = float('inf')
global_max_x = float('-inf')

# ============ Identify Cryomodules & Coerce Numeric Location ============
for file_path in file_paths:
    df = pd.read_excel(file_path)

    # Force column #4 (location) to numeric, drop invalid
    df.iloc[:, 4] = pd.to_numeric(df.iloc[:, 4], errors="coerce")
    df = df.dropna(subset=[df.columns[4]])

    # Clean element column (#1)
    df.iloc[:, 1] = df.iloc[:, 1].astype(str).fillna('').str.strip()

    # Remove cryomodule _CT rows if needed
    df = df[~df.iloc[:, 1].str.contains('CM') | ~df.iloc[:, 1].str.contains('_CT')]

    # Update global min/max
    for idx in range(len(df)):
        location = df.iloc[idx, 4]
        if location < global_min_x:
            global_min_x = location
        if location > global_max_x:
            global_max_x = location

    # Check cryomodules
    for idx in range(len(df)):
        element = df.iloc[idx, 1]
        location = df.iloc[idx, 4]
        if 'CM' in element and ('_UP' in element or '_DN' in element):
            cm_name = element.replace('_UP', '').replace('_DN', '')
            if cm_name not in cryomodules:
                cryomodules[cm_name] = {'start': None, 'end': None}
            if '_UP' in element:
                cryomodules[cm_name]['start'] = location
            elif '_DN' in element:
                cryomodules[cm_name]['end'] = location

# ============= Build the Lattice (Shapes, Icons, Traces) =============
for file_index, file_path in enumerate(file_paths):
    df = pd.read_excel(file_path)

    # Coerce location column
    df.iloc[:, 4] = pd.to_numeric(df.iloc[:, 4], errors="coerce")
    df = df.dropna(subset=[df.columns[4]])

    df.iloc[:, 1] = df.iloc[:, 1].astype(str).fillna('').str.strip()
    df = df[~df.iloc[:, 1].str.contains('CM') | ~df.iloc[:, 1].str.contains('_CT')]

    y_offset = y_offsets[file_index]

    for idx in range(len(df)):
        element = df.iloc[idx, 1]
        location = df.iloc[idx, 4]

        if 'CM' in element:
            continue

        icon_name = get_icon_name(element)
        required_icons.add(f"{icon_name}.svg")
        clean_name = clean_element_name(element)

        if location < global_min_x:
            global_min_x = location
        if location > global_max_x:
            global_max_x = location

        # ========== 3-line element check ===========
        if '_UP' in element:
            up_location = location
            dn_location = None
            central_ct_location = None

            for j in range(idx + 1, len(df)):
                next_elem = df.iloc[j, 1]
                if '_DN' in next_elem and clean_element_name(next_elem) == clean_name:
                    dn_location = df.iloc[j, 4]
                    break

            if dn_location is not None:
                for j in range(idx + 1, len(df)):
                    next_elem = df.iloc[j, 1]
                    if '_CT' in next_elem and clean_element_name(next_elem) == clean_name:
                        if '_P1_' not in next_elem and '_P2_' not in next_elem:
                            central_ct_location = df.iloc[j, 4]
                            break
                if central_ct_location is None:
                    for j in range(idx + 1, len(df)):
                        next_elem = df.iloc[j, 1]
                        if '_CT' in next_elem and clean_element_name(next_elem) == clean_name:
                            central_ct_location = df.iloc[j, 4]
                            break
                    if central_ct_location is None:
                        print(f"No _CT found for {element}, skipping.")
                        continue

                length = dn_location - up_location
                if length == 0:
                    # zero length => line
                    shape = dict(
                        type="line",
                        x0=up_location, x1=up_location,
                        y0=y_offset - 0.5, y1=y_offset + 0.5,
                        line=dict(color="Red", dash="dash"),
                        label=dict(
                            text=clean_name,
                            font=dict(color="rgba(0,0,0,0)", size=1)
                        )
                    )
                    fig.add_shape(shape)

                    hover_txt = f"<b>{clean_name}</b><br>Start/End: {up_location:.2f} m (Zero length)"
                    trace = go.Scatter(
                        x=[up_location],
                        y=[y_offset],
                        mode="markers",
                        marker=dict(size=5, color="red"),
                        hoverinfo="text",
                        text=[hover_txt],
                        name=clean_name,
                        showlegend=False,
                        hovertemplate="%{text}<extra></extra>"
                    )
                    fig.add_trace(trace)
                else:
                    y0 = y_offset - (FIXED_ELEMENT_HEIGHT / 2)
                    y1 = y_offset + (FIXED_ELEMENT_HEIGHT / 2)

                    rect_shape = dict(
                        type="rect",
                        x0=up_location, x1=dn_location,
                        y0=y0, y1=y1,
                        line=dict(color="RoyalBlue"),
                        fillcolor="LightSkyBlue",
                        opacity=0.3,
                        label=dict(
                            text=clean_name,
                            font=dict(color="rgba(0,0,0,0)", size=1)
                        )
                    )
                    fig.add_shape(rect_shape)

                    icon_path = os.path.join(icon_folder, f"{icon_name}.svg")
                    if os.path.exists(icon_path):
                        encoded_image = encode_image_to_base64(icon_path)
                        image_dict = dict(
                            source=encoded_image,
                            xref="x",
                            yref="y",
                            x=(up_location + dn_location) / 2,
                            y=y_offset,
                            sizex=length,
                            sizey=FIXED_ELEMENT_HEIGHT,
                            xanchor="center",
                            yanchor="middle",
                            sizing="stretch",
                            name=clean_name
                        )
                        fig.add_layout_image(image_dict)

                    hover_txt = (f"<b>{clean_name}</b><br>"
                                 f"Start: {up_location:.2f} m<br>"
                                 f"End: {dn_location:.2f} m<br>"
                                 f"Length: {length:.2f} m")
                    trace = go.Scatter(
                        x=[central_ct_location],
                        y=[y_offset],
                        mode="markers",
                        marker=dict(size=5, color="blue"),
                        hoverinfo="text",
                        text=[hover_txt],
                        name=clean_name,
                        showlegend=False,
                        hovertemplate="%{text}<extra></extra>"
                    )
                    fig.add_trace(trace)
            else:
                # _UP but no _DN
                shape = dict(
                    type="line",
                    x0=up_location, x1=up_location,
                    y0=y_offset - 0.5, y1=y_offset + 0.5,
                    line=dict(color="Orange", dash="dash"),
                    label=dict(
                        text=clean_name,
                        font=dict(color="rgba(0,0,0,0)", size=1)
                    )
                )
                fig.add_shape(shape)

                hover_txt = f"<b>{clean_name}</b><br>Only _UP<br>Loc: {up_location:.2f} m"
                trace = go.Scatter(
                    x=[up_location],
                    y=[y_offset],
                    mode="markers",
                    marker=dict(size=5, color="orange"),
                    hoverinfo="text",
                    text=[hover_txt],
                    name=clean_name,
                    showlegend=False,
                    hovertemplate="%{text}<extra></extra>"
                )
                fig.add_trace(trace)

        elif '_CT' in element and not (
            '_UP' in df.iloc[idx - 1, 1] and '_DN' in df.iloc[idx + 1, 1]
        ):
            # Single-line element
            shape = dict(
                type="line",
                x0=location, x1=location,
                y0=y_offset - 0.5, y1=y_offset + 0.5,
                line=dict(color="Green", dash="dash"),
                label=dict(
                    text=clean_name,
                    font=dict(color="rgba(0,0,0,0)", size=1)
                )
            )
            fig.add_shape(shape)

            hover_txt = f"<b>{clean_name}</b><br>Single-line CT<br>Loc: {location:.2f} m"
            trace = go.Scatter(
                x=[location],
                y=[y_offset],
                mode="markers",
                marker=dict(size=5, color="green"),
                hoverinfo="text",
                text=[hover_txt],
                name=clean_name,
                showlegend=False,
                hovertemplate="%{text}<extra></extra>"
            )
            fig.add_trace(trace)
            missing_dimensions_elements.append(f"{clean_name} at {location:.2f} m")

        elif '_DN' in element and not any(
            '_UP' in df.iloc[j, 1] and clean_element_name(df.iloc[j, 1]) == clean_name
            for j in range(0, idx)
        ):
            shape = dict(
                type="line",
                x0=location, x1=location,
                y0=y_offset - 0.5, y1=y_offset + 0.5,
                line=dict(color="Purple", dash="dash"),
                label=dict(
                    text=clean_name,
                    font=dict(color="rgba(0,0,0,0)", size=1)
                )
            )
            fig.add_shape(shape)

            hover_txt = f"<b>{clean_name}</b><br>Only _DN<br>Loc: {location:.2f} m"
            trace = go.Scatter(
                x=[location],
                y=[y_offset],
                mode="markers",
                marker=dict(size=5, color="purple"),
                hoverinfo="text",
                text=[hover_txt],
                name=clean_name,
                showlegend=False,
                hovertemplate="%{text}<extra></extra>"
            )
            fig.add_trace(trace)

# ============== Plot Cryomodules ==============
for cm_name, boundaries in cryomodules.items():
    if boundaries['start'] is not None and boundaries['end'] is not None:
        up_location = boundaries['start']
        dn_location = boundaries['end']
        if up_location > dn_location:
            up_location, dn_location = dn_location, up_location
        cm_length = dn_location - up_location

        shape = dict(
            type="rect",
            x0=up_location, x1=dn_location,
            y0=min(y_offsets) - 1,
            y1=max(y_offsets) + 1,
            line=dict(color="Gray"),
            fillcolor="Gray",
            opacity=0.4,
            label=dict(
                text=f"CRYOMODULE-{cm_name}",
                font=dict(color="rgba(0,0,0,0)", size=1)
            )
        )
        fig.add_shape(shape)

        hover_txt = (f"<b>{cm_name}</b><br>Cryomodule<br>"
                     f"Start: {up_location:.2f} m<br>End: {dn_location:.2f} m<br>"
                     f"Length: {cm_length:.2f} m")
        cm_trace = go.Scatter(
            x=[(up_location + dn_location) / 2],
            y=[0],
            mode="markers",
            marker=dict(size=5, color="gray"),
            hoverinfo="text",
            text=[hover_txt],
            showlegend=False,
            name=f"CRYOMODULE-{cm_name}",
            hovertemplate="%{text}<extra></extra>"
        )
        fig.add_trace(cm_trace)

# ============== Lock aspect ratio, initial range ==============
fig.update_layout(
    title="SC Linac Lattice (Hover, Filter, Mini-Map, Icons Table w/ Preview)",
    xaxis_title="Longitudinal Position (m)",
    yaxis_title="Y Offset (arbitrary units)",
    yaxis=dict(
        scaleanchor="x",
        scaleratio=1,
        range=[min(y_offsets) - 2, max(y_offsets) + 2]
    ),
    showlegend=False
)

# Convert to dict for custom embedding (shapes/images in JS)
fig_dict = fig.to_dict()
original_data = fig_dict["data"]
original_layout = fig_dict["layout"]
original_shapes = original_layout.get("shapes", [])
original_images = original_layout.get("images", [])

layout_for_json = {k: v for k, v in original_layout.items() if k not in ["shapes","images"]}

data_json = json.dumps(original_data)
shapes_json = json.dumps(original_shapes)
images_json = json.dumps(original_images)
layout_json = json.dumps(layout_for_json)

plot_html = fig.to_html(
    full_html=False,
    include_plotlyjs=False,
    div_id="plotly-graph"
)

# Build a list of icons (sorted) plus guessed types, plus a "Preview" if found
icon_rows = []
for icon_filename in sorted(required_icons):
    guessed = guess_icon_type(icon_filename)
    icon_path = os.path.join(icon_folder, icon_filename)
    if os.path.exists(icon_path):
        # Generate a small thumbnail (e.g. 40px wide)
        encoded = encode_image_to_base64(icon_path)
        # We embed an <img> tag
        preview_html = f'<img src="{encoded}" width="40" alt="{icon_filename}" />'
    else:
        preview_html = "NOT FOUND"

    icon_rows.append((icon_filename, guessed, preview_html))

output_html = "interactive_lattice.html"
with open(output_html, "w", encoding="utf-8") as f:
    f.write(f"""
<!DOCTYPE html>
<html>
<head>
    <title>SC Linac Lattice + Icons Table with Preview</title>
    <style>
        .tab {{
            overflow: hidden;
            border: 1px solid #ccc;
            background-color: #f1f1f1;
        }}
        .tab button {{
            background-color: inherit;
            float: left;
            border: none;
            outline: none;
            cursor: pointer;
            padding: 14px 16px;
            transition: 0.3s;
        }}
        .tab button:hover {{
            background-color: #ddd;
        }}
        .tab button.active {{
            background-color: #ccc;
        }}
        .tabcontent {{
            display: none;
            padding: 6px 12px;
            border: 1px solid #ccc;
            border-top: none;
        }}
        .search-container {{
            margin: 10px 0;
        }}
        /* mini-map styling */
        #miniPlot {{
            width: 500px;
            height: 100px;
            border: 1px solid #ccc;
            margin-top: 10px;
        }}
        /* Table of icons */
        table.icons-table {{
            border-collapse: collapse;
            width: 100%;
        }}
        table.icons-table th, table.icons-table td {{
            border: 1px solid #ccc;
            padding: 6px;
            text-align: left;
        }}
        table.icons-table th {{
            background-color: #f2f2f2;
        }}
    </style>
</head>
<body>

<h2>SC Linac Lattice: Hover Info, Substring Filter, Mini-Map, Icons Table w/ Previews</h2>

<div class="search-container">
  <label>Search (substring):</label>
  <input type="text" id="searchTerm" placeholder="e.g. QD or ARC1-QD"/>
  <button onclick="filterElements()">Filter</button>
  <button onclick="resetPlot()">Reset</button>
</div>

<div class="tab">
  <button class="tablinks" onclick="openTab(event, 'Plot')" id="defaultOpen">Plot</button>
  <button class="tablinks" onclick="openTab(event, 'MissingDim')">Missing Dim Elements</button>
  <button class="tablinks" onclick="openTab(event, 'RequiredIcons')">Required Icons</button>
  <button class="tablinks" onclick="openTab(event, 'IconsTable')">Icons Table</button>
</div>

<div id="Plot" class="tabcontent">
  {plot_html}
  <h3>Mini-Map Overview:</h3>
  <div id="miniPlot"></div>
</div>

<div id="MissingDim" class="tabcontent">
  <h3>Missing Dimension Elements</h3>
  <ul>
""")

    for elem in missing_dimensions_elements:
        f.write(f"<li>{elem}</li>")
    f.write("</ul></div>\n")

    f.write("<div id='RequiredIcons' class='tabcontent'><h3>Required Icon Names</h3><ul>")
    for icon in sorted(required_icons):
        f.write(f"<li>{icon}</li>")
    f.write("</ul></div>\n")

    # Our new IconsTable tab with preview
    f.write("<div id='IconsTable' class='tabcontent'><h3>Icons and Their Type & Preview</h3>\n")
    f.write("<table class='icons-table'><thead><tr>")
    f.write("<th>Icon Filename</th><th>Guessed Type</th><th>Preview</th>")
    f.write("</tr></thead><tbody>\n")

    for (icon_filename, guessed_type, preview_html) in icon_rows:
        f.write(f"<tr><td>{icon_filename}</td><td>{guessed_type}</td><td>{preview_html}</td></tr>\n")

    f.write("</tbody></table></div>\n")

    f.write(f"""
<script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
<script>
  var originalData = {data_json};
  var originalShapes = {shapes_json};
  var originalImages = {images_json};
  var originalLayout = {layout_json};

  var masterFigure = {{
    data: originalData,
    layout: originalLayout
  }};
  masterFigure.layout.shapes = originalShapes;
  masterFigure.layout.images = originalImages;

  var graphDiv = document.getElementById('plotly-graph');

  // We'll do an initial newPlot, so user sees the lattice right away:
  Plotly.newPlot(graphDiv, masterFigure.data, masterFigure.layout);

  // Animate-based react function for partial filtering
  function animateReact(newData, newShapes, newImages) {{
    let newLayout = JSON.parse(JSON.stringify(masterFigure.layout));
    newLayout.shapes = newShapes;
    newLayout.images = newImages;

    Plotly.react(graphDiv, newData, newLayout, {{
      transition: {{
        duration: 500,
        easing: 'cubic-in-out'
      }},
      frame: {{
        duration: 500
      }}
    }});
  }}

  function filterElements() {{
    var term = document.getElementById('searchTerm').value.trim().toLowerCase();
    if(!term) {{
      alert("Type something or click Reset.");
      return;
    }}

    let filteredData = originalData.filter(d => {{
      if(!d.name) return false;
      return d.name.toLowerCase().includes(term);
    }});
    let filteredShapes = originalShapes.filter(s => {{
      if(!s.label || !s.label.text) return false;
      return s.label.text.toLowerCase().includes(term);
    }});
    let filteredImages = originalImages.filter(img => {{
      if(!img.name) return false;
      return img.name.toLowerCase().includes(term);
    }});

    animateReact(filteredData, filteredShapes, filteredImages);
  }}

  function resetPlot() {{
    document.getElementById('searchTerm').value = "";
    animateReact(originalData, originalShapes, originalImages);
  }}

  function openTab(evt, tabName) {{
    var i, tabcontent, tablinks;
    tabcontent = document.getElementsByClassName("tabcontent");
    for(i=0; i<tabcontent.length; i++) {{
      tabcontent[i].style.display="none";
    }}
    tablinks = document.getElementsByClassName("tablinks");
    for(i=0; i<tablinks.length; i++) {{
      tablinks[i].className = tablinks[i].className.replace(" active","");
    }}
    document.getElementById(tabName).style.display="block";
    evt.currentTarget.className += " active";
  }}
  document.getElementById("defaultOpen").click();

  // =============== MINI-PLOT ===============
  var miniDiv = document.getElementById('miniPlot');
  var minVal = {float(global_min_x)};
  var maxVal = {float(global_max_x)};

  var miniData = [{{
    x: [minVal, maxVal],
    y: [0, 0],
    mode: 'lines',
    line: {{color:'black'}},
    hoverinfo:'none',
    showlegend:false
  }}];

  var highlightShape = {{
    type:'rect',
    xref:'x',
    yref:'y',
    x0: minVal,
    x1: maxVal,
    y0:-0.3,
    y1:0.3,
    fillcolor:'rgba(255,0,0,0.3)',
    line:{{width:0}}
  }};

  var miniLayout = {{
    margin: {{l:40, r:20, t:20, b:20}},
    xaxis: {{
      range:[minVal-5, maxVal+5],
      showgrid:false
    }},
    yaxis: {{
      range:[-1,1],
      showgrid:false
    }},
    shapes:[highlightShape]
  }};

  Plotly.newPlot(miniDiv, miniData, miniLayout, {{staticPlot:true}});

  graphDiv.on('plotly_relayout', function(ev) {{
    if(ev['xaxis.range[0]']!==undefined && ev['xaxis.range[1]']!==undefined) {{
      let left = ev['xaxis.range[0]'];
      let right = ev['xaxis.range[1]'];

      Plotly.relayout(miniDiv, {{
        'shapes[0].x0': left,
        'shapes[0].x1': right
      }}, [{{
        transition: {{duration:500, easing:'cubic-in-out'}}
      }}]);
    }}
  }});
</script>
</body>
</html>
""")

print(f"Interactive HTML file saved as: {output_html}")
