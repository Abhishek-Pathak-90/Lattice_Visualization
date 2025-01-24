# Lattice Visualization Tool

A Python-based visualization tool for plotting accelerator lattice elements using interactive Plotly graphs.

## Features

- Interactive visualization of accelerator lattice elements using Plotly
- Support for multiple input files with symmetric y-offset plotting
- Automatic handling of cryomodules and element boundaries
- SVG/PNG icon support for element visualization
- Interactive hover information for each element
- Comprehensive element type handling (UP, CT, DN markers)

## Requirements

- Python 3.x
- pandas
- plotly
- tkinter (usually comes with Python)

## Installation

1. Clone this repository:
```bash
git clone [your-repo-url]
cd Lattice_visualization_tool
```

2. Install required packages:
```bash
pip install -r requirements.txt
```

## Usage

1. Run the script:
```bash
python lattice_visualizer.py
```

2. The script will prompt you to:
   - Select one or more Excel files containing lattice element data
   - Choose the directory containing element icons (SVG/PNG)

3. The script will generate an interactive plot showing:
   - Element positions and dimensions
   - Cryomodule boundaries
   - Hover information for each element

## Input File Format

The Excel files should contain the following columns:
1. Element information
2. Element name (with _UP, _CT, _DN suffixes for multi-line elements)
3. Position information
4. Location data

## Features in Detail

### Multi-file Support
- Handles multiple input files with symmetric y-offsets
- Each file's elements are plotted at different vertical positions

### Element Visualization
- Supports both SVG and PNG icons for elements
- Handles zero-length elements with special representation
- Provides distinct visualization for cryomodules

### Interactive Features
- Zoom and pan capabilities
- Detailed hover information for each element
- Customizable layout and appearance

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

[Your chosen license]

## Author

[Your name]
