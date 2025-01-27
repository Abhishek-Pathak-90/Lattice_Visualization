# Lattice Visualization Tool

A Python-based visualization tool for plotting accelerator lattice elements using interactive Plotly graphs.

## Features

- Interactive visualization of accelerator lattice elements using Plotly
- Support for multiple input files with symmetric y-offset plotting
- Automatic handling of cryomodules and element boundaries
- SVG/PNG icon support for element visualization
- Interactive hover information for each element
- Comprehensive element type handling (UP, CT, DN markers)
- Element search and filtering functionality
- Mini-map overview for easy navigation
- Tabbed interface with multiple views:
  - Main Plot View
  - Missing Dimension Elements List
  - Required Icons List
  - Icons Table with Preview
- Color-coded element visualization:
  - Zero-length elements shown as dashed lines
  - Elements with length shown as rectangles with icons
  - Cryomodules displayed as semi-transparent backgrounds
- Detailed hover information showing element positions and dimensions
- Icon preview table with guessed element types
- Aspect ratio locking for better visualization
- Proper axis labels and title

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

3. The script will generate an interactive HTML file with:
   - Main interactive plot showing element positions and dimensions
   - Search bar for filtering elements by name
   - Mini-map for easy navigation
   - Tabbed interface for accessing different views
   - Table of icons with previews and guessed types

4. Navigation and Interaction:
   - Use the search bar to filter elements by name
   - Switch between different views using the tabs
   - Use the mini-map for quick navigation
   - Hover over elements for detailed information
   - View icon previews and types in the Icons Table

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
- Element search and filtering
- Mini-map navigation
- Tabbed interface for different views
- Icon preview table
- Customizable layout and appearance

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Author

Abhishek Pathak
