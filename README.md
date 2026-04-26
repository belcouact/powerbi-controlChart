# Statistical Control Chart for Power BI

A custom Power BI visual for Statistical Process Control (SPC) with comprehensive control chart functionality.

## Features

- **Control Lines**: Auto-calculated Center Line (CL), Upper/Lower Control Limits (UCL/LCL) at ±3σ
- **Sigma Zones**: Optional 1σ and 2σ zone lines for enhanced visual analysis
- **Specification Limits**: Configurable USL/LSL values via format pane
- **Target Line**: Optional target/nominal value indicator
- **Phase Control**: Separate statistics and control limits per phase/group
- **Western Electric Rules**: Six out-of-control detection rules
  - Rule 1: 1 point beyond 3σ
  - Rule 2: 9 points same side of center line
  - Rule 3: 6 points trending
  - Rule 4: 14 points alternating
  - Rule 5: 2 of 3 beyond 2σ
  - Rule 6: 4 of 5 beyond 1σ
- **Process Capability**: Cp, Cpk, Pp, Ppk statistics
- **Zone Shading**: Optional color-coded zones (A, B, C)
- **Statistics Panel**: Display mean, std dev, OOC count/percent, and more

## Data Fields

| Field | Type | Description |
|-------|------|-------------|
| Values | Measure | The measurement values to plot |
| Sample (optional) | Category | Sample name/identifier for x-axis (auto-generated if not provided) |
| Phase (optional) | Category | Phase grouping for separate control limits |

## Installation

1. Download the `.pbiviz` file from the `dist/` folder
2. In Power BI Desktop, go to **Visualizations** pane → **Import a visual from a file**
3. Select the downloaded `.pbiviz` file
4. The visual will appear in your visualizations pane

## Usage

1. Drag the visual onto your report canvas
2. Add your measurement values to the **Values** field
3. (Optional) Add sample identifiers to **Sample** field
4. (Optional) Add phase/group field to **Phase** field for phase-based control
5. Configure settings in the **Format** pane

## Format Pane Options

### Control Lines
- Show/hide control lines (CL, UCL, LCL)
- Toggle 1σ and 2σ zone lines
- Customize colors, styles, and widths

### Specification Limits
- Set USL and LSL values
- Customize appearance

### Target Line
- Set target value
- Customize appearance

### Data Points
- Show/hide data points and connecting line
- Customize in-control and out-of-control colors
- Adjust point size and line width

### Phase Control
- Show/hide phase separators
- Enable recalculate limits per phase
- Customize separator appearance

### Out-of-Control Rules
- Enable/disable individual Western Electric rules

### Statistics Panel
- Show/hide statistics overlay
- Select which statistics to display
- Customize position and font size

### Zone Shading
- Enable color-coded sigma zones
- Customize zone colors and opacity

## Development

### Prerequisites
- Node.js
- Power BI Visuals Tools (`powerbi-visuals-tools`)

### Setup
```bash
npm install
```

### Build
```bash
npx powerbi-visuals-tools package
```

The packaged `.pbiviz` file will be generated in the `dist/` folder.

## License

MIT

## Author

Alex Luo (aluO@wlgore.com)
