# Simple KPI Grid Visual

A customizable KPI grid for Power BI reports. Displays measures as cards with optional categories, comparisons, and drill-through.

## Quick Start
1. Download .pbiviz from [Releases](https://github.com/popenheimer/simple-KPI-grid/releases).
2. In Power BI Desktop, go to Visualizations > ... > Import from file.
3. Drag to canvas, add Value (measure) and Category (for grid mode).

## Data Roles
- **Value**: Required measure for KPI values.
- **Category**: Optional grouping for multi-card grid.
- **Comparison Value**: Optional for deltas (absolute/relative).

## Formatting
- **Grid Settings**: Sizing mode (fixed/dynamic), columns/rows.
- **Card Style**: Fill color, border, min width/height.
- See full options in format pane.

## Samples
[Download sample .pbix](path/to/sample.pbix)

## Support
- Issues: [Create a new issue](https://github.com/popenheimer/simple-KPI-grid/issues)
- Forums: [Power BI Community](https://community.powerbi.com/)
- Email: michaelpope143@gmail.com

## Changelog
[See CHANGELOG.md](CHANGELOG.md)

## Limitations
- Max 1000 cards (performance).
- Drill-through requires report pages setup.