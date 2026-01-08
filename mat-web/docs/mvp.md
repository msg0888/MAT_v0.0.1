MVP scope

- Code Criteria and Hazard Data
  - Inputs per Code sheet; CenterPoint required
  - Derived named ranges for loads
- Geometry import
  - Parse .r3d into Geometry tables (all members)
  - Update Code units, axes, version
- Discrete Loads
  - Import via .arc, RFDS PDF parser, or .tml
  - Calculate CaAa, weights, pressures, point and moment loads
  - CFD lookup (10 deg), CCI front/side with trig and ice interpolation
  - Populate appurtenance summary tables on Code sheet
  - Populate Risa File sheet
- Maintenance Loads
  - Lm/Lv point loads from member labels
  - Area loads for platforms/grating
  - Populate Risa File sheet
- RISA-3D export
  - Auto-generate BLC/LC
  - Wind direction selector (presets + custom)
  - Export .r3d with existing naming
- Results
  - Paste raw envelope results
  - Map to section sets and update Code summaries

Non-goals for MVP
- Hosted deployment
- Full RFDS parser replacement
- Automated RISA analysis
