# Rhino Database Panel Plugin Development

## Goal
Create a native Rhino panel plugin in C# that displays and allows searching/filtering of the door database from the CSV/XLSX files. The panel should enable users to find suitable "Rohling" (raw panels) based on criteria like requirements, thickness, etc.

## Current Status
- Workspace: C:\Users\Adi.Muff\source\repos\RH_DB_Panel
- Existing Python scripts in /python/ build the database CSV.
- Need to port database loading and filtering to C#.
- Basic panel is now functional in Rhino.
- Next: Implement CSV loading, display in DataGridView, and add search/filter functionality.
- Data loading from XLSX works, displayed in GridView.
- Search with wildcard support (* as any) implemented.
- Fixed crash on empty search results.
- Enhanced to multi-criteria column-specific search (e.g., "Tuertyp:VSR Dicke:39").
- Added ability to open and load different XLSX files.
- Added sheet selection for files with multiple sheets.
- Switched to ClosedXML for XLSX handling to avoid licensing issues.

## Development Steps
1. **Setup Basic Plugin Structure**: Use the existing RH_DB_Panel.csproj. Ensure it builds and loads in Rhino.
2. **Create Panel UI**: Use Eto.Forms to create a panel with a DataGridView to display the DataFrame-like data.
3. **Load Data**: Read CSV into a data structure (e.g., DataTable or List of objects).
4. **Implement Search/Filter**: Add text boxes and buttons for filtering.
5. **Integration**: Register the panel in Rhino.
6. **Advanced Features**: Sorting, exporting, etc.

## Next Actions
- Add sorting and export features.
- Improve UI (e.g., dropdown for columns).

Last updated: October 14, 2025
