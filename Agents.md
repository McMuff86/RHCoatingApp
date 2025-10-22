
# Rhino CoatingApp Agent

The **Rhino CoatingApp Agent** is a powerful tool for planning and calculating coating tasks directly within Rhino 8 (default version). It supports users in efficiently calculating material consumption, costs, and working time for coating 3D objects, offering an intuitive user interface for maximum productivity. The plugin is developed in **C#** and adheres to **best practices** for Rhino plugin development to ensure high code quality, maintainability, and performance.

**Note:** This `Agents.md` file is kept up to date to reflect changes, new features, or additions to the plugin. Additional specific documentation for individual functions or modules is provided in separate `*.md` files and linked below.

## Main Features

### 1. **Automatic Surface Selection for 3D Objects**
- **Function:** Selects complete 3D Breps (Boundary Representations) in Rhino and automatically calculates the total surface area of the object.
- **Benefit:** Saves time with precise surface calculations without manual intervention.
- **Details:** Supports complex geometries, including curved surfaces and composite objects.

### 2. **Material Configuration via Intuitive User Interface**
- **Function:** Allows selection and configuration of primer and topcoat via a dockable Eto.Forms panel.
- **Input Options:**
  - Material selection from a predefined list via dropdown controls
  - Editable consumption rates (g/m²) for primer and topcoat
  - Editable prices (Fr./kg) for primer and topcoat
  - Application type toggle (Indoor/Outdoor) with specific additive percentages
  - Hardener percentage configuration for primer and topcoat
  - Thinner percentage configuration for primer and topcoat
  - Time factor adjustment using numeric stepper control
  - Direct object selection via "Select Objects" button
  - Ability to clear current selection and start new calculation
- **Implementation:** Dockable Rhino panel that remains visible during work
- **Benefit:** Flexible adaptation to various coating projects with customizable material properties, additive management, and professional user experience.

### 3. **Cost Calculation**
- **Function:** Calculates material costs based on user-defined prices for primer, topcoat, hardeners, and thinners, as well as the calculated surface area.
- **Details:**
  - Costs displayed in Swiss Francs (Fr.)
  - Automatic summation of costs for all selected materials including additives
  - Hardener costs calculated as percentage of base material (editable price, default: Fr. 23/kg)
  - Thinner costs calculated as percentage of base material (editable price, default: Fr. 9/kg)
  - Supports custom pricing and consumption rates
  - Real-time calculation based on actual surface area
  - Separate calculations for indoor and outdoor applications
- **Benefit:** Transparent cost calculation for accurate budgeting and quoting with Swiss currency support, including all necessary additives with configurable prices.

### 4. **Time Estimation**
- **Function:** Allows input of a time factor per square meter to estimate the working time for coating tasks.
- **Details:**
  - Time factor adjustable based on the coating method (e.g., spray painting, brush, etc.).
  - Accounts for optional breaks or additional efforts (e.g., preparation, drying times).
- **Benefit:** Supports project planning with realistic time estimates.

### 5. **Export Function**
- **Function:** Exports calculation results (surface area, material consumption, costs, working time) in a clear table.
- **Formats:** TXT and CSV formats with built-in file save dialog.
- **Implementation:** Integrated export button within the UI dialog with Eto.Forms SaveFileDialog.
- **Details:** Exported files include structured data with metadata (e.g., project name, date, user).
- **Benefit:** Enables easy sharing of results with clients or other departments.

## Additional Features

### 6. **Material Database & Configuration System**
- **Function:** JSON-based configuration system for materials and default settings
- **Configuration File:** `CoatingConfig.json` loaded from plugin directory
- **Customization:** 
  - Pre-defined materials: Standard Primer, Premium Primer, Basic Topcoat, Premium Topcoat
  - User can edit consumption rates and prices directly in the UI
  - Changes apply to current calculation without modifying config file
  - Dropdown selection automatically populates default values from config
  - Administrators can modify `CoatingConfig.json` to change defaults without recompiling
- **Config Structure:**
  - `ApplicationDefaults`: Indoor/Outdoor additive percentages
  - `MaterialPrices`: Hardener and Thinner prices (Fr./kg)
  - `Materials`: List of available materials with properties
  - `DefaultSettings`: Time factor, default application type, default materials
- **Benefit:** Flexible configuration management with easy customization for different regions or material suppliers.

### 7. **Surface Visualization**
- **Function:** Highlights selected surfaces in the Rhino viewport with color to visually verify the calculation areas.
- **Benefit:** Prevents errors through clear visualization of considered geometries.

### 8. **Multilingual Support**
- **Function:** User interface available in multiple languages (e.g., German, English, Spanish).
- **Benefit:** Enables use in international teams.

## Development Guidelines
- **Programming Language:** The plugin is developed in **C#**, adhering to **best practices** for Rhino plugins, including:
  - Modular code structure for easy maintenance and extensibility.
  - Use of the RhinoCommon API for maximum compatibility.
  - Implementation of error handling and logging for robust performance.
  - Adherence to coding standards (e.g., naming conventions, documentation).
- **Referencing Rhino Developer Docs:** In case of uncertainties or implementation questions, the agent first consults the [Rhino Developer Documentation](https://developer.rhino3d.com/). If no suitable solutions are found, it searches online for appropriate approaches (e.g., forums, Stack Overflow, or other trusted sources).

## Technical Implementation Details

### User Interface Architecture
- **Framework:** Eto.Forms (cross-platform UI framework integrated in Rhino 8)
- **Panel Type:** Dockable Rhino Panel using `Rhino.UI.Panel` base class
  - Registered in plugin `OnLoad` method via `Panels.RegisterPanel()`
  - Can be docked anywhere in Rhino interface (left, right, floating)
  - Remains visible while working in Rhino
  - Opened via `Panels.OpenPanel()` command
  - Provides seamless integration with Rhino's UI system
- **Layout System:** `DynamicLayout` and `StackLayout` for flexible, responsive UI layout
- **Key Controls:**
  - `Button` controls for object selection ("Select Objects", "Clear Objects")
  - `RadioButtonList` for application type selection (Indoor/Outdoor)
  - `DropDown` for material selection (primer and topcoat)
  - `NumericStepper` for editable consumption rates (g/m²)
  - `NumericStepper` for editable material prices (Fr./kg for primer and topcoat)
  - `NumericStepper` for editable additive prices (Fr./kg for hardener and thinner)
  - `NumericStepper` for hardener percentages (primer and topcoat)
  - `NumericStepper` for thinner percentages (primer and topcoat)
  - `NumericStepper` for time factor configuration (h/m²)
  - `TextArea` (read-only) for results display
  - `Button` controls for calculations and export
  - `SaveFileDialog` for file export functionality (TXT, CSV, XLSX)
  - Conditional visibility panels for Indoor/Outdoor additive configurations

### Code Organization
- **CoatingApp.cs:** Main command class handling object selection, surface calculation, and panel updates
  - Implements `CoatingApp` command to open panel and calculate surface areas
  - Uses `Panels.OpenPanel()` to show the dockable panel
  - Updates panel with calculated surface area via `UpdateSurfaceArea()` method
- **CoatingPanel.cs:** Dockable Eto.Forms panel implementation
  - Inherits from `Rhino.UI.Panel` base class
  - Contains all material configuration UI and calculation logic
  - GUID: `A1B2C3D4-E5F6-7890-ABCD-EF1234567890`
  - Registered as "Coating Panel" in Rhino
  - Loads default values from `CoatingConfigManager`
  - Delegates export functionality to `CoatingExporter` class
- **CoatingExporter.cs:** Export functionality for coating calculations
  - Static class responsible for all export operations (TXT, CSV, XLSX)
  - Separates export logic from UI logic for better maintainability
  - Provides consistent export format across all file types
  - Includes helper methods for material amount calculations
- **CoatingConfigManager.cs:** Configuration management system
  - Loads settings from `CoatingConfig.json`
  - Provides singleton access to configuration data
  - Falls back to hardcoded defaults if config file not found
  - Thread-safe lazy loading of configuration
- **CoatingConfig.json:** JSON configuration file
  - Defines all default values for materials, prices, and percentages
  - Copied to output directory during build
  - Can be edited by administrators without recompiling
- **RHCoatingAppPlugin.cs:** Plugin initialization and lifecycle management
  - Registers the dockable panel on plugin load
  - Uses `System.Drawing.SystemIcons.Application` as panel icon
- **Data Classes:** `ApplicationType` enum, `MaterialConfig`, `MaterialInfo`, `CostCalculation`, `ObjectInfo` for structured data management
- **Config Data Classes:** `CoatingConfigData`, `ApplicationDefaults`, `AdditiveSettings`, `MaterialPrices`, `MaterialDefinition`, `DefaultSettings` for configuration structure

### Unit Handling
- **Internal Units:** All calculations use Rhino's native millimeters (mm) for surface areas
- **Display Units:** Surface areas displayed in both mm² and m² for user convenience
- **Material Calculations:** Automatically converts to m² for material consumption and cost calculations

## Workflow

### Typical Usage:
1. **Open Coating Panel**: The panel opens automatically when the plugin loads, or use the `CoatingApp` command
2. **Select Application Type**: Choose "Indoor" or "Outdoor" application
3. **Configure Additives**: Set hardener and thinner percentages for the selected application type
   - Indoor defaults: Primer Hardener 6%, Primer Thinner 10%, Topcoat Hardener 15%, Topcoat Thinner 20%
   - Outdoor defaults: Primer Hardener 8%, Primer Thinner 12%, Topcoat Hardener 18%, Topcoat Thinner 25%
4. **Select Objects**: Click "Select Objects" button or run `CoatingApp` command to select 3D objects
5. **Configure Materials**: 
   - Select primer and/or topcoat from dropdown menus
   - Adjust consumption rates (g/m²) as needed (default: 200 g/m²)
   - Adjust prices (Fr./kg) as needed
   - Set time factor (h/m²) based on coating method
6. **Calculate**: Click "Calculate" button to see results including additives
7. **Review Results**: Check surface area, material amounts (including hardeners and thinners), costs, and time estimates
8. **Export** (optional): Export results to TXT, CSV, or XLSX file with complete additive details
9. **New Calculation**: Click "Clear Objects" to reset and start a new calculation

### Key Features:
- Panel remains docked and visible throughout work session
- Can switch between different object selections without closing panel
- Material settings persist until manually changed
- Application type toggle automatically shows/hides relevant additive controls
- Separate default values for indoor and outdoor applications
- Real-time validation prevents calculation errors
- Comprehensive cost calculation including all additives (hardeners and thinners)

## Installation
1. Download the plugin from the official website or the Rhino Plugin Store.
2. Install the plugin via the Rhino Plugin Manager.
3. Restart Rhino to activate the plugin.
4. Access the Coating Panel via `Panels → Coating Panel` or run the `CoatingApp` command.

## Use Cases
- **Industry:** Calculating coating tasks for mechanical engineering, automotive, or aerospace components.
- **Architecture:** Planning coatings for facades or interior components.
- **Product Design:** Estimating material and labor costs for prototypes or series products.

## Related Documents
Specific documentation for individual modules or functions of the plugin is available in separate Markdown files and linked here as they are created:
- [MaterialDatabase.md](MaterialDatabase.md) (for details on material management, to be created as needed).
- [ExportFunction.md](ExportFunction.md) (for details on export formats and configurations, to be created as needed).
- [Visualization.md](Visualization.md) (for details on surface visualization, to be created as needed).
- [Usage.md](Usage.md) (for step-by-step usage instructions).

## Eto.Forms Documentation
For UI development using Eto.Forms in Rhino plugins, refer to these official guides:
- [Forms and Dialogs](https://developer.rhino3d.com/guides/eto/forms-and-dialogs/) - Understanding modal vs non-modal dialogs and semi-modal dialogs
- [Existing Dialogs](https://developer.rhino3d.com/guides/eto/existing-dialogs/) - Using built-in dialogs like MessageBox, ColorPicker, and File dialogs
- [Containers](https://developer.rhino3d.com/guides/eto/containers/) - Layout containers like Panels, Tables, Stack Layouts, Dynamic Layouts, and Pixel Layouts
- [Controls](https://developer.rhino3d.com/guides/eto/controls/) - Available UI controls and their usage
- [Data Context](https://developer.rhino3d.com/guides/eto/view-and-data/data-context/) - Data binding and context management
- [View Models](https://developer.rhino3d.com/guides/eto/view-and-data/view-models/) - Model-View-ViewModel pattern implementation
- [Binding](https://developer.rhino3d.com/guides/eto/view-and-data/binding/) - Data binding between UI and data models

**Note:** New `*.md` files will be created as needed and linked in this section to ensure comprehensive documentation.

## Support and Feedback
- **Documentation:** Detailed guidance is integrated into the plugin or available online.
- **Support:** Contact us at [support@example.com](mailto:support@example.com) for technical questions.
- **Feedback:** We welcome suggestions for improving the plugin via our feedback form.

## Future Developments
- Integration of AI-supported material recommendations based on object properties.
- Advanced visualization options for coating processes (e.g., simulation of layer thicknesses).
- Cloud integration for team collaboration.

## Version Control
A comprehensive `.gitignore` file is included in the repository to exclude build artifacts, IDE-specific files, and other temporary files from version control.