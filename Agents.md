
# Rhino CoatingApp Agent

The **Rhino CoatingApp Agent** is a powerful tool for planning and calculating coating tasks directly within Rhino 8 (default version). It supports users in efficiently calculating material consumption, costs, and working time for coating 3D objects, offering an intuitive user interface for maximum productivity. The plugin is developed in **C#** and adheres to **best practices** for Rhino plugin development to ensure high code quality, maintainability, and performance.

**Note:** This `Agents.md` file is kept up to date to reflect changes, new features, or additions to the plugin. Additional specific documentation for individual functions or modules is provided in separate `*.md` files and linked below.

## Main Features

### 1. **Automatic Surface Selection for 3D Objects**
- **Function:** Selects complete 3D Breps (Boundary Representations) in Rhino and automatically calculates the total surface area of the object.
- **Benefit:** Saves time with precise surface calculations without manual intervention.
- **Details:** Supports complex geometries, including curved surfaces and composite objects.

### 2. **Material Configuration via Intuitive User Interface**
- **Function:** Allows selection of primer and topcoat via a clear UI.
- **Input Options:**
  - Material selection from a predefined list (extendable with custom materials).
  - Input of material consumption in grams per square meter.
- **Benefit:** Flexible adaptation to various coating projects with immediate configuration preview.

### 3. **Cost Calculation**
- **Function:** Calculates material costs based on the stored prices for primer and topcoat, as well as the calculated surface area.
- **Details:**
  - Automatic summation of costs for all selected materials.
  - Supports multi-layer coatings (e.g., multiple layers of primer or topcoat).
- **Benefit:** Transparent cost calculation for accurate budgeting and quoting.

### 4. **Time Estimation**
- **Function:** Allows input of a time factor per square meter to estimate the working time for coating tasks.
- **Details:**
  - Time factor adjustable based on the coating method (e.g., spray painting, brush, etc.).
  - Accounts for optional breaks or additional efforts (e.g., preparation, drying times).
- **Benefit:** Supports project planning with realistic time estimates.

### 5. **Export Function**
- **Function:** Exports calculation results (surface area, material consumption, costs, working time) in a clear table.
- **Formats:** CSV, XLSX, or JSON (for advanced integrations).
- **Details:** Exported files include structured data with metadata (e.g., project name, date, user).
- **Benefit:** Enables easy sharing of results with clients or other departments.

## Additional Features

### 6. **Material Database**
- **Function:** Integrated database for common primers and topcoats with predefined properties (price, consumption per m²).
- **Customization:** Ability to add or edit custom materials.
- **Benefit:** Speeds up configuration by accessing stored material profiles.

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

## Installation
1. Download the plugin from the official website or the Rhino Plugin Store.
2. Install the plugin via the Rhino Plugin Manager.
3. Restart Rhino to activate the plugin.
4. Access the plugin via the toolbar or the `CoatingApp` command.

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

**Note:** New `*.md` files will be created as needed and linked in this section to ensure comprehensive documentation.

## Support and Feedback
- **Documentation:** Detailed guidance is integrated into the plugin or available online.
- **Support:** Contact us at [support@example.com](mailto:support@example.com) for technical questions.
- **Feedback:** We welcome suggestions for improving the plugin via our feedback form.

## Future Developments
- Integration of AI-supported material recommendations based on object properties.
- Advanced visualization options for coating processes (e.g., simulation of layer thicknesses).
- Cloud integration for team collaboration.

## Create a .gitignore