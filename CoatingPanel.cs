using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Eto.Forms;
using Eto.Drawing;
using Rhino;
using Rhino.UI;
using ClosedXML.Excel;
using System.Linq;

namespace RHCoatingApp
{
    /// <summary>
    /// Dockable Rhino Panel for CoatingApp material configuration and results display
    /// </summary>
    [Guid("A1B2C3D4-E5F6-7890-ABCD-EF1234567890")]
    public class CoatingPanel : Panel
    {
        public static Guid PanelId => typeof(CoatingPanel).GUID;

        private Label surfaceAreaLabel;
        private DropDown primerDropDown;
        private DropDown topcoatDropDown;
        private NumericStepper primerPriceStepper;
        private NumericStepper topcoatPriceStepper;
        private NumericStepper primerConsumptionStepper;
        private NumericStepper topcoatConsumptionStepper;
        private NumericStepper timeFactorStepper;
        private NumericStepper hardenerPriceStepper;
        private NumericStepper thinnerPriceStepper;
        private NumericStepper primerCoatMultiplierStepper;
        private NumericStepper topcoatCoatMultiplierStepper;
        private TextArea resultsTextArea;
        private Button selectObjectsButton;
        private Button clearObjectsButton;
        private Button calculateButton;
        private Button exportButton;
        
        // Application type toggle
        private RadioButtonList applicationTypeRadio;
        
        // Hardener and Thinner controls for Indoor
        private NumericStepper indoorPrimerHardenerStepper;
        private NumericStepper indoorPrimerThinnerStepper;
        private NumericStepper indoorTopcoatHardenerStepper;
        private NumericStepper indoorTopcoatThinnerStepper;
        
        // Hardener and Thinner controls for Outdoor
        private NumericStepper outdoorPrimerHardenerStepper;
        private NumericStepper outdoorPrimerThinnerStepper;
        private NumericStepper outdoorTopcoatHardenerStepper;
        private NumericStepper outdoorTopcoatThinnerStepper;
        
        // Containers for conditional visibility
        private StackLayout indoorAdditivePanel;
        private StackLayout outdoorAdditivePanel;

        // Calculation factor controls
        private Dictionary<string, NumericStepper> calculationFactorSteppers;

        private double surfaceArea_mm2;
        private List<ObjectInfo> selectedObjects;
        private Dictionary<string, MaterialInfo> materials;

        public MaterialConfig MaterialConfig { get; private set; }
        public CostCalculation Costs { get; private set; }
        public double EstimatedTime { get; private set; }
        public CalculationResult CalculationResult { get; private set; }

        public CoatingPanel()
        {
            this.surfaceArea_mm2 = 0;
            this.selectedObjects = new List<ObjectInfo>();
            InitializeMaterials();
            InitializeUI();
        }

        /// <summary>
        /// Updates the surface area and object information displayed in the panel
        /// </summary>
        public void UpdateSurfaceArea(double totalSurfaceArea_mm2, List<ObjectInfo> objects)
        {
            this.surfaceArea_mm2 = totalSurfaceArea_mm2;
            this.selectedObjects = objects ?? new List<ObjectInfo>();
            
            if (surfaceAreaLabel != null)
            {
                string objectCount = selectedObjects.Count > 0 ? $"{selectedObjects.Count} object(s)" : "No objects";
                surfaceAreaLabel.Text = surfaceArea_mm2 > 0
                    ? $"{objectCount} - {surfaceArea_mm2:F2} mm² ({surfaceArea_mm2 / 1000000:F4} m²)"
                    : "No objects selected";
            }
            // Reset results
            if (resultsTextArea != null)
            {
                resultsTextArea.Text = "Click 'Calculate' to see results...";
            }
            if (exportButton != null)
            {
                exportButton.Enabled = false;
            }
        }

        /// <summary>
        /// Legacy method for backward compatibility
        /// </summary>
        public void UpdateSurfaceArea(double surfaceArea_mm2)
        {
            UpdateSurfaceArea(surfaceArea_mm2, new List<ObjectInfo>());
        }

        private void SelectObjectsButton_Click(object sender, EventArgs e)
        {
            try
            {
                // Get the active Rhino document
                var doc = RhinoDoc.ActiveDoc;
                if (doc == null)
                {
                    MessageBox.Show("No active Rhino document found.", MessageBoxType.Error);
                    return;
                }

                // Run the CoatingApp command to select objects
                RhinoApp.RunScript("_CoatingApp", false);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error selecting objects: {ex.Message}", MessageBoxType.Error);
            }
        }

        private void ClearObjectsButton_Click(object sender, EventArgs e)
        {
            try
            {
                // Reset surface area
                UpdateSurfaceArea(0);
                
                // Clear results
                resultsTextArea.Text = "Objects cleared. Select new objects to calculate.";
                exportButton.Enabled = false;
                
                RhinoApp.WriteLine("CoatingApp: Objects cleared.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error clearing objects: {ex.Message}", MessageBoxType.Error);
            }
        }

        private void OnPrimerChanged(object sender, EventArgs e)
        {
            string selectedMaterial = primerDropDown.SelectedKey;
            if (materials.ContainsKey(selectedMaterial) && materials[selectedMaterial] != null)
            {
                var material = materials[selectedMaterial];
                primerConsumptionStepper.Value = material.ConsumptionPerSqM;
                primerPriceStepper.Value = material.PricePerKg;
            }
            else
            {
                // "None" selected
                primerConsumptionStepper.Value = 0;
                primerPriceStepper.Value = 0;
            }
        }

        private void OnTopcoatChanged(object sender, EventArgs e)
        {
            string selectedMaterial = topcoatDropDown.SelectedKey;
            if (materials.ContainsKey(selectedMaterial) && materials[selectedMaterial] != null)
            {
                var material = materials[selectedMaterial];
                topcoatConsumptionStepper.Value = material.ConsumptionPerSqM;
                topcoatPriceStepper.Value = material.PricePerKg;
            }
            else
            {
                // "None" selected
                topcoatConsumptionStepper.Value = 0;
                topcoatPriceStepper.Value = 0;
            }
        }

        private void InitializeMaterials()
        {
            materials = new Dictionary<string, MaterialInfo>();
            materials.Add("None", null);

            // Load materials from configuration
            var config = CoatingConfigManager.Config;
            if (config?.Materials != null)
            {
                foreach (var material in config.Materials)
                {
                    materials.Add(material.Name, new MaterialInfo
                    {
                        Name = material.Name,
                        ConsumptionPerSqM = material.ConsumptionPerSqM,
                        PricePerKg = material.PricePerKg
                    });
                }
            }
            else
            {
                // Fallback to hardcoded values if config not available
                materials.Add("Standard Primer", new MaterialInfo { Name = "Standard Primer", ConsumptionPerSqM = 200.0, PricePerKg = 25.0 });
                materials.Add("Premium Primer", new MaterialInfo { Name = "Premium Primer", ConsumptionPerSqM = 200.0, PricePerKg = 35.0 });
                materials.Add("Basic Topcoat", new MaterialInfo { Name = "Basic Topcoat", ConsumptionPerSqM = 200.0, PricePerKg = 20.0 });
                materials.Add("Premium Topcoat", new MaterialInfo { Name = "Premium Topcoat", ConsumptionPerSqM = 200.0, PricePerKg = 40.0 });
            }
        }

        private void InitializeUI()
        {
            // Create main splitter for horizontal division
            var mainSplitter = new Splitter
            {
                Orientation = Orientation.Horizontal,
                Panel1MinimumSize = 300,
                Panel2MinimumSize = 200,
                Position = 600,
                Panel1 = CreateMaterialConfigurationPanel(),
                Panel2 = CreateCalculationFactorsPanel()
            };

            // Set content for the panel
            Content = mainSplitter;
        }

        private Scrollable CreateMaterialConfigurationPanel()
        {
            var scrollable = new Scrollable();
            var layout = new DynamicLayout();
            layout.Spacing = new Size(5, 5);
            layout.Padding = 10;

            // Object selection section
            layout.BeginVertical();
            layout.AddRow(new Label { Text = "Object Selection:", Font = SystemFonts.Bold() });
            
            var buttonLayout = new DynamicLayout();
            buttonLayout.BeginHorizontal();
            selectObjectsButton = new Button { Text = "Select Objects" };
            selectObjectsButton.Click += SelectObjectsButton_Click;
            buttonLayout.AddRow(selectObjectsButton);
            
            clearObjectsButton = new Button { Text = "Clear Objects" };
            clearObjectsButton.Click += ClearObjectsButton_Click;
            buttonLayout.AddRow(clearObjectsButton);
            buttonLayout.EndHorizontal();
            
            layout.AddRow(buttonLayout);
            
            surfaceAreaLabel = new Label 
            { 
                Text = surfaceArea_mm2 > 0 
                    ? $"{surfaceArea_mm2:F2} mm² ({surfaceArea_mm2 / 1000000:F4} m²)"
                    : "No objects selected",
                Font = new Font(SystemFont.Default, 10)
            };
            layout.AddRow(surfaceAreaLabel);
            layout.EndVertical();

            layout.AddSeparateRow(null);

            // Material configuration section
            layout.BeginVertical();
            layout.AddRow(new Label { Text = "Material Configuration:", Font = SystemFonts.Bold() });
            
            // Primer selection
            layout.BeginHorizontal();
            layout.AddRow(new Label { Text = "Primer:", Width = 120 });
            primerDropDown = new DropDown();
            foreach (var material in materials.Keys)
            {
                primerDropDown.Items.Add(material);
            }
            primerDropDown.SelectedIndex = 1; // Standard Primer (index 1)
            primerDropDown.SelectedIndexChanged += OnPrimerChanged;
            layout.AddRow(primerDropDown);
            layout.EndHorizontal();

            // Primer consumption
            layout.BeginHorizontal();
            layout.AddRow(new Label { Text = "Consumption (g/m²):", Width = 120 });
            primerConsumptionStepper = new NumericStepper
            {
                MinValue = 0,
                MaxValue = 1000,
                Value = 200, // Default 200 g/m²
                DecimalPlaces = 1,
                Increment = 10
            };
            layout.AddRow(primerConsumptionStepper);
            layout.EndHorizontal();

            // Primer price
            layout.BeginHorizontal();
            layout.AddRow(new Label { Text = "Price (Fr./kg):", Width = 120 });
            primerPriceStepper = new NumericStepper
            {
                MinValue = 0,
                MaxValue = 1000,
                Value = 25,
                DecimalPlaces = 2,
                Increment = 1
            };
            layout.AddRow(primerPriceStepper);
            layout.EndHorizontal();

            layout.AddSeparateRow(null);

            // Topcoat selection
            layout.BeginHorizontal();
            layout.AddRow(new Label { Text = "Topcoat:", Width = 120 });
            topcoatDropDown = new DropDown();
            foreach (var material in materials.Keys)
            {
                topcoatDropDown.Items.Add(material);
            }
            topcoatDropDown.SelectedIndex = 3; // Basic Topcoat (index 3)
            topcoatDropDown.SelectedIndexChanged += OnTopcoatChanged;
            layout.AddRow(topcoatDropDown);
            layout.EndHorizontal();

            // Topcoat consumption
            layout.BeginHorizontal();
            layout.AddRow(new Label { Text = "Consumption (g/m²):", Width = 120 });
            topcoatConsumptionStepper = new NumericStepper
            {
                MinValue = 0,
                MaxValue = 1000,
                Value = 200, // Default 200 g/m²
                DecimalPlaces = 1,
                Increment = 10
            };
            layout.AddRow(topcoatConsumptionStepper);
            layout.EndHorizontal();

            // Topcoat price
            layout.BeginHorizontal();
            layout.AddRow(new Label { Text = "Price (Fr./kg):", Width = 120 });
            topcoatPriceStepper = new NumericStepper
            {
                MinValue = 0,
                MaxValue = 1000,
                Value = 20,
                DecimalPlaces = 2,
                Increment = 1
            };
            layout.AddRow(topcoatPriceStepper);
            layout.EndHorizontal();

            layout.AddSeparateRow(null);

            // Time factor
            layout.BeginHorizontal();
            layout.AddRow(new Label { Text = "Time Factor (h/m²):", Width = 120 });
            timeFactorStepper = new NumericStepper
            {
                MinValue = 0.01,
                MaxValue = 10.0,
                Value = 0.5,
                DecimalPlaces = 2,
                Increment = 0.1
            };
            layout.AddRow(timeFactorStepper);
            layout.EndHorizontal();

            layout.AddSeparateRow(null);

            // Load config defaults for multipliers
            var defaultSettings = CoatingConfigManager.Config?.DefaultSettings;

            // Primer coat multiplier
            layout.BeginHorizontal();
            layout.AddRow(new Label { Text = "Primer Coats:", Width = 120 });
            primerCoatMultiplierStepper = new NumericStepper
            {
                MinValue = 1,
                MaxValue = 10,
                Value = defaultSettings?.DefaultPrimerCoatMultiplier ?? 1.0,
                DecimalPlaces = 0,
                Increment = 1
            };
            layout.AddRow(primerCoatMultiplierStepper);
            layout.EndHorizontal();

            // Topcoat coat multiplier
            layout.BeginHorizontal();
            layout.AddRow(new Label { Text = "Topcoat Coats:", Width = 120 });
            topcoatCoatMultiplierStepper = new NumericStepper
            {
                MinValue = 1,
                MaxValue = 10,
                Value = defaultSettings?.DefaultTopcoatCoatMultiplier ?? 1.0,
                DecimalPlaces = 0,
                Increment = 1
            };
            layout.AddRow(topcoatCoatMultiplierStepper);
            layout.EndHorizontal();
            layout.EndVertical();

            layout.AddSeparateRow(null);

            // Additive Prices
            layout.BeginVertical();
            layout.AddRow(new Label { Text = "Additive Prices:", Font = SystemFonts.Bold() });
            
            // Load config defaults
            var config = CoatingConfigManager.Config;
            
            // Hardener Price
            layout.BeginHorizontal();
            layout.AddRow(new Label { Text = "Hardener (Fr./kg):", Width = 120 });
            hardenerPriceStepper = new NumericStepper
            {
                MinValue = 0,
                MaxValue = 1000,
                Value = config?.MaterialPrices?.HardenerPricePerKg ?? 23,
                DecimalPlaces = 2,
                Increment = 1
            };
            layout.AddRow(hardenerPriceStepper);
            layout.EndHorizontal();

            // Thinner Price
            layout.BeginHorizontal();
            layout.AddRow(new Label { Text = "Thinner (Fr./kg):", Width = 120 });
            thinnerPriceStepper = new NumericStepper
            {
                MinValue = 0,
                MaxValue = 1000,
                Value = config?.MaterialPrices?.ThinnerPricePerKg ?? 9,
                DecimalPlaces = 2,
                Increment = 1
            };
            layout.AddRow(thinnerPriceStepper);
            layout.EndHorizontal();
            layout.EndVertical();

            layout.AddSeparateRow(null);

            // Application Type Selection
            layout.BeginVertical();
            layout.AddRow(new Label { Text = "Application Type:", Font = SystemFonts.Bold() });
            applicationTypeRadio = new RadioButtonList
            {
                Orientation = Orientation.Horizontal,
                Spacing = new Size(10, 5)
            };
            applicationTypeRadio.Items.Add(new ListItem { Text = "Indoor", Key = "Indoor" });
            applicationTypeRadio.Items.Add(new ListItem { Text = "Outdoor", Key = "Outdoor" });
            applicationTypeRadio.SelectedIndex = 0; // Default to Indoor
            applicationTypeRadio.SelectedIndexChanged += OnApplicationTypeChanged;
            layout.AddRow(applicationTypeRadio);
            layout.EndVertical();

            layout.AddSeparateRow(null);

            // Indoor Additives Panel
            indoorAdditivePanel = CreateIndoorAdditivePanel();
            layout.AddRow(indoorAdditivePanel);

            // Outdoor Additives Panel
            outdoorAdditivePanel = CreateOutdoorAdditivePanel();
            outdoorAdditivePanel.Visible = false; // Hidden by default
            layout.AddRow(outdoorAdditivePanel);

            layout.AddSeparateRow(null);

            // Calculate button
            calculateButton = new Button { Text = "Calculate" };
            calculateButton.Click += CalculateButton_Click;
            layout.AddRow(calculateButton);

            layout.AddSeparateRow(null);

            // Results section
            layout.BeginVertical();
            layout.AddRow(new Label { Text = "Results:", Font = SystemFonts.Bold() });
            resultsTextArea = new TextArea
            {
                ReadOnly = true,
                Height = 150,
                Text = "Select objects and click 'Calculate' to see results..."
            };
            layout.AddRow(resultsTextArea);
            layout.EndVertical();

            layout.AddSeparateRow(null);

            // Export button
            exportButton = new Button { Text = "Export Results" };
            exportButton.Click += ExportButton_Click;
            exportButton.Enabled = false;
            layout.AddRow(exportButton);

            // Set content for the scrollable panel
            scrollable.Content = layout;
            return scrollable;
        }

        private Scrollable CreateCalculationFactorsPanel()
        {
            var scrollable = new Scrollable();
            var layout = new DynamicLayout();
            layout.Spacing = new Size(5, 5);
            layout.Padding = 10;

            // Initialize calculation factor steppers dictionary
            calculationFactorSteppers = new Dictionary<string, NumericStepper>();

            // Calculation factors section
            layout.BeginVertical();
            layout.AddRow(new Label { Text = "Calculation Factors:", Font = SystemFonts.Bold() });

            // Load calculation factors from config
            var calcConfig = CalculationConfigManager.Config;
            if (calcConfig?.CalculationFactors != null)
            {
                foreach (var factor in calcConfig.CalculationFactors)
                {
                    // Factor name and description
                    layout.BeginHorizontal();
                    layout.AddRow(new Label { Text = $"{factor.Value.Name}:", Width = 180 });
                    layout.AddRow(new Label
                    {
                        Text = factor.Value.Description,
                        Font = new Font(SystemFont.Default, 8),
                        TextColor = Colors.Gray,
                        Wrap = WrapMode.Word
                    });
                    layout.EndHorizontal();

                    // Factor percentage input
                    layout.BeginHorizontal();
                    layout.AddRow(new Label { Text = "Percentage:", Width = 180 });
                    var stepper = new NumericStepper
                    {
                        MinValue = 0,
                        MaxValue = 1000,
                        Value = factor.Value.Percentage,
                        DecimalPlaces = 1,
                        Increment = 1,
                        ToolTip = factor.Value.Description
                    };
                    stepper.ValueChanged += (sender, e) => OnCalculationFactorChanged(factor.Key, stepper.Value);
                    calculationFactorSteppers.Add(factor.Key, stepper);
                    layout.AddRow(stepper);
                    layout.AddRow(new Label { Text = "%" });
                    layout.EndHorizontal();

                    layout.AddSeparateRow(null);
                }
            }

            layout.AddSeparateRow(null);
            layout.AddRow(new Label { Text = "Adjust percentages to calculate final offer price", Font = new Font(SystemFont.Default, 9) });
            layout.EndVertical();

            scrollable.Content = layout;
            return scrollable;
        }

        private StackLayout CreateIndoorAdditivePanel()
        {
            var panel = new StackLayout
            {
                Orientation = Orientation.Vertical,
                Spacing = 5,
                Padding = 10
            };

            var titleLabel = new Label { Text = "Indoor Application - Additives:", Font = SystemFonts.Bold() };
            panel.Items.Add(titleLabel);

            // Load config defaults
            var config = CoatingConfigManager.Config;
            var indoorDefaults = config?.ApplicationDefaults?.Indoor;

            // Primer Hardener
            var primerHardenerLayout = new DynamicLayout();
            primerHardenerLayout.BeginHorizontal();
            primerHardenerLayout.AddRow(new Label { Text = "Primer Hardener (%):", Width = 150 });
            indoorPrimerHardenerStepper = new NumericStepper
            {
                MinValue = 0,
                MaxValue = 100,
                Value = indoorDefaults?.PrimerHardenerPercent ?? 6.0,
                DecimalPlaces = 1,
                Increment = 0.5
            };
            primerHardenerLayout.AddRow(indoorPrimerHardenerStepper);
            primerHardenerLayout.EndHorizontal();
            panel.Items.Add(primerHardenerLayout);

            // Primer Thinner
            var primerThinnerLayout = new DynamicLayout();
            primerThinnerLayout.BeginHorizontal();
            primerThinnerLayout.AddRow(new Label { Text = "Primer Thinner (%):", Width = 150 });
            indoorPrimerThinnerStepper = new NumericStepper
            {
                MinValue = 0,
                MaxValue = 100,
                Value = indoorDefaults?.PrimerThinnerPercent ?? 10.0,
                DecimalPlaces = 1,
                Increment = 0.5
            };
            primerThinnerLayout.AddRow(indoorPrimerThinnerStepper);
            primerThinnerLayout.EndHorizontal();
            panel.Items.Add(primerThinnerLayout);

            // Topcoat Hardener
            var topcoatHardenerLayout = new DynamicLayout();
            topcoatHardenerLayout.BeginHorizontal();
            topcoatHardenerLayout.AddRow(new Label { Text = "Topcoat Hardener (%):", Width = 150 });
            indoorTopcoatHardenerStepper = new NumericStepper
            {
                MinValue = 0,
                MaxValue = 100,
                Value = indoorDefaults?.TopcoatHardenerPercent ?? 15.0,
                DecimalPlaces = 1,
                Increment = 0.5
            };
            topcoatHardenerLayout.AddRow(indoorTopcoatHardenerStepper);
            topcoatHardenerLayout.EndHorizontal();
            panel.Items.Add(topcoatHardenerLayout);

            // Topcoat Thinner
            var topcoatThinnerLayout = new DynamicLayout();
            topcoatThinnerLayout.BeginHorizontal();
            topcoatThinnerLayout.AddRow(new Label { Text = "Topcoat Thinner (%):", Width = 150 });
            indoorTopcoatThinnerStepper = new NumericStepper
            {
                MinValue = 0,
                MaxValue = 100,
                Value = indoorDefaults?.TopcoatThinnerPercent ?? 20.0,
                DecimalPlaces = 1,
                Increment = 0.5
            };
            topcoatThinnerLayout.AddRow(indoorTopcoatThinnerStepper);
            topcoatThinnerLayout.EndHorizontal();
            panel.Items.Add(topcoatThinnerLayout);

            return panel;
        }

        private StackLayout CreateOutdoorAdditivePanel()
        {
            var panel = new StackLayout
            {
                Orientation = Orientation.Vertical,
                Spacing = 5,
                Padding = 10
            };

            var titleLabel = new Label { Text = "Outdoor Application - Additives:", Font = SystemFonts.Bold() };
            panel.Items.Add(titleLabel);

            // Load config defaults
            var config = CoatingConfigManager.Config;
            var outdoorDefaults = config?.ApplicationDefaults?.Outdoor;

            // Primer Hardener
            var primerHardenerLayout = new DynamicLayout();
            primerHardenerLayout.BeginHorizontal();
            primerHardenerLayout.AddRow(new Label { Text = "Primer Hardener (%):", Width = 150 });
            outdoorPrimerHardenerStepper = new NumericStepper
            {
                MinValue = 0,
                MaxValue = 100,
                Value = outdoorDefaults?.PrimerHardenerPercent ?? 8.0,
                DecimalPlaces = 1,
                Increment = 0.5
            };
            primerHardenerLayout.AddRow(outdoorPrimerHardenerStepper);
            primerHardenerLayout.EndHorizontal();
            panel.Items.Add(primerHardenerLayout);

            // Primer Thinner
            var primerThinnerLayout = new DynamicLayout();
            primerThinnerLayout.BeginHorizontal();
            primerThinnerLayout.AddRow(new Label { Text = "Primer Thinner (%):", Width = 150 });
            outdoorPrimerThinnerStepper = new NumericStepper
            {
                MinValue = 0,
                MaxValue = 100,
                Value = outdoorDefaults?.PrimerThinnerPercent ?? 12.0,
                DecimalPlaces = 1,
                Increment = 0.5
            };
            primerThinnerLayout.AddRow(outdoorPrimerThinnerStepper);
            primerThinnerLayout.EndHorizontal();
            panel.Items.Add(primerThinnerLayout);

            // Topcoat Hardener
            var topcoatHardenerLayout = new DynamicLayout();
            topcoatHardenerLayout.BeginHorizontal();
            topcoatHardenerLayout.AddRow(new Label { Text = "Topcoat Hardener (%):", Width = 150 });
            outdoorTopcoatHardenerStepper = new NumericStepper
            {
                MinValue = 0,
                MaxValue = 100,
                Value = outdoorDefaults?.TopcoatHardenerPercent ?? 18.0,
                DecimalPlaces = 1,
                Increment = 0.5
            };
            topcoatHardenerLayout.AddRow(outdoorTopcoatHardenerStepper);
            topcoatHardenerLayout.EndHorizontal();
            panel.Items.Add(topcoatHardenerLayout);

            // Topcoat Thinner
            var topcoatThinnerLayout = new DynamicLayout();
            topcoatThinnerLayout.BeginHorizontal();
            topcoatThinnerLayout.AddRow(new Label { Text = "Topcoat Thinner (%):", Width = 150 });
            outdoorTopcoatThinnerStepper = new NumericStepper
            {
                MinValue = 0,
                MaxValue = 100,
                Value = outdoorDefaults?.TopcoatThinnerPercent ?? 25.0,
                DecimalPlaces = 1,
                Increment = 0.5
            };
            topcoatThinnerLayout.AddRow(outdoorTopcoatThinnerStepper);
            topcoatThinnerLayout.EndHorizontal();
            panel.Items.Add(topcoatThinnerLayout);

            return panel;
        }

        private void OnApplicationTypeChanged(object sender, EventArgs e)
        {
            if (applicationTypeRadio.SelectedKey == "Indoor")
            {
                indoorAdditivePanel.Visible = true;
                outdoorAdditivePanel.Visible = false;
            }
            else
            {
                indoorAdditivePanel.Visible = false;
                outdoorAdditivePanel.Visible = true;
            }
        }

        private void CalculateButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (surfaceArea_mm2 <= 0)
                {
                    MessageBox.Show("Please select objects first using the 'Select Objects' button.", MessageBoxType.Warning);
                    return;
                }

                // Get material configuration from UI fields
                string primerName = primerDropDown.SelectedKey;
                string topcoatName = topcoatDropDown.SelectedKey;
                
                // Determine application type and get appropriate additive percentages
                bool isIndoor = applicationTypeRadio.SelectedKey == "Indoor";
                ApplicationType appType = isIndoor ? ApplicationType.Indoor : ApplicationType.Outdoor;
                
                double primerHardenerPercent = isIndoor ? indoorPrimerHardenerStepper.Value : outdoorPrimerHardenerStepper.Value;
                double primerThinnerPercent = isIndoor ? indoorPrimerThinnerStepper.Value : outdoorPrimerThinnerStepper.Value;
                double topcoatHardenerPercent = isIndoor ? indoorTopcoatHardenerStepper.Value : outdoorTopcoatHardenerStepper.Value;
                double topcoatThinnerPercent = isIndoor ? indoorTopcoatThinnerStepper.Value : outdoorTopcoatThinnerStepper.Value;

                MaterialConfig = new MaterialConfig
                {
                    Primer = primerName != "None" ? new MaterialInfo
                    {
                        Name = primerName,
                        ConsumptionPerSqM = primerConsumptionStepper.Value,
                        PricePerKg = primerPriceStepper.Value
                    } : null,
                    Topcoat = topcoatName != "None" ? new MaterialInfo
                    {
                        Name = topcoatName,
                        ConsumptionPerSqM = topcoatConsumptionStepper.Value,
                        PricePerKg = topcoatPriceStepper.Value
                    } : null,
                    TimeFactor = timeFactorStepper.Value,
                    ApplicationType = appType,
                    PrimerHardenerPercent = primerHardenerPercent,
                    PrimerThinnerPercent = primerThinnerPercent,
                    TopcoatHardenerPercent = topcoatHardenerPercent,
                    TopcoatThinnerPercent = topcoatThinnerPercent,
                    HardenerPricePerKg = hardenerPriceStepper.Value,
                    ThinnerPricePerKg = thinnerPriceStepper.Value,
                    PrimerCoatMultiplier = primerCoatMultiplierStepper.Value,
                    TopcoatCoatMultiplier = topcoatCoatMultiplierStepper.Value
                };

                // Calculate costs
                Costs = CalculateCosts(surfaceArea_mm2, MaterialConfig);

                // Calculate time
                EstimatedTime = EstimateTime(surfaceArea_mm2, MaterialConfig.TimeFactor);

                // Calculate final offer with current factors
                var currentFactors = GetCurrentCalculationFactors();
                CalculationResult = CalculateFinalOffer(Costs, currentFactors);

                // Display results with calculation
                DisplayResultsWithCalculation();

                // Enable export button
                exportButton.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during calculation: {ex.Message}", MessageBoxType.Error);
            }
        }

        private CostCalculation CalculateCosts(double surfaceArea_mm2, MaterialConfig config)
        {
            CostCalculation costs = new CostCalculation();
            double surfaceArea_m2 = surfaceArea_mm2 / 1000000.0;

            if (config.Primer != null)
            {
                // Base primer calculation with coat multiplier
                double primerGrams = surfaceArea_m2 * config.Primer.ConsumptionPerSqM * config.PrimerCoatMultiplier;
                double primerKg = primerGrams / 1000.0;
                costs.PrimerCost = primerKg * config.Primer.PricePerKg;
                
                // Hardener for primer (percentage of primer amount)
                double primerHardenerGrams = primerGrams * (config.PrimerHardenerPercent / 100.0);
                double primerHardenerKg = primerHardenerGrams / 1000.0;
                costs.PrimerHardenerCost = primerHardenerKg * config.HardenerPricePerKg;
                
                // Thinner for primer (percentage of primer amount)
                double primerThinnerGrams = primerGrams * (config.PrimerThinnerPercent / 100.0);
                double primerThinnerKg = primerThinnerGrams / 1000.0;
                costs.PrimerThinnerCost = primerThinnerKg * config.ThinnerPricePerKg;
                
                costs.TotalMaterialCost += costs.PrimerCost + costs.PrimerHardenerCost + costs.PrimerThinnerCost;
            }

            if (config.Topcoat != null)
            {
                // Base topcoat calculation with coat multiplier
                double topcoatGrams = surfaceArea_m2 * config.Topcoat.ConsumptionPerSqM * config.TopcoatCoatMultiplier;
                double topcoatKg = topcoatGrams / 1000.0;
                costs.TopcoatCost = topcoatKg * config.Topcoat.PricePerKg;
                
                // Hardener for topcoat (percentage of topcoat amount)
                double topcoatHardenerGrams = topcoatGrams * (config.TopcoatHardenerPercent / 100.0);
                double topcoatHardenerKg = topcoatHardenerGrams / 1000.0;
                costs.TopcoatHardenerCost = topcoatHardenerKg * config.HardenerPricePerKg;
                
                // Thinner for topcoat (percentage of topcoat amount)
                double topcoatThinnerGrams = topcoatGrams * (config.TopcoatThinnerPercent / 100.0);
                double topcoatThinnerKg = topcoatThinnerGrams / 1000.0;
                costs.TopcoatThinnerCost = topcoatThinnerKg * config.ThinnerPricePerKg;
                
                costs.TotalMaterialCost += costs.TopcoatCost + costs.TopcoatHardenerCost + costs.TopcoatThinnerCost;
            }

            return costs;
        }

        private double EstimateTime(double surfaceArea_mm2, double timeFactor)
        {
            double surfaceArea_m2 = surfaceArea_mm2 / 1000000.0;
            return surfaceArea_m2 * timeFactor;
        }

        private void DisplayResults()
        {
            var results = new System.Text.StringBuilder();
            results.AppendLine("=== COATING CALCULATION RESULTS ===");
            results.AppendLine();
            results.AppendLine($"Application Type: {MaterialConfig.ApplicationType}");
            if (MaterialConfig.Primer != null || MaterialConfig.Topcoat != null)
            {
                results.AppendLine($"Coat Multipliers: Primer={MaterialConfig.PrimerCoatMultiplier:F0}x, Topcoat={MaterialConfig.TopcoatCoatMultiplier:F0}x");
            }
            results.AppendLine();

            // Display individual objects
            if (selectedObjects != null && selectedObjects.Count > 0)
            {
                results.AppendLine("INDIVIDUAL OBJECTS:");
                results.AppendLine();
                
                for (int i = 0; i < selectedObjects.Count; i++)
                {
                    var obj = selectedObjects[i];
                    results.AppendLine($"Object {i + 1}: {obj.Name}");
                    results.AppendLine($"  Surface Area: {obj.SurfaceArea_mm2:F2} mm² ({obj.SurfaceArea_mm2 / 1000000:F4} m²)");
                    
                    // Calculate costs and material amounts for this individual object
                    double objArea_m2 = obj.SurfaceArea_mm2 / 1000000.0;
                    var objCosts = CalculateCosts(obj.SurfaceArea_mm2, MaterialConfig);
                    var objTime = EstimateTime(obj.SurfaceArea_mm2, MaterialConfig.TimeFactor);
                    
                    if (MaterialConfig.Primer != null)
                    {
                        double primerGrams = objArea_m2 * MaterialConfig.Primer.ConsumptionPerSqM * MaterialConfig.PrimerCoatMultiplier;
                        double primerKg = primerGrams / 1000.0;
                        double primerHardenerGrams = primerGrams * (MaterialConfig.PrimerHardenerPercent / 100.0);
                        double primerHardenerKg = primerHardenerGrams / 1000.0;
                        double primerThinnerGrams = primerGrams * (MaterialConfig.PrimerThinnerPercent / 100.0);
                        double primerThinnerKg = primerThinnerGrams / 1000.0;

                        results.AppendLine($"  Primer ({MaterialConfig.PrimerCoatMultiplier:F0}x coats): {primerGrams:F1} g ({primerKg:F3} kg) - Fr. {objCosts.PrimerCost:F2}");
                        results.AppendLine($"    + Hardener ({MaterialConfig.PrimerHardenerPercent:F1}%): {primerHardenerGrams:F1} g ({primerHardenerKg:F3} kg) - Fr. {objCosts.PrimerHardenerCost:F2}");
                        results.AppendLine($"    + Thinner ({MaterialConfig.PrimerThinnerPercent:F1}%): {primerThinnerGrams:F1} g ({primerThinnerKg:F3} kg) - Fr. {objCosts.PrimerThinnerCost:F2}");
                    }
                    if (MaterialConfig.Topcoat != null)
                    {
                        double topcoatGrams = objArea_m2 * MaterialConfig.Topcoat.ConsumptionPerSqM * MaterialConfig.TopcoatCoatMultiplier;
                        double topcoatKg = topcoatGrams / 1000.0;
                        double topcoatHardenerGrams = topcoatGrams * (MaterialConfig.TopcoatHardenerPercent / 100.0);
                        double topcoatHardenerKg = topcoatHardenerGrams / 1000.0;
                        double topcoatThinnerGrams = topcoatGrams * (MaterialConfig.TopcoatThinnerPercent / 100.0);
                        double topcoatThinnerKg = topcoatThinnerGrams / 1000.0;

                        results.AppendLine($"  Topcoat ({MaterialConfig.TopcoatCoatMultiplier:F0}x coats): {topcoatGrams:F1} g ({topcoatKg:F3} kg) - Fr. {objCosts.TopcoatCost:F2}");
                        results.AppendLine($"    + Hardener ({MaterialConfig.TopcoatHardenerPercent:F1}%): {topcoatHardenerGrams:F1} g ({topcoatHardenerKg:F3} kg) - Fr. {objCosts.TopcoatHardenerCost:F2}");
                        results.AppendLine($"    + Thinner ({MaterialConfig.TopcoatThinnerPercent:F1}%): {topcoatThinnerGrams:F1} g ({topcoatThinnerKg:F3} kg) - Fr. {objCosts.TopcoatThinnerCost:F2}");
                    }
                    results.AppendLine($"  Total Material Cost: Fr. {objCosts.TotalMaterialCost:F2}");
                    results.AppendLine($"  Estimated Time: {objTime:F2} hours");
                    results.AppendLine();
                }

                results.AppendLine("---------------------------------");
                results.AppendLine();
            }

            // Display summary
            results.AppendLine("SUMMARY:");
            results.AppendLine();
            results.AppendLine($"Total Objects: {selectedObjects.Count}");
            results.AppendLine($"Total Surface Area: {surfaceArea_mm2:F2} mm²");
            results.AppendLine($"                    ({surfaceArea_mm2 / 1000000:F4} m�)");
            results.AppendLine();

            double totalArea_m2 = surfaceArea_mm2 / 1000000.0;
            
            if (MaterialConfig.Primer != null)
            {
                double totalPrimerGrams = totalArea_m2 * MaterialConfig.Primer.ConsumptionPerSqM * MaterialConfig.PrimerCoatMultiplier;
                double totalPrimerKg = totalPrimerGrams / 1000.0;
                double totalPrimerHardenerGrams = totalPrimerGrams * (MaterialConfig.PrimerHardenerPercent / 100.0);
                double totalPrimerHardenerKg = totalPrimerHardenerGrams / 1000.0;
                double totalPrimerThinnerGrams = totalPrimerGrams * (MaterialConfig.PrimerThinnerPercent / 100.0);
                double totalPrimerThinnerKg = totalPrimerThinnerGrams / 1000.0;

                results.AppendLine($"Primer: {MaterialConfig.Primer.Name}");
                results.AppendLine($"  Consumption: {MaterialConfig.Primer.ConsumptionPerSqM:F1} g/m²");
                results.AppendLine($"  Coats: {MaterialConfig.PrimerCoatMultiplier:F0}x");
                results.AppendLine($"  Total Amount: {totalPrimerGrams:F1} g ({totalPrimerKg:F3} kg)");
                results.AppendLine($"  Price: Fr. {MaterialConfig.Primer.PricePerKg:F2}/kg");
                results.AppendLine($"  Total Cost: Fr. {Costs.PrimerCost:F2}");
                results.AppendLine($"  Hardener ({MaterialConfig.PrimerHardenerPercent:F1}%): {totalPrimerHardenerGrams:F1} g ({totalPrimerHardenerKg:F3} kg) - Fr. {Costs.PrimerHardenerCost:F2}");
                results.AppendLine($"  Thinner ({MaterialConfig.PrimerThinnerPercent:F1}%): {totalPrimerThinnerGrams:F1} g ({totalPrimerThinnerKg:F3} kg) - Fr. {Costs.PrimerThinnerCost:F2}");
                results.AppendLine();
            }

            if (MaterialConfig.Topcoat != null)
            {
                double totalTopcoatGrams = totalArea_m2 * MaterialConfig.Topcoat.ConsumptionPerSqM * MaterialConfig.TopcoatCoatMultiplier;
                double totalTopcoatKg = totalTopcoatGrams / 1000.0;
                double totalTopcoatHardenerGrams = totalTopcoatGrams * (MaterialConfig.TopcoatHardenerPercent / 100.0);
                double totalTopcoatHardenerKg = totalTopcoatHardenerGrams / 1000.0;
                double totalTopcoatThinnerGrams = totalTopcoatGrams * (MaterialConfig.TopcoatThinnerPercent / 100.0);
                double totalTopcoatThinnerKg = totalTopcoatThinnerGrams / 1000.0;

                results.AppendLine($"Topcoat: {MaterialConfig.Topcoat.Name}");
                results.AppendLine($"  Consumption: {MaterialConfig.Topcoat.ConsumptionPerSqM:F1} g/m²");
                results.AppendLine($"  Coats: {MaterialConfig.TopcoatCoatMultiplier:F0}x");
                results.AppendLine($"  Total Amount: {totalTopcoatGrams:F1} g ({totalTopcoatKg:F3} kg)");
                results.AppendLine($"  Price: Fr. {MaterialConfig.Topcoat.PricePerKg:F2}/kg");
                results.AppendLine($"  Total Cost: Fr. {Costs.TopcoatCost:F2}");
                results.AppendLine($"  Hardener ({MaterialConfig.TopcoatHardenerPercent:F1}%): {totalTopcoatHardenerGrams:F1} g ({totalTopcoatHardenerKg:F3} kg) - Fr. {Costs.TopcoatHardenerCost:F2}");
                results.AppendLine($"  Thinner ({MaterialConfig.TopcoatThinnerPercent:F1}%): {totalTopcoatThinnerGrams:F1} g ({totalTopcoatThinnerKg:F3} kg) - Fr. {Costs.TopcoatThinnerCost:F2}");
                results.AppendLine();
            }

            results.AppendLine($"Total Material Cost: Fr. {Costs.TotalMaterialCost:F2}");
            results.AppendLine($"Total Estimated Time: {EstimatedTime:F2} hours");
            results.AppendLine($"Time Factor: {MaterialConfig.TimeFactor:F2} h/m²");

            resultsTextArea.Text = results.ToString();
        }

        private void ExportButton_Click(object sender, EventArgs e)
        {
            try
            {
                var saveDialog = new Eto.Forms.SaveFileDialog
                {
                    Title = "Export Coating Calculation",
                    Filters = 
                    {
                        new FileFilter("Excel Files", ".xlsx"),
                        new FileFilter("CSV Files", ".csv"),
                        new FileFilter("Text Files", ".txt"),
                        new FileFilter("All Files", "*.*")
                    }
                };

                if (saveDialog.ShowDialog(null) == DialogResult.Ok)
                {
                    // Use the CoatingExporter class for all export functionality
                    CoatingExporter.Export(
                        saveDialog.FileName,
                        surfaceArea_mm2,
                        selectedObjects,
                        MaterialConfig,
                        Costs,
                        EstimatedTime,
                        CalculationResult,
                        calculationFactorSteppers?.ToDictionary(kvp => kvp.Key, kvp => kvp.Value.Value)
                    );

                    MessageBox.Show($"Results exported successfully to:\n{saveDialog.FileName}", 
                                    MessageBoxType.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting results: {ex.Message}", MessageBoxType.Error);
            }
        }

        // Export methods moved to CoatingExporter.cs for better maintainability

        private void OnCalculationFactorChanged(string factorKey, double newValue)
        {
            // Update the calculation if we have material config and costs
            if (MaterialConfig != null && Costs != null)
            {
                // Recalculate with new factor values
                CalculationResult = CalculateFinalOffer(Costs, GetCurrentCalculationFactors());

                // Update results display if available
                if (resultsTextArea != null)
                {
                    DisplayResultsWithCalculation();
                }
            }
        }

        private Dictionary<string, double> GetCurrentCalculationFactors()
        {
            var factors = new Dictionary<string, double>();
            if (calculationFactorSteppers != null)
            {
                foreach (var kvp in calculationFactorSteppers)
                {
                    factors.Add(kvp.Key, kvp.Value.Value);
                }
            }
            return factors;
        }

        private CalculationResult CalculateFinalOffer(CostCalculation baseCosts, Dictionary<string, double> factors)
        {
            double baseMaterialCost = baseCosts.TotalMaterialCost;

            // Start with material costs
            double currentCost = baseMaterialCost;

            var result = new CalculationResult
            {
                BaseMaterialCost = baseMaterialCost,
                MaterialCostWithSurcharge = currentCost
            };

            // Apply each factor
            if (factors.ContainsKey("MaterialCostFactor"))
            {
                double materialFactor = factors["MaterialCostFactor"] / 100.0;
                double surcharge = currentCost * materialFactor;
                result.MaterialSurcharge = surcharge;
                currentCost += surcharge;
            }

            if (factors.ContainsKey("ProductionCostFactor"))
            {
                double productionFactor = factors["ProductionCostFactor"] / 100.0;
                double productionCost = currentCost * productionFactor;
                result.ProductionCost = productionCost;
                currentCost += productionCost;
            }

            if (factors.ContainsKey("AdministrationCostFactor"))
            {
                double adminFactor = factors["AdministrationCostFactor"] / 100.0;
                double adminCost = currentCost * adminFactor;
                result.AdministrationCost = adminCost;
                currentCost += adminCost;
            }

            if (factors.ContainsKey("SalesCostFactor"))
            {
                double salesFactor = factors["SalesCostFactor"] / 100.0;
                double salesCost = currentCost * salesFactor;
                result.SalesCost = salesCost;
                currentCost += salesCost;
            }

            if (factors.ContainsKey("ProfitMarginFactor"))
            {
                double profitFactor = factors["ProfitMarginFactor"] / 100.0;
                double profitAmount = currentCost * profitFactor;
                result.ProfitMargin = profitAmount;
                currentCost += profitAmount;
            }

            result.FinalOfferPrice = currentCost;
            result.AppliedFactors = factors;

            return result;
        }

        private void DisplayResultsWithCalculation()
        {
            if (CalculationResult == null)
            {
                DisplayResults();
                return;
            }

            var results = new System.Text.StringBuilder();
            results.AppendLine("=== COATING CALCULATION RESULTS ===");
            results.AppendLine();
            results.AppendLine($"Application Type: {MaterialConfig.ApplicationType}");
            if (MaterialConfig.Primer != null || MaterialConfig.Topcoat != null)
            {
                results.AppendLine($"Coat Multipliers: Primer={MaterialConfig.PrimerCoatMultiplier:F0}x, Topcoat={MaterialConfig.TopcoatCoatMultiplier:F0}x");
            }
            results.AppendLine();

            // Display individual objects (same as before)
            if (selectedObjects != null && selectedObjects.Count > 0)
            {
                results.AppendLine("INDIVIDUAL OBJECTS:");
                results.AppendLine();

                for (int i = 0; i < selectedObjects.Count; i++)
                {
                    var obj = selectedObjects[i];
                    results.AppendLine($"Object {i + 1}: {obj.Name}");
                    results.AppendLine($"  Surface Area: {obj.SurfaceArea_mm2:F2} mm² ({obj.SurfaceArea_mm2 / 1000000:F4} m²)");

                    // Calculate costs and material amounts for this individual object
                    double objArea_m2 = obj.SurfaceArea_mm2 / 1000000.0;
                    var objCosts = CalculateCosts(obj.SurfaceArea_mm2, MaterialConfig);
                    var objTime = EstimateTime(obj.SurfaceArea_mm2, MaterialConfig.TimeFactor);

                    if (MaterialConfig.Primer != null)
                    {
                        double primerGrams = objArea_m2 * MaterialConfig.Primer.ConsumptionPerSqM * MaterialConfig.PrimerCoatMultiplier;
                        double primerKg = primerGrams / 1000.0;
                        double primerHardenerGrams = primerGrams * (MaterialConfig.PrimerHardenerPercent / 100.0);
                        double primerHardenerKg = primerHardenerGrams / 1000.0;
                        double primerThinnerGrams = primerGrams * (MaterialConfig.PrimerThinnerPercent / 100.0);
                        double primerThinnerKg = primerThinnerGrams / 1000.0;

                        results.AppendLine($"  Primer ({MaterialConfig.PrimerCoatMultiplier:F0}x coats): {primerGrams:F1} g ({primerKg:F3} kg) - Fr. {objCosts.PrimerCost:F2}");
                        results.AppendLine($"    + Hardener ({MaterialConfig.PrimerHardenerPercent:F1}%): {primerHardenerGrams:F1} g ({primerHardenerKg:F3} kg) - Fr. {objCosts.PrimerHardenerCost:F2}");
                        results.AppendLine($"    + Thinner ({MaterialConfig.PrimerThinnerPercent:F1}%): {primerThinnerGrams:F1} g ({primerThinnerKg:F3} kg) - Fr. {objCosts.PrimerThinnerCost:F2}");
                    }
                    if (MaterialConfig.Topcoat != null)
                    {
                        double topcoatGrams = objArea_m2 * MaterialConfig.Topcoat.ConsumptionPerSqM * MaterialConfig.TopcoatCoatMultiplier;
                        double topcoatKg = topcoatGrams / 1000.0;
                        double topcoatHardenerGrams = topcoatGrams * (MaterialConfig.TopcoatHardenerPercent / 100.0);
                        double topcoatHardenerKg = topcoatHardenerGrams / 1000.0;
                        double topcoatThinnerGrams = topcoatGrams * (MaterialConfig.TopcoatThinnerPercent / 100.0);
                        double topcoatThinnerKg = topcoatThinnerGrams / 1000.0;

                        results.AppendLine($"  Topcoat ({MaterialConfig.TopcoatCoatMultiplier:F0}x coats): {topcoatGrams:F1} g ({topcoatKg:F3} kg) - Fr. {objCosts.TopcoatCost:F2}");
                        results.AppendLine($"    + Hardener ({MaterialConfig.TopcoatHardenerPercent:F1}%): {topcoatHardenerGrams:F1} g ({topcoatHardenerKg:F3} kg) - Fr. {objCosts.TopcoatHardenerCost:F2}");
                        results.AppendLine($"    + Thinner ({MaterialConfig.TopcoatThinnerPercent:F1}%): {topcoatThinnerGrams:F1} g ({topcoatThinnerKg:F3} kg) - Fr. {objCosts.TopcoatThinnerCost:F2}");
                    }
                    results.AppendLine($"  Total Material Cost: Fr. {objCosts.TotalMaterialCost:F2}");
                    results.AppendLine($"  Estimated Time: {objTime:F2} hours");
                    results.AppendLine();
                }

                results.AppendLine("---------------------------------");
                results.AppendLine();
            }

            // Display summary (same as before)
            results.AppendLine("SUMMARY:");
            results.AppendLine();
            results.AppendLine($"Total Objects: {selectedObjects.Count}");
            results.AppendLine($"Total Surface Area: {surfaceArea_mm2:F2} mm²");
            results.AppendLine($"                    ({surfaceArea_mm2 / 1000000:F4} m²)");
            results.AppendLine();

            double totalArea_m2 = surfaceArea_mm2 / 1000000.0;

            if (MaterialConfig.Primer != null)
            {
                double totalPrimerGrams = totalArea_m2 * MaterialConfig.Primer.ConsumptionPerSqM * MaterialConfig.PrimerCoatMultiplier;
                double totalPrimerKg = totalPrimerGrams / 1000.0;
                double totalPrimerHardenerGrams = totalPrimerGrams * (MaterialConfig.PrimerHardenerPercent / 100.0);
                double totalPrimerHardenerKg = totalPrimerHardenerGrams / 1000.0;
                double totalPrimerThinnerGrams = totalPrimerGrams * (MaterialConfig.PrimerThinnerPercent / 100.0);
                double totalPrimerThinnerKg = totalPrimerThinnerGrams / 1000.0;

                results.AppendLine($"Primer: {MaterialConfig.Primer.Name}");
                results.AppendLine($"  Consumption: {MaterialConfig.Primer.ConsumptionPerSqM:F1} g/m²");
                results.AppendLine($"  Coats: {MaterialConfig.PrimerCoatMultiplier:F0}x");
                results.AppendLine($"  Total Amount: {totalPrimerGrams:F1} g ({totalPrimerKg:F3} kg)");
                results.AppendLine($"  Price: Fr. {MaterialConfig.Primer.PricePerKg:F2}/kg");
                results.AppendLine($"  Total Cost: Fr. {Costs.PrimerCost:F2}");
                results.AppendLine($"  Hardener ({MaterialConfig.PrimerHardenerPercent:F1}%): {totalPrimerHardenerGrams:F1} g ({totalPrimerHardenerKg:F3} kg) - Fr. {Costs.PrimerHardenerCost:F2}");
                results.AppendLine($"  Thinner ({MaterialConfig.PrimerThinnerPercent:F1}%): {totalPrimerThinnerGrams:F1} g ({totalPrimerThinnerKg:F3} kg) - Fr. {Costs.PrimerThinnerCost:F2}");
                results.AppendLine();
            }

            if (MaterialConfig.Topcoat != null)
            {
                double totalTopcoatGrams = totalArea_m2 * MaterialConfig.Topcoat.ConsumptionPerSqM * MaterialConfig.TopcoatCoatMultiplier;
                double totalTopcoatKg = totalTopcoatGrams / 1000.0;
                double totalTopcoatHardenerGrams = totalTopcoatGrams * (MaterialConfig.TopcoatHardenerPercent / 100.0);
                double totalTopcoatHardenerKg = totalTopcoatHardenerGrams / 1000.0;
                double totalTopcoatThinnerGrams = totalTopcoatGrams * (MaterialConfig.TopcoatThinnerPercent / 100.0);
                double totalTopcoatThinnerKg = totalTopcoatThinnerGrams / 1000.0;

                results.AppendLine($"Topcoat: {MaterialConfig.Topcoat.Name}");
                results.AppendLine($"  Consumption: {MaterialConfig.Topcoat.ConsumptionPerSqM:F1} g/m²");
                results.AppendLine($"  Coats: {MaterialConfig.TopcoatCoatMultiplier:F0}x");
                results.AppendLine($"  Total Amount: {totalTopcoatGrams:F1} g ({totalTopcoatKg:F3} kg)");
                results.AppendLine($"  Price: Fr. {MaterialConfig.Topcoat.PricePerKg:F2}/kg");
                results.AppendLine($"  Total Cost: Fr. {Costs.TopcoatCost:F2}");
                results.AppendLine($"  Hardener ({MaterialConfig.TopcoatHardenerPercent:F1}%): {totalTopcoatHardenerGrams:F1} g ({totalTopcoatHardenerKg:F3} kg) - Fr. {Costs.TopcoatHardenerCost:F2}");
                results.AppendLine($"  Thinner ({MaterialConfig.TopcoatThinnerPercent:F1}%): {totalTopcoatThinnerGrams:F1} g ({totalTopcoatThinnerKg:F3} kg) - Fr. {Costs.TopcoatThinnerCost:F2}");
                results.AppendLine();
            }

            results.AppendLine($"Total Material Cost: Fr. {Costs.TotalMaterialCost:F2}");
            results.AppendLine($"Total Estimated Time: {EstimatedTime:F2} hours");
            results.AppendLine($"Time Factor: {MaterialConfig.TimeFactor:F2} h/m²");

            // Add calculation breakdown
            results.AppendLine();
            results.AppendLine("=== OFFER CALCULATION ===");
            results.AppendLine();
            results.AppendLine($"Base Material Cost:     Fr. {CalculationResult.BaseMaterialCost:F2}");
            if (CalculationResult.MaterialSurcharge > 0 && CalculationResult.AppliedFactors.TryGetValue("MaterialCostFactor", out double materialFactor))
                results.AppendLine($"Material Surcharge ({materialFactor:F1}%): Fr. {CalculationResult.MaterialSurcharge:F2}");
            if (CalculationResult.ProductionCost > 0 && CalculationResult.AppliedFactors.TryGetValue("ProductionCostFactor", out double productionFactor))
                results.AppendLine($"Production Cost ({productionFactor:F1}%):    Fr. {CalculationResult.ProductionCost:F2}");
            if (CalculationResult.AdministrationCost > 0 && CalculationResult.AppliedFactors.TryGetValue("AdministrationCostFactor", out double adminFactor))
                results.AppendLine($"Administration Cost ({adminFactor:F1}%): Fr. {CalculationResult.AdministrationCost:F2}");
            if (CalculationResult.SalesCost > 0 && CalculationResult.AppliedFactors.TryGetValue("SalesCostFactor", out double salesFactor))
                results.AppendLine($"Sales Cost ({salesFactor:F1}%):         Fr. {CalculationResult.SalesCost:F2}");
            if (CalculationResult.ProfitMargin > 0 && CalculationResult.AppliedFactors.TryGetValue("ProfitMarginFactor", out double profitFactor))
                results.AppendLine($"Profit Margin ({profitFactor:F1}%):      Fr. {CalculationResult.ProfitMargin:F2}");
            results.AppendLine();
            results.AppendLine($"FINAL OFFER PRICE: Fr. {CalculationResult.FinalOfferPrice:F2}");

            resultsTextArea.Text = results.ToString();
        }
    }

    public class CalculationResult
    {
        public double BaseMaterialCost { get; set; }
        public double MaterialSurcharge { get; set; }
        public double ProductionCost { get; set; }
        public double AdministrationCost { get; set; }
        public double SalesCost { get; set; }
        public double ProfitMargin { get; set; }
        public double FinalOfferPrice { get; set; }
        public double MaterialCostWithSurcharge { get; set; }
        public Dictionary<string, double> AppliedFactors { get; set; }
    }
}