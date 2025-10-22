using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Eto.Forms;
using Eto.Drawing;
using Rhino;
using Rhino.UI;
using ClosedXML.Excel;

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

        private double surfaceArea_mm2;
        private List<ObjectInfo> selectedObjects;
        private Dictionary<string, MaterialInfo> materials;

        public MaterialConfig MaterialConfig { get; private set; }
        public CostCalculation Costs { get; private set; }
        public double EstimatedTime { get; private set; }

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
                    ? $"{objectCount} - {surfaceArea_mm2:F2} mm� ({surfaceArea_mm2 / 1000000:F4} m�)"
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
            // Create layout
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
                    ? $"{surfaceArea_mm2:F2} mm� ({surfaceArea_mm2 / 1000000:F4} m�)"
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
            layout.AddRow(new Label { Text = "Consumption (g/m2):", Width = 120 });
            primerConsumptionStepper = new NumericStepper
            {
                MinValue = 0,
                MaxValue = 1000,
                Value = 200, // Default 200 g/m2
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
            layout.AddRow(new Label { Text = "Consumption (g/m2):", Width = 120 });
            topcoatConsumptionStepper = new NumericStepper
            {
                MinValue = 0,
                MaxValue = 1000,
                Value = 200, // Default 200 g/m2
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
            layout.AddRow(new Label { Text = "Time Factor (h/m2):", Width = 120 });
            timeFactorStepper = new NumericStepper
            {
                MinValue = 0.1,
                MaxValue = 10.0,
                Value = 0.5,
                DecimalPlaces = 2,
                Increment = 0.1
            };
            layout.AddRow(timeFactorStepper);
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

            // Set content for the panel
            Content = layout;
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
                    ThinnerPricePerKg = thinnerPriceStepper.Value
                };

                // Calculate costs
                Costs = CalculateCosts(surfaceArea_mm2, MaterialConfig);

                // Calculate time
                EstimatedTime = EstimateTime(surfaceArea_mm2, MaterialConfig.TimeFactor);

                // Display results
                DisplayResults();

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
                // Base primer calculation
                double primerGrams = surfaceArea_m2 * config.Primer.ConsumptionPerSqM;
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
                // Base topcoat calculation
                double topcoatGrams = surfaceArea_m2 * config.Topcoat.ConsumptionPerSqM;
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
                    results.AppendLine($"  Surface Area: {obj.SurfaceArea_mm2:F2} mm� ({obj.SurfaceArea_mm2 / 1000000:F4} m�)");
                    
                    // Calculate costs and material amounts for this individual object
                    double objArea_m2 = obj.SurfaceArea_mm2 / 1000000.0;
                    var objCosts = CalculateCosts(obj.SurfaceArea_mm2, MaterialConfig);
                    var objTime = EstimateTime(obj.SurfaceArea_mm2, MaterialConfig.TimeFactor);
                    
                    if (MaterialConfig.Primer != null)
                    {
                        double primerGrams = objArea_m2 * MaterialConfig.Primer.ConsumptionPerSqM;
                        double primerKg = primerGrams / 1000.0;
                        double primerHardenerGrams = primerGrams * (MaterialConfig.PrimerHardenerPercent / 100.0);
                        double primerHardenerKg = primerHardenerGrams / 1000.0;
                        double primerThinnerGrams = primerGrams * (MaterialConfig.PrimerThinnerPercent / 100.0);
                        double primerThinnerKg = primerThinnerGrams / 1000.0;
                        
                        results.AppendLine($"  Primer: {primerGrams:F1} g ({primerKg:F3} kg) - Fr. {objCosts.PrimerCost:F2}");
                        results.AppendLine($"    + Hardener ({MaterialConfig.PrimerHardenerPercent:F1}%): {primerHardenerGrams:F1} g ({primerHardenerKg:F3} kg) - Fr. {objCosts.PrimerHardenerCost:F2}");
                        results.AppendLine($"    + Thinner ({MaterialConfig.PrimerThinnerPercent:F1}%): {primerThinnerGrams:F1} g ({primerThinnerKg:F3} kg) - Fr. {objCosts.PrimerThinnerCost:F2}");
                    }
                    if (MaterialConfig.Topcoat != null)
                    {
                        double topcoatGrams = objArea_m2 * MaterialConfig.Topcoat.ConsumptionPerSqM;
                        double topcoatKg = topcoatGrams / 1000.0;
                        double topcoatHardenerGrams = topcoatGrams * (MaterialConfig.TopcoatHardenerPercent / 100.0);
                        double topcoatHardenerKg = topcoatHardenerGrams / 1000.0;
                        double topcoatThinnerGrams = topcoatGrams * (MaterialConfig.TopcoatThinnerPercent / 100.0);
                        double topcoatThinnerKg = topcoatThinnerGrams / 1000.0;
                        
                        results.AppendLine($"  Topcoat: {topcoatGrams:F1} g ({topcoatKg:F3} kg) - Fr. {objCosts.TopcoatCost:F2}");
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
            results.AppendLine($"Total Surface Area: {surfaceArea_mm2:F2} mm�");
            results.AppendLine($"                    ({surfaceArea_mm2 / 1000000:F4} m�)");
            results.AppendLine();

            double totalArea_m2 = surfaceArea_mm2 / 1000000.0;
            
            if (MaterialConfig.Primer != null)
            {
                double totalPrimerGrams = totalArea_m2 * MaterialConfig.Primer.ConsumptionPerSqM;
                double totalPrimerKg = totalPrimerGrams / 1000.0;
                double totalPrimerHardenerGrams = totalPrimerGrams * (MaterialConfig.PrimerHardenerPercent / 100.0);
                double totalPrimerHardenerKg = totalPrimerHardenerGrams / 1000.0;
                double totalPrimerThinnerGrams = totalPrimerGrams * (MaterialConfig.PrimerThinnerPercent / 100.0);
                double totalPrimerThinnerKg = totalPrimerThinnerGrams / 1000.0;
                
                results.AppendLine($"Primer: {MaterialConfig.Primer.Name}");
                results.AppendLine($"  Consumption: {MaterialConfig.Primer.ConsumptionPerSqM:F1} g/m�");
                results.AppendLine($"  Total Amount: {totalPrimerGrams:F1} g ({totalPrimerKg:F3} kg)");
                results.AppendLine($"  Price: Fr. {MaterialConfig.Primer.PricePerKg:F2}/kg");
                results.AppendLine($"  Total Cost: Fr. {Costs.PrimerCost:F2}");
                results.AppendLine($"  Hardener ({MaterialConfig.PrimerHardenerPercent:F1}%): {totalPrimerHardenerGrams:F1} g ({totalPrimerHardenerKg:F3} kg) - Fr. {Costs.PrimerHardenerCost:F2}");
                results.AppendLine($"  Thinner ({MaterialConfig.PrimerThinnerPercent:F1}%): {totalPrimerThinnerGrams:F1} g ({totalPrimerThinnerKg:F3} kg) - Fr. {Costs.PrimerThinnerCost:F2}");
                results.AppendLine();
            }

            if (MaterialConfig.Topcoat != null)
            {
                double totalTopcoatGrams = totalArea_m2 * MaterialConfig.Topcoat.ConsumptionPerSqM;
                double totalTopcoatKg = totalTopcoatGrams / 1000.0;
                double totalTopcoatHardenerGrams = totalTopcoatGrams * (MaterialConfig.TopcoatHardenerPercent / 100.0);
                double totalTopcoatHardenerKg = totalTopcoatHardenerGrams / 1000.0;
                double totalTopcoatThinnerGrams = totalTopcoatGrams * (MaterialConfig.TopcoatThinnerPercent / 100.0);
                double totalTopcoatThinnerKg = totalTopcoatThinnerGrams / 1000.0;
                
                results.AppendLine($"Topcoat: {MaterialConfig.Topcoat.Name}");
                results.AppendLine($"  Consumption: {MaterialConfig.Topcoat.ConsumptionPerSqM:F1} g/m�");
                results.AppendLine($"  Total Amount: {totalTopcoatGrams:F1} g ({totalTopcoatKg:F3} kg)");
                results.AppendLine($"  Price: Fr. {MaterialConfig.Topcoat.PricePerKg:F2}/kg");
                results.AppendLine($"  Total Cost: Fr. {Costs.TopcoatCost:F2}");
                results.AppendLine($"  Hardener ({MaterialConfig.TopcoatHardenerPercent:F1}%): {totalTopcoatHardenerGrams:F1} g ({totalTopcoatHardenerKg:F3} kg) - Fr. {Costs.TopcoatHardenerCost:F2}");
                results.AppendLine($"  Thinner ({MaterialConfig.TopcoatThinnerPercent:F1}%): {totalTopcoatThinnerGrams:F1} g ({totalTopcoatThinnerKg:F3} kg) - Fr. {Costs.TopcoatThinnerCost:F2}");
                results.AppendLine();
            }

            results.AppendLine($"Total Material Cost: Fr. {Costs.TotalMaterialCost:F2}");
            results.AppendLine($"Total Estimated Time: {EstimatedTime:F2} hours");
            results.AppendLine($"Time Factor: {MaterialConfig.TimeFactor:F2} h/m�");

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
                        EstimatedTime
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
    }
}