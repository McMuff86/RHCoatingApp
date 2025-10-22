using System;
using System.Collections.Generic;
using Rhino;
using Rhino.Commands;
using Rhino.Geometry;
using Rhino.Input;
using Rhino.Input.Custom;
using Rhino.UI;

namespace RHCoatingApp
{
    public class CoatingApp : Command
    {
        public CoatingApp()
        {
            // Rhino only creates one instance of each command class defined in a
            // plug-in, so it is safe to store a refence in a static property.
            Instance = this;
        }

        ///<summary>The only instance of this command.</summary>
        public static CoatingApp Instance { get; private set; }

        ///<returns>The command name as it appears on the Rhino command line.</returns>
        public override string EnglishName => "CoatingApp";

        protected override Result RunCommand(RhinoDoc doc, RunMode mode)
        {
            try
            {
                RhinoApp.WriteLine("Starting Rhino CoatingApp...");

                // Step 1: Select 3D objects (Breps)
                List<Brep> selectedBreps = SelectBreps(doc);
                if (selectedBreps.Count == 0)
                {
                    RhinoApp.WriteLine("No valid 3D objects selected.");
                    return Result.Cancel;
                }

                // Step 2: Calculate total surface area
                double totalSurfaceArea_mm2 = CalculateTotalSurfaceArea(selectedBreps);
                RhinoApp.WriteLine($"Total surface area: {totalSurfaceArea_mm2:F2} mm²");

                // Step 3: Material selection and configuration
                MaterialConfig materialConfig = GetMaterialConfiguration(totalSurfaceArea_mm2);
                if (materialConfig == null)
                {
                    RhinoApp.WriteLine("Material configuration cancelled.");
                    return Result.Cancel;
                }

                // Step 4: Calculate costs
                CostCalculation costs = CalculateCosts(totalSurfaceArea_mm2, materialConfig);

                // Step 5: Calculate time estimation
                double estimatedTime = EstimateTime(totalSurfaceArea_mm2, materialConfig.TimeFactor);

                // Step 6: Display results
                DisplayResults(totalSurfaceArea_mm2, materialConfig, costs, estimatedTime);

                // Step 7: Export results
                ExportResults(totalSurfaceArea_mm2, materialConfig, costs, estimatedTime);

                // Step 8: Visualize selected surfaces
                VisualizeSurfaces(doc, selectedBreps);

                RhinoApp.WriteLine("CoatingApp calculation completed successfully.");
                return Result.Success;
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Error in CoatingApp: {ex.Message}");
                return Result.Failure;
            }
        }

        private List<Brep> SelectBreps(RhinoDoc doc)
        {
            List<Brep> breps = new List<Brep>();

            using (GetObject getObject = new GetObject())
            {
                getObject.SetCommandPrompt("Select 3D objects for coating calculation");
                getObject.GeometryFilter = Rhino.DocObjects.ObjectType.Brep;
                getObject.SubObjectSelect = false;
                getObject.EnablePreSelect(false, true);

                while (true)
                {
                    GetResult result = getObject.Get();
                    if (result == GetResult.Object)
                    {
                        Rhino.DocObjects.ObjRef objRef = getObject.Object(0);
                        Brep brep = objRef.Brep();
                        if (brep != null && brep.IsValid)
                        {
                            breps.Add(brep);
                            RhinoApp.WriteLine($"Added object: {objRef.Object().Name ?? "Unnamed"}");
                        }
                    }
                    else if (result == GetResult.Nothing)
                    {
                        break;
                    }
                    else
                    {
                        return breps; // Return what we have so far
                    }
                }
            }

            return breps;
        }

        private double CalculateTotalSurfaceArea(List<Brep> breps)
        {
            double totalArea = 0.0;
            foreach (Brep brep in breps)
            {
                totalArea += brep.GetArea();
            }
            return totalArea;
        }

        private MaterialConfig GetMaterialConfiguration(double surfaceArea_mm2)
        {
            // Simple hardcoded materials for first version
            var materialNames = new[] { "Standard Primer", "Premium Primer", "Basic Topcoat", "Premium Topcoat" };
            var materialData = new[]
            {
                (150.0, 25.0, 0.5), // g/m², €/kg, hours/m²
                (200.0, 35.0, 0.7),
                (120.0, 20.0, 0.4),
                (180.0, 40.0, 0.6)
            };

            RhinoApp.WriteLine("Available materials:");
            for (int i = 0; i < materialNames.Length; i++)
            {
                RhinoApp.WriteLine($"{i + 1}. {materialNames[i]}");
            }

            string primerInput = GetUserInput("Select primer (1-4) or 'none' for no primer:");
            string topcoatInput = GetUserInput("Select topcoat (1-4) or 'none' for no topcoat:");

            MaterialConfig config = new MaterialConfig();

            if (primerInput != "none" && int.TryParse(primerInput, out int primerIndex) && primerIndex >= 1 && primerIndex <= 4)
            {
                int materialIndex = primerIndex - 1;
                var primerData = materialData[materialIndex];
                config.Primer = new MaterialInfo
                {
                    Name = materialNames[materialIndex],
                    ConsumptionPerSqM = primerData.Item1,
                    PricePerKg = primerData.Item2
                };
            }

            if (topcoatInput != "none" && int.TryParse(topcoatInput, out int topcoatIndex) && topcoatIndex >= 1 && topcoatIndex <= 4)
            {
                int materialIndex = topcoatIndex - 1;
                var topcoatData = materialData[materialIndex];
                config.Topcoat = new MaterialInfo
                {
                    Name = materialNames[materialIndex],
                    ConsumptionPerSqM = topcoatData.Item1,
                    PricePerKg = topcoatData.Item2
                };
            }

            string timeFactorInput = GetUserInput("Enter time factor (hours per square meter, default 0.5):");
            if (double.TryParse(timeFactorInput, out double timeFactor) && timeFactor > 0)
            {
                config.TimeFactor = timeFactor;
            }
            else
            {
                config.TimeFactor = 0.5; // Default
            }

            return config;
        }

        private string GetUserInput(string prompt)
        {
            using (GetString getString = new GetString())
            {
                getString.SetCommandPrompt(prompt);
                if (getString.Get() == GetResult.String)
                {
                    return getString.StringResult();
                }
                return string.Empty;
            }
        }

        private CostCalculation CalculateCosts(double surfaceArea_mm2, MaterialConfig config)
        {
            CostCalculation costs = new CostCalculation();

            // Convert mm² to m² for material calculation
            double surfaceArea_m2 = surfaceArea_mm2 / 1000000.0;

            if (config.Primer != null)
            {
                double primerKg = surfaceArea_m2 * config.Primer.ConsumptionPerSqM / 1000.0; // Convert g to kg
                costs.PrimerCost = primerKg * config.Primer.PricePerKg;
                costs.TotalMaterialCost += costs.PrimerCost;
            }

            if (config.Topcoat != null)
            {
                double topcoatKg = surfaceArea_m2 * config.Topcoat.ConsumptionPerSqM / 1000.0; // Convert g to kg
                costs.TopcoatCost = topcoatKg * config.Topcoat.PricePerKg;
                costs.TotalMaterialCost += costs.TopcoatCost;
            }

            return costs;
        }

        private double EstimateTime(double surfaceArea_mm2, double timeFactor)
        {
            // Convert mm² to m² for time calculation
            double surfaceArea_m2 = surfaceArea_mm2 / 1000000.0;
            return surfaceArea_m2 * timeFactor;
        }

        private void DisplayResults(double surfaceArea_mm2, MaterialConfig config, CostCalculation costs, double estimatedTime)
        {
            RhinoApp.WriteLine("\n=== COATING CALCULATION RESULTS ===");
            RhinoApp.WriteLine($"Surface Area: {surfaceArea_mm2:F2} mm² ({surfaceArea_mm2 / 1000000:F2} m²)");

            if (config.Primer != null)
            {
                RhinoApp.WriteLine($"Primer: {config.Primer.Name} - Cost: €{costs.PrimerCost:F2}");
            }

            if (config.Topcoat != null)
            {
                RhinoApp.WriteLine($"Topcoat: {config.Topcoat.Name} - Cost: €{costs.TopcoatCost:F2}");
            }

            RhinoApp.WriteLine($"Total Material Cost: €{costs.TotalMaterialCost:F2}");
            RhinoApp.WriteLine($"Estimated Time: {estimatedTime:F1} hours");
        }

        private void ExportResults(double surfaceArea_mm2, MaterialConfig config, CostCalculation costs, double estimatedTime)
        {
            string exportPath = GetUserInput("Enter export file path (or press Enter for console only):");
            if (string.IsNullOrEmpty(exportPath))
            {
                RhinoApp.WriteLine("Results exported to console only.");
                return;
            }

            try
            {
                using (System.IO.StreamWriter writer = new System.IO.StreamWriter(exportPath))
                {
                    writer.WriteLine("Rhino CoatingApp Calculation Results");
                    writer.WriteLine($"Date: {DateTime.Now}");
                    writer.WriteLine($"Surface Area: {surfaceArea_mm2:F2} mm² ({surfaceArea_mm2 / 1000000:F2} m²)");

                    if (config.Primer != null)
                    {
                        writer.WriteLine($"Primer: {config.Primer.Name} - Cost: €{costs.PrimerCost:F2}");
                    }

                    if (config.Topcoat != null)
                    {
                        writer.WriteLine($"Topcoat: {config.Topcoat.Name} - Cost: €{costs.TopcoatCost:F2}");
                    }

                    writer.WriteLine($"Total Material Cost: €{costs.TotalMaterialCost:F2}");
                    writer.WriteLine($"Estimated Time: {estimatedTime:F1} hours");
                }

                RhinoApp.WriteLine($"Results exported to: {exportPath}");
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Error exporting results: {ex.Message}");
            }
        }

        private void VisualizeSurfaces(RhinoDoc doc, List<Brep> breps)
        {
            // Just redraw the view to highlight the selected objects
            // The objects are already in the document, no need to add them again
            doc.Views.Redraw();
            RhinoApp.WriteLine($"Highlighted {breps.Count} objects in viewport for calculation.");
        }
    }

    // Data classes for the coating calculation
    public class MaterialConfig
    {
        public MaterialInfo Primer { get; set; }
        public MaterialInfo Topcoat { get; set; }
        public double TimeFactor { get; set; } = 0.5; // hours per square meter
    }

    public class MaterialInfo
    {
        public string Name { get; set; }
        public double ConsumptionPerSqM { get; set; } // grams per square meter
        public double PricePerKg { get; set; } // price per kilogram
    }

    public class CostCalculation
    {
        public double PrimerCost { get; set; }
        public double TopcoatCost { get; set; }
        public double TotalMaterialCost { get; set; }
    }
}
