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

                // Step 1: Open the Coating Panel (if not already visible)
                Panels.OpenPanel(CoatingPanel.PanelId);

                // Step 2: Select 3D objects (Breps)
                List<Brep> selectedBreps = SelectBreps(doc);
                if (selectedBreps.Count == 0)
                {
                    RhinoApp.WriteLine("No valid 3D objects selected.");
                    return Result.Cancel;
                }

                // Step 3: Calculate total surface area
                double totalSurfaceArea_mm2 = CalculateTotalSurfaceArea(selectedBreps);
                RhinoApp.WriteLine($"Total surface area: {totalSurfaceArea_mm2:F2} mm²");

                // Step 4: Update the panel with the calculated surface area
                var panel = Panels.GetPanel<CoatingPanel>(doc);
                if (panel != null)
                {
                    panel.UpdateSurfaceArea(totalSurfaceArea_mm2);
                }
                else
                {
                    RhinoApp.WriteLine("Warning: Could not find Coating Panel to update.");
                }

                // Step 5: Visualize selected surfaces
                VisualizeSurfaces(doc, selectedBreps);

                RhinoApp.WriteLine("CoatingApp: Surface area calculated and panel updated.");
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
