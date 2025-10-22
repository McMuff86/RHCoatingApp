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

                // Step 2: Select 3D objects (Breps) and collect object information
                var selectedObjects = new List<ObjectInfo>();
                double totalSurfaceArea_mm2 = 0;
                
                using (GetObject getObject = new GetObject())
                {
                    getObject.SetCommandPrompt("Select 3D objects for coating calculation");
                    getObject.GeometryFilter = Rhino.DocObjects.ObjectType.Brep;
                    getObject.SubObjectSelect = false;
                    getObject.EnablePreSelect(false, true);
                    getObject.GetMultiple(1, 0);

                    if (getObject.CommandResult() != Result.Success)
                    {
                        RhinoApp.WriteLine("No valid 3D objects selected.");
                        return getObject.CommandResult();
                    }

                    for (int i = 0; i < getObject.ObjectCount; i++)
                    {
                        Rhino.DocObjects.ObjRef objRef = getObject.Object(i);
                        Brep brep = objRef.Brep();
                        
                        if (brep != null && brep.IsValid)
                        {
                            double surfaceArea = brep.GetArea();
                            totalSurfaceArea_mm2 += surfaceArea;

                            var objInfo = new ObjectInfo
                            {
                                Name = objRef.Object().Name ?? $"Object {i + 1}",
                                SurfaceArea_mm2 = surfaceArea,
                                ObjectId = objRef.ObjectId
                            };
                            selectedObjects.Add(objInfo);

                            RhinoApp.WriteLine($"Added: {objInfo.Name} - {surfaceArea:F2} mm²");
                        }
                    }
                }

                if (selectedObjects.Count == 0)
                {
                    RhinoApp.WriteLine("No valid 3D objects selected.");
                    return Result.Cancel;
                }

                RhinoApp.WriteLine($"Total surface area: {totalSurfaceArea_mm2:F2} mm² ({selectedObjects.Count} objects)");

                // Step 3: Update the panel with the calculated surface area and object information
                var panel = Panels.GetPanel<CoatingPanel>(doc);
                if (panel != null)
                {
                    panel.UpdateSurfaceArea(totalSurfaceArea_mm2, selectedObjects);
                }
                else
                {
                    RhinoApp.WriteLine("Warning: Could not find Coating Panel to update.");
                }

                // Step 4: Visualize selected surfaces
                doc.Views.Redraw();

                RhinoApp.WriteLine("CoatingApp: Surface area calculated and panel updated.");
                return Result.Success;
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Error in CoatingApp: {ex.Message}");
                return Result.Failure;
            }
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

    public class ObjectInfo
    {
        public string Name { get; set; }
        public double SurfaceArea_mm2 { get; set; }
        public Guid ObjectId { get; set; }
    }
}
