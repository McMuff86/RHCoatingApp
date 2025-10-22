using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;

namespace RHCoatingApp
{
    /// <summary>
    /// Handles all export functionality for coating calculations
    /// Supports TXT, CSV, and XLSX formats
    /// </summary>
    public class CoatingExporter
    {
        /// <summary>
        /// Exports coating calculation results to a file in the specified format
        /// </summary>
        /// <param name="filename">Full path to the output file</param>
        /// <param name="surfaceArea_mm2">Total surface area in mm²</param>
        /// <param name="selectedObjects">List of selected objects with their details</param>
        /// <param name="config">Material configuration used for calculation</param>
        /// <param name="costs">Calculated costs</param>
        /// <param name="estimatedTime">Estimated time for the coating job</param>
        public static void Export(string filename, double surfaceArea_mm2, List<ObjectInfo> selectedObjects, 
                                  MaterialConfig config, CostCalculation costs, double estimatedTime)
        {
            string extension = Path.GetExtension(filename).ToLower();
            
            switch (extension)
            {
                case ".xlsx":
                    ExportToExcel(filename, surfaceArea_mm2, selectedObjects, config, costs, estimatedTime);
                    break;
                case ".csv":
                    ExportToCSV(filename, surfaceArea_mm2, selectedObjects, config, costs, estimatedTime);
                    break;
                case ".txt":
                default:
                    ExportToText(filename, surfaceArea_mm2, selectedObjects, config, costs, estimatedTime);
                    break;
            }
        }

        #region Text Export

        private static void ExportToText(string filename, double surfaceArea_mm2, List<ObjectInfo> selectedObjects, 
                                         MaterialConfig config, CostCalculation costs, double estimatedTime)
        {
            using (var writer = new StreamWriter(filename))
            {
                writer.WriteLine("Rhino CoatingApp Calculation Results");
                writer.WriteLine($"Date: {DateTime.Now}");
                writer.WriteLine($"Application Type: {config.ApplicationType}");
                writer.WriteLine();

                // Export individual objects
                if (selectedObjects != null && selectedObjects.Count > 0)
                {
                    writer.WriteLine("INDIVIDUAL OBJECTS:");
                    writer.WriteLine();
                    
                    for (int i = 0; i < selectedObjects.Count; i++)
                    {
                        var obj = selectedObjects[i];
                        writer.WriteLine($"Object {i + 1}: {obj.Name}");
                        writer.WriteLine($"  Surface Area: {obj.SurfaceArea_mm2:F2} mm² ({obj.SurfaceArea_mm2 / 1000000:F4} m²)");
                        
                        double objArea_m2 = obj.SurfaceArea_mm2 / 1000000.0;
                        var objCosts = CalculateIndividualCosts(obj.SurfaceArea_mm2, config);
                        var objTime = EstimateIndividualTime(obj.SurfaceArea_mm2, config.TimeFactor);
                        
                        WriteObjectMaterialDetails(writer, objArea_m2, config, objCosts);
                        writer.WriteLine($"  Total Material Cost: Fr. {objCosts.TotalMaterialCost:F2}");
                        writer.WriteLine($"  Estimated Time: {objTime:F2} hours");
                        writer.WriteLine();
                    }

                    writer.WriteLine("─────────────────────────────────");
                    writer.WriteLine();
                }

                // Export summary
                WriteSummary(writer, surfaceArea_mm2, selectedObjects.Count, config, costs, estimatedTime);
            }
        }

        private static void WriteObjectMaterialDetails(StreamWriter writer, double objArea_m2, MaterialConfig config, CostCalculation objCosts)
        {
            if (config.Primer != null)
            {
                double primerGrams = objArea_m2 * config.Primer.ConsumptionPerSqM;
                double primerKg = primerGrams / 1000.0;
                double primerHardenerGrams = primerGrams * (config.PrimerHardenerPercent / 100.0);
                double primerHardenerKg = primerHardenerGrams / 1000.0;
                double primerThinnerGrams = primerGrams * (config.PrimerThinnerPercent / 100.0);
                double primerThinnerKg = primerThinnerGrams / 1000.0;
                
                writer.WriteLine($"  Primer: {primerGrams:F1} g ({primerKg:F3} kg) - Fr. {objCosts.PrimerCost:F2}");
                writer.WriteLine($"    + Hardener ({config.PrimerHardenerPercent:F1}%): {primerHardenerGrams:F1} g ({primerHardenerKg:F3} kg) - Fr. {objCosts.PrimerHardenerCost:F2}");
                writer.WriteLine($"    + Thinner ({config.PrimerThinnerPercent:F1}%): {primerThinnerGrams:F1} g ({primerThinnerKg:F3} kg) - Fr. {objCosts.PrimerThinnerCost:F2}");
            }
            
            if (config.Topcoat != null)
            {
                double topcoatGrams = objArea_m2 * config.Topcoat.ConsumptionPerSqM;
                double topcoatKg = topcoatGrams / 1000.0;
                double topcoatHardenerGrams = topcoatGrams * (config.TopcoatHardenerPercent / 100.0);
                double topcoatHardenerKg = topcoatHardenerGrams / 1000.0;
                double topcoatThinnerGrams = topcoatGrams * (config.TopcoatThinnerPercent / 100.0);
                double topcoatThinnerKg = topcoatThinnerGrams / 1000.0;
                
                writer.WriteLine($"  Topcoat: {topcoatGrams:F1} g ({topcoatKg:F3} kg) - Fr. {objCosts.TopcoatCost:F2}");
                writer.WriteLine($"    + Hardener ({config.TopcoatHardenerPercent:F1}%): {topcoatHardenerGrams:F1} g ({topcoatHardenerKg:F3} kg) - Fr. {objCosts.TopcoatHardenerCost:F2}");
                writer.WriteLine($"    + Thinner ({config.TopcoatThinnerPercent:F1}%): {topcoatThinnerGrams:F1} g ({topcoatThinnerKg:F3} kg) - Fr. {objCosts.TopcoatThinnerCost:F2}");
            }
        }

        private static void WriteSummary(StreamWriter writer, double surfaceArea_mm2, int objectCount, 
                                         MaterialConfig config, CostCalculation costs, double estimatedTime)
        {
            writer.WriteLine("SUMMARY:");
            writer.WriteLine();
            writer.WriteLine($"Total Objects: {objectCount}");
            writer.WriteLine($"Total Surface Area: {surfaceArea_mm2:F2} mm² ({surfaceArea_mm2 / 1000000:F4} m²)");
            writer.WriteLine();

            double totalArea_m2 = surfaceArea_mm2 / 1000000.0;

            if (config.Primer != null)
            {
                double totalPrimerGrams = totalArea_m2 * config.Primer.ConsumptionPerSqM;
                double totalPrimerKg = totalPrimerGrams / 1000.0;
                double totalPrimerHardenerGrams = totalPrimerGrams * (config.PrimerHardenerPercent / 100.0);
                double totalPrimerHardenerKg = totalPrimerHardenerGrams / 1000.0;
                double totalPrimerThinnerGrams = totalPrimerGrams * (config.PrimerThinnerPercent / 100.0);
                double totalPrimerThinnerKg = totalPrimerThinnerGrams / 1000.0;
                
                writer.WriteLine($"Primer: {config.Primer.Name}");
                writer.WriteLine($"  Consumption: {config.Primer.ConsumptionPerSqM:F1} g/m²");
                writer.WriteLine($"  Total Amount: {totalPrimerGrams:F1} g ({totalPrimerKg:F3} kg)");
                writer.WriteLine($"  Price: Fr. {config.Primer.PricePerKg:F2}/kg");
                writer.WriteLine($"  Total Cost: Fr. {costs.PrimerCost:F2}");
                writer.WriteLine($"  Hardener ({config.PrimerHardenerPercent:F1}%): {totalPrimerHardenerGrams:F1} g ({totalPrimerHardenerKg:F3} kg) - Fr. {costs.PrimerHardenerCost:F2}");
                writer.WriteLine($"  Thinner ({config.PrimerThinnerPercent:F1}%): {totalPrimerThinnerGrams:F1} g ({totalPrimerThinnerKg:F3} kg) - Fr. {costs.PrimerThinnerCost:F2}");
                writer.WriteLine();
            }

            if (config.Topcoat != null)
            {
                double totalTopcoatGrams = totalArea_m2 * config.Topcoat.ConsumptionPerSqM;
                double totalTopcoatKg = totalTopcoatGrams / 1000.0;
                double totalTopcoatHardenerGrams = totalTopcoatGrams * (config.TopcoatHardenerPercent / 100.0);
                double totalTopcoatHardenerKg = totalTopcoatHardenerGrams / 1000.0;
                double totalTopcoatThinnerGrams = totalTopcoatGrams * (config.TopcoatThinnerPercent / 100.0);
                double totalTopcoatThinnerKg = totalTopcoatThinnerGrams / 1000.0;
                
                writer.WriteLine($"Topcoat: {config.Topcoat.Name}");
                writer.WriteLine($"  Consumption: {config.Topcoat.ConsumptionPerSqM:F1} g/m²");
                writer.WriteLine($"  Total Amount: {totalTopcoatGrams:F1} g ({totalTopcoatKg:F3} kg)");
                writer.WriteLine($"  Price: Fr. {config.Topcoat.PricePerKg:F2}/kg");
                writer.WriteLine($"  Total Cost: Fr. {costs.TopcoatCost:F2}");
                writer.WriteLine($"  Hardener ({config.TopcoatHardenerPercent:F1}%): {totalTopcoatHardenerGrams:F1} g ({totalTopcoatHardenerKg:F3} kg) - Fr. {costs.TopcoatHardenerCost:F2}");
                writer.WriteLine($"  Thinner ({config.TopcoatThinnerPercent:F1}%): {totalTopcoatThinnerGrams:F1} g ({totalTopcoatThinnerKg:F3} kg) - Fr. {costs.TopcoatThinnerCost:F2}");
                writer.WriteLine();
            }

            writer.WriteLine($"Total Material Cost: Fr. {costs.TotalMaterialCost:F2}");
            writer.WriteLine($"Total Estimated Time: {estimatedTime:F2} hours");
            writer.WriteLine($"Time Factor: {config.TimeFactor:F2} h/m²");
        }

        #endregion

        #region CSV Export

        private static void ExportToCSV(string filename, double surfaceArea_mm2, List<ObjectInfo> selectedObjects, 
                                        MaterialConfig config, CostCalculation costs, double estimatedTime)
        {
            using (var writer = new StreamWriter(filename))
            {
                // Header
                writer.WriteLine("Rhino CoatingApp Calculation Results");
                writer.WriteLine($"Date,{DateTime.Now}");
                writer.WriteLine($"Application Type,{config.ApplicationType}");
                writer.WriteLine();

                // Individual objects table
                writer.WriteLine("INDIVIDUAL OBJECTS");
                writer.WriteLine("Object,Name,Surface Area (mm²),Surface Area (m²),Primer (g),Primer (kg),Primer Cost (Fr.),Primer Hardener (g),Primer Hardener (kg),Primer Hardener Cost (Fr.),Primer Thinner (g),Primer Thinner (kg),Primer Thinner Cost (Fr.),Topcoat (g),Topcoat (kg),Topcoat Cost (Fr.),Topcoat Hardener (g),Topcoat Hardener (kg),Topcoat Hardener Cost (Fr.),Topcoat Thinner (g),Topcoat Thinner (kg),Topcoat Thinner Cost (Fr.),Total Cost (Fr.),Time (hours)");
                
                if (selectedObjects != null && selectedObjects.Count > 0)
                {
                    for (int i = 0; i < selectedObjects.Count; i++)
                    {
                        var obj = selectedObjects[i];
                        double objArea_m2 = obj.SurfaceArea_mm2 / 1000000.0;
                        var objCosts = CalculateIndividualCosts(obj.SurfaceArea_mm2, config);
                        var objTime = EstimateIndividualTime(obj.SurfaceArea_mm2, config.TimeFactor);
                        
                        var amounts = CalculateMaterialAmounts(objArea_m2, config);
                        
                        writer.WriteLine($"{i + 1},\"{obj.Name}\",{obj.SurfaceArea_mm2:F2},{objArea_m2:F4}," +
                                       $"{amounts.PrimerGrams:F1},{amounts.PrimerKg:F3},{objCosts.PrimerCost:F2}," +
                                       $"{amounts.PrimerHardenerGrams:F1},{amounts.PrimerHardenerKg:F3},{objCosts.PrimerHardenerCost:F2}," +
                                       $"{amounts.PrimerThinnerGrams:F1},{amounts.PrimerThinnerKg:F3},{objCosts.PrimerThinnerCost:F2}," +
                                       $"{amounts.TopcoatGrams:F1},{amounts.TopcoatKg:F3},{objCosts.TopcoatCost:F2}," +
                                       $"{amounts.TopcoatHardenerGrams:F1},{amounts.TopcoatHardenerKg:F3},{objCosts.TopcoatHardenerCost:F2}," +
                                       $"{amounts.TopcoatThinnerGrams:F1},{amounts.TopcoatThinnerKg:F3},{objCosts.TopcoatThinnerCost:F2}," +
                                       $"{objCosts.TotalMaterialCost:F2},{objTime:F2}");
                    }
                }
                
                writer.WriteLine();
                WriteCSVSummary(writer, surfaceArea_mm2, selectedObjects.Count, config, costs, estimatedTime);
            }
        }

        private static void WriteCSVSummary(StreamWriter writer, double surfaceArea_mm2, int objectCount,
                                            MaterialConfig config, CostCalculation costs, double estimatedTime)
        {
            writer.WriteLine("SUMMARY");
            writer.WriteLine($"Total Objects,{objectCount}");
            writer.WriteLine($"Total Surface Area (mm²),{surfaceArea_mm2:F2}");
            writer.WriteLine($"Total Surface Area (m²),{surfaceArea_mm2 / 1000000:F4}");
            writer.WriteLine();
            
            double totalArea_m2 = surfaceArea_mm2 / 1000000.0;
            var totalAmounts = CalculateMaterialAmounts(totalArea_m2, config);
            
            if (config.Primer != null)
            {
                writer.WriteLine($"Primer,{config.Primer.Name}");
                writer.WriteLine($"Primer Consumption (g/m²),{config.Primer.ConsumptionPerSqM:F1}");
                writer.WriteLine($"Primer Total Amount (g),{totalAmounts.PrimerGrams:F1}");
                writer.WriteLine($"Primer Total Amount (kg),{totalAmounts.PrimerKg:F3}");
                writer.WriteLine($"Primer Price (Fr./kg),{config.Primer.PricePerKg:F2}");
                writer.WriteLine($"Primer Total Cost (Fr.),{costs.PrimerCost:F2}");
                writer.WriteLine($"Primer Hardener ({config.PrimerHardenerPercent:F1}%),{totalAmounts.PrimerHardenerGrams:F1} g ({totalAmounts.PrimerHardenerKg:F3} kg) - Fr. {costs.PrimerHardenerCost:F2}");
                writer.WriteLine($"Primer Thinner ({config.PrimerThinnerPercent:F1}%),{totalAmounts.PrimerThinnerGrams:F1} g ({totalAmounts.PrimerThinnerKg:F3} kg) - Fr. {costs.PrimerThinnerCost:F2}");
            }
            
            if (config.Topcoat != null)
            {
                writer.WriteLine($"Topcoat,{config.Topcoat.Name}");
                writer.WriteLine($"Topcoat Consumption (g/m²),{config.Topcoat.ConsumptionPerSqM:F1}");
                writer.WriteLine($"Topcoat Total Amount (g),{totalAmounts.TopcoatGrams:F1}");
                writer.WriteLine($"Topcoat Total Amount (kg),{totalAmounts.TopcoatKg:F3}");
                writer.WriteLine($"Topcoat Price (Fr./kg),{config.Topcoat.PricePerKg:F2}");
                writer.WriteLine($"Topcoat Total Cost (Fr.),{costs.TopcoatCost:F2}");
                writer.WriteLine($"Topcoat Hardener ({config.TopcoatHardenerPercent:F1}%),{totalAmounts.TopcoatHardenerGrams:F1} g ({totalAmounts.TopcoatHardenerKg:F3} kg) - Fr. {costs.TopcoatHardenerCost:F2}");
                writer.WriteLine($"Topcoat Thinner ({config.TopcoatThinnerPercent:F1}%),{totalAmounts.TopcoatThinnerGrams:F1} g ({totalAmounts.TopcoatThinnerKg:F3} kg) - Fr. {costs.TopcoatThinnerCost:F2}");
            }
            
            writer.WriteLine($"Total Material Cost (Fr.),{costs.TotalMaterialCost:F2}");
            writer.WriteLine($"Total Estimated Time (hours),{estimatedTime:F2}");
            writer.WriteLine($"Time Factor (h/m²),{config.TimeFactor:F2}");
        }

        #endregion

        #region Excel Export

        private static void ExportToExcel(string filename, double surfaceArea_mm2, List<ObjectInfo> selectedObjects, 
                                          MaterialConfig config, CostCalculation costs, double estimatedTime)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Coating Calculation");
                int row = 1;

                // Header
                worksheet.Cell(row, 1).Value = "Rhino CoatingApp Calculation Results";
                worksheet.Cell(row, 1).Style.Font.Bold = true;
                worksheet.Cell(row, 1).Style.Font.FontSize = 14;
                row += 2;
                
                worksheet.Cell(row, 1).Value = "Date:";
                worksheet.Cell(row, 2).Value = DateTime.Now.ToString();
                row += 1;
                
                worksheet.Cell(row, 1).Value = "Application Type:";
                worksheet.Cell(row, 2).Value = config.ApplicationType.ToString();
                row += 2;

                // Individual objects table
                row = WriteExcelObjectsTable(worksheet, row, selectedObjects, config);
                
                // Summary
                row = WriteExcelSummary(worksheet, row, surfaceArea_mm2, selectedObjects.Count, config, costs, estimatedTime);

                // Auto-fit columns
                worksheet.Columns().AdjustToContents();
                
                // Save workbook
                workbook.SaveAs(filename);
            }
        }

        private static int WriteExcelObjectsTable(IXLWorksheet worksheet, int row, List<ObjectInfo> selectedObjects, MaterialConfig config)
        {
            worksheet.Cell(row, 1).Value = "INDIVIDUAL OBJECTS";
            worksheet.Cell(row, 1).Style.Font.Bold = true;
            worksheet.Cell(row, 1).Style.Font.FontSize = 12;
            row += 1;
            
            // Table header
            string[] headers = new[] { 
                "Object", "Name", "Surface Area (mm²)", "Surface Area (m²)",
                "Primer (g)", "Primer (kg)", "Primer Cost (Fr.)",
                "Primer Hardener (g)", "Primer Hardener (kg)", "Primer Hardener Cost (Fr.)",
                "Primer Thinner (g)", "Primer Thinner (kg)", "Primer Thinner Cost (Fr.)",
                "Topcoat (g)", "Topcoat (kg)", "Topcoat Cost (Fr.)",
                "Topcoat Hardener (g)", "Topcoat Hardener (kg)", "Topcoat Hardener Cost (Fr.)",
                "Topcoat Thinner (g)", "Topcoat Thinner (kg)", "Topcoat Thinner Cost (Fr.)",
                "Total Cost (Fr.)", "Time (hours)"
            };
            
            for (int col = 0; col < headers.Length; col++)
            {
                worksheet.Cell(row, col + 1).Value = headers[col];
            }
            worksheet.Range(row, 1, row, headers.Length).Style.Font.Bold = true;
            worksheet.Range(row, 1, row, headers.Length).Style.Fill.BackgroundColor = XLColor.LightGray;
            row += 1;
            
            // Individual objects data
            if (selectedObjects != null && selectedObjects.Count > 0)
            {
                for (int i = 0; i < selectedObjects.Count; i++)
                {
                    var obj = selectedObjects[i];
                    double objArea_m2 = obj.SurfaceArea_mm2 / 1000000.0;
                    var objCosts = CalculateIndividualCosts(obj.SurfaceArea_mm2, config);
                    var objTime = EstimateIndividualTime(obj.SurfaceArea_mm2, config.TimeFactor);
                    var amounts = CalculateMaterialAmounts(objArea_m2, config);
                    
                    int col = 1;
                    worksheet.Cell(row, col++).Value = i + 1;
                    worksheet.Cell(row, col++).Value = obj.Name;
                    worksheet.Cell(row, col++).Value = obj.SurfaceArea_mm2;
                    worksheet.Cell(row, col++).Value = objArea_m2;
                    worksheet.Cell(row, col++).Value = amounts.PrimerGrams;
                    worksheet.Cell(row, col++).Value = amounts.PrimerKg;
                    worksheet.Cell(row, col++).Value = objCosts.PrimerCost;
                    worksheet.Cell(row, col++).Value = amounts.PrimerHardenerGrams;
                    worksheet.Cell(row, col++).Value = amounts.PrimerHardenerKg;
                    worksheet.Cell(row, col++).Value = objCosts.PrimerHardenerCost;
                    worksheet.Cell(row, col++).Value = amounts.PrimerThinnerGrams;
                    worksheet.Cell(row, col++).Value = amounts.PrimerThinnerKg;
                    worksheet.Cell(row, col++).Value = objCosts.PrimerThinnerCost;
                    worksheet.Cell(row, col++).Value = amounts.TopcoatGrams;
                    worksheet.Cell(row, col++).Value = amounts.TopcoatKg;
                    worksheet.Cell(row, col++).Value = objCosts.TopcoatCost;
                    worksheet.Cell(row, col++).Value = amounts.TopcoatHardenerGrams;
                    worksheet.Cell(row, col++).Value = amounts.TopcoatHardenerKg;
                    worksheet.Cell(row, col++).Value = objCosts.TopcoatHardenerCost;
                    worksheet.Cell(row, col++).Value = amounts.TopcoatThinnerGrams;
                    worksheet.Cell(row, col++).Value = amounts.TopcoatThinnerKg;
                    worksheet.Cell(row, col++).Value = objCosts.TopcoatThinnerCost;
                    worksheet.Cell(row, col++).Value = objCosts.TotalMaterialCost;
                    worksheet.Cell(row, col++).Value = objTime;
                    
                    row += 1;
                }
            }
            
            return row + 2;
        }

        private static int WriteExcelSummary(IXLWorksheet worksheet, int row, double surfaceArea_mm2, int objectCount,
                                             MaterialConfig config, CostCalculation costs, double estimatedTime)
        {
            worksheet.Cell(row, 1).Value = "SUMMARY";
            worksheet.Cell(row, 1).Style.Font.Bold = true;
            worksheet.Cell(row, 1).Style.Font.FontSize = 12;
            row += 1;
            
            worksheet.Cell(row, 1).Value = "Total Objects:";
            worksheet.Cell(row, 2).Value = objectCount;
            row += 1;
            
            worksheet.Cell(row, 1).Value = "Total Surface Area (mm²):";
            worksheet.Cell(row, 2).Value = surfaceArea_mm2;
            row += 1;
            
            worksheet.Cell(row, 1).Value = "Total Surface Area (m²):";
            worksheet.Cell(row, 2).Value = surfaceArea_mm2 / 1000000;
            row += 2;
            
            double totalArea_m2 = surfaceArea_mm2 / 1000000.0;
            var totalAmounts = CalculateMaterialAmounts(totalArea_m2, config);
            
            if (config.Primer != null)
            {
                worksheet.Cell(row, 1).Value = "Primer:";
                worksheet.Cell(row, 2).Value = config.Primer.Name;
                row += 1;
                worksheet.Cell(row, 1).Value = "  Consumption (g/m²):";
                worksheet.Cell(row, 2).Value = config.Primer.ConsumptionPerSqM;
                row += 1;
                worksheet.Cell(row, 1).Value = "  Total Amount (g):";
                worksheet.Cell(row, 2).Value = totalAmounts.PrimerGrams;
                row += 1;
                worksheet.Cell(row, 1).Value = "  Total Amount (kg):";
                worksheet.Cell(row, 2).Value = totalAmounts.PrimerKg;
                row += 1;
                worksheet.Cell(row, 1).Value = "  Price (Fr./kg):";
                worksheet.Cell(row, 2).Value = config.Primer.PricePerKg;
                row += 1;
                worksheet.Cell(row, 1).Value = "  Total Cost (Fr.):";
                worksheet.Cell(row, 2).Value = costs.PrimerCost;
                row += 1;
                worksheet.Cell(row, 1).Value = $"  Hardener ({config.PrimerHardenerPercent:F1}%):";
                worksheet.Cell(row, 2).Value = $"{totalAmounts.PrimerHardenerGrams:F1} g ({totalAmounts.PrimerHardenerKg:F3} kg) - Fr. {costs.PrimerHardenerCost:F2}";
                row += 1;
                worksheet.Cell(row, 1).Value = $"  Thinner ({config.PrimerThinnerPercent:F1}%):";
                worksheet.Cell(row, 2).Value = $"{totalAmounts.PrimerThinnerGrams:F1} g ({totalAmounts.PrimerThinnerKg:F3} kg) - Fr. {costs.PrimerThinnerCost:F2}";
                row += 2;
            }
            
            if (config.Topcoat != null)
            {
                worksheet.Cell(row, 1).Value = "Topcoat:";
                worksheet.Cell(row, 2).Value = config.Topcoat.Name;
                row += 1;
                worksheet.Cell(row, 1).Value = "  Consumption (g/m²):";
                worksheet.Cell(row, 2).Value = config.Topcoat.ConsumptionPerSqM;
                row += 1;
                worksheet.Cell(row, 1).Value = "  Total Amount (g):";
                worksheet.Cell(row, 2).Value = totalAmounts.TopcoatGrams;
                row += 1;
                worksheet.Cell(row, 1).Value = "  Total Amount (kg):";
                worksheet.Cell(row, 2).Value = totalAmounts.TopcoatKg;
                row += 1;
                worksheet.Cell(row, 1).Value = "  Price (Fr./kg):";
                worksheet.Cell(row, 2).Value = config.Topcoat.PricePerKg;
                row += 1;
                worksheet.Cell(row, 1).Value = "  Total Cost (Fr.):";
                worksheet.Cell(row, 2).Value = costs.TopcoatCost;
                row += 1;
                worksheet.Cell(row, 1).Value = $"  Hardener ({config.TopcoatHardenerPercent:F1}%):";
                worksheet.Cell(row, 2).Value = $"{totalAmounts.TopcoatHardenerGrams:F1} g ({totalAmounts.TopcoatHardenerKg:F3} kg) - Fr. {costs.TopcoatHardenerCost:F2}";
                row += 1;
                worksheet.Cell(row, 1).Value = $"  Thinner ({config.TopcoatThinnerPercent:F1}%):";
                worksheet.Cell(row, 2).Value = $"{totalAmounts.TopcoatThinnerGrams:F1} g ({totalAmounts.TopcoatThinnerKg:F3} kg) - Fr. {costs.TopcoatThinnerCost:F2}";
                row += 2;
            }
            
            worksheet.Cell(row, 1).Value = "Total Material Cost (Fr.):";
            worksheet.Cell(row, 2).Value = costs.TotalMaterialCost;
            worksheet.Cell(row, 1).Style.Font.Bold = true;
            row += 1;
            
            worksheet.Cell(row, 1).Value = "Total Estimated Time (hours):";
            worksheet.Cell(row, 2).Value = estimatedTime;
            row += 1;
            
            worksheet.Cell(row, 1).Value = "Time Factor (h/m²):";
            worksheet.Cell(row, 2).Value = config.TimeFactor;
            
            return row;
        }

        #endregion

        #region Helper Methods

        private static CostCalculation CalculateIndividualCosts(double surfaceArea_mm2, MaterialConfig config)
        {
            CostCalculation costs = new CostCalculation();
            double surfaceArea_m2 = surfaceArea_mm2 / 1000000.0;

            if (config.Primer != null)
            {
                double primerGrams = surfaceArea_m2 * config.Primer.ConsumptionPerSqM;
                double primerKg = primerGrams / 1000.0;
                costs.PrimerCost = primerKg * config.Primer.PricePerKg;
                
                double primerHardenerGrams = primerGrams * (config.PrimerHardenerPercent / 100.0);
                double primerHardenerKg = primerHardenerGrams / 1000.0;
                costs.PrimerHardenerCost = primerHardenerKg * config.HardenerPricePerKg;
                
                double primerThinnerGrams = primerGrams * (config.PrimerThinnerPercent / 100.0);
                double primerThinnerKg = primerThinnerGrams / 1000.0;
                costs.PrimerThinnerCost = primerThinnerKg * config.ThinnerPricePerKg;
                
                costs.TotalMaterialCost += costs.PrimerCost + costs.PrimerHardenerCost + costs.PrimerThinnerCost;
            }

            if (config.Topcoat != null)
            {
                double topcoatGrams = surfaceArea_m2 * config.Topcoat.ConsumptionPerSqM;
                double topcoatKg = topcoatGrams / 1000.0;
                costs.TopcoatCost = topcoatKg * config.Topcoat.PricePerKg;
                
                double topcoatHardenerGrams = topcoatGrams * (config.TopcoatHardenerPercent / 100.0);
                double topcoatHardenerKg = topcoatHardenerGrams / 1000.0;
                costs.TopcoatHardenerCost = topcoatHardenerKg * config.HardenerPricePerKg;
                
                double topcoatThinnerGrams = topcoatGrams * (config.TopcoatThinnerPercent / 100.0);
                double topcoatThinnerKg = topcoatThinnerGrams / 1000.0;
                costs.TopcoatThinnerCost = topcoatThinnerKg * config.ThinnerPricePerKg;
                
                costs.TotalMaterialCost += costs.TopcoatCost + costs.TopcoatHardenerCost + costs.TopcoatThinnerCost;
            }

            return costs;
        }

        private static double EstimateIndividualTime(double surfaceArea_mm2, double timeFactor)
        {
            double surfaceArea_m2 = surfaceArea_mm2 / 1000000.0;
            return surfaceArea_m2 * timeFactor;
        }

        private static MaterialAmounts CalculateMaterialAmounts(double area_m2, MaterialConfig config)
        {
            var amounts = new MaterialAmounts();
            
            if (config.Primer != null)
            {
                amounts.PrimerGrams = area_m2 * config.Primer.ConsumptionPerSqM;
                amounts.PrimerKg = amounts.PrimerGrams / 1000.0;
                amounts.PrimerHardenerGrams = amounts.PrimerGrams * (config.PrimerHardenerPercent / 100.0);
                amounts.PrimerHardenerKg = amounts.PrimerHardenerGrams / 1000.0;
                amounts.PrimerThinnerGrams = amounts.PrimerGrams * (config.PrimerThinnerPercent / 100.0);
                amounts.PrimerThinnerKg = amounts.PrimerThinnerGrams / 1000.0;
            }
            
            if (config.Topcoat != null)
            {
                amounts.TopcoatGrams = area_m2 * config.Topcoat.ConsumptionPerSqM;
                amounts.TopcoatKg = amounts.TopcoatGrams / 1000.0;
                amounts.TopcoatHardenerGrams = amounts.TopcoatGrams * (config.TopcoatHardenerPercent / 100.0);
                amounts.TopcoatHardenerKg = amounts.TopcoatHardenerGrams / 1000.0;
                amounts.TopcoatThinnerGrams = amounts.TopcoatGrams * (config.TopcoatThinnerPercent / 100.0);
                amounts.TopcoatThinnerKg = amounts.TopcoatThinnerGrams / 1000.0;
            }
            
            return amounts;
        }

        #endregion
    }

    #region Helper Classes

    internal class MaterialAmounts
    {
        public double PrimerGrams { get; set; }
        public double PrimerKg { get; set; }
        public double PrimerHardenerGrams { get; set; }
        public double PrimerHardenerKg { get; set; }
        public double PrimerThinnerGrams { get; set; }
        public double PrimerThinnerKg { get; set; }
        
        public double TopcoatGrams { get; set; }
        public double TopcoatKg { get; set; }
        public double TopcoatHardenerGrams { get; set; }
        public double TopcoatHardenerKg { get; set; }
        public double TopcoatThinnerGrams { get; set; }
        public double TopcoatThinnerKg { get; set; }
    }

    #endregion
}

