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
        // Shorthand references to config and labels for cleaner code
        private static UnitsAndLabelsConfigData Config => UnitsAndLabelsConfigManager.Config;
        private static LabelsConfig L => Config.Labels;
        private static UnitsConfig U => Config.Units;
        /// <summary>
        /// Exports coating calculation results to a file in the specified format
        /// </summary>
        /// <param name="filename">Full path to the output file</param>
        /// <param name="surfaceArea_mm2">Total surface area in mm²</param>
        /// <param name="selectedObjects">List of selected objects with their details</param>
        /// <param name="config">Material configuration used for calculation</param>
        /// <param name="costs">Calculated costs</param>
        /// <param name="estimatedTime">Estimated time for the coating job</param>
        /// <param name="calculationResult">Calculation result with final offer price</param>
        /// <param name="calculationFactors">Current calculation factors as dictionary</param>
        public static void Export(string filename, double surfaceArea_mm2, List<ObjectInfo> selectedObjects,
                                  MaterialConfig config, CostCalculation costs, double estimatedTime,
                                  CalculationResult calculationResult = null, Dictionary<string, double> calculationFactors = null)
        {
            string extension = Path.GetExtension(filename).ToLower();
            
            switch (extension)
            {
                case ".xlsx":
                    ExportToExcel(filename, surfaceArea_mm2, selectedObjects, config, costs, estimatedTime, calculationResult, calculationFactors);
                    break;
                case ".csv":
                    ExportToCSV(filename, surfaceArea_mm2, selectedObjects, config, costs, estimatedTime, calculationResult, calculationFactors);
                    break;
                case ".txt":
                default:
                    ExportToText(filename, surfaceArea_mm2, selectedObjects, config, costs, estimatedTime, calculationResult, calculationFactors);
                    break;
            }
        }

        #region Text Export

        private static void ExportToText(string filename, double surfaceArea_mm2, List<ObjectInfo> selectedObjects,
                                         MaterialConfig config, CostCalculation costs, double estimatedTime,
                                         CalculationResult calculationResult = null, Dictionary<string, double> calculationFactors = null)
        {
            using (var writer = new StreamWriter(filename))
            {
                writer.WriteLine(L.Sections.CalculationResults);
                writer.WriteLine($"{L.General.Date}: {DateTime.Now}");
                writer.WriteLine($"{L.General.ApplicationType}: {config.ApplicationType}");
                writer.WriteLine();

                // Export individual objects
                if (selectedObjects != null && selectedObjects.Count > 0)
                {
                    writer.WriteLine($"{L.Sections.IndividualObjects}:");
                    writer.WriteLine();
                    
                    for (int i = 0; i < selectedObjects.Count; i++)
                    {
                        var obj = selectedObjects[i];
                        writer.WriteLine($"{L.General.Object} {i + 1}: {obj.Name}");
                        writer.WriteLine($"  {L.Properties.SurfaceArea}: {UnitsAndLabelsConfigManager.FormatAreaBoth(obj.SurfaceArea_mm2)}");
                        
                        double objArea_m2 = obj.SurfaceArea_mm2 / UnitsAndLabelsConfigManager.GetAreaConversionFactor();
                        var objCosts = CalculateIndividualCosts(obj.SurfaceArea_mm2, config);
                        var objTime = EstimateIndividualTime(obj.SurfaceArea_mm2, config.TimeFactor);
                        
                        WriteObjectMaterialDetails(writer, objArea_m2, config, objCosts);
                        writer.WriteLine($"  {L.Costs.TotalMaterialCost}: {UnitsAndLabelsConfigManager.FormatCurrency(objCosts.TotalMaterialCost)}");
                        writer.WriteLine($"  {L.Time.EstimatedTime}: {UnitsAndLabelsConfigManager.FormatTime(objTime)}");
                        writer.WriteLine();
                    }

                    writer.WriteLine("─────────────────────────────────");
                    writer.WriteLine();
                }

                // Export summary
                WriteSummary(writer, surfaceArea_mm2, selectedObjects.Count, config, costs, estimatedTime, calculationResult, calculationFactors);
            }
        }

        private static void WriteObjectMaterialDetails(StreamWriter writer, double objArea_m2, MaterialConfig config, CostCalculation objCosts)
        {
            if (config.Primer != null)
            {
                double primerGrams = objArea_m2 * config.Primer.ConsumptionPerSqM * config.PrimerCoatMultiplier;
                double primerHardenerGrams = primerGrams * (config.PrimerHardenerPercent / 100.0);
                double primerThinnerGrams = primerGrams * (config.PrimerThinnerPercent / 100.0);

                writer.WriteLine($"  {L.Materials.Primer} ({config.PrimerCoatMultiplier:F0}x {L.Properties.Coats.ToLower()}): {UnitsAndLabelsConfigManager.FormatWeightBoth(primerGrams)} - {UnitsAndLabelsConfigManager.FormatCurrency(objCosts.PrimerCost)}");
                writer.WriteLine($"    + {L.Materials.Hardener} ({UnitsAndLabelsConfigManager.FormatPercentage(config.PrimerHardenerPercent)}): {UnitsAndLabelsConfigManager.FormatWeightBoth(primerHardenerGrams)} - {UnitsAndLabelsConfigManager.FormatCurrency(objCosts.PrimerHardenerCost)}");
                writer.WriteLine($"    + {L.Materials.Thinner} ({UnitsAndLabelsConfigManager.FormatPercentage(config.PrimerThinnerPercent)}): {UnitsAndLabelsConfigManager.FormatWeightBoth(primerThinnerGrams)} - {UnitsAndLabelsConfigManager.FormatCurrency(objCosts.PrimerThinnerCost)}");
            }

            if (config.Topcoat != null)
            {
                double topcoatGrams = objArea_m2 * config.Topcoat.ConsumptionPerSqM * config.TopcoatCoatMultiplier;
                double topcoatHardenerGrams = topcoatGrams * (config.TopcoatHardenerPercent / 100.0);
                double topcoatThinnerGrams = topcoatGrams * (config.TopcoatThinnerPercent / 100.0);

                writer.WriteLine($"  {L.Materials.Topcoat} ({config.TopcoatCoatMultiplier:F0}x {L.Properties.Coats.ToLower()}): {UnitsAndLabelsConfigManager.FormatWeightBoth(topcoatGrams)} - {UnitsAndLabelsConfigManager.FormatCurrency(objCosts.TopcoatCost)}");
                writer.WriteLine($"    + {L.Materials.Hardener} ({UnitsAndLabelsConfigManager.FormatPercentage(config.TopcoatHardenerPercent)}): {UnitsAndLabelsConfigManager.FormatWeightBoth(topcoatHardenerGrams)} - {UnitsAndLabelsConfigManager.FormatCurrency(objCosts.TopcoatHardenerCost)}");
                writer.WriteLine($"    + {L.Materials.Thinner} ({UnitsAndLabelsConfigManager.FormatPercentage(config.TopcoatThinnerPercent)}): {UnitsAndLabelsConfigManager.FormatWeightBoth(topcoatThinnerGrams)} - {UnitsAndLabelsConfigManager.FormatCurrency(objCosts.TopcoatThinnerCost)}");
            }
        }

        private static void WriteSummary(StreamWriter writer, double surfaceArea_mm2, int objectCount,
                                         MaterialConfig config, CostCalculation costs, double estimatedTime,
                                         CalculationResult calculationResult = null, Dictionary<string, double> calculationFactors = null)
        {
            writer.WriteLine($"{L.Sections.Summary}:");
            writer.WriteLine();
            writer.WriteLine($"{L.General.TotalObjects}: {objectCount}");
            writer.WriteLine($"{L.Properties.SurfaceArea}: {UnitsAndLabelsConfigManager.FormatAreaBoth(surfaceArea_mm2)}");
            writer.WriteLine();

            double totalArea_m2 = surfaceArea_mm2 / UnitsAndLabelsConfigManager.GetAreaConversionFactor();

            if (config.Primer != null)
            {
                double totalPrimerGrams = totalArea_m2 * config.Primer.ConsumptionPerSqM * config.PrimerCoatMultiplier;
                double totalPrimerHardenerGrams = totalPrimerGrams * (config.PrimerHardenerPercent / 100.0);
                double totalPrimerThinnerGrams = totalPrimerGrams * (config.PrimerThinnerPercent / 100.0);

                writer.WriteLine($"{L.Materials.Primer}: {config.Primer.Name}");
                writer.WriteLine($"  {L.Properties.Consumption}: {UnitsAndLabelsConfigManager.FormatWeightSmall(config.Primer.ConsumptionPerSqM)}/{U.Area.LargeUnit}");
                writer.WriteLine($"  {L.Properties.Coats}: {config.PrimerCoatMultiplier:F0}x");
                writer.WriteLine($"  {L.Properties.TotalAmount}: {UnitsAndLabelsConfigManager.FormatWeightBoth(totalPrimerGrams)}");
                writer.WriteLine($"  {L.Properties.Price}: {UnitsAndLabelsConfigManager.FormatCurrency(config.Primer.PricePerKg)}/{U.Weight.LargeUnit}");
                writer.WriteLine($"  {L.Properties.TotalCost}: {UnitsAndLabelsConfigManager.FormatCurrency(costs.PrimerCost)}");
                writer.WriteLine($"  {L.Materials.Hardener} ({UnitsAndLabelsConfigManager.FormatPercentage(config.PrimerHardenerPercent)}): {UnitsAndLabelsConfigManager.FormatWeightBoth(totalPrimerHardenerGrams)} - {UnitsAndLabelsConfigManager.FormatCurrency(costs.PrimerHardenerCost)}");
                writer.WriteLine($"  {L.Materials.Thinner} ({UnitsAndLabelsConfigManager.FormatPercentage(config.PrimerThinnerPercent)}): {UnitsAndLabelsConfigManager.FormatWeightBoth(totalPrimerThinnerGrams)} - {UnitsAndLabelsConfigManager.FormatCurrency(costs.PrimerThinnerCost)}");
                writer.WriteLine();
            }

            if (config.Topcoat != null)
            {
                double totalTopcoatGrams = totalArea_m2 * config.Topcoat.ConsumptionPerSqM * config.TopcoatCoatMultiplier;
                double totalTopcoatHardenerGrams = totalTopcoatGrams * (config.TopcoatHardenerPercent / 100.0);
                double totalTopcoatThinnerGrams = totalTopcoatGrams * (config.TopcoatThinnerPercent / 100.0);

                writer.WriteLine($"{L.Materials.Topcoat}: {config.Topcoat.Name}");
                writer.WriteLine($"  {L.Properties.Consumption}: {UnitsAndLabelsConfigManager.FormatWeightSmall(config.Topcoat.ConsumptionPerSqM)}/{U.Area.LargeUnit}");
                writer.WriteLine($"  {L.Properties.Coats}: {config.TopcoatCoatMultiplier:F0}x");
                writer.WriteLine($"  {L.Properties.TotalAmount}: {UnitsAndLabelsConfigManager.FormatWeightBoth(totalTopcoatGrams)}");
                writer.WriteLine($"  {L.Properties.Price}: {UnitsAndLabelsConfigManager.FormatCurrency(config.Topcoat.PricePerKg)}/{U.Weight.LargeUnit}");
                writer.WriteLine($"  {L.Properties.TotalCost}: {UnitsAndLabelsConfigManager.FormatCurrency(costs.TopcoatCost)}");
                writer.WriteLine($"  {L.Materials.Hardener} ({UnitsAndLabelsConfigManager.FormatPercentage(config.TopcoatHardenerPercent)}): {UnitsAndLabelsConfigManager.FormatWeightBoth(totalTopcoatHardenerGrams)} - {UnitsAndLabelsConfigManager.FormatCurrency(costs.TopcoatHardenerCost)}");
                writer.WriteLine($"  {L.Materials.Thinner} ({UnitsAndLabelsConfigManager.FormatPercentage(config.TopcoatThinnerPercent)}): {UnitsAndLabelsConfigManager.FormatWeightBoth(totalTopcoatThinnerGrams)} - {UnitsAndLabelsConfigManager.FormatCurrency(costs.TopcoatThinnerCost)}");
                writer.WriteLine();
            }

            writer.WriteLine($"{L.Costs.TotalMaterialCost}: {UnitsAndLabelsConfigManager.FormatCurrency(costs.TotalMaterialCost)}");
            writer.WriteLine($"{L.Time.TotalEstimatedTime}: {UnitsAndLabelsConfigManager.FormatTime(estimatedTime)}");
            writer.WriteLine($"{L.Time.TimeFactor}: {config.TimeFactor:F2} {U.Time.Unit}/{U.Area.LargeUnit}");

            // Add calculation breakdown if available
            if (calculationResult != null && calculationFactors != null)
            {
                writer.WriteLine();
                writer.WriteLine($"=== {L.Sections.OfferCalculation} ===");
                writer.WriteLine();
                writer.WriteLine($"{L.Costs.BaseMaterialCost}:     {UnitsAndLabelsConfigManager.FormatCurrency(calculationResult.BaseMaterialCost)}");

                if (calculationResult.MaterialSurcharge > 0 && calculationFactors.TryGetValue("MaterialCostFactor", out double materialFactor))
                    writer.WriteLine($"{L.Costs.MaterialSurcharge} ({UnitsAndLabelsConfigManager.FormatPercentage(materialFactor)}): {UnitsAndLabelsConfigManager.FormatCurrency(calculationResult.MaterialSurcharge)}");
                if (calculationResult.ProductionCost > 0 && calculationFactors.TryGetValue("ProductionCostFactor", out double productionFactor))
                    writer.WriteLine($"{L.Costs.ProductionCost} ({UnitsAndLabelsConfigManager.FormatPercentage(productionFactor)}):    {UnitsAndLabelsConfigManager.FormatCurrency(calculationResult.ProductionCost)}");
                if (calculationResult.AdministrationCost > 0 && calculationFactors.TryGetValue("AdministrationCostFactor", out double adminFactor))
                    writer.WriteLine($"{L.Costs.AdministrationCost} ({UnitsAndLabelsConfigManager.FormatPercentage(adminFactor)}): {UnitsAndLabelsConfigManager.FormatCurrency(calculationResult.AdministrationCost)}");
                if (calculationResult.SalesCost > 0 && calculationFactors.TryGetValue("SalesCostFactor", out double salesFactor))
                    writer.WriteLine($"{L.Costs.SalesCost} ({UnitsAndLabelsConfigManager.FormatPercentage(salesFactor)}):         {UnitsAndLabelsConfigManager.FormatCurrency(calculationResult.SalesCost)}");
                if (calculationResult.ProfitMargin > 0 && calculationFactors.TryGetValue("ProfitMarginFactor", out double profitFactor))
                    writer.WriteLine($"{L.Costs.ProfitMargin} ({UnitsAndLabelsConfigManager.FormatPercentage(profitFactor)}):      {UnitsAndLabelsConfigManager.FormatCurrency(calculationResult.ProfitMargin)}");

                writer.WriteLine();
                writer.WriteLine($"{L.Costs.FinalOfferPrice}: {UnitsAndLabelsConfigManager.FormatCurrency(calculationResult.FinalOfferPrice)}");
            }
        }

        #endregion

        #region CSV Export

        private static void ExportToCSV(string filename, double surfaceArea_mm2, List<ObjectInfo> selectedObjects,
                                        MaterialConfig config, CostCalculation costs, double estimatedTime,
                                        CalculationResult calculationResult = null, Dictionary<string, double> calculationFactors = null)
        {
            using (var writer = new StreamWriter(filename))
            {
                // Header
                writer.WriteLine(L.Sections.CalculationResults);
                writer.WriteLine($"{L.General.Date},{DateTime.Now}");
                writer.WriteLine($"{L.General.ApplicationType},{config.ApplicationType}");
                writer.WriteLine();

                // Individual objects table
                writer.WriteLine(L.Sections.IndividualObjects);
                writer.WriteLine($"{L.General.Object},{L.General.Name},{L.Properties.SurfaceArea} ({U.Area.SmallUnit}),{L.Properties.SurfaceArea} ({U.Area.LargeUnit})," +
                               $"{L.Materials.Primer} ({U.Weight.SmallUnit}),{L.Materials.Primer} ({U.Weight.LargeUnit}),{L.Materials.Primer} {L.Properties.Coats},{L.Materials.Primer} {L.Properties.TotalCost} ({U.Currency.Symbol})," +
                               $"{L.Materials.Primer} {L.Materials.Hardener} ({U.Weight.SmallUnit}),{L.Materials.Primer} {L.Materials.Hardener} ({U.Weight.LargeUnit}),{L.Materials.Primer} {L.Materials.Hardener} {L.Properties.TotalCost} ({U.Currency.Symbol})," +
                               $"{L.Materials.Primer} {L.Materials.Thinner} ({U.Weight.SmallUnit}),{L.Materials.Primer} {L.Materials.Thinner} ({U.Weight.LargeUnit}),{L.Materials.Primer} {L.Materials.Thinner} {L.Properties.TotalCost} ({U.Currency.Symbol})," +
                               $"{L.Materials.Topcoat} ({U.Weight.SmallUnit}),{L.Materials.Topcoat} ({U.Weight.LargeUnit}),{L.Materials.Topcoat} {L.Properties.Coats},{L.Materials.Topcoat} {L.Properties.TotalCost} ({U.Currency.Symbol})," +
                               $"{L.Materials.Topcoat} {L.Materials.Hardener} ({U.Weight.SmallUnit}),{L.Materials.Topcoat} {L.Materials.Hardener} ({U.Weight.LargeUnit}),{L.Materials.Topcoat} {L.Materials.Hardener} {L.Properties.TotalCost} ({U.Currency.Symbol})," +
                               $"{L.Materials.Topcoat} {L.Materials.Thinner} ({U.Weight.SmallUnit}),{L.Materials.Topcoat} {L.Materials.Thinner} ({U.Weight.LargeUnit}),{L.Materials.Topcoat} {L.Materials.Thinner} {L.Properties.TotalCost} ({U.Currency.Symbol})," +
                               $"{L.Costs.TotalMaterialCost} ({U.Currency.Symbol}),{L.Time.EstimatedTime} ({U.Time.UnitLong})");
                
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
                                       $"{amounts.PrimerGrams:F1},{amounts.PrimerKg:F3},{config.PrimerCoatMultiplier:F0},{objCosts.PrimerCost:F2}," +
                                       $"{amounts.PrimerHardenerGrams:F1},{amounts.PrimerHardenerKg:F3},{objCosts.PrimerHardenerCost:F2}," +
                                       $"{amounts.PrimerThinnerGrams:F1},{amounts.PrimerThinnerKg:F3},{objCosts.PrimerThinnerCost:F2}," +
                                       $"{amounts.TopcoatGrams:F1},{amounts.TopcoatKg:F3},{config.TopcoatCoatMultiplier:F0},{objCosts.TopcoatCost:F2}," +
                                       $"{amounts.TopcoatHardenerGrams:F1},{amounts.TopcoatHardenerKg:F3},{objCosts.TopcoatHardenerCost:F2}," +
                                       $"{amounts.TopcoatThinnerGrams:F1},{amounts.TopcoatThinnerKg:F3},{objCosts.TopcoatThinnerCost:F2}," +
                                       $"{objCosts.TotalMaterialCost:F2},{objTime:F2}");
                    }
                }
                
                writer.WriteLine();
                WriteCSVSummary(writer, surfaceArea_mm2, selectedObjects.Count, config, costs, estimatedTime, calculationResult, calculationFactors);
            }
        }

        private static void WriteCSVSummary(StreamWriter writer, double surfaceArea_mm2, int objectCount,
                                            MaterialConfig config, CostCalculation costs, double estimatedTime,
                                            CalculationResult calculationResult = null, Dictionary<string, double> calculationFactors = null)
        {
            writer.WriteLine(L.Sections.Summary);
            writer.WriteLine($"{L.General.TotalObjects},{objectCount}");
            writer.WriteLine($"{L.Properties.SurfaceArea} ({U.Area.SmallUnit}),{surfaceArea_mm2:F2}");
            writer.WriteLine($"{L.Properties.SurfaceArea} ({U.Area.LargeUnit}),{surfaceArea_mm2 / UnitsAndLabelsConfigManager.GetAreaConversionFactor():F4}");
            writer.WriteLine();
            
            double totalArea_m2 = surfaceArea_mm2 / UnitsAndLabelsConfigManager.GetAreaConversionFactor();
            var totalAmounts = CalculateMaterialAmounts(totalArea_m2, config);
            
            if (config.Primer != null)
            {
                writer.WriteLine($"Primer,{config.Primer.Name}");
                writer.WriteLine($"Primer Consumption (g/m²),{config.Primer.ConsumptionPerSqM:F1}");
                writer.WriteLine($"Primer Coats,{config.PrimerCoatMultiplier:F0}x");
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
                writer.WriteLine($"Topcoat Coats,{config.TopcoatCoatMultiplier:F0}x");
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

            // Add calculation breakdown if available
            if (calculationResult != null && calculationFactors != null)
            {
                writer.WriteLine();
                writer.WriteLine("OFFER CALCULATION");
                writer.WriteLine($"Base Material Cost (Fr.),{calculationResult.BaseMaterialCost:F2}");

                if (calculationResult.MaterialSurcharge > 0 && calculationFactors.TryGetValue("MaterialCostFactor", out double materialFactor))
                    writer.WriteLine($"Material Surcharge ({materialFactor:F1}%),{calculationResult.MaterialSurcharge:F2}");
                if (calculationResult.ProductionCost > 0 && calculationFactors.TryGetValue("ProductionCostFactor", out double productionFactor))
                    writer.WriteLine($"Production Cost ({productionFactor:F1}%),{calculationResult.ProductionCost:F2}");
                if (calculationResult.AdministrationCost > 0 && calculationFactors.TryGetValue("AdministrationCostFactor", out double adminFactor))
                    writer.WriteLine($"Administration Cost ({adminFactor:F1}%),{calculationResult.AdministrationCost:F2}");
                if (calculationResult.SalesCost > 0 && calculationFactors.TryGetValue("SalesCostFactor", out double salesFactor))
                    writer.WriteLine($"Sales Cost ({salesFactor:F1}%),{calculationResult.SalesCost:F2}");
                if (calculationResult.ProfitMargin > 0 && calculationFactors.TryGetValue("ProfitMarginFactor", out double profitFactor))
                    writer.WriteLine($"Profit Margin ({profitFactor:F1}%),{calculationResult.ProfitMargin:F2}");

                writer.WriteLine();
                writer.WriteLine($"Final Offer Price (Fr.),{calculationResult.FinalOfferPrice:F2}");
            }
        }

        #endregion

        #region Excel Export

        private static void ExportToExcel(string filename, double surfaceArea_mm2, List<ObjectInfo> selectedObjects,
                                          MaterialConfig config, CostCalculation costs, double estimatedTime,
                                          CalculationResult calculationResult = null, Dictionary<string, double> calculationFactors = null)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Coating Calculation");
                int row = 1;

                // Header
                worksheet.Cell(row, 1).Value = L.Sections.CalculationResults;
                worksheet.Cell(row, 1).Style.Font.Bold = true;
                worksheet.Cell(row, 1).Style.Font.FontSize = 14;
                row += 2;
                
                worksheet.Cell(row, 1).Value = $"{L.General.Date}:";
                worksheet.Cell(row, 2).Value = DateTime.Now.ToString();
                row += 1;
                
                worksheet.Cell(row, 1).Value = $"{L.General.ApplicationType}:";
                worksheet.Cell(row, 2).Value = config.ApplicationType.ToString();
                row += 2;

                // Individual objects table
                row = WriteExcelObjectsTable(worksheet, row, selectedObjects, config);
                
                // Summary
                row = WriteExcelSummary(worksheet, row, surfaceArea_mm2, selectedObjects.Count, config, costs, estimatedTime, calculationResult, calculationFactors);

                // Auto-fit columns
                worksheet.Columns().AdjustToContents();
                
                // Save workbook
                workbook.SaveAs(filename);
            }
        }

        private static int WriteExcelObjectsTable(IXLWorksheet worksheet, int row, List<ObjectInfo> selectedObjects, MaterialConfig config)
        {
            worksheet.Cell(row, 1).Value = L.Sections.IndividualObjects;
            worksheet.Cell(row, 1).Style.Font.Bold = true;
            worksheet.Cell(row, 1).Style.Font.FontSize = 12;
            row += 1;
            
            // Table header
            string[] headers = new[] {
                L.General.Object, L.General.Name, $"{L.Properties.SurfaceArea} ({U.Area.SmallUnit})", $"{L.Properties.SurfaceArea} ({U.Area.LargeUnit})",
                $"{L.Materials.Primer} ({U.Weight.SmallUnit})", $"{L.Materials.Primer} ({U.Weight.LargeUnit})", $"{L.Materials.Primer} {L.Properties.Coats}", $"{L.Materials.Primer} {L.Properties.TotalCost} ({U.Currency.Symbol})",
                $"{L.Materials.Primer} {L.Materials.Hardener} ({U.Weight.SmallUnit})", $"{L.Materials.Primer} {L.Materials.Hardener} ({U.Weight.LargeUnit})", $"{L.Materials.Primer} {L.Materials.Hardener} {L.Properties.TotalCost} ({U.Currency.Symbol})",
                $"{L.Materials.Primer} {L.Materials.Thinner} ({U.Weight.SmallUnit})", $"{L.Materials.Primer} {L.Materials.Thinner} ({U.Weight.LargeUnit})", $"{L.Materials.Primer} {L.Materials.Thinner} {L.Properties.TotalCost} ({U.Currency.Symbol})",
                $"{L.Materials.Topcoat} ({U.Weight.SmallUnit})", $"{L.Materials.Topcoat} ({U.Weight.LargeUnit})", $"{L.Materials.Topcoat} {L.Properties.Coats}", $"{L.Materials.Topcoat} {L.Properties.TotalCost} ({U.Currency.Symbol})",
                $"{L.Materials.Topcoat} {L.Materials.Hardener} ({U.Weight.SmallUnit})", $"{L.Materials.Topcoat} {L.Materials.Hardener} ({U.Weight.LargeUnit})", $"{L.Materials.Topcoat} {L.Materials.Hardener} {L.Properties.TotalCost} ({U.Currency.Symbol})",
                $"{L.Materials.Topcoat} {L.Materials.Thinner} ({U.Weight.SmallUnit})", $"{L.Materials.Topcoat} {L.Materials.Thinner} ({U.Weight.LargeUnit})", $"{L.Materials.Topcoat} {L.Materials.Thinner} {L.Properties.TotalCost} ({U.Currency.Symbol})",
                $"{L.Costs.TotalMaterialCost} ({U.Currency.Symbol})", $"{L.Time.EstimatedTime} ({U.Time.UnitLong})"
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
                    worksheet.Cell(row, col++).Value = $"{config.PrimerCoatMultiplier:F0}x";
                    worksheet.Cell(row, col++).Value = objCosts.PrimerCost;
                    worksheet.Cell(row, col++).Value = amounts.PrimerHardenerGrams;
                    worksheet.Cell(row, col++).Value = amounts.PrimerHardenerKg;
                    worksheet.Cell(row, col++).Value = objCosts.PrimerHardenerCost;
                    worksheet.Cell(row, col++).Value = amounts.PrimerThinnerGrams;
                    worksheet.Cell(row, col++).Value = amounts.PrimerThinnerKg;
                    worksheet.Cell(row, col++).Value = objCosts.PrimerThinnerCost;
                    worksheet.Cell(row, col++).Value = amounts.TopcoatGrams;
                    worksheet.Cell(row, col++).Value = amounts.TopcoatKg;
                    worksheet.Cell(row, col++).Value = $"{config.TopcoatCoatMultiplier:F0}x";
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
                                             MaterialConfig config, CostCalculation costs, double estimatedTime,
                                             CalculationResult calculationResult = null, Dictionary<string, double> calculationFactors = null)
        {
            worksheet.Cell(row, 1).Value = L.Sections.Summary;
            worksheet.Cell(row, 1).Style.Font.Bold = true;
            worksheet.Cell(row, 1).Style.Font.FontSize = 12;
            row += 1;
            
            worksheet.Cell(row, 1).Value = $"{L.General.TotalObjects}:";
            worksheet.Cell(row, 2).Value = objectCount;
            row += 1;
            
            worksheet.Cell(row, 1).Value = $"{L.Properties.SurfaceArea} ({U.Area.SmallUnit}):";
            worksheet.Cell(row, 2).Value = surfaceArea_mm2;
            row += 1;
            
            worksheet.Cell(row, 1).Value = $"{L.Properties.SurfaceArea} ({U.Area.LargeUnit}):";
            worksheet.Cell(row, 2).Value = surfaceArea_mm2 / UnitsAndLabelsConfigManager.GetAreaConversionFactor();
            row += 2;
            
            double totalArea_m2 = surfaceArea_mm2 / UnitsAndLabelsConfigManager.GetAreaConversionFactor();
            var totalAmounts = CalculateMaterialAmounts(totalArea_m2, config);
            
            if (config.Primer != null)
            {
                worksheet.Cell(row, 1).Value = $"{L.Materials.Primer}:";
                worksheet.Cell(row, 2).Value = config.Primer.Name;
                row += 1;
                worksheet.Cell(row, 1).Value = $"  {L.Properties.Consumption} ({U.Weight.SmallUnit}/{U.Area.LargeUnit}):";
                worksheet.Cell(row, 2).Value = config.Primer.ConsumptionPerSqM;
                row += 1;
                worksheet.Cell(row, 1).Value = $"  {L.Properties.Coats}:";
                worksheet.Cell(row, 2).Value = $"{config.PrimerCoatMultiplier:F0}x";
                row += 1;
                worksheet.Cell(row, 1).Value = $"  {L.Properties.TotalAmount} ({U.Weight.SmallUnit}):";
                worksheet.Cell(row, 2).Value = totalAmounts.PrimerGrams;
                row += 1;
                worksheet.Cell(row, 1).Value = $"  {L.Properties.TotalAmount} ({U.Weight.LargeUnit}):";
                worksheet.Cell(row, 2).Value = totalAmounts.PrimerKg;
                row += 1;
                worksheet.Cell(row, 1).Value = $"  {L.Properties.Price} ({U.Currency.Symbol}/{U.Weight.LargeUnit}):";
                worksheet.Cell(row, 2).Value = config.Primer.PricePerKg;
                row += 1;
                worksheet.Cell(row, 1).Value = $"  {L.Properties.TotalCost} ({U.Currency.Symbol}):";
                worksheet.Cell(row, 2).Value = costs.PrimerCost;
                row += 1;
                worksheet.Cell(row, 1).Value = $"  {L.Materials.Hardener} ({UnitsAndLabelsConfigManager.FormatPercentage(config.PrimerHardenerPercent)}):";
                worksheet.Cell(row, 2).Value = $"{UnitsAndLabelsConfigManager.FormatWeightBoth(totalAmounts.PrimerHardenerGrams)} - {UnitsAndLabelsConfigManager.FormatCurrency(costs.PrimerHardenerCost)}";
                row += 1;
                worksheet.Cell(row, 1).Value = $"  {L.Materials.Thinner} ({UnitsAndLabelsConfigManager.FormatPercentage(config.PrimerThinnerPercent)}):";
                worksheet.Cell(row, 2).Value = $"{UnitsAndLabelsConfigManager.FormatWeightBoth(totalAmounts.PrimerThinnerGrams)} - {UnitsAndLabelsConfigManager.FormatCurrency(costs.PrimerThinnerCost)}";
                row += 2;
            }

            if (config.Topcoat != null)
            {
                worksheet.Cell(row, 1).Value = $"{L.Materials.Topcoat}:";
                worksheet.Cell(row, 2).Value = config.Topcoat.Name;
                row += 1;
                worksheet.Cell(row, 1).Value = $"  {L.Properties.Consumption} ({U.Weight.SmallUnit}/{U.Area.LargeUnit}):";
                worksheet.Cell(row, 2).Value = config.Topcoat.ConsumptionPerSqM;
                row += 1;
                worksheet.Cell(row, 1).Value = $"  {L.Properties.Coats}:";
                worksheet.Cell(row, 2).Value = $"{config.TopcoatCoatMultiplier:F0}x";
                row += 1;
                worksheet.Cell(row, 1).Value = $"  {L.Properties.TotalAmount} ({U.Weight.SmallUnit}):";
                worksheet.Cell(row, 2).Value = totalAmounts.TopcoatGrams;
                row += 1;
                worksheet.Cell(row, 1).Value = $"  {L.Properties.TotalAmount} ({U.Weight.LargeUnit}):";
                worksheet.Cell(row, 2).Value = totalAmounts.TopcoatKg;
                row += 1;
                worksheet.Cell(row, 1).Value = $"  {L.Properties.Price} ({U.Currency.Symbol}/{U.Weight.LargeUnit}):";
                worksheet.Cell(row, 2).Value = config.Topcoat.PricePerKg;
                row += 1;
                worksheet.Cell(row, 1).Value = $"  {L.Properties.TotalCost} ({U.Currency.Symbol}):";
                worksheet.Cell(row, 2).Value = costs.TopcoatCost;
                row += 1;
                worksheet.Cell(row, 1).Value = $"  {L.Materials.Hardener} ({UnitsAndLabelsConfigManager.FormatPercentage(config.TopcoatHardenerPercent)}):";
                worksheet.Cell(row, 2).Value = $"{UnitsAndLabelsConfigManager.FormatWeightBoth(totalAmounts.TopcoatHardenerGrams)} - {UnitsAndLabelsConfigManager.FormatCurrency(costs.TopcoatHardenerCost)}";
                row += 1;
                worksheet.Cell(row, 1).Value = $"  {L.Materials.Thinner} ({UnitsAndLabelsConfigManager.FormatPercentage(config.TopcoatThinnerPercent)}):";
                worksheet.Cell(row, 2).Value = $"{UnitsAndLabelsConfigManager.FormatWeightBoth(totalAmounts.TopcoatThinnerGrams)} - {UnitsAndLabelsConfigManager.FormatCurrency(costs.TopcoatThinnerCost)}";
                row += 2;
            }
            
            worksheet.Cell(row, 1).Value = $"{L.Costs.TotalMaterialCost} ({U.Currency.Symbol}):";
            worksheet.Cell(row, 2).Value = costs.TotalMaterialCost;
            worksheet.Cell(row, 1).Style.Font.Bold = true;
            row += 1;
            
            worksheet.Cell(row, 1).Value = $"{L.Time.TotalEstimatedTime} ({U.Time.UnitLong}):";
            worksheet.Cell(row, 2).Value = estimatedTime;
            row += 1;
            
            worksheet.Cell(row, 1).Value = $"{L.Time.TimeFactor} ({U.Time.Unit}/{U.Area.LargeUnit}):";
            worksheet.Cell(row, 2).Value = config.TimeFactor;
            row += 2;

            // Add calculation breakdown if available
            if (calculationResult != null && calculationFactors != null)
            {
                worksheet.Cell(row, 1).Value = L.Sections.OfferCalculation;
                worksheet.Cell(row, 1).Style.Font.Bold = true;
                worksheet.Cell(row, 1).Style.Font.FontSize = 12;
                row += 1;

                worksheet.Cell(row, 1).Value = $"{L.Costs.BaseMaterialCost} ({U.Currency.Symbol}):";
                worksheet.Cell(row, 2).Value = calculationResult.BaseMaterialCost;
                row += 1;

                if (calculationResult.MaterialSurcharge > 0 && calculationFactors.TryGetValue("MaterialCostFactor", out double materialFactor))
                {
                    worksheet.Cell(row, 1).Value = $"{L.Costs.MaterialSurcharge} ({UnitsAndLabelsConfigManager.FormatPercentage(materialFactor)}):";
                    worksheet.Cell(row, 2).Value = calculationResult.MaterialSurcharge;
                    row += 1;
                }

                if (calculationResult.ProductionCost > 0 && calculationFactors.TryGetValue("ProductionCostFactor", out double productionFactor))
                {
                    worksheet.Cell(row, 1).Value = $"{L.Costs.ProductionCost} ({UnitsAndLabelsConfigManager.FormatPercentage(productionFactor)}):";
                    worksheet.Cell(row, 2).Value = calculationResult.ProductionCost;
                    row += 1;
                }

                if (calculationResult.AdministrationCost > 0 && calculationFactors.TryGetValue("AdministrationCostFactor", out double adminFactor))
                {
                    worksheet.Cell(row, 1).Value = $"{L.Costs.AdministrationCost} ({UnitsAndLabelsConfigManager.FormatPercentage(adminFactor)}):";
                    worksheet.Cell(row, 2).Value = calculationResult.AdministrationCost;
                    row += 1;
                }

                if (calculationResult.SalesCost > 0 && calculationFactors.TryGetValue("SalesCostFactor", out double salesFactor))
                {
                    worksheet.Cell(row, 1).Value = $"{L.Costs.SalesCost} ({UnitsAndLabelsConfigManager.FormatPercentage(salesFactor)}):";
                    worksheet.Cell(row, 2).Value = calculationResult.SalesCost;
                    row += 1;
                }

                if (calculationResult.ProfitMargin > 0 && calculationFactors.TryGetValue("ProfitMarginFactor", out double profitFactor))
                {
                    worksheet.Cell(row, 1).Value = $"{L.Costs.ProfitMargin} ({UnitsAndLabelsConfigManager.FormatPercentage(profitFactor)}):";
                    worksheet.Cell(row, 2).Value = calculationResult.ProfitMargin;
                    row += 1;
                }

                row += 1;
                worksheet.Cell(row, 1).Value = $"{L.Costs.FinalOfferPrice} ({U.Currency.Symbol}):";
                worksheet.Cell(row, 2).Value = calculationResult.FinalOfferPrice;
                worksheet.Cell(row, 1).Style.Font.Bold = true;
                worksheet.Cell(row, 2).Style.Font.Bold = true;
                worksheet.Cell(row, 2).Style.Fill.BackgroundColor = XLColor.LightYellow;
                row += 1;
            }

            return row;
        }

        #endregion

        #region Helper Methods

        private static CostCalculation CalculateIndividualCosts(double surfaceArea_mm2, MaterialConfig config)
        {
            CostCalculation costs = new CostCalculation();
            double surfaceArea_m2 = surfaceArea_mm2 / UnitsAndLabelsConfigManager.GetAreaConversionFactor();

            if (config.Primer != null)
            {
                double primerGrams = surfaceArea_m2 * config.Primer.ConsumptionPerSqM;
                double primerKg = primerGrams / UnitsAndLabelsConfigManager.GetWeightConversionFactor();
                costs.PrimerCost = primerKg * config.Primer.PricePerKg;
                
                double primerHardenerGrams = primerGrams * (config.PrimerHardenerPercent / 100.0);
                double primerHardenerKg = primerHardenerGrams / UnitsAndLabelsConfigManager.GetWeightConversionFactor();
                costs.PrimerHardenerCost = primerHardenerKg * config.HardenerPricePerKg;
                
                double primerThinnerGrams = primerGrams * (config.PrimerThinnerPercent / 100.0);
                double primerThinnerKg = primerThinnerGrams / UnitsAndLabelsConfigManager.GetWeightConversionFactor();
                costs.PrimerThinnerCost = primerThinnerKg * config.ThinnerPricePerKg;
                
                costs.TotalMaterialCost += costs.PrimerCost + costs.PrimerHardenerCost + costs.PrimerThinnerCost;
            }

            if (config.Topcoat != null)
            {
                double topcoatGrams = surfaceArea_m2 * config.Topcoat.ConsumptionPerSqM;
                double topcoatKg = topcoatGrams / UnitsAndLabelsConfigManager.GetWeightConversionFactor();
                costs.TopcoatCost = topcoatKg * config.Topcoat.PricePerKg;
                
                double topcoatHardenerGrams = topcoatGrams * (config.TopcoatHardenerPercent / 100.0);
                double topcoatHardenerKg = topcoatHardenerGrams / UnitsAndLabelsConfigManager.GetWeightConversionFactor();
                costs.TopcoatHardenerCost = topcoatHardenerKg * config.HardenerPricePerKg;
                
                double topcoatThinnerGrams = topcoatGrams * (config.TopcoatThinnerPercent / 100.0);
                double topcoatThinnerKg = topcoatThinnerGrams / UnitsAndLabelsConfigManager.GetWeightConversionFactor();
                costs.TopcoatThinnerCost = topcoatThinnerKg * config.ThinnerPricePerKg;
                
                costs.TotalMaterialCost += costs.TopcoatCost + costs.TopcoatHardenerCost + costs.TopcoatThinnerCost;
            }

            return costs;
        }

        private static double EstimateIndividualTime(double surfaceArea_mm2, double timeFactor)
        {
            double surfaceArea_m2 = surfaceArea_mm2 / UnitsAndLabelsConfigManager.GetAreaConversionFactor();
            return surfaceArea_m2 * timeFactor;
        }

        private static MaterialAmounts CalculateMaterialAmounts(double area_m2, MaterialConfig config)
        {
            var amounts = new MaterialAmounts();

            if (config.Primer != null)
            {
                amounts.PrimerGrams = area_m2 * config.Primer.ConsumptionPerSqM * config.PrimerCoatMultiplier;
                amounts.PrimerKg = amounts.PrimerGrams / UnitsAndLabelsConfigManager.GetWeightConversionFactor();
                amounts.PrimerHardenerGrams = amounts.PrimerGrams * (config.PrimerHardenerPercent / 100.0);
                amounts.PrimerHardenerKg = amounts.PrimerHardenerGrams / UnitsAndLabelsConfigManager.GetWeightConversionFactor();
                amounts.PrimerThinnerGrams = amounts.PrimerGrams * (config.PrimerThinnerPercent / 100.0);
                amounts.PrimerThinnerKg = amounts.PrimerThinnerGrams / UnitsAndLabelsConfigManager.GetWeightConversionFactor();
            }

            if (config.Topcoat != null)
            {
                amounts.TopcoatGrams = area_m2 * config.Topcoat.ConsumptionPerSqM * config.TopcoatCoatMultiplier;
                amounts.TopcoatKg = amounts.TopcoatGrams / UnitsAndLabelsConfigManager.GetWeightConversionFactor();
                amounts.TopcoatHardenerGrams = amounts.TopcoatGrams * (config.TopcoatHardenerPercent / 100.0);
                amounts.TopcoatHardenerKg = amounts.TopcoatHardenerGrams / UnitsAndLabelsConfigManager.GetWeightConversionFactor();
                amounts.TopcoatThinnerGrams = amounts.TopcoatGrams * (config.TopcoatThinnerPercent / 100.0);
                amounts.TopcoatThinnerKg = amounts.TopcoatThinnerGrams / UnitsAndLabelsConfigManager.GetWeightConversionFactor();
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

