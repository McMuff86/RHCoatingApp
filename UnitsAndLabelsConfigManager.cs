using System;
using System.IO;
using System.Text.Json;

namespace RHCoatingApp
{
    /// <summary>
    /// Manages units and labels configuration for internationalization and flexibility
    /// Provides centralized access to units, labels, and formatting options
    /// </summary>
    public class UnitsAndLabelsConfigManager
    {
        private static readonly Lazy<UnitsAndLabelsConfigData> _config = new Lazy<UnitsAndLabelsConfigData>(LoadConfig);

        /// <summary>
        /// Gets the current units and labels configuration
        /// </summary>
        public static UnitsAndLabelsConfigData Config => _config.Value;

        private static UnitsAndLabelsConfigData LoadConfig()
        {
            try
            {
                string configPath = Path.Combine(
                    Path.GetDirectoryName(typeof(UnitsAndLabelsConfigManager).Assembly.Location),
                    "UnitsAndLabels.json"
                );

                if (File.Exists(configPath))
                {
                    string json = File.ReadAllText(configPath);
                    var config = JsonSerializer.Deserialize<UnitsAndLabelsConfigData>(json, new JsonSerializerOptions
                    {
                        PropertyNameCaseInsensitive = true,
                        ReadCommentHandling = JsonCommentHandling.Skip
                    });

                    if (config != null)
                    {
                        return config;
                    }
                }

                Rhino.RhinoApp.WriteLine("UnitsAndLabels.json not found. Using default configuration.");
            }
            catch (Exception ex)
            {
                Rhino.RhinoApp.WriteLine($"Error loading UnitsAndLabels.json: {ex.Message}. Using default configuration.");
            }

            return GetDefaultConfig();
        }

        private static UnitsAndLabelsConfigData GetDefaultConfig()
        {
            return new UnitsAndLabelsConfigData
            {
                Units = new UnitsConfig
                {
                    Area = new AreaUnits { SmallUnit = "mm²", LargeUnit = "m²", ConversionFactor = 1000000.0 },
                    Weight = new WeightUnits { SmallUnit = "g", LargeUnit = "kg", ConversionFactor = 1000.0 },
                    Currency = new CurrencyConfig { Symbol = "Fr.", Position = "before" },
                    Time = new TimeUnits { Unit = "h", UnitLong = "hours" }
                },
                Labels = new LabelsConfig
                {
                    Materials = new MaterialLabels
                    {
                        Primer = "Primer",
                        Topcoat = "Topcoat",
                        Hardener = "Hardener",
                        Thinner = "Thinner"
                    },
                    Properties = new PropertyLabels
                    {
                        Consumption = "Consumption",
                        Price = "Price",
                        Coats = "Coats",
                        TotalAmount = "Total Amount",
                        TotalCost = "Total Cost",
                        SurfaceArea = "Surface Area"
                    },
                    Costs = new CostLabels
                    {
                        BaseMaterialCost = "Base Material Cost",
                        TotalMaterialCost = "Total Material Cost",
                        MaterialSurcharge = "Material Surcharge",
                        ProductionCost = "Production Cost",
                        AdministrationCost = "Administration Cost",
                        SalesCost = "Sales Cost",
                        ProfitMargin = "Profit Margin",
                        FinalOfferPrice = "FINAL OFFER PRICE"
                    },
                    Time = new TimeLabels
                    {
                        EstimatedTime = "Estimated Time",
                        TotalEstimatedTime = "Total Estimated Time",
                        TimeFactor = "Time Factor"
                    },
                    Sections = new SectionLabels
                    {
                        IndividualObjects = "INDIVIDUAL OBJECTS",
                        Summary = "SUMMARY",
                        OfferCalculation = "OFFER CALCULATION",
                        Results = "Results",
                        CalculationResults = "Rhino CoatingApp Calculation Results"
                    },
                    General = new GeneralLabels
                    {
                        Object = "Object",
                        Name = "Name",
                        TotalObjects = "Total Objects",
                        Date = "Date",
                        ApplicationType = "Application Type",
                        CoatMultipliers = "Coat Multipliers"
                    },
                    UI = new UILabels
                    {
                        NoObjectsSelected = "No objects selected",
                        ClickToCalculate = "Click 'Calculate' to see results...",
                        SelectObjectsFirst = "Select objects and click 'Calculate' to see results...",
                        ObjectsCleared = "Objects cleared. Select new objects to calculate."
                    },
                    Export = new ExportLabels
                    {
                        ExportTitle = "Export Coating Calculation",
                        ExcelFiles = "Excel Files",
                        CsvFiles = "CSV Files",
                        TextFiles = "Text Files",
                        AllFiles = "All Files",
                        ExportSuccess = "Results exported successfully to:",
                        ExportError = "Error exporting results:"
                    }
                },
                FormatStrings = new FormatStringsConfig
                {
                    AreaSmall = "{0:F2} {1}",
                    AreaLarge = "{0:F4} {1}",
                    WeightSmall = "{0:F1} {1}",
                    WeightLarge = "{0:F3} {1}",
                    Currency = "{0} {1:F2}",
                    Time = "{0:F2} {1}",
                    Percentage = "{0:F1}%",
                    PerUnit = "{0}/{1}"
                }
            };
        }

        #region Helper Methods for Formatting

        /// <summary>
        /// Formats area in small units (mm²)
        /// </summary>
        public static string FormatAreaSmall(double areaInSmallUnits)
        {
            var config = Config;
            return string.Format(config.FormatStrings.AreaSmall, areaInSmallUnits, config.Units.Area.SmallUnit);
        }

        /// <summary>
        /// Formats area in large units (m²)
        /// </summary>
        public static string FormatAreaLarge(double areaInSmallUnits)
        {
            var config = Config;
            double areaInLargeUnits = areaInSmallUnits / config.Units.Area.ConversionFactor;
            return string.Format(config.FormatStrings.AreaLarge, areaInLargeUnits, config.Units.Area.LargeUnit);
        }

        /// <summary>
        /// Formats area with both units
        /// </summary>
        public static string FormatAreaBoth(double areaInSmallUnits)
        {
            return $"{FormatAreaSmall(areaInSmallUnits)} ({FormatAreaLarge(areaInSmallUnits)})";
        }

        /// <summary>
        /// Formats weight in small units (g)
        /// </summary>
        public static string FormatWeightSmall(double weightInSmallUnits)
        {
            var config = Config;
            return string.Format(config.FormatStrings.WeightSmall, weightInSmallUnits, config.Units.Weight.SmallUnit);
        }

        /// <summary>
        /// Formats weight in large units (kg)
        /// </summary>
        public static string FormatWeightLarge(double weightInSmallUnits)
        {
            var config = Config;
            double weightInLargeUnits = weightInSmallUnits / config.Units.Weight.ConversionFactor;
            return string.Format(config.FormatStrings.WeightLarge, weightInLargeUnits, config.Units.Weight.LargeUnit);
        }

        /// <summary>
        /// Formats weight with both units
        /// </summary>
        public static string FormatWeightBoth(double weightInSmallUnits)
        {
            return $"{FormatWeightSmall(weightInSmallUnits)} ({FormatWeightLarge(weightInSmallUnits)})";
        }

        /// <summary>
        /// Formats currency value
        /// </summary>
        public static string FormatCurrency(double amount)
        {
            var config = Config;
            if (config.Units.Currency.Position == "before")
            {
                return string.Format(config.FormatStrings.Currency, config.Units.Currency.Symbol, amount);
            }
            else
            {
                return string.Format("{1:F2} {0}", config.Units.Currency.Symbol, amount);
            }
        }

        /// <summary>
        /// Formats time value
        /// </summary>
        public static string FormatTime(double timeValue)
        {
            var config = Config;
            return string.Format(config.FormatStrings.Time, timeValue, config.Units.Time.UnitLong);
        }

        /// <summary>
        /// Formats percentage value
        /// </summary>
        public static string FormatPercentage(double percentage)
        {
            return string.Format(Config.FormatStrings.Percentage, percentage);
        }

        /// <summary>
        /// Formats per-unit value (e.g., "g/m²", "Fr./kg")
        /// </summary>
        public static string FormatPerUnit(string numerator, string denominator)
        {
            return string.Format(Config.FormatStrings.PerUnit, numerator, denominator);
        }

        /// <summary>
        /// Gets conversion factor from small to large units for area
        /// </summary>
        public static double GetAreaConversionFactor()
        {
            return Config.Units.Area.ConversionFactor;
        }

        /// <summary>
        /// Gets conversion factor from small to large units for weight
        /// </summary>
        public static double GetWeightConversionFactor()
        {
            return Config.Units.Weight.ConversionFactor;
        }

        #endregion
    }

    #region Data Classes

    public class UnitsAndLabelsConfigData
    {
        public UnitsConfig Units { get; set; }
        public LabelsConfig Labels { get; set; }
        public FormatStringsConfig FormatStrings { get; set; }
    }

    public class UnitsConfig
    {
        public AreaUnits Area { get; set; }
        public WeightUnits Weight { get; set; }
        public CurrencyConfig Currency { get; set; }
        public TimeUnits Time { get; set; }
    }

    public class AreaUnits
    {
        public string SmallUnit { get; set; }
        public string LargeUnit { get; set; }
        public double ConversionFactor { get; set; }
    }

    public class WeightUnits
    {
        public string SmallUnit { get; set; }
        public string LargeUnit { get; set; }
        public double ConversionFactor { get; set; }
    }

    public class CurrencyConfig
    {
        public string Symbol { get; set; }
        public string Position { get; set; }
    }

    public class TimeUnits
    {
        public string Unit { get; set; }
        public string UnitLong { get; set; }
    }

    public class LabelsConfig
    {
        public MaterialLabels Materials { get; set; }
        public PropertyLabels Properties { get; set; }
        public CostLabels Costs { get; set; }
        public TimeLabels Time { get; set; }
        public SectionLabels Sections { get; set; }
        public GeneralLabels General { get; set; }
        public UILabels UI { get; set; }
        public ExportLabels Export { get; set; }
    }

    public class MaterialLabels
    {
        public string Primer { get; set; }
        public string Topcoat { get; set; }
        public string Hardener { get; set; }
        public string Thinner { get; set; }
    }

    public class PropertyLabels
    {
        public string Consumption { get; set; }
        public string Price { get; set; }
        public string Coats { get; set; }
        public string TotalAmount { get; set; }
        public string TotalCost { get; set; }
        public string SurfaceArea { get; set; }
    }

    public class CostLabels
    {
        public string BaseMaterialCost { get; set; }
        public string TotalMaterialCost { get; set; }
        public string MaterialSurcharge { get; set; }
        public string ProductionCost { get; set; }
        public string AdministrationCost { get; set; }
        public string SalesCost { get; set; }
        public string ProfitMargin { get; set; }
        public string FinalOfferPrice { get; set; }
    }

    public class TimeLabels
    {
        public string EstimatedTime { get; set; }
        public string TotalEstimatedTime { get; set; }
        public string TimeFactor { get; set; }
    }

    public class SectionLabels
    {
        public string IndividualObjects { get; set; }
        public string Summary { get; set; }
        public string OfferCalculation { get; set; }
        public string Results { get; set; }
        public string CalculationResults { get; set; }
    }

    public class GeneralLabels
    {
        public string Object { get; set; }
        public string Name { get; set; }
        public string TotalObjects { get; set; }
        public string Date { get; set; }
        public string ApplicationType { get; set; }
        public string CoatMultipliers { get; set; }
    }

    public class UILabels
    {
        public string NoObjectsSelected { get; set; }
        public string ClickToCalculate { get; set; }
        public string SelectObjectsFirst { get; set; }
        public string ObjectsCleared { get; set; }
    }

    public class ExportLabels
    {
        public string ExportTitle { get; set; }
        public string ExcelFiles { get; set; }
        public string CsvFiles { get; set; }
        public string TextFiles { get; set; }
        public string AllFiles { get; set; }
        public string ExportSuccess { get; set; }
        public string ExportError { get; set; }
    }

    public class FormatStringsConfig
    {
        public string AreaSmall { get; set; }
        public string AreaLarge { get; set; }
        public string WeightSmall { get; set; }
        public string WeightLarge { get; set; }
        public string Currency { get; set; }
        public string Time { get; set; }
        public string Percentage { get; set; }
        public string PerUnit { get; set; }
    }

    #endregion
}

