using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace RHCoatingApp
{
    /// <summary>
    /// Manages loading and saving of calculation configuration from JSON file
    /// </summary>
    public class CalculationConfigManager
    {
        private const string ConfigFileName = "CalculationConfig.json";
        private static CalculationConfigData _config;
        private static readonly object _lock = new object();

        /// <summary>
        /// Gets the current calculation configuration, loading it if necessary
        /// </summary>
        public static CalculationConfigData Config
        {
            get
            {
                if (_config == null)
                {
                    lock (_lock)
                    {
                        if (_config == null)
                        {
                            _config = LoadConfig();
                        }
                    }
                }
                return _config;
            }
        }

        /// <summary>
        /// Loads calculation configuration from JSON file
        /// </summary>
        private static CalculationConfigData LoadConfig()
        {
            try
            {
                // Try to load from plugin directory
                string pluginDir = Path.GetDirectoryName(typeof(CalculationConfigManager).Assembly.Location);
                string configPath = Path.Combine(pluginDir, ConfigFileName);

                if (File.Exists(configPath))
                {
                    string json = File.ReadAllText(configPath);
                    var options = new JsonSerializerOptions
                    {
                        PropertyNameCaseInsensitive = true,
                        ReadCommentHandling = JsonCommentHandling.Skip,
                        AllowTrailingCommas = true
                    };
                    return JsonSerializer.Deserialize<CalculationConfigData>(json, options);
                }
                else
                {
                    Rhino.RhinoApp.WriteLine($"Calculation config file not found at: {configPath}. Using default values.");
                }
            }
            catch (Exception ex)
            {
                Rhino.RhinoApp.WriteLine($"Error loading calculation config: {ex.Message}. Using default values.");
            }

            // Return default configuration if loading fails
            return GetDefaultConfig();
        }

        /// <summary>
        /// Returns default calculation configuration values
        /// </summary>
        private static CalculationConfigData GetDefaultConfig()
        {
            return new CalculationConfigData
            {
                CalculationFactors = new System.Collections.Generic.Dictionary<string, CalculationFactor>
                {
                    {
                        "MaterialCostFactor",
                        new CalculationFactor
                        {
                            Key = "MaterialCostFactor",
                            Name = "Material Cost Factor",
                            Percentage = 18.0,
                            Description = "Additional costs for materials (surcharge)"
                        }
                    },
                    {
                        "ProductionCostFactor",
                        new CalculationFactor
                        {
                            Key = "ProductionCostFactor",
                            Name = "Production Cost Factor",
                            Percentage = 120.0,
                            Description = "Production and manufacturing costs"
                        }
                    },
                    {
                        "AdministrationCostFactor",
                        new CalculationFactor
                        {
                            Key = "AdministrationCostFactor",
                            Name = "Administration Cost Factor",
                            Percentage = 31.0,
                            Description = "Administrative and overhead costs"
                        }
                    },
                    {
                        "SalesCostFactor",
                        new CalculationFactor
                        {
                            Key = "SalesCostFactor",
                            Name = "Sales Cost Factor",
                            Percentage = 5.0,
                            Description = "Sales and distribution costs"
                        }
                    },
                    {
                        "ProfitMarginFactor",
                        new CalculationFactor
                        {
                            Key = "ProfitMarginFactor",
                            Name = "Profit Margin Factor",
                            Percentage = 19.0,
                            Description = "Desired profit margin"
                        }
                    }
                },
                DefaultSettings = new CalculationDefaultSettings
                {
                    ShowDetailedCalculation = true,
                    DefaultCalculationType = "Standard"
                }
            };
        }

        /// <summary>
        /// Reloads calculation configuration from file
        /// </summary>
        public static void ReloadConfig()
        {
            lock (_lock)
            {
                _config = LoadConfig();
            }
        }
    }

    #region Calculation Configuration Data Classes

    public class CalculationConfigData
    {
        public System.Collections.Generic.Dictionary<string, CalculationFactor> CalculationFactors { get; set; }
        public CalculationDefaultSettings DefaultSettings { get; set; }
    }

    public class CalculationFactor
    {
        public string Key { get; set; }
        public string Name { get; set; }
        public double Percentage { get; set; }
        public string Description { get; set; }
    }

    public class CalculationDefaultSettings
    {
        public bool ShowDetailedCalculation { get; set; }
        public string DefaultCalculationType { get; set; }
    }

    #endregion
}
