using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace RHCoatingApp
{
    /// <summary>
    /// Manages loading and saving of coating configuration from JSON file
    /// </summary>
    public class CoatingConfigManager
    {
        private const string ConfigFileName = "CoatingConfig.json";
        private static CoatingConfigData _config;
        private static readonly object _lock = new object();

        /// <summary>
        /// Gets the current configuration, loading it if necessary
        /// </summary>
        public static CoatingConfigData Config
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
        /// Loads configuration from JSON file
        /// </summary>
        private static CoatingConfigData LoadConfig()
        {
            try
            {
                // Try to load from plugin directory
                string pluginDir = Path.GetDirectoryName(typeof(CoatingConfigManager).Assembly.Location);
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
                    return JsonSerializer.Deserialize<CoatingConfigData>(json, options);
                }
                else
                {
                    Rhino.RhinoApp.WriteLine($"Config file not found at: {configPath}. Using default values.");
                }
            }
            catch (Exception ex)
            {
                Rhino.RhinoApp.WriteLine($"Error loading config: {ex.Message}. Using default values.");
            }

            // Return default configuration if loading fails
            return GetDefaultConfig();
        }

        /// <summary>
        /// Returns default configuration values
        /// </summary>
        private static CoatingConfigData GetDefaultConfig()
        {
            return new CoatingConfigData
            {
                ApplicationDefaults = new ApplicationDefaults
                {
                    Indoor = new AdditiveSettings
                    {
                        PrimerHardenerPercent = 6.0,
                        PrimerThinnerPercent = 10.0,
                        TopcoatHardenerPercent = 15.0,
                        TopcoatThinnerPercent = 20.0,
                        PrimerCoatMultiplier = 1.0,
                        TopcoatCoatMultiplier = 1.0
                    },
                    Outdoor = new AdditiveSettings
                    {
                        PrimerHardenerPercent = 8.0,
                        PrimerThinnerPercent = 12.0,
                        TopcoatHardenerPercent = 18.0,
                        TopcoatThinnerPercent = 25.0,
                        PrimerCoatMultiplier = 1.0,
                        TopcoatCoatMultiplier = 2.0
                    }
                },
                MaterialPrices = new MaterialPrices
                {
                    HardenerPricePerKg = 23.0,
                    ThinnerPricePerKg = 9.0
                },
                Materials = new List<MaterialDefinition>
                {
                    new MaterialDefinition { Name = "Standard Primer", Type = "Primer", ConsumptionPerSqM = 200.0, PricePerKg = 25.0 },
                    new MaterialDefinition { Name = "Premium Primer", Type = "Primer", ConsumptionPerSqM = 200.0, PricePerKg = 35.0 },
                    new MaterialDefinition { Name = "Basic Topcoat", Type = "Topcoat", ConsumptionPerSqM = 200.0, PricePerKg = 20.0 },
                    new MaterialDefinition { Name = "Premium Topcoat", Type = "Topcoat", ConsumptionPerSqM = 200.0, PricePerKg = 40.0 }
                },
                DefaultSettings = new DefaultSettings
                {
                    TimeFactor = 0.5,
                    TimePrice = 80.0,
                    DefaultApplicationType = "Indoor",
                    DefaultPrimer = "Standard Primer",
                    DefaultTopcoat = "Basic Topcoat",
                    DefaultPrimerCoatMultiplier = 1.0,
                    DefaultTopcoatCoatMultiplier = 1.0
                }
            };
        }

        /// <summary>
        /// Reloads configuration from file
        /// </summary>
        public static void ReloadConfig()
        {
            lock (_lock)
            {
                _config = LoadConfig();
            }
        }
    }

    #region Configuration Data Classes

    public class CoatingConfigData
    {
        public ApplicationDefaults ApplicationDefaults { get; set; }
        public MaterialPrices MaterialPrices { get; set; }
        public List<MaterialDefinition> Materials { get; set; }
        public DefaultSettings DefaultSettings { get; set; }
    }

    public class ApplicationDefaults
    {
        public AdditiveSettings Indoor { get; set; }
        public AdditiveSettings Outdoor { get; set; }
    }

    public class AdditiveSettings
    {
        public double PrimerHardenerPercent { get; set; }
        public double PrimerThinnerPercent { get; set; }
        public double TopcoatHardenerPercent { get; set; }
        public double TopcoatThinnerPercent { get; set; }
        public double PrimerCoatMultiplier { get; set; }
        public double TopcoatCoatMultiplier { get; set; }
    }

    public class MaterialPrices
    {
        public double HardenerPricePerKg { get; set; }
        public double ThinnerPricePerKg { get; set; }
    }

    public class MaterialDefinition
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public double ConsumptionPerSqM { get; set; }
        public double PricePerKg { get; set; }
    }

    public class DefaultSettings
    {
        public double TimeFactor { get; set; }
        public double TimePrice { get; set; }
        public string DefaultApplicationType { get; set; }
        public string DefaultPrimer { get; set; }
        public string DefaultTopcoat { get; set; }
        public double DefaultPrimerCoatMultiplier { get; set; }
        public double DefaultTopcoatCoatMultiplier { get; set; }
    }

    #endregion
}

