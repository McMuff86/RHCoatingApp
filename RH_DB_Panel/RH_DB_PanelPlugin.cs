using System;
using Rhino;
using Rhino.PlugIns;
using Rhino.UI;
using Eto.Forms;
using Eto.Drawing;
using System.Drawing;
using System.Runtime.InteropServices;
using System.IO;
using System.Data;
using CsvHelper;
using System.Globalization;
using System.Linq;
using ClosedXML.Excel;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using Rhino.Display;
using Rhino.Geometry;
using Rhino.Input.Custom;

namespace RH_DB_Panel
{
    ///<summary>
    /// <para>Every RhinoCommon .rhp assembly must have one and only one PlugIn-derived
    /// class. DO NOT create instances of this class yourself. It is the
    /// responsibility of Rhino to create an instance of this class.</para>
    /// <para>To complete plug-in information, please also see all PlugInDescription
    /// attributes in AssemblyInfo.cs (you might need to click "Project" ->
    /// "Show All Files" to see it in the "Solution Explorer" window).</para>
    ///</summary>
    public class RH_DB_PanelPlugin : PlugIn
    {
        public RH_DB_PanelPlugin()
        {
            Instance = this;
        }

        ///<summary>Gets the only instance of the RH_DB_PanelPlugin plug-in.</summary>
        public static RH_DB_PanelPlugin Instance { get; private set; }

        internal HoverTooltipService HoverService { get; private set; }

        protected override LoadReturnCode OnLoad(ref string errorMessage)
        {
            var panel_type = typeof(DatabasePanel);
            // If System.Drawing.Common is added via NuGet, you can use System.Drawing.SystemIcons.Application
            Panels.RegisterPanel(this, panel_type, "DB Panel", System.Drawing.SystemIcons.Application);
            HoverService = new HoverTooltipService();
            return LoadReturnCode.Success;
        }

        protected override void OnShutdown()
        {
            try
            {
                HoverService?.Shutdown();
            }
            catch { }
            base.OnShutdown();
        }
    }

    [Guid("03375280-E2AB-4CCB-A3C0-3B8565B4FB0D")]
    public class DatabasePanel : Panel
    {
        // Fields
        private DataTable dataTable;  // existing
        private DataTable currentTable;  // new, tracks what's shown
        private readonly GridView gridView;
        private readonly TextBox searchBox;
        private readonly Button searchButton;
        private readonly Button openButton;
        private readonly ComboBox sheetSelector;
        private readonly Button assignButton;
        private readonly CheckBox hoverTooltipToggle;
        private readonly NumericStepper hoverDelayMs;
        private readonly NumericStepper maxTooltipLines;
        private string currentFilePath;

        public DatabasePanel()
        {
            // Create UI controls first
            searchBox = new TextBox { PlaceholderText = "Search..." };
            searchButton = new Button { Text = "Search" };
            searchButton.Click += OnSearchButtonClick;

            openButton = new Button { Text = "Open XLSX" };
            openButton.Click += OnOpenButtonClick;

            assignButton = new Button { Text = "Assign to Object" };
            assignButton.Click += OnAssignButtonClick;

            sheetSelector = new ComboBox();
            sheetSelector.SelectedIndexChanged += OnSheetSelected;

            gridView = new GridView
            {
                AllowMultipleSelection = false,
                Height = -1, // Allow natural height
                Width = -1   // Allow natural width
            };

            hoverTooltipToggle = new CheckBox { Text = "Hover tooltip", Checked = true };
            hoverTooltipToggle.CheckedChanged += (s, e2) =>
            {
                var svc = RH_DB_PanelPlugin.Instance?.HoverService;
                if (svc != null)
                    svc.SetActive(hoverTooltipToggle.Checked == true);
            };

            hoverDelayMs = new NumericStepper { MinValue = 100, MaxValue = 5000, Value = 500, Increment = 100, Width = 80 };
            hoverDelayMs.ValueChanged += (s, e2) =>
            {
                var svc = RH_DB_PanelPlugin.Instance?.HoverService;
                if (svc != null)
                    svc.DelayMs = (int)Math.Round(hoverDelayMs.Value);
            };

            maxTooltipLines = new NumericStepper { MinValue = 0, MaxValue = 1000, Value = 12, Increment = 1, Width = 70 };
            maxTooltipLines.ValueChanged += (s, e2) =>
            {
                var svc = RH_DB_PanelPlugin.Instance?.HoverService;
                if (svc != null)
                    svc.MaxLines = (int)Math.Round(maxTooltipLines.Value);
            };

            // Tooltip text color selector
            var tooltipColorCombo = new ComboBox();
            tooltipColorCombo.Items.Add("Blue");
            tooltipColorCombo.Items.Add("White");
            tooltipColorCombo.Items.Add("Yellow");
            tooltipColorCombo.Items.Add("Green");
            tooltipColorCombo.Items.Add("Red");
            tooltipColorCombo.Items.Add("Black");
            tooltipColorCombo.SelectedIndex = 0; // Default to Blue
            tooltipColorCombo.SelectedIndexChanged += (s, e2) =>
            {
                var svc = RH_DB_PanelPlugin.Instance?.HoverService;
                if (svc != null)
                    svc.TextColor = GetColorFromName(tooltipColorCombo.SelectedValue?.ToString() ?? "Blue");
            };

            // Tooltip font size selector
            var tooltipFontSize = new NumericStepper { MinValue = 8, MaxValue = 24, Value = 12, Increment = 1, Width = 70 };
            tooltipFontSize.ValueChanged += (s, e2) =>
            {
                var svc = RH_DB_PanelPlugin.Instance?.HoverService;
                if (svc != null)
                    svc.FontSize = (int)Math.Round(tooltipFontSize.Value);
            };

            var layout = new DynamicLayout();
            layout.Padding = new Eto.Drawing.Padding(10);
            layout.Spacing = new Eto.Drawing.Size(5, 5);

            // Search controls in top row
            layout.AddRow(searchBox, searchButton);

            // Main controls in left column, tooltip controls in right column
            layout.BeginHorizontal();

            // Left column: main controls
            layout.BeginVertical();
            layout.AddRow(openButton);
            layout.AddRow(sheetSelector);
            layout.AddRow(assignButton);
            layout.EndVertical();

            // Right column: tooltip controls
            layout.BeginVertical();
            layout.AddRow(null, hoverTooltipToggle, null); // Center the checkbox
            layout.AddRow(new Label { Text = "Delay (ms):" }, hoverDelayMs);
            layout.AddRow(new Label { Text = "Max lines:" }, maxTooltipLines);
            layout.AddRow(new Label { Text = "Text color:" }, tooltipColorCombo);
            layout.AddRow(new Label { Text = "Font size:" }, tooltipFontSize);
            layout.EndVertical();

            layout.EndHorizontal();

            // DataGrid takes remaining space and is resizable
            layout.Add(gridView, xscale: true, yscale: true);

            Content = layout;

            // Set panel size constraints for better resizeability
            MinimumSize = new Eto.Drawing.Size(600, 400);
            Size = new Eto.Drawing.Size(1000, 700);

            // Initialize hover service defaults
            var initSvc = RH_DB_PanelPlugin.Instance?.HoverService;
            if (initSvc != null)
            {
                initSvc.DelayMs = (int)Math.Round(hoverDelayMs.Value);
                initSvc.MaxLines = (int)Math.Round(maxTooltipLines.Value);
                initSvc.TextColor = GetColorFromName(tooltipColorCombo.SelectedValue?.ToString() ?? "Blue");
                initSvc.FontSize = (int)Math.Round(tooltipFontSize.Value);
                initSvc.SetActive(hoverTooltipToggle.Checked == true);
            }

            try
            {
                // Now set initial file and populate sheets (this will trigger OnSheetSelected)
                // Try loading the primary file (this will trigger OnSheetSelected)
                currentFilePath = @"C:\Users\Adi.Muff\source\repos\RH_DB_Panel\csv\Schuler_Tueren_Complete_Database.xlsx";
                PopulateSheetSelector(currentFilePath);
            }
            catch (Exception)
            {
                // If primary file is missing or fails, check fallback directory
                string fallbackDir = @"C:\Users\adrian.muff\source\repos\work\RH_DB_Panel\csv";

                var files = Directory.GetFiles(fallbackDir, "*.xlsx");
                if (files.Length > 0)
                {
                    currentFilePath = files[0]; // take first match
                    PopulateSheetSelector(currentFilePath);
                }
                else
                {
                    MessageBox.Show("Keine Excel-Datei gefunden im Fallback-Ordner:\n" + fallbackDir);
                }
            }

        }

        private void LoadData(string filePath, string sheetName = null)
        {
            try
            {
                dataTable = new DataTable();

                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheet(sheetName ?? workbook.Worksheets.First().Name);
                    if (worksheet == null)
                    {
                        RhinoApp.WriteLine($"Sheet '{sheetName}' not found in XLSX.");
                        return;
                    }

                    var rowCount = worksheet.LastRowUsed()?.RowNumber() ?? 0;
                    var colCount = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;

                    if (rowCount == 0 || colCount == 0)
                    {
                        RhinoApp.WriteLine("No data in sheet.");
                        return;
                    }

                    for (int col = 1; col <= colCount; col++)
                    {
                        var headerText = worksheet.Cell(1, col).GetValue<string>();
                        var colName = string.IsNullOrWhiteSpace(headerText) ? $"Column{col}" : headerText;
                        dataTable.Columns.Add(colName);
                    }

                    for (int row = 2; row <= rowCount; row++)
                    {
                        var dataRow = dataTable.NewRow();
                        for (int col = 1; col <= colCount; col++)
                        {
                            dataRow[col - 1] = worksheet.Cell(row, col).Value;
                        }
                        dataTable.Rows.Add(dataRow);
                    }

                    RhinoApp.WriteLine($"Loaded {dataTable.Rows.Count} rows from {filePath} sheet '{worksheet.Name}'.");
                }
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Error loading XLSX: {ex.Message}");
            }
        }

        private void UpdateGridView(DataTable dt)
        {
            currentTable = dt;  // keep reference
            gridView.Columns.Clear();
            for (int colIndex = 0; colIndex < dt.Columns.Count; colIndex++)
            {
                var column = dt.Columns[colIndex];
                gridView.Columns.Add(new GridColumn
                {
                    DataCell = new TextBoxCell(colIndex),
                    HeaderText = column.ColumnName,
                    Editable = false,
                    AutoSize = true
                });
            }

            gridView.DataStore = dt.Rows.Cast<DataRow>().Select(row => row.ItemArray).ToArray();
            gridView.AllowColumnReordering = true;
            gridView.ShowHeader = true;
        }

        private void OnSearchButtonClick(object sender, EventArgs e)
        {
            var searchText = searchBox.Text.Trim();
            if (string.IsNullOrEmpty(searchText))
            {
                UpdateGridView(dataTable);
                return;
            }

            // Tokenize
            var parts = searchText.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            var criteria = parts
                .Select(p => p.Split(new[] { ':' }, 2))
                .Where(p => p.Length == 2)
                .Select(p => new { Key = p[0].Trim(), Value = p[1].Trim() })
                .ToList();

            IEnumerable<DataRow> filteredRowsEnum;

            // Build case-insensitive column lookup with aliases
            var colLookup = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (DataColumn c in dataTable.Columns)
                colLookup[c.ColumnName] = c.ColumnName;

            // Common aliases
            if (!colLookup.ContainsKey("Dicke") && colLookup.ContainsKey("Dicke_mm"))
                colLookup["Dicke"] = "Dicke_mm";

            if (criteria.Count == 0)
            {
                // Global search across all columns. Support '*' wildcard.
                var containsWildcard = searchText.Contains('*');
                if (containsWildcard)
                {
                    var regexPattern = Regex.Escape(searchText).Replace("\\*", ".*");
                    var rx = new Regex(regexPattern, RegexOptions.IgnoreCase);
                    filteredRowsEnum = dataTable.Rows.Cast<DataRow>()
                        .Where(row => row.ItemArray.Any(f => rx.IsMatch(f?.ToString() ?? string.Empty)));
                }
                else
                {
                    var needle = searchText.ToLowerInvariant();
                    filteredRowsEnum = dataTable.Rows.Cast<DataRow>()
                        .Where(row => row.ItemArray.Any(f => (f?.ToString() ?? string.Empty).ToLowerInvariant().Contains(needle)));
                }
            }
            else
            {
                // Column-specific multi-criteria search
                filteredRowsEnum = dataTable.Rows.Cast<DataRow>();
                foreach (var kv in criteria)
                {
                    // Resolve column name (exact, alias, or fuzzy match)
                    string columnName;
                    if (!colLookup.TryGetValue(kv.Key, out columnName))
                    {
                        var match = dataTable.Columns.Cast<DataColumn>()
                            .Select(c => c.ColumnName)
                            .FirstOrDefault(n => n.Equals(kv.Key, StringComparison.OrdinalIgnoreCase) ||
                                                 n.IndexOf(kv.Key, StringComparison.OrdinalIgnoreCase) >= 0);
                        if (match == null)
                            continue; // skip unknown key
                        columnName = match;
                    }

                    if (!dataTable.Columns.Contains(columnName))
                        continue;

                    if (kv.Value.Contains('*'))
                    {
                        var regexPattern = Regex.Escape(kv.Value).Replace("\\*", ".*");
                        var rx = new Regex(regexPattern, RegexOptions.IgnoreCase);
                        filteredRowsEnum = filteredRowsEnum.Where(row => rx.IsMatch(row[columnName]?.ToString() ?? string.Empty));
                    }
                    else
                    {
                        var needle = kv.Value.ToLowerInvariant();
                        filteredRowsEnum = filteredRowsEnum.Where(row => (row[columnName]?.ToString() ?? string.Empty).ToLowerInvariant().Contains(needle));
                    }
                }
            }

            DataTable filteredTable;
            if (filteredRowsEnum.Any())
                filteredTable = filteredRowsEnum.CopyToDataTable();
            else
                filteredTable = dataTable.Clone();

            UpdateGridView(filteredTable);
        }

        private void OnOpenButtonClick(object sender, EventArgs e)
        {
            var dialog = new Eto.Forms.OpenFileDialog  // Fully qualified
            {
                Title = "Select XLSX File",
                Filters = { new FileFilter("Excel Files", ".xlsx") }
            };

            if (dialog.ShowDialog(this) == DialogResult.Ok)
            {
                currentFilePath = dialog.FileName;
                PopulateSheetSelector(currentFilePath);
            }
        }

        private void PopulateSheetSelector(string filePath)
        {
            if (sheetSelector == null)
            {
                RhinoApp.WriteLine("Sheet selector not initialized.");
                return;
            }

            sheetSelector.Items.Clear();
            using (var workbook = new XLWorkbook(filePath))
            {
                foreach (var sheet in workbook.Worksheets)
                {
                    sheetSelector.Items.Add(sheet.Name);
                }
            }

            if (sheetSelector.Items.Count > 0)
            {
                sheetSelector.SelectedIndex = 0;  // Load first sheet
            }
        }

        private void OnSheetSelected(object sender, EventArgs e)
        {
            if (sheetSelector.SelectedIndex >= 0 && !string.IsNullOrEmpty(currentFilePath))
            {
                var selectedSheet = sheetSelector.SelectedValue.ToString();
                LoadData(currentFilePath, selectedSheet);
                UpdateGridView(dataTable);
            }
        }

        private void OnAssignButtonClick(object sender, EventArgs e)
        {
            try
            {
                if (currentTable == null || currentTable.Rows.Count == 0)
                {
                    RhinoApp.WriteLine("No data loaded.");
                    return;
                }

                var selectedIndex = gridView.SelectedRows != null && gridView.SelectedRows.Any()
                    ? gridView.SelectedRows.First()
                    : -1;
                if (selectedIndex < 0)
                {
                    RhinoApp.WriteLine("Please select a row in the table first.");
                    return;
                }

                if (selectedIndex >= currentTable.Rows.Count)
                {
                    RhinoApp.WriteLine("Selected row index is out of range.");
                    return;
                }

                var row = currentTable.Rows[selectedIndex];

                // Ask user to pick a single Rhino object
                var go = new Rhino.Input.Custom.GetObject();
                go.SetCommandPrompt("Select object to assign user text");
                go.SubObjectSelect = false;
                go.DisablePreSelect();
                go.OneByOnePostSelect = true;
                go.Get();
                if (go.CommandResult() != Rhino.Commands.Result.Success)
                    return;

                var objRef = go.Object(0);
                var rhObj = objRef?.Object();
                if (rhObj == null)
                {
                    RhinoApp.WriteLine("No object selected.");
                    return;
                }

                var doc = rhObj.Document ?? Rhino.RhinoDoc.ActiveDoc;
                int added = 0;

                if (doc != null)
                {
                    var attrs = rhObj.Attributes.Duplicate();
                    foreach (DataColumn col in currentTable.Columns)
                    {
                        var key = col.ColumnName;
                        var value = row[col] != null ? row[col].ToString() : string.Empty;
                        if (value == null)
                            value = string.Empty;
                        attrs.SetUserString(key, value);
                        added++;
                    }

                    var ok = doc.Objects.ModifyAttributes(rhObj, attrs, true);
                    if (ok)
                        RhinoApp.WriteLine($"Assigned {added} key/value pairs to object {rhObj.Id}.");
                    else
                        RhinoApp.WriteLine("Failed to modify object attributes.");
                }
                else
                {
                    foreach (DataColumn col in currentTable.Columns)
                    {
                        var key = col.ColumnName;
                        var value = row[col] != null ? row[col].ToString() : string.Empty;
                        if (value == null)
                            value = string.Empty;
                        rhObj.Attributes.SetUserString(key, value);
                        added++;
                    }
                    rhObj.CommitChanges();
                    RhinoApp.WriteLine($"Assigned {added} key/value pairs to object {rhObj.Id}.");
                }
            }
            catch (Exception ex)
            {
                RhinoApp.WriteLine($"Error assigning user text: {ex.Message}");
            }
        }

        public static Guid PanelId => typeof(DatabasePanel).GUID;

        private static System.Drawing.Color GetColorFromName(string colorName)
        {
            var lowerName = colorName?.ToLower() ?? "blue";
            switch (lowerName)
            {
                case "blue": return System.Drawing.Color.Blue;
                case "white": return System.Drawing.Color.White;
                case "yellow": return System.Drawing.Color.Yellow;
                case "green": return System.Drawing.Color.Green;
                case "red": return System.Drawing.Color.Red;
                case "black": return System.Drawing.Color.Black;
                default: return System.Drawing.Color.Blue;
            }
        }
    }

    internal class HoverTooltipService : MouseCallback
    {
        private readonly HoverTooltipConduit _conduit = new HoverTooltipConduit();
        private bool _active;
        private DateTime _lastMoveUtc;
        private System.Drawing.Point _lastScreenPt;
        private RhinoView _lastView;
        private RhinoViewport _lastViewport;
        private readonly TimeSpan _stabilityTolerance = TimeSpan.FromMilliseconds(25);
        private System.Drawing.Point _lastShownPt;
        private Guid _lastShownObjectId = Guid.Empty;

        public int DelayMs { get; set; } = 500;
        public int MaxLines { get; set; } = 50; // 0 = all
        public System.Drawing.Color TextColor { get; set; } = System.Drawing.Color.Blue;
        public int FontSize { get; set; } = 12;

        public void SetActive(bool active)
        {
            if (_active == active)
                return;
            _active = active;
            try { Enabled = active; } catch { }
            if (active)
            {
                RhinoApp.Idle += OnIdle;
            }
            else
            {
                RhinoApp.Idle -= OnIdle;
                _conduit.Enabled = false;
                _lastShownObjectId = Guid.Empty;
            }
        }

        public void Shutdown()
        {
            SetActive(false);
        }

        protected override void OnMouseMove(MouseCallbackEventArgs e)
        {
            if (!_active || e == null || e.View == null)
                return;
            _lastView = e.View;
            _lastViewport = e.View.ActiveViewport;
            _lastScreenPt = e.ViewportPoint;
            _lastMoveUtc = DateTime.UtcNow;
        }

        private void OnIdle(object sender, EventArgs e)
        {
            if (!_active || _lastView == null || _lastViewport == null)
                return;

            var elapsed = DateTime.UtcNow - _lastMoveUtc;
            if (elapsed < TimeSpan.FromMilliseconds(Math.Max(0, DelayMs)) - _stabilityTolerance)
                return;

            // Perform a point pick at the last position
            var doc = _lastView.Document ?? RhinoDoc.ActiveDoc;
            if (doc == null)
                return;

            var pc = new PickContext
            {
                View = _lastView,
                PickStyle = PickStyle.PointPick
            };
            var pickXform = _lastViewport.GetPickTransform(_lastScreenPt);
            pc.SetPickTransform(pickXform);

            var picked = doc.Objects.PickObjects(pc);
            if (picked != null && picked.Length > 0)
            {
                Rhino.DocObjects.RhinoObject rhObj = null;
                foreach (var objRef in picked)
                {
                    rhObj = objRef.Object();
                    if (rhObj != null)
                        break;
                }

                if (rhObj != null)
                {
                    var userStrings = rhObj.Attributes.GetUserStrings();
                    if (userStrings != null && userStrings.Count > 0)
                    {
                        var lines = new List<string>();
                        foreach (var key in userStrings.AllKeys)
                        {
                            var val = rhObj.Attributes.GetUserString(key) ?? string.Empty;
                            lines.Add($"{key}: {val}");
                            if (MaxLines > 0 && lines.Count >= MaxLines)
                                break;
                        }

                        _conduit.Update(_lastScreenPt, lines, TextColor, FontSize);
                        if (!_conduit.Enabled || _lastShownObjectId != rhObj.Id || _lastShownPt != _lastScreenPt)
                        {
                            _conduit.Enabled = true;
                            _lastShownObjectId = rhObj.Id;
                            _lastShownPt = _lastScreenPt;
                            _lastView.Redraw();
                        }
                        return;
                    }
                }
            }

            HideConduit();
        }

        private void HideConduit()
        {
            if (_conduit.Enabled)
            {
                _conduit.Enabled = false;
                _lastShownObjectId = Guid.Empty;
                _lastView?.Redraw();
            }
        }
    }

    internal class HoverTooltipConduit : DisplayConduit
    {
        private System.Drawing.Point _screenPoint;
        private List<string> _lines;
        private System.Drawing.Color _textColor;
        private int _fontSize;

        public void Update(System.Drawing.Point screenPoint, IEnumerable<string> lines, System.Drawing.Color textColor, int fontSize)
        {
            _screenPoint = screenPoint;
            _lines = lines != null ? new List<string>(lines) : null;
            _textColor = textColor;
            _fontSize = fontSize;
        }

        protected override void DrawForeground(DrawEventArgs e)
        {
            if (_lines == null || _lines.Count == 0)
                return;

            var start = new Point2d(_screenPoint.X + 16, _screenPoint.Y + 16);
            var lineHeight = _fontSize + 2; // Add some spacing between lines

            // Draw text with selected color and font size
            for (int i = 0; i < _lines.Count; i++)
            {
                var pt = new Point2d(start.X, start.Y + i * lineHeight);
                e.Display.Draw2dText(_lines[i], _textColor, pt, true, _fontSize);
            }
        }
    }
}