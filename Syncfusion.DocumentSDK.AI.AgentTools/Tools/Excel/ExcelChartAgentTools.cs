#region Copyright Syncfusion Inc. 2001 - 2026
//
//  Copyright Syncfusion Inc. 2001 - 2026. All rights reserved.
//
//  Use of this code is subject to the terms of our license.
//  A copy of the current license can be obtained at any time by e-mailing
//  licensing@syncfusion.com. Re-distribution in any form is strictly
//  prohibited. Any infringement will be prosecuted under applicable laws. 
//
#endregion

using System;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Syncfusion.AI.AgentTools.Core;
using Syncfusion.XlsIO;
using Syncfusion.Drawing;

namespace Syncfusion.AI.AgentTools.Excel
{
    /// <summary>
    /// Provides AI agent tools for Excel chart management operations.
    /// Handles chart creation, modification, customization, and removal.
    /// </summary>
    public class ExcelChartAgentTools : AgentToolBase<IWorkbook>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelChartAgentTools"/> class (Mode 1 — InMemory).
        /// </summary>
        /// <param name="manager">The Excel workbook manager.</param>
        public ExcelChartAgentTools(ExcelWorkbookManager manager)
            : base(manager, DocumentType.Excel) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelChartAgentTools"/> class (Mode 2 — DocumentStorage).
        /// </summary>
        /// <param name="manager">The document storage manager.</param>
        public ExcelChartAgentTools(DocumentStorageManager manager)
            : base(manager, DocumentType.Excel) { }

        // Static cache of normalized names -> enum value to avoid per-call allocations
        private static readonly Dictionary<string, ExcelChartType> s_nameMap;

        // Build the name map once using enum names and common variants.
        static ExcelChartAgentTools()
        {
            s_nameMap = new Dictionary<string, ExcelChartType>(StringComparer.OrdinalIgnoreCase);

            foreach (var name in Enum.GetNames(typeof(ExcelChartType)))
            {
                if (!Enum.TryParse<ExcelChartType>(name, out var val))
                    continue;

                // canonical enum name as key
                AddKey(name, val);

                // underscore -> space (Column_Clustered -> column clustered)
                AddKey(name.Replace("_", " "), val);

                // remove underscores (PieOfPie)
                AddKey(name.Replace("_", ""), val);

                // lowercase, spaces normalized are handled by AddKey

                // handle 3D variants: allow "3d" with/without space
                if (name.EndsWith("_3D", StringComparison.OrdinalIgnoreCase))
                {
                    var baseName = name.Substring(0, name.Length - 3);
                    AddKey(baseName + " 3d", val);
                    AddKey(baseName + "3d", val);
                }

                // handle 100 vs 100% (e.g., Column_Stacked_100)
                if (name.IndexOf("100", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    AddKey(name.Replace("100", "100%"), val);
                    AddKey(name.Replace("100", " percent"), val);
                }

                // swap two-word phrases to support both "Clustered Column" and "Column Clustered"
                var words = name.Replace("_", " ").Split(' ', StringSplitOptions.RemoveEmptyEntries);
                if (words.Length == 2)
                {
                    AddKey(words[1] + " " + words[0], val);
                }
            }

            // Manual synonyms for ambiguous/common friendly names
            void addSyn(string friendly, string enumName)
            {
                if (Enum.TryParse<ExcelChartType>(enumName, true, out var v))
                    AddKey(friendly, v);
            }

            addSyn("clustered column", "Column_Clustered");
            addSyn("stacked column", "Column_Stacked");
            addSyn("100 stacked column", "Column_Stacked_100");
            addSyn("100% stacked column", "Column_Stacked_100");
            addSyn("clustered column 3d", "Column_Clustered_3D");
            addSyn("clustered bar", "Bar_Clustered");
            addSyn("stacked bar", "Bar_Stacked");
            addSyn("100% stacked bar", "Bar_Stacked_100");
            addSyn("line markers", "Line_Markers");
            addSyn("pie of pie", "PieOfPie");
            addSyn("bar of pie", "Pie_Bar");
            addSyn("pie exploded", "Pie_Exploded");
            addSyn("pie 3d", "Pie_3D");
            addSyn("area stacked 100", "Area_Stacked_100");
            addSyn("scatter markers", "Scatter_Markers");
            addSyn("column 3d", "Column_3D");
            addSyn("doughnut exploded", "Doughnut_Exploded");
            addSyn("doughnut", "Doughnut");
            addSyn("radar markers", "Radar_Markers");
            addSyn("radar filled", "Radar_Filled");
            addSyn("surface 3d", "Surface_3D");
            addSyn("surface contour", "Surface_Contour");
            addSyn("bubble", "Bubble");
            addSyn("bubble 3d", "Bubble_3D");
            addSyn("stock high low close", "Stock_HighLowClose");
            addSyn("stock open high low close", "Stock_OpenHighLowClose");
            addSyn("cylinder clustered", "Cylinder_Clustered");
            addSyn("cone clustered", "Cone_Clustered");
            addSyn("pyramid clustered", "Pyramid_Clustered");
            addSyn("combination chart", "Combination_Chart");
            addSyn("funnel", "Funnel");
            addSyn("waterfall", "WaterFall");
            addSyn("box and whisker", "BoxAndWhisker");
            addSyn("histogram", "Histogram");
            addSyn("pareto", "Pareto");
            addSyn("treemap", "TreeMap");
            addSyn("sunburst", "SunBurst");
        }

        private static void AddKey(string raw, ExcelChartType value)
        {
            if (string.IsNullOrWhiteSpace(raw))
                return;
            // normalize key: lowercase, collapse non-alphanum to single space, trim
            string key = Regex.Replace(raw.ToLowerInvariant(), "[^a-z0-9]+", " ").Trim();
            if (!s_nameMap.ContainsKey(key))
                s_nameMap[key] = value;
        }

        private ExcelChartType? GetChartTypeFromFriendlyName(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return null;

            string s = input.Trim().ToLowerInvariant();

            // Pie charts
            if (s.Contains("pie"))
            {
                if (s.Contains("bar")) return ExcelChartType.Pie_Bar;
                if (s.Contains("of")) return ExcelChartType.PieOfPie;
                if (s.Contains("exploded") && s.Contains("3d")) return ExcelChartType.Pie_Exploded_3D;
                if (s.Contains("exploded")) return ExcelChartType.Pie_Exploded;
                if (s.Contains("3d")) return ExcelChartType.Pie_3D;
                return ExcelChartType.Pie;
            }

            // Column charts
            if (s.Contains("column"))
            {
                if (s.Contains("clustered") && s.Contains("3d")) return ExcelChartType.Column_Clustered_3D;
                if (s.Contains("stacked") && s.Contains("3d") && s.Contains("100")) return ExcelChartType.Column_Stacked_100_3D;
                if (s.Contains("stacked") && s.Contains("3d")) return ExcelChartType.Column_Stacked_3D;
                if (s.Contains("stacked") && s.Contains("100")) return ExcelChartType.Column_Stacked_100;
                if (s.Contains("stacked")) return ExcelChartType.Column_Stacked;
                if (s.Contains("3d")) return ExcelChartType.Column_3D;
                return ExcelChartType.Column_Clustered;
            }

            // Bar charts
            if (s.Contains("bar"))
            {
                if (s.Contains("clustered") && s.Contains("3d")) return ExcelChartType.Bar_Clustered_3D;
                if (s.Contains("stacked") && s.Contains("3d") && s.Contains("100")) return ExcelChartType.Bar_Stacked_100_3D;
                if (s.Contains("stacked") && s.Contains("3d")) return ExcelChartType.Bar_Stacked_3D;
                if (s.Contains("stacked") && s.Contains("100")) return ExcelChartType.Bar_Stacked_100;
                if (s.Contains("stacked")) return ExcelChartType.Bar_Stacked;
                return ExcelChartType.Bar_Clustered;
            }

            // Line charts
            if (s.Contains("line"))
            {
                if (s.Contains("markers") && s.Contains("stacked") && s.Contains("100")) return ExcelChartType.Line_Markers_Stacked_100;
                if (s.Contains("markers") && s.Contains("stacked")) return ExcelChartType.Line_Markers_Stacked;
                if (s.Contains("markers")) return ExcelChartType.Line_Markers;
                if (s.Contains("stacked") && s.Contains("100")) return ExcelChartType.Line_Stacked_100;
                if (s.Contains("stacked")) return ExcelChartType.Line_Stacked;
                if (s.Contains("3d")) return ExcelChartType.Line_3D;
                return ExcelChartType.Line;
            }

            // Area charts
            if (s.Contains("area"))
            {
                if (s.Contains("stacked") && s.Contains("3d") && s.Contains("100")) return ExcelChartType.Area_Stacked_100_3D;
                if (s.Contains("stacked") && s.Contains("3d")) return ExcelChartType.Area_Stacked_3D;
                if (s.Contains("stacked") && s.Contains("100")) return ExcelChartType.Area_Stacked_100;
                if (s.Contains("stacked")) return ExcelChartType.Area_Stacked;
                if (s.Contains("3d")) return ExcelChartType.Area_3D;
                return ExcelChartType.Area;
            }

            // Doughnut charts
            if (s.Contains("doughnut"))
            {
                if (s.Contains("exploded")) return ExcelChartType.Doughnut_Exploded;
                return ExcelChartType.Doughnut;
            }

            // Radar charts
            if (s.Contains("radar"))
            {
                if (s.Contains("filled")) return ExcelChartType.Radar_Filled;
                if (s.Contains("markers")) return ExcelChartType.Radar_Markers;
                return ExcelChartType.Radar;
            }

            // Scatter charts
            if (s.Contains("scatter"))
            {
                if (s.Contains("smoothed") && s.Contains("markers")) return ExcelChartType.Scatter_SmoothedLine_Markers;
                if (s.Contains("smoothed")) return ExcelChartType.Scatter_SmoothedLine;
                if (s.Contains("line") && s.Contains("markers")) return ExcelChartType.Scatter_Line_Markers;
                if (s.Contains("line")) return ExcelChartType.Scatter_Line;
                if (s.Contains("markers")) return ExcelChartType.Scatter_Markers;
            }

            // Bubble charts
            if (s.Contains("bubble"))
            {
                if (s.Contains("3d")) return ExcelChartType.Bubble_3D;
                return ExcelChartType.Bubble;
            }

            // Stock charts
            if (s.Contains("stock"))
            {
                if (s.Contains("volume") && s.Contains("open")) return ExcelChartType.Stock_VolumeOpenHighLowClose;
                if (s.Contains("volume") && s.Contains("high")) return ExcelChartType.Stock_VolumeHighLowClose;
                if (s.Contains("open")) return ExcelChartType.Stock_OpenHighLowClose;
                return ExcelChartType.Stock_HighLowClose;
            }

            // Cylinder charts
            if (s.Contains("cylinder"))
            {
                if (s.Contains("bar") && s.Contains("stacked") && s.Contains("100")) return ExcelChartType.Cylinder_Bar_Stacked_100;
                if (s.Contains("bar") && s.Contains("stacked")) return ExcelChartType.Cylinder_Bar_Stacked;
                if (s.Contains("bar") && s.Contains("clustered")) return ExcelChartType.Cylinder_Bar_Clustered;
                if (s.Contains("stacked") && s.Contains("100")) return ExcelChartType.Cylinder_Stacked_100;
                if (s.Contains("stacked")) return ExcelChartType.Cylinder_Stacked;
                if (s.Contains("clustered") && s.Contains("3d")) return ExcelChartType.Cylinder_Clustered_3D;
                return ExcelChartType.Cylinder_Clustered;
            }

            // Cone charts
            if (s.Contains("cone"))
            {
                if (s.Contains("bar") && s.Contains("stacked") && s.Contains("100")) return ExcelChartType.Cone_Bar_Stacked_100;
                if (s.Contains("bar") && s.Contains("stacked")) return ExcelChartType.Cone_Bar_Stacked;
                if (s.Contains("bar") && s.Contains("clustered")) return ExcelChartType.Cone_Bar_Clustered;
                if (s.Contains("stacked") && s.Contains("100")) return ExcelChartType.Cone_Stacked_100;
                if (s.Contains("stacked")) return ExcelChartType.Cone_Stacked;
                if (s.Contains("clustered") && s.Contains("3d")) return ExcelChartType.Cone_Clustered_3D;
                return ExcelChartType.Cone_Clustered;
            }

            // Pyramid charts
            if (s.Contains("pyramid"))
            {
                if (s.Contains("bar") && s.Contains("stacked") && s.Contains("100")) return ExcelChartType.Pyramid_Bar_Stacked_100;
                if (s.Contains("bar") && s.Contains("stacked")) return ExcelChartType.Pyramid_Bar_Stacked;
                if (s.Contains("bar") && s.Contains("clustered")) return ExcelChartType.Pyramid_Bar_Clustered;
                if (s.Contains("stacked") && s.Contains("100")) return ExcelChartType.Pyramid_Stacked_100;
                if (s.Contains("stacked")) return ExcelChartType.Pyramid_Stacked;
                if (s.Contains("clustered") && s.Contains("3d")) return ExcelChartType.Pyramid_Clustered_3D;
                return ExcelChartType.Pyramid_Clustered;
            }

            // Surface charts
            if (s.Contains("surface"))
            {
                if (s.Contains("noc") && s.Contains("3d")) return ExcelChartType.Surface_NoColor_3D;
                if (s.Contains("noc") && s.Contains("contour")) return ExcelChartType.Surface_NoColor_Contour;
                if (s.Contains("contour")) return ExcelChartType.Surface_Contour;
                if (s.Contains("3d")) return ExcelChartType.Surface_3D;
            }

            // Funnel, Waterfall, Box & Whisker, Histogram, Pareto, TreeMap, SunBurst
            if (s.Contains("funnel")) return ExcelChartType.Funnel;
            if (s.Contains("waterfall")) return ExcelChartType.WaterFall;
            if (s.Contains("box") || s.Contains("whisker")) return ExcelChartType.BoxAndWhisker;           
            if (s.Contains("histogram")) return ExcelChartType.Histogram;
            if (s.Contains("pareto")) return ExcelChartType.Pareto;
            if (s.Contains("treemap")) return ExcelChartType.TreeMap;
            if (s.Contains("sunburst")) return ExcelChartType.SunBurst;

            return null;
        }

        // Attempts to parse a user-friendly chart type string into the
        // strongly-typed `ExcelChartType` enum used by XlsIO.
        private bool TryGetChartType(string input, out ExcelChartType chartType)
        {
            chartType = default;
            if (string.IsNullOrWhiteSpace(input))
                return false;

            string s = input.Trim();

            // 1) Direct parse (case-insensitive)
            if (Enum.TryParse<ExcelChartType>(s, true, out chartType))
                return true;

            // 2) Try normalized enum-like variants
            string variant = s.Replace(" ", "_").Replace("-", "_").Replace("%", "");
            if (Enum.TryParse<ExcelChartType>(variant, true, out chartType))
                return true;

            // 3) Remove all spaces (helps with names like PieOfPie)
            variant = Regex.Replace(s, " ", "");
            if (Enum.TryParse<ExcelChartType>(variant, true, out chartType))
                return true;

            // 4) Lookup in prebuilt name map
            string key = Regex.Replace(s.ToLowerInvariant(), "[^a-z0-9]+", " ").Trim();
            if (s_nameMap.TryGetValue(key, out chartType))
                return true;

            return false;
        }

        private static string Capitalize(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            if (s.Length == 1) return s.ToUpperInvariant();
            return char.ToUpperInvariant(s[0]) + s.Substring(1);
        }

        /// <summary>
        /// Creates a chart in the specified worksheet from a data range.
        /// </summary>
        /// <param name="workbookIdOrFilePath">The workbook ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="worksheetName">The name of the worksheet.</param>
        /// <param name="chartType">The type of chart (e.g., Column_Clustered, Line, Pie, Bar_Clustered).</param>
        /// <param name="dataRange">The data range for the chart (e.g., "A1:C6").</param>
        /// <param name="isSeriesInRows">True if series are in rows, false if in columns.</param>
        /// <param name="topRow">Optional top row for chart positioning (default: 8).</param>
        /// <param name="leftColumn">Optional left column for chart positioning (default: 1).</param>
        /// <param name="bottomRow">Optional bottom row for chart positioning (default: 23).</param>
        /// <param name="rightColumn">Optional right column for chart positioning (default: 8).</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result containing the chart index.</returns>
        [Tool(Name = "CreateChart", Description = "Creates a chart from a data range in the worksheet. Supports various chart types like Column_Clustered, Line, Pie, Bar_Clustered, Area, etc. workbookIdOrFilePath: The workbook ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult CreateChart(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The name of the worksheet")] string worksheetName,
            [ToolParameter(Description = "Chart type (e.g., Column_Clustered, Line, Pie, Bar_Clustered, Area, Column_Stacked, Line_Markers)")] string chartType,
            [ToolParameter(Description = "Data range for the chart (e.g., A1:C6)")] string dataRange,
            [ToolParameter(Description = "True if series are in rows, false if in columns (default: false)")] bool isSeriesInRows = false,
            [ToolParameter(Description = "Top row for chart positioning (default: 8)")] int topRow = 8,
            [ToolParameter(Description = "Left column for chart positioning (default: 1)")] int leftColumn = 1,
            [ToolParameter(Description = "Bottom row for chart positioning (default: 23)")] int bottomRow = 23,
            [ToolParameter(Description = "Right column for chart positioning (default: 8)")] int rightColumn = 8,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")] string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);
                ArgumentNullException.ThrowIfNull(worksheetName);
                ArgumentNullException.ThrowIfNull(chartType);
                ArgumentNullException.ThrowIfNull(dataRange);

                var workbook = OpenDocument(workbookIdOrFilePath);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                var worksheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == worksheetName);
                if (worksheet == null)
                    return AgentToolResult.Fail($"Worksheet not found: {worksheetName}");

                // Parse chart type (accepts friendly names)
                if (!TryGetChartType(chartType, out var excelChartType))
                {
                    ExcelChartType? excelFrienlyChartType = GetChartTypeFromFriendlyName(chartType);
                    if (excelFrienlyChartType != null)
                        excelChartType = excelFrienlyChartType.Value;
                    else
                        return AgentToolResult.Fail($"Invalid chart type: {chartType}. Examples: Column_Clustered, Line, Pie, Bar_Clustered, Area");
                }

                // Create chart
                IChartShape chart = worksheet.Charts.Add();
                chart.ChartType = excelChartType;
                chart.DataRange = worksheet.Range[dataRange];
                chart.IsSeriesInRows = isSeriesInRows;

                // Position the chart
                chart.TopRow = topRow;
                chart.LeftColumn = leftColumn;
                chart.BottomRow = bottomRow;
                chart.RightColumn = rightColumn;

                // Get chart index by counting charts (since Index property doesn't exist)
                int chartIndex = worksheet.Charts.Count - 1;
                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_chart.xlsx";
                    
                string outputKey = outputFilePath;
                SaveDocument(outputKey, workbook);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = workbookIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok(
                    $"Chart created successfully in worksheet '{worksheetName}' at position (Row {topRow}, Col {leftColumn}) into document {outputKey}",
                    new { ChartIndex = chartIndex, ChartType = chartType, DataRange = dataRange, OutputKey = outputKey });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to create chart: {ex.Message}");
            }
        }

        /// <summary>
        /// Creates a chart by adding series one by one.
        /// </summary>
        /// <param name="workbookIdOrFilePath">The workbook ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="worksheetName">The name of the worksheet.</param>
        /// <param name="chartType">The type of chart (e.g., Line, Column_Clustered, Bar_Clustered).</param>
        /// <param name="seriesName">The name of the series.</param>
        /// <param name="valuesRange">The range containing the values (e.g., "B2:B6").</param>
        /// <param name="categoryLabelsRange">The range containing category labels (e.g., "A2:A6").</param>
        /// <param name="topRow">Optional top row for chart positioning (default: 8).</param>
        /// <param name="leftColumn">Optional left column for chart positioning (default: 1).</param>
        /// <param name="bottomRow">Optional bottom row for chart positioning (default: 23).</param>
        /// <param name="rightColumn">Optional right column for chart positioning (default: 8).</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result containing the chart index and series information.</returns>
        [Tool(Name = "CreateChartWithSeries", Description = "Creates a chart and adds a series with values and category labels. workbookIdOrFilePath: The workbook ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult CreateChartWithSeries(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The name of the worksheet")] string worksheetName,
            [ToolParameter(Description = "Chart type (e.g., Line, Column_Clustered, Bar_Clustered)")] string chartType,
            [ToolParameter(Description = "Name of the series")] string seriesName,
            [ToolParameter(Description = "Range containing values (e.g., B2:B6)")] string valuesRange,
            [ToolParameter(Description = "Range containing category labels (e.g., A2:A6)")] string categoryLabelsRange,
            [ToolParameter(Description = "Top row for chart positioning (default: 8)")] int topRow = 8,
            [ToolParameter(Description = "Left column for chart positioning (default: 1)")] int leftColumn = 1,
            [ToolParameter(Description = "Bottom row for chart positioning (default: 23)")] int bottomRow = 23,
            [ToolParameter(Description = "Right column for chart positioning (default: 8)")] int rightColumn = 8,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")] string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);
                ArgumentNullException.ThrowIfNull(worksheetName);
                ArgumentNullException.ThrowIfNull(chartType);
                ArgumentNullException.ThrowIfNull(seriesName);
                ArgumentNullException.ThrowIfNull(valuesRange);
                ArgumentNullException.ThrowIfNull(categoryLabelsRange);

                var workbook = OpenDocument(workbookIdOrFilePath);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                var worksheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == worksheetName);
                if (worksheet == null)
                    return AgentToolResult.Fail($"Worksheet not found: {worksheetName}");

                // Parse chart type (accepts friendly names)
                if (!TryGetChartType(chartType, out var excelChartType))
                {
                    ExcelChartType? excelFrienlyChartType = GetChartTypeFromFriendlyName(chartType);
                    if (excelFrienlyChartType != null)
                        excelChartType = excelFrienlyChartType.Value;
                    else
                        return AgentToolResult.Fail($"Invalid chart type: {chartType}. Examples: Column_Clustered, Line, Pie, Bar_Clustered, Area");
                }

                // Create chart
                IChartShape chart = worksheet.Charts.Add();
                chart.ChartType = excelChartType;

                // Add series
                IChartSerie series = chart.Series.Add(seriesName);
                series.Values = worksheet.Range[valuesRange];
                series.CategoryLabels = worksheet.Range[categoryLabelsRange];

                // Position the chart
                chart.TopRow = topRow;
                chart.LeftColumn = leftColumn;
                chart.BottomRow = bottomRow;
                chart.RightColumn = rightColumn;

                // Get chart index by counting charts
                int chartIndex = worksheet.Charts.Count - 1;

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_chart_with_series.xlsx";

                string outputKey = outputFilePath;
                SaveDocument(outputKey, workbook);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = workbookIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok(
                    $"Chart with series '{seriesName}' created successfully in worksheet '{worksheetName}' into document {outputKey}",
                    new { ChartIndex = chartIndex, SeriesName = seriesName, ChartType = chartType, OutputKey = outputKey });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to create chart with series: {ex.Message}");
            }
        }

        /// <summary>
        /// Adds a series to an existing chart.
        /// </summary>
        /// <param name="workbookIdOrFilePath">The workbook ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="worksheetName">The name of the worksheet.</param>
        /// <param name="chartIndex">The index of the chart (0-based).</param>
        /// <param name="seriesName">The name of the series to add.</param>
        /// <param name="valuesRange">The range containing the values (e.g., "C2:C6").</param>
        /// <param name="categoryLabelsRange">The range containing category labels (e.g., "A2:A6").</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(Name = "AddSeriesToChart", Description = "Adds a new series to an existing chart. workbookIdOrFilePath: The workbook ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult AddSeriesToChart(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The name of the worksheet")] string worksheetName,
            [ToolParameter(Description = "Index of the chart (0-based)")] int chartIndex,
            [ToolParameter(Description = "Name of the series to add")] string seriesName,
            [ToolParameter(Description = "Range containing values (e.g., C2:C6)")] string valuesRange,
            [ToolParameter(Description = "Range containing category labels (e.g., A2:A6)")] string categoryLabelsRange = "",
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")] string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);
                ArgumentNullException.ThrowIfNull(worksheetName);
                ArgumentNullException.ThrowIfNull(seriesName);
                ArgumentNullException.ThrowIfNull(valuesRange);
                ArgumentNullException.ThrowIfNull(categoryLabelsRange);

                var workbook = OpenDocument(workbookIdOrFilePath);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                var worksheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == worksheetName);
                if (worksheet == null)
                    return AgentToolResult.Fail($"Worksheet not found: {worksheetName}");

                if (chartIndex < 0 || chartIndex >= worksheet.Charts.Count)
                    return AgentToolResult.Fail($"Chart index {chartIndex} is out of range");

                var chart = worksheet.Charts[chartIndex];

                // Add series
                IChartSerie series = chart.Series.Add(seriesName);
                series.Values = worksheet.Range[valuesRange];
                if (!string.IsNullOrEmpty(categoryLabelsRange))
                    series.CategoryLabels = worksheet.Range[categoryLabelsRange];
                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_add_series.xlsx";

                string outputKey = outputFilePath;
                SaveDocument(outputKey, workbook);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = workbookIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok(
                    $"Series '{seriesName}' added successfully to chart at index {chartIndex} into document {outputKey}",
                    new { SeriesName = seriesName, ChartIndex = chartIndex, OutputKey = outputKey });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to add series to chart: {ex.Message}");
            }
        }

        /// <summary>
        /// Sets the chart title.
        /// </summary>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(Name = "SetChartElement", Description = "Sets the elements of a chart. workbookIdOrFilePath: The workbook ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult SetChartElements(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The name of the worksheet")] string worksheetName,
            [ToolParameter(Description = "Index of the chart (0-based)")] int chartIndex,
            [ToolParameter(Description = "Index of the series (0-based)")] int seriesIndex,
            [ToolParameter(Description = "Title text for the chart")] string title,
            [ToolParameter(Description = "True to show legend, false to hide")] bool hasLegend,
            [ToolParameter(Description = "Legend position(ONLY exact values): Bottom, Top, Left, Right, Corner (default: Bottom)")] string position = "Bottom",
            [ToolParameter(Description = "Show values in data labels (default: true)")] bool showValue = true,
            [ToolParameter(Description = "Show category names in data labels (default: false)")] bool showCategoryName = false,
            [ToolParameter(Description = "Show series names in data labels (default: false)")] bool showSeriesName = false,
            [ToolParameter(Description = "Data label position: Outside, Inside, Center, etc. (default: Outside)")] string dataLabelposition = "Outside",
            [ToolParameter(Description = "Title for category (horizontal) axis")] string? categoryAxisTitle = null,
            [ToolParameter(Description = "Title for value (vertical) axis")] string? valueAxisTitle = null,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")] string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);
                ArgumentNullException.ThrowIfNull(worksheetName);
                ArgumentNullException.ThrowIfNull(title);

                var workbook = OpenDocument(workbookIdOrFilePath);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                var worksheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == worksheetName);
                if (worksheet == null)
                    return AgentToolResult.Fail($"Worksheet not found: {worksheetName}");

                if (chartIndex < 0 || chartIndex >= worksheet.Charts.Count)
                    return AgentToolResult.Fail($"Chart index {chartIndex} is out of range");

                var chart = worksheet.Charts[chartIndex];
                chart.ChartTitle = title;
                chart.HasLegend = hasLegend;

                if (hasLegend && Enum.TryParse<ExcelLegendPosition>(position, true, out var legendPosition))
                {
                    chart.Legend.Position = legendPosition;
                }
                ExcelChartType excelChartType = chart.ChartType;
                var series = chart.Series[seriesIndex];
                var dataLabels = series.DataPoints.DefaultDataPoint.DataLabels;

                dataLabels.IsValue = showValue;
                dataLabels.IsCategoryName = showCategoryName;
                dataLabels.IsSeriesName = showSeriesName;

                if (Enum.TryParse<ExcelDataLabelPosition>(dataLabelposition, true, out var labelPosition))
                {
                    if ((!excelChartType.ToString().ToLower().Contains("area") && excelChartType != ExcelChartType.TreeMap &&
                        excelChartType != ExcelChartType.SunBurst && !excelChartType.ToString().ToLower().Contains("radar") &&
                        excelChartType != ExcelChartType.Funnel && !excelChartType.ToString().ToLower().Contains("surface") &&
                        !excelChartType.ToString().ToLower().Contains("volume") && !excelChartType.ToString().ToLower().Contains("3d")) ||
                        excelChartType == ExcelChartType.Bubble_3D)
                    {
                        if (excelChartType == ExcelChartType.WaterFall || excelChartType.ToString().ToLower().Contains("column") ||
                            excelChartType.ToString().ToLower().Contains("bar") || excelChartType == ExcelChartType.Histogram ||
                            excelChartType == ExcelChartType.Pareto)
                        {
                            if (labelPosition == ExcelDataLabelPosition.Center || labelPosition == ExcelDataLabelPosition.Inside ||
                                (labelPosition == ExcelDataLabelPosition.Outside && !excelChartType.ToString().ToLower().Contains("stack")))
                            {
                                dataLabels.Position = labelPosition;
                            }
                        }
                        else if (excelChartType.ToString().ToLower().Contains("line") || excelChartType.ToString().ToLower().Contains("scatter") ||
                            excelChartType.ToString().ToLower().Contains("bubble") || excelChartType.ToString().ToLower().Contains("stock"))
                        {
                            if (labelPosition == ExcelDataLabelPosition.Center || labelPosition == ExcelDataLabelPosition.Left ||
                                labelPosition == ExcelDataLabelPosition.Right || labelPosition == ExcelDataLabelPosition.Above ||
                                labelPosition == ExcelDataLabelPosition.Below)
                            {
                                dataLabels.Position = labelPosition;
                            }
                        }
                        else if (excelChartType.ToString().ToLower().Contains("pie"))
                        {
                            if (labelPosition == ExcelDataLabelPosition.Center || labelPosition == ExcelDataLabelPosition.Inside ||
                                labelPosition == ExcelDataLabelPosition.Outside || labelPosition == ExcelDataLabelPosition.BestFit)
                            {
                                dataLabels.Position = labelPosition;
                            }
                        }
                        else if (excelChartType == ExcelChartType.BoxAndWhisker)
                        {
                            if (labelPosition == ExcelDataLabelPosition.Left || labelPosition == ExcelDataLabelPosition.Right ||
                                labelPosition == ExcelDataLabelPosition.Above || labelPosition == ExcelDataLabelPosition.Below)
                            {
                                dataLabels.Position = labelPosition;
                            }
                        }
                    }
                }
                if (!string.IsNullOrEmpty(categoryAxisTitle))
                {
                    chart.PrimaryCategoryAxis.Title = categoryAxisTitle;
                }
				if (!string.IsNullOrEmpty(valueAxisTitle))
                {
                    chart.PrimaryValueAxis.Title = valueAxisTitle;
                }
                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_set_chart_elements.xlsx";

                string outputKey = outputFilePath;
                SaveDocument(outputKey, workbook);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = workbookIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok($"Chart element set successfully into document {outputKey}", new { OutputKey = outputKey });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to set chart Elements: {ex.Message}");
            }
        }
    }
}
