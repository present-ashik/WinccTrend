using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.IO;
using System.Data;
using ExcelDataReader;
using System.Text;

namespace TrendControlExcel
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [DefaultEvent("PointAdded")]
    public class CustomTrendControl : UserControl
    {
        private Chart chart;

        public CustomTrendControl()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            chart = new Chart();
            chart.Dock = DockStyle.Fill;

            var area = new ChartArea("DefaultArea");
            area.AxisX.Title = "X";
            area.AxisY.Title = "Y";
            chart.ChartAreas.Add(area);

            chart.Legends.Add(new Legend("Legend"));
            this.Controls.Add(chart);
        }

        // Adds a series if not exists. seriesType: "Line", "Spline", etc.
        [ComVisible(true)]
        public void AddSeries(string seriesName, string seriesType)
        {
            if (string.IsNullOrWhiteSpace(seriesName)) return;
            if (chart.Series.IsUniqueName(seriesName) == false) return;
            var s = new Series(seriesName);
            s.ChartType = (SeriesChartType)Enum.Parse(typeof(SeriesChartType), seriesType ?? "Line");
            s.XValueType = ChartValueType.Double;
            s.YValueType = ChartValueType.Double;
            chart.Series.Add(s);
        }

        // Add a point to a series (create series if not present)
        [ComVisible(true)]
        public void AddPoint(string seriesName, double x, double y)
        {
            if (chart.Series.IsUniqueName(seriesName))
            {
                var s = new Series(seriesName);
                s.ChartType = SeriesChartType.Line;
                chart.Series.Add(s);
            }
            chart.Series[seriesName].Points.AddXY(x, y);
        }

        [ComVisible(true)]
        public void Clear()
        {
            chart.Series.Clear();
        }

        // Load Excel file (.xlsx or .xls) and plot columns.
        // Assumes first row is header. Plots first numeric column as X (or row index) and subsequent numeric columns as series.
        [ComVisible(true)]
        public bool LoadExcel(string filePath, int sheetIndex)
        {
            try
            {
                if (!File.Exists(filePath)) return false;
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var conf = new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                        };
                        var ds = reader.AsDataSet(conf);
                        if (sheetIndex < 0 || sheetIndex >= ds.Tables.Count) sheetIndex = 0;
                        var table = ds.Tables[sheetIndex];
                        if (table.Rows.Count == 0) return false;

                        // Clear existing series
                        chart.Series.Clear();

                        // Determine numeric columns
                        var numericCols = new System.Collections.Generic.List<int>();
                        for (int c = 0; c < table.Columns.Count; c++)
                        {
                            // check first data row for numeric value
                            double tmp;
                            if (table.Rows.Count > 0 && double.TryParse(Convert.ToString(table.Rows[0][c]), out tmp))
                                numericCols.Add(c);
                        }

                        // If no numeric columns found, attempt to plot using row index as X and first column as Y (if numeric)
                        if (numericCols.Count == 0)
                        {
                            for (int r = 0; r < table.Rows.Count; r++)
                            {
                                double y;
                                if (double.TryParse(Convert.ToString(table.Rows[r][0]), out y))
                                {
                                    string sname = table.Columns.Count > 1 ? table.Columns[1].ColumnName : "Series1";
                                    if (chart.Series.IsUniqueName(sname))
                                        chart.Series.Add(new Series(sname) { ChartType = SeriesChartType.Line });
                                    chart.Series[sname].Points.AddXY(r, y);
                                }
                            }
                        }
                        else
                        {
                            // Use first numeric column as X if more than one numeric column and user wants that; otherwise use row index as X
                            int xCol = numericCols[0];
                            // For each remaining numeric column, create a series
                            for (int i = 1; i < numericCols.Count; i++)
                            {
                                int colIdx = numericCols[i];
                                string sname = table.Columns[colIdx].ColumnName ?? ($"Series{colIdx}");
                                var s = new Series(sname) { ChartType = SeriesChartType.Line };
                                chart.Series.Add(s);
                                for (int r = 0; r < table.Rows.Count; r++)
                                {
                                    double xVal, yVal;
                                    string xStr = Convert.ToString(table.Rows[r][xCol]);
                                    string yStr = Convert.ToString(table.Rows[r][colIdx]);
                                    if (!double.TryParse(xStr, out xVal)) xVal = r;
                                    if (!double.TryParse(yStr, out yVal)) continue;
                                    s.Points.AddXY(xVal, yVal);
                                }
                            }
                        }
                    }
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        // Load CSV helper (simple)
        [ComVisible(true)]
        public bool LoadCSV(string filePath, char sep)
        {
            try
            {
                if (!File.Exists(filePath)) return false;
                var lines = System.IO.File.ReadAllLines(filePath);
                if (lines.Length == 0) return false;

                chart.Series.Clear();
                // assume header exists
                var headers = lines[0].Split(sep);
                int cols = headers.Length;
                for (int c = 1; c < cols; c++)
                {
                    var s = new Series(headers[c]) { ChartType = SeriesChartType.Line };
                    chart.Series.Add(s);
                }
                for (int r = 1; r < lines.Length; r++)
                {
                    var parts = lines[r].Split(sep);
                    double x = r - 1;
                    if (parts.Length > 0 && double.TryParse(parts[0], out double xv)) x = xv;
                    for (int c = 1; c < parts.Length && c <= chart.Series.Count; c++)
                    {
                        if (double.TryParse(parts[c], out double yv))
                            chart.Series[c-1].Points.AddXY(x, yv);
                    }
                }
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
