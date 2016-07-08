using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using GemBox.Spreadsheet;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Web;
using GuildfordBoroughCouncil.Security;

namespace GuildfordBoroughCouncil.Linq
{
    public static class Spreadsheet
    {
        private static ExcelFile GenerateSpreadsheet<T>(IEnumerable<T> list = null, InformationProtectiveMarking.Distribution Distribution = InformationProtectiveMarking.Distribution.Internal, string Title = "")
        {
            SpreadsheetInfo.SetLicense(Properties.Settings.Default.GemBoxSpreadsheetLicenseKey);
            SpreadsheetInfo.FreeLimitReached += (object sender, FreeLimitEventArgs e) => { e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial; };

            var ef = new ExcelFile();

            ef.DocumentProperties.BuiltIn[BuiltInDocumentProperties.Title] = Title;

            ef.DocumentProperties.Custom.Add("bjDocumentLabelXML", @"<?xml version=""1.0"" encoding=""us-ascii""?><sisl xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" sislVersion=""0"" policy=""b62b33be-a5f6-4f9d-86d0-ef32eac2404f"" xmlns=""http://www.boldonjames.com/2008/01/sie/i");
            ef.DocumentProperties.Custom.Add("bjDocumentLabelXML-0", @"nternal/label""><element uid=""id_protective_marking_new_item_1"" value="""" /><element uid=""id_distribution_newvalue1"" value="""" /></sisl>");

            ef.DocumentProperties.Custom.Add("bjDocumentSecurityLabel", "Guildford Borough Council UNCLASSIFIED INTERNAL");

            var ws = ef.Worksheets.Add("Sheet1");

            #region Header and footer

            var hf = ws.HeadersFooters;

            hf.FirstPage.Header.CenterSection.Content = Title;
            hf.DefaultPage.Header = hf.FirstPage.Header;

            hf.FirstPage.Footer.CenterSection.Append("Page ").Append(HeaderFooterFieldType.PageNumber).Append(" of ").Append(HeaderFooterFieldType.NumberOfPages);
            hf.DefaultPage.Footer = hf.FirstPage.Footer;

            #endregion

            #region Print options

            var printOptions = ws.PrintOptions;
            printOptions.Portrait = false;

            #endregion

            PropertyInfo[] properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

            #region Header row

            var i = 0;

            var DataFormat = new Dictionary<string, SpreadsheetDisplayFormatAttribute>();

            foreach (PropertyInfo p in properties)
            {
                if (!p.CanRead) { continue; }

                string displayName = p.Name;

                var da = p.GetCustomAttributes(typeof(DisplayAttribute), true).Cast<DisplayAttribute>().SingleOrDefault();

                if (da != null)
                {
                    displayName = da.Name;
                }
                
                ws.Cells[0, i].Value = displayName;

                var df = p.GetCustomAttributes(typeof(SpreadsheetDisplayFormatAttribute), true).Cast<SpreadsheetDisplayFormatAttribute>().SingleOrDefault();

                if (df != null)
                {
                    DataFormat.Add(p.Name, df);
                }

                i++;
            }

            var HeaderStyle = new CellStyle();
            HeaderStyle.HorizontalAlignment = HorizontalAlignmentStyle.Left;
            HeaderStyle.FillPattern.SetSolid(ColorTranslator.FromHtml("#538DD5"));
            HeaderStyle.Font.Weight = ExcelFont.BoldWeight;
            HeaderStyle.Font.Color = Color.White;
            //HeaderStyle.Borders.SetBorders(MultipleBorders.Right | MultipleBorders.Top, Color.Black, LineStyle.Thin);

            ws.Cells.GetSubrangeAbsolute(0, 0, 0, Math.Max(i - 1, 0)).Style = HeaderStyle;
            // Freeze top row
            ws.Panes = new WorksheetPanes(PanesState.Frozen, 0, 1, "A2", PanePosition.BottomLeft);

            #endregion

            if (list != null)
            {
                var rowId = 1;

                foreach (T row in list)
                {
                    var rowData = row.GetType();
                    var cellId = 0;

                    foreach (PropertyInfo p in properties)
                    {
                        if (!p.CanRead) { continue; }

                        var DataFormatString = "";

                        var value = rowData.GetProperty(p.Name).GetValue(row);

                        var cell = ws.Cells[rowId, cellId];

                        if (DataFormat.ContainsKey(p.Name))
                        {
                            var da = DataFormat[p.Name];

                            DataFormatString = da.DataFormatString;
                            
                            if (value == null)
                            {
                                value = da.NullDisplayText;
                            }

                            if (p.PropertyType == typeof(string))
                            {
                                if ((string)value == "")
                                {
                                    value = null;
                                }
                            }
                        }

                        if (p.PropertyType.ToString().Contains("System.Int64") && string.IsNullOrWhiteSpace(DataFormatString))
                        {
                            cell.Style = new CellStyle { NumberFormat = "0" };
                        }
                        else
                        {
                            cell.Style = new CellStyle { NumberFormat = DataFormatString };
                        }

                        cell.Value = value;

                        cellId++;
                    }

                    rowId++;
                }
            }

            if (ws.Rows.Count > 0)
            {
                // Autofit columns
                for (i = 0; i < ws.CalculateMaxUsedColumns(); i++)
                {
                    ws.Columns[i].AutoFit(1.5, ws.Rows[0], ws.Rows[ws.Rows.Count - 1]);
                }
            }

            return ef;
        }

        public static ExcelFile GetSpreadsheet<T>(IEnumerable<T> list = null, string fileName = "report.ods", string Title = "")
        {
            return GenerateSpreadsheet(list, Title: Title);
        }

        public static void AsSpreadsheet<T>(this HttpResponseBase httpResponse, IEnumerable<T> list = null, string fileName = "report.ods", string Title = "")
        {
            var ef = GenerateSpreadsheet(list, Title: Title);

            ef.Save(httpResponse, fileName);
        }

        public static void ToSpreadsheet<T>(this IEnumerable<T> list, string fileName = "report.ods", string Title = "")
        {
            var ef = GenerateSpreadsheet(list, Title: Title);
            
            ef.Save(fileName);
        }

        #region IDictionary

        private static ExcelFile GenerateSpreadsheet(IEnumerable<IDictionary<string, object>> list = null, InformationProtectiveMarking.Distribution Distribution = InformationProtectiveMarking.Distribution.Internal, string Title = "")
        {
            if (list == null || list.Count() == 0)
            {
                return null;
            }

            SpreadsheetInfo.SetLicense(Properties.Settings.Default.GemBoxSpreadsheetLicenseKey);
            SpreadsheetInfo.FreeLimitReached += (object sender, FreeLimitEventArgs e) => { e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial; };

            var ef = new ExcelFile();

            ef.DocumentProperties.BuiltIn[BuiltInDocumentProperties.Title] = Title;

            ef.DocumentProperties.Custom.Add("bjDocumentLabelXML", @"<?xml version=""1.0"" encoding=""us-ascii""?><sisl xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" sislVersion=""0"" policy=""b62b33be-a5f6-4f9d-86d0-ef32eac2404f"" xmlns=""http://www.boldonjames.com/2008/01/sie/i");
            ef.DocumentProperties.Custom.Add("bjDocumentLabelXML-0", @"nternal/label""><element uid=""id_protective_marking_new_item_1"" value="""" /><element uid=""id_distribution_newvalue1"" value="""" /></sisl>");

            ef.DocumentProperties.Custom.Add("bjDocumentSecurityLabel", "Guildford Borough Council UNCLASSIFIED INTERNAL");

            var ws = ef.Worksheets.Add("Sheet1");

            #region Header and footer

            var hf = ws.HeadersFooters;

            hf.FirstPage.Header.CenterSection.Content = Title;
            hf.DefaultPage.Header = hf.FirstPage.Header;

            hf.FirstPage.Footer.CenterSection.Append("Page ").Append(HeaderFooterFieldType.PageNumber).Append(" of ").Append(HeaderFooterFieldType.NumberOfPages);
            hf.DefaultPage.Footer = hf.FirstPage.Footer;

            #endregion

            #region Print options

            var printOptions = ws.PrintOptions;
            printOptions.Portrait = false;

            #endregion

            #region Header row

            var i = 0;

            var DataFormat = new Dictionary<string, SpreadsheetDisplayFormatAttribute>();

            foreach (var d in list.First())
            {
                string displayName = d.Key;
                
                ws.Cells[0, i].Value = displayName;

                i++;
            }

            var HeaderStyle = new CellStyle();
            HeaderStyle.HorizontalAlignment = HorizontalAlignmentStyle.Left;
            HeaderStyle.FillPattern.SetSolid(ColorTranslator.FromHtml("#538DD5"));
            HeaderStyle.Font.Weight = ExcelFont.BoldWeight;
            HeaderStyle.Font.Color = Color.White;

            ws.Cells.GetSubrangeAbsolute(0, 0, 0, Math.Max(i - 1, 0)).Style = HeaderStyle;
            // Freeze top row
            ws.Panes = new WorksheetPanes(PanesState.Frozen, 0, 1, "A2", PanePosition.BottomLeft);

            #endregion

            if (list != null)
            {
                var rowId = 1;

                foreach (var row in list)
                {
                    var cellId = 0;

                    foreach (var k in row.Keys)
                    {
                        if (row.ContainsKey(k))
                        {
                            var value = row[k];

                            var cell = ws.Cells[rowId, cellId];

                            if (value != null && row[k].GetType().ToString().Contains("System.Int64"))
                            {
                                cell.Style = new CellStyle { NumberFormat = "0" };
                            }

                            cell.Value = value;
                        }
                        cellId++;
                    }

                    rowId++;
                }
            }

            if (ws.Rows.Count > 0)
            {
                // Autofit columns
                for (i = 0; i < ws.CalculateMaxUsedColumns(); i++)
                {
                    ws.Columns[i].AutoFit(1.5, ws.Rows[0], ws.Rows[ws.Rows.Count - 1]);
                }
            }

            return ef;
        }

        public static ExcelFile GetSpreadsheet(IEnumerable<IDictionary<string, object>> list = null, string fileName = "report.ods", string Title = "")
        {
            return GenerateSpreadsheet(list, Title: Title);
        }

        public static void AsSpreadsheet(this HttpResponseBase httpResponse, IEnumerable<IDictionary<string, object>> list = null, string fileName = "report.ods", string Title = "")
        {
            var ef = GenerateSpreadsheet(list, Title: Title);

            ef.Save(httpResponse, fileName);
        }

        public static void ToSpreadsheet(this IEnumerable<IDictionary<string, object>> list, string fileName = "report.ods", string Title = "")
        {
            var ef = GenerateSpreadsheet(list, Title: Title);

            ef.Save(fileName);
        }

        #endregion
    }
}
