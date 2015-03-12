// <copyright file="ExcelWriter.cs" company="Axinesis">
//     Axinesis All rights reserved.
// </copyright>
// <author>Adrien Denis</author>

namespace ExcelWriterCSharp
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Drawing;
    using System.Linq;
    using System.Text;
    using Excel = Microsoft.Office.Interop.Excel;

    /// <summary>
    /// to do
    /// </summary>
    public class ExcelWriter : IDisposable
    {
        #region Fields

        /// <summary>
        /// to do
        /// </summary>
        private Excel.Application app;

        /// <summary>
        /// to do
        /// </summary>
        private Excel.Workbook book;

        /// <summary>
        /// to do
        /// </summary>
        private Excel.Worksheet sheetToWrite;

        /// <summary>
        /// to do
        /// </summary>
        private List<Excel.Worksheet> sheets = new List<Excel.Worksheet>();

        /// <summary>
        /// to do
        /// </summary>
        private string path;

        /// <summary>
        /// to do
        /// </summary>
        private List<string> sheetsNames;

        /// <summary>
        /// to do
        /// </summary>
        private ExcelConvert convert; 
        #endregion

        #region Ctor

        /// <summary>
        /// Initializes a new instance of the ExcelWriter class.
        /// </summary>
        /// <param name="path">to do</param>
        /// <param name="sheetName">to do 2</param>
        public ExcelWriter(string path, string sheetName)
        {
            this.path = path;

            this.sheetsNames = new List<string>();

            this.app = new Excel.Application();
            this.app.DisplayAlerts = false;
            this.book = this.app.Workbooks.Add();

            this.app.ActiveSheet.Name = sheetName;
            this.AddNameToSheetsNames(sheetName);
            this.sheets.Add(this.app.ActiveSheet);
            this.sheetToWrite = this.app.ActiveSheet;

            this.convert = new ExcelConvert();
        } 
        #endregion

        #region DCtor

        /// <summary>
        /// Finalizes an instance of the ExcelWriter class.
        /// </summary>
        ~ExcelWriter()
        {
            // Finalizer calls Dispose(false)
            this.Dispose(false);
        } 
        #endregion

        #region Properties

        /// <summary>
        /// Gets : to do
        /// </summary>
        public ReadOnlyCollection<string> SheetsNames
        {
            get { return this.sheetsNames.AsReadOnly(); }
        } 
        #endregion

        #region Public Methodes

        /// <summary>
        /// to do
        /// </summary>
        /// <param name="name">to do 2</param>
        public void AddSheet(string name)
        {
            this.sheets.Add(this.book.Worksheets.Add());

            this.sheets.Last().Name = name;
            this.AddNameToSheetsNames(name);
        }

        /// <summary>
        /// to do
        /// </summary>
        /// <param name="name">to do 2</param>
        public void RemoveSheet(string name)
        {
            if (this.sheets.Count > 0)
            {
                var sheetToRemove = this.FindSheet(name);
                if (sheetToRemove != null)
                {
                    int index = sheetToRemove.Index;
                    this.book.Sheets[index].Delete();
                    this.sheets.Remove(sheetToRemove);
                    this.sheetsNames.Remove(name);
                }
            }
        }

        /// <summary>
        /// to do
        /// </summary>
        /// <param name="name">to do 2</param>
        public void SheetToWrite(string name)
        {
            if (!this.sheetToWrite.Name.Equals(name))
            {
                this.sheetToWrite = this.FindSheet(name);
            }
        }

        /// <summary>
        /// to do
        /// </summary>
        /// <param name="row">to do 2</param>
        /// <param name="column">to do 3</param>
        /// <param name="text">to do 4</param>
        /// <param name="options">to do 5</param>
        public void WriteOnCell(int row, int column, string text, CellOptions options)
        {
            this.WriteOnCell(row, column, text);
            if (this.sheetToWrite != null)
            {
                var cell = this.sheetToWrite.Cells[row, column];
                this.AttributeOptions(options, cell);
            }
        }

        /// <summary>
        /// to do
        /// </summary>
        /// <param name="row">to do 2</param>
        /// <param name="column">to do 3</param>
        /// <param name="text">to do 4</param>
        public void WriteOnCell(int row, int column, string text)
        {
            if (this.sheetToWrite != null)
            {
                this.sheetToWrite.Cells[row, column] = text;
            }
        }

        /// <summary>
        /// to do
        /// </summary>
        /// <param name="row">to do 2</param>
        /// <param name="startColumn">to do 3</param>
        /// <param name="texts">to do 4</param>
        /// <param name="options">to do 5</param>
        public void WriteHorizontalHeaders(int row, int startColumn, string[] texts, CellOptions options)
        {
            int numCol = startColumn;
            for (int i = 0; i < texts.Length; i++)
            {
                this.WriteOnCell(row, numCol, texts[i], options);
                numCol++;
            }
        }

        /// <summary>
        /// to do
        /// </summary>
        /// <param name="row">to do 2</param>
        /// <param name="startColumn">to do 3</param>
        /// <param name="texts">to do 4</param>
        public void WriteHorizontalHeaders(int row, int startColumn, string[] texts)
        {
            int numCol = startColumn;
            for (int i = 0; i < texts.Length; i++)
            {
                this.WriteOnCell(row, numCol, texts[i]);
                numCol++;
            }
        }

        /// <summary>
        /// to do
        /// </summary>
        /// <param name="startRow">to do2</param>
        /// <param name="startColumn">to do 3</param>
        /// <param name="endRow">to do 4</param>
        /// <param name="endColumn">to do 5</param>
        /// <param name="options">to do 6</param>
        public void FormatRangeOptions(int startRow, int startColumn, int endRow, int endColumn, CellOptions options)
        {
            Excel.Range range = this.FindRange(startRow, startColumn, endRow, endColumn);
            AttributeOptions(options, range);
        }

        /// <summary>
        /// to do
        /// </summary>
        /// <param name="startRow">to do 2</param>
        /// <param name="startColumn">to do 3</param>
        /// <param name="endRow">to do 4</param>
        /// <param name="endColumn">to do 5</param>
        /// <param name="options">to do 6</param>
        public void FormatRangeAsTable(int startRow, int startColumn, int endRow, int endColumn, TableOptions options)
        {
            Excel.Range range = this.FindRange(startRow, startColumn, endRow, endColumn);
            range.Worksheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, range, System.Type.Missing, (Excel.XlYesNoGuess)options.Headers, System.Type.Missing).Name = options.Name;
            range.Select();

            string tabStyle = (options.Style != TableStyle.None) ? options.Style.ToString() : string.Empty;
            range.Worksheet.ListObjects[options.Name].TableStyle = tabStyle;
        }

        /// <summary>
        /// to do
        /// </summary>
        /// <param name="startRow">to do 2</param>
        /// <param name="startColumn">to do 3</param>
        /// <param name="endRow">to do 4</param>
        /// <param name="endColumn">to do 5</param>
        /// <param name="options">to do 6</param>
        public void CreateChartFromRange(int startRow, int startColumn, int endRow, int endColumn, ChartOptions options)
        {
            var charts = this.sheetToWrite.ChartObjects() as Excel.ChartObjects;
            var chartObject = charts.Add(options.Left, options.Top, 500, 500) as Excel.ChartObject;
            var chart = chartObject.Chart;

            var range = this.FindRange(startRow, startColumn, endRow, endColumn);
            chart.SetSourceData(range);

            if (options.Title != null)
            {
                chart.ChartWizard(Source: range, Title: options.Title);
            }

            if (options.XAxeTitle != null)
            {
                chart.ChartWizard(Source: range, CategoryTitle: options.XAxeTitle);
            }

            if (options.YAxeTitle != null)
            {
                chart.ChartWizard(Source: range, ValueTitle: options.YAxeTitle);
            }

            // Important le style en dernier sinon ca déonne avec le nom des axes.
            if (options.Style != ChartStyle.None)
            {
                chart.ChartType = (Excel.XlChartType)options.Style;
            }
        }

        /// <summary>
        /// to do
        /// </summary>
        public void AutoFitColumns()
        {
            this.sheetToWrite.Columns.AutoFit();
        }

        /// <summary>
        /// to do
        /// </summary>
        public void Save()
        {
            this.book.SaveAs(this.path);
        }

        /// <summary>
        /// to do
        /// </summary>
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// to do
        /// </summary>
        /// <param name="disposing">to do 2</param>
        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Libère les objets COM
                if (this.app != null)
                {
                    this.app.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(this.app);
                }

                if (this.book != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(this.book);
                }

                if (this.sheetToWrite != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(this.sheetToWrite);
                }

                this.app = null;
                this.book = null;
                this.sheetToWrite = null;
            }
        }
        #endregion

        #region Private Methodes

        /// <summary>
        /// to do
        /// </summary>
        /// <param name="options">to do 2</param>
        /// <param name="cell">to do 3</param>
        private static void AttributeOptions(CellOptions options, dynamic cell)
        {
            if (options.Bold)
            {
                cell.Font.Bold = options.Bold;
            }

            if (options.Italic)
            {
                cell.Font.Italic = options.Italic;
            }

            if (options.Underline)
            {
                cell.Font.Underline = options.Underline;
            }

            if (options.FontName != null)
            {
                cell.Font.Name = options.FontName;
            }

            if (options.FontSize > 0)
            {
                cell.Font.Size = options.FontSize;
            }

            if (options.FontColor != null)
            {
                cell.Font.Color = options.FontColor;
            }

            if (options.CellColor != Color.Empty)
            {
                cell.Interior.Color = options.CellColor;
            }

            if (options.TextHorizontalAlignment != HorizontalAlignment.None)
            {
                cell.HorizontalAlignment = (Excel.XlHAlign)options.TextHorizontalAlignment;
            }

            if (options.TextVerticalAlignment != VerticalAlignment.None)
            {
                cell.VerticalAlignment = (Excel.XlVAlign)options.TextVerticalAlignment;
            }

            if (options.Borders.Count > 0)
            {
                var borders = cell.Borders;

                foreach (var border in options.Borders)
                {
                    if (border.Style != BorderStyle.None)
                    {
                        var index = (Excel.XlBordersIndex)border.Position;

                        borders[index].LineStyle = (Excel.XlLineStyle)border.Style;

                        if (border.Thickness > 0)
                        {
                            var thickness = (border.Thickness > 4) ? 4 : border.Thickness;
                            borders[index].Weight = thickness;
                        }

                        if (border.Color != null)
                        {
                            borders[index].Color = border.Color;
                        }
                    }
                }
            }
        } 

        /// <summary>
        /// to do
        /// </summary>
        /// <param name="startRow">to do 2</param>
        /// <param name="startColumn">to do 3</param>
        /// <param name="endRow">to do 4</param>
        /// <param name="endColumn">to do 5</param>
        /// <returns>to do 6</returns>
        private Excel.Range FindRange(int startRow, int startColumn, int endRow, int endColumn)
        {
            string startCell = this.convert.ConvertCellNumToLetter(startRow, startColumn);
            string endCell = this.convert.ConvertCellNumToLetter(endRow, endColumn);
            return this.sheetToWrite.get_Range(startCell, endCell);
        }

        /// <summary>
        /// to do
        /// </summary>
        /// <param name="name">to do 2</param>
        /// <returns>to do 3</returns>
        private Excel.Worksheet FindSheet(string name)
        {
            return this.sheets.Where(s => s.Name.Equals(name)).FirstOrDefault();
        }

        /// <summary>
        /// to do
        /// </summary>
        /// <param name="name">to do 2</param>
        private void AddNameToSheetsNames(string name)
        {
            this.sheetsNames.Add(name);
        }
        #endregion
    }
}
