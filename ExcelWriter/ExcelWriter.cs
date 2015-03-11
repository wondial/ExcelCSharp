using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.ObjectModel;
using System.Drawing;

namespace ExcelWriterCSharp
{
    public class ExcelWriter : IDisposable
    {
        private Excel.Application _app;
        private Excel.Workbook _book;
        private Excel.Worksheet _sheetToWrite;
        private List<Excel.Worksheet> _sheets = new List<Excel.Worksheet>();
        private string _path;
        private List<string> _sheetsNames;
        private ExcelConvert _convert;

        public ReadOnlyCollection<string> SheetsNames
        {
            get { return _sheetsNames.AsReadOnly(); }
        }

        public ExcelWriter(string path, string sheetName)
        {
            _path = path;

            _sheetsNames = new List<string>();

            _app = new Excel.Application();
            _app.DisplayAlerts = false;
            _book = _app.Workbooks.Add();

            _app.ActiveSheet.Name = sheetName;
            AddNameToSheetsNames(sheetName);
            _sheets.Add(_app.ActiveSheet);
            _sheetToWrite = _app.ActiveSheet;

            _convert = new ExcelConvert();
        }

        public void AddSheet(string name)
        {
            _sheets.Add(_book.Worksheets.Add());
            
            _sheets.Last().Name = name;
            AddNameToSheetsNames(name);
        }

        public void RemoveSheet(string name)
        {
            if (_sheets.Count > 0)
            {
                var sheetToRemove = FindSheet(name);
                if (sheetToRemove != null)
                {
                    int index = sheetToRemove.Index;
                    _book.Sheets[index].Delete();
                    _sheets.Remove(sheetToRemove);
                    _sheetsNames.Remove(name);
                }
            }
        }

        public void SheetToWrite(string name)
        {
            if (!_sheetToWrite.Name.Equals(name))
                _sheetToWrite = FindSheet(name);
        }

        public void WriteOnCell(int row, int column, string text, CellOptions options)
        {
            WriteOnCell(row, column, text);
            if (_sheetToWrite != null)
            {
                var cell = _sheetToWrite.Cells[row,column];

                AttributeOptions(options, cell);
            }
        }

        public void WriteOnCell(int row, int column, string text)
        {
            if (_sheetToWrite != null)
                _sheetToWrite.Cells[row, column] = text;
        }

        public void WriteHorizontalHeaders(int row, int startColumn, string[] texts, CellOptions options)
        {
            int numCol = startColumn;
            for (int i = 0; i < texts.Length; i++)
            {
                WriteOnCell(row, numCol, texts[i], options);
                numCol++;
            }
        }

        public void WriteHorizontalHeaders(int row, int startColumn, string[] texts)
        {
            int numCol = startColumn;
            for (int i = 0; i < texts.Length; i++)
            {
                WriteOnCell(row, numCol, texts[i]);
                numCol++;
            }
        }

        public void FormatRangeOptions(int startRow, int startColumn,int endRow, int endColumn, CellOptions options)
        {
            Excel.Range range = FindRange(startRow, startColumn, endRow, endColumn);
            AttributeOptions(options, range);
        }

        public void FormatRangeAsTable(int startRow, int startColumn,int endRow, int endColumn, string tableName)
        {
            Excel.Range range = FindRange(startRow, startColumn, endRow, endColumn);
            range.Worksheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, range, System.Type.Missing, Excel.XlYesNoGuess.xlYes, System.Type.Missing).Name = tableName;
            range.Select();
            range.Worksheet.ListObjects[tableName].TableStyle = "TableStyleMedium14";
        }

        private Excel.Range FindRange(int startRow, int startColumn,int endRow, int endColumn)
        {
            string startCell = _convert.ConvertCellNumToLetter(startRow, startColumn);
            string endCell = _convert.ConvertCellNumToLetter(endRow, endColumn);
            return _sheetToWrite.get_Range(startCell, endCell);
        }

        public void AutoFitColumns()
        {
            _sheetToWrite.Columns.AutoFit();
        }

        private static void AttributeOptions(CellOptions options, dynamic cell)
        {
            if (options.Bold)
                cell.Font.Bold = options.Bold;
            if (options.Italic)
                cell.Font.Italic = options.Italic;
            if (options.Underline)
                cell.Font.Underline = options.Underline;

            if (options.FontName != null)
                cell.Font.Name = options.FontName;
            if (options.FontSize > 0)
                cell.Font.Size = options.FontSize;
            if (options.FontColor != null)
                cell.Font.Color = options.FontColor;

            if (options.CellColor != Color.Empty )
                cell.Interior.Color = options.CellColor;

            if(options.TextHorizontalAlignment != 0)
                cell.HorizontalAlignment = (Excel.XlHAlign)options.TextHorizontalAlignment;
            if (options.TextVerticalAlignment != 0)
                cell.VerticalAlignment = (Excel.XlVAlign)options.TextVerticalAlignment;

            if (options.Borders.Count > 0)
            {
                var borders = cell.Borders;
                foreach (var border in options.Borders)
                {
                    var index = (Excel.XlBordersIndex)border.Position;

                    if(border.Style != 0)
                        borders[index].LineStyle = (Excel.XlLineStyle)border.Style;
                    if (border.Thickness > 0)
                    {
                        var thickness = (border.Thickness > 4) ? 4 : border.Thickness;
                        borders[index].Weight = thickness;
                    }
                    if(border.Color != null)
                        borders[index].Color = border.Color;
                }
            }
        }

        public void Save()
        {
            _book.SaveAs(_path);
        }

        private Excel.Worksheet FindSheet(string name)
        {
            return _sheets.Where(s => s.Name.Equals(name)).FirstOrDefault();
        }

        private void AddNameToSheetsNames(string name)
        {
            _sheetsNames.Add(name);
        }

        public void Dispose()
        {
            // Libère les objets COM
            if (_app != null)
            {
                _app.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_app);
            }
            if(_book != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_book);
            if(_sheetToWrite != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_sheetToWrite);

            _app = null;
            _book = null;
            _sheetToWrite = null;

            GC.Collect();
        }
    }
}
