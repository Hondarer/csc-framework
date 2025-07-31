using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelApp
{
    public class ExcelHandler
    {
        /// <summary>
        /// Excelファイルを読み込む
        /// </summary>
        /// <param name="filePath">ファイルパス</param>
        /// <param name="worksheetName">ワークシート名（nullの場合は最初のシート）</param>
        /// <returns>読み込んだデータ</returns>
        public static List<string[]> ReadExcel(string filePath, string worksheetName = null)
        {
            var result = new List<string[]>();
            
            try
            {
                using (var document = SpreadsheetDocument.Open(filePath, false))
                {
                    var workbookPart = document.WorkbookPart;
                    var worksheetPart = GetWorksheetPart(workbookPart, worksheetName);
                    
                    if (worksheetPart == null)
                    {
                        Console.WriteLine($"ワークシート '{worksheetName}' が見つかりません。");
                        return result;
                    }
                    
                    var worksheet = worksheetPart.Worksheet;
                    var sharedStringTablePart = workbookPart.SharedStringTablePart;
                    var sharedStringTable = sharedStringTablePart?.SharedStringTable;
                    
                    var rows = worksheet.GetFirstChild<SheetData>().Elements<Row>();
                    
                    foreach (var row in rows.OrderBy(r => r.RowIndex))
                    {
                        var rowData = new List<string>();
                        var cells = row.Elements<Cell>().OrderBy(c => c.CellReference.Value);
                        
                        string lastColumnName = "";
                        foreach (var cell in cells)
                        {
                            var columnName = GetColumnName(cell.CellReference);
                            
                            // 空の列を埋める
                            while (GetNextColumnName(lastColumnName) != columnName && !string.IsNullOrEmpty(lastColumnName))
                            {
                                rowData.Add("");
                                lastColumnName = GetNextColumnName(lastColumnName);
                            }
                            
                            var cellValue = GetCellValue(cell, sharedStringTable);
                            rowData.Add(cellValue);
                            lastColumnName = columnName;
                        }
                        
                        result.Add(rowData.ToArray());
                    }
                }
                
                Console.WriteLine($"Excelファイルを読み込みました: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Excel読み込みエラー: {ex.Message}");
            }
            
            return result;
        }
        
        /// <summary>
        /// Excelファイルに書き込む
        /// </summary>
        /// <param name="filePath">ファイルパス</param>
        /// <param name="data">書き込むデータ</param>
        /// <param name="worksheetName">ワークシート名</param>
        public static void WriteExcel(string filePath, List<string[]> data, string worksheetName = "Sheet1")
        {
            try
            {
                // ファイルが存在する場合は削除
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
                
                using (var document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
                {
                    // Workbookパートを作成
                    var workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();
                    
                    // Worksheetパートを作成
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());
                    
                    // Sheetを追加
                    var sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());
                    var sheet = new Sheet()
                    {
                        Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = worksheetName
                    };
                    sheets.Append(sheet);
                    
                    // データを書き込み
                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                    
                    for (uint rowIndex = 0; rowIndex < data.Count; rowIndex++)
                    {
                        var row = new Row() { RowIndex = rowIndex + 1 };
                        sheetData.Append(row);
                        
                        for (int colIndex = 0; colIndex < data[(int)rowIndex].Length; colIndex++)
                        {
                            var cellReference = GetCellReference(rowIndex + 1, colIndex);
                            var cell = new Cell()
                            {
                                CellReference = cellReference,
                                DataType = CellValues.InlineString,
                                InlineString = new InlineString() { Text = new Text(data[(int)rowIndex][colIndex]) }
                            };
                            row.Append(cell);
                        }
                    }
                    
                    workbookPart.Workbook.Save();
                }
                
                Console.WriteLine($"Excelファイルを保存しました: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Excel書き込みエラー: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 複数シートでの書き込み
        /// </summary>
        /// <param name="filePath">ファイルパス</param>
        /// <param name="sheetsData">シート名とデータの辞書</param>
        public static void WriteMultipleSheets(string filePath, Dictionary<string, List<string[]>> sheetsData)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
                
                using (var document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
                {
                    var workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();
                    var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                    
                    uint sheetId = 1;
                    foreach (var sheetData in sheetsData)
                    {
                        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                        worksheetPart.Worksheet = new Worksheet(new SheetData());
                        
                        var sheet = new Sheet()
                        {
                            Id = workbookPart.GetIdOfPart(worksheetPart),
                            SheetId = sheetId++,
                            Name = sheetData.Key
                        };
                        sheets.Append(sheet);
                        
                        // データ書き込み
                        var sheetDataElement = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                        var data = sheetData.Value;
                        
                        for (uint rowIndex = 0; rowIndex < data.Count; rowIndex++)
                        {
                            var row = new Row() { RowIndex = rowIndex + 1 };
                            sheetDataElement.Append(row);
                            
                            for (int colIndex = 0; colIndex < data[(int)rowIndex].Length; colIndex++)
                            {
                                var cellReference = GetCellReference(rowIndex + 1, colIndex);
                                var cell = new Cell()
                                {
                                    CellReference = cellReference,
                                    DataType = CellValues.InlineString,
                                    InlineString = new InlineString() { Text = new Text(data[(int)rowIndex][colIndex]) }
                                };
                                row.Append(cell);
                            }
                        }
                    }
                    
                    workbookPart.Workbook.Save();
                }
                
                Console.WriteLine($"複数シートのExcelファイルを保存しました: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Excel書き込みエラー: {ex.Message}");
            }
        }
        
        #region ヘルパーメソッド
        
        private static WorksheetPart GetWorksheetPart(WorkbookPart workbookPart, string worksheetName)
        {
            if (string.IsNullOrEmpty(worksheetName))
            {
                return workbookPart.WorksheetParts.First();
            }
            
            var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>()
                .FirstOrDefault(s => s.Name == worksheetName);
            
            if (sheet == null) return null;
            
            return (WorksheetPart)workbookPart.GetPartById(sheet.Id);
        }
        
        private static string GetCellValue(Cell cell, SharedStringTable sharedStringTable)
        {
            if (cell.CellValue == null) return "";
            
            var value = cell.CellValue.InnerXml;
            
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                if (sharedStringTable != null)
                {
                    return sharedStringTable.ChildElements[int.Parse(value)].InnerText;
                }
            }
            
            return value;
        }
        
        private static string GetColumnName(string cellReference)
        {
            var columnName = "";
            foreach (char c in cellReference)
            {
                if (char.IsLetter(c))
                {
                    columnName += c;
                }
                else
                {
                    break;
                }
            }
            return columnName;
        }
        
        private static string GetNextColumnName(string columnName)
        {
            if (string.IsNullOrEmpty(columnName))
            {
                return "A";
            }
            
            var chars = columnName.ToCharArray();
            for (int i = chars.Length - 1; i >= 0; i--)
            {
                if (chars[i] == 'Z')
                {
                    chars[i] = 'A';
                }
                else
                {
                    chars[i]++;
                    return new string(chars);
                }
            }
            
            return "A" + new string(chars);
        }
        
        private static string GetCellReference(uint row, int column)
        {
            string columnName = "";
            int columnNumber = column;
            
            while (columnNumber >= 0)
            {
                columnName = (char)('A' + (columnNumber % 26)) + columnName;
                columnNumber = columnNumber / 26 - 1;
            }
            
            return columnName + row;
        }
        
        #endregion
    }
}
