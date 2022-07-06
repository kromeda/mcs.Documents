using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Documents.Reports
{
    public class SentIndividualNotifications
    {
        public byte[] Export(SentIndividualArgs args)
        {
            using var ms = new MemoryStream();
            using var document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            AddStyles(workbookPart);

            DetailsSheet(args, workbookPart);
            AggregateSheet(args, workbookPart);

            workbookPart.Workbook.Save();
            document.Close();

            static void AddStyles(WorkbookPart workbookPart)
            {
                var wbsp = workbookPart.AddNewPart<WorkbookStylesPart>();
                wbsp.Stylesheet = GenerateStyles();
                wbsp.Stylesheet.Save();
            }

            return ms.ToArray();
        }

        private static void DetailsSheet(
            SentIndividualArgs args,
            WorkbookPart workbookPart)
        {
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            var sheetId = workbookPart.GetIdOfPart(worksheetPart);
            var sheet = new Sheet { Id = sheetId, SheetId = 1, Name = "Поименный отчет" };
            sheets.Append(sheet);
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            SetupColumns(worksheetPart);
            InsertHeader(args.Range, sheetData);
            InsertEmptyRow(sheetData);
            InsertTableHeaders(sheetData);
            InsertTableData(sheetData, args.Data);
            MergeCells(worksheetPart);

            static void SetupColumns(WorksheetPart worksheetPart)
            {
                var lstColumns = worksheetPart.Worksheet.GetFirstChild<Columns>();
                var needToInsertColumns = false;

                if (lstColumns == null)
                {
                    lstColumns = new Columns();
                    needToInsertColumns = true;
                }

                lstColumns.Append(new Column() { Min = 1, Max = 11, Width = 16, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 3, Max = 3, Width = 30, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 11, Max = 11, Width = 30, CustomWidth = true });

                if (needToInsertColumns)
                {
                    worksheetPart.Worksheet.InsertAt(lstColumns, 0);
                }
            }

            static void InsertHeader(StartEndRange range, SheetData sheetData)
            {
                var row = new Row { RowIndex = 2 };
                sheetData.Append(row);

                var header = $"Отчет по отправленным уведомлениям за период с {range.Start.ToShortDateString()} по {range.End.ToShortDateString()}";
                InsertCell(row, "A", header, CellValues.String, 1);
            }

            static void InsertEmptyRow(SheetData sheetData)
            {
                var row = new Row { RowIndex = 3 };
                sheetData.Append(row);
            }

            static void InsertTableHeaders(SheetData sheetData)
            {
                var row = new Row { RowIndex = 4 };
                sheetData.Append(row);

                InsertCell(row, "A", "Дата дела (пост. на контроль)", CellValues.String, 2);
                InsertCell(row, "B", "Номер дела", CellValues.String, 2);
                InsertCell(row, "C", "Адрес потребителя", CellValues.String, 2);
                InsertCell(row, "D", "Лицевой счет", CellValues.String, 2);
                InsertCell(row, "E", "ФИО", CellValues.String, 2);
                InsertCell(row, "F", "Сумма по уведомлению", CellValues.String, 2);
                InsertCell(row, "G", "Номер уведомления", CellValues.String, 2);
                InsertCell(row, "H", "Дата уведомления", CellValues.String, 2);
                InsertCell(row, "I", "Наименование ПУ", CellValues.String, 2);
                InsertCell(row, "J", "Наименование филиала", CellValues.String, 2);
                InsertCell(row, "K", "Адрес доставки уведомления", CellValues.String, 2);
            }

            static void InsertTableData(SheetData sheetData, List<SentIndividualNotificationsData> reportData)
            {
                var indentRows = 5;
                var detailsData = reportData.Where(rowItem => rowItem.RowLevel == 0).ToList();

                foreach (var dataRow in detailsData)
                {
                    var index = detailsData.IndexOf(dataRow);
                    var currentRowIndex = index + indentRows;
                    var row = new Row { RowIndex = new UInt32Value((uint)currentRowIndex) };
                    sheetData.Append(row);

                    InsertCell(row, "A", ReplaceHexadecimalSymbols(dataRow.CaseDate.Value.ToShortDateString()), CellValues.String, 3);
                    InsertCell(row, "B", ReplaceHexadecimalSymbols(dataRow.CaseNumber.ToString()), CellValues.Number, 4);
                    InsertCell(row, "C", ReplaceHexadecimalSymbols(dataRow.PointAddress), CellValues.String, 3);
                    InsertCell(row, "D", ReplaceHexadecimalSymbols(dataRow.ContractNumber.ToString()), CellValues.Number, 4);
                    InsertCell(row, "E", ReplaceHexadecimalSymbols(dataRow.CustomerName), CellValues.String, 3);
                    InsertCell(row, "F", dataRow.DebtSum.Value, CellValues.Number, 5);
                    InsertCell(row, "G", dataRow.NotificationNumber.Value, CellValues.Number, 4);
                    InsertCell(row, "E", dataRow.NotificationDate.Value, CellValues.Date, 6);
                    InsertCell(row, "I", ReplaceHexadecimalSymbols(dataRow.OfficeName), CellValues.String, 3);
                    InsertCell(row, "J", ReplaceHexadecimalSymbols(dataRow.BranchName), CellValues.String, 3);
                    InsertCell(row, "K", ReplaceHexadecimalSymbols(dataRow.DeliveryAddress), CellValues.String, 3);
                }
            }

            static void MergeCells(WorksheetPart worksheetPart)
            {
                var mergeCells = new MergeCells();
                mergeCells.Append(new MergeCell() { Reference = new StringValue("A2:K2") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("A3:K3") });
                worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetData>().First());
            }
        }

        private static void AggregateSheet(SentIndividualArgs args, WorkbookPart workbookPart)
        {
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            var sheets = workbookPart.Workbook.Sheets;
            var sheetId = workbookPart.GetIdOfPart(worksheetPart);
            var sheet = new Sheet { Id = sheetId, SheetId = 2, Name = "Сводный отчет" };
            sheets.Append(sheet);
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            SetupColumns(worksheetPart);
            InsertHeader(args.Range, sheetData);
            InsertTableHeaders(sheetData);
            InsertTableData(sheetData, args.Data);
            MergeCells(worksheetPart);

            static void SetupColumns(WorksheetPart worksheetPart)
            {
                var lstColumns = worksheetPart.Worksheet.GetFirstChild<Columns>();
                var needToInsertColumns = false;

                if (lstColumns == null)
                {
                    lstColumns = new Columns();
                    needToInsertColumns = true;
                }

                lstColumns.Append(new Column() { Min = 1, Max = 3, Width = 35, CustomWidth = true });

                if (needToInsertColumns)
                {
                    worksheetPart.Worksheet.InsertAt(lstColumns, 0);
                }
            }

            static void InsertHeader(StartEndRange range, SheetData sheetData)
            {
                var row = new Row { RowIndex = 2 };
                sheetData.Append(row);

                var header = $"Сводный отчет по отправленным уведомлениям за период с {range.Start.ToShortDateString()} по {range.End.ToShortDateString()}";
                InsertCell(row, "A", header, CellValues.String, 7);
            }

            static void InsertTableHeaders(SheetData sheetData)
            {
                var row = new Row { RowIndex = 4 };
                sheetData.Append(row);

                InsertCell(row, "A", "Наименование филиала", CellValues.String, 2);
                InsertCell(row, "B", "Наименование ПУ", CellValues.String, 2);
                InsertCell(row, "C", "Количество направленных уведомлений", CellValues.String, 2);
            }

            static void InsertTableData(SheetData sheetData, List<SentIndividualNotificationsData> reportData)
            {
                var indentRows = 5;

                var aggregatedData = reportData.
                    Where(rowItem => rowItem.RowLevel > 0).
                    OrderBy(rowItem => rowItem.RowOrder).ToList();

                foreach (var dataRow in aggregatedData)
                {
                    var index = aggregatedData.IndexOf(dataRow);
                    var currentRowIndex = index + indentRows;
                    var row = new Row { RowIndex = (uint)currentRowIndex };
                    sheetData.Append(row);

                    var boldRowsLevels = new[] { 1, 3 };
                    var textStyleId = (uint)(boldRowsLevels.Contains(dataRow.RowLevel) ? 2 : 3);
                    var numberStyleId = (uint)(boldRowsLevels.Contains(dataRow.RowLevel) ? 2 : 4);

                    InsertCell(row, "A", ReplaceHexadecimalSymbols(dataRow.OfficeName), CellValues.String, textStyleId);
                    InsertCell(row, "B", ReplaceHexadecimalSymbols(dataRow.BranchName), CellValues.String, textStyleId);
                    InsertCell(row, "C", dataRow.Amount, CellValues.Number, numberStyleId);
                }
            }

            static void MergeCells(WorksheetPart worksheetPart)
            {
                var mergeCells = new MergeCells();
                mergeCells.Append(new MergeCell() { Reference = new StringValue("A2:C2") });
                worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetData>().First());
            }
        }

        private static void InsertCell(Row row, string cell_letter, string val, CellValues type, uint styleIndex)
        {
            Cell refCell = null;
            var newCell = new Cell { CellReference = cell_letter + row.RowIndex.ToString(), StyleIndex = styleIndex };
            row.InsertBefore(newCell, refCell);

            newCell.CellValue = new CellValue(val);
            newCell.DataType = new EnumValue<CellValues>(type);
        }

        private static void InsertCell(Row row, string cell_letter, decimal val, CellValues type, uint styleIndex)
        {
            Cell refCell = null;
            var newCell = new Cell { CellReference = cell_letter + ":" + row.RowIndex.ToString(), StyleIndex = styleIndex };
            row.InsertBefore(newCell, refCell);

            newCell.CellValue = new CellValue(val);
            newCell.DataType = new EnumValue<CellValues>(type);
        }

        private static void InsertCell(Row row, string cell_letter, DateTime val, CellValues type, uint styleIndex)
        {
            Cell refCell = null;
            var  newCell = new Cell { CellReference = cell_letter + ":" + row.RowIndex.ToString(), StyleIndex = styleIndex };
            row.InsertBefore(newCell, refCell);

            newCell.CellValue = new CellValue(val);
            newCell.DataType = new EnumValue<CellValues>(type);
        }

        private static string ReplaceHexadecimalSymbols(string txt)
        {
            string hexSymbolsPattern = "[\x00-\x08\x0B\x0C\x0E-\x1F\x26]";
            return Regex.Replace(txt, hexSymbolsPattern, "", RegexOptions.Compiled);
        }

        private static Stylesheet GenerateStyles()
        {
            var centerAlignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            return new Stylesheet(
                new Fonts(
                    new Font(
                        new FontSize() { Val = 11 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" }),
                    new Font(
                        new Bold(),
                        new FontSize() { Val = 16 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Arial" }),
                    new Font(
                        new Bold(),
                        new FontSize() { Val = 10 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Arial" }),
                    new Font(
                        new FontSize() { Val = 10 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Arial" }),
                    new Font(
                        new Bold(),
                        new FontSize() { Val = 11 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" })),
                new Fills(
                    new Fill(
                        new PatternFill() { PatternType = PatternValues.None })),
                new Borders(
                    new Border(
                        new LeftBorder(),
                        new RightBorder(),
                        new TopBorder(),
                        new BottomBorder(),
                        new DiagonalBorder()),
                    new Border(
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color() { Indexed = (UInt32Value)64U }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color() { Indexed = (UInt32Value)64U }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())),
                new CellFormats(
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 0 },
                    new CellFormat(centerAlignment) { FontId = 1, FillId = 0, BorderId = 0 },
                    new CellFormat(new Alignment()
                    { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Top, WrapText = true })
                    { FontId = 2, FillId = 0, BorderId = 1 },
                    new CellFormat(new Alignment()
                    { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top, WrapText = true })
                    { FontId = 3, FillId = 0, BorderId = 1 },
                    new CellFormat(new Alignment()
                    { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Top, WrapText = true })
                    { FontId = 3, FillId = 0, BorderId = 1, NumberFormatId = 1 },
                    new CellFormat(new Alignment()
                    { Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Top })
                    { FontId = 3, FillId = 0, BorderId = 1, NumberFormatId = 4 },
                    new CellFormat(new Alignment()
                    { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Top })
                    { FontId = 3, FillId = 0, BorderId = 1, NumberFormatId = 14, ApplyNumberFormat = true },
                    new CellFormat(new Alignment()
                    { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom })
                    { FontId = 4, FillId = 0, BorderId = 0 }));
        }
    }
}