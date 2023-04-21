using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;

namespace Controlling
{
    /*
    internal class PivotCreator
    {

        static void Main(string[] args)
        {
            string inputFilePath = @"C:\path\to\your\input\excel\file.xlsx";
            string outputFilePath = @"C:\path\to\your\output\excel\file.xlsx";

            CreatePivotTable(inputFilePath, outputFilePath);
        }

        private static void CreatePivotTable(string inputFilePath, string outputFilePath)
        {
            File.Copy(inputFilePath, outputFilePath, true);

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(outputFilePath, true))
            {
                WorkbookPart workbookPart = document.WorkbookPart;

                // Create Table
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                TableDefinitionPart tableDefinitionPart = worksheetPart.AddNewPart<TableDefinitionPart>();
                Table table = CreateTable(worksheetPart);
                tableDefinitionPart.Table = table;
                tableDefinitionPart.Table.Save();

                // Create PivotTable
                WorksheetPart pivotTableWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                string pivotTableSheetId = workbookPart.GetIdOfPart(pivotTableWorksheetPart);
                pivotTableWorksheetPart.Worksheet = new Worksheet();
                pivotTableWorksheetPart.Worksheet.Save();
                Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
                sheets.Append(new Sheet()
                {
                    Id = pivotTableSheetId,
                    SheetId = (uint)sheets.Count() + 1,
                    Name = "PivotTable"
                });

                PivotTablePart pivotTablePart = pivotTableWorksheetPart.AddNewPart<PivotTablePart>();
                PivotCacheDefinitionPart pivotCacheDefinitionPart = workbookPart.AddNewPart<PivotCacheDefinitionPart>();

                pivotTablePart.PivotTableDefinition = CreatePivotTableDefinition();
                pivotTablePart.PivotTableDefinition.Save();

                pivotCacheDefinitionPart.PivotCacheDefinition = CreatePivotCacheDefinition(workbookPart, table);
                pivotCacheDefinitionPart.PivotCacheDefinition.Save();

                pivotTablePart.AddPart(pivotCacheDefinitionPart);

                workbookPart.Workbook.Save();
            }
        }

        private static Table CreateTable(WorksheetPart worksheetPart)
        {
            string sheetName = worksheetPart.Worksheet.Descendants<Sheet>().First().Name;
            string reference = worksheetPart.Worksheet.Descendants<SheetData>().First().OuterXml;
            int firstRow = 1;
            int firstColumn = 1;
            int lastRow = reference.Split(new string[] { "<row" }, StringSplitOptions.None).Length - 2;
            int lastColumn = reference.Split(new string[] { "<c" }, StringSplitOptions.None)[1].Split('r')[1].First() - 65 + 1;

            string range = $"{GetColumnName(firstColumn)}{firstRow}:{GetColumnName(lastColumn)}{lastRow}";

            Table table = new Table()
            {
                Id = 1,
                Name = "DataTable",
                DisplayName = "DataTable",
                Reference = range,
                TotalsRowShown = false
            };

            // Table Columns
            TableColumn[] tableColumns = new TableColumn[lastColumn];
            for (int i = 0; i < lastColumn; i++)
            {
                tableColumns[i] = new TableColumn()
                {
                    Id = (uint)(i + 1),
                    Name = $"Column{i + 1}"
                };
            }

            table.Append(new TableColumns() { Count = (uint)tableColumns.Length }, new TableStyleInfo() { Name = "TableStyleMedium2", ShowColumnStripes = false, ShowRowStripes = true });
            table.TableColumns.Append(tableColumns);

            // Add table to the worksheet
            TablePart tablePart = worksheetPart.AddNewPart<TablePart>();
            tablePart.Table = table;
            tablePart.Table.Save();
            worksheetPart.Worksheet.Append(new TablePart() { Id = worksheetPart.GetIdOfPart(tablePart) });
            worksheetPart.Worksheet.Save();

            return table;
        }

        private static PivotTableDefinition CreatePivotTableDefinition()
        {
            PivotTableDefinition pivotTableDefinition = new PivotTableDefinition()
            {
                Name = "PivotTable1",
                CacheId = 0,
                DataOnRows = true,
                AutoFormatId = 0,
                ApplyNumberFormats = true,
                ApplyBorderFormats = true,
                ApplyFontFormats = true,
                ApplyPatternFormats = true,
                ApplyAlignmentFormats = true,
                ApplyWidthHeightFormats = true,
                DataCaption = "Data"
            };

            // Layout for the PivotTable
            pivotTableDefinition.Append(new Location() { Reference = "A3" });
            pivotTableDefinition.Append(new PivotTableStyleInfo() { Name = "PivotStyleMedium9", ShowRowStripes = true });

            return pivotTableDefinition;
        }

        private static PivotCacheDefinition CreatePivotCacheDefinition(WorkbookPart workbookPart, Table table)
        {
            PivotCacheDefinition pivotCacheDefinition = new PivotCacheDefinition()
            {
                SaveData = true,
                RefreshOnLoad = true,
                BackgroundQuery = false,
                SupportSubquery = false,
                SupportAdvancedDrill = true
            };

            CacheSource cacheSource = new CacheSource() { Type = CacheSourceType.Worksheet };
            cacheSource.Append(new WorksheetSource() { Sheet = "Sheet1", Table = table.Name });

            pivotCacheDefinition.Append(cacheSource);
            pivotCacheDefinition.Append(CreatePivotCacheFields(workbookPart, table));

            return pivotCacheDefinition;
        }

        private static PivotCacheFields CreatePivotCacheFields(WorkbookPart workbookPart, Table table)
        {
            PivotCacheFields pivotCacheFields = new PivotCacheFields() { Count = (uint)table.TableColumns.Count() };

            foreach (TableColumn tableColumn in table.TableColumns)
            {
                PivotCacheField pivotCacheField = new PivotCacheField() { Name = tableColumn.Name, ShowNewItems = true };
                pivotCacheFields.Append(pivotCacheField);
            }

            return pivotCacheFields;
        }

        private static string GetColumnName(int columnIndex)
        {
            int dividend = columnIndex;
            string columnName = String.Empty;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
    }
    */
}


