using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using CodeFluent.Runtime.Utilities;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExportTable
{
    public class SpreadsheetGenerator<T>
    {
        // ReSharper disable PossiblyMistakenUseOfParamsMethod

        private readonly Table<T> _table;

        public Table<T> Table
        {
            get { return _table; }
        }

        public SpreadsheetGenerator(Table<T> table)
        {
            if (table == null) throw new ArgumentNullException("table");

            _table = table;
        }

        public void Generate(string path)
        {
            if (path == null) throw new ArgumentNullException("path");

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
            {
                Generate(document);
            }
        }

        public void Generate(Stream stream)
        {
            if (stream == null) throw new ArgumentNullException("stream");

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                Generate(document);
            }
        }

        public void Generate(SpreadsheetDocument document)
        {
            if (document == null) throw new ArgumentNullException("document");

            if (Table.Columns.Count == 0)
                return;

            WorkbookPart part = document.AddWorkbookPart();
            part.Workbook = new Workbook();



            // Add style
            WorkbookStylesPart stylePart = part.AddNewPart<WorkbookStylesPart>();
            CreateWorkbookStylesPart(stylePart);

            // Add Sheets to the Workbook.
            WorksheetPart worksheetPart = part.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            worksheetPart.Worksheet = new Worksheet();

            Sheets sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());    
            Sheet sheet = new Sheet
            {
                Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = Table.Title
            };
            sheets.Append(sheet);

            //TableParts tableParts;
            //if (true)
            //{
            //    TableDefinitionPart tableDefinitionPart1 = worksheetPart.AddNewPart<TableDefinitionPart>("rId2");
            //    Table table1 = new Table() { Id = (UInt32Value)1U, Name = "Table1", DisplayName = "Table1", Reference = "A1:B2", TotalsRowShown = false };
            //    AutoFilter autoFilter1 = new AutoFilter() { Reference = "A1:B2" };

            //    TableColumns tableColumns1 = new TableColumns() { Count = (UInt32Value)2U };
            //    TableColumn tableColumn1 = new TableColumn() { Id = (UInt32Value)1U, Name = "Title 1" };
            //    //TableColumn tableColumn2 = new TableColumn() { Id = (UInt32Value)2U, Name = "Title2", DataFormatId = (UInt32Value)0U };
            //    TableColumn tableColumn2 = new TableColumn() { Id = (UInt32Value)2U, Name = "Title2" };

            //    tableColumns1.Append(tableColumn1);
            //    tableColumns1.Append(tableColumn2);
            //    TableStyleInfo tableStyleInfo1 = new TableStyleInfo() { Name = "TableStyleLight9", ShowFirstColumn = false, ShowLastColumn = false, ShowRowStripes = true, ShowColumnStripes = false };

            //    table1.Append(autoFilter1);
            //    table1.Append(tableColumns1);
            //    table1.Append(tableStyleInfo1);

            //    tableDefinitionPart1.Table = table1;

            //    tableParts = new TableParts() { Count = (UInt32Value)1U };
            //    TablePart tablePart = new TablePart() { Id = "rId2" };
            //    tableParts.Append(tablePart);
            //}

            // Add rows
            uint rowIndex = 0;
            if (Table.ShowHeader)
            {
                AppendHeaders(sheetData, rowIndex);
                rowIndex++;
            }

            if (Table.DataSource != null)
            {
                foreach (var item in Table.DataSource)
                {
                    AppendRow(sheetData, rowIndex, item);
                    rowIndex++;
                }
            }

            worksheetPart.Worksheet.AppendChild(sheetData);
            //worksheetPart.Worksheet.AppendChild(tableParts);

            part.Workbook.Save();
        }

        private readonly Dictionary<TableColumn<T>, uint> _styleIndex = new Dictionary<TableColumn<T>, uint>();
        private uint GetStyleIndex(TableColumn<T> column)
        {
            return _styleIndex[column];
        }

        private void CreateWorkbookStylesPart(WorkbookStylesPart part)
        {
            Stylesheet stylesheet = new Stylesheet { MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "x14ac" } };
            stylesheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            Fonts fonts = new Fonts { Count = 2, KnownFonts = true };
            Font font = new Font();
            font.Append(new FontSize { Val = 11 });
            font.Append(new Color { Theme = 1 });
            font.Append(new FontName { Val = "Calibri" });
            font.Append(new FontFamilyNumbering { Val = 2 });
            font.Append(new FontScheme { Val = FontSchemeValues.Minor });
            fonts.Append(font);

            Font fontHeader = new Font();
            fontHeader.Append(new Bold());
            fontHeader.Append(new FontSize { Val = 11 });
            fontHeader.Append(new Color { Theme = 1 });
            fontHeader.Append(new FontName { Val = "Calibri" });
            fontHeader.Append(new FontFamilyNumbering { Val = 2 });
            fontHeader.Append(new FontScheme { Val = FontSchemeValues.Minor });
            fonts.Append(fontHeader);

            Fills fills = new Fills { Count = 1 };
            Fill fill1 = new Fill(new PatternFill { PatternType = PatternValues.None });
            fills.Append(fill1);

            Borders borders = new Borders { Count = 1 };
            Border border = new Border();
            border.Append(new LeftBorder());
            border.Append(new RightBorder());
            border.Append(new TopBorder());
            border.Append(new BottomBorder());
            border.Append(new DiagonalBorder());
            borders.Append(border);

            // Numbers
            NumberingFormats numberingFormats = new NumberingFormats { Count = 0 };
            for (int index = 0; index < Table.Columns.Count; index++)
            {
                var tableColumn = Table.Columns[index];
                if (tableColumn.Format != null)
                {
                    numberingFormats.Count++;
                    numberingFormats.Append(new NumberingFormat { NumberFormatId = 164 + (uint)index, FormatCode = tableColumn.Format });
                }
            }

            CellStyleFormats cellStyleFormats = new CellStyleFormats { Count = 2 };
            cellStyleFormats.Append(new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 });
            
            CellFormats cellFormats = new CellFormats { Count = 3 };
            // Default
            cellFormats.Append(new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 });
            // Date
            cellFormats.Append(new CellFormat { NumberFormatId = 14, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0, ApplyNumberFormat = true });
            // Header
            cellFormats.Append(new CellFormat { NumberFormatId = 0, FontId = 1, FillId = 0, BorderId = 0, FormatId = 0, ApplyFont = true }); 

            // Columns
            for (int index = 0; index < Table.Columns.Count; index++)
            {
                var tableColumn = Table.Columns[index];
                _styleIndex[tableColumn] = cellFormats.Count;
                if (tableColumn.Format != null)
                {
                    cellFormats.Append(new CellFormat { NumberFormatId = 164 + (uint)index, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0, ApplyNumberFormat = true });
                }
                else
                {
                    cellFormats.Append(new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 });
                }

                cellFormats.Count++;
            }


            CellStyles cellStyles = new CellStyles { Count = 1 };
            CellStyle cellStyle1 = new CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 };

            cellStyles.Append(cellStyle1);
            DifferentialFormats differentialFormats = new DifferentialFormats { Count = 0 };
            TableStyles tableStyles = new TableStyles { Count = 0, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            StylesheetExtensionList stylesheetExtensionList = new StylesheetExtensionList();

            stylesheet.Append(numberingFormats);
            stylesheet.Append(fonts);
            stylesheet.Append(fills);
            stylesheet.Append(borders);
            stylesheet.Append(cellStyleFormats);
            stylesheet.Append(cellFormats);
            stylesheet.Append(cellStyles);
            stylesheet.Append(differentialFormats);
            stylesheet.Append(tableStyles);
            stylesheet.Append(stylesheetExtensionList);

            part.Stylesheet = stylesheet;
        }

        private Row AppendHeaders(SheetData sheetData, uint rowIndex)
        {
            Row row = new Row();
            row.RowIndex = new UInt32Value(rowIndex + 1); // Base index 1

            uint columnIndex = 0;
            foreach (TableColumn<T> column in Table.Columns)
            {
                Cell cell = new Cell();
                cell.CellReference = GetColumnIndex(columnIndex) + row.RowIndex;
                cell.DataType = CellValues.InlineString;
                cell.AppendChild(new InlineString(new Text(string.Format(CultureInfo.InvariantCulture, "{0}", column.Title))));
                cell.StyleIndex = 2u;
                row.AppendChild(cell);
                columnIndex++;
            }

            sheetData.AppendChild(row);
            return row;
        }

        private Row AppendRow(SheetData sheetData, uint rowIndex, T value)
        {
            Row row = new Row();
            row.RowIndex = new UInt32Value(rowIndex + 1); // Base index 1

            uint columnIndex = 0;
            foreach (TableColumn<T> column in Table.Columns)
            {
                AppendCell(row, row.RowIndex, columnIndex, column, value);
                columnIndex++;
            }

            sheetData.AppendChild(row);
            return row;
        }

        private Cell AppendCell(Row row, uint rowIndex, uint columnIndex, TableColumn<T> column, T value)
        {
            if (value == null)
                return null;

            var displayValue = column.GetValue(value);
            if (displayValue == null)
                return null;

            Cell cell = new Cell();
            cell.CellReference = GetColumnIndex(columnIndex) + rowIndex;

            CellValues datatype;
            if (column.DataType == null)
            {
                datatype = GetDataType(displayValue.GetType());
            }
            else
            {
                datatype = GetDataType(column.DataType);
            }

            switch (datatype)
            {
                case CellValues.Boolean:
                    bool boolean;
                    if (!ConvertUtilities.TryChangeType(displayValue, column.Culture, out boolean))
                        return null;

                    cell.DataType = CellValues.Boolean;
                    cell.CellValue = new CellValue(BooleanValue.FromBoolean(boolean));
                    break;

                case CellValues.Number:
                    string number;
                    if (!ConvertUtilities.TryChangeType(displayValue, CultureInfo.InvariantCulture, out number))
                        return null;

                    cell.DataType = CellValues.Number;
                    cell.CellValue = new CellValue(number);
                    break;

                case CellValues.String:
                    string formula;
                    if (!ConvertUtilities.TryChangeType(displayValue, column.Culture, out formula))
                        return null;

                    cell.DataType = CellValues.InlineString;
                    cell.CellFormula = new CellFormula(formula);
                    break;

                case CellValues.InlineString:
                    string text;
                    if (!ConvertUtilities.TryChangeType(displayValue, column.Culture, out text))
                        return null;

                    cell.DataType = CellValues.InlineString;
                    cell.AppendChild(new InlineString(new Text(text)));
                    break;

                case CellValues.Date:
                    DateTime datetime;
                    if (!ConvertUtilities.TryChangeType(displayValue, column.Culture, out datetime))
                        return null;

                    //cell.DataType = CellValues.Date;
                    cell.CellValue = new CellValue(datetime.ToOADate().ToString(CultureInfo.InvariantCulture));
                    cell.StyleIndex = 1;
                    break;

                default:
                    throw new ArgumentOutOfRangeException();
            }

            if (column.Format != null)
            {
                cell.StyleIndex = GetStyleIndex(column);
            }

            row.AppendChild(cell);
            return cell;
        }

        private static string GetColumnIndex(uint columnIndex)
        {
            StringBuilder sb = new StringBuilder();

            const string columnNames = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            while (columnIndex != 0)
            {
                sb.Append(columnNames[(int)(columnIndex % 26)]);
                columnIndex = columnIndex / 26;
            }

            return sb.ToString();
        }

        public static CellValues GetDataType(Type type)
        {
            if (type == null) throw new ArgumentNullException("type");

            TypeCode typeCode = Type.GetTypeCode(type);

            switch (typeCode)
            {
                case TypeCode.Empty:
                case TypeCode.DBNull:
                case TypeCode.Object:
                    return CellValues.String;
                case TypeCode.Boolean:
                    return CellValues.Boolean;
                case TypeCode.SByte:
                case TypeCode.Byte:
                case TypeCode.Int16:
                case TypeCode.UInt16:
                case TypeCode.Int32:
                case TypeCode.UInt32:
                case TypeCode.Int64:
                case TypeCode.UInt64:
                case TypeCode.Single:
                case TypeCode.Double:
                case TypeCode.Decimal:
                    return CellValues.Number;
                case TypeCode.DateTime:
                    return CellValues.Date;
                case TypeCode.Char:
                case TypeCode.String:
                    return CellValues.InlineString;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        // ReSharper restore PossiblyMistakenUseOfParamsMethod
    }
}