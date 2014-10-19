using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq.Expressions;
using CodeFluent.Runtime.Utilities;
using DocumentFormat.OpenXml.Packaging;

namespace ExportTable
{
    public static class Fluent
    {
        public static Table<T> ToTable<T>(this IEnumerable<T> values, bool showHeader = true, string sheetName = null)
        {
            var table = new Table<T>();
            table.DataSource = values;
            table.ShowHeader = showHeader;
            table.Title = sheetName ?? ConvertUtilities.Decamelize(typeof(T).Name);
            return table;
        }

        public static Table<T> AddColumn<T>(
            this Table<T> table,
            string title = null,
            Type dataType = null,
            string format = null,
            bool? convertEmptyStringToNull = null,
            bool? encodeValue = null,
            string nullDisplayText = null,
            CultureInfo culture = null,
            Func<T, object> select = null)
        {
            return AddColumn<T, object>(table, null, title, dataType, format, convertEmptyStringToNull, encodeValue, nullDisplayText, culture, select);
        }

        public static Table<T> AddColumn<T, TValue>(
            this Table<T> table,
            Expression<Func<T, TValue>> expression = null,
            string title = null,
            Type dataType = null,
            string format = null,
            bool? convertEmptyStringToNull = null,
            bool? encodeValue = null,
            string nullDisplayText = null,
            CultureInfo culture = null,
            Func<T, object> select = null)
        {
            var tableColumn = new TableColumn<T>();
            if (expression != null)
            {
                var metadata = Metadata.Create(expression);
                tableColumn.DataType = dataType ?? (metadata.DataType == null ? typeof(string) : ConvertUtilities.GetType(metadata.DataType)) ?? typeof(string);
                tableColumn.Format = format ?? metadata.DisplayFormat;
                tableColumn.Title = title ?? metadata.DisplayName;
                tableColumn.ConvertEmptyStringToNull = convertEmptyStringToNull ?? metadata.ConvertEmptyStringToNull;
                tableColumn.NullDisplayText = nullDisplayText ?? metadata.NullDisplayText;
                tableColumn.EncodeValue = encodeValue ?? metadata.EncodeValue;
            }
            else
            {
                tableColumn.DataType = dataType ?? typeof(string);
                tableColumn.Format = format;
                tableColumn.Title = title;
                if (convertEmptyStringToNull.HasValue)
                {
                    tableColumn.ConvertEmptyStringToNull = convertEmptyStringToNull.Value;
                }
                tableColumn.NullDisplayText = nullDisplayText;
                if (encodeValue.HasValue)
                {
                    tableColumn.EncodeValue = encodeValue.Value;
                }
            }

            if (culture != null)
            {
                tableColumn.Culture = culture;
            }

            if (select == null && expression != null)
            {
                var func = expression.Compile();
                tableColumn.SelectFunction = obj => func(obj);
            }
            else
            {
                tableColumn.SelectFunction = select;
            }

            table.Columns.Add(tableColumn);
            return table;
        }

        public static void GenerateSpreadsheet<T>(this Table<T> table, string path)
        {
            if (table == null) throw new ArgumentNullException("table");
            if (path == null) throw new ArgumentNullException("path");
            var generator = new SpreadsheetGenerator<T>(table);
            generator.Generate(path);
        }

        public static void GenerateSpreadsheet<T>(this Table<T> table, SpreadsheetDocument document)
        {
            if (table == null) throw new ArgumentNullException("table");
            if (document == null) throw new ArgumentNullException("document");
            var generator = new SpreadsheetGenerator<T>(table);
            generator.Generate(document);
        }

        public static void GenerateSpreadsheet<T>(this Table<T> table, Stream stream)
        {
            if (table == null) throw new ArgumentNullException("table");
            if (stream == null) throw new ArgumentNullException("stream");
            var generator = new SpreadsheetGenerator<T>(table);
            generator.Generate(stream);
        }

        public static string GenerateHtmlTable<T>(this Table<T> table)
        {
            if (table == null) throw new ArgumentNullException("table");
            HtmlTableGenerator<T> generator = new HtmlTableGenerator<T>(table);
            return generator.Generate();
        }

        public static void GenerateHtmlTable<T>(this Table<T> table, TextWriter writer)
        {
            if (table == null) throw new ArgumentNullException("table");
            HtmlTableGenerator<T> generator = new HtmlTableGenerator<T>(table);
            generator.Generate(writer);
        }

        public static void GenerateHtmlTable<T>(this Table<T> table, string path)
        {
            if (table == null) throw new ArgumentNullException("table");
            HtmlTableGenerator<T> generator = new HtmlTableGenerator<T>(table);
            generator.Generate(path);
        }

        public static void GenerateCsv<T>(this Table<T> table, string path, CsvOptions options = null)
        {
            if (table == null) throw new ArgumentNullException("table");
            CsvGenerator<T> generator = new CsvGenerator<T>(table);
            generator.Generate(path, options);
        }

        public static string GenerateCsv<T>(this Table<T> table, CsvOptions options = null)
        {
            if (table == null) throw new ArgumentNullException("table");
            CsvGenerator<T> generator = new CsvGenerator<T>(table);
            return generator.Generate(options);
        }

        public static void GenerateCsv<T>(this Table<T> table, TextWriter writer, CsvOptions options = null)
        {
            if (table == null) throw new ArgumentNullException("table");
            CsvGenerator<T> generator = new CsvGenerator<T>(table);
            generator.Generate(writer, options);
        }
    }
}