using System;
using System.IO;

namespace ExportTable
{
    public class CsvGenerator
    {
        public const string DefaultSeparator = ",";
        public const string DefaultQuote = "\"";
    }

    public class CsvGenerator<T>
    {
        private readonly Table<T> _table;

        public Table<T> Table
        {
            get { return _table; }
        }

        public CsvGenerator(Table<T> table)
        {
            if (table == null) throw new ArgumentNullException("table");
            _table = table;
        }

        public void Generate(string path, CsvOptions options = null)
        {
            using (StreamWriter sw = File.CreateText(path))
            {
                Generate(sw, options);
            }
        }

        public string Generate(CsvOptions options = null)
        {
            using (StringWriter sw = new StringWriter())
            {
                Generate(sw, options);
                return sw.ToString();
            }
        }

        public void Generate(TextWriter writer, CsvOptions options = null)
        {
            if (writer == null) throw new ArgumentNullException("writer");

            if (options == null)
            {
                options = new CsvOptions();
            }

            if (Table.ShowHeader)
            {
                bool first = true;
                foreach (var tableColumn in Table.Columns)
                {
                    if (!first)
                    {
                        writer.Write(options.Separator);
                    }

                    writer.Write(Encode(tableColumn.Title, options.Separator, options.Quote));
                    first = false;
                }

                writer.WriteLine();
            }

            foreach (var item in Table.DataSource)
            {
                bool first = true;
                foreach (var tableColumn in Table.Columns)
                {
                    if (!first)
                    {
                        writer.Write(options.Separator);
                    }

                    string formattedValue = tableColumn.GetFormattedValue(item);
                    writer.Write(Encode(formattedValue, options.Separator, options.Quote));
                    first = false;
                }
                writer.WriteLine();
            }
        }

        private string Encode(string str, string separator, string quote)
        {
            if (string.IsNullOrEmpty(str))
                return str;

            if (string.IsNullOrEmpty(quote))
                return str;

            if (str.Contains(separator))
            {
                string escapedValue = str.Replace(quote, quote + quote);
                str = quote + escapedValue + quote;
            }

            return str;
        }
    }
}