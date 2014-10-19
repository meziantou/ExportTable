using System;
using System.IO;
using System.Web;

namespace ExportTable
{
    public class HtmlTableGenerator<T>
    {
        private readonly Table<T> _table;

        public Table<T> Table
        {
            get { return _table; }
        }

        public HtmlTableGenerator(Table<T> table)
        {
            _table = table;
        }

        public string Generate()
        {
            using (StringWriter sw = new StringWriter())
            {
                Generate(sw);
                return sw.ToString();
            }
        }

        public void Generate(string path, bool createFullPage = true)
        {
            using (StreamWriter sw = File.CreateText(path))
            {
                Generate(sw, createFullPage);
            }
        }

        public void Generate(TextWriter writer, bool createFullPage = false)
        {
            if (writer == null) throw new ArgumentNullException("writer");

            if (createFullPage)
            {
                writer.Write(@"<!doctype html><html><head><title>{0}</title></head><body>", HttpUtility.HtmlEncode(Table.Title));
            }

            writer.Write("<table>");
            if (Table.ShowHeader)
            {
                writer.Write("<thead>");
                writer.Write("<tr>");
                foreach (var tableColumn in Table.Columns)
                {
                    writer.Write("<th>");
                    writer.Write(HttpUtility.HtmlEncode(tableColumn.Title));
                    writer.Write("</th>");
                }
                writer.Write("</tr>");
                writer.Write("</thead>");
            }

            writer.Write("<tbody>");
            foreach (var item in Table.DataSource)
            {
                writer.Write("<tr>");
                foreach (var tableColumn in Table.Columns)
                {
                    writer.Write("<td>");
                    string formattedValue = tableColumn.GetFormattedValue(item);
                    if (tableColumn.EncodeValue)
                    {
                        formattedValue = HttpUtility.HtmlEncode(formattedValue);
                    }

                    writer.Write(formattedValue);
                    writer.Write("</td>");
                }
                writer.Write("</tr>");
            }

            writer.Write("</tbody>");
            writer.Write("</table>");


            if (createFullPage)
            {
                writer.Write("</body></html>");
            }
        }
    }
}