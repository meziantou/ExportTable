using System.Collections.Generic;

namespace ExportTable
{
    public class Table<T>
    {
        public Table()
        {
            Columns = new List<TableColumn<T>>();
        }

        public IList<TableColumn<T>> Columns { get; set; }
        public bool ShowHeader { get; set; }
        public string Title { get; set; }
        public IEnumerable<T> DataSource { get; set; }
    }
}