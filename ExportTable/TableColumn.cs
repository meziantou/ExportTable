using System;
using System.Globalization;

namespace ExportTable
{
    public class TableColumn<T>
    {
        public TableColumn()
        {
            Culture = CultureInfo.CurrentCulture;
        }

        public string Format { get; set; }
        public Type DataType { get; set; }
        public string Title { get; set; }
        public Func<T, object> SelectFunction { get; set; }
        public IFormatProvider Culture { get; set; }
        public bool EncodeValue { get; set; }
        public bool ConvertEmptyStringToNull { get; set; }
        public string NullDisplayText { get; set; }

        public object GetValue(T obj)
        {
            try
            {
                if (SelectFunction == null)
                    return null;

                object value = SelectFunction(obj);
                if (value == null)
                    return NullDisplayText;

                if (ConvertEmptyStringToNull && value is string && (string)value == string.Empty)
                {
                    return NullDisplayText;
                }

                return value;
            }
            catch
            {
                return null;
            }
        }

        public string FormatValue(object value)
        {
            if (value == null)
                return null;

            if (this.Format == null)
                return string.Format(Culture, "{0}", value);
            return string.Format(Culture, Format, value);
        }

        public string GetFormattedValue(T obj)
        {
            return FormatValue(GetValue(obj));
        }
    }
}