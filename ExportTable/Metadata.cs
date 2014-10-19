using System;
using System.ComponentModel.DataAnnotations;
using System.Linq.Expressions;
using System.Reflection;
using CodeFluent.Runtime.Utilities;

namespace ExportTable
{
    public class Metadata
    {
        public Metadata()
        {
            EncodeValue = true;
        }

        public bool Hidden { get; set; }
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public string ShortName { get; set; }
        public string DisplayFormat { get; set; }
        public string Description { get; set; }
        public string DataType { get; set; }
        public string GroupName { get; set; }
        public int Order { get; set; }
        public string NullDisplayText { get; set; }
        public bool ConvertEmptyStringToNull { get; set; }
        public bool EncodeValue { get; set; }

        public static Metadata Create<T, TValue>(Expression<Func<T, TValue>> expression)
        {
            if (expression == null) throw new ArgumentNullException("expression");

            Expression expressionBody = expression.Body;
            MemberExpression memberExpression = expressionBody as MemberExpression;
            if (memberExpression != null)
                return Create(memberExpression.Member);

            return new Metadata();
        }

        public static Metadata Create(MemberInfo memberInfo)
        {
            if (memberInfo == null) throw new ArgumentNullException("memberInfo");

            Metadata metadata = new Metadata();

            // Name
            metadata.Name = memberInfo.Name;

            // DisplayAttribute
            DisplayAttribute displayAttribute = memberInfo.GetCustomAttribute<DisplayAttribute>();
            if (displayAttribute != null)
            {
                metadata.DisplayName = displayAttribute.GetName();
                metadata.ShortName = displayAttribute.GetShortName();
                metadata.GroupName = displayAttribute.GetGroupName();
                metadata.Description = displayAttribute.GetDescription();

                int? order = displayAttribute.GetOrder();
                if (order != null)
                {
                    metadata.Order = order.Value;
                }
            }

            if (metadata.DisplayName == null)
            {
                metadata.DisplayName = ConvertUtilities.Decamelize(metadata.Name);
            }

            // DataType
            DataTypeAttribute dataTypeAttribute = memberInfo.GetCustomAttribute<DataTypeAttribute>();
            if (dataTypeAttribute != null)
            {
                metadata.DataType = dataTypeAttribute.GetDataTypeName();
                Fill(metadata, dataTypeAttribute.DisplayFormat);
            }
            if (metadata.DataType == null)
            {
                PropertyInfo pi = memberInfo as PropertyInfo;
                if (pi != null)
                {
                    metadata.DataType = pi.PropertyType.AssemblyQualifiedName;
                }
                else
                {
                    FieldInfo fi = memberInfo as FieldInfo;
                    if (fi != null)
                    {
                        metadata.DataType = fi.FieldType.AssemblyQualifiedName;
                    }
                }

            }

            // DisplayFormat
            DisplayFormatAttribute displayFormatAttribute = memberInfo.GetCustomAttribute<DisplayFormatAttribute>();
            if (displayFormatAttribute != null)
            {
                Fill(metadata, displayFormatAttribute);
            }

            // ScaffoldColumnAttribute
            ScaffoldColumnAttribute scaffoldColumnAttribute = memberInfo.GetCustomAttribute<ScaffoldColumnAttribute>();
            if (scaffoldColumnAttribute != null)
            {
                metadata.Hidden = scaffoldColumnAttribute.Scaffold;
            }

            return metadata;
        }

        private static void Fill(Metadata metadata, DisplayFormatAttribute attribute)
        {
            if (metadata == null) throw new ArgumentNullException("metadata");
            if (attribute == null)
                return;

            metadata.DisplayFormat = attribute.DataFormatString;
            metadata.EncodeValue = attribute.HtmlEncode;
            metadata.ConvertEmptyStringToNull = attribute.ConvertEmptyStringToNull;
            metadata.NullDisplayText = attribute.NullDisplayText;

        }

    }
}