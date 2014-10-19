using System;
using System.Collections.Generic;

namespace ExportTable
{
    public class ModelMetadata
    {
        public const int DefaultOrder = 10000;
        
        private Dictionary<string, object> _additionalValues = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        private bool _convertEmptyStringToNull = true;
        private bool _encodeValue = true;
        private int _order = DefaultOrder;
        private bool _showForDisplay = true;
        private string _simpleDisplayText;


        public virtual Dictionary<string, object> AdditionalValues
        {
            get { return _additionalValues; }
        }
        
        public virtual bool ConvertEmptyStringToNull
        {
            get { return _convertEmptyStringToNull; }
            set { _convertEmptyStringToNull = value; }
        }

        public virtual string DataTypeName { get; set; }

        public virtual string Description { get; set; }

        public virtual string DisplayFormatString { get; set; }
        
        public virtual string DisplayName { get; set; }
        
        public virtual bool HideSurroundingHtml { get; set; }

        public virtual bool EncodeValue
        {
            get { return _encodeValue; }
            set { _encodeValue = value; }
        }
        
        public virtual string NullDisplayText { get; set; }

        public virtual int Order
        {
            get { return _order; }
            set { _order = value; }
        }

        public virtual string ShortDisplayName { get; set; }

        public virtual bool ShowForDisplay
        {
            get { return _showForDisplay; }
            set { _showForDisplay = value; }
        }

        public virtual string TemplateHint { get; set; }

        public virtual string Watermark { get; set; }
        
        //internal static ModelMetadata FromLambdaExpression<TParameter, TValue>(Expression<Func<TParameter, TValue>> expression)
        //{
        //    if (expression == null)
        //    {
        //        throw new ArgumentNullException("expression");
        //    }

        //    string propertyName = null;
        //    Type containerType = null;
        //    bool legalExpression = false;

        //    // Need to verify the expression is valid; it needs to at least end in something
        //    // that we can convert to a meaningful string for model binding purposes

        //    switch (expression.Body.NodeType)
        //    {
        //        case ExpressionType.ArrayIndex:
        //            // ArrayIndex always means a single-dimensional indexer; multi-dimensional indexer is a method call to Get()
        //            legalExpression = true;
        //            break;

        //        case ExpressionType.Call:
        //            // Only legal method call is a single argument indexer/DefaultMember call
        //            legalExpression = ExpressionHelper.IsSingleArgumentIndexer(expression.Body);
        //            break;

        //        case ExpressionType.MemberAccess:
        //            // Property/field access is always legal
        //            MemberExpression memberExpression = (MemberExpression)expression.Body;
        //            propertyName = memberExpression.Member is PropertyInfo ? memberExpression.Member.Name : null;
        //            containerType = memberExpression.Expression.Type;
        //            legalExpression = true;
        //            break;

        //        case ExpressionType.Parameter:
        //            // Parameter expression means "model => model", so we delegate to FromModel
        //            return FromModel(viewData, metadataProvider);
        //    }

        //    if (!legalExpression)
        //    {
        //        throw new InvalidOperationException(MvcResources.TemplateHelpers_TemplateLimitations);
        //    }

        //    TParameter container = viewData.Model;
        //    Func<object> modelAccessor = () =>
        //    {
        //        try
        //        {
        //            return CachedExpressionCompiler.Process(expression)(container);
        //        }
        //        catch (NullReferenceException)
        //        {
        //            return null;
        //        }
        //    };

        //    return GetMetadataFromProvider(modelAccessor, typeof(TValue), propertyName, container, containerType, metadataProvider);
        //}
    }
}
