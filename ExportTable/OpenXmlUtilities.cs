using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace ExportTable
{
    public static class OpenXmlUtilities
    {
        public static bool ValidateSpreadsheet(string path)
        {
            IList<ValidationErrorInfo> errors;
            return ValidateSpreadsheet(path, out errors);
        }

        public static bool ValidateSpreadsheet(string path, out IList<ValidationErrorInfo> errors)
        {
            using (OpenXmlPackage package = SpreadsheetDocument.Open(path, false))
            {
                return ValidateSpreadsheet(package, out errors);
            }
        }

        private static bool ValidateSpreadsheet(OpenXmlPackage document)
        {
            IList<ValidationErrorInfo> errors;
            return ValidateSpreadsheet(document, out errors);
        }

        public static bool ValidateSpreadsheet(OpenXmlPackage document, out IList<ValidationErrorInfo> errors)
        {
            OpenXmlValidator openXmlValidator = new OpenXmlValidator();
            var validationErrorInfos = openXmlValidator.Validate(document);
            if (validationErrorInfos == null)
            {
                errors = null;
                return true;
            }
            else
            {
                errors = validationErrorInfos.ToList();
                return !errors.Any();
            }
        }
    }
}