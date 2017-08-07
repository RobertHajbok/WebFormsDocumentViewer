using WebFormsDocumentViewer.Converters;

namespace WebFormsDocumentViewer.Infrastructure
{
    internal class ConverterFactory
    {
        internal static IConverter GetConverter(SupportedExtensions extension)
        {
            IConverter converter = null;
            switch (extension)
            {
                case SupportedExtensions.doc:
                case SupportedExtensions.docx:
                    converter = new WordToPdfConverter();
                    break;
                case SupportedExtensions.ppt:
                case SupportedExtensions.pptx:
                    converter = new PowerPointToPdfConverter();
                    break;
                case SupportedExtensions.xls:
                case SupportedExtensions.xlsx:
                    converter = new ExcelToHtmlConverter();
                    break;
            }
            return converter;
        }
    }
}
