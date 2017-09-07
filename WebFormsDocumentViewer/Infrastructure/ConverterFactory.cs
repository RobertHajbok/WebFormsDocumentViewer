using WebFormsDocumentViewer.Converters;

namespace WebFormsDocumentViewer.Infrastructure
{
    public class ConverterFactory
    {
        public static IConverter GetConverter(SupportedExtensions extension)
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
                case SupportedExtensions.rtf:
                    converter = new RtfToPdfConverter();
                    break;
                case SupportedExtensions.eml:
                case SupportedExtensions.msg:
                    converter = new MailToHtmlConverter();
                    break;
            }
            return converter;
        }
    }
}
