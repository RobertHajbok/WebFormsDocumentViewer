using NUnit.Framework;
using WebFormsDocumentViewer.Converters;

namespace WebFormsDocumentViewer.Infrastructure.Tests
{
    [TestFixture]
    public class ConverterFactoryTests
    {
        [Test]
        public void GetConverter_When_ConverterIsRequested_Then_FactoryReturnsCorrectInstance()
        {
            Assert.That(ConverterFactory.GetConverter(SupportedExtensions.doc), Is.TypeOf<WordToPdfConverter>());
            Assert.That(ConverterFactory.GetConverter(SupportedExtensions.docx), Is.TypeOf<WordToPdfConverter>());
            Assert.That(ConverterFactory.GetConverter(SupportedExtensions.pdf), Is.Null);
            Assert.That(ConverterFactory.GetConverter(SupportedExtensions.ppt), Is.TypeOf<PowerPointToPdfConverter>());
            Assert.That(ConverterFactory.GetConverter(SupportedExtensions.pptx), Is.TypeOf<PowerPointToPdfConverter>());
            Assert.That(ConverterFactory.GetConverter(SupportedExtensions.xls), Is.TypeOf<ExcelToHtmlConverter>());
            Assert.That(ConverterFactory.GetConverter(SupportedExtensions.xlsx), Is.TypeOf<ExcelToHtmlConverter>());
            Assert.That(ConverterFactory.GetConverter(SupportedExtensions.rtf), Is.TypeOf<RtfToPdfConverter>());
            Assert.That(ConverterFactory.GetConverter(SupportedExtensions.txt), Is.Null);
            Assert.That(ConverterFactory.GetConverter(SupportedExtensions.msg), Is.TypeOf<MailToHtmlConverter>());
            Assert.That(ConverterFactory.GetConverter(SupportedExtensions.eml), Is.TypeOf<MailToHtmlConverter>());
        }
    }
}