using NUnit.Framework;
using System;
using System.IO;

namespace WebFormsDocumentViewer.Converters.Tests
{
    [TestFixture]
    public class ExcelToHtmlConverterTests
    {
        [Test]
        [Category("INTEGRATION")]
        public void Convert_When_ExcelFileNotFound_Then_NullIsRetuned()
        {
            ExcelToHtmlConverter converter = new ExcelToHtmlConverter();
            string newFilePath = converter.Convert("test.xlsx", "Temp", "");
            Assert.That(newFilePath, Is.Null);
        }

        [Test]
        [Category("INTEGRATION")]
        public void Convert_When_ExcelFileIsFound_Then_NewFilePathIsReturned()
        {
            ExcelToHtmlConverter converter = new ExcelToHtmlConverter();
            string root = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory));
            string newFilePath = converter.Convert("Samples\\sample.xlsx", "Temp", root);
            Assert.That(newFilePath, Is.Not.Null);
            Assert.That(Path.GetExtension(newFilePath), Is.EqualTo(".html"));
        }

        [Test]
        [Category("INTEGRATION")]
        public void Convert_When_ExcelDestinationPathDoesNotExist_Then_DirectoryIsCreated()
        {
            string currentDateSpan = DateTime.Now.Ticks.ToString();
            ExcelToHtmlConverter converter = new ExcelToHtmlConverter();
            string root = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory));

            Assert.That(Directory.Exists(Path.Combine(root, "Temp" + currentDateSpan)), Is.False);
            converter.Convert("Samples\\sample.xlsx", "Temp" + currentDateSpan, root);
            Assert.That(Directory.Exists(Path.Combine(root, "Temp" + currentDateSpan)), Is.True);
            Directory.Delete(Path.Combine(root, "Temp" + currentDateSpan), true);
        }
    }
}