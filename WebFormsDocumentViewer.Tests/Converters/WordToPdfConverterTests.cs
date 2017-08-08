using NUnit.Framework;
using System;
using System.IO;

namespace WebFormsDocumentViewer.Converters.Tests
{
    [TestFixture]
    public class WordToPdfConverterTests
    {
        [Test]
        [Category("INTEGRATION")]
        public void Convert_When_WordFileNotFound_Then_NullIsRetuned()
        {
            WordToPdfConverter converter = new WordToPdfConverter();
            string newFilePath = converter.Convert("test.docx", "Temp", "");
            Assert.That(newFilePath, Is.Null);
        }

        [Test]
        [Category("INTEGRATION")]
        public void Convert_When_WordFileIsFound_Then_NewFilePathIsReturned()
        {
            WordToPdfConverter converter = new WordToPdfConverter();
            string root = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory));
            string newFilePath = converter.Convert("Samples\\sample.docx", "Temp", root);
            Assert.That(newFilePath, Is.Not.Null);
            Assert.That(Path.GetExtension(newFilePath), Is.EqualTo(".pdf"));
        }

        [Test]
        [Category("INTEGRATION")]
        public void Convert_When_WordDestinationPathDoesNotExist_Then_DirectoryIsCreated()
        {
            string currentDateSpan = DateTime.Now.Ticks.ToString();
            WordToPdfConverter converter = new WordToPdfConverter();
            string root = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory));

            Assert.That(Directory.Exists(Path.Combine(root, "Temp" + currentDateSpan)), Is.False);
            converter.Convert("Samples\\sample.docx", "Temp" + currentDateSpan, root);
            Assert.That(Directory.Exists(Path.Combine(root, "Temp" + currentDateSpan)), Is.True);
            Directory.Delete(Path.Combine(root, "Temp" + currentDateSpan), true);
        }
    }
}