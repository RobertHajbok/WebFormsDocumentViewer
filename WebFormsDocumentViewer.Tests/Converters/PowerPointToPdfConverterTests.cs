using NUnit.Framework;
using System;
using System.IO;

namespace WebFormsDocumentViewer.Converters.Tests
{
    [TestFixture]
    public class PowerPointToPdfConverterTests
    {
        [Test]
        [Category("INTEGRATION")]
        public void Convert_When_PowerPointFileNotFound_Then_NullIsRetuned()
        {
            PowerPointToPdfConverter converter = new PowerPointToPdfConverter();
            string newFilePath = converter.Convert("test.pptx", "Temp");
            Assert.That(newFilePath, Is.Null);
        }

        [Test]
        [Category("INTEGRATION")]
        public void Convert_When_PowerPointFileIsFound_Then_NewFilePathIsReturned()
        {
            PowerPointToPdfConverter converter = new PowerPointToPdfConverter();
            string root = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory));
            string newFilePath = converter.Convert(Path.Combine(root, "Samples\\sample.pptx"), Path.Combine(root, "Temp"));
            Assert.That(newFilePath, Is.Not.Null);
            Assert.That(Path.GetExtension(newFilePath), Is.EqualTo(".pdf"));
        }

        [Test]
        [Category("INTEGRATION")]
        public void Convert_When_PowerPointIsProtected_Then_FileCanBeConverted()
        {
            PowerPointToPdfConverter converter = new PowerPointToPdfConverter();
            string root = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory));
            string newFilePath = converter.Convert(Path.Combine(root, "Samples\\sample-protected.pptx"), Path.Combine(root, "Temp"));
            Assert.That(newFilePath, Is.Not.Null);
            Assert.That(Path.GetExtension(newFilePath), Is.EqualTo(".pdf"));
        }

        [Test]
        [Category("INTEGRATION")]
        public void Convert_When_PowerPointDestinationPathDoesNotExist_Then_DirectoryIsCreated()
        {
            string currentDateSpan = DateTime.Now.Ticks.ToString();
            PowerPointToPdfConverter converter = new PowerPointToPdfConverter();
            string root = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory));

            Assert.That(Directory.Exists(Path.Combine(root, "Temp" + currentDateSpan)), Is.False);
            converter.Convert(Path.Combine(root, "Samples\\sample.pptx"), Path.Combine(root, "Temp" + currentDateSpan));
            Assert.That(Directory.Exists(Path.Combine(root, "Temp" + currentDateSpan)), Is.True);
            Directory.Delete(Path.Combine(root, "Temp" + currentDateSpan), true);
        }
    }
}