using NUnit.Framework;
using System;
using System.IO;

namespace WebFormsDocumentViewer.Converters.Tests
{
    [TestFixture]
    public class RtfToPdfConverterTests
    {
        [Test]
        [Category("INTEGRATION")]
        public void Convert_When_RtfFileNotFound_Then_NullIsRetuned()
        {
            RtfToPdfConverter converter = new RtfToPdfConverter();
            string newFilePath = converter.Convert("test.rtf", "Temp", "");
            Assert.That(newFilePath, Is.Null);
        }

        [Test]
        [Category("INTEGRATION")]
        public void Convert_When_RtfFileIsFound_Then_NewFilePathIsReturned()
        {
            RtfToPdfConverter converter = new RtfToPdfConverter();
            string root = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory));
            string newFilePath = converter.Convert("Samples\\sample.rtf", "Temp", root);
            Assert.That(newFilePath, Is.Not.Null);
            Assert.That(Path.GetExtension(newFilePath), Is.EqualTo(".pdf"));
        }

        [Test]
        [Category("INTEGRATION")]
        public void Convert_When_RtfDestinationPathDoesNotExist_Then_DirectoryIsCreated()
        {
            string currentDateSpan = DateTime.Now.Ticks.ToString();
            RtfToPdfConverter converter = new RtfToPdfConverter();
            string root = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory));

            Assert.That(Directory.Exists(Path.Combine(root, "Temp" + currentDateSpan)), Is.False);
            converter.Convert("Samples\\sample.rtf", "Temp" + currentDateSpan, root);
            Assert.That(Directory.Exists(Path.Combine(root, "Temp" + currentDateSpan)), Is.True);
            Directory.Delete(Path.Combine(root, "Temp" + currentDateSpan), true);
        }
    }
}