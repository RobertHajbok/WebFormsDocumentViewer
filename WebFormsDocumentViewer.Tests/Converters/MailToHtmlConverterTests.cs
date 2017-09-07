using NUnit.Framework;
using System;
using System.IO;

namespace WebFormsDocumentViewer.Converters.Tests
{
    [TestFixture]
    public class MailToHtmlConverterTests
    {
        [Test]
        [Category("INTEGRATION")]
        public void Convert_When_MailFileNotFound_Then_NullIsRetuned()
        {
            MailToHtmlConverter converter = new MailToHtmlConverter();
            string newFilePath = converter.Convert("test.msg", "Temp");
            Assert.That(newFilePath, Is.Null);
        }

        [Test]
        [Category("INTEGRATION")]
        public void Convert_When_MsgFileIsFound_Then_NewFilePathIsReturned()
        {
            MailToHtmlConverter converter = new MailToHtmlConverter();
            string root = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory));
            string newFilePath = converter.Convert(Path.Combine(root, "Samples\\sample.msg"), Path.Combine(root, "Temp"));
            Assert.That(newFilePath, Is.Not.Null);
            Assert.That(Path.GetExtension(newFilePath), Is.EqualTo(".html"));
        }

        [Test]
        [Category("INTEGRATION")]
        public void Convert_When_EmlFileIsFound_Then_NewFilePathIsReturned()
        {
            MailToHtmlConverter converter = new MailToHtmlConverter();
            string root = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory));
            string newFilePath = converter.Convert(Path.Combine(root, "Samples\\sample.eml"), Path.Combine(root, "Temp"));
            Assert.That(newFilePath, Is.Not.Null);
            Assert.That(Path.GetExtension(newFilePath), Is.EqualTo(".html"));
        }

        [Test]
        [Category("INTEGRATION")]
        public void Convert_When_MailDestinationPathDoesNotExist_Then_DirectoryIsCreated()
        {
            string currentDateSpan = DateTime.Now.Ticks.ToString();
            MailToHtmlConverter converter = new MailToHtmlConverter();
            string root = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory));

            Assert.That(Directory.Exists(Path.Combine(root, "Temp" + currentDateSpan)), Is.False);
            converter.Convert(Path.Combine(root, "Samples\\sample.msg"), Path.Combine(root, "Temp" + currentDateSpan));
            Assert.That(Directory.Exists(Path.Combine(root, "Temp" + currentDateSpan)), Is.True);
            Directory.Delete(Path.Combine(root, "Temp" + currentDateSpan), true);
        }
    }
}