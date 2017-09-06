using NUnit.Framework;
using System;
using System.Text;
using System.Web.UI.WebControls;

namespace WebFormsDocumentViewer.Tests
{
    [TestFixture]
    public class DocumentViewerTests
    {
        [Test]
        public void FilePath_When_PathStartsWithTilde_Then_TidleIsRemoved()
        {
            DocumentViewer documentViewer = new DocumentViewer
            {
                FilePath = "~/Samples/sample.pdf"
            };
            Assert.That(documentViewer.FilePath, Is.EqualTo("Samples/sample.pdf"));
        }

        [Test]
        public void TempDirectoryPath_When_PathStartsWithTilde_Then_TidleIsRemoved()
        {
            DocumentViewer documentViewer = new DocumentViewer
            {
                TempDirectoryPath = "~/Temp"
            };
            Assert.That(documentViewer.TempDirectoryPath, Is.EqualTo("Temp"));
        }

        [Test]
        public void BuildControl_When_ExtensionIsNotSupported_Then_DocumentIsNotDisplayed()
        {
            DocumentViewer documentViewer = new DocumentViewer
            {
                FilePath = ""
            };
            Assert.That(documentViewer.BuildControl("", "").ToString(), Is.EqualTo(new StringBuilder("Cannot display document viewer").ToString()));
        }

        [Test]
        public void BuildControl_When_PdfJsIsUsed_Then_FilePathIsSetAccordingly()
        {
            DocumentViewer documentViewer = new DocumentViewer
            {
                PdfRenderer = PdfRenderers.PdfJs,
                FilePath = "sample.pdf"
            };
            documentViewer.BuildControl("", "");
            Assert.That(documentViewer.FilePath.StartsWith("/Scripts/pdf.js/web/viewer.html?file=../../../"));
        }

        [Test]
        public void BuildControl_When_NoErrorsOccur_Then_IFrameIsSetUp()
        {
            DocumentViewer documentViewer = new DocumentViewer
            {
                FilePath = "sample.pdf",
                Width = Unit.Pixel(500),
                Height = Unit.Pixel(500)
            };
            Assert.That(documentViewer.BuildControl("", "").ToString(), Is.EqualTo("<iframe src=sample.pdf width=500px height=500px></iframe>"));
        }

        [Test]
        public void BuildControl_When_TempDirectoryIsEmpty_Then_ItIsSetToTemp()
        {
            DocumentViewer documentViewer = new DocumentViewer
            {
                FilePath = "sample.xlsx",
                TempDirectoryPath = ""
            };
            documentViewer.BuildControl("", "");
            Assert.That(documentViewer.TempDirectoryPath, Is.EqualTo("Temp"));
        }

        [Test]
        [Category("INTEGRATION")]
        public void BuildControl_When_FileCannotBeConverted_Then_DocumentIsNotDisplayed()
        {
            DocumentViewer documentViewer = new DocumentViewer
            {
                FilePath = "sample.pptx"
            };
            Assert.That(documentViewer.BuildControl("", "").ToString(), Is.EqualTo(new StringBuilder("Cannot display document viewer").ToString()));
        }
    }
}