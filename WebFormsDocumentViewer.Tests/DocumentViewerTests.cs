using NUnit.Framework;
using System.IO;
using System.Text;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebFormsDocumentViewer.Tests
{
    [TestFixture]
    public class DocumentViewerTests
    {
        [Test]
        public void TempDirectoryPath_When_PathIsEmpty_Then_TempFolderIsSet()
        {
            DocumentViewer documentViewer = new DocumentViewer
            {
                TempDirectoryPath = ""
            };
            Assert.That(documentViewer.TempDirectoryPath, Is.EqualTo("~/Temp"));
        }

        [Test]
        public void BuildControl_When_ExtensionIsNotSupported_Then_DocumentIsNotDisplayed()
        {
            DocumentViewer documentViewer = new DocumentViewer
            {
                FilePath = ""
            };
            Assert.That(documentViewer.BuildControl("", "", "", "", "", "").ToString(), Is.EqualTo(new StringBuilder("Cannot display document viewer").ToString()));
        }

        [Test]
        public void BuildControl_When_PdfJsIsUsed_Then_FilePathIsSetAccordingly()
        {
            DocumentViewer documentViewer = new DocumentViewer
            {
                PdfRenderer = PdfRenderers.PdfJs
            };
            Assert.That(documentViewer.BuildControl("", "sample.pdf", "", "", "", "").ToString().Contains("Scripts/pdf.js/web/viewer.html?file="));
        }

        [Test]
        public void BuildControl_When_NoErrorsOccur_Then_IFrameIsSetUp()
        {
            DocumentViewer documentViewer = new DocumentViewer
            {
                Width = Unit.Pixel(500),
                Height = Unit.Pixel(500)
            };
            Assert.That(documentViewer.BuildControl("", "sample.pdf", "", "", "", "").ToString(), Is.EqualTo("<iframe src=/sample.pdf width=500px height=500px></iframe>"));
        }

        [Test]
        [Category("INTEGRATION")]
        public void BuildControl_When_FileCannotBeConverted_Then_DocumentIsNotDisplayed()
        {
            DocumentViewer documentViewer = new DocumentViewer
            {
                FilePath = "sample.pptx",
                TempDirectoryPath = "TmpFolder"
            };
            Assert.That(documentViewer.BuildControl("", "sample.pptx", "", "", "", "").ToString(), Is.EqualTo(new StringBuilder("Cannot display document viewer").ToString()));
        }

        [Test]
        public void RenderControl_When_ErrorOccursOnRetrievingServerParameters_Then_ControlIsNotRendered()
        {
            DocumentViewer documentViewer = new DocumentViewer();
            HtmlTextWriter textWriter = new HtmlTextWriter(new StringWriter());
            documentViewer.RenderControl(textWriter);
            Assert.That(textWriter.InnerWriter.ToString(), Is.EqualTo("<div>\r\n\tCannot display document viewer\r\n</div>"));
        }
    }
}