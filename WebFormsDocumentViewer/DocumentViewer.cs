using System;
using System.ComponentModel;
using System.Drawing.Design;
using System.IO;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.Design;
using System.Web.UI.WebControls;
using WebFormsDocumentViewer.Infrastructure;

namespace WebFormsDocumentViewer
{
    [DefaultProperty("FilePath")]
    [ToolboxData("<{0}:DocumentViewer runat=server></{0}:DocumentViewer>")]
    public class DocumentViewer : WebControl
    {
        private string filePath;
        private string tempDirectoryPath;
        private PdfRenderers pdfRenderers;

        [Category("Source File")]
        [Browsable(true)]
        [Description("Set path to source file.")]
        [UrlProperty, Editor(typeof(UrlEditor), typeof(UITypeEditor))]
        public string FilePath
        {
            get
            {
                return filePath;
            }
            set
            {
                filePath = string.IsNullOrEmpty(value) ? string.Empty : value;
            }
        }

        [Category("Temporary Directory Path")]
        [Browsable(true)]
        [Description("Set path to the directory where the files will be converted.")]
        [UrlProperty, Editor(typeof(UrlEditor), typeof(UITypeEditor))]
        public string TempDirectoryPath
        {
            get
            {
                return string.IsNullOrEmpty(tempDirectoryPath) ? "~/Temp" : tempDirectoryPath;
            }
            set
            {
                tempDirectoryPath = string.IsNullOrEmpty(value) ? string.Empty : value;
            }
        }

        [Category("PDF Renderer")]
        [Browsable(true)]
        [Description("Set the PDF renderer for PDF documents or documents that are converted to PDF. Adobe Reader is used by default")]
        public PdfRenderers PdfRenderer
        {
            get
            {
                return pdfRenderers;
            }
            set
            {
                pdfRenderers = value;
            }
        }

        public override void RenderControl(HtmlTextWriter writer)
        {
            writer.RenderBeginTag(HtmlTextWriterTag.Div);
            try
            {
                writer.Write(BuildControl(HttpContext.Current.Server.MapPath(FilePath), ResolveUrl(FilePath), HttpContext.Current.Server.MapPath(TempDirectoryPath),
                    ResolveUrl(TempDirectoryPath), HttpContext.Current.Request.Url.GetLeftPart(UriPartial.Authority), ResolveUrl("~/")).ToString());
            }
            catch
            {
                writer.Write("Cannot display document viewer");
            }
            writer.RenderEndTag();
        }

        public StringBuilder BuildControl(string filePhysicalPath, string fileVirtualPath, string tempDirectoryPhysicalPath,
            string tempDirectoryVirtualPath, string appDomain, string appRootUrl)
        {
            try
            {
                string fileExtension = Path.GetExtension(fileVirtualPath);
                string frameSource = fileVirtualPath;
                SupportedExtensions extension = (SupportedExtensions)Enum.Parse(typeof(SupportedExtensions), fileExtension.Replace(".", ""));
                IConverter converter = ConverterFactory.GetConverter(extension);
                if (converter != null)
                {
                    string tempFileName = converter.Convert(filePhysicalPath, tempDirectoryPhysicalPath);
                    if (string.IsNullOrEmpty(tempFileName))
                        throw new Exception("An error ocurred while trying to convert the file");

                    frameSource = string.Format("{0}/{1}", tempDirectoryVirtualPath, tempFileName);
                }

                if (PdfRenderer == PdfRenderers.PdfJs && Enum.IsDefined(typeof(PdfJsSupportedExtensions), extension.ToString()))
                    frameSource = string.Format("{0}{1}Scripts/pdf.js/web/viewer.html?file={0}{2}", appDomain, appRootUrl, frameSource);
                else
                    frameSource = string.Format("{0}/{1}", appDomain, frameSource);

                StringBuilder sb = new StringBuilder();
                sb.Append("<iframe src=" + frameSource + " ");
                sb.Append("width=" + Width.ToString() + " ");
                sb.Append("height=" + Height.ToString() + ">");
                sb.Append("</iframe>");
                return sb;
            }
            catch
            {
                return new StringBuilder("Cannot display document viewer");
            }
        }

    }
}
