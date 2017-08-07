using System;
using System.ComponentModel;
using System.Drawing.Design;
using System.IO;
using System.Text;
using System.Web.UI;
using System.Web.UI.Design;
using System.Web.UI.WebControls;

namespace WebFormsDocumentViewer
{
    [DefaultProperty("FilePath")]
    [ToolboxData("<{0}:DocumentViewer runat=server></{0}:DocumentViewer>")]
    public class DocumentViewer : WebControl
    {
        private string filePath;
        private string tempDirectoryPath;

        [Category("Source File")]
        [Browsable(true)]
        [Description("Set path to source file.")]
        [Editor(typeof(UrlEditor), typeof(UITypeEditor))]
        public string FilePath
        {
            get
            {
                return filePath;
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    filePath = string.Empty;
                }
                else
                {
                    int tilde = -1;
                    tilde = value.IndexOf('~');
                    filePath = tilde != -1 ? value.Substring((tilde + 2)).Trim() : value;
                }
            }
        }

        public string TempDirectoryPath
        {
            get
            {
                return tempDirectoryPath;
            }
            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    tempDirectoryPath = string.Empty;
                }
                else
                {
                    int tilde = -1;
                    tilde = value.IndexOf('~');
                    tempDirectoryPath = tilde != -1 ? value.Substring((tilde + 2)).Trim() : value;
                }
            }
        }

        public override void RenderControl(HtmlTextWriter writer)
        {
            try
            {
                string fileExtension = Path.GetExtension(FilePath);
                if (fileExtension == ".doc" || fileExtension == ".docx")
                {
                    if (string.IsNullOrEmpty(TempDirectoryPath))
                        TempDirectoryPath = "Temp";
                    FilePath = WordToPdfConverter.Convert(FilePath, TempDirectoryPath);
                    if (string.IsNullOrEmpty(FilePath))
                        throw new Exception("An error ocurred while trying to convert Word to PDF");
                }
                else if (fileExtension == ".ppt" || fileExtension == ".pptx")
                {
                    if (string.IsNullOrEmpty(TempDirectoryPath))
                        TempDirectoryPath = "Temp";
                    FilePath = PowerPointToPdfConverter.Convert(FilePath, TempDirectoryPath);
                    if (string.IsNullOrEmpty(FilePath))
                        throw new Exception("An error ocurred while trying to convert PowerPoint to PDF");
                }
                StringBuilder sb = new StringBuilder();
                sb.Append("<iframe src=" + FilePath?.ToString() + " ");
                sb.Append("width=" + Width.ToString() + " ");
                sb.Append("height=" + Height.ToString() + ">");
                sb.Append("</iframe>");

                writer.RenderBeginTag(HtmlTextWriterTag.Div);
                writer.Write(sb.ToString());
                writer.RenderEndTag();
            }
            catch
            {
                writer.RenderBeginTag(HtmlTextWriterTag.Div);
                writer.Write("Cannot display document viewer");
                writer.RenderEndTag();
            }
        }
    }
}
