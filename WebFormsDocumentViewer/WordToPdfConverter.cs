using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using System.Web;

namespace WebFormsDocumentViewer
{
    /// <summary>
    /// Helper class to convert Word documents to PDF
    /// </summary>
    /// <remarks>
    /// You should also have Microsoft Office installed on the server for this to work
    /// </remarks>
    public static class WordToPdfConverter
    {
        /// <summary>
        /// Converts the Word document to PDF
        /// </summary>
        /// <param name="filePath">Path to the Word file</param>
        /// <param name="destinationPath">Directory where the PDF file will be saved</param>
        /// <returns>Path of the converted file</returns>
        public static string Convert(string filePath, string destinationPath)
        {
            Application appWord = new Application();
            try
            {
                Document wordDocument = appWord.Documents.Open(HttpContext.Current.Server.MapPath(filePath));
                string fileName = Path.GetFileNameWithoutExtension(filePath) + DateTime.Now.Ticks + ".pdf";

                if (!Directory.Exists(HttpContext.Current.Server.MapPath(destinationPath)))
                    Directory.CreateDirectory(HttpContext.Current.Server.MapPath(destinationPath));
                string newFilePath = Path.Combine(destinationPath, fileName);

                wordDocument.ExportAsFixedFormat(HttpContext.Current.Server.MapPath(newFilePath), WdExportFormat.wdExportFormatPDF);
                return newFilePath;
            }
            catch
            {
                return null;
            }
            finally
            {
                appWord.Documents.Close(WdSaveOptions.wdDoNotSaveChanges, WdOriginalFormat.wdOriginalDocumentFormat, false);
                appWord.Quit();
            }
        }
    }
}
