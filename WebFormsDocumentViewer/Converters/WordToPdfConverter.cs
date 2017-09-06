using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using WebFormsDocumentViewer.Infrastructure;

namespace WebFormsDocumentViewer.Converters
{
    /// <summary>
    /// Helper class to convert Word documents to PDF
    /// </summary>
    /// <remarks>
    /// You should also have Microsoft Office installed on the server for this to work
    /// </remarks>
    public class WordToPdfConverter : IConverter
    {
        /// <summary>
        /// Converts the Word document to PDF
        /// </summary>
        /// <param name="filePath">Path to the Word file</param>
        /// <param name="destinationPath">Directory where the PDF file will be saved</param>
        /// <returns>Name of the converted file</returns>
        public string Convert(string filePath, string destinationPath)
        {
            Application appWord = new Application();
            try
            {
                Document wordDocument = appWord.Documents.Open(filePath);
                string fileName = Path.GetFileNameWithoutExtension(filePath) + DateTime.Now.Ticks + ".pdf";

                if (!Directory.Exists(destinationPath))
                    Directory.CreateDirectory(destinationPath);
                wordDocument.ExportAsFixedFormat(Path.Combine(destinationPath, fileName), WdExportFormat.wdExportFormatPDF);

                return fileName;
            }
            catch
            {
                return null;
            }
            finally
            {
                // appWord.Documents.Close removed because sometimes it prompts for save changes
                // even if WdSaveOptions.wdDoNotSaveChanges is set
                appWord.Quit(WdSaveOptions.wdDoNotSaveChanges, WdOriginalFormat.wdOriginalDocumentFormat, false);
            }
        }
    }
}
