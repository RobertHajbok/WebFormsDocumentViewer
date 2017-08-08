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
        /// <param name="projectRootPath">Root directory of the project or server</param>
        /// <returns>Path of the converted file</returns>
        public string Convert(string filePath, string destinationPath, string projectRootPath)
        {
            Application appWord = new Application();
            try
            {
                Document wordDocument = appWord.Documents.Open(Path.Combine(projectRootPath, filePath));
                string fileName = Path.GetFileNameWithoutExtension(filePath) + DateTime.Now.Ticks + ".pdf";

                if (!Directory.Exists(Path.Combine(projectRootPath, destinationPath)))
                    Directory.CreateDirectory(Path.Combine(projectRootPath, destinationPath));
                string newFilePath = Path.Combine(destinationPath, fileName);

                wordDocument.ExportAsFixedFormat(Path.Combine(projectRootPath, newFilePath), WdExportFormat.wdExportFormatPDF);
                return newFilePath;
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
