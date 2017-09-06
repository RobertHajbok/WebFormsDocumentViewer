using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using WebFormsDocumentViewer.Infrastructure;

namespace WebFormsDocumentViewer.Converters
{
    /// <summary>
    /// Helper class to convert RichTextFormat documents to PDF. For now this does exactly what the Word converter does,
    /// but there might be some differences, that's why they are separated
    /// </summary>
    /// <remarks>
    /// You should also have Microsoft Office installed on the server for this to work
    /// </remarks>
    public class RtfToPdfConverter : IConverter
    {
        /// <summary>
        /// Converts the RichTextFormat document to PDF
        /// </summary>
        /// <param name="filePath">Path to the RichTextFormat file</param>
        /// <param name="destinationPath">Directory where the PDF file will be saved</param>
        /// <returns>Name of the converted file</returns>
        public string Convert(string filePath, string destinationPath)
        {
            Application appWord = new Application();
            try
            {
                Document rtfDocument = appWord.Documents.Open(filePath);
                string fileName = Path.GetFileNameWithoutExtension(filePath) + DateTime.Now.Ticks + ".pdf";

                if (!Directory.Exists(destinationPath))
                    Directory.CreateDirectory(destinationPath);
                rtfDocument.ExportAsFixedFormat(Path.Combine(destinationPath, fileName), WdExportFormat.wdExportFormatPDF);

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
