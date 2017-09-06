using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using WebFormsDocumentViewer.Infrastructure;

namespace WebFormsDocumentViewer.Converters
{
    /// <summary>
    /// Helper class to convert Excel documents to HTML
    /// </summary>
    /// <remarks>
    /// You should also have Microsoft Office installed on the server for this to work
    /// </remarks>
    public class ExcelToHtmlConverter : IConverter
    {
        /// <summary>
        /// Converts the Excel document to HTML
        /// </summary>
        /// <param name="filePath">Path to the Excel file</param>
        /// <param name="destinationPath">Directory where the HTML file will be saved</param>
        /// <returns>Name of the converted file</returns>
        public string Convert(string filePath, string destinationPath)
        {
            Application appExcel = new Application();
            try
            {
                Workbook workbook = appExcel.Workbooks.Open(filePath);
                string fileName = Path.GetFileNameWithoutExtension(filePath) + DateTime.Now.Ticks + ".html";

                if (!Directory.Exists(destinationPath))
                    Directory.CreateDirectory(destinationPath);
                workbook.SaveAs(Path.Combine(destinationPath, fileName), XlFileFormat.xlHtml);

                return fileName;
            }
            catch
            {
                return null;
            }
            finally
            {
                appExcel.Workbooks.Close();
                appExcel.Quit();
            }
        }
    }
}
