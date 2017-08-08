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
        /// <param name="projectRootPath">Root directory of the project or server</param>
        /// <returns>Path of the converted file</returns>
        public string Convert(string filePath, string destinationPath, string projectRootPath)
        {
            Application appExcel = new Application();
            try
            {
                Workbook workbook = appExcel.Workbooks.Open(Path.Combine(projectRootPath, filePath));
                string fileName = Path.GetFileNameWithoutExtension(filePath) + DateTime.Now.Ticks + ".html";

                if (!Directory.Exists(Path.Combine(projectRootPath, destinationPath)))
                    Directory.CreateDirectory(Path.Combine(projectRootPath, destinationPath));
                string newFilePath = Path.Combine(destinationPath, fileName);

                workbook.SaveAs(Path.Combine(projectRootPath, newFilePath), XlFileFormat.xlHtml);
                return newFilePath;
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
