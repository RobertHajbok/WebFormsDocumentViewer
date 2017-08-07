using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Web;
using WebFormsDocumentViewer.Infrastructure;

namespace WebFormsDocumentViewer.Converters
{
    /// <summary>
    /// Helper class to convert Excel documents to HTML
    /// </summary>
    /// <remarks>
    /// You should also have Microsoft Office installed on the server for this to work
    /// </remarks>
    internal class ExcelToHtmlConverter : IConverter
    {
        /// <summary>
        /// Converts the Excel document to HTML
        /// </summary>
        /// <param name="filePath">Path to the Excel file</param>
        /// <param name="destinationPath">Directory where the HTML file will be saved</param>
        /// <returns>Path of the converted file</returns>
        public string Convert(string filePath, string destinationPath)
        {
            Application appExcel = new Application();
            try
            {
                Workbook workbook = appExcel.Workbooks.Open(HttpContext.Current.Server.MapPath(filePath));
                string fileName = Path.GetFileNameWithoutExtension(filePath) + DateTime.Now.Ticks + ".html";

                if (!Directory.Exists(HttpContext.Current.Server.MapPath(destinationPath)))
                    Directory.CreateDirectory(HttpContext.Current.Server.MapPath(destinationPath));
                string newFilePath = Path.Combine(destinationPath, fileName);

                workbook.SaveAs(HttpContext.Current.Server.MapPath(newFilePath), XlFileFormat.xlHtml);
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
