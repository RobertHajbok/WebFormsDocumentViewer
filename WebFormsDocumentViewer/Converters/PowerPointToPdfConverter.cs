using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Web;
using WebFormsDocumentViewer.Infrastructure;

namespace WebFormsDocumentViewer.Converters
{
    /// <summary>
    /// Helper class to convert PowerPoint documents to PDF
    /// </summary>
    /// <remarks>
    /// You should also have Microsoft Office installed on the server for this to work
    /// </remarks>
    internal class PowerPointToPdfConverter : IConverter
    {
        /// <summary>
        /// Converts the PowerPoint document to PDF
        /// </summary>
        /// <param name="filePath">Path to the PowerPoint file</param>
        /// <param name="destinationPath">Directory where the PDF file will be saved</param>
        /// <returns>Path of the converted file</returns>
        public string Convert(string filePath, string destinationPath)
        {
            Application appPowerPoint = new Application();
            Presentation powerPointDocument = null;
            try
            {
                powerPointDocument = appPowerPoint.Presentations.Open(HttpContext.Current.Server.MapPath(filePath));
                string fileName = Path.GetFileNameWithoutExtension(filePath) + DateTime.Now.Ticks + ".pdf";

                if (!Directory.Exists(HttpContext.Current.Server.MapPath(destinationPath)))
                    Directory.CreateDirectory(HttpContext.Current.Server.MapPath(destinationPath));
                string newFilePath = Path.Combine(destinationPath, fileName);

                powerPointDocument.ExportAsFixedFormat(HttpContext.Current.Server.MapPath(newFilePath), PpFixedFormatType.ppFixedFormatTypePDF);
                return newFilePath;
            }
            catch
            {
                return null;
            }
            finally
            {
                powerPointDocument?.Close();
                appPowerPoint.Quit();

                // Power point not closing up so easily, so this works after a couple of tries from
                // https://stackoverflow.com/questions/981547/powerpoint-launched-via-c-sharp-does-not-quit
                Process[] processes = Process.GetProcesses();
                foreach(var process in processes.Where(x => x.ProcessName.ToLower().Contains("powerpnt")))
                {
                    process.Kill();
                }
            }
        }
    }
}
