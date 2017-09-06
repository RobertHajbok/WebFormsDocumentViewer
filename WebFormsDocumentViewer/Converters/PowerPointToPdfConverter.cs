using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace WebFormsDocumentViewer.Converters
{
    /// <summary>
    /// Helper class to convert PowerPoint documents to PDF
    /// </summary>
    /// <remarks>
    /// You should also have Microsoft Office installed on the server for this to work
    /// </remarks>
    public class PowerPointToPdfConverter : Infrastructure.IConverter
    {
        /// <summary>
        /// Converts the PowerPoint document to PDF
        /// </summary>
        /// <param name="filePath">Path to the PowerPoint file</param>
        /// <param name="destinationPath">Directory where the PDF file will be saved</param>
        /// <returns>Name of the converted file</returns>
        public string Convert(string filePath, string destinationPath)
        {           
            Application appPowerPoint = new Application();
            Presentation powerPointDocument = null;
            try
            {
                powerPointDocument = appPowerPoint.Presentations.Open(filePath, MsoTriState.msoFalse, 
                    MsoTriState.msoFalse, MsoTriState.msoFalse);
                powerPointDocument.Final = false;
                string fileName = Path.GetFileNameWithoutExtension(filePath) + DateTime.Now.Ticks + ".pdf";

                if (!Directory.Exists(destinationPath))
                    Directory.CreateDirectory(destinationPath);
                powerPointDocument.ExportAsFixedFormat(Path.Combine(destinationPath, fileName), PpFixedFormatType.ppFixedFormatTypePDF);

                return fileName;
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
