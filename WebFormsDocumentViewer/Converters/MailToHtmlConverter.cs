using Microsoft.Office.Interop.Outlook;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using WebFormsDocumentViewer.Infrastructure;

namespace WebFormsDocumentViewer.Converters
{
    /// <summary>
    /// Helper class to convert Mail files to HTML
    /// </summary>
    /// <remarks>
    /// You should also have Microsoft Office installed on the server for this to work
    /// </remarks>
    public class MailToHtmlConverter : IConverter
    {
        /// <summary>
        /// Converts the Mail file to HTML
        /// </summary>
        /// <param name="filePath">Path to the mail file</param>
        /// <param name="destinationPath">Directory where the HTML file will be saved</param>
        /// <returns>Name of the converted file</returns>
        public string Convert(string filePath, string destinationPath)
        {
            Process process = new Process();
            Application appOutlook = null;
            MailItem mailItem = null;
            try
            {
                if (Path.GetExtension(filePath) == "." + SupportedExtensions.eml.ToString())
                {
                    foreach (var oldProcess in Process.GetProcessesByName("OUTLOOK"))
                    {
                        oldProcess.Kill();
                    }

                    process.StartInfo.FileName = filePath;
                    process.Start();

                    bool outlookOpen = false;
                    while (!outlookOpen)
                    {
                        try
                        {
                            appOutlook = (Application)Marshal.GetActiveObject("Outlook.Application");
                            outlookOpen = true;
                        }
                        catch
                        {
                            Thread.Sleep(100);
                        }
                    }
                    mailItem = (MailItem)appOutlook.ActiveInspector().CurrentItem;
                }
                else
                {
                    appOutlook = new Application();
                    mailItem = (MailItem)appOutlook.Session.OpenSharedItem(filePath);
                }
                string fileName = Path.GetFileNameWithoutExtension(filePath) + DateTime.Now.Ticks + ".html";

                if (!Directory.Exists(destinationPath))
                    Directory.CreateDirectory(destinationPath);
                mailItem.SaveAs(Path.Combine(destinationPath, fileName), OlSaveAsType.olHTML);

                return fileName;
            }
            catch
            {
                return null;
            }
            finally
            {
                mailItem?.Close(OlInspectorClose.olDiscard);
                appOutlook?.Quit();
                process.Dispose();
            }
        }
    }
}
