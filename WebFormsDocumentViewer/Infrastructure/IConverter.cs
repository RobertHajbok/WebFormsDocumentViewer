namespace WebFormsDocumentViewer.Infrastructure
{
    /// <summary>
    /// Defines methods to be implemented in converters
    /// </summary>
    internal interface IConverter
    {
        string Convert(string filePath, string destinationPath);
    }
}
