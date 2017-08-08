namespace WebFormsDocumentViewer.Infrastructure
{
    /// <summary>
    /// Defines methods to be implemented in converters
    /// </summary>
    public interface IConverter
    {
        string Convert(string filePath, string destinationPath, string projectRootPath);
    }
}
