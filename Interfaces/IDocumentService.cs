namespace EchoTest.Interfaces
{
    public interface IDocumentService
    {
        void GenerateAndUploadGoogleDocument(string googleSheetUrl, string destination);
    }
}
