using Domain.Dto;

namespace Web.Interfaces
{
    public interface IDocumentService
    {
        string GenerateAndUploadGoogleDocument(string googleSheetUrl, string? shareWith);
        IEnumerable<SpreadSheetDto> GetGoogleSheet();
        void LogoutFromGoogleAPI();
    }
}
