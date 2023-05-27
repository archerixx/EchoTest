using Domain.Dto;

namespace Api.Interfaces
{
    public interface IDocumentService
    {
        string GenerateAndUploadGoogleDocument(string googleSheetUrl, string? shareWith);
        IEnumerable<SpreadSheetDto> GetGoogleSheet();
        void LogoutFromGoogleAPI();
    }
}
