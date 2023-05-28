using Domain.Dto;
using Web.HelperMethods;
using Web.Interfaces;
using Google.Apis.Sheets.v4.Data;

namespace Web.Services
{
    public class DocumentService : IDocumentService
    {
        public string GenerateAndUploadGoogleDocument(string spreadsheetId, string? shareWith)
        {
            Spreadsheet spreadsheet = GoogleApiHelper.GetSpreadsheet(spreadsheetId);

            #region soltion#2 - save file then upload to google docs
            //var currentPath = Directory.GetCurrentDirectory();
            //var filePath = Path.Combine(currentPath, "Output.docx");
            //// following line comes after GenerateWordDocumentFromSpreadsheet
            //// update fields (TOC, Tables) - works only on machines with Office installed
            //GenerateWordHelper.UpdateFieldsInWordDocument(filePath);
            #endregion

            // generate document stream
            var stream = GenerateWordHelper.GenerateWordDocumentFromSpreadsheet(spreadsheet);
            // upload to GoogleDocs
            return GoogleApiHelper.UploadToGoogleDocs(stream, spreadsheet.Properties.Title, shareWith);
        }

        public IEnumerable<SpreadSheetDto> GetGoogleSheet()
        {
            var sheetFiles = GoogleApiHelper.GetSpreadsheet();
            return sheetFiles.Select(sf => new SpreadSheetDto { 
                Id = sf.Id,
                Name = sf.Name
            });
        }

        public void LogoutFromGoogleAPI()
        {
            GoogleApiHelper.Logout();
        }
    }
}
