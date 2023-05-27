using Google.Apis.Sheets.v4.Data;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.Util;
using NPOI.XWPF.UserModel;
using System.Drawing;
using Word = Microsoft.Office.Interop.Word;
using Api.Models;
using Sheet = Google.Apis.Sheets.v4.Data.Sheet;

namespace Api.HelperMethods
{
    public class GenerateWordHelper
    {
        public static MemoryStream GenerateWordDocumentFromSpreadsheet(Spreadsheet spreadsheet, string? filePath = null)
        {
            // Since Image does not exist inside Excel, has to be either picked up from internet from URL or from local drive
            var imageData = GetImageData("CompanyName.png");
            
            MemoryStream mainStream = new MemoryStream();
            using (MemoryStream sw = new MemoryStream())
            //using (FileStream sw = File.Create(filePath))
            {
                // initialize document
                XWPFDocument doc = new();

                // first paragraph, place picture
                AddPictureToDocument(ref doc, imageData);

                // each sheet is one page
                for (var i = 0; i < spreadsheet.Sheets.Count; i++)
                {
                    // extract sheet and data from spreadsheet
                    Sheet sheet = spreadsheet.Sheets[i];
                    var sheetData = sheet.Data[0].RowData;

                    switch (sheet.Properties.Title)
                    {
                        case "Cover Page":
                            // create dictonary from existing data
                            Dictionary<string, string> coverPageDictionary = sheetData
                                .ToDictionary(
                                    r => r.Values[0].FormattedValue,
                                    r => r.Values[1].FormattedValue
                                );

                            AddTitle(ref doc, coverPageDictionary.GetValueOrDefault("Title"));
                            // discard title value from dictionary
                            coverPageDictionary.Remove("Title");
                            AddAuthorInfo(ref doc, "Zerin", coverPageDictionary);

                            // after this, Table of contents is set
                            var paragraph = doc.CreateParagraph();
                            paragraph.IsPageBreak = true;
                            doc.CreateTOC();
                            break;
                        case "Abstract":
                            // toc header
                            AddTOCHeader(ref doc, sheet.Properties.Title);
                            // only second column (Values[1]) is populated
                            var abstartRowData = sheetData.Where(r => r.Values is not null);

                            AddTable(ref doc, abstartRowData);
                            break;
                        case "Classification Codes":
                            // header is not same as title is populated
                            var classificationCodesRowData = sheetData.Where(r => r.Values is not null);
                            var classificationCodesHeaderRow = classificationCodesRowData
                                .First(rowData => rowData.Values.FirstOrDefault(v => v.EffectiveFormat != null)?.EffectiveFormat.Borders == null);

                            // toc header
                            AddTOCHeader(ref doc, classificationCodesHeaderRow.Values.First(v => v.FormattedValue != null).FormattedValue);

                            AddTable(ref doc, classificationCodesRowData);
                            break;
                        case "Claims":
                            // toc header
                            AddTOCHeader(ref doc, sheet.Properties.Title);
                            // get data
                            var claimsRowData = sheetData.Where(r => r.Values is not null);
                            // create two-dimensional array
                            var columnsData = claimsRowData.Select(r => r.Values.Where(v => v.FormattedValue is not null).ToArray()).Skip(1).ToArray();

                            AddBullets(ref doc, columnsData);
                            break;
                        case "Application Events":
                            // header is not same as title is populated
                            var applicationEventsRowData = sheetData.Where(r => r.Values is not null);
                            var applicationEventsHeader = applicationEventsRowData
                                .First(rowData => rowData.Values.FirstOrDefault(v => v.EffectiveFormat != null)?.EffectiveFormat.Borders == null);

                            // toc header
                            AddTOCHeader(ref doc, applicationEventsHeader.Values.First(v => v.FormattedValue != null).FormattedValue);

                            AddTable(ref doc, applicationEventsRowData);
                            break;
                    }
                }

                doc.Write(sw);
                mainStream = new MemoryStream(sw.ToArray());
            }

            return mainStream;
        }

        private static void AddPictureToDocument(ref XWPFDocument document, ImageData imageData)
        {
            var imageLine = document.CreateParagraph().CreateRun();
            var heigh = Units.PixelToEMU(imageData.Height);
            var width = Units.PixelToEMU(imageData.Width);
            imageLine.AddPicture(imageData.Stream, (int)imageData.Type, imageData.Name, width, heigh);
        }

        private static void AddTitle(ref XWPFDocument document, string? titleText)
        {
            XWPFParagraph paragraph = document.CreateParagraph();
            paragraph.Alignment = ParagraphAlignment.CENTER;
            // two line breaks between image and title
            var titleLineBreaks = paragraph.CreateRun();
            titleLineBreaks.FontSize = 10;
            titleLineBreaks.AddBreak(BreakType.TEXTWRAPPING);
            titleLineBreaks.AddBreak(BreakType.TEXTWRAPPING);
            // title
            var title = paragraph.CreateRun();
            title.FontFamily = "Arial";
            title.FontSize = 26;
            title.AppendText(titleText);
        }

        private static void AddAuthorInfo(ref XWPFDocument document, string preparedBy, Dictionary<string, string> coverPageDictionary)
        {
            XWPFParagraph paragraph = document.CreateParagraph();
            paragraph.IsPageBreak = false;
            paragraph.Alignment = ParagraphAlignment.LEFT;
            var lineBreaks = paragraph.CreateRun();
            // create 11 breaks to space out AuthorInfo
            for (int i = 0; i <= 10; i++)
            {
                lineBreaks.AddBreak(BreakType.TEXTWRAPPING);
            }

            coverPageDictionary.TryAdd("Prepared by", preparedBy);
            var authorInfoFields = new List<string> { "Prepared by", "Inventor", "Patent No" };
            for (int field = 0; field < coverPageDictionary.Count; field++)
            {
                var fieldTypeName = paragraph.CreateRun();
                fieldTypeName.FontFamily = "Arial";
                fieldTypeName.FontSize = 14;
                fieldTypeName.SetText($"{authorInfoFields[field]}: ");
                var fieldTypeValue = paragraph.CreateRun();
                fieldTypeValue.IsBold = true;
                fieldTypeValue.FontSize = 14;
                fieldTypeValue.AppendText(coverPageDictionary.GetValueOrDefault(authorInfoFields[field]));
                fieldTypeValue.AddBreak(BreakType.TEXTWRAPPING);
            }
        }

        private static void AddTOCHeader(ref XWPFDocument document, string header)
        {
            // TOC header
            var paragraph = document.CreateParagraph();
            paragraph.IsPageBreak = true;
            paragraph.Style = "Heading1";
            // header value
            var abstractHeader = paragraph.CreateRun();
            abstractHeader.FontSize = 14;
            abstractHeader.IsBold = true;
            abstractHeader.SetText(header);
            abstractHeader.FontFamily = "Arial";
            abstractHeader.AddBreak(BreakType.TEXTWRAPPING);
        }

        private static void AddTable(ref XWPFDocument document, IEnumerable<RowData> rowData)
        {
            List<RowData> tableRowDataList = rowData
                .Where(rowData => rowData.Values != null && rowData.Values.Count > 0
                    && rowData.Values.FirstOrDefault(v => v.EffectiveFormat != null)?.EffectiveFormat.Borders != null)
                .ToList();

            // extract colums from rowData
            var columns = tableRowDataList.First().Values.Where(v => v.FormattedValue is not null).Count();
            var columnsData = tableRowDataList.Select(r => r.Values.Where(v => v.FormattedValue is not null).ToArray()).ToArray();

            // table
            var table = document.CreateTable(tableRowDataList.Count(), columns);
            table.SetCellMargins(100, 100, 100, 100);
            
            // this part does not seem to be recognized by Google docs
            var tblLayout = table.GetCTTbl().tblPr.AddNewTblLayout();
            tblLayout.type = ST_TblLayoutType.@fixed;
            tblLayout.typeSpecified = true;

            for (int i = 0; i < tableRowDataList.Count; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    // tableLayout no recognized above, so column width has to be set 9600 is approx 16,8 cm (as fixed size is)
                    table.SetColumnWidth(j, (ulong)(9600 / columns));
                    var dataCell = table.Rows[i].GetCell(j);
                    dataCell.GetParagraphArray(0).SpacingAfterLines = 1;
                    dataCell.SetColor(GetHexString(columnsData[i][j].EffectiveFormat.BackgroundColorStyle.RgbColor));

                    // VerticalAlignemnt, but does not match output doc format
                    //dataCell.SetVerticalAlignment((XWPFTableCell.XWPFVertAlign)Enum.Parse(typeof(XWPFTableCell.XWPFVertAlign), columnsData[i][j].EffectiveFormat.VerticalAlignment));

                    dataCell.SetBorderTop(XWPFTable.XWPFBorderType.SINGLE, 10, 0, GetHexString(columnsData[i][j].EffectiveFormat.Borders.Top.ColorStyle.RgbColor));
                    dataCell.SetBorderLeft(XWPFTable.XWPFBorderType.SINGLE, 10, 0, GetHexString(columnsData[i][j].EffectiveFormat.Borders.Left.ColorStyle.RgbColor));
                    dataCell.SetBorderBottom(XWPFTable.XWPFBorderType.SINGLE, 10, 0, GetHexString(columnsData[i][j].EffectiveFormat.Borders.Bottom.ColorStyle.RgbColor));
                    dataCell.SetBorderRight(XWPFTable.XWPFBorderType.SINGLE, 10, 0, GetHexString(columnsData[i][j].EffectiveFormat.Borders.Right.ColorStyle.RgbColor));

                    // if text is hyperlink
                    if (columnsData[i][j].Hyperlink != null)
                    {
                        XWPFHyperlinkRun hyperlinkrun = CreateHyperlinkRun(dataCell.GetParagraphArray(0), columnsData[i][j].Hyperlink);
                        hyperlinkrun.SetText(columnsData[i][j].FormattedValue);
                        hyperlinkrun.SetColor(GetHexString(columnsData[i][j].EffectiveFormat.TextFormat.ForegroundColorStyle.RgbColor));
                        if (columnsData[i][j].EffectiveFormat.TextFormat.Underline.Value)
                            hyperlinkrun.Underline = UnderlinePatterns.Single;
                        hyperlinkrun.IsBold = columnsData[i][j].EffectiveFormat.TextFormat.Bold.Value;
                        hyperlinkrun.FontFamily = columnsData[i][j].EffectiveFormat.TextFormat.FontFamily;
                        hyperlinkrun.FontSize = columnsData[i][j].EffectiveFormat.TextFormat.FontSize.Value;
                    }
                    else
                    {
                        var dataCellText = dataCell.GetParagraphArray(0).CreateRun();
                        dataCellText.SetText(columnsData[i][j].FormattedValue);
                        dataCellText.IsBold = columnsData[i][j].EffectiveFormat.TextFormat.Bold.Value;
                        dataCellText.FontFamily = columnsData[i][j].EffectiveFormat.TextFormat.FontFamily;
                        dataCellText.FontSize = columnsData[i][j].EffectiveFormat.TextFormat.FontSize.Value;
                    }
                }
            }
        }

        private static void AddBullets(ref XWPFDocument document, CellData[][] cellData)
        {
            XWPFNumbering numbering = document.CreateNumbering();
            var mltTest = new CT_MultiLevelType()
            {
                val = ST_MultiLevelType.multilevel
            };
            var ct_abnTest = new CT_AbstractNum()
            {
                multiLevelType = mltTest,
                lvl = new List<CT_Lvl>()
            };

            ulong ind = 360;
            int leftSize = 360;
            int levlNo = 1;
            string levl = "";
            // create bullet style and levels, max 8 possible
            for (int i = 0; i < 8; i++)
            {
                levl += $"%{levlNo}.";

                var nbmLvl = new CT_Lvl
                {
                    ilvl = $"{i}",
                    start = new CT_DecimalNumber() { val = $"1" },
                    numFmt = new CT_NumFmt() { val = ST_NumberFormat.@decimal },
                    lvlText = new CT_LevelText() { val = levl },
                    lvlJc = new CT_Jc() { val = ST_Jc.left },
                    pPr = new CT_PPr { ind = new CT_Ind { left = $"{leftSize}", hanging = ind } }
                };

                levlNo++;
                ind += 144; // indentation
                leftSize += 504; // left padding
                ct_abnTest.lvl.Add(nbmLvl);
            }
            string abstractNumId = numbering.AddAbstractNum(new XWPFAbstractNum(ct_abnTest));
            string numId = numbering.AddNum(abstractNumId);

            var paragraph = document.CreateParagraph();
            var paragraphText = paragraph.CreateRun();

            for (int i = 0; i < cellData.Count(); i++)
            {
                paragraph = document.CreateParagraph();
                paragraphText = paragraph.CreateRun();
                paragraphText.FontFamily = cellData[i][1].EffectiveFormat.TextFormat.FontFamily;
                paragraphText.FontSize = cellData[i][1].EffectiveFormat.TextFormat.FontSize.Value;
                paragraphText.IsBold = cellData[i][1].EffectiveFormat.TextFormat.Bold.Value;
                paragraphText.SetText(cellData[i][1].FormattedValue);
                paragraph.SetNumID(numId, $"{cellData[i][0].FormattedValue.Split(".").Count() - 1}");
            }
        }

        private static XWPFHyperlinkRun CreateHyperlinkRun(XWPFParagraph paragraph, String uri)
        {
            String rId = paragraph.Document.GetPackagePart().AddExternalRelationship(
              uri,
              XWPFRelation.HYPERLINK.Relation
             ).Id;

            return paragraph.CreateHyperlinkRun(rId);
        }

        private static string GetHexString(Google.Apis.Sheets.v4.Data.Color color)
        {
            if (color.Red is null || color.Green is null || color.Blue is null)
                return string.Empty;

            return string.Format("#{0}{1}{2}",
                ((int)(color.Red * 255)).ToString("X2"),
                ((int)(color.Green * 255)).ToString("X2"),
                ((int)(color.Blue * 255)).ToString("X2"));
        }

        // Since Image does not exist inside Excel, has to be either picked up from internet from URL or from local drive
        private static ImageData GetImageData(string imagePath)
        {
            // get Image info
            var imageName = Path.GetFileName(imagePath);
            var fileExtension = imageName.Split(".")[1].ToUpper();
            var imageType = (PictureType)Enum.Parse(typeof(PictureType), fileExtension);
            Image image = Image.FromFile(imagePath);

            return new ImageData()
            {
                Name = imageName,
                Stream = File.OpenRead(imagePath) as Stream,
                Type = imageType,
                Width = image.Width,
                Height = image.Height
            };
        }

        public static void UpdateFieldsInWordDocument(string filePath)
        {
            // Create an instance of Word application
            var wordApp = new Word.Application();

            // Disable alerts
            wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

            // Open the Word document
            var wordDoc = wordApp.Documents.Open(filePath);

            // Update fields
            wordDoc.Fields.Update();

            // Save and close the document
            wordDoc.Save();
            wordDoc.Close();

            // Enable alerts
            wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsAll;

            // Quit Word application
            wordApp.Quit();
        }
    }
}
