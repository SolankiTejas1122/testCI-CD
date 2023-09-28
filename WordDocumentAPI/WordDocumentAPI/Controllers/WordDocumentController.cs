using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Microsoft.AspNetCore.Mvc;

namespace WordDocumentAPI.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class WordDocumentController : ControllerBase
    {
        [HttpGet]
        public IActionResult GenerateWordDocument()
        {
            try
            {
                string fileName = "DocumentWithTable.docx";

                using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document))
                {
                    // Add a main document part
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                    // Create the document structure
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());

                    // Create a table
                    Table table = new Table();

                    // Define table properties
                    TableProperties tableProperties = new TableProperties(
                        new TableBorders(
                            new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                            new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                            new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                            new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                            new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                            new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 }
                        )
                    );

                    table.AppendChild(tableProperties);

                    // Define table columns
                    TableGrid tableGrid = new TableGrid(
                        new GridColumn { Width = new GridColumnWidth { Width = "2000" } },  // Adjust column widths as needed
                        new GridColumn { Width = new GridColumnWidth { Width = "2000" } },
                        new GridColumn { Width = new GridColumnWidth { Width = "2000" } }
                        // Add more columns as needed
                    );

                    table.AppendChild(tableGrid);

                    // Add table header row
                    TableRow headerRow = new TableRow();
                    headerRow.AppendChild(CreateTableCell("Header 1"));
                    headerRow.AppendChild(CreateTableCell("Header 2"));
                    headerRow.AppendChild(CreateTableCell("Header 3"));
                    // Add more header cells as needed
                    table.AppendChild(headerRow);
                    // Add the table to the document
                    body.AppendChild(table);
                    // Save the document
                    wordDocument.Save();
                }

                string filePath = Path.Combine(Directory.GetCurrentDirectory(), fileName);

                // Return the file as a download response
                var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                return File(fileStream, "application/octet-stream", fileName);
            }
            catch (Exception ex)
            {
                return BadRequest($"Error: {ex.Message}");
            }
        }

        private TableCell CreateTableCell(string text)
        {
            TableCell cell = new TableCell();
            Paragraph paragraph = new Paragraph();
            Run run = new Run();
            Text cellText = new Text(text);
            run.Append(cellText);
            paragraph.Append(run);
            cell.Append(paragraph);
            return cell;
        }
    }
}
