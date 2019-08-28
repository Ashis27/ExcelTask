using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.AspNetCore.Hosting;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;

namespace UploadExcelTask.Controllers
{
    [Route("api/[Controller]")]
    [ApiController]
    public class UploadExcelController : ControllerBase
    {
        [HttpPost("UpdateExcelSheet")]
        public async Task<IActionResult> UploadExcelFile(IFormFile fileInfo)
        {
            try
            {
                if (fileInfo == null || fileInfo.Length == 0)
                    return BadRequest("File not found");
                Dictionary<string, Dictionary<string, string>> rowData = new Dictionary<string, Dictionary<string, string>>();
                Dictionary<string, string> colData = new Dictionary<string, string>();
                var filePath = Path.GetTempFileName();
                var memory = new MemoryStream();
                await fileInfo.CopyToAsync(memory);
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(memory, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                    OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);

                    foreach (Row rw in sheetData.Elements<Row>())
                    {
                        foreach (Cell cl in rw.Elements<Cell>())
                        {
                            if (int.Parse(cl.InnerText) >= 0)
                            {
                                int stringId = int.Parse(cl.InnerText);
                                colData.Add(cl.CellReference, workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(stringId).InnerText);
                                //colData.Add(cl.CellReference, cl.InnerText);
                            }
                        }
                        rowData.Add(Guid.NewGuid().ToString(), colData);
                        colData = new Dictionary<string, string>();
                    }
                }
                //Insert new data in newly added sheet
                CreateNewSheetAndInsertData(memory, rowData);
                memory.Position = 0;
                var path = Path.Combine(Directory.GetCurrentDirectory(), $"wwwroot/upload", fileInfo.FileName);
                using (var stream = new FileStream(path, FileMode.Create))
                {
                    await memory.CopyToAsync(stream);
                }
                return Ok();
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        private string GetContentType(string fileName)
        {
            var types = GetMimeTypes();
            var ext = Path.GetExtension(fileName).ToLowerInvariant();
            return types[ext];
        }
        public static void CreateNewSheetAndInsertData(Stream docName, Dictionary<string, Dictionary<string, string>> rowData)
        {
            // Open the document for editing.
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
            {
                // Get the SharedStringTablePart. If it does not exist, create a new one.
                SharedStringTablePart shareStringPart;
                if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }
                // Insert the text into the SharedStringTablePart.
                // Insert a new worksheet.
                WorksheetPart worksheetPart = InsertNewWorksheet(spreadSheet.WorkbookPart);
                // Insert cell A1 into the new worksheet.
                Row r = new Row();
                uint count = 0;
                foreach (var rw in rowData)
                {
                    count++;
                    foreach (var val in rw.Value)
                    {
                        Cell cell = InsertCellInWorksheet(val.Key.Remove(val.Key.Length - 1, 1), count, worksheetPart);
                        // Set the value of cell A1.
                        cell.CellValue = new CellValue(val.Value + "_" + val.Key);
                        //var test1 = cell.CellValue;
                        //cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                        //var test = cell.DataType;
                    }
                }
                worksheetPart.Worksheet.Save();
            }
        }
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;
            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }
            // If there is not a cell with the specified column name, insert one.
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellReference.Value.Length == cellReference.Length)
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }
                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);
                worksheet.Save();
                return newCell;
            }
        }
        private static WorksheetPart InsertNewWorksheet(WorkbookPart workbookPart)
        {
            // Add a blank WorksheetPart.
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();
            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new worksheet.
            uint sheetId = 0;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            // Give the new worksheet a name.
            string sheetName = "Sheet" + sheetId;

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();

            return newWorksheetPart;
        }
        private Dictionary<string, string> GetMimeTypes()
        {
            return new Dictionary<string, string>
            {
                {".txt", "text/plain"},
                {".pdf", "application/pdf"},
                {".doc", "application/vnd.ms-word"},
                {".docx", "application/vnd.ms-word"},
                {".xls", "application/vnd.ms-excel"},
                {".xlsx", "application/vnd.openxmlformatsofficedocument.spreadsheetml.sheet"},
                {".png", "image/png"},
                {".jpg", "image/jpeg"},
                {".jpeg", "image/jpeg"},
                {".gif", "image/gif"},
                {".csv", "text/csv"},
                {".zip", "application/zip"}
            };
        }

        [HttpGet("DownloadUpdatedFile")]
        public async Task<IActionResult> DownloadFile(string fileName)
        {
            if (String.IsNullOrEmpty(fileName))
                return BadRequest("File name can not be empty");
            var path = Path.Combine(Directory.GetCurrentDirectory(), $"wwwroot/upload", fileName);
            
            var memory = new MemoryStream();
            using (var stream = new FileStream(path, FileMode.Open))
            {
                await stream.CopyToAsync(memory);
            }
            if (memory == null)
                return NotFound();
            memory.Position = 0;
            FileStreamResult file = File(memory, GetContentType(fileName), fileName);
            return file;
            // return file;
        }
    }
}