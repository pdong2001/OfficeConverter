using OfficeOpenXml;
using Newtonsoft.Json;
using SixLabors.ImageSharp;

namespace ExcelToJsonConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Set EPPlus license context (NonCommercial or Commercial)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            Console.WriteLine("Excel to JSON Converter");
            Console.WriteLine("=======================\n");

            if (args.Length == 0)
            {
                Console.WriteLine("Usage: ExcelToJsonConverter <excel-file-path> [output-json-path]");
                Console.WriteLine("\nExample:");
                Console.WriteLine("  ExcelToJsonConverter input.xlsx output.json");
                Console.WriteLine("  ExcelToJsonConverter input.xlsx");
                return;
            }

            string inputPath = args[0];
            string outputPath = args.Length > 1 ? args[1] : Path.ChangeExtension(inputPath, ".json");

            try
            {
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: File '{inputPath}' not found.");
                    return;
                }

                Console.WriteLine($"Reading Excel file: {inputPath}");
                var result = ConvertExcelToJson(inputPath);

                string jsonOutput = JsonConvert.SerializeObject(result, Formatting.Indented);
                File.WriteAllText(outputPath, jsonOutput);

                Console.WriteLine($"\nConversion completed successfully!");
                Console.WriteLine($"Output file: {outputPath}");
                Console.WriteLine($"Worksheets processed: {result.Worksheets.Count}");
                Console.WriteLine($"Total images extracted: {result.Worksheets.Sum(w => w.Images.Count)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
            }
        }

        static ExcelData ConvertExcelToJson(string excelFilePath)
        {
            var excelData = new ExcelData
            {
                FileName = Path.GetFileName(excelFilePath),
                Worksheets = new List<WorksheetData>()
            };

            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    Console.WriteLine($"Processing worksheet: {worksheet.Name}");
                    
                    var worksheetData = new WorksheetData
                    {
                        Name = worksheet.Name,
                        Rows = new List<List<object>>(),
                        Images = new List<ImageData>()
                    };

                    // Extract cell data
                    if (worksheet.Dimension != null)
                    {
                        int rowCount = worksheet.Dimension.End.Row;
                        int colCount = worksheet.Dimension.End.Column;

                        for (int row = 1; row <= rowCount; row++)
                        {
                            var rowData = new List<object>();
                            for (int col = 1; col <= colCount; col++)
                            {
                                var cell = worksheet.Cells[row, col];
                                rowData.Add(cell.Value ?? "");
                            }
                            worksheetData.Rows.Add(rowData);
                        }
                    }

                    // Extract images
                    if (worksheet.Drawings.Count > 0)
                    {
                        Console.WriteLine($"  Found {worksheet.Drawings.Count} drawing(s)");
                        
                        foreach (var drawing in worksheet.Drawings)
                        {
                            if (drawing is OfficeOpenXml.Drawing.ExcelPicture picture)
                            {
                                var imageData = new ImageData
                                {
                                    Name = picture.Name,
                                    Description = picture.Description ?? "",
                                    Row = picture.From.Row,
                                    Column = picture.From.Column
                                };

                                // Save image to a file
                                string imageDirectory = Path.Combine(Path.GetDirectoryName(excelFilePath) ?? "", "extracted_images");
                                Directory.CreateDirectory(imageDirectory);

                                string imageFileName = $"{worksheet.Name}_{picture.Name}.png";
                                string imagePath = Path.Combine(imageDirectory, imageFileName);

                                try
                                {
                                    // Get image bytes
                                    var imageBytes = picture.Image?.ImageBytes;
                                    
                                    if (imageBytes != null && imageBytes.Length > 0)
                                    {
                                        File.WriteAllBytes(imagePath, imageBytes);
                                        
                                        // Get dimensions from the saved image using ImageSharp
                                        using (var image = Image.Load(imageBytes))
                                        {
                                            imageData.Width = image.Width;
                                            imageData.Height = image.Height;
                                        }
                                        
                                        imageData.FilePath = imagePath;
                                        Console.WriteLine($"  Extracted image: {imageFileName} ({imageData.Width}x{imageData.Height})");
                                    }
                                    else
                                    {
                                        Console.WriteLine($"  Warning: Image '{picture.Name}' has no image bytes");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"  Warning: Could not save image '{picture.Name}': {ex.Message}");
                                }

                                worksheetData.Images.Add(imageData);
                            }
                        }
                    }

                    excelData.Worksheets.Add(worksheetData);
                }
            }

            return excelData;
        }
    }

    // Data structures for JSON output
    public class ExcelData
    {
        public string FileName { get; set; } = "";
        public List<WorksheetData> Worksheets { get; set; } = new();
    }

    public class WorksheetData
    {
        public string Name { get; set; } = "";
        public List<List<object>> Rows { get; set; } = new();
        public List<ImageData> Images { get; set; } = new();
    }

    public class ImageData
    {
        public string Name { get; set; } = "";
        public string Description { get; set; } = "";
        public int Row { get; set; }
        public int Column { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public string FilePath { get; set; } = "";
    }
}
