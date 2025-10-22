# Excel to JSON Converter

A C# console application that converts Excel files (.xlsx) to JSON format. The application supports multiple worksheets, extracts embedded images, and includes comprehensive image metadata in the output.

## Features

- **Multiple Worksheet Support**: Converts all worksheets in an Excel file
- **Image Extraction**: Automatically extracts and saves images embedded in worksheets
- **Image Metadata**: Includes image names, dimensions, positions, and file paths in JSON output
- **Cross-Platform**: Uses SixLabors.ImageSharp for cross-platform image handling
- **Simple CLI**: Easy-to-use command-line interface

## Prerequisites

- .NET 9.0 SDK or later
- EPPlus library (automatically installed via NuGet)
- Newtonsoft.Json library (automatically installed via NuGet)
- SixLabors.ImageSharp library (automatically installed via NuGet)

## Installation

1. Clone the repository:
```bash
git clone https://github.com/pdong2001/OfficeConverter.git
cd OfficeConverter/ExcelToJsonConverter
```

2. Restore dependencies:
```bash
dotnet restore
```

3. Build the application:
```bash
dotnet build
```

## Usage

### Basic Usage

Convert an Excel file to JSON:
```bash
dotnet run -- <excel-file-path> [output-json-path]
```

### Examples

Convert with automatic output filename:
```bash
dotnet run -- input.xlsx
# Creates input.json in the same directory
```

Convert with custom output filename:
```bash
dotnet run -- data/products.xlsx output/products.json
```

### Output Structure

The generated JSON has the following structure:

```json
{
  "FileName": "example.xlsx",
  "Worksheets": [
    {
      "Name": "Sheet1",
      "Rows": [
        ["Header1", "Header2", "Header3"],
        ["Value1", "Value2", "Value3"]
      ],
      "Images": [
        {
          "Name": "Image 1",
          "Description": "Picture",
          "Row": 1,
          "Column": 4,
          "Width": 100,
          "Height": 100,
          "FilePath": "extracted_images/Sheet1_Image 1.png"
        }
      ]
    }
  ]
}
```

### Extracted Images

Images are automatically extracted and saved to a folder named `extracted_images` in the same directory as the input Excel file. The naming convention for extracted images is:

```
{WorksheetName}_{ImageName}.png
```

## JSON Output Format

### Root Object
- `FileName` (string): The name of the input Excel file
- `Worksheets` (array): Array of worksheet objects

### Worksheet Object
- `Name` (string): Name of the worksheet
- `Rows` (array): 2D array containing cell values
- `Images` (array): Array of image objects found in the worksheet

### Image Object
- `Name` (string): Name of the image in Excel
- `Description` (string): Image description/alt text
- `Row` (integer): Row index where the image starts (0-based)
- `Column` (integer): Column index where the image starts (0-based)
- `Width` (integer): Image width in pixels
- `Height` (integer): Image height in pixels
- `FilePath` (string): Path to the extracted image file

## License

This project uses EPPlus library which requires a commercial license for commercial use. For non-commercial use, it's free under the Polyform Noncommercial license.

## Dependencies

- **EPPlus**: Excel file manipulation library
- **Newtonsoft.Json**: JSON serialization
- **SixLabors.ImageSharp**: Cross-platform image processing

## Troubleshooting

### Common Issues

1. **File not found error**: Ensure the Excel file path is correct and the file exists
2. **Permission denied**: Make sure you have read permissions for the input file and write permissions for the output directory
3. **Invalid Excel file**: The application only supports .xlsx files (Office Open XML format)

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Authors

- Created for the OfficeConverter project
