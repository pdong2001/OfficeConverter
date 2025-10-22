# OfficeConverter

A collection of office document conversion tools.

## Excel to JSON Converter

A C# console application that converts Excel files to JSON format with support for multiple worksheets and embedded images.

### Quick Start

```bash
cd ExcelToJsonConverter
dotnet run -- input.xlsx output.json
```

See [ExcelToJsonConverter/README.md](ExcelToJsonConverter/README.md) for detailed documentation.

## Features

- Convert Excel files (.xlsx) to JSON
- Support for multiple worksheets
- Automatic image extraction with metadata
- Cross-platform compatibility
- Simple command-line interface

## Requirements

- .NET 9.0 SDK or later

## License

This project uses EPPlus library which requires a commercial license for commercial use. For non-commercial use, it's free under the Polyform Noncommercial license.
