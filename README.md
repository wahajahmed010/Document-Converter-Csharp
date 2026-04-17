# Document Converter

A Windows Forms desktop application for converting between popular document formats: PDF, Word (DOC/DOCX), and Text.

## Features

- **PDF to Word** — Convert PDF documents to editable DOCX format
- **Word to PDF** — Export Word documents to PDF format
- **Word to Text** — Extract plain text from Word documents
- **Async Operations** — Non-blocking conversions with progress reporting
- **Cancellation Support** — Cancel ongoing conversions
- **Structured Logging** — Detailed operation logs
- **Global Exception Handling** — Graceful error management

## Supported Formats

| Input | Output | Status |
|-------|--------|--------|
| PDF | DOCX | ✅ |
| DOC/DOCX | PDF | ✅ |
| DOC/DOCX | TXT | ✅ |

## Requirements

- Windows 10/11
- .NET 8.0 Runtime
- Microsoft Office (for Word interop conversions)
- Visual Studio 2022 (for building from source)

## Installation

### From Source

1. Clone the repository:
   ```bash
   git clone https://github.com/wahajahmed010/Document-Converter-Csharp.git
   ```

2. Open `Document converter.sln` in Visual Studio 2022

3. Restore NuGet packages:
   ```bash
   dotnet restore
   ```

4. Build and run:
   ```bash
   dotnet build
   dotnet run
   ```

### Dependencies (NuGet)

- `SautinSoft.PdfFocus` — PDF conversion engine
- `Microsoft.Office.Interop.Word` — Word document automation

## Usage

1. Launch the application
2. Select input file using the Browse button
3. Choose output format from the dropdown
4. Select save location
5. Click Convert
6. Monitor progress in the status bar
7. Cancel anytime with the Cancel button

## Project Structure

```
Document converter/
├── Form1.cs              # Main application form
├── Conversion.cs         # Central conversion dispatcher
├── DocToPdf.cs           # Word to PDF converter
├── DocToTxt.cs           # Word to Text converter
├── PdfToWord.cs          # PDF to Word converter
├── PdfToTxt.cs           # PDF to Text converter
├── Logger.cs             # Structured logging
└── Program.cs            # Application entry point
```

## Architecture

The application follows a modular converter pattern:

- **Central Dispatcher** — `Conversion.cs` routes requests to appropriate converters
- **Converter Classes** — Each implements `IDisposable` for proper resource cleanup
- **Async Pattern** — All conversions run on background threads
- **Progress Reporting** — Real-time progress updates via `IProgress<T>`
- **Cancellation** — `CancellationToken` support for user-initiated cancellation

## Recent Enhancements (2026-04-17)

This project was enhanced using the Hermes Adaptation auto-skill pipeline:

### Critical Fixes
- ✅ Fixed COM resource leaks (Word process cleanup)
- ✅ Implemented `IDisposable` pattern on all converters
- ✅ Added comprehensive input validation
- ✅ Fixed `PdfToTxt` to accept explicit save paths
- ✅ Migrated from local DLL to NuGet packages

### Production Hardening
- ✅ Added structured logging with file output
- ✅ Implemented global exception handlers
- ✅ Converted to async/await pattern
- ✅ Added cancellation support
- ✅ Added progress reporting
- ✅ Upgraded from .NET 4.0 to .NET 8.0
- ✅ Removed dead code and commented blocks
- ✅ Cleaned up naming conventions

## Logging

Logs are written to:
```
%LOCALAPPDATA%\DocumentConverter\Logs\app_YYYYMMDD.log
```

## Error Handling

The application includes:
- Global unhandled exception catchers
- Input validation (null/empty checks, file existence)
- COM object cleanup in finally blocks
- Graceful degradation on conversion failures

## Development

### Building

```bash
dotnet build --configuration Release
```

### Testing

> ⚠️ This project requires Windows and Microsoft Office for full testing. Cross-platform testing is limited.

## Limitations

- Windows-only (requires Microsoft Office interop)
- Requires Office installation for Word conversions
- Large files may take significant time
- OCR not supported for image-based PDFs

## License

Student project — created while learning C# and .NET

## Credits

Enhanced using the [OpenClaw Hermes Adaptation](https://github.com/wahajahmed010/openclaw-hermes-adaptation) auto-skill pipeline.

Council: Mnemosyne, Sibyl, Fabricator  
Enhancement Agents: Enhancer-Critical, Enhancer-High
