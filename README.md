# MSG File Parser

A .NET 8.0 console application that extracts content from Microsoft Outlook MSG files and converts them to readable text format.

## Features

- 📧 Extract email metadata (subject, sender, recipients, date)
- 📝 Support for multiple body formats (Plain Text, HTML, RTF)
- 🎯 Flexible output options (specific filename or directory)
- ⚡ Fast and efficient processing
- 🔧 Cross-platform compatibility (Windows, Linux, macOS)
- 🛡️ Comprehensive input validation and error handling
- ⚠️ Smart file format detection and warnings
- 💾 Large file size warnings and memory protection
- 🔒 File permission checking and access validation

## Prerequisites

- [.NET 8.0 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)
- For Linux users: `libgdiplus` library (usually pre-installed)

## Installation of .NET SDK

```bask
wget https://dot.net/v1/dotnet-install.sh -O dotnet-install.sh

chmod +x ./dotnet-install.sh

./dotnet-install.sh --version latest
```

## Installation

1. Clone or download this repository
2. Navigate to the project directory
3. Build the project:
   ```bash
   dotnet build
   ```

## Usage

The application supports three different usage patterns:

### 1. Specific Output Filename
```bash
dotnet run "input.msg" "output.txt"
```
Creates a file with your exact specified name and location.

### 2. Directory Output
```bash
dotnet run "input.msg" "./output/"
dotnet run "input.msg" "."
```
Auto-generates filename based on MSG file name and places it in the specified directory.

### 3. Default Behavior
```bash
dotnet run "input.msg"
```
Auto-generates filename in the same directory as the MSG file.

## Examples

```bash
# Extract to custom filename
dotnet run "email.msg" "extracted_email.txt"

# Extract to current directory with auto-generated name
dotnet run "email.msg" "."

# Extract to same directory as MSG file
dotnet run "email.msg"

# Extract to specific directory
dotnet run "email.msg" "/path/to/output/"
```

## Output Format

The extracted text file contains:

- **Subject**: Email subject line
- **From**: Sender information
- **Date**: Email timestamp (YYYY-MM-DD HH:mm:ss format)
- **Recipients**: List of all recipients with email addresses and display names
- **Body**: Email content (preserves formatting from original)

### Sample Output
```
Subject: Meeting Tomorrow
From: john.doe@example.com
Date: 2023-04-11 11:55:37

Recipients:
- jane.smith@example.com (Jane Smith)
- team@example.com (Team)

Body:
Hello Team,

Please join us for the meeting tomorrow at 2 PM.

Best regards,
John
```

## Dependencies

This project uses the following NuGet packages:

- [MSGReader](https://www.nuget.org/packages/MSGReader/) (4.3.0) - For parsing MSG files
- [System.Text.Encoding.CodePages](https://www.nuget.org/packages/System.Text.Encoding.CodePages/) (8.0.0) - For proper text encoding support

## Technical Details

- **Target Framework**: .NET 8.0
- **Platform**: Cross-platform (Windows, Linux, macOS)
- **Architecture**: Console Application
- **Language**: C#

## Error Handling & Validation

The application includes comprehensive error handling and validation for:

### Input Validation
- ✅ **File existence checking** - Verifies MSG file exists before processing
- ✅ **File extension validation** - Warns if file doesn't have .msg extension
- ✅ **File access permissions** - Checks read permissions before processing
- ✅ **Large file warnings** - Alerts for files over 50MB that may take longer
- ✅ **Empty/invalid path detection** - Prevents processing of invalid file paths

### Output Validation
- ✅ **Output directory verification** - Ensures target directory exists
- ✅ **Write permission checking** - Verifies write access to output location
- ✅ **File overwrite warnings** - Alerts when output file already exists
- ✅ **Path validation** - Validates output path format and accessibility

### MSG File Processing
- ✅ **Invalid file format detection** - Identifies non-MSG or corrupted files
- ✅ **Memory exhaustion protection** - Handles extremely large files gracefully
- ✅ **Missing data handling** - Safely processes emails with missing components
- ✅ **Encoding error recovery** - Handles text encoding issues gracefully

### Error Messages
The application provides clear, actionable error messages:
```bash
# File not found
Error: File not found - nonexistent.msg

# Wrong file type
Warning: File does not have .msg extension - .txt
Attempting to process anyway...

# Permission issues
Error: Access denied to file - protected.msg

# Large files
Warning: Large file detected (75MB). Processing may take longer...

# Invalid MSG format
Unexpected error: File is not a valid MSG file: document.pdf
```

## Troubleshooting

### Linux: "The type initializer for 'Gdip' threw an exception"

If you encounter GDI+ related errors on Linux, install the required graphics library:

```bash
sudo apt update
sudo apt install libgdiplus libc6-dev
```

### Common Issues and Solutions

#### File Not Found Errors
Ensure that:
- The MSG file path is correct and the file exists
- You have read permissions for the input file
- The file path doesn't contain special characters that need escaping

#### Permission Errors
```bash
Error: Access denied to file - protected.msg
```
**Solution**: Check file permissions and ensure you have read access to the MSG file and write access to the output directory.

#### Invalid File Format
```bash
Error: File is not a valid MSG file
```
**Solution**: Verify the file is actually an Outlook MSG file. The application will warn about non-.msg extensions but will attempt processing anyway.

#### Large File Processing
```bash
Warning: Large file detected (75MB). Processing may take longer...
```
**Solution**: This is normal for large emails with attachments. Allow extra processing time or consider the file size limitations of your system.

#### Output Directory Issues
```bash
Error: Output directory not found - /nonexistent/path/
```
**Solution**: Ensure the output directory exists or use a valid path. The application will create the output file but not missing directories.

### Memory Issues
If processing very large MSG files causes memory issues:
- Close other applications to free up memory
- Process files individually rather than in batches
- Consider processing on a system with more available RAM

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

This project is open source. Please check the license file for details.

## Support

For issues or questions:
- Check existing issues in the repository
- Create a new issue with detailed information about the problem
- Include sample files (if possible) and error messages

---

**Note**: This tool is designed for legitimate email processing purposes. Always ensure you have proper authorization before processing email files that don't belong to you.
