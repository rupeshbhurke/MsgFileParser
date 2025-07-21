# MSG File Parser

A .NET 8.0 console application that extracts content from Microsoft Outlook MSG files and converts them to readable text format.

## Features

- 📧 Extract email metadata (subject, sender, recipients, date)
- 📝 Support for multiple body formats (Plain Text, HTML, RTF) - **Not tested!**
- 🎯 Flexible output options (specific filename or directory)
- ⚡ Fast and efficient processing
- 🔧 Cross-platform compatibility (Windows, Linux, macOS)

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

## Error Handling

The application includes comprehensive error handling for:

- Missing input files
- Invalid MSG file format
- Non-existent output directories
- File permission issues
- Encoding problems

## Troubleshooting

### Linux: "The type initializer for 'Gdip' threw an exception"

If you encounter GDI+ related errors on Linux, install the required graphics library:

```bash
sudo apt update
sudo apt install libgdiplus libc6-dev
```

### File Not Found Errors

Ensure that:
- The MSG file path is correct and the file exists
- You have read permissions for the input file
- The output directory exists and you have write permissions

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
