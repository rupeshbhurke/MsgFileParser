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

The application supports four different usage patterns:

- **Subject**: Email subject line
- **From**: Sender information
- **Date**: Email timestamp (YYYY-MM-DD HH:mm:ss format)
- **Recipients**: List of all recipients with email addresses and display names
Creates a file with your exact specified name and location. You can specify the export format as a third argument:

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
 
### 4. Export Format
You can specify the export format as a third argument:

- `--text` (default): Export as plain text
- `--html`: Export as HTML


Please join us for the meeting tomorrow at 2 PM.

Best regards,
John
```

## Dependencies

This project uses the following NuGet packages:

- [MSGReader](https://www.nuget.org/packages/MSGReader/) (4.3.0) - For parsing MSG files
- [System.Text.Encoding.CodePages](https://www.nuget.org/packages/System.Text.Encoding.CodePages/) (8.0.0) - For proper text encoding support

dotnet run "email.msg" "/path/to/output.html" --html
## Technical Details

- **Target Framework**: .NET 8.0
- **Platform**: Cross-platform (Windows, Linux, macOS)
- **Architecture**: Console Application
- **Language**: C#

## Error Handling & Exit Codes

The application provides standardized exit codes and structured output messages for easy integration with other processes.

### Exit Codes

| Code | Category | Description |
|------|----------|-------------|
| **0** | SUCCESS | Processing completed successfully |
| **1** | INVALID_USAGE | No input file specified |
| **2** | INVALID_INPUT | MSG file path is empty or invalid |
| **3** | FILE_NOT_FOUND | Input MSG file does not exist |
| **4** | ACCESS_DENIED | Cannot read the input file (permission denied) |
| **5** | IO_ERROR | Input/output error accessing the file |
| **6** | INVALID_OUTPUT | Output path is empty or invalid |
| **7** | DIRECTORY_NOT_FOUND | Output directory does not exist |
| **8** | WRITE_ACCESS_DENIED | No write permission for output directory |
| **9** | WRITE_IO_ERROR | Cannot write to output directory |
| **10** | INVALID_PATH | Invalid file path format |
| **11** | NOT_SUPPORTED | Operation not supported |
| **20** | INVALID_MSG_FORMAT | Invalid MSG file format |
| **21** | OUT_OF_MEMORY | File too large to process |
| **22** | INVALID_MSG_FILE | Not a valid MSG file |
| **23** | WRITE_FAILED | Failed to write output file |
| **24** | WRITE_ACCESS_DENIED | Access denied writing output file |
| **25** | PROCESSING_FAILED | General processing failure |
| **99** | UNEXPECTED | Unexpected error occurred |

## Exit Codes

Exit codes are defined in the `ExitCode` enum (`ExitCode.cs`) for maintainability:

| Enum Value         | Code | Description                                 |
|--------------------|------|---------------------------------------------|
| Success            | 0    | Success                                    |
| InvalidUsage       | 1    | Invalid usage (missing or incorrect args)   |
| FileNotFound       | 3    | File not found                             |
| AccessDenied       | 4    | Access denied                              |
| IoError            | 5    | IO error                                   |
| DirectoryNotFound  | 7    | Directory not found                        |
| InvalidPath        | 10   | Invalid path                               |
| NotSupported       | 11   | Not supported                              |
| Unexpected         | 99   | Unexpected error                           |

Reference: See `ExitCode.cs` for the authoritative list.

### Output Message Format

All messages follow a structured format for easy parsing:

#### Success Messages
```
SUCCESS: Processing completed - Output: /path/to/output.txt
```

#### Error Messages
```
ERROR: <ERROR_CODE> - <Human readable description>
```

#### Warning Messages
```
WARNING: <WARNING_CODE> - <Description>
```

#### Info Messages
```
INFO: <Description>
```

### Example Output

#### Successful Processing
```bash
$ dotnet run "email.msg" "output.txt"
MSG File Parser v1.0
===================
INFO: Using HTML body (contains formatting)
INFO: File created successfully: output.txt
SUCCESS: Processing completed - Output: output.txt
$ echo $?
0
```

#### Error Case
```bash
$ dotnet run "missing.msg"
MSG File Parser v1.0
===================
ERROR: FILE_NOT_FOUND - missing.msg
$ echo $?
3
```

#### Warning Case
```bash
$ dotnet run "document.pdf"
MSG File Parser v1.0
===================
WARNING: INVALID_EXTENSION - File has .pdf extension, expected .msg
INFO: Attempting to process anyway...
ERROR: INVALID_MSG_FILE - Not a valid MSG file: document.pdf
$ echo $?
22
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

## Integration Guide

### For Calling Applications

When integrating this tool into other applications, monitor both exit codes and output messages:

#### Shell Script Integration
```bash
#!/bin/bash
output=$(dotnet run "email.msg" "output.txt" 2>&1)
exit_code=$?

if [ $exit_code -eq 0 ]; then
    echo "Success: $output"
    # Process the generated file
else
    echo "Error (Code $exit_code): $output"
    # Handle the error based on exit code
fi
```

#### Python Integration
```python
import subprocess
import sys

def process_msg_file(msg_path, output_path):
    try:
        result = subprocess.run([
            "dotnet", "run", msg_path, output_path
        ], capture_output=True, text=True, cwd="/path/to/MsgFileParser")
        
        if result.returncode == 0:
            print(f"Success: {result.stdout}")
            return True
        else:
            print(f"Error {result.returncode}: {result.stdout}")
            return False
    except Exception as e:
        print(f"Failed to execute: {e}")
        return False
```

#### Error Code Handling
```bash
case $exit_code in
    0) echo "Processing completed successfully" ;;
    1) echo "Usage error - check command arguments" ;;
    3) echo "Input file not found" ;;
    7) echo "Output directory missing" ;;
    22) echo "Invalid MSG file format" ;;
    *) echo "Unexpected error occurred" ;;
esac
```

## Testing

The project includes a comprehensive test script that validates all functionality and error conditions.

### Running Tests

Execute the test script to validate all scenarios:

```bash
./test_msgparser.sh
```

**Note**: Make sure the script is executable. If not, run:
```bash
chmod +x test_msgparser.sh
```

### Test Coverage

The test script validates:

| Test Scenario | Validates |
|---------------|-----------|
| **No arguments** | INVALID_USAGE (Exit Code 1) |
| **Empty file path** | INVALID_INPUT (Exit Code 2) |
| **Non-existent file** | FILE_NOT_FOUND (Exit Code 3) |
| **Invalid file format** | INVALID_MSG_FILE (Exit Code 22) |
| **Missing output directory** | DIRECTORY_NOT_FOUND (Exit Code 7) |
| **No write permissions** | WRITE_ACCESS_DENIED (Exit Code 8) |
| **Valid MSG processing** | SUCCESS (Exit Code 0) |
| **File overwrite warning** | Warning handling + SUCCESS |
| **Default output behavior** | Auto-filename generation |
| **Directory output** | Auto-filename in specified directory |

### Test Output

The script provides color-coded results:
- 🟢 **Green**: Passed tests
- 🔴 **Red**: Failed tests  
- 🟡 **Yellow**: Skipped tests (missing prerequisites)
- 🔵 **Blue**: Test information

```bash
==========================================
TEST RESULTS SUMMARY
==========================================
Total Tests:  10
Passed:       10
Failed:       0
All tests passed! ✓
```

### Test Requirements

- .NET 8.0 SDK installed
- Valid MSG file in the project directory (copies `1.msg` if available)
- Write permissions in the project directory
- Linux/Unix environment (for permission testing)

### Automated Integration

The test script returns appropriate exit codes:
- **Exit Code 0**: All tests passed
- **Exit Code 1**: Some tests failed

This makes it suitable for CI/CD pipelines:

```bash
# In your CI/CD pipeline
./test_msgparser.sh
if [ $? -eq 0 ]; then
    echo "All tests passed - ready for deployment"
else
    echo "Tests failed - deployment blocked"
    exit 1
fi
```

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. **Run the test suite**: `./test_msgparser.sh`
5. Ensure all tests pass
6. Test thoroughly with your own MSG files
7. Submit a pull request

### Development Guidelines

- Follow existing code style and patterns
- Add appropriate error handling for new features
- Update exit codes table if adding new error conditions
- Test edge cases and error scenarios
- Update README.md if adding new functionality

## License

This project is open source. Please check the license file for details.

## Support

For issues or questions:
- Check existing issues in the repository
- Create a new issue with detailed information about the problem
- Include sample files (if possible) and error messages

---

**Note**: This tool is designed for legitimate email processing purposes. Always ensure you have proper authorization before processing email files that don't belong to you.
