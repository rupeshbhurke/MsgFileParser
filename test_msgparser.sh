#!/bin/bash

# MSG File Parser Test Script
# Tests all error conditions and success scenarios

set -e  # Exit on any error during setup

echo "=========================================="
echo "MSG File Parser - Comprehensive Test Suite"
echo "=========================================="
echo

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# Test counters
TOTAL_TESTS=0
PASSED_TESTS=0
FAILED_TESTS=0

# Function to run a test
run_test() {
    local test_name="$1"
    local command="$2"
    local expected_exit_code="$3"
    local expected_output_pattern="$4"
    
    echo -e "${BLUE}TEST:${NC} $test_name"
    echo "Command: $command"
    
    TOTAL_TESTS=$((TOTAL_TESTS + 1))
    
    # Run the command and capture output and exit code
    set +e  # Don't exit on command failure
    output=$(eval "$command" 2>&1)
    actual_exit_code=$?
    set -e  # Re-enable exit on error
    
    # Check exit code
    if [ "$actual_exit_code" -eq "$expected_exit_code" ]; then
        echo -e "${GREEN}✓ Exit code: $actual_exit_code (expected: $expected_exit_code)${NC}"
        exit_code_ok=true
    else
        echo -e "${RED}✗ Exit code: $actual_exit_code (expected: $expected_exit_code)${NC}"
        exit_code_ok=false
    fi
    
    # Check output pattern if provided
    if [ -n "$expected_output_pattern" ]; then
        if echo "$output" | grep -q "$expected_output_pattern"; then
            echo -e "${GREEN}✓ Output contains: '$expected_output_pattern'${NC}"
            output_ok=true
        else
            echo -e "${RED}✗ Output does not contain: '$expected_output_pattern'${NC}"
            echo "Actual output:"
            echo "$output"
            output_ok=false
        fi
    else
        output_ok=true
    fi
    
    # Overall test result
    if [ "$exit_code_ok" = true ] && [ "$output_ok" = true ]; then
        echo -e "${GREEN}✓ PASSED${NC}"
        PASSED_TESTS=$((PASSED_TESTS + 1))
    else
        echo -e "${RED}✗ FAILED${NC}"
        FAILED_TESTS=$((FAILED_TESTS + 1))
    fi
    
    echo "Output:"
    echo "$output"
    echo
    echo "---"
    echo
}

# Setup - Build the project
echo "Setting up test environment..."
echo "Building project..."
dotnet build > /dev/null 2>&1

# Create test files and directories
echo "Creating test files..."

# Create a valid MSG file (copy the existing one if available)
if [ -f "1.msg" ]; then
    cp "1.msg" "test_valid.msg"
else
    echo "Warning: No valid MSG file found. Some tests may fail."
fi

# Create invalid files
echo "not a msg file" > "test_invalid.txt"
echo "fake msg content" > "test_fake.msg"

# Create read-only file for permission test
echo "test content" > "test_readonly.msg"
chmod 444 "test_readonly.msg"

# Create directory for tests
mkdir -p test_output
mkdir -p test_readonly_dir
chmod 555 test_readonly_dir

echo "Setup complete!"
echo

# Test 1: No arguments (INVALID_USAGE)
run_test "No arguments provided" \
    "dotnet run" \
    1 \
    "ERROR: INVALID_USAGE"

# Test 2: Empty argument (INVALID_INPUT)
run_test "Empty file path" \
    "dotnet run ''" \
    2 \
    "ERROR: INVALID_INPUT"

# Test 3: File not found (FILE_NOT_FOUND)
run_test "Non-existent file" \
    "dotnet run 'nonexistent.msg'" \
    3 \
    "ERROR: FILE_NOT_FOUND"

# Test 4: Wrong file extension warning + invalid format
run_test "Invalid file format (.txt)" \
    "dotnet run 'test_invalid.txt'" \
    22 \
    "WARNING: INVALID_EXTENSION"

# Test 5: Output directory doesn't exist (DIRECTORY_NOT_FOUND)
run_test "Non-existent output directory" \
    "dotnet run 'test_invalid.txt' '/nonexistent/path/output.txt'" \
    7 \
    "ERROR: DIRECTORY_NOT_FOUND"

# Test 6: No write permission to output directory (WRITE_ACCESS_DENIED)
run_test "No write permission to output directory" \
    "dotnet run 'test_invalid.txt' 'test_readonly_dir/output.txt'" \
    8 \
    "ERROR: WRITE_ACCESS_DENIED"

# Test 7: Valid MSG file processing (SUCCESS)
if [ -f "test_valid.msg" ]; then
    run_test "Valid MSG file processing" \
        "dotnet run 'test_valid.msg' 'test_output/success.txt'" \
        0 \
        "SUCCESS: Processing completed"
else
    echo -e "${YELLOW}SKIPPED: Valid MSG file test (no valid MSG file available)${NC}"
    echo
fi

# Test 8: File overwrite warning
if [ -f "test_valid.msg" ]; then
    # Create a file first
    touch "test_output/overwrite_test.txt"
    run_test "File overwrite warning" \
        "dotnet run 'test_valid.msg' 'test_output/overwrite_test.txt'" \
        0 \
        "WARNING: FILE_EXISTS"
else
    echo -e "${YELLOW}SKIPPED: File overwrite test (no valid MSG file available)${NC}"
    echo
fi

# Test 9: Default output behavior
if [ -f "test_valid.msg" ]; then
    run_test "Default output behavior" \
        "dotnet run 'test_valid.msg'" \
        0 \
        "SUCCESS: Processing completed"
else
    echo -e "${YELLOW}SKIPPED: Default output test (no valid MSG file available)${NC}"
    echo
fi

# Test 10: Directory output behavior
if [ -f "test_valid.msg" ]; then
    run_test "Directory output behavior" \
        "dotnet run 'test_valid.msg' 'test_output/'" \
        0 \
        "SUCCESS: Processing completed"
else
    echo -e "${YELLOW}SKIPPED: Directory output test (no valid MSG file available)${NC}"
    echo
fi

# Cleanup
echo "Cleaning up test files..."
rm -f test_invalid.txt test_fake.msg test_readonly.msg test_valid.msg
rm -rf test_output
chmod 755 test_readonly_dir
rmdir test_readonly_dir
echo

# Test Results Summary
echo "=========================================="
echo "TEST RESULTS SUMMARY"
echo "=========================================="
echo "Total Tests:  $TOTAL_TESTS"
echo -e "Passed:       ${GREEN}$PASSED_TESTS${NC}"
echo -e "Failed:       ${RED}$FAILED_TESTS${NC}"

if [ $FAILED_TESTS -eq 0 ]; then
    echo -e "${GREEN}All tests passed! ✓${NC}"
    exit 0
else
    echo -e "${RED}Some tests failed! ✗${NC}"
    exit 1
fi
