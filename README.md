# ExcelMetadata

A Go library for extracting comprehensive metadata from Excel files (.xlsx) using [excelize](https://github.com/xuri/excelize). This library allows you to extract all structural information, styles, formulas, and content from Excel files and export them as JSON.

## Features

- 📊 **Complete Excel Structure Extraction**
  - Document properties (title, author, dates, etc.)
  - Sheet information (names, visibility, dimensions)
  - Cell data, formulas, and types
  - Merged cells with values
  - Row heights and column widths

- 🎨 **Style Information**
  - Font styles (bold, italic, color, size, etc.)
  - Fill patterns and colors
  - Borders and alignment
  - Number formats
  - Cell protection settings

- 🔗 **Rich Content Support**
  - Hyperlinks
  - Data validations
  - Comments (Note: Comment extraction not implemented in current version)
  - Images with formatting details
  - Named ranges (defined names)

- ⚡ **Performance Options**
  - Configurable extraction options
  - Cell limit per sheet
  - Selective feature extraction

## Installation

```bash
go get github.com/prongbang/excelmetadata
```

## Requirements

- Go 1.18 or higher
- github.com/xuri/excelize/v2 v2.9.1+

## Quick Start

### Basic Usage

```go
package main

import (
    "fmt"
    "log"
    "github.com/prongbang/excelmetadata"
)

func main() {
    // Extract metadata from Excel file
    metadata, err := excelmetadata.QuickExtract("sample.xlsx")
    if err != nil {
        log.Fatal(err)
    }

    fmt.Printf("File: %s\n", metadata.Filename)
    fmt.Printf("Created by: %s\n", metadata.Properties.Creator)
    fmt.Printf("Sheets: %d\n", len(metadata.Sheets))
}
```

### Export to JSON

```go
// Export to JSON string
jsonStr, err := excelmetadata.QuickExtractToJSON("sample.xlsx", true)
if err != nil {
    log.Fatal(err)
}
fmt.Println(jsonStr)

// Export directly to file
err = excelmetadata.QuickExtractToFile("sample.xlsx", "metadata.json", true)
if err != nil {
    log.Fatal(err)
}
```

### Advanced Usage with Options

```go
// Configure extraction options
options := &excelmetadata.Options{
    IncludeCellData:       true,
    IncludeStyles:         true,
    IncludeImages:         true,
    IncludeDefinedNames:   true,
    IncludeDataValidation: true,
    MaxCellsPerSheet:      1000, // Limit cells per sheet
}

// Create extractor with options
extractor, err := excelmetadata.New("large_file.xlsx", options)
if err != nil {
    log.Fatal(err)
}
defer extractor.Close()

// Extract metadata
metadata, err := extractor.Extract()
if err != nil {
    log.Fatal(err)
}

// Process metadata...
```

## Data Structures

### Metadata
The main structure containing all extracted information:

```go
type Metadata struct {
    Filename     string               // Original filename
    Properties   DocumentProperties   // Document properties
    Sheets       []SheetMetadata      // Sheet information
    DefinedNames []DefinedName        // Named ranges
    Styles       map[int]StyleDetails // Unique styles
    ExtractedAt  time.Time           // Extraction timestamp
}
```

### SheetMetadata
Information about each sheet:

```go
type SheetMetadata struct {
    Index           int                // Sheet index
    Name            string             // Sheet name
    Visible         bool               // Visibility status
    Dimensions      SheetDimensions    // Used range
    MergedCells     []MergedCell       // Merged cells
    DataValidations []DataValidation   // Validation rules
    Protection      *SheetProtection   // Protection settings
    RowHeights      map[int]float64    // Custom row heights
    ColWidths       map[string]float64 // Custom column widths
    Cells           []CellMetadata     // Cell data
    Images          []ImageMetadata    // Embedded images
}
```

### CellMetadata
Individual cell information:

```go
type CellMetadata struct {
    Address   string            // Cell address (e.g., "A1")
    Value     interface{}       // Cell value
    Formula   string            // Formula if present
    StyleID   int               // Style reference
    Type      excelize.CellType // Cell type
    Hyperlink *Hyperlink        // Hyperlink if present
}
```

## Extraction Options

| Option | Description | Default |
|--------|-------------|---------|
| `IncludeCellData` | Extract cell values and formulas | `true` |
| `IncludeStyles` | Extract style information | `true` |
| `IncludeImages` | Extract embedded images | `true` |
| `IncludeDefinedNames` | Extract named ranges | `true` |
| `IncludeDataValidation` | Extract data validation rules | `true` |
| `MaxCellsPerSheet` | Maximum cells to extract per sheet (0 = unlimited) | `0` |

## JSON Output Example

```json
{
  "filename": "sample.xlsx",
  "properties": {
    "title": "Sales Report",
    "creator": "John Doe",
    "created": "2024-01-15T10:30:00Z",
    "modified": "2024-01-20T14:45:00Z"
  },
  "sheets": [
    {
      "index": 0,
      "name": "Sheet1",
      "visible": true,
      "dimensions": {
        "startCell": "A1",
        "endCell": "D10",
        "rowCount": 10,
        "colCount": 4
      },
      "cells": [
        {
          "address": "A1",
          "value": "Product",
          "styleId": 1,
          "type": 3
        },
        {
          "address": "B2",
          "value": "100",
          "formula": "SUM(C2:D2)",
          "styleId": 2,
          "type": 2
        }
      ],
      "mergedCells": [
        {
          "startCell": "A1",
          "endCell": "B1",
          "value": "Header"
        }
      ]
    }
  ],
  "styles": {
    "1": {
      "font": {
        "bold": true,
        "size": 14,
        "color": "#000000"
      },
      "fill": {
        "type": "pattern",
        "pattern": 1,
        "color": ["#E0E0E0"]
      }
    }
  },
  "extractedAt": "2024-01-20T15:00:00Z"
}
```

## Use Cases

1. **Excel to JSON Conversion** - Convert Excel files to JSON for web applications
2. **Document Analysis** - Analyze Excel file structure and complexity
3. **Version Control** - Track changes in Excel files by comparing metadata
4. **Data Migration** - Extract data with complete formatting information
5. **Backup and Archival** - Create searchable metadata for Excel archives
6. **Template Generation** - Extract structure to create new Excel templates

## Performance Considerations

- For large files, use `MaxCellsPerSheet` to limit extraction
- Disable unnecessary features (styles, images) for faster extraction
- Image extraction includes binary data, which can significantly increase JSON size

## Limitations

- Currently supports .xlsx files only (not .xls)
- Some advanced Excel features may not be captured
- Images are extracted as binary data (base64 encoded in JSON)

## Error Handling

The library provides detailed error messages:

```go
extractor, err := excelmetadata.New("file.xlsx", nil)
if err != nil {
    // Handle file opening errors
    log.Printf("Failed to open file: %v", err)
    return
}
defer extractor.Close()

metadata, err := extractor.Extract()
if err != nil {
    // Handle extraction errors
    log.Printf("Failed to extract metadata: %v", err)
    return
}
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

This project uses the [excelize](https://github.com/xuri/excelize) library, which is licensed under the BSD 3-Clause License.

## Acknowledgments

- Built on top of [excelize](https://github.com/xuri/excelize)
- Inspired by the need for better Excel file introspection tools

## Example Projects

### 1. Excel Diff Tool
```go
// Compare two Excel files
metadata1, _ := excelmetadata.QuickExtract("version1.xlsx")
metadata2, _ := excelmetadata.QuickExtract("version2.xlsx")

// Compare sheet counts, cell values, styles, etc.
```

### 2. Excel Search Engine
```go
// Index Excel files for searching
files := []string{"report1.xlsx", "report2.xlsx", "report3.xlsx"}
var index []excelmetadata.Metadata

for _, file := range files {
    metadata, _ := excelmetadata.QuickExtract(file)
    index = append(index, *metadata)
}

// Search through metadata for specific content
```

### 3. Excel Structure Validator
```go
// Validate Excel files against a template
template, _ := excelmetadata.QuickExtract("template.xlsx")
submission, _ := excelmetadata.QuickExtract("submission.xlsx")

// Check if submission matches template structure
if len(template.Sheets) != len(submission.Sheets) {
    log.Println("Sheet count mismatch")
}
```

## Support

For issues, questions, or contributions, please visit the [GitHub repository](https://github.com/prongbang/excelmetadata).
