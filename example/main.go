package main

import (
	"fmt"
	"log"
	"os"

	"github.com/prongbang/excelmetadata"
	"github.com/xuri/excelize/v2"
)

func main() {
	// Example 1: Quick extraction with default options
	quickExample()

	// Example 2: Custom options extraction
	customOptionsExample()

	// Example 3: Save to file
	saveToFileExample()

	// Example 4: Process specific metadata
	processMetadataExample()

	// Example 5: Extract from existing file handle
	existingFileExample()
}

// Example 1: Quick extraction with default options
func quickExample() {
	fmt.Println("=== Quick Extraction Example ===")

	// Method 1: Get metadata object
	metadata, err := excelmetadata.QuickExtract("sample.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	fmt.Printf("File: %s\n", metadata.Filename)
	fmt.Printf("Creator: %s\n", metadata.Properties.Creator)
	fmt.Printf("Number of sheets: %d\n", len(metadata.Sheets))

	// Method 2: Get JSON string directly
	jsonStr, err := excelmetadata.QuickExtractToJSON("sample.xlsx", true)
	if err != nil {
		log.Fatal(err)
	}

	fmt.Println("\nJSON Output (first 500 chars):")
	if len(jsonStr) > 500 {
		fmt.Println(jsonStr[:500] + "...")
	} else {
		fmt.Println(jsonStr)
	}

	// Method 3: Save directly to file
	err = excelmetadata.QuickExtractToFile("sample.xlsx", "metadata.json", true)
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println("\nMetadata saved to metadata.json")
}

// Example 2: Custom options extraction
func customOptionsExample() {
	fmt.Println("\n=== Custom Options Example ===")

	// Configure options
	options := &excelmetadata.Options{
		IncludeCellData:       true,
		IncludeStyles:         true,
		IncludeComments:       true,
		IncludeImages:         false,
		IncludeDefinedNames:   true,
		IncludeDataValidation: true,
		MaxCellsPerSheet:      1000, // Limit to 1000 cells per sheet
	}

	// Create extractor with custom options
	extractor, err := excelmetadata.New("sample.xlsx", options)
	if err != nil {
		log.Fatal(err)
	}
	defer extractor.Close()

	// Extract metadata
	metadata, err := extractor.Extract()
	if err != nil {
		log.Fatal(err)
	}

	// Process results
	fmt.Printf("Extracted %d sheets\n", len(metadata.Sheets))
	for _, sheet := range metadata.Sheets {
		fmt.Printf("  Sheet '%s': %d cells, %d merged regions\n",
			sheet.Name, len(sheet.Cells), len(sheet.MergedCells))
	}
}

// Example 3: Save to file with options
func saveToFileExample() {
	fmt.Println("\n=== Save to File Example ===")

	// Create extractor with minimal options for smaller output
	options := &excelmetadata.Options{
		IncludeCellData: false, // Only structure, no cell data
		IncludeStyles:   false,
		IncludeComments: false,
		IncludeImages:   true,
	}

	extractor, err := excelmetadata.New("sample.xlsx", options)
	if err != nil {
		log.Fatal(err)
	}
	defer extractor.Close()

	// Save as minified JSON
	err = extractor.ExtractToFile("structure.json", false)
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println("Saved structure to structure.json (minified)")

	// Save as pretty JSON
	err = extractor.ExtractToFile("structure_pretty.json", true)
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println("Saved structure to structure_pretty.json (formatted)")
}

// Example 4: Process specific metadata
func processMetadataExample() {
	fmt.Println("\n=== Process Metadata Example ===")

	metadata, err := excelmetadata.QuickExtract("sample.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	// Process document properties
	fmt.Println("\nDocument Properties:")
	fmt.Printf("  Title: %s\n", metadata.Properties.Title)
	fmt.Printf("  Created: %v\n", metadata.Properties.Created)
	fmt.Printf("  Modified: %v\n", metadata.Properties.Modified)

	// Process defined names
	if len(metadata.DefinedNames) > 0 {
		fmt.Println("\nDefined Names:")
		for _, name := range metadata.DefinedNames {
			fmt.Printf("  %s -> %s (scope: %s)\n",
				name.Name, name.RefersTo, name.Scope)
		}
	}

	// Process each sheet
	for _, sheet := range metadata.Sheets {
		fmt.Printf("\nSheet: %s\n", sheet.Name)

		// Dimensions
		fmt.Printf("  Dimensions: %s:%s (%d rows x %d cols)\n",
			sheet.Dimensions.StartCell, sheet.Dimensions.EndCell,
			sheet.Dimensions.RowCount, sheet.Dimensions.ColCount)

		// Merged cells
		if len(sheet.MergedCells) > 0 {
			fmt.Println("  Merged Cells:")
			for _, mc := range sheet.MergedCells {
				fmt.Printf("    %s:%s = %s\n",
					mc.StartCell, mc.EndCell, mc.Value)
			}
		}

		// Data validations
		if len(sheet.DataValidations) > 0 {
			fmt.Println("  Data Validations:")
			for _, dv := range sheet.DataValidations {
				fmt.Printf("    Range: %s, Type: %s\n",
					dv.Range, dv.Type)
			}
		}

		// Find formulas
		formulaCount := 0
		for _, cell := range sheet.Cells {
			if cell.Formula != "" {
				formulaCount++
			}
		}
		if formulaCount > 0 {
			fmt.Printf("  Formulas: %d cells\n", formulaCount)
		}

		// Hyperlinks
		hyperlinkCount := 0
		for _, cell := range sheet.Cells {
			if cell.Hyperlink != nil {
				hyperlinkCount++
			}
		}
		if hyperlinkCount > 0 {
			fmt.Printf("  Hyperlinks: %d\n", hyperlinkCount)
		}

		// Images and Charts
		if len(sheet.Images) > 0 {
			fmt.Printf("  Images: %d\n", len(sheet.Images))
		}
	}

	// Process styles
	if len(metadata.Styles) > 0 {
		fmt.Printf("\nUnique Styles: %d\n", len(metadata.Styles))

		// Count different style features
		fontStyles := 0
		fillStyles := 0
		borderStyles := 0

		for _, style := range metadata.Styles {
			if style.Font != nil {
				fontStyles++
			}
			if style.Fill != nil {
				fillStyles++
			}
			if len(style.Border) > 0 {
				borderStyles++
			}
		}

		fmt.Printf("  With custom font: %d\n", fontStyles)
		fmt.Printf("  With fill color: %d\n", fillStyles)
		fmt.Printf("  With borders: %d\n", borderStyles)
	}
}

// Example 5: Extract from existing file handle
func existingFileExample() {
	fmt.Println("\n=== Existing File Handle Example ===")

	// This is useful when you already have an excelize.File open
	// and want to extract metadata without reopening the file

	// Open file normally with excelize
	f, err := excelize.OpenFile("sample.xlsx")
	if err != nil {
		log.Fatal(err)
	}
	defer f.Close()

	// Do some operations with the file...
	// For example, read a specific cell
	value, _ := f.GetCellValue("Sheet1", "A1")
	fmt.Printf("Cell A1 value: %s\n", value)

	// Now extract metadata from the same file handle
	extractor, _ := excelmetadata.New("sample.xlsx", nil)
	metadata, err := extractor.Extract()
	if err != nil {
		log.Fatal(err)
	}

	fmt.Printf("Extracted metadata from open file: %d sheets found\n",
		len(metadata.Sheets))
}

// Utility function to analyze metadata size
func analyzeMetadataSize() {
	fmt.Println("\n=== Metadata Size Analysis ===")

	testOptions := []struct {
		name string
		opts *excelmetadata.Options
	}{
		{
			name: "Full extraction",
			opts: excelmetadata.DefaultOptions(),
		},
		{
			name: "Structure only",
			opts: &excelmetadata.Options{
				IncludeCellData: false,
				IncludeStyles:   false,
				IncludeComments: false,
			},
		},
		{
			name: "Cells only (no styles)",
			opts: &excelmetadata.Options{
				IncludeCellData: true,
				IncludeStyles:   false,
				IncludeComments: false,
				IncludeImages:   false,
			},
		},
		{
			name: "Limited cells",
			opts: &excelmetadata.Options{
				IncludeCellData:  true,
				MaxCellsPerSheet: 100,
			},
		},
	}

	for _, test := range testOptions {
		extractor, err := excelmetadata.New("sample.xlsx", test.opts)
		if err != nil {
			continue
		}

		jsonStr, err := extractor.ExtractToJSON(false)
		if err != nil {
			extractor.Close()
			continue
		}

		fmt.Printf("%s: %d bytes\n", test.name, len(jsonStr))
		extractor.Close()
	}
}

// Example: Search and filter specific data
func searchExample() {
	fmt.Println("\n=== Search Example ===")

	metadata, err := excelmetadata.QuickExtract("sample.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	// Find all cells with formulas
	fmt.Println("Cells with formulas:")
	for _, sheet := range metadata.Sheets {
		for _, cell := range sheet.Cells {
			if cell.Formula != "" {
				fmt.Printf("  %s!%s: %s\n",
					sheet.Name, cell.Address, cell.Formula)
			}
		}
	}

	// Find all cells with specific style
	targetStyleID := 15 // example style ID
	fmt.Printf("\nCells with style ID %d:\n", targetStyleID)
	for _, sheet := range metadata.Sheets {
		for _, cell := range sheet.Cells {
			if cell.StyleID == targetStyleID {
				fmt.Printf("  %s!%s: %v\n",
					sheet.Name, cell.Address, cell.Value)
			}
		}
	}

	// Find all hyperlinks
	fmt.Println("\nAll hyperlinks:")
	for _, sheet := range metadata.Sheets {
		for _, cell := range sheet.Cells {
			if cell.Hyperlink != nil {
				fmt.Printf("  %s!%s -> %s\n",
					sheet.Name, cell.Address, cell.Hyperlink.Link)
			}
		}
	}
}

// Example: Create a summary report
func summaryReportExample() {
	fmt.Println("\n=== Summary Report Example ===")

	metadata, err := excelmetadata.QuickExtract("sample.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	report := fmt.Sprintf(`Excel File Analysis Report
========================
File: %s
Created by: %s
Last modified by: %s
Last modified: %v

Summary:
- Total sheets: %d
- Visible sheets: %d
- Total cells with data: %d
- Total formulas: %d
- Total merged regions: %d
- Total hyperlinks: %d
- Unique styles: %d

`,
		metadata.Filename,
		metadata.Properties.Creator,
		metadata.Properties.LastModifiedBy,
		metadata.Properties.Modified,
		len(metadata.Sheets),
		countVisibleSheets(metadata),
		countTotalCells(metadata),
		countFormulas(metadata),
		countMergedCells(metadata),
		countHyperlinks(metadata),
		len(metadata.Styles),
	)

	// Add per-sheet details
	report += "Sheet Details:\n"
	for _, sheet := range metadata.Sheets {
		report += fmt.Sprintf("- %s: %d×%d (%d cells)\n",
			sheet.Name,
			sheet.Dimensions.RowCount,
			sheet.Dimensions.ColCount,
			len(sheet.Cells),
		)
	}

	fmt.Println(report)

	// Save report to file
	err = os.WriteFile("excel_analysis_report.txt", []byte(report), 0644)
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println("Report saved to excel_analysis_report.txt")
}

// Helper functions for summary report
func countVisibleSheets(m *excelmetadata.Metadata) int {
	count := 0
	for _, sheet := range m.Sheets {
		if sheet.Visible {
			count++
		}
	}
	return count
}

func countTotalCells(m *excelmetadata.Metadata) int {
	count := 0
	for _, sheet := range m.Sheets {
		count += len(sheet.Cells)
	}
	return count
}

func countFormulas(m *excelmetadata.Metadata) int {
	count := 0
	for _, sheet := range m.Sheets {
		for _, cell := range sheet.Cells {
			if cell.Formula != "" {
				count++
			}
		}
	}
	return count
}

func countMergedCells(m *excelmetadata.Metadata) int {
	count := 0
	for _, sheet := range m.Sheets {
		count += len(sheet.MergedCells)
	}
	return count
}

func countHyperlinks(m *excelmetadata.Metadata) int {
	count := 0
	for _, sheet := range m.Sheets {
		for _, cell := range sheet.Cells {
			if cell.Hyperlink != nil {
				count++
			}
		}
	}
	return count
}
