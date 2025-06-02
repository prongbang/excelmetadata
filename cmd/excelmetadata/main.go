package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/prongbang/excelmetadata"
)

const version = "v1.0.3"

type CommandFlags struct {
	// Extract command flags
	outputFile  string
	prettyPrint bool
	maxCells    int
	noStyles    bool
	noImages    bool
	inputFile   string
	command     string
	// Compare command flags
	file1       string
	file2       string
	showDetails bool
	// Search command flags
	searchDir       string
	searchPattern   string
	searchRecursive bool
}

func main() {
	flags := parseFlags()

	switch flags.command {
	case "extract":
		if err := handleExtract(flags); err != nil {
			log.Fatal(err)
		}
	case "compare":
		if err := handleCompare(flags); err != nil {
			log.Fatal(err)
		}
	case "search":
		if err := handleSearch(flags); err != nil {
			log.Fatal(err)
		}
	case "version":
		fmt.Printf("excelmetada version %s\n", version)
	default:
		fmt.Printf("Unknown command: %s\n", flags.command)
		printUsage()
		os.Exit(1)
	}
}

func parseFlags() *CommandFlags {
	flags := &CommandFlags{}

	// Define command-line flags
	extractCmd := flag.NewFlagSet("extract", flag.ExitOnError)
	extractCmd.StringVar(&flags.outputFile, "output", "", "Output JSON file path")
	extractCmd.StringVar(&flags.outputFile, "o", "", "Output JSON file path (shorthand)")
	extractCmd.BoolVar(&flags.prettyPrint, "pretty", false, "Pretty print JSON output")
	extractCmd.BoolVar(&flags.prettyPrint, "p", false, "Pretty print JSON output (shorthand)")
	extractCmd.IntVar(&flags.maxCells, "max-cells", 0, "Maximum cells per sheet (0 for unlimited)")
	extractCmd.IntVar(&flags.maxCells, "m", 0, "Maximum cells per sheet (shorthand)")
	extractCmd.BoolVar(&flags.noStyles, "no-styles", false, "Exclude styles from extraction")
	extractCmd.BoolVar(&flags.noImages, "no-images", false, "Exclude images from extraction")

	compareCmd := flag.NewFlagSet("compare", flag.ExitOnError)
	compareCmd.BoolVar(&flags.showDetails, "detail", false, "Show detailed comparison")
	compareCmd.BoolVar(&flags.showDetails, "d", false, "Show detailed comparison (shorthand)")

	searchCmd := flag.NewFlagSet("search", flag.ExitOnError)
	searchCmd.StringVar(&flags.searchPattern, "pattern", "", "Search pattern")
	searchCmd.StringVar(&flags.searchPattern, "p", "", "Search pattern (shorthand)")
	searchCmd.BoolVar(&flags.searchRecursive, "recursive", false, "Search in subdirectories")
	searchCmd.BoolVar(&flags.searchRecursive, "r", false, "Search in subdirectories (shorthand)")

	if len(os.Args) < 2 {
		printUsage()
		os.Exit(1)
	}

	flags.command = os.Args[1]

	switch flags.command {
	case "extract":
		extractCmd.Parse(os.Args[2:])
		if extractCmd.NArg() > 0 {
			flags.inputFile = extractCmd.Arg(0)
		}
	case "compare":
		compareCmd.Parse(os.Args[2:])
		if compareCmd.NArg() > 1 {
			flags.file1 = compareCmd.Arg(0)
			flags.file2 = compareCmd.Arg(1)
		}
	case "search":
		searchCmd.Parse(os.Args[2:])
		if searchCmd.NArg() > 0 {
			flags.searchDir = searchCmd.Arg(0)
		}
	case "version":
		fmt.Println(version)
	default:
		printUsage()
		os.Exit(1)
	}

	return flags
}

func printUsage() {
	fmt.Println(`Excel Metadata CLI Tool

Usage:
  excelmetadata <command> [flags]

Commands:
  extract     Extract metadata from Excel file
             excelmetadata extract [-o output] [-p] [-m max-cells] [--no-styles] [--no-images] <file>

  compare     Compare two Excel files
             excelmetadata compare [-d] <file1> <file2>

  search      Search Excel files in directory
             excelmetadata search [-p pattern] [-r] <directory>

  version     Show version information

Examples:
  excelmetadata extract -o metadata.json -p input.xlsx
  excelmetadata compare -d file1.xlsx file2.xlsx
  excelmetadata search -p "Sales" -r /path/to/documents
`)
}

func handleExtract(flags *CommandFlags) error {
	if flags.inputFile == "" {
		return fmt.Errorf("please provide an input file")
	}

	options := &excelmetadata.Options{
		IncludeCellData:       true,
		IncludeStyles:         !flags.noStyles,
		IncludeImages:         !flags.noImages,
		IncludeDefinedNames:   true,
		IncludeDataValidation: true,
		MaxCellsPerSheet:      flags.maxCells,
	}

	extractor, err := excelmetadata.New(flags.inputFile, options)
	if err != nil {
		return fmt.Errorf("failed to create extractor: %v", err)
	}
	defer extractor.Close()

	metadata, err := extractor.Extract()
	if err != nil {
		return fmt.Errorf("failed to extract metadata: %v", err)
	}

	if flags.outputFile == "" {
		// Print to stdout
		var jsonData []byte
		if flags.prettyPrint {
			jsonData, err = json.MarshalIndent(metadata, "", "  ")
		} else {
			jsonData, err = json.Marshal(metadata)
		}
		if err != nil {
			return fmt.Errorf("failed to marshal JSON: %v", err)
		}
		fmt.Println(string(jsonData))
	} else {
		// Save to file
		err = excelmetadata.QuickExtractToFile(flags.inputFile, flags.outputFile, flags.prettyPrint)
		if err != nil {
			return fmt.Errorf("failed to save to file: %v", err)
		}
		fmt.Printf("Metadata saved to %s\n", flags.outputFile)
	}

	return nil
}

func handleCompare(flags *CommandFlags) error {
	if flags.file1 == "" || flags.file2 == "" {
		return fmt.Errorf("please provide two files to compare")
	}

	metadata1, err := excelmetadata.QuickExtract(flags.file1)
	if err != nil {
		return fmt.Errorf("failed to extract metadata from %s: %v", flags.file1, err)
	}

	metadata2, err := excelmetadata.QuickExtract(flags.file2)
	if err != nil {
		return fmt.Errorf("failed to extract metadata from %s: %v", flags.file2, err)
	}

	fmt.Printf("Comparing %s with %s:\n", flags.file1, flags.file2)
	fmt.Printf("Sheets: %d vs %d\n", len(metadata1.Sheets), len(metadata2.Sheets))

	if flags.showDetails {
		for i, sheet1 := range metadata1.Sheets {
			if i < len(metadata2.Sheets) {
				sheet2 := metadata2.Sheets[i]
				fmt.Printf("\nSheet %d:\n", i+1)
				fmt.Printf("  Name: %s vs %s\n", sheet1.Name, sheet2.Name)
				fmt.Printf("  Cells: %d vs %d\n", len(sheet1.Cells), len(sheet2.Cells))
			}
		}
	}

	return nil
}

func handleSearch(flags *CommandFlags) error {
	if flags.searchDir == "" {
		return fmt.Errorf("please provide a directory to search")
	}

	if flags.searchPattern == "" {
		return fmt.Errorf("please provide a search pattern")
	}

	walkFunc := func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		if !info.IsDir() && strings.HasSuffix(strings.ToLower(path), ".xlsx") {
			metadata, err := excelmetadata.QuickExtract(path)
			if err != nil {
				fmt.Printf("Error processing %s: %v\n", path, err)
				return nil
			}

			if searchInMetadata(metadata, flags.searchPattern) {
				fmt.Printf("Match found in: %s\n", path)
			}
		}
		return nil
	}

	if flags.searchRecursive {
		return filepath.Walk(flags.searchDir, walkFunc)
	}

	files, err := filepath.Glob(filepath.Join(flags.searchDir, "*.xlsx"))
	if err != nil {
		return fmt.Errorf("failed to list Excel files: %v", err)
	}

	for _, file := range files {
		if err := walkFunc(file, nil, nil); err != nil {
			return err
		}
	}

	return nil
}

func searchInMetadata(metadata *excelmetadata.Metadata, pattern string) bool {
	if pattern == "" {
		return false
	}

	for _, sheet := range metadata.Sheets {
		if strings.Contains(sheet.Name, pattern) {
			return true
		}

		for _, cell := range sheet.Cells {
			if value, ok := cell.Value.(string); ok {
				if strings.Contains(value, pattern) {
					return true
				}
			}
		}
	}

	return false
}
