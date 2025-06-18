package main

import (
	"encoding/json"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/prongbang/excelmetadata"
	"github.com/urfave/cli/v2"
)

const version = "v1.0.3"

func main() {
	app := &cli.App{
		Name:    "excelmetadata",
		Usage:   "Excel Metadata CLI Tool",
		Version: version,
		Commands: []*cli.Command{
			{
				Name:    "extract",
				Aliases: []string{"e"},
				Usage:   "Extract metadata from Excel file",
				Flags: []cli.Flag{
					&cli.StringFlag{
						Name:    "output",
						Aliases: []string{"o"},
						Usage:   "Output JSON file path",
					},
					&cli.BoolFlag{
						Name:    "pretty",
						Aliases: []string{"p"},
						Usage:   "Pretty print JSON output",
					},
					&cli.IntFlag{
						Name:    "max-cells",
						Aliases: []string{"m"},
						Usage:   "Maximum cells per sheet (0 for unlimited)",
						Value:   0,
					},
					&cli.BoolFlag{
						Name:  "no-styles",
						Usage: "Exclude styles from extraction",
					},
					&cli.BoolFlag{
						Name:  "no-images",
						Usage: "Exclude images from extraction",
					},
				},
				Action: handleExtract,
			},
			{
				Name:    "compare",
				Aliases: []string{"c"},
				Usage:   "Compare two Excel files",
				Flags: []cli.Flag{
					&cli.BoolFlag{
						Name:    "detail",
						Aliases: []string{"d"},
						Usage:   "Show detailed comparison",
					},
				},
				Action: handleCompare,
			},
			{
				Name:    "search",
				Aliases: []string{"s"},
				Usage:   "Search Excel files in directory",
				Flags: []cli.Flag{
					&cli.StringFlag{
						Name:    "pattern",
						Aliases: []string{"p"},
						Usage:   "Search pattern",
					},
					&cli.BoolFlag{
						Name:    "recursive",
						Aliases: []string{"r"},
						Usage:   "Search in subdirectories",
					},
				},
				Action: handleSearch,
			},
		},
	}

	if err := app.Run(os.Args); err != nil {
		log.Fatal(err)
	}
}

func handleExtract(c *cli.Context) error {
	inputFile := c.Args().First()
	if inputFile == "" {
		return fmt.Errorf("please provide an input file")
	}

	options := &excelmetadata.Options{
		IncludeCellData:       true,
		IncludeStyles:         !c.Bool("no-styles"),
		IncludeImages:         !c.Bool("no-images"),
		IncludeDefinedNames:   true,
		IncludeDataValidation: true,
		MaxCellsPerSheet:      c.Int("max-cells"),
	}

	extractor, err := excelmetadata.New(inputFile, options)
	if err != nil {
		return fmt.Errorf("failed to create extractor: %v", err)
	}
	defer func(extractor *excelmetadata.Extractor) {
		_ = extractor.Close()
	}(extractor)

	metadata, err := extractor.Extract()
	if err != nil {
		return fmt.Errorf("failed to extract metadata: %v", err)
	}

	outputFile := c.String("output")
	fmt.Println("output file:", outputFile)
	if outputFile == "" {
		// Print to stdout
		var jsonData []byte
		if c.Bool("pretty") {
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
		err = excelmetadata.QuickExtractToFile(inputFile, outputFile, c.Bool("pretty"))
		if err != nil {
			return fmt.Errorf("failed to save to file: %v", err)
		}
		fmt.Printf("Metadata saved to %s\n", outputFile)
	}

	return nil
}

func handleCompare(c *cli.Context) error {
	if c.Args().Len() < 2 {
		return fmt.Errorf("please provide two files to compare")
	}

	file1 := c.Args().Get(0)
	file2 := c.Args().Get(1)

	metadata1, err := excelmetadata.QuickExtract(file1)
	if err != nil {
		return fmt.Errorf("failed to extract metadata from %s: %v", file1, err)
	}

	metadata2, err := excelmetadata.QuickExtract(file2)
	if err != nil {
		return fmt.Errorf("failed to extract metadata from %s: %v", file2, err)
	}

	fmt.Printf("Comparing %s with %s:\n", file1, file2)
	fmt.Printf("Sheets: %d vs %d\n", len(metadata1.Sheets), len(metadata2.Sheets))

	if c.Bool("detail") {
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

func handleSearch(c *cli.Context) error {
	searchDir := c.Args().First()
	if searchDir == "" {
		return fmt.Errorf("please provide a directory to search")
	}

	searchPattern := c.String("pattern")
	if searchPattern == "" {
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

			if searchInMetadata(metadata, searchPattern) {
				fmt.Printf("Match found in: %s\n", path)
			}
		}
		return nil
	}

	if c.Bool("recursive") {
		return filepath.Walk(searchDir, walkFunc)
	}

	files, err := filepath.Glob(filepath.Join(searchDir, "*.xlsx"))
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
