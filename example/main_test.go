package main_test

import (
	"fmt"
	"testing"

	"github.com/prongbang/excelmetadata"
)

func TestGetMetadata(t *testing.T) {
	fmt.Println("=== Quick Extraction JSON ===")
	err := excelmetadata.QuickExtractToFile("sample.xlsx", "output/sample.metadata.json", true)
	fmt.Println(err)
}

func TestGetMetadataWithSheet(t *testing.T) {
	fmt.Println("=== Quick Extraction GO ===")
	err := excelmetadata.QuickExtractToFile("sample.xlsx", "output/sample.metadata.go", false)
	fmt.Println(err)
}
