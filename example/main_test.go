package main_test

import (
	"fmt"
	"testing"

	"github.com/prongbang/excelmetadata"
)

func TestGetMetadata(t *testing.T) {
	fmt.Println("=== Quick Extraction Example ===")

	var err error
	err = excelmetadata.QuickExtractToFile("sample.xlsx", "excelmeta/simaple.metadata.json", true)
	fmt.Println(err)
	err = excelmetadata.QuickExtractToFile("sample.xlsx", "excelmeta/simaple_metadata.go", false)
	fmt.Println(err)
}
