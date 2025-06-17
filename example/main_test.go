package main_test

import (
	"encoding/json"
	"fmt"
	"log"
	"os"
	"testing"

	"github.com/prongbang/excelmetadata"
)

func TestGetMetadata(t *testing.T) {
	fmt.Println("=== Quick Extraction Example ===")

	metadata, err := excelmetadata.QuickExtract("sample2.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	data, _ := json.Marshal(metadata)
	os.WriteFile("simaple2.metadata.json", data, 0644)
}
