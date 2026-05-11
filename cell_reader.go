package excelmetadata

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

// rawCell represents a single cell parsed from worksheet XML.
type rawCell struct {
	Ref     string // e.g. "A1"
	Row     int
	Col     int
	Value   string
	Type    string
	StyleID int
}

// rawRow represents a row of cells parsed from worksheet XML.
type rawRow struct {
	Index int
	Cells []rawCell
}

// readSheetCells reads the sheet XML directly from excelize.File.Pkg
// to extract ALL cells including those that have only a style attribute
// and no value — excelize's GetRows() skips those.
func readSheetCells(f *excelize.File, sheetName string) ([]rawCell, error) {
	sheetPath, err := getSheetXMLPath(f, sheetName)
	if err != nil {
		return nil, err
	}

	raw, ok := f.Pkg.Load(sheetPath)
	if !ok {
		return nil, fmt.Errorf("sheet xml not found in Pkg: %s", sheetPath)
	}

	data, ok := raw.([]byte)
	if !ok {
		return nil, fmt.Errorf("invalid xml data type for: %s", sheetPath)
	}

	dec := xml.NewDecoder(bytes.NewReader(data))
	var cells []rawCell
	var currentRow *rawRow

	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "row":
				currentRow = &rawRow{}
				for _, a := range t.Attr {
					if a.Name.Local == "r" {
						currentRow.Index, _ = strconv.Atoi(a.Value)
					}
				}

			case "c":
				if currentRow == nil {
					continue
				}

				cell := rawCell{}
				for _, a := range t.Attr {
					switch a.Name.Local {
					case "r":
						cell.Ref = a.Value
						col, row := splitCellRef(a.Value)
						cell.Col = colToIndex(col)
						cell.Row = row
					case "s":
						cell.StyleID, _ = strconv.Atoi(a.Value)
					case "t":
						cell.Type = a.Value
					}
				}

				// Read inner elements (<v>, <f>, etc.)
				for {
					innerTok, err := dec.Token()
					if err != nil {
						break
					}
					switch inner := innerTok.(type) {
					case xml.StartElement:
						if inner.Name.Local == "v" {
							var value string
							if err := dec.DecodeElement(&value, &inner); err == nil {
								cell.Value = value
							}
						}
					case xml.EndElement:
						if inner.Name.Local == "c" {
							goto doneCell
						}
					}
				}
			doneCell:
				currentRow.Cells = append(currentRow.Cells, cell)
			}

		case xml.EndElement:
			if t.Name.Local == "row" && currentRow != nil {
				cells = append(cells, currentRow.Cells...)
				currentRow = nil
			}
		}
	}

	return cells, nil
}

// getSheetXMLPath resolves a sheet name to its XML path inside the xlsx.
// Uses GetSheetList() to get the sheet order, then maps to file paths via
// the relationships file.
func getSheetXMLPath(f *excelize.File, sheetName string) (string, error) {
	sheetList := f.GetSheetList()
	sheetIdx := -1
	for i, name := range sheetList {
		if name == sheetName {
			sheetIdx = i
			break
		}
	}
	if sheetIdx < 0 {
		return "", fmt.Errorf("sheet not found: %s", sheetName)
	}

	// Standard path naming: xl/worksheets/sheet{N}.xml
	// N is 1-based and corresponds to the sheet's order in the workbook.
	return fmt.Sprintf("xl/worksheets/sheet%d.xml", sheetIdx+1), nil
}

// splitCellRef splits "A1" into ("A", 1).
func splitCellRef(ref string) (string, int) {
	i := 0
	for ; i < len(ref); i++ {
		if ref[i] >= '0' && ref[i] <= '9' {
			break
		}
	}
	col := ref[:i]
	row, _ := strconv.Atoi(ref[i:])
	return col, row
}

// colToIndex converts "A"->0, "B"->1, ..., "Z"->25, "AA"->26, etc.
func colToIndex(col string) int {
	col = strings.ToUpper(col)
	n := 0
	for _, c := range col {
		n = n*26 + int(c-'A'+1)
	}
	return n - 1
}
