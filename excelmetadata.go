package excelmetadata

import (
	"encoding/json"
	"fmt"
	"os"
	"path"
	"path/filepath"
	"strings"
	"time"

	"github.com/pkg/errors"
	"github.com/xuri/excelize/v2"
)

// Extractor is the main interface for extracting Excel metadata
type Extractor struct {
	file     *excelize.File
	filename string
	options  *Options
}

// Options configures the extraction behavior
type Options struct {
	IncludeCellData       bool
	IncludeStyles         bool
	IncludeImages         bool
	IncludeDefinedNames   bool
	IncludeDataValidation bool
	MaxCellsPerSheet      int
}

// DefaultOptions returns recommended default options
func DefaultOptions() *Options {
	return &Options{
		IncludeCellData:       true,
		IncludeStyles:         true,
		IncludeImages:         true,
		IncludeDefinedNames:   true,
		IncludeDataValidation: true,
		MaxCellsPerSheet:      0,
	}
}

// Metadata represents the complete Excel file metadata
type Metadata struct {
	Filename     string               `json:"filename"`
	Properties   DocumentProperties   `json:"properties"`
	Sheets       []SheetMetadata      `json:"sheets"`
	DefinedNames []DefinedName        `json:"definedNames,omitempty"`
	Styles       map[int]StyleDetails `json:"styles,omitempty"`
	ExtractedAt  time.Time            `json:"extractedAt"`
}

// DocumentProperties contains Excel document properties
type DocumentProperties struct {
	Title          string `json:"title,omitempty"`
	Subject        string `json:"subject,omitempty"`
	Creator        string `json:"creator,omitempty"`
	Keywords       string `json:"keywords,omitempty"`
	Description    string `json:"description,omitempty"`
	LastModifiedBy string `json:"lastModifiedBy,omitempty"`
	Category       string `json:"category,omitempty"`
	Version        string `json:"version,omitempty"`
	Created        string `json:"created,omitempty"`
	Modified       string `json:"modified,omitempty"`
}

// SheetMetadata contains metadata for a single sheet
type SheetMetadata struct {
	Index           int                `json:"index"`
	Name            string             `json:"name"`
	Visible         bool               `json:"visible"`
	Dimensions      SheetDimensions    `json:"dimensions"`
	MergedCells     []MergedCell       `json:"mergedCells,omitempty"`
	DataValidations []DataValidation   `json:"dataValidations,omitempty"`
	Protection      *SheetProtection   `json:"protection,omitempty"`
	RowHeights      map[int]float64    `json:"rowHeights,omitempty"`
	ColWidths       map[string]float64 `json:"colWidths,omitempty"`
	Cells           []CellMetadata     `json:"cells,omitempty"`
	Images          []ImageMetadata    `json:"images,omitempty"`
}

// SheetDimensions represents the used range of a sheet
type SheetDimensions struct {
	StartCell string `json:"startCell"`
	EndCell   string `json:"endCell"`
	RowCount  int    `json:"rowCount"`
	ColCount  int    `json:"colCount"`
}

// CellMetadata contains metadata for a single cell
type CellMetadata struct {
	Address   string            `json:"address"`
	Value     interface{}       `json:"value,omitempty"`
	Formula   string            `json:"formula,omitempty"`
	StyleID   int               `json:"styleId,omitempty"`
	Type      excelize.CellType `json:"type"`
	Hyperlink *Hyperlink        `json:"hyperlink,omitempty"`
}

// MergedCell represents a merged cell range
type MergedCell struct {
	StartCell string `json:"startCell"`
	EndCell   string `json:"endCell"`
	Value     string `json:"value,omitempty"`
}

// DataValidation represents data validation rules
type DataValidation struct {
	Range        string  `json:"range"`
	Type         string  `json:"type"`
	Operator     string  `json:"operator,omitempty"`
	Formula1     string  `json:"formula1,omitempty"`
	Formula2     string  `json:"formula2,omitempty"`
	ShowError    bool    `json:"showError"`
	ErrorTitle   *string `json:"errorTitle,omitempty"`
	ErrorMessage *string `json:"errorMessage,omitempty"`
}

// SheetProtection represents sheet protection settings
type SheetProtection struct {
	Protected           bool   `json:"protected"`
	Password            string `json:"password,omitempty"`
	EditObjects         bool   `json:"editObjects"`
	EditScenarios       bool   `json:"editScenarios"`
	SelectLockedCells   bool   `json:"selectLockedCells"`
	SelectUnlockedCells bool   `json:"selectUnlockedCells"`
}

// DefinedName represents a named range
type DefinedName struct {
	Name     string `json:"name"`
	RefersTo string `json:"refersTo"`
	Scope    string `json:"scope,omitempty"`
}

// StyleDetails contains detailed style information
type StyleDetails struct {
	Font         *FontStyle      `json:"font,omitempty"`
	Fill         *FillStyle      `json:"fill,omitempty"`
	Border       []BorderStyle   `json:"border,omitempty"`
	Alignment    *AlignmentStyle `json:"alignment,omitempty"`
	NumberFormat int             `json:"numberFormat,omitempty"`
	Protection   *Protection     `json:"protection,omitempty"`
}

// FontStyle represents font formatting
type FontStyle struct {
	Bold      bool    `json:"bold,omitempty"`
	Italic    bool    `json:"italic,omitempty"`
	Underline string  `json:"underline,omitempty"`
	Strike    bool    `json:"strike,omitempty"`
	Family    string  `json:"family,omitempty"`
	Size      float64 `json:"size,omitempty"`
	Color     string  `json:"color,omitempty"`
}

// FillStyle represents cell fill formatting
type FillStyle struct {
	Type    string   `json:"type,omitempty"`
	Pattern int      `json:"pattern,omitempty"`
	Color   []string `json:"color,omitempty"`
}

// BorderStyle represents cell border formatting
type BorderStyle struct {
	Type  string `json:"type,omitempty"`
	Color string `json:"color,omitempty"`
	Style int    `json:"style,omitempty"`
}

// AlignmentStyle represents text alignment
type AlignmentStyle struct {
	Horizontal   string `json:"horizontal,omitempty"`
	Vertical     string `json:"vertical,omitempty"`
	WrapText     bool   `json:"wrapText,omitempty"`
	TextRotation int    `json:"textRotation,omitempty"`
	Indent       int    `json:"indent,omitempty"`
	ShrinkToFit  bool   `json:"shrinkToFit,omitempty"`
}

// Protection represents cell protection settings
type Protection struct {
	Hidden bool `json:"hidden,omitempty"`
	Locked bool `json:"locked,omitempty"`
}

// Hyperlink represents a cell hyperlink
type Hyperlink struct {
	Link string `json:"link"`
}

// ImageMetadata represents image information
type ImageMetadata struct {
	Cell       string       `json:"cell"`
	File       []byte       `json:"file"`
	Extension  string       `json:"extension"`
	InsertType byte         `json:"insertType"`
	Format     *ImageFormat `json:"format"`
}

type ImageFormat struct {
	AltText             string
	PrintObject         *bool
	Locked              *bool
	LockAspectRatio     bool
	AutoFit             bool
	AutoFitIgnoreAspect bool
	OffsetX             int
	OffsetY             int
	ScaleX              float64
	ScaleY              float64
	Hyperlink           string
	HyperlinkType       string
	Positioning         string
}

// New creates a new Extractor instance
func New(filename string, options *Options) (*Extractor, error) {
	if options == nil {
		options = DefaultOptions()
	}

	f, err := excelize.OpenFile(filename)
	if err != nil {
		return nil, fmt.Errorf("failed to open file: %w", err)
	}

	return &Extractor{
		file:     f,
		filename: filename,
		options:  options,
	}, nil
}

// Extract performs the metadata extraction
func (e *Extractor) Extract() (*Metadata, error) {
	metadata := &Metadata{
		Filename:    e.filename,
		ExtractedAt: time.Now(),
		Sheets:      []SheetMetadata{},
	}

	// Extract document properties
	if props, err := e.extractDocumentProperties(); err == nil {
		metadata.Properties = props
	}

	// Extract defined names
	if e.options.IncludeDefinedNames {
		metadata.DefinedNames = e.extractDefinedNames()
	}

	// Extract sheet metadata
	sheets := e.file.GetSheetList()
	for idx, sheetName := range sheets {
		sheetMeta, err := e.extractSheetMetadata(idx, sheetName)
		if err != nil {
			continue
		}
		metadata.Sheets = append(metadata.Sheets, sheetMeta)
	}

	// Extract unique styles if requested
	if e.options.IncludeStyles {
		metadata.Styles = e.extractUniqueStyles()
	}

	return metadata, nil
}

// ExtractToJSON extracts metadata and returns it as JSON string
func (e *Extractor) ExtractToJSON(pretty bool) (string, error) {
	metadata, err := e.Extract()
	if err != nil {
		return "", err
	}

	var jsonData []byte
	if pretty {
		jsonData, err = json.MarshalIndent(metadata, "", "  ")
	} else {
		jsonData, err = json.Marshal(metadata)
	}

	if err != nil {
		return "", fmt.Errorf("failed to marshal to JSON: %w", err)
	}

	return string(jsonData), nil
}

// ExtractToGO extracts metadata and returns it as GO string
func (e *Extractor) ExtractToGO() (string, error) {
	metadata, err := e.Extract()
	if err != nil {
		return "", err
	}

	// Helper function to marshal Go values as Go code
	var marshalGo func(v interface{}, indent string) string
	marshalGo = func(v interface{}, indent string) string {
		switch val := v.(type) {
		case string:
			return fmt.Sprintf("%q", val)
		case time.Time:
			return fmt.Sprintf("time.Date(%d, %d, %d, %d, %d, %d, %d, time.UTC)", val.Year(), val.Month(), val.Day(), val.Hour(), val.Minute(), val.Second(), val.Nanosecond())
		case []byte:
			return fmt.Sprintf("%#v", val)
		case nil:
			return "nil"
		case bool:
			return fmt.Sprintf("%v", val)
		case int:
			return fmt.Sprintf("%d", val)
		case float64:
			return fmt.Sprintf("%v", val)
		case *string:
			if val == nil {
				return "nil"
			}
			return fmt.Sprintf("%v", *val)
		case *bool:
			if val == nil {
				return "nil"
			}
			return fmt.Sprintf("%v", *val)
		case map[int]float64:
			if len(val) == 0 {
				return "nil"
			}
			s := "map[int]float64{"
			for k, v := range val {
				s += fmt.Sprintf("%d: %v, ", k, v)
			}
			s += "}"
			return s
		case map[string]float64:
			if len(val) == 0 {
				return "nil"
			}
			s := "map[string]float64{"
			for k, v := range val {
				s += fmt.Sprintf("%q: %v, ", k, v)
			}
			s += "}"
			return s
		case map[int]StyleDetails:
			if len(val) == 0 {
				return "nil"
			}
			s := "map[int]excelmetadata.StyleDetails{\n"
			for k, v := range val {
				s += fmt.Sprintf("%s%d: %s,\n", indent+"  ", k, marshalGo(v, indent+"  "))
			}
			s += indent + "}"
			return s
		case []string:
			if len(val) == 0 {
				return "nil"
			}
			s := "[]string{"
			for _, v := range val {
				s += fmt.Sprintf("%q, ", v)
			}
			s += "}"
			return s
		case []int:
			if len(val) == 0 {
				return "nil"
			}
			s := "[]int{"
			for _, v := range val {
				s += fmt.Sprintf("%d, ", v)
			}
			s += "}"
			return s
		case []SheetMetadata:
			if len(val) == 0 {
				return "nil"
			}
			s := "[]excelmetadata.SheetMetadata{\n"
			for _, v := range val {
				s += indent + "  " + marshalGo(v, indent+"  ") + ",\n"
			}
			s += indent + "}"
			return s
		case []MergedCell:
			if len(val) == 0 {
				return "nil"
			}
			s := "[]excelmetadata.MergedCell{\n"
			for _, v := range val {
				s += indent + "  " + marshalGo(v, indent+"  ") + ",\n"
			}
			s += indent + "}"
			return s
		case []DataValidation:
			if len(val) == 0 {
				return "nil"
			}
			s := "[]excelmetadata.DataValidation{\n"
			for _, v := range val {
				s += indent + "  " + marshalGo(v, indent+"  ") + ",\n"
			}
			s += indent + "}"
			return s
		case []CellMetadata:
			if len(val) == 0 {
				return "nil"
			}
			s := "[]excelmetadata.CellMetadata{\n"
			for _, v := range val {
				s += indent + "  " + marshalGo(v, indent+"  ") + ",\n"
			}
			s += indent + "}"
			return s
		case []ImageMetadata:
			if len(val) == 0 {
				return "nil"
			}
			s := "[]excelmetadata.ImageMetadata{\n"
			for _, v := range val {
				s += indent + "  " + marshalGo(v, indent+"  ") + ",\n"
			}
			s += indent + "}"
			return s
		case []DefinedName:
			if len(val) == 0 {
				return "nil"
			}
			s := "[]excelmetadata.DefinedName{\n"
			for _, v := range val {
				s += indent + "  " + marshalGo(v, indent+"  ") + ",\n"
			}
			s += indent + "}"
			return s
		case *SheetProtection:
			if val == nil {
				return "nil"
			}
			return "&" + marshalGo(*val, indent)
		case *FontStyle:
			if val == nil {
				return "nil"
			}
			return "&" + marshalGo(*val, indent)
		case *FillStyle:
			if val == nil {
				return "nil"
			}
			return "&" + marshalGo(*val, indent)
		case *AlignmentStyle:
			if val == nil {
				return "nil"
			}
			return "&" + marshalGo(*val, indent)
		case *Protection:
			if val == nil {
				return "nil"
			}
			return "&" + marshalGo(*val, indent)
		case *Hyperlink:
			if val == nil {
				return "nil"
			}
			return "&" + marshalGo(*val, indent)
		case *ImageFormat:
			if val == nil {
				return "nil"
			}
			return "&" + marshalGo(*val, indent)
		case StyleDetails:
			s := "excelmetadata.StyleDetails{\n"
			s += indent + "  Font: " + marshalGo(val.Font, indent+"  ") + ",\n"
			s += indent + "  Fill: " + marshalGo(val.Fill, indent+"  ") + ",\n"
			s += indent + "  Border: " + marshalGo(val.Border, indent+"  ") + ",\n"
			s += indent + "  Alignment: " + marshalGo(val.Alignment, indent+"  ") + ",\n"
			s += indent + "  NumberFormat: " + marshalGo(val.NumberFormat, indent+"  ") + ",\n"
			s += indent + "  Protection: " + marshalGo(val.Protection, indent+"  ") + ",\n"
			s += indent + "}"
			return s
		case FontStyle:
			s := "excelmetadata.FontStyle{\n"
			s += indent + "  Bold: " + marshalGo(val.Bold, indent+"  ") + ",\n"
			s += indent + "  Italic: " + marshalGo(val.Italic, indent+"  ") + ",\n"
			s += indent + "  Underline: " + marshalGo(val.Underline, indent+"  ") + ",\n"
			s += indent + "  Strike: " + marshalGo(val.Strike, indent+"  ") + ",\n"
			s += indent + "  Family: " + marshalGo(val.Family, indent+"  ") + ",\n"
			s += indent + "  Size: " + marshalGo(val.Size, indent+"  ") + ",\n"
			s += indent + "  Color: " + marshalGo(val.Color, indent+"  ") + ",\n"
			s += indent + "}"
			return s
		case FillStyle:
			s := "excelmetadata.FillStyle{\n"
			s += indent + "  Type: " + marshalGo(val.Type, indent+"  ") + ",\n"
			s += indent + "  Pattern: " + marshalGo(val.Pattern, indent+"  ") + ",\n"
			s += indent + "  Color: " + marshalGo(val.Color, indent+"  ") + ",\n"
			s += indent + "}"
			return s
		case []BorderStyle:
			if len(val) == 0 {
				return "nil"
			}
			s := "[]excelmetadata.BorderStyle{\n"
			for _, v := range val {
				s += indent + "  " + marshalGo(v, indent+"  ") + ",\n"
			}
			s += indent + "}"
			return s
		case BorderStyle:
			s := "excelmetadata.BorderStyle{\n"
			s += indent + "  Type: " + marshalGo(val.Type, indent+"  ") + ",\n"
			s += indent + "  Color: " + marshalGo(val.Color, indent+"  ") + ",\n"
			s += indent + "  Style: " + marshalGo(val.Style, indent+"  ") + ",\n"
			s += indent + "}"
			return s
		case AlignmentStyle:
			s := "excelmetadata.AlignmentStyle{\n"
			s += indent + "  Horizontal: " + marshalGo(val.Horizontal, indent+"  ") + ",\n"
			s += indent + "  Vertical: " + marshalGo(val.Vertical, indent+"  ") + ",\n"
			s += indent + "  WrapText: " + marshalGo(val.WrapText, indent+"  ") + ",\n"
			s += indent + "  TextRotation: " + marshalGo(val.TextRotation, indent+"  ") + ",\n"
			s += indent + "  Indent: " + marshalGo(val.Indent, indent+"  ") + ",\n"
			s += indent + "  ShrinkToFit: " + marshalGo(val.ShrinkToFit, indent+"  ") + ",\n"
			s += indent + "}"
			return s
		case Protection:
			s := "excelmetadata.Protection{\n"
			s += indent + "  Hidden: " + marshalGo(val.Hidden, indent+"  ") + ",\n"
			s += indent + "  Locked: " + marshalGo(val.Locked, indent+"  ") + ",\n"
			s += indent + "}"
			return s
		case SheetMetadata:
			s := "excelmetadata.SheetMetadata{\n"
			s += indent + "  Index: " + marshalGo(val.Index, indent+"  ") + ",\n"
			s += indent + "  Name: " + marshalGo(val.Name, indent+"  ") + ",\n"
			s += indent + "  Visible: " + marshalGo(val.Visible, indent+"  ") + ",\n"
			s += indent + "  Dimensions: " + marshalGo(val.Dimensions, indent+"  ") + ",\n"
			s += indent + "  MergedCells: " + marshalGo(val.MergedCells, indent+"  ") + ",\n"
			s += indent + "  DataValidations: " + marshalGo(val.DataValidations, indent+"  ") + ",\n"
			s += indent + "  Protection: " + marshalGo(val.Protection, indent+"  ") + ",\n"
			s += indent + "  RowHeights: " + marshalGo(val.RowHeights, indent+"  ") + ",\n"
			s += indent + "  ColWidths: " + marshalGo(val.ColWidths, indent+"  ") + ",\n"
			s += indent + "  Cells: " + marshalGo(val.Cells, indent+"  ") + ",\n"
			s += indent + "  Images: " + marshalGo(val.Images, indent+"  ") + ",\n"
			s += indent + "}"
			return s
		case SheetDimensions:
			s := "excelmetadata.SheetDimensions{\n"
			s += indent + "  StartCell: " + marshalGo(val.StartCell, indent+"  ") + ",\n"
			s += indent + "  EndCell: " + marshalGo(val.EndCell, indent+"  ") + ",\n"
			s += indent + "  RowCount: " + marshalGo(val.RowCount, indent+"  ") + ",\n"
			s += indent + "  ColCount: " + marshalGo(val.ColCount, indent+"  ") + ",\n"
			s += indent + "}"
			return s
		case MergedCell:
			s := "excelmetadata.MergedCell{\n"
			s += indent + "  StartCell: " + marshalGo(val.StartCell, indent+"  ") + ",\n"
			s += indent + "  EndCell: " + marshalGo(val.EndCell, indent+"  ") + ",\n"
			s += indent + "  Value: " + marshalGo(val.Value, indent+"  ") + ",\n"
			s += indent + "}"
			return s
		case DataValidation:
			s := "excelmetadata.DataValidation{\n"
			s += indent + "  Range: " + marshalGo(val.Range, indent+"  ") + ",\n"
			s += indent + "  Type: " + marshalGo(val.Type, indent+"  ") + ",\n"
			s += indent + "  Operator: " + marshalGo(val.Operator, indent+"  ") + ",\n"
			s += indent + "  Formula1: " + marshalGo(val.Formula1, indent+"  ") + ",\n"
			s += indent + "  Formula2: " + marshalGo(val.Formula2, indent+"  ") + ",\n"
			s += indent + "  ShowError: " + marshalGo(val.ShowError, indent+"  ") + ",\n"
			s += indent + "  ErrorTitle: " + marshalGo(val.ErrorTitle, indent+"  ") + ",\n"
			s += indent + "  ErrorMessage: " + marshalGo(val.ErrorMessage, indent+"  ") + ",\n"
			s += indent + "}"
			return s
		case SheetProtection:
			s := "excelmetadata.SheetProtection{\n"
			s += indent + "  Protected: " + marshalGo(val.Protected, indent+"  ") + ",\n"
			s += indent + "  Password: " + marshalGo(val.Password, indent+"  ") + ",\n"
			s += indent + "  EditObjects: " + marshalGo(val.EditObjects, indent+"  ") + ",\n"
			s += indent + "  EditScenarios: " + marshalGo(val.EditScenarios, indent+"  ") + ",\n"
			s += indent + "  SelectLockedCells: " + marshalGo(val.SelectLockedCells, indent+"  ") + ",\n"
			s += indent + "  SelectUnlockedCells: " + marshalGo(val.SelectUnlockedCells, indent+"  ") + ",\n"
			s += indent + "}"
			return s
		case CellMetadata:
			s := "excelmetadata.CellMetadata{\n"
			s += indent + "  Address: " + marshalGo(val.Address, indent+"  ") + ",\n"
			s += indent + "  Value: " + marshalGo(val.Value, indent+"  ") + ",\n"
			s += indent + "  Formula: " + marshalGo(val.Formula, indent+"  ") + ",\n"
			s += indent + "  StyleID: " + marshalGo(val.StyleID, indent+"  ") + ",\n"
			s += indent + "  Type: " + strings.ReplaceAll(fmt.Sprintf("excelize.CellType('%q')", string(val.Type)), "\"", "") + ",\n"
			s += indent + "  Hyperlink: " + marshalGo(val.Hyperlink, indent+"  ") + ",\n"
			s += indent + "}"
			return s
		case Hyperlink:
			s := "excelmetadata.Hyperlink{\n"
			s += indent + "  Link: " + marshalGo(val.Link, indent+"  ") + ",\n"
			s += indent + "}"
			return s
		case ImageMetadata:
			s := "excelmetadata.ImageMetadata{\n"
			s += indent + "  Cell: " + marshalGo(val.Cell, indent+"  ") + ",\n"
			s += indent + "  File: " + marshalGo(val.File, indent+"  ") + ",\n"
			s += indent + "  Extension: " + marshalGo(val.Extension, indent+"  ") + ",\n"
			s += indent + "  InsertType: " + fmt.Sprintf("%#v", val.InsertType) + ",\n"
			s += indent + "  Format: " + marshalGo(val.Format, indent+"  ") + ",\n"
			s += indent + "}"
			return s
		case ImageFormat:
			s := "excelmetadata.ImageFormat{\n"
			s += indent + "  AltText: " + marshalGo(val.AltText, indent+"  ") + ",\n"
			s += indent + "  PrintObject: " + marshalGo(val.PrintObject, indent+"  ") + ",\n"
			s += indent + "  Locked: " + marshalGo(val.Locked, indent+"  ") + ",\n"
			s += indent + "  LockAspectRatio: " + marshalGo(val.LockAspectRatio, indent+"  ") + ",\n"
			s += indent + "  AutoFit: " + marshalGo(val.AutoFit, indent+"  ") + ",\n"
			s += indent + "  AutoFitIgnoreAspect: " + marshalGo(val.AutoFitIgnoreAspect, indent+"  ") + ",\n"
			s += indent + "  OffsetX: " + marshalGo(val.OffsetX, indent+"  ") + ",\n"
			s += indent + "  OffsetY: " + marshalGo(val.OffsetY, indent+"  ") + ",\n"
			s += indent + "  ScaleX: " + marshalGo(val.ScaleX, indent+"  ") + ",\n"
			s += indent + "  ScaleY: " + marshalGo(val.ScaleY, indent+"  ") + ",\n"
			s += indent + "  Hyperlink: " + marshalGo(val.Hyperlink, indent+"  ") + ",\n"
			s += indent + "  HyperlinkType: " + marshalGo(val.HyperlinkType, indent+"  ") + ",\n"
			s += indent + "  Positioning: " + marshalGo(val.Positioning, indent+"  ") + ",\n"
			s += indent + "}"
			return s
		case DefinedName:
			s := "excelmetadata.DefinedName{\n"
			s += indent + "  Name: " + marshalGo(val.Name, indent+"  ") + ",\n"
			s += indent + "  RefersTo: " + marshalGo(val.RefersTo, indent+"  ") + ",\n"
			s += indent + "  Scope: " + marshalGo(val.Scope, indent+"  ") + ",\n"
			s += indent + "}"
			return s
		case DocumentProperties:
			s := "excelmetadata.DocumentProperties{\n"
			s += indent + "  Title: " + marshalGo(val.Title, indent+"  ") + ",\n"
			s += indent + "  Subject: " + marshalGo(val.Subject, indent+"  ") + ",\n"
			s += indent + "  Creator: " + marshalGo(val.Creator, indent+"  ") + ",\n"
			s += indent + "  Keywords: " + marshalGo(val.Keywords, indent+"  ") + ",\n"
			s += indent + "  Description: " + marshalGo(val.Description, indent+"  ") + ",\n"
			s += indent + "  LastModifiedBy: " + marshalGo(val.LastModifiedBy, indent+"  ") + ",\n"
			s += indent + "  Category: " + marshalGo(val.Category, indent+"  ") + ",\n"
			s += indent + "  Version: " + marshalGo(val.Version, indent+"  ") + ",\n"
			s += indent + "  Created: " + marshalGo(val.Created, indent+"  ") + ",\n"
			s += indent + "  Modified: " + marshalGo(val.Modified, indent+"  ") + ",\n"
			s += indent + "}"
			return s
		case Metadata:
			s := "excelmetadata.Metadata{\n"
			s += indent + "  Filename: " + marshalGo(val.Filename, indent+"  ") + ",\n"
			s += indent + "  Properties: " + marshalGo(val.Properties, indent+"  ") + ",\n"
			s += indent + "  Sheets: " + marshalGo(val.Sheets, indent+"  ") + ",\n"
			s += indent + "  DefinedNames: " + marshalGo(val.DefinedNames, indent+"  ") + ",\n"
			s += indent + "  Styles: " + marshalGo(val.Styles, indent+"  ") + ",\n"
			s += indent + "  ExtractedAt: " + marshalGo(val.ExtractedAt, indent+"  ") + ",\n"
			s += indent + "}"
			return s
		default:
			return fmt.Sprintf("%#v", v)
		}
	}

	goStr := fmt.Sprintf(`package main

import (
	"github.com/prongbang/excelmetadata"
	"github.com/prongbang/excelrecreator"
	"github.com/xuri/excelize/v2"
)

func main() {
	f := excelize.NewFile()

	metadata := &%s

	reCreator := &excelrecreator.Recreator{
		File:     f,
		Metadata: metadata,
		Options:  excelrecreator.DefaultOptions(),
		StyleMap: make(map[int]int),
	}
	_ = reCreator.Recreate()

	_ = f.SaveAs("sample.clone.xlsx")
}`,
		marshalGo(*metadata, ""),
	)

	return goStr, nil
}

// ExtractToFile extracts metadata and saves it to a JSON or GO file
func (e *Extractor) ExtractToFile(outputPath string, pretty bool) error {
	ext := path.Ext(outputPath)
	var data []byte
	if ext == ".json" {
		jsonStr, err := e.ExtractToJSON(pretty)
		if err != nil {
			return err
		}
		data = []byte(jsonStr)
	} else if ext == ".go" {
		goStr, err := e.ExtractToGO()
		if err != nil {
			return err
		}
		data = []byte(goStr)
	} else {
		return errors.New(fmt.Sprintf("unsupported %s file", ext))
	}

	dir := filepath.Dir(outputPath)
	if err := os.MkdirAll(dir, 0755); err != nil {
		return err
	}

	return os.WriteFile(outputPath, data, 0644)
}

// Close closes the underlying Excel file
func (e *Extractor) Close() error {
	return e.file.Close()
}

// Private extraction methods

func (e *Extractor) extractDocumentProperties() (DocumentProperties, error) {
	props, err := e.file.GetDocProps()
	if err != nil {
		return DocumentProperties{}, err
	}

	return DocumentProperties{
		Title:          props.Title,
		Subject:        props.Subject,
		Creator:        props.Creator,
		Keywords:       props.Keywords,
		Description:    props.Description,
		LastModifiedBy: props.LastModifiedBy,
		Category:       props.Category,
		Version:        props.Version,
		Created:        props.Created,
		Modified:       props.Modified,
	}, nil
}

func (e *Extractor) extractDefinedNames() []DefinedName {
	var names []DefinedName

	// GetDefinedName returns []DefinedName
	definedNames := e.file.GetDefinedName()
	for _, dn := range definedNames {
		names = append(names, DefinedName{
			Name:     dn.Name,
			RefersTo: dn.RefersTo,
			Scope:    dn.Scope,
		})
	}

	return names
}

func (e *Extractor) extractSheetMetadata(index int, sheetName string) (SheetMetadata, error) {
	visible, _ := e.file.GetSheetVisible(sheetName)
	sheet := SheetMetadata{
		Index:   index,
		Name:    sheetName,
		Visible: visible,
	}

	// Get sheet dimensions
	if dimensions, err := e.getSheetDimensions(sheetName); err == nil {
		sheet.Dimensions = dimensions
	}

	// Extract merged cells
	if mergedCells, err := e.file.GetMergeCells(sheetName); err == nil {
		for _, mc := range mergedCells {
			sheet.MergedCells = append(sheet.MergedCells, MergedCell{
				StartCell: mc.GetStartAxis(),
				EndCell:   mc.GetEndAxis(),
				Value:     mc.GetCellValue(),
			})
		}
	}

	// Extract data validations
	if e.options.IncludeDataValidation {
		// GetDataValidations returns ([]*DataValidation, error)
		if dvs, err := e.file.GetDataValidations(sheetName); err == nil {
			for _, dv := range dvs {
				sheet.DataValidations = append(sheet.DataValidations, DataValidation{
					Range:        dv.Sqref,
					Type:         dv.Type,
					Operator:     dv.Operator,
					Formula1:     dv.Formula1,
					Formula2:     dv.Formula2,
					ShowError:    dv.ShowErrorMessage,
					ErrorTitle:   dv.ErrorTitle,
					ErrorMessage: dv.Error,
				})
			}
		}
	}

	// Extract row heights and column widths
	sheet.RowHeights = make(map[int]float64)
	sheet.ColWidths = make(map[string]float64)

	// Get column widths
	cols, _ := e.file.GetCols(sheetName)
	for idx := range cols {
		col, _ := excelize.ColumnNumberToName(idx + 1)
		width, _ := e.file.GetColWidth(sheetName, col)
		if width != 9.140625 { // default width
			sheet.ColWidths[col] = width
		}
	}

	// Extract cell data
	if e.options.IncludeCellData {
		cells, err := e.extractCellData(sheetName)
		if err == nil {
			sheet.Cells = cells
		}
	}

	// Extract images
	if e.options.IncludeImages {
		sheet.Images = e.extractImages(sheetName)
	}

	return sheet, nil
}

func (e *Extractor) getSheetDimensions(sheetName string) (SheetDimensions, error) {
	rows, err := e.file.GetRows(sheetName)
	if err != nil {
		return SheetDimensions{}, err
	}

	rowCount := len(rows)
	maxColCount := 0

	for _, row := range rows {
		if len(row) > maxColCount {
			maxColCount = len(row)
		}
	}

	if rowCount == 0 || maxColCount == 0 {
		return SheetDimensions{
			StartCell: "A1",
			EndCell:   "A1",
			RowCount:  0,
			ColCount:  0,
		}, nil
	}

	startCell := "A1"
	endCol, _ := excelize.ColumnNumberToName(maxColCount)
	endCell := fmt.Sprintf("%s%d", endCol, rowCount)

	return SheetDimensions{
		StartCell: startCell,
		EndCell:   endCell,
		RowCount:  rowCount,
		ColCount:  maxColCount,
	}, nil
}

func (e *Extractor) extractCellData(sheetName string) ([]CellMetadata, error) {
	var cells []CellMetadata
	cellCount := 0

	rows, err := e.file.GetRows(sheetName)
	if err != nil {
		return nil, err
	}

	for rowIdx, row := range rows {
		for colIdx, value := range row {
			if value == "" {
				continue
			}

			if e.options.MaxCellsPerSheet > 0 && cellCount >= e.options.MaxCellsPerSheet {
				return cells, nil
			}

			col, _ := excelize.ColumnNumberToName(colIdx + 1)
			cellAddr := fmt.Sprintf("%s%d", col, rowIdx+1)

			cellMeta := CellMetadata{
				Address: cellAddr,
				Value:   value,
			}

			// Get formula
			if formula, err := e.file.GetCellFormula(sheetName, cellAddr); err == nil && formula != "" {
				cellMeta.Formula = formula
			}

			// Get style ID
			if styleID, err := e.file.GetCellStyle(sheetName, cellAddr); err == nil {
				cellMeta.StyleID = styleID
			}

			// Get cell type
			if cellType, err := e.file.GetCellType(sheetName, cellAddr); err == nil {
				cellMeta.Type = cellType
			}

			// Get hyperlink - GetCellHyperLink returns (HyperlinkOpts, string, error)
			if link, target, err := e.file.GetCellHyperLink(sheetName, cellAddr); err == nil && link {
				cellMeta.Hyperlink = &Hyperlink{
					Link: target,
				}
			}

			cells = append(cells, cellMeta)
			cellCount++
		}
	}

	return cells, nil
}

func (e *Extractor) extractImages(sheetName string) []ImageMetadata {
	var images []ImageMetadata

	cellAddress, err := e.file.GetPictureCells(sheetName)
	if err != nil {
		return images
	}
	for _, cellAddr := range cellAddress {
		// GetPictures returns ([]Picture, error)
		pictures, err := e.file.GetPictures(sheetName, cellAddr)
		if err != nil {
			continue
		}

		// Use index to avoid issues with range variable
		for _, picture := range pictures {
			img := ImageMetadata{
				Cell:       cellAddr,
				File:       picture.File,
				Extension:  picture.Extension,
				InsertType: byte(picture.InsertType),
			}
			if picture.Format != nil {
				img.Format = &ImageFormat{
					AltText:             picture.Format.AltText,
					PrintObject:         picture.Format.PrintObject,
					Locked:              picture.Format.Locked,
					LockAspectRatio:     picture.Format.LockAspectRatio,
					AutoFit:             picture.Format.AutoFit,
					AutoFitIgnoreAspect: picture.Format.AutoFitIgnoreAspect,
					OffsetX:             picture.Format.OffsetX,
					OffsetY:             picture.Format.OffsetY,
					ScaleX:              picture.Format.ScaleX,
					ScaleY:              picture.Format.ScaleY,
					Hyperlink:           picture.Format.Hyperlink,
					HyperlinkType:       picture.Format.HyperlinkType,
					Positioning:         picture.Format.Positioning,
				}
			}
			images = append(images, img)
		}
	}

	return images
}

func (e *Extractor) extractUniqueStyles() map[int]StyleDetails {
	styles := make(map[int]StyleDetails)
	processedStyles := make(map[int]bool)

	for _, sheetName := range e.file.GetSheetList() {
		rows, err := e.file.GetRows(sheetName)
		if err != nil {
			continue
		}

		for rowIdx, row := range rows {
			for colIdx := range row {
				col, _ := excelize.ColumnNumberToName(colIdx + 1)
				cellAddr := fmt.Sprintf("%s%d", col, rowIdx+1)

				if styleID, err := e.file.GetCellStyle(sheetName, cellAddr); err == nil && styleID != 0 {
					if !processedStyles[styleID] {
						if style, err := e.extractStyleDetails(styleID); err == nil {
							styles[styleID] = style
							processedStyles[styleID] = true
						}
					}
				}
			}
		}
	}

	return styles
}

func (e *Extractor) extractStyleDetails(styleID int) (StyleDetails, error) {
	// GetStyle returns (*Style, error)
	style, err := e.file.GetStyle(styleID)
	if err != nil {
		return StyleDetails{}, err
	}

	details := StyleDetails{
		NumberFormat: style.NumFmt,
	}

	// Extract font details
	if style.Font != nil {
		details.Font = &FontStyle{
			Bold:      style.Font.Bold,
			Italic:    style.Font.Italic,
			Underline: style.Font.Underline,
			Strike:    style.Font.Strike,
			Family:    style.Font.Family,
			Size:      style.Font.Size,
			Color:     style.Font.Color,
		}
	}

	// Extract fill details
	if len(style.Fill.Color) > 0 {
		details.Fill = &FillStyle{
			Type:    style.Fill.Type,
			Pattern: style.Fill.Pattern,
			Color:   style.Fill.Color,
		}
	}

	// Extract border details
	if len(style.Border) > 0 {
		for _, border := range style.Border {
			details.Border = append(details.Border, BorderStyle{
				Type:  border.Type,
				Color: border.Color,
				Style: border.Style,
			})
		}
	}

	// Extract alignment details
	if style.Alignment != nil {
		details.Alignment = &AlignmentStyle{
			Horizontal:   style.Alignment.Horizontal,
			Vertical:     style.Alignment.Vertical,
			WrapText:     style.Alignment.WrapText,
			TextRotation: style.Alignment.TextRotation,
			Indent:       style.Alignment.Indent,
			ShrinkToFit:  style.Alignment.ShrinkToFit,
		}
	}

	// Extract protection details
	if style.Protection != nil {
		details.Protection = &Protection{
			Hidden: style.Protection.Hidden,
			Locked: style.Protection.Locked,
		}
	}

	return details, nil
}

// Utility functions

// QuickExtract is a convenience function for simple extraction
func QuickExtract(filename string) (*Metadata, error) {
	extractor, err := New(filename, DefaultOptions())
	if err != nil {
		return nil, err
	}
	defer func(extractor *Extractor) {
		_ = extractor.Close()
	}(extractor)

	return extractor.Extract()
}

// QuickExtractToJSON is a convenience function for extracting to JSON
func QuickExtractToJSON(filename string, pretty bool) (string, error) {
	extractor, err := New(filename, DefaultOptions())
	if err != nil {
		return "", err
	}
	defer func(extractor *Extractor) {
		_ = extractor.Close()
	}(extractor)

	return extractor.ExtractToJSON(pretty)
}

// QuickExtractToGO is a convenience function for extracting to GO
func QuickExtractToGO(filename string) (string, error) {
	extractor, err := New(filename, DefaultOptions())
	if err != nil {
		return "", err
	}
	defer func(extractor *Extractor) {
		_ = extractor.Close()
	}(extractor)

	return extractor.ExtractToGO()
}

// QuickExtractToFile is a convenience function for extracting to a JSON file
func QuickExtractToFile(excelFile, file string, pretty bool) error {
	extractor, err := New(excelFile, DefaultOptions())
	if err != nil {
		return err
	}
	defer func(extractor *Extractor) {
		_ = extractor.Close()
	}(extractor)

	return extractor.ExtractToFile(file, pretty)
}
