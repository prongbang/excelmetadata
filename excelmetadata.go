package excelmetadata

import (
	"encoding/json"
	"fmt"
	"os"
	"time"

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

// ExtractToFile extracts metadata and saves it to a JSON file
func (e *Extractor) ExtractToFile(outputPath string, pretty bool) error {
	jsonStr, err := e.ExtractToJSON(pretty)
	if err != nil {
		return err
	}

	return os.WriteFile(outputPath, []byte(jsonStr), 0644)
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
	defer extractor.Close()

	return extractor.Extract()
}

// QuickExtractToJSON is a convenience function for extracting to JSON
func QuickExtractToJSON(filename string, pretty bool) (string, error) {
	extractor, err := New(filename, DefaultOptions())
	if err != nil {
		return "", err
	}
	defer extractor.Close()

	return extractor.ExtractToJSON(pretty)
}

// QuickExtractToFile is a convenience function for extracting to a JSON file
func QuickExtractToFile(excelFile, jsonFile string, pretty bool) error {
	extractor, err := New(excelFile, DefaultOptions())
	if err != nil {
		return err
	}
	defer extractor.Close()

	return extractor.ExtractToFile(jsonFile, pretty)
}
