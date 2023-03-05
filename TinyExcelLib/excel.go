package tinyexcellib

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"strings"
)

type ExcelFileHandler struct {
	excelFilePath       string
	workbook            internalFile
	sheets              map[string]internalFile
	sharedStrings       internalFile
	otherFiles          []internalFile
	sharedStringsLoaded ExcelXML_sharedstrings__
}

type internalFile struct {
	isChanged bool
	structure interface{}
	zipFile   *zip.File
	content   *[]byte
}

type Cell struct {
	Value interface{}
}

func ReadExcelFile(path string) (*ExcelFileHandler, error) {
	archive, err := zip.OpenReader(path)
	if err != nil {
		return nil, err
	}

	var workbook, sharedStrings internalFile
	sheets := make(map[string]internalFile, 0)

	otherFiles := make([]internalFile, 0)

	for _, f := range archive.File {

		// Get sheets: use the actual file name as
		// a temporary name
		if strings.HasPrefix(f.Name, "xl/worksheets/sheet") {
			sheets[f.Name] = internalFile{zipFile: f}
		} else

		// Get workbook
		if f.Name == "xl/workbook.xml" {
			workbook = internalFile{zipFile: f}
		} else

		// Get shared strings file
		if f.Name == "xl/sharedStrings.xml" {
			sharedStrings = internalFile{zipFile: f}
		} else

		// Everything else
		{
			otherFiles = append(otherFiles, internalFile{zipFile: f})
		}
	}

	efh := &ExcelFileHandler{
		excelFilePath: path,
		workbook:      workbook,
		sheets:        sheets,
		sharedStrings: sharedStrings,
		otherFiles:    otherFiles,
	}

	if err = efh.readContent(); err != nil {
		return nil, err
	}

	if err = efh.renameSheets(); err != nil {
		return nil, err
	}

	if err = efh.loadSharedStrings(); err != nil {
		return nil, err
	}

	return efh, archive.Close()
}

func (efh *ExcelFileHandler) readContent() error {
	var err error
	var n, tot int
	var r io.ReadCloser

	// read workbook
	r, err = efh.workbook.zipFile.Open()
	if err != nil {
		return err
	}

	workbookContent := make([]byte, efh.workbook.zipFile.UncompressedSize64)
	tot = 0
	for tot < len(workbookContent) {
		if n, err = r.Read(workbookContent[tot:]); err != nil {
			return err
		}
		tot += n
	}

	if err = r.Close(); err != nil {
		return err
	}
	efh.workbook.content = &workbookContent

	// read sheets
	for name, f := range efh.sheets {
		r, err = f.zipFile.Open()
		if err != nil {
			return err
		}

		sheetContent := make([]byte, f.zipFile.UncompressedSize64)
		tot = 0
		for tot < len(sheetContent) {
			if n, err = r.Read(sheetContent[tot:]); err != nil {
				return err
			}
			tot += n
		}

		if err = r.Close(); err != nil {
			return err
		}

		efh.sheets[name] = internalFile{isChanged: false, zipFile: f.zipFile, content: &sheetContent}
	}

	// read shared strings
	// but the shared strings file might not be present
	if efh.sharedStrings.zipFile != nil {
		r, err = efh.sharedStrings.zipFile.Open()
		if err != nil {
			return err
		}

		sharedStringsContent := make([]byte, efh.sharedStrings.zipFile.UncompressedSize64)
		tot = 0
		for tot < len(sharedStringsContent) {
			if n, err = r.Read(sharedStringsContent[tot:]); err != nil {
				return err
			}
			tot += n
		}

		if err = r.Close(); err != nil {
			return err
		}
		efh.sharedStrings.content = &sharedStringsContent
	}

	// Other files
	for i, f := range efh.otherFiles {
		r, err = f.zipFile.Open()
		if err != nil {
			return err
		}

		fileContent := make([]byte, f.zipFile.UncompressedSize64)
		tot = 0
		for tot < len(fileContent) {
			if n, err = r.Read(fileContent[tot:]); err != nil {
				return err
			}
			tot += n
		}

		if err = r.Close(); err != nil {
			return err
		}

		efh.otherFiles[i] = internalFile{isChanged: false, zipFile: f.zipFile, content: &fileContent}
	}

	return nil
}

func (efh *ExcelFileHandler) renameSheets() error {
	var wb ExcelXML_workbook__

	if err := xml.Unmarshal(*efh.workbook.content, &wb); err != nil {
		return err
	}

	// put the actual sheets' names here
	names := make(map[int]string)
	for _, s := range wb.Sheets.List {
		idx, err := strconv.Atoi(s.RId[3:])
		if err != nil {
			return nil
		}
		names[idx] = s.Name
	}

	// get the id of the sheet from the archive path
	newSheets := make(map[string]internalFile)
	for k, s := range efh.sheets {
		idx, err := strconv.Atoi(strings.Split(k[19:], ".")[0])
		if err != nil {
			return nil
		}

		newSheets[names[idx]] = internalFile{content: s.content, zipFile: s.zipFile}
	}

	efh.sheets = newSheets

	return nil
}

func (efh *ExcelFileHandler) loadSharedStrings() error {
	var ss ExcelXML_sharedstrings__

	if efh.sharedStrings.zipFile != nil {
		if err := xml.Unmarshal(*efh.sharedStrings.content, &ss); err != nil {
			return err
		}
		efh.sharedStringsLoaded = ss
	} else

	// initialize shared strings file
	{
		efh.sharedStringsLoaded.XMLName = xml.Name{Local: "sst"}
		efh.sharedStringsLoaded.Xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
		efh.sharedStringsLoaded.Si = make([]ExcelXML_Si__, 0)
	}

	return nil
}

func (efh *ExcelFileHandler) addStringValue(value string) int {
	for idx, t := range efh.sharedStringsLoaded.Si {
		if t.T == value {
			efh.sharedStringsLoaded.Count += 1
			return idx
		}
	}

	// not found
	idx := len(efh.sharedStringsLoaded.Si)

	efh.sharedStringsLoaded.Count += 1
	efh.sharedStringsLoaded.UniqueCount += 1

	efh.sharedStringsLoaded.Si = append(efh.sharedStringsLoaded.Si, ExcelXML_Si__{T: value})

	return idx
}

// func (efh *ExcelFileHandler) GetSheetByName(sheetName string) {
// }

func (efh *ExcelFileHandler) InsertNewRow(sheetName string, index int, cells []Cell) error {

	var ws ExcelXML_worksheet_unmarshal__
	if (*efh).sheets[sheetName].isChanged {
		if s, ok := (*efh).sheets[sheetName].structure.(ExcelXML_worksheet_unmarshal__); ok {
			ws = s
		}
	} else {
		if err := xml.Unmarshal(*efh.sheets[sheetName].content, &ws); err != nil {
			panic(err)
		}
	}

	newRow := ExcelXML_row__{
		R:     index,
		Spans: fmt.Sprintf("%d:%d", 1, len(cells)), // TODO: I'm not sure about this
		Cols:  make([]ExcelXML_cell__, len(cells)),
	}

	var cellFinalValue, cellFinalType string
	for j, c := range cells {

		switch v := c.Value.(type) {
		case int:
			cellFinalValue = fmt.Sprintf("%d", v)
			cellFinalType = ""
		case float64:
			cellFinalValue = fmt.Sprintf("%f", v)
			cellFinalType = ""
		case string:
			cellFinalValue = fmt.Sprintf("%d", efh.addStringValue(v))
			cellFinalType = "s"
		}

		newRow.Cols[j] = ExcelXML_cell__{
			R:     zeroBasedIndecesToExcel(index-1, j),
			Value: cellFinalValue,
			T:     cellFinalType,
		}
	}

	// locate the index of the row in the list
	// of the sheet data rows
	innerRowIndex := -1
	for i, row := range ws.SheetData.List {
		if row.R >= index {
			innerRowIndex = i
			break
		}
	}

	if innerRowIndex == -1 {
		innerRowIndex = len(ws.SheetData.List) - 1
	}

	startRows := ws.SheetData.List[0:innerRowIndex]
	endRows := ws.SheetData.List[innerRowIndex:]

	// update endRows
	for i, row := range endRows {
		endRows[i].R = endRows[i].R + 1
		for j, c := range row.Cols {
			rowIdx, colIdx := excelToZeroBasedIndeces(c.R)
			endRows[i].Cols[j].R = zeroBasedIndecesToExcel(rowIdx+1, colIdx)
		}
	}

	// append new row
	startRows = append(startRows, newRow)
	ws.SheetData.List = append(startRows, endRows...)

	// Update sheet dimension attribute
	originalDim := strings.Split(ws.Dimension.Ref, ":")

	bottomRighDimUpdated := getMaxExcelIndex(
		originalDim[1],
		ws.SheetData.List[len(ws.SheetData.List)-1].Cols[len(ws.SheetData.List[len(ws.SheetData.List)-1].Cols)-1].R,
	)

	ws.Dimension.Ref = fmt.Sprintf("%s:%s", originalDim[0], bottomRighDimUpdated)

	efh.sheets[sheetName] = internalFile{
		isChanged: true,
		zipFile:   efh.sheets[sheetName].zipFile,
		content:   efh.sheets[sheetName].content,
		structure: ws,
	}

	return nil
}

func (efh *ExcelFileHandler) UpdateSheet(sheetName string) error {

	if s, ok := ((*efh).sheets)[sheetName].structure.(ExcelXML_worksheet_unmarshal__); ok {

		content, err := s.Marshal()
		if err != nil {
			panic(err)
		}

		efh.sheets[sheetName] = internalFile{
			isChanged: false,
			zipFile:   efh.sheets[sheetName].zipFile,
			content:   &content,
		}
	}

	return fmt.Errorf("UpdateSheet: sheet name not found")
}

// Get all the cells in a given work sheet.
// The values are provided as strings.
func (efh *ExcelFileHandler) ReadCells(sheetName string) ([][]string, error) {

	var ws ExcelXML_worksheet_unmarshal__
	if err := xml.Unmarshal(*efh.sheets[sheetName].content, &ws); err != nil {
		panic(err)
	}

	dimension := strings.Split(ws.Dimension.Ref, ":")

	// startRow, startCol := excelToZeroBasedIndeces(dimension[0])
	endRow, endCol := excelToZeroBasedIndeces(dimension[1])

	// Initialize the empty matrix
	var rowIdx, colIdx int
	cells := make([][]string, endRow+1)
	for i := range cells {
		cells[i] = make([]string, endCol+1)
	}

	for _, row := range ws.SheetData.List {
		for _, cell := range row.Cols {

			rowIdx, colIdx = excelToZeroBasedIndeces(cell.R)

			// string cell
			if cell.T == "s" {
				ssIndex, err := strconv.Atoi(cell.Value)
				if err != nil {
					return nil, err
				}
				cells[rowIdx][colIdx] = efh.sharedStringsLoaded.Si[ssIndex].T
			} else {
				cells[rowIdx][colIdx] = cell.Value
			}
		}
	}

	return cells, nil
}

func (efh *ExcelFileHandler) Save() error {
	return efh.SaveAs(efh.excelFilePath)
}

func (efh *ExcelFileHandler) SaveAs(path string) error {
	var zipWriter *zip.Writer
	var err error
	var f *os.File

	os.Remove(path)

	f, err = os.OpenFile(filepath.Clean(path), os.O_WRONLY|os.O_TRUNC|os.O_CREATE, os.ModePerm)
	if err != nil {
		panic(err)
	}
	defer f.Close()

	zipWriter = zip.NewWriter(f)
	defer zipWriter.Close()

	var buff []byte
	var w io.Writer

	// workbook
	// efh.workbook.zipFile.SetMode(os.ModePerm)
	w, err = zipWriter.Create(efh.workbook.zipFile.Name)
	if err != nil {
		panic(err)
	}
	_, err = w.Write(*efh.workbook.content)
	if err != nil {
		panic(err)
	}

	// worksheet
	for _, sheet := range efh.sheets {
		// sheet.zipFile.SetMode(os.ModePerm)
		w, err = zipWriter.Create(sheet.zipFile.Name)
		if err != nil {
			panic(err)
		}
		_, err = w.Write(*sheet.content)
		if err != nil {
			panic(err)
		}
	}

	// shared strings
	// efh.sharedStrings.zipFile.SetMode(os.ModePerm)
	w, err = zipWriter.Create(efh.sharedStrings.zipFile.Name)
	if err != nil {
		panic(err)
	}

	buff, err = xml.Marshal(efh.sharedStringsLoaded)
	if err != nil {
		panic(err)
	}
	_, err = w.Write(buff)
	if err != nil {
		panic(err)
	}

	// other files
	for _, o := range efh.otherFiles {
		// o.zipFile.SetMode(os.ModePerm)
		w, err = zipWriter.Create(o.zipFile.Name)
		if err != nil {
			panic(err)
		}
		_, err = w.Write(*o.content)
		if err != nil {
			panic(err)
		}
	}

	return nil
}
