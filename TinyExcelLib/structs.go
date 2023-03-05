package tinyexcellib

import "encoding/xml"

///////////////////////////////////////////////////////////////////////////////
///////////////////					WORKBOOK

type ExcelXML_workbook__ struct {
	XMLName xml.Name `xml:"workbook"`

	// FileVersion string `xml:"fileVersion"`
	// workbookPr string`xml:"workbookPr"`
	// revisionPtr string`xml:"xr:revisionPtr"`
	// BookViews string `xml:"bookViews,"`

	Sheets ExcelXML_sheets__ `xml:"sheets"`
}

type ExcelXML_sheets__ struct {
	List []ExcelXML_sheet_unmarshal__ `xml:"sheet"`
}

type ExcelXML_sheet_unmarshal__ struct {
	XMLName xml.Name `xml:"sheet"`
	Name    string   `xml:"name,attr"`
	Id      int      `xml:"sheetId,attr"`
	RId     string   `xml:"id,attr"`
}

type ExcelXML_sheet_marshal__ struct {
	XMLName xml.Name `xml:"sheet"`
	Name    string   `xml:"name,attr"`
	Id      int      `xml:"sheetId,attr"`
	RId     string   `xml:"r:id,attr"`
}

///////////////////////////////////////////////////////////////////////////////
///////////////////					WORKSHEET

type ExcelXML_worksheet_unmarshal__ struct {
	XMLName     xml.Name `xml:"worksheet"`
	Xmlns       string   `xml:"xmlns,attr"`
	R           string   `xml:"r,attr"`
	MC          string   `xml:"mc,attr"`
	MCIgnorable string   `xml:"Ignorable,attr"`
	X14ac       string   `xml:"x14ac,attr"`
	XR          string   `xml:"xr,attr"`
	XR2         string   `xml:"xr2,attr"`
	XR3         string   `xml:"xr3,attr"`
	XRUid       string   `xml:"uid,attr"`

	SheetPr       ExcelXML_sheetpr__       `xml:"sheetPr"`
	Dimension     ExcelXML_dimension__     `xml:"dimension"`
	SheetViews    ExcelXML_sheetviews__    `xml:"sheetViews"`
	SheetFormatPr ExcelXML_sheetformatpr__ `xml:"sheetFormatPr"`

	Cols      *ExcelXML_cols__     `xml:"cols"`
	SheetData ExcelXML_sheetdata__ `xml:"sheetData"`

	PhoneticPr  ExcelXML_phoneticpr__          `xml:"phoneticPr"`
	PageMargins ExcelXML_pagemargins__         `xml:"pageMargins"`
	PageSetup   ExcelXML_pagesetup_unmarshal__ `xml:"pageSetup"`
	Drawing     *ExcelXML_drawing_unmarshal__  `xml:"drawing"`
}

type ExcelXML_worksheet_marshal__ struct {
	XMLName     xml.Name `xml:"worksheet"`
	Xmlns       string   `xml:"xmlns,attr"`
	R           string   `xml:"xmlns:r,attr"`
	MC          string   `xml:"xmlns:mc,attr"`
	MCIgnorable string   `xml:"mc:Ignorable,attr,omitempty"`
	X14ac       string   `xml:"xmlns:x14ac,attr"`
	XR          string   `xml:"xmlns:xr,attr,omitempty"`
	XR2         string   `xml:"xmlns:xr2,attr,omitempty"`
	XR3         string   `xml:"xmlns:xr3,attr,omitempty"`
	XRUid       string   `xml:"xr:uid,attr,omitempty"`

	SheetPr       ExcelXML_sheetpr__       `xml:"sheetPr"`
	Dimension     ExcelXML_dimension__     `xml:"dimension"`
	SheetViews    ExcelXML_sheetviews__    `xml:"sheetViews"`
	SheetFormatPr ExcelXML_sheetformatpr__ `xml:"sheetFormatPr"`

	Cols      *ExcelXML_cols__     `xml:"cols,omitempty"`
	SheetData ExcelXML_sheetdata__ `xml:"sheetData"`

	PhoneticPr  ExcelXML_phoneticpr__        `xml:"phoneticPr"`
	PageMargins ExcelXML_pagemargins__       `xml:"pageMargins"`
	PageSetup   ExcelXML_pagesetup_marshal__ `xml:"pageSetup"`
	Drawing     *ExcelXML_drawing_marshal__  `xml:"drawing,omitempty"`
}

type ExcelXML_sheetpr__ struct {
	CodeName string `xml:"codeName,attr"`
}

type ExcelXML_dimension__ struct {
	Ref string `xml:"ref,attr"`
}

type ExcelXML_sheetviews__ struct {
	List []ExcelXML_sheetview__ `xml:"sheetView"`
}

type ExcelXML_sheetview__ struct {
	TabSelected    int `xml:"tabSelected,attr,omitempty"`
	WorkbookViewId int `xml:"workbookViewId,attr"`
}

type ExcelXML_sheetformatpr__ struct {
	DefaultColWidth  float64 `xml:"defaultColWidth,attr,omitempty"`
	DefaultRowHeight float64 `xml:"defaultRowHeight,attr,omitempty"`
}

type ExcelXML_cols__ struct {
	List []ExcelXML_col__ `xml:"col"`
}

type ExcelXML_col__ struct {
	Min         int     `xml:"min,attr"`
	Max         int     `xml:"max,attr"`
	Width       float64 `xml:"width,attr"`
	Style       int     `xml:"style,attr,omitempty"`
	CustomWidth int     `xml:"customWidth,attr"`
}

type ExcelXML_sheetdata__ struct {
	List []ExcelXML_row__ `xml:"row"`
}

type ExcelXML_row__ struct {
	R            int               `xml:"r,attr"`
	Spans        string            `xml:"spans,attr"`
	S            int               `xml:"s,attr,omitempty"`
	HT           string            `xml:"ht,attr,omitempty"`
	CustomFormat string            `xml:"customFormat,attr,omitempty"`
	Cols         []ExcelXML_cell__ `xml:"c"`
}

type ExcelXML_cell__ struct {
	R     string `xml:"r,attr"`
	S     string `xml:"s,attr,omitempty"`
	T     string `xml:"t,attr,omitempty"`
	Value string `xml:"v"`
}

type ExcelXML_phoneticpr__ struct {
	FontId int    `xml:"fontId,attr"`
	Type   string `xml:"type,attr"`
}

type ExcelXML_pagemargins__ struct {
	Left   string `xml:"left,attr"`
	Right  string `xml:"right,attr"`
	Top    string `xml:"top,attr"`
	Bottom string `xml:"bottom,attr"`
	Header string `xml:"header,attr"`
	Footer string `xml:"footer,attr"`
}

type ExcelXML_pagesetup_unmarshal__ struct {
	PaperSize   int    `xml:"paperSize,attr"`
	Orientation string `xml:"orientation,attr"`
	RId         string `xml:"id,attr"`
}

type ExcelXML_pagesetup_marshal__ struct {
	PaperSize   int    `xml:"paperSize,attr"`
	Orientation string `xml:"orientation,attr"`
	RId         string `xml:"r:id,attr"`
}

type ExcelXML_drawing_unmarshal__ struct {
	RId string `xml:"id,attr"`
}

type ExcelXML_drawing_marshal__ struct {
	RId string `xml:"r:id,attr"`
}

func (ws *ExcelXML_worksheet_unmarshal__) Marshal() ([]byte, error) {

	var drawing *ExcelXML_drawing_marshal__ = nil
	if ws.Drawing != nil {
		tmp := ExcelXML_drawing_marshal__(*ws.Drawing)
		drawing = &tmp
	}

	return xml.Marshal(ExcelXML_worksheet_marshal__{
		XMLName:     ws.XMLName,
		Xmlns:       ws.Xmlns,
		R:           ws.R,
		MC:          ws.MC,
		MCIgnorable: ws.MCIgnorable,
		X14ac:       ws.X14ac,
		XR:          ws.XR,
		XR2:         ws.XR2,
		XR3:         ws.XR3,
		XRUid:       ws.XRUid,

		SheetPr:       ws.SheetPr,
		Dimension:     ws.Dimension,
		SheetViews:    ws.SheetViews,
		SheetFormatPr: ws.SheetFormatPr,

		Cols:      ws.Cols,
		SheetData: ws.SheetData,

		PhoneticPr:  ws.PhoneticPr,
		PageMargins: ws.PageMargins,
		PageSetup:   ExcelXML_pagesetup_marshal__(ws.PageSetup),
		Drawing:     drawing,
	})
}

///////////////////////////////////////////////////////////////////////////////
///////////////////					OTHER

type ExcelXML_sharedstrings__ struct {
	XMLName     xml.Name `xml:"sst"`
	Xmlns       string   `xml:"xmlns,attr"`
	Count       int      `xml:"count,attr"`
	UniqueCount int      `xml:"uniqueCount,attr"`

	Si []ExcelXML_Si__ `xml:"si"`
}

type ExcelXML_Si__ struct {
	T string `xml:"t"`
}
