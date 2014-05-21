package goxlsx

import (
	"encoding/xml"
)

type sheet struct {
	Name    string `xml:"name,attr"`
	SheetID string `xml:"sheetId,attr"`
	Rid     string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
}
type workbook struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main workbook"`
	Sheets  []sheet  `xml:"sheets>sheet"`
}

type si struct {
	T string `xml:"t"`
}
type sst struct {
	XMLName     xml.Name `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main sst"`
	Count       int      `xml:"count,attr"`
	UniqueCount int      `xml:"uniqueCount,attr"`
	Si          []si     `xml:"si"`
}

type xlsxColumn struct {
	R    string `xml:"r,attr"`
	T    string `xml:"t,attr"`
	V    string `xml:"v"`
	Text string `xml:"is>t"`
}
type xlsx_row struct {
	Rownumber int          `xml:"r,attr"`
	Cols      []xlsxColumn `xml:"c"`
}

type xslxDimension struct {
	Ref string `xml:"ref,attr"`
}

type xlsxWorksheet struct {
	XMLName   xml.Name      `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main worksheet"`
	Dimension xslxDimension `xml:"dimension"`
	Row       []xlsx_row    `xml:"sheetData>row"`
}
