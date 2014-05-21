package reader

import (
	"encoding/xml"
)

type sheet struct {
	Name    string `xml:"name,attr"`
	SheetId string `xml:"sheetId,attr"`
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

type xlsx_column struct {
	R    string `xml:"r,attr"`
	T    string `xml:"t,attr"`
	V    string `xml:"v"`
	Text string `xml:"is>t"`
}
type xlsx_row struct {
	Rownumber int           `xml:"r,attr"`
	Cols      []xlsx_column `xml:"c"`
}

type xslx_dimension struct {
	Ref string `xml:"ref,attr"`
}

type xlsx_worksheet struct {
	XMLName   xml.Name       `xml:"http://schemas.openxmlformats.org/spreadsheetml/2006/main worksheet"`
	Dimension xslx_dimension `xml:"dimension"`
	Row       []xlsx_row     `xml:"sheetData>row"`
}
