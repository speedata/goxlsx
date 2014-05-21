// Package goxlsx accesses Excel 2007 (.xslx) for reading.
package goxlsx

import (
	"archive/zip"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"strconv"
	"strings"
)

// Represent a worksheet in an excel file. A worksheet is a rectangular area of cells, each cell can contain a value.
type Worksheet struct {
	Name          string
	MaxRow        int
	MaxColumn     int
	MinRow        int
	MinColumn     int
	NumWorksheets int
	filename      string
	id            string
	rows          map[int]*row
	spreadsheet   *Spreadsheet
}

type cell struct {
	Name  string
	Type  string
	Value string
}

type row struct {
	Num   int
	Cells map[int]*cell
}

// A spreadsheet represents the .xlsx file.
type Spreadsheet struct {
	filepath          string
	compressedFiles   []zip.File
	worksheets        []*Worksheet
	NumWorksheets     int
	sharedStrings     []string
	uncompressedFiles map[string][]byte
}

func readWorkbook(data []byte, s *Spreadsheet) ([]*Worksheet, error) {
	wb := &workbook{}
	err := xml.Unmarshal(data, wb)
	if err != nil {
		return nil, err
	}
	var worksheets []*Worksheet

	for i := 0; i < len(wb.Sheets); i++ {
		w := &Worksheet{}
		w.spreadsheet = s
		w.Name = wb.Sheets[i].Name
		w.id = wb.Sheets[i].SheetId
		worksheets = append(worksheets, w)
	}
	return worksheets, nil
}

func readStrings(data []byte) []string {
	sst := &sst{}
	xml.Unmarshal(data, sst)
	ret := make([]string, sst.UniqueCount)
	for i := 0; i < sst.UniqueCount; i++ {
		ret[i] = sst.Si[i].T
	}
	return ret
}

// OpenFile reads a file located at the given path and returns a spreadsheet object.
func OpenFile(path string) (*Spreadsheet, error) {
	xlsx := new(Spreadsheet)
	xlsx.filepath = path
	xlsx.uncompressedFiles = make(map[string][]byte)

	r, err := zip.OpenReader(path)
	if err != nil {
		return nil, err
	}
	defer r.Close()

	for _, f := range r.File {
		buf := make([]byte, f.UncompressedSize64)
		rc, err := f.Open()
		if err != nil {
			return nil, err
		}
		pos := 0
	readfile:
		for {
			size, err := rc.Read(buf[pos:])
			if err == io.EOF {
				// ok, fine
				break readfile
			} else if err != nil {
				return nil, err
			}
			pos += size
		}
		if pos != int(f.UncompressedSize64) {
			return nil, errors.New(fmt.Sprintf("read (%d) not equal to uncompressed size (%d)", pos, f.UncompressedSize64))
		}

		xlsx.uncompressedFiles[f.Name] = buf
	}
	xlsx.worksheets, err = readWorkbook(xlsx.uncompressedFiles["xl/workbook.xml"], xlsx)
	if err != nil {
		return nil, err
	}
	xlsx.NumWorksheets = len(xlsx.worksheets)
	xlsx.sharedStrings = readStrings(xlsx.uncompressedFiles["xl/sharedStrings.xml"])

	return xlsx, nil
}

// excelpos is something like "AC101"
func stringToPosition(excelpos string) (int, int) {
	var columnnumber, rownumber rune
	for _, v := range excelpos {
		if v >= 'A' && v <= 'Z' {
			columnnumber = columnnumber*26 + v - 'A' + 1
		}
		if v >= '0' && v <= '9' {
			rownumber = rownumber*10 + v - '0'
		}
	}
	return int(columnnumber), int(rownumber)
}

// Cell returns the contents of cell at column, row, where 1,1 is the top left corner. The return value is always a string.
// The user is in charge to convert this value to a number, if necessary. Formulae are not returned.
func (ws *Worksheet) Cell(column, row int) string {
	xrow := ws.rows[row]
	if xrow == nil {
		return ""
	}
	if xrow.Cells[column] == nil {
		return ""
	}
	return xrow.Cells[column].Value
}

func (s *Spreadsheet) readWorksheet(data []byte) (*Worksheet, error) {
	ws_xlsx := &xlsx_worksheet{}
	err := xml.Unmarshal(data, ws_xlsx)
	if err != nil {
		return nil, err
	}
	ws := &Worksheet{}
	ws.rows = make(map[int]*row)
	tmp := strings.Split(ws_xlsx.Dimension.Ref, ":")
	ws.MinColumn, ws.MinRow = stringToPosition(tmp[0])
	ws.MaxColumn, ws.MaxRow = stringToPosition(tmp[1])

	var currentRow *row

	for xrow := 0; xrow < len(ws_xlsx.Row); xrow++ {
		thisrow := ws_xlsx.Row[xrow]

		currentRow = &row{}
		currentRow.Cells = make(map[int]*cell)
		currentRow.Num = thisrow.Rownumber
		ws.rows[thisrow.Rownumber] = currentRow

		for col := 0; col < len(thisrow.Cols); col++ {
			var cellnumber rune
			thiscol := thisrow.Cols[col]
			for _, v := range thiscol.R {
				if v >= 'A' && v <= 'Z' {
					cellnumber = cellnumber*26 + v - 'A' + 1
				}
			}
			currentCell := &cell{}

			currentRow.Cells[int(cellnumber)] = currentCell

			if thiscol.T == "s" {
				v, err := strconv.Atoi(thiscol.V)
				if err != nil {
					return nil, err
				}
				currentCell.Value = s.sharedStrings[v]
				currentCell.Type = "s"
			} else if thiscol.T == "" {
				currentCell.Type = "v"
				currentCell.Value = thiscol.V
			}

		}
	}
	return ws, nil
}

// GetWorksheet returns the worksheet with the given number, starting at 0.
func (s *Spreadsheet) GetWorksheet(number int) (*Worksheet, error) {
	if number >= len(s.worksheets) || number < 0 {
		return nil, errors.New("index out of range")
	}
	filename := fmt.Sprintf("xl/worksheets/sheet%s.xml", s.worksheets[number].id)
	ws, err := s.readWorksheet(s.uncompressedFiles[filename])
	ws.filename = filename
	if err != nil {
		return nil, err
	}
	return ws, nil
}
