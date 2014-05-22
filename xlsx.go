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

// Worksheet represents a single worksheet in an excel file.
// A worksheet is a rectangular area of cells, each cell can contain a value.
type Worksheet struct {
	Name        string
	MaxRow      int
	MaxColumn   int
	MinRow      int
	MinColumn   int
	filename    string
	id          string
	rid         string
	rows        map[int]*row
	spreadsheet *Spreadsheet
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

// Spreadsheet represents the whole .xlsx file.
type Spreadsheet struct {
	filepath          string
	compressedFiles   []zip.File
	worksheets        []*Worksheet
	sharedStrings     []string
	uncompressedFiles map[string][]byte
	relationships     map[string]relationship
}

// NumWorksheets returns the number of worksheets in a file.
func (s *Spreadsheet) NumWorksheets() int {
	return len(s.worksheets)
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
		w.id = wb.Sheets[i].SheetID
		w.rid = wb.Sheets[i].Rid
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
			return nil, fmt.Errorf("read (%d) not equal to uncompressed size (%d)", pos, f.UncompressedSize64)
		}

		xlsx.uncompressedFiles[f.Name] = buf
	}
	xlsx.relationships, err = readRelationships(xlsx.uncompressedFiles["xl/_rels/workbook.xml.rels"])
	if err != nil {
		return nil, err
	}
	xlsx.worksheets, err = readWorkbook(xlsx.uncompressedFiles["xl/workbook.xml"], xlsx)
	if err != nil {
		return nil, err
	}
	xlsx.sharedStrings = readStrings(xlsx.uncompressedFiles["xl/sharedStrings.xml"])

	return xlsx, nil
}

func readRelationships(data []byte) (map[string]relationship, error) {
	rels := &xslxRelationships{}
	err := xml.Unmarshal(data, rels)
	if err != nil {
		return nil, err
	}
	ret := make(map[string]relationship)
	for _, v := range rels.Relationship {
		ret[v.Id] = relationship{Type: v.Type, Target: v.Target}
	}
	return ret, nil
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
	wsXlsx := &xlsxWorksheet{}
	err := xml.Unmarshal(data, wsXlsx)
	if err != nil {
		return nil, err
	}
	ws := &Worksheet{}
	ws.rows = make(map[int]*row)
	tmp := strings.Split(wsXlsx.Dimension.Ref, ":")
	ws.MinColumn, ws.MinRow = stringToPosition(tmp[0])
	ws.MaxColumn, ws.MaxRow = stringToPosition(tmp[1])

	var currentRow *row

	for xrow := 0; xrow < len(wsXlsx.Row); xrow++ {
		thisrow := wsXlsx.Row[xrow]

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
	rid := s.worksheets[number].rid
	filename := "xl/" + s.relationships[rid].Target
	ws, err := s.readWorksheet(s.uncompressedFiles[filename])
	ws.filename = filename
	ws.Name = s.worksheets[number].Name
	if err != nil {
		return nil, err
	}
	return ws, nil
}
