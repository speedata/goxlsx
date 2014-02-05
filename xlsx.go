package goxlsx

import (
	"archive/zip"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"log"
	"strconv"
	"strings"
)

type Worksheet struct {
	Name        string
	filename    string
	id          string
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

type Spreadsheet struct {
	filepath        string
	compressedFiles []zip.File
	Worksheets      []*Worksheet
	sharedStrings   []string
}

func readWorkbook(d *xml.Decoder, s *Spreadsheet) []*Worksheet {
	worksheets := make([]*Worksheet, 0, 5)
	var (
		err   error
		token xml.Token
	)

	for {
		token, err = d.Token()
		if err != nil {
			if err != io.EOF {
				log.Fatal(err)
			}
			break
		}
		switch x := token.(type) {
		case xml.StartElement:
			switch x.Name.Local {
			case "sheet":
				ws := new(Worksheet)
				ws.spreadsheet = s
				for _, a := range x.Attr {
					if a.Name.Local == "name" {
						ws.Name = a.Value
					}
					if a.Name.Local == "sheetId" {
						ws.id = a.Value
					}
				}
				worksheets = append(worksheets, ws)
			}
		}
	}
	return worksheets
}

func readStrings(d *xml.Decoder, s *Spreadsheet) {
	var (
		err   error
		data  []byte
		token xml.Token
	)
	for {
		token, err = d.Token()
		if err != nil {
			if err != io.EOF {
				log.Fatal(err)
			}
			break
		}
		switch x := token.(type) {
		case xml.StartElement:
			switch x.Name.Local {
			case "sst":
				// root element
				for i := 0; i < len(x.Attr); i++ {
					if x.Attr[i].Name.Local == "uniqueCount" {
						count, err := strconv.Atoi(x.Attr[i].Value)
						if err != nil {
							log.Fatal(err)
						}
						s.sharedStrings = make([]string, 0, count)
					}
				}
			default:
				// log.Println(x.Name.Local)
			}
		case xml.CharData:
			data = x.Copy()
		case xml.EndElement:
			switch x.Name.Local {
			case "t":
				s.sharedStrings = append(s.sharedStrings, string(data))
			}
		}

	}
}

func OpenFile(path string) (*Spreadsheet, error) {
	xlsx := new(Spreadsheet)
	xlsx.filepath = path

	r, err := zip.OpenReader(path)
	if err != nil {
		return nil, err
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == "xl/workbook.xml" {
			rc, err := f.Open()
			if err != nil {
				return nil, err
			}
			xlsx.Worksheets = readWorkbook(xml.NewDecoder(rc), xlsx)
			rc.Close()
		}
		if f.Name == "xl/sharedStrings.xml" {
			rc, err := f.Open()
			if err != nil {
				return nil, err
			}
			readStrings(xml.NewDecoder(rc), xlsx)
			rc.Close()
		}
	}
	return xlsx, nil
}

func readWorksheetXML(dec *xml.Decoder, s *Spreadsheet) (map[int]*row, error) {
	rows := make(map[int]*row)
	var (
		err         error
		token       xml.Token
		rownum      int
		currentCell *cell
		currentRow  *row
	)
	for {
		token, err = dec.Token()
		if err != nil {
			if err != io.EOF {
				return nil, err
			}
			break
		}
		switch x := token.(type) {
		case xml.StartElement:
			switch x.Name.Local {
			case "row":
				currentRow = &row{}
				currentRow.Cells = make(map[int]*cell)
				for _, a := range x.Attr {
					if a.Name.Local == "r" {
						rownum, err = strconv.Atoi(a.Value)
						if err != nil {
							return nil, err
						}
					}
				}
				currentRow.Num = rownum
				rows[rownum] = currentRow
			case "c":
				currentCell = &cell{}
				var cellnumber rune
				for _, a := range x.Attr {
					switch a.Name.Local {
					case "r":
						for _, v := range a.Value {
							if v >= 'A' && v <= 'Z' {
								cellnumber = cellnumber*26 + v - 'A' + 1
							}
						}
					case "t":
						if a.Value == "s" {
							currentCell.Type = "s"
						} else if a.Value == "n" {
							currentCell.Type = "n"
						}

					}

				}
				currentRow.Cells[int(cellnumber)] = currentCell
			}
		case xml.EndElement:
			switch x.Name.Local {
			case "c":
				currentCell = nil
			}
		case xml.CharData:
			if currentCell != nil {
				val := string(x.Copy())
				if currentCell.Type == "s" {
					valInt, _ := strconv.Atoi(val)
					currentCell.Value = s.sharedStrings[valInt]
				} else if currentCell.Type == "n" {
					currentCell.Value = strings.TrimSuffix(val, ".0")
				} else {
					currentCell.Value = val
				}
			}
		}
	}
	return rows, nil
}

func (ws *Worksheet) readWorksheetZIP() error {
	r, err := zip.OpenReader(ws.spreadsheet.filepath)
	if err != nil {
		return err
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == ws.filename {
			rc, err := f.Open()
			if err != nil {
				return err
			}
			defer rc.Close()
			rows, err := readWorksheetXML(xml.NewDecoder(rc), ws.spreadsheet)
			ws.rows = rows
			if err != nil {
				return err
			}
		}
	}
	return nil
}

func (ws *Worksheet) Cell(row, column int) string {
	xrow := ws.rows[row]
	if xrow == nil {
		return ""
	}
	if xrow.Cells[column] == nil {
		return ""
	}
	return xrow.Cells[column].Value
}

func (s *Spreadsheet) GetWorksheet(number int) (*Worksheet, error) {
	if number >= len(s.Worksheets) || number < 0 {
		return nil, errors.New("Index out of range")
	}
	ws := s.Worksheets[number]
	ws.filename = "xl/worksheets" + "/" + fmt.Sprintf("sheet%s.xml", ws.id)
	err := ws.readWorksheetZIP()
	if err != nil {
		return nil, err
	}
	return ws, nil
}
