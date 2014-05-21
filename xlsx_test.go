package goxlsx

import (
	"path/filepath"
	"testing"
)

func TestOpenFile(t *testing.T) {
	xlsx, err := OpenFile(filepath.Join("_testdata", "Worksheet1.xlsx"))
	if err != nil {
		t.Error(err)
	}
	if xlsx.NumWorksheets() != 2 {
		t.Error("num of worksheets != 2")
	}

	ws, err := xlsx.GetWorksheet(0)
	if err != nil {
		t.Error(err)
	}
	if ws.filename != "xl/worksheets/sheet1.xml" {
		t.Error("filename mismatch, got", ws.filename)
	}
	if len(ws.rows) != 5 {
		t.Error("ws.rows != 5")
	}

	row := ws.rows[1]
	if row.Cells[1].Value != "A" {
		t.Error("First value should be A")
	}
	if row.Cells[2].Value != "B" {
		t.Error("Second value should be B")
	}
	if ws.Cell(1, 1) != "A" {
		t.Error("1,1 should be A")
	}

}
