package goxlsx

import (
	. "github.com/smartystreets/goconvey/convey"
	"path/filepath"
	"testing"
)

func TestOpenFile(t *testing.T) {
	Convey("Trying to open sample worksheet", t, func() {
		xlsx, err := OpenFile(filepath.Join("_testdata", "Worksheet1.xlsx"))
		if err != nil {
			t.Error(err)
		}
		So(len(xlsx.Worksheets), ShouldEqual, 2)

		Convey("Looking at the first worksheet", func() {
			ws, err := xlsx.GetWorksheet(0)
			So(err, ShouldBeNil)
			So(ws.filename, ShouldEqual, "xl/worksheets/sheet1.xml")
			So(len(ws.rows), ShouldEqual, 5)
			Convey("Looking at first row", func() {
				row := ws.rows[1]
				So(row.Cells[1].Value, ShouldEqual, "A")
				So(row.Cells[2].Value, ShouldEqual, "B")
			})

			Convey("Getting Cell (1,1)", func() {
				So(ws.Cell(1, 1), ShouldEqual, "A")
			})
		})
	})
}
