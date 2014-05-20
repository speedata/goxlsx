goxlsx
======

Excel-XML reader for Go

Installation
------------
    go get github.com/speedata/goxlsx


Usage example
-------------
````go
package main

import (
    "fmt"
    "github.com/speedata/goxlsx"
    "log"
)

func main() {
    excelfile := "src/github.com/speedata/goxlsx/_testdata/Worksheet1.xlsx"
    spreadsheet, err := goxlsx.OpenFile(excelfile)
    if err != nil {
        log.Fatal(err)
    }
    ws1, err := spreadsheet.GetWorksheet(0)
    if err != nil {
        log.Fatal(err)
    }
    fmt.Println(ws1.Cell(1, 1))

}
````

Other:
-----

Status: usable<br>
Supported/maintained: yes<br>
Contribution welcome: yes (pull requests, issues)<br>
Main page: https://github.com/speedata/goxlsx<br>
License: MIT<br>
Contact: gundlach@speedata.de<br>
