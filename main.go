package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"os"

	"github.com/xuri/excelize/v2"
)

func main() {
	flag.Parse()
	dir := flag.Args()
	if len(dir) == 0 {
		fmt.Println(fmt.Sprintf("DirRequired"))
		return
	}
	getExcelData(dir[0])
}

func getExcelData(dir string) {
	f, err := excelize.OpenFile(dir)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	for _, name := range f.GetSheetMap() {
		fmt.Println("SheetName: " + name)

		WriterCSV := csv.NewWriter(os.Stdout)

		rows, err := f.Rows(name)
		if err != nil {
			fmt.Println(err)
			return
		}
		for rows.Next() {
			row, err := rows.Columns()
			if err != nil {
				fmt.Println(err)
			}
			csverr := WriterCSV.Write(row)
			if csverr != nil {
				return
			}
		}
		WriterCSV.Flush()
	}

}
