package main

import "tinyexcellib"

func main() {

	wb, err := tinyexcellib.ReadExcelFile("book1.xlsx")
	if err != nil {
		panic(err)
	}

	wb.InsertNewRow("Sheet1", 2, []tinyexcellib.Cell{{Value: "Hello"}, {Value: "World"}})
	// wb.InsertNewRow("Sheet1", 2, []tinyexcellib.Cell{{Value: 7493840.89}, {Value: 53.5380}, {Value: 7949.322}})
	wb.UpdateSheet("Sheet1")

	wb.Save()
}
