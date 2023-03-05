package main

import "tinyexcellib"

func main() {

	wb, err := tinyexcellib.ReadExcelFile("book1.xlsx")
	if err != nil {
		panic(err)
	}

	wb.InsertNewRow("Sheet1", 1, []tinyexcellib.Cell{{Value: "Hello"}, {Value: "World"}})
	wb.UpdateSheet("Sheet1")

	wb.Save()
}
