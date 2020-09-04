package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"log"
)

/*
	Sammidev
	8:49 PM
	Thursday, September 3
	Thunderstorm 30C
*/

func main() {
	writeExcell()
	read()
	writeExcellStyle()
}

type M map[string]interface{}
var data = []M{
	M{"Nisn": "00001", "Name": "Ayatullah Ramadhan Jacoeb", "Gender": "male", "Age": 55, "Email": "Ayatullah@gmail.com", "Telp": "08xxxx"},
	M{"Nisn": "00002", "Name": "Abdul Rauf",   			    "Gender": "male", "Age": 20, "Email": "Abdul@gmail.com",     "Telp": "08xxxx"},
	M{"Nisn": "00003", "Name": "Dandi Arnanda",    			"Gender": "male", "Age": 39, "Email": "Dandi@gmail.com",     "Telp": "08xxxx"},
	M{"Nisn": "00004", "Name": "Sammi Aldhi Yanto", 	    "Gender": "male", "Age": 19, "Email": "Sammidev@gmail.com",  "Telp": "08xxxx"},
	M{"Nisn": "00005", "Name": "Aditya Andika Putra",       "Gender": "male", "Age": 39, "Email": "Aditya@gmail.com",    "Telp": "08xxxx"},
	M{"Nisn": "00006", "Name": "Gusnur",   				    "Gender": "male", "Age": 59, "Email": "Gus@gmail.com",       "Telp": "08xxxx"},
	M{"Nisn": "00007", "Name": "Aditya Fauzan Nul Haq",     "Gender": "male", "Age": 69, "Email": "Haq@gmail.com", 	  	 "Telp": "08xxxx"},
}

// write excell
func writeExcell()  {

	xlsx := excelize.NewFile()
	sheet1Name := "Sheet One"

	xlsx.SetSheetName(xlsx.GetSheetName(1), sheet1Name)
	xlsx.SetCellValue(sheet1Name, "A1", "Nisn")
	xlsx.SetCellValue(sheet1Name, "B1", "Name")
	xlsx.SetCellValue(sheet1Name, "C1", "Gender")
	xlsx.SetCellValue(sheet1Name, "D1", "Age")
	xlsx.SetCellValue(sheet1Name, "E1", "Email")
	xlsx.SetCellValue(sheet1Name, "F1", "Telp")

	err := xlsx.AutoFilter(sheet1Name, "A1","F1", "")
	if err != nil {
		log.Fatal("error", err.Error())
	}
	for i, each := range data {
			xlsx.SetCellValue(sheet1Name, fmt.Sprintf("A%d", i+2), each["Nisn"])
			xlsx.SetCellValue(sheet1Name, fmt.Sprintf("B%d", i+2), each["Name"])
			xlsx.SetCellValue(sheet1Name, fmt.Sprintf("C%d", i+2), each["Gender"])
			xlsx.SetCellValue(sheet1Name, fmt.Sprintf("D%d", i+2), each["Age"])
			xlsx.SetCellValue(sheet1Name, fmt.Sprintf("E%d", i+2), each["Email"])
			xlsx.SetCellValue(sheet1Name, fmt.Sprintf("F%d", i+2), each["Telp"])
	}
	err = xlsx.SaveAs("./data.xlsx")
	if err != nil {
		fmt.Println(err)
	}
	fmt.Println("DONE")
}

// write excell with style
func writeExcellStyle()  {
	xlsx := excelize.NewFile()
	sheet1Name := "Sheet One"

	xlsx.SetSheetName(xlsx.GetSheetName(1), sheet1Name)
	xlsx.SetCellValue(sheet1Name, "A1", "Nisn")
	xlsx.SetCellValue(sheet1Name, "B1", "Name")
	xlsx.SetCellValue(sheet1Name, "C1", "Gender")
	xlsx.SetCellValue(sheet1Name, "D1", "Age")
	xlsx.SetCellValue(sheet1Name, "E1", "Email")
	xlsx.SetCellValue(sheet1Name, "F1", "Telp")

	err := xlsx.AutoFilter(sheet1Name, "A1","F1", "")
	if err != nil {
		log.Fatal("error", err.Error())
	}
	for i, each := range data {
		xlsx.SetCellValue(sheet1Name, fmt.Sprintf("A%d", i+2), each["Nisn"])
		xlsx.SetCellValue(sheet1Name, fmt.Sprintf("B%d", i+2), each["Name"])
		xlsx.SetCellValue(sheet1Name, fmt.Sprintf("C%d", i+2), each["Gender"])
		xlsx.SetCellValue(sheet1Name, fmt.Sprintf("D%d", i+2), each["Age"])
		xlsx.SetCellValue(sheet1Name, fmt.Sprintf("E%d", i+2), each["Email"])
		xlsx.SetCellValue(sheet1Name, fmt.Sprintf("F%d", i+2), each["Telp"])
	}

	sheet2Name := "Sheet two"
	sheetIndex := xlsx.NewSheet(sheet2Name)
	xlsx.SetActiveSheet(sheetIndex)

	xlsx.SetCellValue(sheet2Name, "A1", "Hello Sam")
	xlsx.MergeCell(sheet2Name, "A1", "B1")

	style, err := xlsx.NewStyle(`{
        "font": {
            "bold": true,
            "size": 20
        },
        "fill": {
            "type": "pattern",
            "color": ["#E0EBF5"],
            "pattern": 1
        }
    }`)
	if err != nil {
		log.Fatal("ERROR", err.Error())
	}
	xlsx.SetCellStyle(sheet2Name, "A1", "A1", style)

	err = xlsx.SaveAs("./file2.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}


// read excell
func read()  {

	xlsx, err := excelize.OpenFile("./data.xlsx")
	if err != nil {
		log.Fatal("ERROR", err.Error())
	}
	sheet1Name := "Sheet One"
	rows := make([]M, 0)

	for i := 2; i < 9; i++ {
		row := M{
			"Nisn":   xlsx.GetCellValue(sheet1Name, fmt.Sprintf("A%d", i)),
			"Name":	  xlsx.GetCellValue(sheet1Name, fmt.Sprintf("B%d", i)),
			"Gender": xlsx.GetCellValue(sheet1Name, fmt.Sprintf("C%d", i)),
			"Age":    xlsx.GetCellValue(sheet1Name, fmt.Sprintf("D%d", i)),
			"Email":  xlsx.GetCellValue(sheet1Name, fmt.Sprintf("E%d", i)),
			"Telp":   xlsx.GetCellValue(sheet1Name, fmt.Sprintf("F%d", i)),
		}
		rows = append(rows, row)
	}

	for i, each := range rows{
		fmt.Println(i + 1, " -> ", each)
	}
}