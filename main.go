package main

import (
	"fmt"
	"github.com/eqinox76/xls"
	"github.com/xuri/excelize/v2"
	"io/ioutil"
	"log"
	"path"
	"path/filepath"
)

func main() {
	// 数据分析.xls文件路径
	templateFilePath := `D:\MIP\模板.xlsx`

	// 获取当前文件夹下所有Excel文件
	files, err := ioutil.ReadDir(".")
	if err != nil {
		log.Fatal(err)
	}

	// 创建新的Excel文件
	newFile := excelize.NewFile()

	for _, file := range files {
		// 检查文件是否为Excel文件
		if (filepath.Ext(file.Name()) == ".XLS" || filepath.Ext(file.Name()) == ".xls") && (file.Name() != "模板.xlsx" && file.Name() != "数据分析.xlsx") {
			// 打开Excel文件

			xlFile, err := xls.Open(file.Name(), "utf-8")
			if err != nil {
				log.Println(err)
				continue
			}
			sheet1 := xlFile.GetSheet(0)
			if sheet1 == nil {
				continue
			}

			// 读取模板文件
			templateFile, err := excelize.OpenFile(templateFilePath)
			if err != nil {
				log.Fatal(err)
			}

			valA1 := sheet1.Row(16).Col(1)
			templateFile.SetCellValue("1", "B3", removeLast(valA1))

			valA2 := sheet1.Row(17).Col(3)
			templateFile.SetCellValue("1", "C3", removeLast(valA2))

			valA3 := sheet1.Row(11).Col(3)
			templateFile.SetCellValue("1", "D3", removeLast(valA3))
			templateFile.SetCellValue("1", "H3", removeLast(valA3))

			valA4 := sheet1.Row(17).Col(1)
			templateFile.SetCellValue("1", "I3", removeLast(valA4))

			valA5 := sheet1.Row(28).Col(1)
			templateFile.SetCellValue("1", "K3", removeLast(valA5))

			valA6 := sheet1.Row(32).Col(1)
			templateFile.SetCellValue("1", "P3", removeLast(valA6))

			valA7 := sheet1.Row(30).Col(1)
			templateFile.SetCellValue("1", "Q3", removeLast(valA7))

			valA8 := sheet1.Row(31).Col(1)
			templateFile.SetCellValue("1", "R3", removeLast(valA8))

			valA9 := sheet1.Row(29).Col(1)
			templateFile.SetCellValue("1", "T3", removeLast(valA9))

			valA10 := sheet1.Row(5).Col(1)
			templateFile.SetCellValue("1", "A3", valA10)

			// 将当前Excel表格数据作为新的sheet添加到新的Excel文件中
			newSheetName := valA10
			newFile.NewSheet(newSheetName)
			rows, err := templateFile.GetRows("1")
			if err != nil {
				log.Println(err)
				continue
			}
			for rowIndex, row := range rows {
				rowNum := rowIndex + 1
				for colIndex, cellValue := range row {
					cell := ToAlphaString(colIndex) + fmt.Sprint(rowNum)
					formula, _ := templateFile.GetCellFormula("1", cell)
					if formula != "" {
						newFile.SetCellFormula(newSheetName, cell, formula)
						continue
					}
					newFile.SetCellValue(newSheetName, cell, cellValue)
				}
			}
		}
	}
	err = newFile.DeleteSheet("sheet1")
	if err != nil {
		log.Fatal(err)
	}
	// 保存新的Excel文件
	err = newFile.SaveAs("数据分析.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	fmt.Println("数据已成功导出到数据.xls文件")
}

func removeExtension(fileName string) string {
	extension := path.Ext(fileName)
	return fileName[:len(fileName)-len(extension)]
}

func ToAlphaString(value int) string {
	if value < 0 {
		return ""
	}
	var ans string
	i := value + 1
	for i > 0 {
		ans = string((i-1)%26+65) + ans
		i = (i - 1) / 26
	}
	return ans
}

func removeLast(value string) string {
	var numStr string
	for _, char := range value {
		if (char >= '0' && char <= '9') || char == '.' {
			numStr += string(char)
		} else {
			return numStr
		}
	}
	return numStr
}
