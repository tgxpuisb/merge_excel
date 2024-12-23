package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func getExcel() {
	f, err := excelize.OpenFile("1.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	rows, err := f.GetRows("评价表")
	for _, row := range rows {
		for _, colCell := range row {
			fmt.Print(colCell, "\t")
		}
		fmt.Println()
	}
}

func genNewFile() {
	srcFile, err := excelize.OpenFile("1.xlsx")
	if err != nil {
		fmt.Println("打开源文件失败:", err)
		return
	}
	defer func() {
		if err := srcFile.Close(); err != nil {
			fmt.Println("关闭源文件失败:", err)
		}
	}()

	srcFile.SaveAs("new.xlsx")
}

func copyData() {
	f, err := excelize.OpenFile("1.xlsx")
	targetFile, e := excelize.OpenFile("new.xlsx")
	if err != nil || e != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
		if err := targetFile.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	rows, err := f.GetRows("评价表")
	targetRows, e := f.GetRows("评价表")
	targetRowsLen := len(targetRows)
	for rowNum, row := range rows[5:] {
		for colNum, cellValue := range row {
			colName, e := excelize.ColumnNumberToName(colNum + 1)
			if e != nil {
				fmt.Printf("转换列名失败: %v\n", err)
				continue
			}

			cellAxis := fmt.Sprintf("%s%d", colName, rowNum+targetRowsLen+1)
			if colNum == 0 {
				targetFile.SetCellValue("评价表", cellAxis, rowNum+targetRowsLen+1)
			} else {
				targetFile.SetCellValue("评价表", cellAxis, cellValue)
			}

			if err != nil {
				fmt.Printf("转换列名失败: %v\n", err)
				continue
			}
		}
	}
	targetFile.SaveAs("new1.xlsx")
}

func main() {
	// fmt.Println("hello")
	// f := excelize.NewFile()

	// defer func() {
	// 	if err := f.Close(); err != nil {
	// 		fmt.Println(err)
	// 	}
	// }()
	// getExcel()
	// genNewFile()
	copyData()

}
