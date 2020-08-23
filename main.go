package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"github.com/tealeg/xlsx"
	"io/ioutil"
	"os"
	"strings"
)

var (
	path    string
	outPath string
	single  bool
	array   bool
)

func max(a, b int) int {
	if a > b {
		return a
	}

	return b
}

func init() {
	flag.StringVar(&path, "p", "", "Excel文件路径")
	flag.StringVar(&outPath, "op", "", "Json文件输出路径 (相对于excel)")
	flag.BoolVar(&single, "s", false, "true: 多个sheet导出一个json文件, false: 对应sheet导出对应json文件")
	flag.BoolVar(&array, "a", false, "true: 转换为数组, false: 转换为对象")
}

func main() {
	flag.Parse()
	if path == "" {
		flag.PrintDefaults()
		return
	}

	excel, err := xlsx.OpenFile(path)
	if err != nil {
		fmt.Println(err)
		return
	}

	lastIndex := max(strings.LastIndex(path, "\\"), strings.LastIndex(path, "/")) + 1
	outPath = path[:lastIndex] + outPath + "/"
	excelName := path[lastIndex:strings.LastIndex(path, ".")]
	if array {
		convertToArray(excel, excelName)
	} else {
		convertToJson(excel, excelName)
	}
}

func convertToArray(excel *xlsx.File, excelName string) {
	jsonMap := make(map[string]interface{})
	for _, sheet := range excel.Sheets {
		if len(sheet.Rows) == 0 || len(sheet.Rows[0].Cells) == 0 {
			continue
		}

		fileName := sheet.Rows[0].Cells[0].Value
		if fileName == "" {
			fileName = sheet.Name
		}

		sheetArray := sheetToArray(sheet)
		if single {
			jsonMap[fileName] = sheetArray
			continue
		}

		jsonByte, _ := json.Marshal(sheetArray)
		writeJson(jsonByte, fileName)
	}

	if single {
		jsonByte, _ := json.Marshal(jsonMap)
		writeJson(jsonByte, excelName)
	}
}

func convertToJson(excel *xlsx.File, excelName string) {
	jsonMap := make(map[string]interface{})
	for _, sheet := range excel.Sheets {
		if len(sheet.Rows) == 0 || len(sheet.Rows[0].Cells) == 0 {
			continue
		}

		fileName := sheet.Rows[0].Cells[0].Value
		if fileName == "" {
			fileName = sheet.Name
		}

		sheetJson := sheetToJson(sheet)
		if single {
			jsonMap[fileName] = sheetJson
			continue
		}

		jsonByte, _ := json.Marshal(sheetJson)
		writeJson(jsonByte, fileName)
	}

	if single {
		jsonByte, _ := json.Marshal(jsonMap)
		writeJson(jsonByte, excelName)
	}
}

func writeJson(json []byte, fileName string) {
	err := os.MkdirAll(outPath, os.ModePerm)
	if err != nil {
		fmt.Println(err)
		return
	}

	writePath := outPath + fileName + ".json"
	err = ioutil.WriteFile(writePath, json, os.ModePerm)
	if err != nil {
		fmt.Println(err)
	} else {
		fmt.Println("导出成功:", writePath)
	}
}

func sheetToArray(sheet *xlsx.Sheet) []map[string]interface{} {
	jsonArray := make([]map[string]interface{}, 0)
	for i, row := range sheet.Rows {
		if i <= 1 || len(row.Cells) == 0 || row.Cells[0].Value == "" {
			continue
		}

		rowJson := rowToJson(row, sheet.Rows[0])
		if len(rowJson) != 0 {
			jsonArray = append(jsonArray, rowJson)
		}
	}

	return jsonArray
}

func sheetToJson(sheet *xlsx.Sheet) map[string]interface{} {
	jsonMap := make(map[string]interface{})
	for i, row := range sheet.Rows {
		if i <= 1 || len(row.Cells) == 0 {
			continue
		}

		key := row.Cells[0].Value
		if key == "" {
			continue
		}

		rowJson := rowToJson(row, sheet.Rows[0])
		if len(rowJson) != 0 {
			jsonMap[key] = rowJson
		}
	}

	return jsonMap
}

// 把一行数据转换为json对象
func rowToJson(row *xlsx.Row, keys *xlsx.Row) map[string]interface{} {
	jsonMap := make(map[string]interface{})
	for i, cell := range row.Cells {
		if i == 0 {
			continue
		}

		key := keys.Cells[i].Value
		if key == "" {
			break
		}

		val := cellToVal(cell)
		jsonMap[key] = val
	}

	return jsonMap
}

func cellToVal(cell *xlsx.Cell) interface{} {
	switch cell.Type() {
	case xlsx.CellTypeBool:
		return cell.Bool()
	case xlsx.CellTypeNumeric:
		val, err := cell.Float()
		if err != nil {
			return 0
		}

		return val
	default:
		return cell.Value
	}
}
