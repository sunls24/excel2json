package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"github.com/tealeg/xlsx"
	"io/ioutil"
	"os"
	"path/filepath"
	"strings"
	"time"
)

var (
	dir     string
	paths   string
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
	flag.StringVar(&dir, "d", "", "导出一个文件夹下的所有Excel (参数p失效)")
	flag.StringVar(&paths, "p", "", "Excel文件路径, 多个以','分割")
	flag.StringVar(&outPath, "op", "", "Json文件输出路径 (相对于excel)")
	flag.BoolVar(&single, "s", false, "多个sheet导出一个json文件, 默认对应sheet导出对应json文件")
	flag.BoolVar(&array, "a", false, "转换为数组, 默认转换为对象")
}

func main() {
	flag.Parse()
	if paths == "" && dir == "" {
		flag.PrintDefaults()
		return
	}

	startTime := time.Now()
	var pathArray []string
	if dir != "" {
		pathArray, _ = filepath.Glob(dir + "/*.xlsx")
	} else {
		pathArray = strings.Split(paths, ",")
	}

	ch := make(chan int, len(pathArray))
	for _, path := range pathArray {
		go func(path string) {
			excel, err := xlsx.OpenFile(path)
			if err != nil {
				fmt.Println(err)
				ch <- 0
				return
			}

			lastIndex := max(strings.LastIndex(path, "\\"), strings.LastIndex(path, "/")) + 1
			writePath := path[:lastIndex]
			if outPath != "" {
				writePath += outPath + "/"
			}

			excelName := path[lastIndex:strings.LastIndex(path, ".")]
			if array {
				convertToArray(excel, excelName, writePath)
			} else {
				convertToJson(excel, excelName, writePath)
			}

			ch <- 0
		}(path)
	}

	for range pathArray {
		<-ch
	}

	fmt.Println("用时:", time.Since(startTime))
}

func convertToArray(excel *xlsx.File, excelName string, writePath string) {
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
		writeJson(jsonByte, fileName, writePath)
	}

	if single {
		jsonByte, _ := json.Marshal(jsonMap)
		writeJson(jsonByte, excelName, writePath)
	}
}

func convertToJson(excel *xlsx.File, excelName string, writePath string) {
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
		writeJson(jsonByte, fileName, writePath)
	}

	if single {
		jsonByte, _ := json.Marshal(jsonMap)
		writeJson(jsonByte, excelName, writePath)
	}
}

func writeJson(json []byte, fileName string, writePath string) {
	err := os.MkdirAll(writePath, os.ModePerm)
	if err != nil {
		fmt.Println(err)
		return
	}

	writePath += fileName + ".json"
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
