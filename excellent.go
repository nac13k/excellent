package excellent

import (
	"fmt"
	"net/http"
	"strconv"
	"strings"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
)

// New TODO: Doc
func New() *excelize.File {
	return excelize.NewFile()
}

func toChar(i int) string {
	if i/26 > 0 {
		return string(rune('A'-1+i/26)) + string(rune('A'-1+i%26))
	}
	return string(rune('A' - 1 + i))
}

// GetAxis TODO: Doc
func getAxis(x, y int) string {
	return fmt.Sprintf("%v%v", toChar(x+1), y+1)
}

func getSheet(index int) string {
	return fmt.Sprintf("Sheet%d", index)
}

// SetHeaders TODO: Doc
func setHeaders(headers *HeadersStruct, xlsx *excelize.File) {
	index := 0
	for k, v := range headers.Data {
		index++
		if index == 1 {
			xlsx.SetSheetName("Sheet1", k)
		} else {
			xlsx.NewSheet(index, k)
		}
		xlsx.SetActiveSheet(index)
		for i, d := range v {
			xlsx.SetCellValue(getSheet(index), getAxis(i, 0), d)
		}
	}
}

func setValues(values *ValuesStruct, xlsx *excelize.File) {
	index := 0
	for _, matrix := range values.Data {
		index++
		for y, ary := range matrix {
			for x, v := range ary {
				xlsx.SetCellValue(getSheet(index), getAxis(x, y+1), v)
			}
		}
		setCellColor(getSheet(index), getAxis(0, 0), getAxis(len(matrix[0])-1, len(matrix)-1), xlsx)
	}
}

func setCellColor(sheet, start, end string, xlsx *excelize.File) {
	format := `{"table_style":"TableStyleMedium2", "show_first_column":true,"show_last_column":true,"show_row_stripes":true,"show_column_stripes":false}`
	xlsx.AddTable(sheet, start, end, format)
}

func getActiveSheet(activeSheet int) int {
	if activeSheet == 0 {
		return 1
	}
	return activeSheet
}

func splitSelector(selector string) (string, string) err {
	s := strings.Split(selector, ":")
	return s[0], s[1]
}

// WIP TODO: weas
// func addImageFromURL(url string) {
// 	response, err := http.Get(url)

// 	if err != null {
// 		return err
// 	}

// 	defer response.Body.Close()
// 	file, err := os.Create(fmt.Sprintf("/tmp/%v.%v", strconv.Itoa(time.Now().Unix()), "jpg"))

// 	if err != nil {
// 		return err
// 	}

// 	_, err = io.Copy(file, response.Body)
// 	if err != nil {
// 		return err
// 	}
// 	file.Close()
// 	return err
// }

func addImageFromPath(selector, path string, xlsx *excelize.File) err {
	sheet, cell := splitSelector(selector)
	return xlsx.AddPicture(sheet, cell, path, "")
}

// SaveFile TODO: Doc
func saveFile(name, folder string, xlsx *excelize.File) (f string, e error) {
	f = fmt.Sprintf("%v/%v.xlsx", folder, name)
	e = xlsx.SaveAs(f)
	return
}
