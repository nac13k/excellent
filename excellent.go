package excellent

import (
	"fmt"

	"github.com/Luxurioust/excelize"
)

// New TODO: Doc
func New() *excelize.File {
	return excelize.CreateFile()
}

func toChar(i int) string {
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
	index := 1
	// fmt.Println("hoakksakas", xlsx.GetSheetName(1))
	for k, v := range headers.Data {
		index++
		xlsx.NewSheet(index, k)
		xlsx.SetActiveSheet(index)
		for i, d := range v {
			xlsx.SetCellValue(getSheet(index), getAxis(i, 0), d)
		}
	}
}

func setValues(values *ValuesStruct, xlsx *excelize.File) {
	index := 1
	for _, matrix := range values.Data {
		index++
		for y, ary := range matrix {
			for x, v := range ary {
				xlsx.SetCellValue(getSheet(index), getAxis(x, y+1), v)
			}
		}
		setCellColor(getSheet(index), getAxis(0, 0), getAxis(len(matrix[0]), len(matrix)), xlsx)
	}
}

func setCellColor(sheet, start, end string, xlsx *excelize.File) {
	format := `{"table_style":"TableStyleMedium2", "show_first_column":true,"show_last_column":true,"show_row_stripes":true,"show_column_stripes":false}`
	xlsx.AddTable(sheet, start, end, format)
	// err := xlsx.AutoFilter(sheet, start, end, "")
	// if err != nil {
	// 	fmt.Println(err)
	// }
}

func getActiveSheet(activeSheet int) int {
	if activeSheet == 0 {
		return 1
	}
	return activeSheet
}

// SaveFile TODO: Doc
func saveFile(name string, xlsx *excelize.File) (f string, e error) {
	f = fmt.Sprintf("./files/%v.xlsx", name)
	e = xlsx.WriteTo(f)
	return
}