package excellent

import (
	"fmt"
	"io"
	"os"

	"github.com/Luxurioust/excelize"
)

// New TODO: Doc
func New(name, folder string) (xlsx *excelize.File, f string) {
	f = fmt.Sprintf("%v/%v.xlsx", folder, name)
	println(f)
	err := copy("/Users/nac13k/Documents/bitlab/report_service/Book.xlsx", f)
	println(err)
	println("/Users/nac13k/Documents/bitlab/report_service/Book.xlsx")
	xlsx, _ = excelize.OpenFile(f)
	return
}

func copy(src, dest string) error {
	srcFile, err := os.Open(src)

	if err != nil {
		return err
	}

	defer srcFile.Close()

	destFile, err := os.Create(dest)

	if err != nil {
		return err
	}

	defer destFile.Close()

	_, err = io.Copy(destFile, srcFile) // check first var for number of bytes copied

	if err != nil {
		return err
	}

	err = destFile.Sync()
	return err
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

// SaveFile TODO: Doc
func saveFile(name, folder string, xlsx *excelize.File) (f string, e error) {
	f = fmt.Sprintf("%v/%v.xlsx", folder, name)
	e = xlsx.WriteTo(f)
	return
}
