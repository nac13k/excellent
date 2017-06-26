package excellent

import "encoding/json"

//X  TODO: Doc
type X struct {
	Name    string        `json:"name"`
	Headers HeadersStruct `json:"headers"`
	Values  ValuesStruct  `json:"values"`
}

//HeadersStruct  TODO: Doc
type HeadersStruct struct {
	Data  map[string][]string `json:"data"`
	Color string              `json:"color"`
}

// ValuesStruct TODO: Doc
type ValuesStruct struct {
	Data   map[string][][]interface{} `json:"data"`
	Format ValuesFormatStruct         `json:"format"`
}

// ValuesFormatStruct TODO: Doc
type ValuesFormatStruct struct {
	Index  int               `json:"index"`
	Colors map[string]string `json:"colors"`
}

// Unmarshal TODO: Doc
func (x *X) Unmarshal(data []byte) error {
	return json.Unmarshal(data, x)
}

// ToXLSX TODO: Doc
func (x *X) ToXLSX() (f string, e error) {
	xlsx := New()
	setHeaders(&x.Headers, xlsx)
	xlsx.DeleteSheet(xlsx.GetSheetName(1))
	setValues(&x.Values, xlsx)
	f, e = saveFile(x.Name, xlsx)
	return
}
