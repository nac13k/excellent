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
	Data map[string][]string `json:"data"`
}

// ValuesStruct TODO: Doc
type ValuesStruct struct {
	Data map[string][][]interface{} `json:"data"`
}

// Unmarshal TODO: Doc
func (x *X) Unmarshal(data []byte) error {
	return json.Unmarshal(data, x)
}

// ToXLSX TODO: Doc
func (x *X) ToXLSX(folder string) (f string, e error) {
	xlsx, f := New(x.Name, folder)
	setHeaders(&x.Headers, xlsx)
	// xlsx.DeleteSheet(xlsx.GetSheetName(1))
	setValues(&x.Values, xlsx)
	e = xlsx.Save()
	// f, e = saveFile(xlsx)
	return
}
