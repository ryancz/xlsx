package xlsx

import "github.com/tealeg/xlsx"

const DefSheetName = "Sheet1"

type File struct {
	f *xlsx.File
}

func (f *File) DefSheet() *Sheet {
	return f.AddSheet(DefSheetName)
}

func (f *File) AddSheet(name string) *Sheet {
	var s *xlsx.Sheet
	if f.f.Sheet[name] != nil {
		s = f.f.Sheet[name]
	} else {
		s, _ = f.f.AddSheet(name)
	}
	return &Sheet{
		s: s,
	}
}

func (f *File) WriteRow(data ...interface{}) {
	f.DefSheet().WriteRow(data...)
}

func (f *File) Save(path string) error {
	return f.f.Save(path)
}

func NewFile() *File {
	return &File{
		f: xlsx.NewFile(),
	}
}

type Sheet struct {
	s *xlsx.Sheet
}

func (s *Sheet) WriteRow(data ...interface{}) {
	r := s.s.AddRow()
	for _, value := range data {
		switch vv := value.(type) {
		case string:
			r.AddCell().SetString(vv)
		case int:
			r.AddCell().SetInt(vv)
		case int64:
			r.AddCell().SetInt64(vv)
		case float64:
			r.AddCell().SetFloat(vv)
		default:
			r.AddCell().SetString("#")
		}
	}
}

func (s *Sheet) WriteRows(rows [][]interface{}) {
	for _, row := range rows {
		s.WriteRow(row...)
	}
}
