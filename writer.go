package eb

import (
	"errors"
	"github.com/tealeg/xlsx/v3"
	"os"
	"path/filepath"
	"reflect"
)

type WriterBind interface {
	Write(option *WriterOption)
}

var (
	errTsIsNil = errors.New("exl: ts is nil")
)

func write(sheet *xlsx.Sheet, data []any) {
	r := sheet.AddRow()
	for _, cell := range data {
		r.AddCell().SetValue(cell)
	}
}

// Write defines write []T to excel file
//
// params: file,excel file full path
//
// params: typed parameter T, must be implements exl.Bind
func Write[T WriterBind](file string, ts []T) error {
	if ts == nil {
		return errTsIsNil
	}
	f := xlsx.NewFile()
	wm := NewWriterOption()
	if len(ts) > 0 {
		ts[0].Write(wm)
	}
	tT := new(T)
	if sheet, _ := f.AddSheet(wm.SheetName); sheet != nil {
		typ := reflect.TypeOf(tT).Elem().Elem()
		numField := typ.NumField()
		header := make([]any, numField, numField)
		for i := 0; i < numField; i++ {
			fe := typ.Field(i)
			name := fe.Name
			if tt, have := fe.Tag.Lookup(wm.TagName); have {
				name = tt
			}
			header[i] = name
		}
		// write header
		write(sheet, header)
		if len(ts) > 0 {
			// write data
			for _, t := range ts {
				data := make([]any, numField, numField)
				for i := 0; i < numField; i++ {
					data[i] = reflect.ValueOf(t).Elem().Field(i).Interface()
				}
				write(sheet, data)
			}
		}
	}

	_ = os.MkdirAll(filepath.Dir(file), 0600)

	return f.Save(file)
}

// WriteExcel defines write [][]string to excel
//
// params: file, excel file pull path
//
// params: data, write data to excel
func WriteExcel(file string, data [][]string) error {
	f := xlsx.NewFile()
	sheet, _ := f.AddSheet("Sheet1")

	for _, row := range data {
		r := sheet.AddRow()
		for _, cell := range row {
			r.AddCell().SetString(cell)
		}
	}

	return f.Save(file)
}
