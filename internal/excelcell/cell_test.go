package excelcell

import (
	"testing"
	"time"

	"github.com/xuri/excelize/v2"
)

func TestSetCellAutoWritesSupportedTypes(t *testing.T) {
	file := excelize.NewFile()
	const sheet = "Sheet1"
	const excelDefaultNumberType = excelize.CellTypeUnset

	cases := []struct {
		name     string
		cell     string
		value    interface{}
		want     string
		wantType excelize.CellType
	}{
		{name: "bool true", cell: "A1", value: true, want: "TRUE", wantType: excelize.CellTypeBool},
		{name: "bool false", cell: "A2", value: false, want: "FALSE", wantType: excelize.CellTypeBool},
		{name: "int", cell: "A3", value: int(-1), want: "-1", wantType: excelDefaultNumberType},
		{name: "int8", cell: "A4", value: int8(-2), want: "-2", wantType: excelDefaultNumberType},
		{name: "int16", cell: "A5", value: int16(-3), want: "-3", wantType: excelDefaultNumberType},
		{name: "int32", cell: "A6", value: int32(-4), want: "-4", wantType: excelDefaultNumberType},
		{name: "int64", cell: "A7", value: int64(-5), want: "-5", wantType: excelDefaultNumberType},
		{name: "uint", cell: "A8", value: uint(1), want: "1", wantType: excelDefaultNumberType},
		{name: "uint8", cell: "A9", value: uint8(2), want: "2", wantType: excelDefaultNumberType},
		{name: "uint16", cell: "A10", value: uint16(3), want: "3", wantType: excelDefaultNumberType},
		{name: "uint32", cell: "A11", value: uint32(4), want: "4", wantType: excelDefaultNumberType},
		{name: "uint64", cell: "A12", value: uint64(5), want: "5", wantType: excelDefaultNumberType},
		{name: "float32", cell: "A13", value: float32(1.25), want: "1.25", wantType: excelDefaultNumberType},
		{name: "float64", cell: "A14", value: float64(2.5), want: "2.5", wantType: excelDefaultNumberType},
		{name: "string", cell: "A15", value: "00123", want: "00123", wantType: excelize.CellTypeSharedString},
		{name: "bytes", cell: "A16", value: []byte("bytes"), want: "bytes", wantType: excelize.CellTypeSharedString},
		{name: "nil", cell: "A17", value: nil, want: "", wantType: excelize.CellTypeUnset},
		{name: "default string", cell: "A18", value: DefaultString("raw text"), want: "raw text", wantType: excelize.CellTypeInlineString},
	}

	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			if err := SetCellAuto(file, sheet, tc.cell, tc.value); err != nil {
				t.Fatalf("SetCellAuto failed: %v", err)
			}

			got, err := file.GetCellValue(sheet, tc.cell)
			if err != nil {
				t.Fatalf("GetCellValue failed: %v", err)
			}
			if got != tc.want {
				t.Fatalf("unexpected cell value, got %q, want %q", got, tc.want)
			}

			gotType, err := file.GetCellType(sheet, tc.cell)
			if err != nil {
				t.Fatalf("GetCellType failed: %v", err)
			}
			if gotType != tc.wantType {
				t.Fatalf("unexpected cell type, got %v, want %v", gotType, tc.wantType)
			}
		})
	}
}

func TestSetCellAutoUsesExcelizeValueForTimeTypes(t *testing.T) {
	file := excelize.NewFile()
	const sheet = "Sheet1"

	if err := SetCellAuto(file, sheet, "B1", 90*time.Minute); err != nil {
		t.Fatalf("SetCellAuto duration failed: %v", err)
	}
	if err := SetCellAuto(file, sheet, "B2", time.Date(2026, 5, 28, 12, 30, 0, 0, time.UTC)); err != nil {
		t.Fatalf("SetCellAuto time failed: %v", err)
	}

	for _, cell := range []string{"B1", "B2"} {
		got, err := file.GetCellValue(sheet, cell)
		if err != nil {
			t.Fatalf("GetCellValue %s failed: %v", cell, err)
		}
		if got == "" {
			t.Fatalf("expected %s to have a formatted value", cell)
		}
	}
}

func TestSetCellAutoRejectsNilFile(t *testing.T) {
	if err := SetCellAuto(nil, "Sheet1", "A1", "value"); err == nil {
		t.Fatal("expected nil file error")
	}
}
