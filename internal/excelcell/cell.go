package excelcell

import (
	"fmt"
	"time"

	"github.com/xuri/excelize/v2"
)

// DefaultString writes a string with excelize.SetCellDefault instead of
// SetCellStr. Use it only when the caller intentionally wants Excel's default
// string behavior without excelize's special-character filtering.
type DefaultString string

// SetCellAuto writes value into a cell by dispatching to the most specific
// excelize setter for the Go value type.
func SetCellAuto(file *excelize.File, sheet, cell string, value interface{}) error {
	if file == nil {
		return fmt.Errorf("excel file is nil")
	}

	switch v := value.(type) {
	case nil:
		return file.SetCellValue(sheet, cell, nil)
	case bool:
		return file.SetCellBool(sheet, cell, v)
	case int:
		return file.SetCellInt(sheet, cell, int64(v))
	case int8:
		return file.SetCellInt(sheet, cell, int64(v))
	case int16:
		return file.SetCellInt(sheet, cell, int64(v))
	case int32:
		return file.SetCellInt(sheet, cell, int64(v))
	case int64:
		return file.SetCellInt(sheet, cell, v)
	case uint:
		return file.SetCellUint(sheet, cell, uint64(v))
	case uint8:
		return file.SetCellUint(sheet, cell, uint64(v))
	case uint16:
		return file.SetCellUint(sheet, cell, uint64(v))
	case uint32:
		return file.SetCellUint(sheet, cell, uint64(v))
	case uint64:
		return file.SetCellUint(sheet, cell, v)
	case float32:
		return file.SetCellFloat(sheet, cell, float64(v), -1, 32)
	case float64:
		return file.SetCellFloat(sheet, cell, v, -1, 64)
	case string:
		return file.SetCellStr(sheet, cell, v)
	case []byte:
		return file.SetCellStr(sheet, cell, string(v))
	case time.Duration:
		return file.SetCellValue(sheet, cell, v)
	case time.Time:
		return file.SetCellValue(sheet, cell, v)
	case DefaultString:
		return file.SetCellDefault(sheet, cell, string(v))
	default:
		return file.SetCellValue(sheet, cell, value)
	}
}
