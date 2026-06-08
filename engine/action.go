package engine

import (
	"fmt"
	"github.com/iEvan-lhr/go-excel-agent/workbook"
	"regexp"
	"strconv"
	"strings"
)

var numericValueRegex = regexp.MustCompile(`[-+]?\d+(?:\.\d+)?`)

func applyAction(currentValue string, action UpdateAction) (display string, typed any, err error) {
	actionType := strings.ToLower(strings.TrimSpace(action.Type))
	switch actionType {
	case "overwrite":
		return workbook.DisplayValue(action.Value), action.Value, nil
	case "append_suffix":
		value := currentValue + workbook.DisplayValue(action.Value)
		return value, value, nil
	case "prepend_prefix":
		value := workbook.DisplayValue(action.Value) + currentValue
		return value, value, nil
	case "find_and_replace":
		value := strings.ReplaceAll(currentValue, action.Find, action.Replace)
		return value, value, nil
	case "multiply":
		left, ok := parseNumber(currentValue)
		if !ok {
			// Gracefully skip non-numeric or empty cells without failing the batch update
			return currentValue, nil, nil
		}
		right, ok := parseNumber(action.Value)
		if !ok {
			return "", nil, fmt.Errorf("无法将乘数 '%v' 解析为数字", action.Value)
		}
		value := left * right
		return strconv.FormatFloat(value, 'f', -1, 64), value, nil
	default:
		return "", nil, fmt.Errorf("不支持的更新动作: %s", action.Type)
	}
}

func parseNumber(value any) (float64, bool) {
	switch v := value.(type) {
	case float64:
		return v, true
	case float32:
		return float64(v), true
	case int:
		return float64(v), true
	case int8:
		return float64(v), true
	case int16:
		return float64(v), true
	case int32:
		return float64(v), true
	case int64:
		return float64(v), true
	case uint:
		return float64(v), true
	case uint8:
		return float64(v), true
	case uint16:
		return float64(v), true
	case uint32:
		return float64(v), true
	case uint64:
		return float64(v), true
	}

	text := strings.TrimSpace(fmt.Sprintf("%v", value))
	if text == "" || text == "<nil>" {
		return 0, false
	}
	text = strings.ReplaceAll(text, ",", "")
	text = strings.ReplaceAll(text, "，", "")
	text = strings.ReplaceAll(text, "￥", "")
	text = strings.ReplaceAll(text, "$", "")

	if num, err := strconv.ParseFloat(text, 64); err == nil {
		return num, true
	}

	match := numericValueRegex.FindString(text)
	if match == "" {
		return 0, false
	}
	num, err := strconv.ParseFloat(match, 64)
	if err != nil {
		return 0, false
	}
	return num, true
}

func firstNonEmpty(values ...string) string {
	for _, value := range values {
		if strings.TrimSpace(value) != "" {
			return value
		}
	}
	return ""
}
