package convert

import (
	"fmt"
	"reflect"
	"strconv"
)

var (
	BaseConvertMap = map[string]func(reflect.Value) interface{}{
		"FormatDataToFloat64": FormatDataToFloat64,
	}
)

// FormatDataToFloat64
//
//	@Description:  数据转换
//	@param req  需要转换的数据
//	@return string 转后后的数据
func FormatDataToFloat64(req reflect.Value) interface{} {
	resp := 0.0
	switch req.Kind() {
	case reflect.Int:
		resp = float64(req.Int())
	case reflect.String:
		fl, err := strconv.ParseFloat(req.String(), 64)
		if err != nil {
			fmt.Println("Error converting string to float64")
			return req.String()
		}
		resp = fl
	case reflect.Float64:
		resp = req.Float()
	default:
		fmt.Println("Value is of a different type")
	}
	return fmt.Sprintf("%.2f", resp)
}
