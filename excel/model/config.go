package model

import "reflect"

type ExportConfig struct {
	// Sheet
	SheetName string
	// 操作类
	E *Excel
	// 标题
	Title string
	// 是否使用默认样式
	IsDefaultStyle bool
	// 数据
	ExportData interface{}
	// 表头控制
	ChangeHead map[string]string
	DynamicColConfig
}

// DynamicColConfig 动态表头配置
type DynamicColConfig struct {
	// 需要处理的字段
	Fields string
	// 控制是否忽略
	IsIgnore bool
}

type DataRowBuildReq struct {
	// 操作类
	Excel      *Excel
	SheetName  string
	EndColName string
	// 开始行数
	StartRow int
	// 是否使用默认样式
	IsDefaultStyle bool
	// 数据
	DataValue reflect.Value
	DynamicColConfig
}

type MapExportConfig struct {
	Heads          interface{}
	ExportData     []map[string]interface{}
	SheetName      string
	Title          string
	IsDefaultStyle bool
}

func NewExportConfig() *ExportConfig {
	return &ExportConfig{
		E: NewExcel(),
		DynamicColConfig: DynamicColConfig{
			Fields:   "",
			IsIgnore: false,
		},
		IsDefaultStyle: false,
		SheetName:      "Sheet1",
	}
}

func (e *ExportConfig) SetExportData(exportData interface{}) *ExportConfig {
	e.ExportData = exportData
	return e
}

func (e *ExportConfig) SetTitle(title string) *ExportConfig {
	e.Title = title
	return e
}

func (e *ExportConfig) SetChangeHead(changeHead map[string]string) *ExportConfig {
	e.ChangeHead = changeHead
	return e
}
