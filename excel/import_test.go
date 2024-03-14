package excel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"testing"
)

// 导入测试
func TestImport(t *testing.T) {
	// 这里需要注意：excelize 不支持xls格式的excel文件，如果是xls格式，会报错提示：不支持的工作簿文件格式
	f, err := excelize.OpenFile(basePath() + "1.xlsx")
	if err != nil {
		fmt.Println("文件打开失败")
	}
	importList := []TestStruct{}
	err = ImportExcel(f, &importList, 4, 5)
	if err != nil {
		fmt.Println(err)
	}
	for _, t := range importList {
		fmt.Println(t) // 输出导入结果
	}
}

func basePath() (res string) {
	res, _ = os.Getwd()
	return res + string(os.PathSeparator)
}
