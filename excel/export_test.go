package excel

import (
	"fmt"
	"go-excel/excel/model"
	"os"
	"testing"
)

// TestStruct 测试结构体
type TestStruct struct {
	Id       string `excel:"name:用户账号;"`
	Name     string `excel:"name:用户姓名;"`
	Email    string `excel:"name:用户邮箱;width:25;"`
	Com      string `excel:"name:所属公司;"`
	Dept     bool   `excel:"name:所在部门;replace:false_超级管理员,true_普通用户;"`
	RoleName string `excel:"name:角色名称;replace:1_超级管理员,2_普通用户;"`
	Remark   int    `excel:"name:备注;replace:1_超级管理员,2_普通用户;width:40;"`
}

type TestConvertStruct struct {
	Id   string `excel:"name:用户账号;convert:FormatDataToFloat64;"`
	Name string `excel:"name:用户姓名;"`
}

// =============================== excel 导入、导出测试 ===============================
// 要导出的列表
var testList = []TestStruct{
	{"fuhua", "符华", "fuhua@123.com", "太虚剑派", false, "1", 1},
	{"baiye", "白夜", "baiye@123.com", "天命科技有限公司", false, "2", 1},
	{"chiling", "炽翎", "chiling@123.com", "太虚剑派", false, "2", 2},
	{"yunmo", "云墨", "yunmo@123.com", "太虚剑派", false, "1", 2},
	{"yuelun", "月轮", "yuelun@123.com", "天命科技有限公司", false, "1", 1},
	{"xunyu", "迅羽",
		"xunyu@123.com哈哈哈哈哈哈哈哈这里是最大行高测试哈哈哈哈哈哈哈哈这11111111111里是最大行高测试哈哈哈哈哈哈哈哈这里是最大行高测试",
		"天命科技有限公司", true, "2",
		124},
}

// 单个sheet导出
func TestExportSheet(t *testing.T) {
	// 获取导出的数据
	changeHead := map[string]string{"Id": "账号", "Name": "真实姓名"}
	//单个sheet导出
	f, err := NormalDynamicExport("Sheet1", "用户信息", "Id,Email,", true, true, testList, changeHead)
	//f, err := NormalDynamicExport(testList, "Sheet1", "用户信息", "", true, false, changeHead)
	if err != nil {
		fmt.Println(err)
		return
	}
	f.Path = basePath() + "单个sheet导出.xlsx"
	if err := f.Save(); err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println("Excel文件已生成")
}

// 测试转换函数
func TestConvert(t *testing.T) {
	// 获取导出的数据
	dataList := []TestConvertStruct{
		{"1.23333", "白夜"},
		{"测试数据", "黑日"},
	}
	//单个sheet导出
	f, err := NormalDynamicExport("Sheet1", "用户信息", "", true, false, dataList, nil)
	if err != nil {
		fmt.Println(err)
		return
	}
	f.Path = basePath() + "转换函数导出.xlsx"
	if err := f.Save(); err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println("Excel文件已生成")
}

// 多个sheet导出
func TestExportSheets(t *testing.T) {
	// 获取导出的数据
	changeHead := map[string]string{"Id": "账号", "Name": "真实姓名"}
	// 多个sheet导出
	e := model.ExcelInit()
	for i := 0; i < 3; i++ {
		sheet := "Sheet" + fmt.Sprintf("%d", i+1)
		title := "用户信息" + fmt.Sprintf("%d", i+1)
		fmt.Println(sheet)
		// 其实就是相当于普通sheet导出，只不过是每个sheet分别传对应的数据过去
		err := ExportExcel(sheet, title, "", true, false, testList, changeHead, e)
		if err != nil {
			fmt.Println(err)
			return
		}
	}
	e.F.Path = basePath() + "多个sheet导出.xlsx"
	if err := e.F.Save(); err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println("Excel文件已生成")
}

// 基于map的导出
func TestExportMap(t *testing.T) {
	// 表头
	header := []string{"Name", "Email", "City"}
	// map数据
	data := []map[string]interface{}{
		{"Name": "符华", "Email": "fuhua@123.com", "City": "惠州"},
		{"Name": "陈悦", "Email": "chenyue@qq.com", "City": "深圳"},
		{"Name": "鹤熙", "Email": "hexi@123.com", "City": "广州"},
	}
	f, err := MapExport(header, data, "Sheet1", "", false)
	if err != nil {
		fmt.Println("导出失败", err)
		return
	}
	// 保存文件
	err = f.SaveAs(basePath() + "map导出.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println("Excel文件已成功导出")
}

// 复杂表头导出
func TestExportTree(t *testing.T) {
	// map导出
	/*header := [][]string{
		{"一级表头1", "一级表头1", "一级表头1", "一级表头1", "一级表头1", "一级表头1", "一级表头1", "一级表头1", "一级表头2", "一级表头2", "一级表头2", "一级表头2"},
		{"二级表头1", "二级表头1", "二级表头1", "二级表头1", "二级表头1", "二级表头2", "二级表头2", "二级表头2", "二级表头3", "二级表头3", "二级表头3", "二级表头3"},
		{"三级表头1", "三级表头1", "三级表头2", "三级表头3", "三级表头3", "三级表头4", "三级表头4", "三级表头5", "三级表头6", "三级表头6", "三级表头6", "三级表头7"},
		{"四级表头1", "四级表头2", "四级表头3", "四级表头4", "四级表头5", "四级表头6", "四级表头7", "四级表头8", "四级表头9", "四级表头10", "四级表头11", "四级表头12"},
	}
	var data = []map[string]interface{}{
		{"四级表头1": "1", "四级表头2": "1", "四级表头3": "1", "四级表头4": "4", "四级表头5": "5", "四级表头6": "6", "四级表头7": "7", "四级表头8": "8", "四级表头9": "9", "四级表头10": "10", "四级表头11": "11", "四级表头12": "12"},
		{"四级表头1": "11", "四级表头2": "22", "四级表头3": "33", "四级表头4": "44", "四级表头5": "55", "四级表头6": "66", "四级表头7": "77", "四级表头8": "88", "四级表头9": "99", "四级表头10": "100", "四级表头11": "111", "四级表头12": "122"},
		{"四级表头1": "111", "四级表头2": "222", "四级表头3": "333", "四级表头4": "444", "四级表头5": "555", "四级表头6": "666", "四级表头7": "777", "四级表头8": "888", "四级表头9": "999", "四级表头10": "1000", "四级表头11": "1111", "四级表头12": "1222"},
	}
	f, err := excel.MapExport(header, data, "Sheet1", "这里是标题", false)
	// 合并表头单元格
	// 没有title时，表头从第一行开始合并startRowNum=1；
	// 有title时，表头从第二行开始合并，startRowNum=2；
	// endRowNum=6，表示内容行开始不再需要合并
	excel.HorizontalMerge(f, "Sheet1", 2, 6)*/

	// 结构体导出（自定义表头）
	header := [][]string{
		{"基本信息", "基本信息", "基本信息", "基本信息", "基本信息", "其他信息", "其他信息"},
		{"用户信息", "用户信息", "用户信息", "部门信息", "部门信息", "角色信息", "备注"},
		{"用户信息", "用户信息", "用户信息", "所属公司", "所在部门", "角色信息", "备注"},
		{"用户账号", "用户姓名", "用户邮箱", "所属公司", "所在部门", "角色名称", "备注"},
	}
	var data = []TestStruct{
		{"云墨", "云墨", "云墨", "太虚剑派", false, "1", 1},
		{"fuhua", "炽翎", "炽翎", "炽翎", false, "1", 1},
		{"月轮", "月轮", "yuelun@123.com", "yuelun@123.com", true, "2", 2},
		{"admin", "admin", "admin", "admin", false, "1", 2},
		{"符华", "符华", "admin@123.com", "天命科技有限公司", false, "1", 1},
		{"chenyue", "chenyue", "chenyue@123.com", "天命科技有限公司", true, "2", 124},
		{"鹤熙", "鹤熙", "鹤熙", "天命科技有限公司", true, "2", 124},
	}
	f, err := CustomHeaderExport("Sheet1", "这里是标题", true, header, data)
	if err != nil {
		panic(err)
	}
	// 合并表头单元格
	HorizontalMerge(f, "Sheet1", 2, 5)
	// 纵向合并数据行内容
	VerticalMerge(f, "Sheet1", 1, nil)
	// 保存文件
	err = f.SaveAs(basePath() + "复杂表头导出.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println("Excel文件已生成")
}

// 横向合并单元格导出
func TestExportHorizontal(t *testing.T) {
	// 结构体数据导出
	var data = []TestStruct{
		{"云墨", "云墨", "云墨", "太虚剑派", false, "1", 1},                             // A2:C2 , E2:G2
		{"fuhua", "炽翎", "炽翎", "炽翎", false, "1", 1},                            // B3:D3 , E2:G2
		{"月轮", "月轮", "yuelun@123.com", "yuelun@123.com", true, "2", 2},        // A4:B4 , C4:D4 , E4:G4
		{"admin", "admin", "admin", "admin", false, "1", 2},                   // A5:D5 , E5:F5
		{"符华", "符华", "admin@123.com", "天命科技有限公司", false, "1", 1},              // A6:B6 , E6:G6
		{"chenyue", "chenyue", "chenyue@123.com", "天命科技有限公司", true, "2", 124}, // A7:B7 , E7:F7
		{"鹤熙", "鹤熙", "鹤熙", "天命科技有限公司", true, "2", 124},                        // A8:C8 , E8:F8
	}
	f, err := NormalDynamicExport("Sheet1", "", "", false, false, data, nil)
	// map数据导出
	/*header := []string{"账号", "姓名", "部门", "角色", "备注"}
	var data = []map[string]interface{}{
		{"账号": "符华", "姓名": "符华", "部门": "符华", "角色": "太虚剑派", "备注": "太虚剑派"},           // A2:C2 , D2:E2
		{"账号": "云墨", "姓名": "云墨", "部门": "太虚剑派", "角色": "太虚剑派", "备注": "太虚剑派"},         // A3:B3 , C3:E3
		{"账号": "月轮", "姓名": "月轮", "部门": "天命科技有限公司", "角色": "天命科技有限公司", "备注": "太虚剑派"}, // A4:B4 , C4:D4
		{"账号": "鹤熙", "姓名": "天命科技有限公司", "部门": "天命科技有限公司", "角色": "鹤熙", "备注": "鹤熙"},   // B5:C5 , D5:E5
	}
	f, err := excel.MapExport(header, data, "Sheet1", "", false)*/
	if err != nil {
		fmt.Println(err)
		return
	}
	// 横向合并单元格：没有标题只有一行表头，所以内容从第二行开始合并 startRowNum=2
	HorizontalMerge(f, "Sheet1", 2, -1) // endRowNum = -1，表示全部每一行都需要合并
	//excel.HorizontalMerge(f, "Sheet1", 2, 6) // endRowNum = 6，表示第6行开始，后面的行不进行合并（包括第6行）
	// 保存文件
	err = f.SaveAs(basePath() + "横向合并导出.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println("Excel文件已生成")
}

// 纵向合并单元格导出
func TestExportVertical(t *testing.T) {
	// map数据
	//header := []string{"所属公司", "所属部门", "姓名", "职位", "联系电话"}
	header := [][]string{
		{"基本信息", "基本信息", "基本信息", "基本信息", "基本信息"},
		{"部门信息", "部门信息", "用户信息", "用户信息", "用户信息"},
		{"所属公司", "所属部门", "姓名", "职位", "联系电话"},
	}
	var data = []map[string]interface{}{
		{"所属公司": "太虚剑派", "所属部门": "开发部", "姓名": "符华", "职位": "部门经理", "联系电话": "1321"},
		{"所属公司": "太虚剑派", "所属部门": "开发部", "姓名": "炽翎", "职位": "部门经理", "联系电话": "46545"},
		{"所属公司": "太虚剑派", "所属部门": "开发部", "姓名": "云墨", "职位": "部门主管", "联系电话": "13212"},
		{"所属公司": "太虚剑派", "所属部门": "财务部", "姓名": "赤鸢", "职位": "部门主管", "联系电话": "13212"},
		{"所属公司": "太虚剑派", "所属部门": "财务部", "姓名": "华", "职位": "员工", "联系电话": "13212"},
		{"所属公司": "天命科技", "所属部门": "财务部", "姓名": "白夜", "职位": "组长", "联系电话": "dfgdfg"},
		{"所属公司": "天命科技", "所属部门": "研发部", "姓名": "月轮", "职位": "组长", "联系电话": "45645"},
		{"所属公司": "天命科技", "所属部门": "研发部", "姓名": "迅羽", "职位": "组长", "联系电话": "45645"},
	}
	f, err := MapExport(header, data, "Sheet1", "", false)
	if err != nil {
		panic(err)
	}
	needColIndex := []int{1, 2} // 需要合并的列号，比如只需要合并第一列和第二列
	// 横向合并表头行内容
	HorizontalMerge(f, "Sheet1", 1, 4)
	// 纵向合并数据行内容
	VerticalMerge(f, "Sheet1", 0, needColIndex)

	// 结构体数据导出
	/*var data = []TestStruct{
		{"赤鸢", "云墨", "chiyuan", "太虚剑派", false, "1", 1},
		{"赤鸢", "炽翎", "chiyuan", "太虚剑派", false, "1", 1},
		{"赤鸢", "炽翎", "chiyuan", "太虚剑派", true, "2", 2},
		{"赤鸢", "云墨", "chiyuan", "太虚剑派", false, "1", 2},
		{"符华", "白夜", "fuhua", "天命科技", false, "1", 1},
		{"符华", "月轮", "fuhua", "天命科技", true, "2", 121},
		{"符华", "迅羽", "fuhua", "天命科技", true, "2", 121},
	}
	f, err := excel.NormalDynamicExport("Sheet1", "", "", false, false, data, nil)
	if err != nil {
		panic(err)
	}
	// 纵向合并内容单元格
	excel.VerticalMerge(f, "Sheet1", 0, nil)*/

	// 保存文件
	err = f.SaveAs(basePath() + "纵向合并导出.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	fmt.Println("Excel文件已生成")
}

// 模板导出
func TestExportTemplate(t *testing.T) {
	/** 关于模板导出：
	excelize并没有提供类似Java easypoi的模板指令功能，假如需要根据excel模板导出excel文件的话，我现在的解决办法就是通过
	text/template 或 html/template 来解析模板文件，然后填充数据，输出到excel文件。
	但是这两个库并不能直接解析excel模板，它们只能解析文本格式的文件，不能解析二进制格式的文件。
	所以我们准备好了一个excel模板时，还需要将这个excel文件另存为xml文件；
	有了xml文件也不能直接用，有些地方的数据可能需要处理一下，比如遍历list的时候，需要将range遍历放到正确的位置。
	检查好之后，再使用 template.ParseFiles 函数解析这个xml文件，接下来就可以填充数据了。填充好后，将渲染模板并输出为excel文件。
	*/
	// 准备结构体数据
	var list = []TestStruct{
		{"fuhua", "符华", "fuhua@123.com", "太虚剑派", false, "1", 1},
		{"baiye", "白夜", "baiye@123.com", "天命科技有限公司", false, "2", 1},
		{"chiling", "炽翎", "chiling@123.com", "太虚剑派", false, "2", 2},
		{"yunmo", "云墨", "yunmo@123.com", "太虚剑派", false, "1", 2},
		{"yuelun", "月轮", "yuelun@123.com", "天命科技有限公司", false, "1", 1},
		{"xunyu", "迅羽", "xunyu@123.com", "天命科技有限公司", true, "2", 124},
	}
	// 准备map数据
	data := map[string]any{
		"name":     "符华",
		"sex":      "女",
		"birthday": "2000-10-01",
		"jg":       "广东省",
		"sfz":      "123456789111111111",
		"gzdw":     "天命科技",
		"email":    "fuhua@123.com",
		"phone":    "13188888888",
		"workList": []map[string]any{
			{"date": "2019-01-01~2020-10-01", "unit": "逐火之蛾", "post": "普通员工", "sfzz": "否"},
			{"date": "2020-10-20~2022-09-15", "unit": "太虚剑派", "post": "CEO", "sfzz": "否"},
			{"date": "2022-10-01~至今", "unit": "天命科技", "post": "经理", "sfzz": "是"},
		},
		"list": list,
	}
	templatePath := basePath() + "template" + string(os.PathSeparator) + "测试模板导出.xml" // 模板路径
	outPath := basePath() + "output.xlsx"                                             // 输出路径
	err := TemplateExport(templatePath, outPath, data)
	if err != nil {
		fmt.Println(err.Error())
	}
	fmt.Println("导出成功！")
	// 将excel下载到浏览器，第一个参数是模板路径、第二个参数是下载的文件名称（要带上后缀）、第三个参数是模板数据、第四个参数是http响应
	// excel.DownLoadByTemplate(templatePath,"output.xlsx",data,res)
}
