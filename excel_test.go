package go_struct_excel

import (
	"fmt"
	"os"
	"reflect"
	"testing"
)

type foo struct {
	Name    string          `excel:"姓名"`
	Age     *int            `excel:"年龄,allowempty"`
	Height  int             `excel:"身高,font{color:ff0000 size:16}"`
	Holiday map[string]bool `excel:"假期,expand:regexp(^\\d{4}-\\d{2}-\\d{2}$)"`
}

func (f foo) GatherHeaderRows() int {
	return 1
}

func (f foo) GatherHeader(sheet *Sheet) error {
	style, _ := sheet.GetCenterStyle()
	sheet.Excel.SetCellValue(sheet.SheetName, "A2", "个人信息")
	sheet.Excel.MergeCell(sheet.SheetName, "A2", "C2")
	sheet.Excel.SetCellStyle(sheet.SheetName, "A2", "C2", style)
	sheet.Excel.SetCellValue(sheet.SheetName, "D2", "假期信息")
	sheet.Excel.MergeCell(sheet.SheetName, "D2", "I2")
	return nil
}

func (f foo) Remarks() (string, float64, float64) {
	return `.特别注意：导入数值加逗号格式，很经常被Excel带成数值，可在前面加个'号，或设置单元格格式为文本			
			
.导入规则：
.全局名称不允许重复			
.各种包含类型枚举：可为空表示不定向，或输入：不限、包含、不包含			
.小时定向：周一到周六枚举：0~6，周日枚举值0，如定向周日22、23点，周二0、1点，格式 0:22,23|2:0,1			
.媒体平台：VendorName，请输入全称			
.流量定向：请输入ID`, 150, 85
}

func TestNewExcel(t *testing.T) {
	excel := NewExcel()
	defer excel.File.Close()
	sheet, err := excel.AddSheet("hello")
	if err != nil {
		t.Error(err)
		return
	}
	// 1994-05-25
	data := make([]*foo, 0)
	age := 28
	data = append(data, &foo{
		Name:   "h",
		Age:    &age,
		Height: 181,
		Holiday: map[string]bool{
			"2022-01-27": false,
			"2022-01-28": true,
			"2022-01-29": true,
		},
	}, &foo{
		Name:   "o",
		Age:    &age,
		Height: 182,
		Holiday: map[string]bool{
			"2022-01-27": true,
			"2022-01-28": true,
			"2022-01-30": true,
			"2022-02-09": true,
			"2022-12-09": true,
		},
	})

	if err = sheet.AddData(data); err != nil {
		t.Error(err)
		return
	}

	if err = excel.SaveAs("helloword"); err != nil {
		t.Errorf("文件保存失败: %s", err.Error())
		return
	}
	dir, _ := os.Getwd()
	fmt.Println("当前路径：", dir)
	return
}

func TestParseExcelHeaderTag(t *testing.T) {
	f := foo{
		Name: "刘宁",
	}

	val := reflect.ValueOf(f)
	fmt.Println(val.Type().Field(0).Name)
}

func TestReadData(t *testing.T) {
	//m := make(map[string]int, 0)
	//fmt.Println(reflect.TypeOf(m).Elem().Kind())

	excel, err := OpenExcel("helloword.xlsx")
	if err != nil {
		t.Error(err)
	}

	sheet, err := excel.OpenSheet("hello")
	if err != nil {
		t.Error(err)
	}

	if data, err := sheet.ReadData(foo{}); err != nil {
		t.Error(err)
	} else if d, ok := data.([]*foo); ok {
		//fmt.Printf("%+v", )
		for _, v := range d {
			fmt.Println(v.Name)
		}
	}
}
