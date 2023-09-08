package structexcel

import (
	"encoding/json"
	"fmt"
	"os"
	"reflect"
	"testing"
)

type foo struct {
	Name    string          `excel:"姓名" json:"name"`
	Age     *int            `excel:"年龄,allowempty" json:"age"`
	Height  int             `excel:"身高,font{color:ff0000 size:16}" json:"height"`
	Holiday map[string]bool `excel:"假期,expand:regexp(^\\d{4}-\\d{2}-\\d{2}$)" json:"holiday"`
}

func (f foo) GatherHeaderRows() int {
	return 1
}

func (f foo) GatherHeader(sheet *Sheet) error {
	style, _ := sheet.GetCenterStyle()
	sheet.SetCellValueByName("A7", "个人信息")
	sheet.MergeCellByName("A7", "C7")
	sheet.SetCellStyleByName("A7", "C7", style)
	sheet.SetCellValueByName("D7", "假期信息")
	sheet.MergeCellByName("D7", "I7")
	return nil
}

func (f foo) Remarks() (string, int, int) {
	return `.特别注意：导入数值加逗号格式，很经常被Excel带成数值，可在前面加个'号，或设置单元格格式为文本			
			
.导入规则：
.全局名称不允许重复			
.各种包含类型枚举：可为空表示不定向，或输入：不限、包含、不包含			
`, 6, 9
}

func TestNewPositionExcel(t *testing.T) {
	excel := NewExcel("position.xlsx")
	defer excel.File.Close()
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

	sheet, err := excel.AddSheet("hello")
	if err != nil {
		t.Fatal(err)
	}
	sheet.SetPosition(2, 1)
	if err = sheet.AddData(data); err != nil {
		t.Error(err)
		return
	}
	if err = excel.SaveAs(); err != nil {
		t.Errorf("文件保存失败: %s", err.Error())
		return
	}
	dir, _ := os.Getwd()
	fmt.Println("当前路径：", dir)
}

func TestNewExcel(t *testing.T) {
	excel := NewExcel("helloworld.xlsx")
	defer excel.File.Close()
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

	sheet, err := excel.AddSheet("hello")
	if err != nil {
		t.Fatal(err)
	}
	if err = sheet.AddData(data); err != nil {
		t.Error(err)
		return
	}
	if err = excel.SaveAs(); err != nil {
		t.Errorf("文件保存失败: %s", err.Error())
		return
	}
	dir, _ := os.Getwd()
	fmt.Println("当前路径：", dir)
}

func TestParseExcelHeaderTag(t *testing.T) {
	f := foo{
		Name: "douyacun",
	}

	val := reflect.ValueOf(f)
	fmt.Println(val.Type().Field(0).Name)
}

func TestReadData(t *testing.T) {
	excel, err := OpenExcel("helloworld.xlsx")
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
		if str, err := json.Marshal(d); err != nil {
			t.Error(err)
		} else {
			fmt.Println(string(str))
		}
	}
}
