# go-struct-excel

1. struct支持导出为excel
2. excel导入为struct
3. 表头支持扩展，如：日期表头不确定长度
4. excel第一行为备注
5. excel表头进行汇总
6. 标记某行为特殊颜色
7. 如果字段为空，就不生成该表头

> 表头不支持重复

实际效果：

![](helloworld.png)

# 安装

```shell
go get github.com/douyacun/go-struct-excel
```

# 导出 excel

```go

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
`, 3, 8
}

func TestNewExcel(t *testing.T) {
	excel := NewExcel()
	defer excel.File.Close()
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

	if err := excel.AddSheet("hello").AddData(data); err != nil {
		t.Error(err)
		return
	}

	if err := excel.SaveAs("helloword"); err != nil {
		t.Errorf("文件保存失败: %s", err.Error())
		return
	}
	dir, _ := os.Getwd()
	fmt.Println("当前路径：", dir)
	return
}

```

excel tag说明：

- tag英文逗号分隔，第一个作为表头名称，其他没有顺序要求
- `allowempty`: 表头在场景一中需要展示，其他不需要。字段为指针类型，tag标记为 `allowmepty`
- `font`: 设定此列富文本样式，支持 `{color:ff0000 size:16 italic:true blod:true underline:true}`
- `expand`: 自动扩展表头，支持正则匹配表头，`expand:regexp(^\\d{4}-\\d{2}-\\d{2}$)"`， 其中内置正则 
  + `expand:date`: 2022-06-18
  + `expand:datetime`: 2022-06-18 08:27:39
  + `expand:month`: 2022-06

表头备注：
```go
type ExcelRemarks interface {
	Remarks() (remark string, row, col int) // 占用几行几列
}
```

汇总表头
```go
type ExcelGatherHeader interface {
	GatherHeaderRows() int           // 汇总表头占几行，不包括字段行
	GatherHeader(sheet *Sheet) error // 汇总表头合并单元格，单元格样式需要自己实现
}
```


## 导入用法

```go
func TestReadData(t *testing.T) {
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
		if str, err := json.Marshal(d); err != nil {
			t.Error(err)
		} else {
			fmt.Println(string(str))
		}
	}
}

```

输出：
```shell
=== RUN   TestReadData
[{"name":"h","age":28,"height":181,"holiday":{"2022-01-27":false,"2022-01-28":true,"2022-01-29":true}},{"name":"o","age":28,"height":182,"holiday":{"2022-01-27":true,"2022-01-28":true,"2022-01-29":false,"2022-01-30":true,"2022-02-09":true,"2022-12-09":true}}]
--- PASS: TestReadData (0.00s)
PASS
```