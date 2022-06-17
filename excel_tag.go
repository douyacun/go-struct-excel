package structexcel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"regexp"
	"strconv"
	"strings"
)

type excelHeaderField struct {
	Col int

	fieldName   string
	headerName  string
	allowEmpty  bool
	expand      bool
	expandRegex *regexp.Regexp
	skip        bool
	level       int
	split       string
	font        *excelize.Font
}

type excelHeaderNode struct {
	Start    int
	Height   int
	Width    int
	Name     string
	Children []*excelHeaderNode
}

func ParseExcelHeaderTag(tag string, col int) *excelHeaderField {
	h := &excelHeaderField{
		Col:   col,
		level: 1,
	}
	if tag == "" || tag == "-" {
		h.skip = true
		return h
	}

	tagList := strings.Split(tag, ",")
	for k, v := range tagList {
		if v == "allowempty" {
			h.allowEmpty = true
		}

		if strings.HasPrefix(v, "expand:") {
			h.expand = true
			h.parseExpand(v)
		}

		if strings.HasPrefix(v, "split:") {
			h.split = v[6:]
		}

		if strings.HasPrefix(v, "font{") {
			h.font = &excelize.Font{}
			prop := v[5 : len(v)-1]
			for _, f := range strings.Split(prop, " ") {
				rich := strings.Split(f, ":")
				if len(rich) != 2 {
					panic("无效富文本tag：font{size:14 color:FF0000} 当前：" + prop)
				}
				switch rich[0] {
				case "size":
					s, err := strconv.ParseFloat(rich[1], 64)
					if err != nil {
						panic(fmt.Sprintf("excel font fontsize parse error: %s", err.Error()))
					}
					h.font.Size = s
				case "bold":
					if rich[1] == "true" {
						h.font.Bold = true
					}
				case "color":
					if len(rich[1]) == 6 {
						h.font.Color = rich[1]
					} else {
						panic("无效富文本tag color：font{size:14 color:FF0000} 当前：" + rich[1])
					}
				case "italic":
					if rich[1] == "true" {
						h.font.Italic = true
					}
				case "family":
					if rich[1] != "" {
						h.font.Family = rich[1]
					}
				case "strike":
					if rich[1] == "true" {
						h.font.Strike = true
					}
				case "underline":
					if rich[1] == "single" {
						h.font.Underline = "single"
					}
				}
			}
		}

		if k == 0 {
			h.headerName = v
		}
	}

	return h
}

func (e *excelHeaderField) parseExpand(expand string) {
	exp := expand[7:]
	if strings.HasPrefix(exp, "date") {
		e.expandRegex = regexp.MustCompile(`^\d{4}-\d{2}-\d{2}$`)
	} else if strings.HasPrefix(exp, "datetime") {
		e.expandRegex = regexp.MustCompile(`^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$`)
	} else if strings.HasPrefix(exp, "month") {
		e.expandRegex = regexp.MustCompile(`^\d{4}-\d{2}$`)
	} else if strings.HasPrefix(exp, "regexp") {
		e.expandRegex = regexp.MustCompile(exp[7 : len(exp)-1])
	}
}

func (e excelHeaderField) IsSkip() bool {
	return e.skip
}

type excelHeaderMap map[string]*excelHeaderField

type excelHeaderSlice []*excelHeaderField

func (x excelHeaderSlice) Len() int           { return len(x) }
func (x excelHeaderSlice) Less(i, j int) bool { return x[i].Col < x[j].Col }
func (x excelHeaderSlice) Swap(i, j int)      { x[i], x[j] = x[j], x[i] }

func (x excelHeaderSlice) getFieldMap() excelHeaderMap {
	res := make(excelHeaderMap, 0)
	for _, v := range x {
		if v.level == 1 {
			res[v.fieldName] = v
		} else {
			res[v.headerName] = v
		}
	}
	return res
}

func (x excelHeaderSlice) getHeaderMap() excelHeaderMap {
	res := make(excelHeaderMap, 0)
	for _, v := range x {
		res[v.headerName] = v
	}
	return res
}

func (x excelHeaderSlice) getExpandHeaderSlice() excelHeaderSlice {
	res := make(excelHeaderSlice, 0)
	for _, v := range x {
		if v.expand {
			res = append(res, v)
		}
	}
	return res
}

func (x excelHeaderSlice) getColHeaderMap() map[int]*excelHeaderField {
	res := make(map[int]*excelHeaderField)
	for _, v := range x {
		res[v.Col] = v
	}
	return res
}
