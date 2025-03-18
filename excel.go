package structexcel

import (
	"bytes"
	"fmt"
	"io"
	"net/http"
	"reflect"
	"sort"
	"strconv"
	"strings"

	"github.com/pkg/errors"
	"github.com/xuri/excelize/v2"
)

type ExcelRemarks interface {
	Remarks() (remark string, row, col int)
}

type ExcelGatherHeader interface {
	GatherHeaderRows() int           // 汇总表头占几行，不包括字段行
	GatherHeader(sheet *Sheet) error // 汇总表头合并单元格，单元格样式需要自己实现
}

type Excel struct {
	File        *excelize.File
	activeSheet int
	Filename    string
}

func NewExcel(filename string) *Excel {
	f := excelize.NewFile()
	return &Excel{
		File:        f,
		activeSheet: -1,
		Filename:    filename,
	}
}

func OpenExcel(filename string) (*Excel, error) {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		return nil, errors.Wrap(err, "excel")
	}
	return &Excel{File: f, activeSheet: -1}, nil
}

func OpenReader(r io.Reader) (*Excel, error) {
	f, err := excelize.OpenReader(r)
	if err != nil {
		return nil, errors.Wrap(err, "excel")
	}
	return &Excel{File: f, activeSheet: -1}, nil
}

func OpenFromUrl(url string) (*Excel, error) {
	resp, err := http.Get(url)
	if err != nil {
		return nil, errors.Wrapf(err, "excel读取(%s)失败了", url)
	}
	defer resp.Body.Close()
	if resp.StatusCode < 200 && resp.StatusCode > 299 {
		return nil, errors.Errorf("excel读取(%s)失败了", url)
	}
	return OpenReader(resp.Body)
}

func (e *Excel) Close() error {
	return e.File.Close()
}

// Bytes 吐字节
func (e *Excel) Bytes() ([]byte, error) {
	buf := bytes.NewBuffer(nil)
	if err := e.File.Write(buf); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

// Response proto
//func (e *Excel) Response(filename string) (*commonProto.Excel, error) {
//	bt, err := e.Bytes()
//	if err != nil {
//		return nil, err
//	}
//	return &commonProto.Excel{
//		FileName: filename,
//		Raw:      bt,
//	}, nil
//}

// SaveAs 保存为文件
func (e *Excel) SaveAs() error {
	if err := e.File.SaveAs(e.Filename); err != nil {
		return err
	}
	return nil
}

// AddSheet 添加sheet
func (e *Excel) AddSheet(name string) (*Sheet, error) {
	index, err := e.File.NewSheet(name)
	if err != nil {
		return nil, err
	}
	if e.activeSheet == -1 {
		e.File.SetActiveSheet(index)
		// 移除默认sheet1，好像没办法重命名sheet
		e.File.DeleteSheet("Sheet1")
	}
	return &Sheet{
		Excel:            e.File,
		SheetName:        name,
		autoCreateHeader: true,
		row:              0,
		col:              0,
		header:           make(excelHeaderSlice, 0),
	}, nil
}

func (e *Excel) OpenSheet(sheetName string) (*Sheet, error) {
	index, err := e.File.GetSheetIndex(sheetName)
	if err != nil {
		return nil, err
	}
	if index == -1 {
		return nil, errors.Errorf("%s sheet缺失", sheetName)
	}
	return &Sheet{
		index:            index,
		Excel:            e.File,
		SheetName:        sheetName,
		autoCreateHeader: false,
		row:              0,
		col:              0,
	}, nil
}

func (e *Excel) OpenSheetByMap(sheetName string) (*Sheet, error) {
	sheetMap := e.GetSheetMap()

	for i, name := range sheetMap {
		if name == sheetName {
			return &Sheet{
				index:            i,
				Excel:            e.File,
				SheetName:        name,
				autoCreateHeader: false,
				row:              0,
				col:              0,
			}, nil
		}
	}
	return nil, errors.Errorf("%s sheet缺失", sheetName)
}

func (e *Excel) GetSheetMap() map[int]string {
	return e.File.GetSheetMap()
}

func (e *Excel) Response(w http.ResponseWriter) error {
	header := w.Header()

	byt, err := e.Bytes()
	if err != nil {
		return err
	}
	header["Accept-Length"] = []string{strconv.Itoa(len(byt))}
	header["Content-Type"] = []string{"application/vnd.ms-excel"}
	header["Access-Control-Expose-Headers"] = []string{"Content-Disposition"}
	header["Content-Disposition"] = []string{fmt.Sprintf("attachment; filename=\"%s\"", e.Filename)}
	w.Write(byt)
	return nil
}

type Sheet struct {
	Excel     *excelize.File
	SheetName string

	header           excelHeaderSlice
	index            int // sheet index
	autoCreateHeader bool
	hasRemarks       bool
	row              int
	col              int
}

func (s *Sheet) GetIndex() int {
	return s.index
}

func (s *Sheet) SetAutoCreateHeader(on bool) {
	s.autoCreateHeader = on
}

func (s *Sheet) addRow(n ...int) *Sheet {
	if len(n) == 0 {
		s.row += 1
	} else {
		for _, v := range n {
			s.row += v
		}
	}
	s.col = 0
	return s
}

func (s *Sheet) addCol() *Sheet {
	s.col += 1
	return s
}

func (s *Sheet) axis(row, col int) (string, error) {
	_col, err := excelize.ColumnNumberToName(col)
	if err != nil {
		return "", errors.Wrap(err, "excelize")
	}
	return excelize.JoinCellName(_col, row)
}

func (s *Sheet) fieldIsNil(data reflect.Value, index int) bool {
	dataValue := getElem(data)
	for k := 0; k < dataValue.Len(); k++ {
		v := getElem(dataValue.Index(k))
		if !v.Field(index).IsNil() {
			return false
		}
	}
	return true
}

// expandHeader
// fieldIndex 字段index
// index 表头开始位置
func (s *Sheet) expandHeader(dataValue reflect.Value, index int, col int) int {
	keySet := make(map[string]struct{}, 0)
	keyList := make([]string, 0)
	// 遍历所有数据，保证扩展字段表头是最完整的
	for k := 0; k < dataValue.Len(); k++ {
		v := getElem(dataValue.Index(k))
		field := v.Field(index)
		if field.Kind() == reflect.Map {
			for _, key := range field.MapKeys() {
				if _, ok := keySet[key.String()]; !ok {
					keySet[key.String()] = struct{}{}
					keyList = append(keyList, key.String())
				}
			}
		}
		if field.Kind() == reflect.Slice {
		}
	}
	sort.Strings(keyList)
	header := getElem(dataValue.Index(0))
	for _, v := range keyList {
		s.header = append(s.header, &excelHeaderField{
			fieldName:   header.Type().Field(index).Name,
			headerName:  v,
			Col:         col,
			allowEmpty:  false,
			expand:      false,
			expandRegex: nil,
			skip:        false,
			level:       2,
		})
		col += 1
	}
	return col
}

// transferHeaders
// 展开表头
func (s *Sheet) transferHeaders(data reflect.Value) *Sheet {
	// var data reflect.Type
	col := 1
	var value reflect.Value
	if data.Kind() == reflect.Slice {
		value = getElem(data.Index(0))
	} else if data.Kind() == reflect.Struct {
		value = data
	} else {
		panic("表头解析支持 struct | slice")
	}
	typee := value.Type()

	for i := 0; i < typee.NumField(); i++ {
		fieldType := value.Field(i)
		header := ParseExcelHeaderTag(typee.Field(i).Tag.Get("excel"), col)
		if header.IsSkip() {
			continue
		}
		header.fieldName = typee.Field(i).Name
		// 字段非nil，设置表头
		if data.Kind() == reflect.Slice && header.allowEmpty {
			header.allowEmpty = s.fieldIsNil(data, i)
		}
		// 展开扩展表头
		if header.expand {
			if fieldType.Kind() != reflect.Map {
				panic("expand表头非map[string]类型")
			}
			if value.Field(i).Len() > 0 {
				col = s.expandHeader(data, i, col)
			}
		} else {
			col += 1
		}
		s.header = append(s.header, header)
	}
	s.addRow()
	return s
}

func (s Sheet) GetCenterStyle() (int, error) {
	return s.Excel.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
			WrapText:   true,
		},
	})
}

// AddHeader 添加表头
func (s *Sheet) AddHeader(data interface{}) error {
	dataValue := reflect.ValueOf(data)
	if dataValue.Kind() != reflect.Slice ||
		dataValue.Len() == 0 {
		return nil
	}

	headerValue := getElem(dataValue.Index(0))

	switch headerValue.Kind() {
	case reflect.Struct:
		s.transferHeaders(dataValue)
		headerList := s.header
		sort.Sort(headerList)

		gatherHeader, ok := headerValue.Interface().(ExcelGatherHeader)
		if ok {
			if err := gatherHeader.GatherHeader(s); err != nil {
				return err
			}
			s.addRow(gatherHeader.GatherHeaderRows())
		}

		for _, v := range headerList {
			if v.IsSkip() || v.allowEmpty || v.expand {
				continue
			}
			s.addCol()
			axis, err := s.axis(s.row, s.col)
			if err != nil {
				return err
			}
			if err = s.setCellValue(axis, v, v.headerName); err != nil {
				return err
			}
		}
	default:
		return errors.New("行数据类型必须是struct")
	}
	return nil
}

func (s *Sheet) AddRemark(remark string, row, col int) error {
	s.addRow(row)
	axis, err := s.axis(row, col)
	if err != nil {
		return err
	}
	s.Excel.MergeCell(s.SheetName, "A1", axis)
	if err = s.Excel.SetCellValue(s.SheetName, "A1", remark); err != nil {
		return err
	}
	s.hasRemarks = true
	if style, err := s.Excel.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "left",
			Vertical:   "top",
			WrapText:   true,
		},
	}); err != nil {
		return nil
	} else {
		return s.Excel.SetCellStyle(s.SheetName, "A1", axis, style)
	}
}

func (s *Sheet) autoAddRemarks(data reflect.Value) error {
	if data.Len() == 0 {
		return nil
	}
	valueStruct := data.Index(0)
	if r, ok := valueStruct.Interface().(ExcelRemarks); ok {
		remarks, height, width := r.Remarks()
		return s.AddRemark(remarks, height, width)
	}
	return nil
}

func (s *Sheet) setCellValue(axis string, header *excelHeaderField, data interface{}) error {
	if header.font == nil {
		return s.Excel.SetCellValue(s.SheetName, axis, data)
	} else {
		return s.Excel.SetCellRichText(s.SheetName, axis, []excelize.RichTextRun{
			{
				Font: header.font,
				Text: fmt.Sprint(data),
			},
		})
	}
}

// AddData 遍历slice，导出数据
func (s *Sheet) AddData(data interface{}) error {
	dataType := reflect.TypeOf(data)
	dataValue := reflect.ValueOf(data)
	if dataType.Kind() != reflect.Slice {
		return errors.New("数据必须是slice")
	}

	if dataValue.Len() == 0 {
		_ = s.Excel.SetCellValue(s.SheetName, "A1", "没有数据")
		_ = s.Excel.MergeCell(s.SheetName, "A1", "C1")
		return nil
	}

	if !s.hasRemarks {
		if err := s.autoAddRemarks(dataValue); err != nil {
			return err
		}
	}

	if s.autoCreateHeader {
		if err := s.AddHeader(data); err != nil {
			return errors.Wrap(err, "创建表头失败")
		}
	}

	for k := 0; k < dataValue.Len(); k++ {
		valueStruct := getElem(dataValue.Index(k))
		if s.autoCreateHeader && valueStruct.Kind() == reflect.Slice && k == 0 {
			continue
		}
		s.addRow()
		switch valueStruct.Kind() {
		case reflect.Struct:
			for i := 0; i < valueStruct.NumField(); i++ {
				value := getElem(valueStruct.Field(i))
				headerNameMap := s.header.getFieldMap()
				header, ok := headerNameMap[valueStruct.Type().Field(i).Name]
				if !ok || header.IsSkip() {
					continue
				}
				if header.allowEmpty && s.fieldIsNil(dataValue, i) {
					continue
				}
				if !header.expand {
					axis, _ := s.axis(s.row, header.Col)
					if err := s.setCellValue(axis, header, value); err != nil {
						return err
					}
				} else {
					for _, key := range value.MapKeys() {
						if eHeader, ok := headerNameMap[key.String()]; ok {
							axis, _ := s.axis(s.row, eHeader.Col)
							if err := s.setCellValue(axis, header, value.MapIndex(key)); err != nil {
								return err
							}
						}
					}
				}
			}
		// case reflect.Slice:
		//	for i := 0; i < valueStruct.Len(); i++ {
		//		field := valueStruct.Index(i)
		//		_ = s.addCol().setCellValue(field)
		//	}
		default:
			return errors.New("行数据类型必须是struct或slice")
		}
	}
	return nil
}

// readHeader 读取表头, 确定表头位置
func (s *Sheet) readHeader(header []string) {
	headerMap := s.header.getHeaderMap()
	expandHeader := s.header.getExpandHeaderSlice()

	for col, cell := range header {
		cell = strings.TrimSpace(cell)
		if h, ok := headerMap[cell]; ok {
			h.Col = col + 1
			h.isMatch = true
		} else {
			for _, v := range expandHeader {
				if v.expandRegex.MatchString(cell) {
					v.Col = -1
					s.header = append(s.header, &excelHeaderField{
						Col:         col + 1,
						fieldName:   v.fieldName,
						headerName:  cell,
						allowEmpty:  false,
						expand:      false,
						expandRegex: nil,
						skip:        false,
						level:       2,
						isMatch:     true,
					})
				}
			}
		}
	}
}

func (s *Sheet) ExpandHeaderLen() int {
	count := 0
	for _, v := range s.header {
		if v.level == 2 {
			count += 1
		}
	}
	return count
}

func (s *Sheet) cellToValue(field reflect.Type, cell string, axis string) (reflect.Value, error) {
	cell = strings.TrimSpace(cell)
	switch field.Kind() {
	case reflect.String:
		return reflect.ValueOf(cell), nil
	case reflect.Ptr:
		v, err := s.cellToValue(field.Elem(), cell, axis)
		if err != nil {
			return reflect.Value{}, err
		}
		x := reflect.New(field.Elem())
		x.Elem().Set(v)
		return x, err
	case reflect.Int8:
		if cell == "" {
			return reflect.ValueOf(int8(0)), nil
		}
		i, err := strconv.ParseInt(cell, 10, 64)
		if err != nil {
			return reflect.Value{}, errors.Wrapf(err, "%s表格(%s)转int8失败", axis, cell)
		}
		return reflect.ValueOf(int8(i)), nil
	case reflect.Int16:
		if cell == "" {
			return reflect.ValueOf(int16(0)), nil
		}
		i, err := strconv.ParseInt(cell, 10, 64)
		if err != nil {
			return reflect.Value{}, errors.Wrapf(err, "%s表格(%s)转int16失败", axis, cell)
		}
		return reflect.ValueOf(int16(i)), nil
	case reflect.Int32:
		if cell == "" {
			return reflect.ValueOf(int32(0)), nil
		}
		i, err := strconv.ParseInt(cell, 10, 64)
		if err != nil {
			return reflect.Value{}, errors.Wrapf(err, "%s表格(%s)转int32失败", axis, cell)
		}
		return reflect.ValueOf(int32(i)), nil
	case reflect.Uint32:
		if cell == "" {
			return reflect.ValueOf(int32(0)), nil
		}
		i, err := strconv.ParseInt(cell, 10, 64)
		if err != nil {
			return reflect.Value{}, errors.Wrapf(err, "%s表格(%s)转int32失败", axis, cell)
		}
		return reflect.ValueOf(uint32(i)), nil
	case reflect.Int:
		if cell == "" {
			return reflect.ValueOf(int(0)), nil
		}
		i, err := strconv.ParseInt(cell, 10, 64)
		if err != nil {
			return reflect.Value{}, errors.Wrapf(err, "%s表格(%s)转int失败", axis, cell)
		}
		return reflect.ValueOf(int(i)), nil
	case reflect.Int64:
		if cell == "" {
			return reflect.ValueOf(int64(0)), nil
		}
		i, err := strconv.ParseInt(cell, 10, 64)
		if err != nil {
			return reflect.Value{}, errors.Wrapf(err, "%s表格(%s)转int64失败", axis, cell)
		}
		return reflect.ValueOf(i), nil
	case reflect.Uint64:
		if cell == "" {
			return reflect.ValueOf(int64(0)), nil
		}
		i, err := strconv.ParseInt(cell, 10, 64)
		if err != nil {
			return reflect.Value{}, errors.Wrapf(err, "%s表格(%s)转int64失败", axis, cell)
		}
		return reflect.ValueOf(uint64(i)), nil
	case reflect.Bool:
		lower := strings.ToLower(cell)
		if lower == "true" || lower == "1" || lower == "t" {
			return reflect.ValueOf(true), nil
		} else if lower == "" || lower == "false" || lower == "f" || lower == "0" {
			return reflect.ValueOf(false), nil
		} else {
			return reflect.Value{}, errors.Errorf("%s表格(%s)转bool失败", axis, cell)
		}
	case reflect.Float32:
		if cell == "" {
			return reflect.ValueOf(float32(0)), nil
		}
		f, err := strconv.ParseFloat(cell, 64)
		if err != nil {
			return reflect.Value{}, errors.Wrapf(err, "%s表格(%s)转float32失败", axis, cell)
		}
		return reflect.ValueOf(f), nil
	case reflect.Float64:
		if cell == "" {
			return reflect.ValueOf(float64(0)), nil
		}
		f, err := strconv.ParseFloat(cell, 64)
		if err != nil {
			return reflect.Value{}, errors.Wrapf(err, "%s表格(%s)转float64失败", axis, cell)
		}
		return reflect.ValueOf(f), nil
	}
	return reflect.Value{}, errors.Errorf("暂不支持的类型: %s，需要添加一下switch case", field.Kind())
}

func (s *Sheet) readBody(rows [][]string, data reflect.Value) (interface{}, error) {
	hMap := s.header.getColHeaderMap()
	res := reflect.MakeSlice(reflect.SliceOf(reflect.New(data.Type()).Type()), 0, len(rows))
	for rn, row := range rows {
		itemPtr := reflect.New(data.Type())
		item := itemPtr.Elem()
		for col, cell := range row {
			if h, ok := hMap[col+1]; ok {
				field := item.FieldByName(h.fieldName)
				if !field.CanSet() {
					continue
				}
				axis, _ := s.axis(rn+2, col+1)

				switch field.Kind() {
				case reflect.Map:
					if value, err := s.cellToValue(field.Type().Elem(), cell, axis); err != nil {
						return nil, err
					} else {
						if field.IsNil() {
							field.Set(reflect.MakeMap(field.Type()))
						}
						field.SetMapIndex(reflect.ValueOf(h.headerName), value)
					}
				default:
					if value, err := s.cellToValue(field.Type(), cell, axis); err != nil {
						return nil, err
					} else {
						field.Set(value)
					}
				}
			}
		}
		res = reflect.Append(res, itemPtr)
	}
	return res.Interface(), nil
}

// ReadData 读取表格数据,
// @return []*struct{}
func (s *Sheet) ReadData(data interface{}) (interface{}, error) {
	dataValue := getElem(reflect.ValueOf(data))
	dataType := dataValue.Type()
	if dataType.Kind() != reflect.Struct {
		return nil, errors.New("data必须是struct类型")
	}

	s.transferHeaders(dataValue)

	rows, err := s.Excel.GetRows(s.SheetName)
	if err != nil {
		return nil, err
	}

	if len(rows) == 0 {
		return nil, errors.New("excel没有数据")
	}

	rows = s.filterEmpty(rows)
	// 头部备注
	start := 0
	if remarker, ok := data.(ExcelRemarks); ok {
		remarks, _, _ := remarker.Remarks()
		if len(rows[0]) == 1 {
			if strings.TrimSpace(rows[0][0]) == strings.TrimSpace(remarks) {
				start += 1
			}
		}
	}
	if gatherHeader, ok := data.(ExcelGatherHeader); ok {
		start += gatherHeader.GatherHeaderRows()
	}
	s.readHeader(rows[start])
	return s.readBody(rows[start+1:], dataValue)
}

func (s *Sheet) filterEmpty(rows [][]string) [][]string {
	res := make([][]string, 0)

	for _, row := range rows {
		isEmpty := true
		for _, col := range row {
			if len(strings.TrimSpace(col)) > 0 {
				isEmpty = false
				continue
			}
		}
		if !isEmpty {
			res = append(res, row)
		}
	}

	return res
}
