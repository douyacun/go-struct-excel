package structexcel

import "reflect"

func getElem(elem reflect.Value) reflect.Value {
	for elem.Kind() == reflect.Ptr || elem.Kind() == reflect.Interface {
		elem = elem.Elem()
	}
	return elem
}
