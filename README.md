# go-struct-excel

需求：

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

# 用法

