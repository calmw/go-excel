package main

import (
	"github.com/calmw/go-excel/excel"
)

func main() {
	// 将列表转换成excel文件
	list := make([]User, 0)
	user1 := User{
		Name:    "张三",
		Age:     18,
		Address: "北京东三环",
	}
	user2 := User{
		Name:    "李四",
		Age:     21,
		Address: "上海人民路",
	}
	user3 := User{
		Name:    "王五",
		Age:     22,
		Address: "长沙开福区",
	}
	list = append(list, user1, user2, user3)
	f := excel.ListToExcel(list, "员工信息表", "员工表")
	_ = f.SaveAs("./test.xlsx")
}

type User struct {
	Name    string `excel:"column:B;desc:姓名;width:30"`
	Age     int    `excel:"column:C;desc:年龄;width:10"`
	Address string `excel:"column:D;desc:地址;width:50"`
}
