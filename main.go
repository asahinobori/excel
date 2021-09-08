package main

import (
	"excel/collect"
	"fmt"
	"os"
)

func main() {
	collectInstance := collect.NewCollect("src", "dst")
	err := collectInstance.Run()
	if err != nil {
		fmt.Println(err)
		fmt.Println("处理失败！！！请到下面链接反馈问题")
		fmt.Println("https://docs.google.com/spreadsheets/d/1GkcPa0WjVt2UBVnRNQ-1SO49vYsR0CQgk3qtA-VxE-Y/edit#gid=0")
	} else {
		fmt.Println("处理成功")
	}
	fmt.Println("请按回车退出")
	b := make([]byte, 1)
	os.Stdin.Read(b)
}
