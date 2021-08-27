package main

import (
	"excel/collect"
	"fmt"
	"os"
)

func main() {
	collectInstance := collect.NewCollect("src", "dst")
	collectInstance.Run()
	fmt.Println("若本句上面有其它信息输出，应该不要相信本次运行的结果，并下面链接反馈问题")
	fmt.Println("https://docs.google.com/spreadsheets/d/1GkcPa0WjVt2UBVnRNQ-1SO49vYsR0CQgk3qtA-VxE-Y/edit#gid=0")
	fmt.Println("请按回车退出")
	b := make([]byte, 1)
	os.Stdin.Read(b)
}
