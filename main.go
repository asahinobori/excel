package main

import (
	"excel/collect"
	"excel/config"
	"excel/log"
	"fmt"
	"os"
)

func main() {
	conf := config.InitConf()
	collectInstance := collect.NewCollect(conf)
	err := collectInstance.Run()

	if err != nil {
		fmt.Println(err)
		fmt.Println("运行过程中出现问题！！！")
		log.Error(err)
	} else {
		fmt.Println("运行过程中没有问题")
		log.Info("no problem cause before exit")
	}
	fmt.Println("请按回车退出")
	b := make([]byte, 1)
	os.Stdin.Read(b)
}
