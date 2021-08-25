package main

import "excel/collect"

func main() {
    collectInstance := collect.NewCollect("src", "dst")
	collectInstance.Run()
}
