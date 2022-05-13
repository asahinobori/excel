package config

import (
	"excel/log"
	"github.com/spf13/viper"
)

type Config struct {
	Concurrent       bool
	TaskMap          map[string]bool
	SrcPath, DstPath string
	LogLevel         string
}

func InitConf() *Config {
	// setup default config
	// about concurrency, default is enable
	concur := true

	// about task, default is all enable
	taskMap := make(map[string]bool, 5)
	taskMap["content"] = true
	taskMap["campaign"] = true
	taskMap["cps"] = true
	taskMap["newgame"] = true
	taskMap["mcn"] = true

	// about src and dst path
	src := "src"
	dst := "dst"

	// about log
	logl := "error"

	// parse config file
	viper.SetConfigName("config.ini")
	viper.SetConfigType("toml")
	viper.AddConfigPath(".")
	if err := viper.ReadInConfig(); err != nil {
		// return default config
		return &Config{
			Concurrent: concur,
			TaskMap:    taskMap,
			SrcPath:    src,
			DstPath:    dst,
			LogLevel:   logl,
		}
	}

	if viper.IsSet("concurrency.enable") && viper.GetInt("concurrency.enable") <= 0 {
		concur = false
	}

	tasks := viper.GetStringMap("task")
	for task, v := range tasks {
		isOn, ok := v.(int64)
		if ok && isOn > 0 {
			taskMap[task] = true
		} else {
			taskMap[task] = false
		}
	}

	src = viper.GetString("directory.src")
	dst = viper.GetString("directory.dst")
	logl = viper.GetString("log.level")
	if src == "" {
		src = "src"
	}
	if dst == "" {
		dst = "dst"
	}
	if logl == "" {
		logl = "error"
	}

	// set log level
	if err := log.SetLevel(logl); err != nil {
		logl = "error"
	}

	return &Config{
		Concurrent: concur,
		TaskMap:    taskMap,
		SrcPath:    src,
		DstPath:    dst,
		LogLevel:   logl,
	}
}
