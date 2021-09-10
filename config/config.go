package config

import (
	"github.com/spf13/viper"
)

type Config struct {
	Concurrent       bool
	TaskMap          map[string]bool
	SrcPath, DstPath string
}

func InitConf() *Config {
	// setup default config
	// about concurrency, default is enable
	concur := true

	// about task, default is all enable
	taskMap := make(map[string]bool, 4)
	taskMap["content"] = true
	taskMap["campaign"] = true
	taskMap["cps"] = true
	taskMap["newgame"] = true

	// about src and dst path
	src := "src"
	dst := "dst"

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
	if src == "" {
		src = "src"
	}
	if dst == "" {
		dst = "dst"
	}

	return &Config{
		Concurrent: concur,
		TaskMap:    taskMap,
		SrcPath:    src,
		DstPath:    dst,
	}
}
