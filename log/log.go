package log

import (
	"github.com/sirupsen/logrus"
	"os"
)

func init() {
	logrus.SetFormatter(&logrus.TextFormatter{})
	logFile, _ := os.OpenFile("./excel.log", os.O_CREATE|os.O_APPEND|os.O_WRONLY, 0644)
	logrus.SetOutput(logFile)
}

func SetLevel(lvl string) error {
	if l, err := logrus.ParseLevel(lvl); err != nil {
		logrus.SetLevel(logrus.ErrorLevel)
		logrus.Error("not a valid log level:", lvl, ", set level to error")
		return err
	} else {
		logrus.SetLevel(l)
	}
	return nil
}

func Info(args ...interface{}) {
	logrus.Info(args...)
}

func Debug(args ...interface{}) {
	logrus.Debug(args...)
}

func Error(args ...interface{}) {
	logrus.Error(args...)
}
