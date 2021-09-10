package collect

import (
	"errors"
	"excel/config"
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"path/filepath"
	"strings"
	"sync"
)

type Collect struct {
	conf               *config.Config
	srcDir, dstDir     string
	srcFiles, dstFiles map[string]*excelize.File // file_name, fd
	srcCsvFiles        map[string]*os.File
	dstFilesMutex      map[string]*sync.Mutex
}

type Sheet struct {
	name      string // sheet name
	start     string // start coordinates value for search
	row, col  int    // row and col index now
	file      *excelize.File
	fileMutex *sync.Mutex
	data      [][]string // each col data of each row
	month     string
	indexs    []int             // special col index
	org       map[string]string // uid, org
}

func NewCollect(config *config.Config) *Collect {
	return &Collect{
		conf:          config,
		srcDir:        config.SrcPath,
		dstDir:        config.DstPath,
		srcFiles:      make(map[string]*excelize.File),
		dstFiles:      make(map[string]*excelize.File),
		srcCsvFiles:   make(map[string]*os.File),
		dstFilesMutex: make(map[string]*sync.Mutex),
	}
}

func (c *Collect) loadSrcFiles() error {
	files, err := os.ReadDir(c.srcDir)
	if err != nil {
		return err
	}
	for _, file := range files {
		if strings.HasSuffix(file.Name(), "xlsx") || strings.HasSuffix(file.Name(), "xls") {
			f, err := excelize.OpenFile(c.srcDir + "/" + file.Name())
			if err != nil {
				return err
			}
			c.srcFiles[file.Name()] = f
			// fmt.Println("successfully load", file.Name())
		} else if strings.HasSuffix(file.Name(), "csv") {
			f, err := os.OpenFile(c.srcDir+"/"+file.Name(), os.O_RDONLY, os.ModePerm)
			if err != nil {
				return err
			}
			c.srcCsvFiles[file.Name()] = f
		}
	}

	return nil
}

func (c *Collect) createDstFile(filename string) error {
	f := excelize.NewFile()
	s, err := os.Stat(filepath.Dir(c.dstDir + "/" + filename))
	if err != nil && os.IsNotExist(err) {
		err := os.MkdirAll(filepath.Dir(c.dstDir+"/"+filename), 0755)
		if err != nil {
			return err
		}
	} else {
		if !s.IsDir() {
			return fmt.Errorf("path %s is not dir", filepath.Dir(c.dstDir+"/"+filename))
		}
	}
	// TODO: backup old file before save as new file
	if err := f.SaveAs(c.dstDir + "/" + filename); err != nil {
		return err
	}
	c.dstFiles[filename] = f
	c.dstFilesMutex[filename] = new(sync.Mutex)
	return nil
}

func (c *Collect) Run() error {
	// load all src files, record fd
	if err := c.loadSrcFiles(); err != nil {
		return err
	}
	// create dst file for output
	if err := c.createDstFile("项目立项及实际费用明细.xlsx"); err != nil {
		return err
	}

	// do task concurrently or sequentially
	var runErr []error
	errChan := make(chan error, len(c.conf.TaskMap))
	wg := &sync.WaitGroup{}
	for task, enabled := range c.conf.TaskMap {
		if !enabled {
			continue
		}
		if c.conf.Concurrent {
			wg.Add(1)
			go c.doTask(task, wg, errChan)
		} else {
			c.doTask(task, wg, errChan)
		}

		wg.Add(1)
		go func(wg *sync.WaitGroup) {
			defer wg.Done()
			if err := <-errChan; err != nil {
				runErr = append(runErr, err)
			}
		}(wg)
	}

	wg.Wait()
	if runErr != nil {
		return runErr[0] // only return the first error
	} else {
		return nil
	}
}

func (c *Collect) doTask(task string, wg *sync.WaitGroup, errChan chan error) {
	var err error
	switch task {
	case "content":
		err = c.CollectForContent()
	case "campaign":
		err = c.CollectForAll("活动")
	case "cps":
		err = c.CollectForAll("CPS分发")
	case "newgame":
		err = c.CollectForAll("新游预约")
	case "mcn":
		err = c.CollectForMcn()
	default:
		err = errors.New("unsupported task")
	}
	if err != nil {
		fmt.Println("task[", task, "]: failed")
		errChan <- err
	} else {
		fmt.Println("task[", task, "]: successful")
		errChan <- nil
	}
	if c.conf.Concurrent {
		wg.Done()
	}
}
