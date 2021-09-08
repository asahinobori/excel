package collect

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"path/filepath"
	"strings"
	"sync"
)

var dirDefault = "."

type Collect struct {
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

func NewCollect(args ...string) *Collect {
	srcDirSet := dirDefault
	dstDirSet := dirDefault
	if len(args) >= 1 && len(args[0]) != 0 {
		srcDirSet = args[0]
	}
	if len(args) >= 2 && len(args[1]) != 0 {
		dstDirSet = args[1]
	}

	return &Collect{
		srcDir:        srcDirSet,
		dstDir:        dstDirSet,
		srcFiles:      make(map[string]*excelize.File),
		dstFiles:      make(map[string]*excelize.File),
		srcCsvFiles:   make(map[string]*os.File),
		dstFilesMutex: make(map[string]*sync.Mutex),
	}
}

func (c *Collect) LoadSrcFiles() error {
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

func (c *Collect) CreateDstFile(filename string) error {
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
	if err := c.LoadSrcFiles(); err != nil {
		return err
	}
	// create dst file for output
	if err := c.CreateDstFile("项目立项及实际费用明细.xlsx"); err != nil {
		return err
	}

	var runErr error = nil
	errChan := make(chan error)
	wg := sync.WaitGroup{}
	wg.Add(1)
	go func() {
		// collect for content
		defer wg.Done()
		if err := c.CollectForContent(); err != nil {
			errChan <- err
		} else {
			errChan <- nil
		}
	}()
	// collect for common
	wg.Add(1)
	go func() {
		defer wg.Done()
		if err := c.CollectForAll("活动"); err != nil {
			errChan <- err
		} else {
			errChan <- nil
		}
	}()
	wg.Add(1)
	go func() {
		defer wg.Done()
		if err := c.CollectForAll("CPS分发"); err != nil {
			errChan <- err
		} else {
			errChan <- nil
		}
	}()
	wg.Add(1)
	go func() {
		defer wg.Done()
		if err := c.CollectForAll("新游预约"); err != nil {
			errChan <- err
		} else {
			errChan <- nil
		}
	}()
	go func() {
		for err := range errChan {
			if err != nil {
				runErr = err
			}
		}
	}()
	wg.Wait()
	return runErr
}
