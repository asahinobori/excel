package collect

import (
	"fmt"
	"github.com/dimchansky/utfbom"
	"github.com/gocarina/gocsv"
	"github.com/xuri/excelize/v2"
	"io/ioutil"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
)

var dirDefault = "."

var collectMap = map[int]int{
	1: 3,
	2: 4,
	3: 13,
	5: 5,
	6: 6,
}

const (
	// start from zero
	uidIndex      = 4
	nickNameIndex = 5
	moneyIndex    = 9
)

type Collect struct {
	srcDir, dstDir     string
	srcFiles, dstFiles map[string]*excelize.File // file_name, fd
	srcCsvFiles        map[string]*os.File
}

type Sheet struct {
	name      string // sheet name
	start     string // start coordinates value for search
	row, col  int    // row and col index now
	file      *excelize.File
	data      [][]string // each col data of each row
	month     string
	typeIndex int
	org       map[string]string // uid, org
}

type Org struct {
	KolType string `csv:"kol_type"`
	Uid     string `csv:"uid"`
	AddDate string `csv:"add_date"`
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
		srcDir:      srcDirSet,
		dstDir:      dstDirSet,
		srcFiles:    make(map[string]*excelize.File),
		dstFiles:    make(map[string]*excelize.File),
		srcCsvFiles: make(map[string]*os.File),
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
			f, err := os.OpenFile(c.srcDir+"/"+file.Name(), os.O_RDWR, os.ModePerm)
			if err != nil {
				panic(err)
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
	return nil
}

func (c *Collect) ReadCSV(orgsMap map[string]string) error {
	orgs := make([]*Org, 0)
	orgMap := make(map[string]int)
	for _, f := range c.srcCsvFiles {
		csvContent, err := ioutil.ReadAll(utfbom.SkipOnly(f))
		if err != nil {
			return err
		}
		if err := gocsv.UnmarshalBytes(csvContent, &orgs); err != nil {
			return err
		}
		for id, org := range orgs {
			if orgsId, exist := orgMap[org.Uid]; exist {
				if strings.Contains(orgs[orgsId].KolType, "其他") && !strings.Contains(org.KolType, "其他") {
					orgMap[org.Uid] = id
					orgsMap[org.Uid] = org.KolType
				} else if newDate, err := strconv.Atoi(org.AddDate); err == nil {
					if oldDate, err := strconv.Atoi(orgs[orgsId].AddDate); err == nil {
						if newDate > oldDate {
							orgMap[org.Uid] = id
							orgsMap[org.Uid] = org.KolType
						}
					}
				}
			} else {
				orgMap[org.Uid] = id
				orgsMap[org.Uid] = org.KolType
			}
		}
	}
	return nil
}

func (s *Sheet) ReadSheet() error {
	sheetList := s.file.GetSheetList()
	for _, sheetName := range sheetList {
		// skip hidden sheet
		if !s.file.GetSheetVisible(sheetName) {
			continue
		}
		if strings.Contains(sheetName, s.name) {
			startFound, err := s.file.SearchSheet(sheetName, s.start)
			if err != nil {
				return err
			} else if startFound == nil {
				continue
			}
			s.col, s.row, err = excelize.CellNameToCoordinates(startFound[0])

			// traverse this sheet and get data from start coordinate
			curRow := 0
			rowsIt, err := s.file.Rows(sheetName)
			if err != nil {
				return err
			}
			for rowsIt.Next() {
				curRow++
				colsData, err := rowsIt.Columns()
				if err != nil {
					return err
				} else if curRow < s.row {
					continue
				} else if curRow == s.row {
					for id, colData := range colsData {
						if strings.Contains(colData, "动态类型") {
							s.typeIndex = id
							break
						}
					}
					continue
				} else if colsData == nil {
					break // absolutely data end
				} else if (len(colsData) > moneyIndex) &&
					(len(colsData[uidIndex]) == 0) && (len(colsData[nickNameIndex]) == 0) && (len(colsData[moneyIndex]) == 0) {
					break // maybe data end
				} else if (len(colsData) > moneyIndex) &&
					((len(colsData[uidIndex]) == 0) || (len(colsData[nickNameIndex]) == 0) || (len(colsData[moneyIndex-1]) == 0) || (len(colsData[moneyIndex]) == 0)) {
					continue // data not enough
				} else if len(colsData) <= moneyIndex {
					break // maybe data end
				}
				s.data = append(s.data, colsData)
			}
		}
	}
	return nil
}

func (s *Sheet) WriteSheet(from *Sheet) error {
	sheetList := s.file.GetSheetList()
	foundSheet := false
	for _, sheetName := range sheetList {
		if strings.Contains(sheetName, s.name) {
			s.name = sheetName
			foundSheet = true
		}
	}
	if !foundSheet {
		// create new sheet
		index := s.file.NewSheet(s.name)
		s.file.SetActiveSheet(index)
		if err := s.file.SetSheetVisible("Sheet1", false); err != nil {
			return err
		}
		if err := s.file.Save(); err != nil {
			return err
		}
	}

	var dstAxis string
	var err error
	for _, colsData := range from.data {
		s.row++
		for col, colData := range colsData {
			if dstCol, ok := collectMap[col+1]; ok {
				dstAxis, _ = excelize.CoordinatesToCellName(dstCol, s.row)
				err = s.file.SetCellValue(s.name, dstAxis, colData)
				if err != nil {
					return err
				}
			}
		}

		// deal with month
		dstAxis, _ = excelize.CoordinatesToCellName(1, s.row)
		err = s.file.SetCellValue(s.name, dstAxis, "2021/"+from.month+"/1")
		if err != nil {
			return err
		}
		exp := "yyyy\"年\"m\"月\""
		style, err := s.file.NewStyle(&excelize.Style{CustomNumFmt: &exp})
		err = s.file.SetCellStyle(s.name, dstAxis, dstAxis, style)
		if err != nil {
			return err
		}

		// deal with org
		if org, exist := s.org[colsData[uidIndex]]; exist {
			dstAxis, _ := excelize.CoordinatesToCellName(2, s.row)
			err = s.file.SetCellValue(s.name, dstAxis, org)
		} else {
			dstAxis, _ := excelize.CoordinatesToCellName(2, s.row)
			err = s.file.SetCellValue(s.name, dstAxis, "其他_付费kol")
		}

		// deal with sum
		if (from.typeIndex == 0) || (colsData[from.typeIndex] == "") {
			dstAxis, _ := excelize.CoordinatesToCellName(9, s.row)
			err = s.file.SetCellValue(s.name, dstAxis, colsData[moneyIndex])
		} else if strings.Contains(colsData[from.typeIndex], "视频") {
			dstAxis, _ := excelize.CoordinatesToCellName(7, s.row)
			err = s.file.SetCellValue(s.name, dstAxis, colsData[moneyIndex])
		} else if strings.Contains(colsData[from.typeIndex], "文") {
			dstAxis, _ := excelize.CoordinatesToCellName(8, s.row)
			err = s.file.SetCellValue(s.name, dstAxis, colsData[moneyIndex])
		} else {
			dstAxis, _ := excelize.CoordinatesToCellName(9, s.row)
			err = s.file.SetCellValue(s.name, dstAxis, colsData[moneyIndex])
		}
		if err != nil {
			return err
		}

		// deal with type
		dstAxis, _ = excelize.CoordinatesToCellName(11, s.row)
		err = s.file.SetCellValue(s.name, dstAxis, from.name)
		if err != nil {
			return err
		}
	}

	if err = s.file.Save(); err != nil {
		return err
	}

	return nil
}

func (c *Collect) Run() error {
	// load all src files, record fd
	if err := c.LoadSrcFiles(); err != nil {
		return err
	}
	// create dst file for output
	if err := c.CreateDstFile("collect.xlsx"); err != nil {
		return err
	}

	// Step 1: read csv
	orgsMap := make(map[string]string)
	if err := c.ReadCSV(orgsMap); err != nil {
		return err
	}

	// Step 2: read sheet in src excel files, and write to dst file
	sheets := make([]Sheet, 0)
	for fname, f := range c.srcFiles {
		monthReg1 := regexp.MustCompile(`[^\d]\d+月`)
		monthReg2 := regexp.MustCompile(`\d+`)
		monthRes := monthReg2.FindStringSubmatch(monthReg1.FindStringSubmatch(fname)[0])[0]
		sheets = append(sheets, Sheet{
			name:      "内容创作者",
			start:     "运营部门",
			file:      f,
			month:     monthRes,
			typeIndex: 0,
		})
		sheets = append(sheets, Sheet{
			name:      "内容采购",
			start:     "运营部门",
			file:      f,
			month:     monthRes,
			typeIndex: 0,
		})
	}

	targetSheet := &Sheet{
		name: "大神内域作者费用明细",
		row:  1,
		col:  1,
		file: c.dstFiles["collect.xlsx"],
		org:  orgsMap,
	}

	for _, sheet := range sheets {
		if err := sheet.ReadSheet(); err != nil {
			return err
		}
		if err := targetSheet.WriteSheet(&sheet); err != nil {
			return err
		}
	}
	return nil
}
