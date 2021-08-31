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

const (
	// start from zero for src
	department = 0 // 运营部门
	game       = 1 // 游戏产品
	sponsor    = 2 // 出资方
	uid        = 4 // UID
	nickName   = 5 // 昵称
	money      = 9 // 金额（橙列）

	// start from zero for dst
	monthD      = 0  // 月份
	orgD        = 1  // 机构
	departmentD = 2  // 部门
	gameD       = 3  // 游戏
	uidD        = 4  // UID
	nickNameD   = 5  // 昵称
	videoMoneyD = 6  // 视频费用
	textMoneyD  = 7  // 图文费用
	unclsMoneyD = 8  // 不能区分（费用）
	typeD       = 10 // 类别
	sponsorD    = 12 // 出资方
	readCntD    = 13 // 阅读量
)

var collectMap = map[int]int{
	department: departmentD,
	game:       gameD,
	sponsor:    sponsorD,
	uid:        uidD,
	nickName:   nickNameD,
}

type Collect struct {
	srcDir, dstDir     string
	srcFiles, dstFiles map[string]*excelize.File // file_name, fd
	srcCsvFiles        map[string]*os.File
}

type Sheet struct {
	name         string // sheet name
	start        string // start coordinates value for search
	row, col     int    // row and col index now
	file         *excelize.File
	data         [][]string // each col data of each row
	month        string
	dynTypeIndex int
	readCntIndex int
	org          map[string]string // uid, org
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
							s.dynTypeIndex = id
						} else if strings.Contains(colData, "阅读量") {
							s.readCntIndex = id
							break
						}
					}
					continue
				} else if colsData == nil {
					break // absolutely data end
				} else if (len(colsData) > money) &&
					(len(colsData[uid]) == 0) && (len(colsData[nickName]) == 0) && (len(colsData[money]) == 0) {
					break // maybe data end
				} else if (len(colsData) > money) &&
					((len(colsData[uid]) == 0) || (len(colsData[nickName]) == 0) || (len(colsData[money-1]) == 0) || (len(colsData[money]) == 0)) {
					continue // data not enough
				} else if len(colsData) <= money {
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
			if dstCol, ok := collectMap[col]; ok {
				dstAxis, _ = excelize.CoordinatesToCellName(dstCol+1, s.row)
				err = s.file.SetCellValue(s.name, dstAxis, colData)
				if err != nil {
					return err
				}
			}
		}

		// deal with month
		dstAxis, _ = excelize.CoordinatesToCellName(monthD+1, s.row)
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
		if org, exist := s.org[colsData[uid]]; exist {
			dstAxis, _ := excelize.CoordinatesToCellName(orgD+1, s.row)
			err = s.file.SetCellValue(s.name, dstAxis, org)
		} else {
			dstAxis, _ := excelize.CoordinatesToCellName(orgD+1, s.row)
			err = s.file.SetCellValue(s.name, dstAxis, "其他_付费kol")
		}

		// deal with sum
		if (from.dynTypeIndex == 0) || (colsData[from.dynTypeIndex] == "") {
			dstAxis, _ := excelize.CoordinatesToCellName(unclsMoneyD+1, s.row)
			err = s.file.SetCellValue(s.name, dstAxis, colsData[money])
		} else if strings.Contains(colsData[from.dynTypeIndex], "视频") {
			dstAxis, _ := excelize.CoordinatesToCellName(videoMoneyD+1, s.row)
			err = s.file.SetCellValue(s.name, dstAxis, colsData[money])
		} else {
			dstAxis, _ := excelize.CoordinatesToCellName(textMoneyD+1, s.row)
			err = s.file.SetCellValue(s.name, dstAxis, colsData[money])
		}
		if err != nil {
			return err
		}

		// deal with type
		dstAxis, _ = excelize.CoordinatesToCellName(typeD+1, s.row)
		err = s.file.SetCellValue(s.name, dstAxis, from.name)
		if err != nil {
			return err
		}

		// deal with readCnt
		if from.readCntIndex != 0 {
			dstAxis, _ = excelize.CoordinatesToCellName(readCntD+1, s.row)
			err = s.file.SetCellValue(s.name, dstAxis, colsData[from.readCntIndex])
			if err != nil {
				return err
			}
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
			name:         "内容创作者",
			start:        "运营部门",
			file:         f,
			month:        monthRes,
			dynTypeIndex: 0,
			readCntIndex: 0,
		})
		sheets = append(sheets, Sheet{
			name:         "内容采购",
			start:        "运营部门",
			file:         f,
			month:        monthRes,
			dynTypeIndex: 0,
			readCntIndex: 0,
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
