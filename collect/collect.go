package collect

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"path/filepath"
	"regexp"
	"strings"
)

var dirDefault = "."

var collectMap  = map[int]int {
	1 : 3,
	2 : 4,
	3 : 13,
	5 : 5,
	6 : 6,
}

const moneyIndex = 9

type Collect struct {
	srcDir, dstDir string
	SrcFiles, DstFiles map[string]*excelize.File  // file_name, fd
}

type Sheet struct {
	name string  // sheet name
	start string  // start coordinates value for search
	row, col int  // row and col index now
	file *excelize.File
	data map[int][]string  // row, col_data
	month string
	typeIndex int
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

	return &Collect {
		srcDir: srcDirSet,
		dstDir: dstDirSet,
		SrcFiles: make(map[string]*excelize.File),
		DstFiles: make(map[string]*excelize.File),
	}
}

func (c *Collect) LoadSrcExcels() error {
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
			c.SrcFiles[file.Name()] = f
			// fmt.Println("successfully load", file.Name())
		}
	}

	return nil
}

func (c *Collect) CreateDstExcel(filename string) error {
	f := excelize.NewFile()
	s, err := os.Stat(filepath.Dir(c.dstDir + "/" + filename))
	if err != nil && os.IsNotExist(err) {
		err := os.MkdirAll(filepath.Dir(c.dstDir + "/" + filename), 0755)
		if err != nil {
			return err
		}
	} else {
		if !s.IsDir() {
			return fmt.Errorf("path %s is not dir", filepath.Dir(c.dstDir + "/" + filename))
		}
	}
	// TODO: backup old file before save as new file
	if err := f.SaveAs(c.dstDir + "/" + filename); err != nil {
		return err
	}
	c.DstFiles[filename] = f
	return nil
}

func (s *Sheet) ReadSheet() error {
	sheetList := s.file.GetSheetList()
	for _, sheetName := range sheetList {
		if strings.Contains(sheetName, s.name) {
			startFound, err := s.file.SearchSheet(sheetName, s.start)
			if err != nil {
				return err
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
				} else if (colsData == nil) || (len(colsData[0]) == 0) || (len(colsData) > moneyIndex && len(colsData[moneyIndex]) == 0) {
					break
				}
				s.data[curRow] = colsData
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
		_ = s.file.NewSheet(s.name)
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
		err = s.file.SetCellValue(s.name, dstAxis, "2021/" + from.month + "/1")
		if err != nil {
			return err
		}

		// deal with sum
		if (from.typeIndex == 0) || (colsData[from.typeIndex] == "") {
			dstAxis, _ := excelize.CoordinatesToCellName(9, s.row)
			err = s.file.SetCellValue(s.name, dstAxis, colsData[moneyIndex])
		} else if strings.Contains(colsData[from.typeIndex], "视频") {
			dstAxis, _ := excelize.CoordinatesToCellName(7, s.row)
			err = s.file.SetCellValue(s.name, dstAxis, colsData[moneyIndex])
		} else if strings.Contains(colsData[from.typeIndex], "文", ) {
			dstAxis, _ := excelize.CoordinatesToCellName(8, s.row)
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

func (c *Collect) Run() {
	if err := c.LoadSrcExcels(); err != nil {
		fmt.Println(err)
	}
	if err := c.CreateDstExcel("collect.xlsx"); err != nil {
		fmt.Println(err)
	}

	sheets := make([]Sheet, 0)
	for fname, f := range c.SrcFiles {
		monthReg1 := regexp.MustCompile(`[^\d]\d+月`)
		monthReg2 := regexp.MustCompile(`\d+`)
		monthRes := monthReg2.FindStringSubmatch(monthReg1.FindStringSubmatch(fname)[0])[0]
		sheets = append(sheets, Sheet {
			name: "内容创作者",
			start: "运营部门",
			file: f,
			data: make(map[int][]string),
			month: monthRes,
			typeIndex: 0,
		})
	}

	targetSheet := &Sheet {
		name: "大神内域作者费用明细",
		row: 1,
		col: 1,
		file: c.DstFiles["collect.xlsx"],
		data: make(map[int][]string),
	}

	for _, sheet := range sheets {
		if err := sheet.ReadSheet(); err != nil {
			fmt.Println(err)
			continue
		}
		if err := targetSheet.WriteSheet(&sheet); err != nil {
			fmt.Println(err)
		}
	}
}
