// code for collect content("内容"), namely "内容创作者，内容采购", sheets into "大神内域作者费用明细" sheet
// parse month("月份") from file name
// parse organization("机构") from file which suffix is "csv"

package collect

import (
	"fmt"
	"github.com/dimchansky/utfbom"
	"github.com/gocarina/gocsv"
	"github.com/xuri/excelize/v2"
	"io/ioutil"
	"regexp"
	"strconv"
	"strings"
)

const (
	// col start from zero for src
	department = 0 // 运营部门
	game       = 1 // 游戏产品
	sponsor    = 2 // 出资方
	uid        = 4 // UID
	nickName   = 5 // 昵称
	srcEnd     = 9 // 可能结束的列数(目前应该是金额（橙列），它可能是第J列或第K列)

	// col start from zero for dst
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

	// indexs in Sheet struct
	dynType = 0
	readCnt = 1
	money   = 2 // 金额（橙列）
	ctIdEnd = 3 // increase this number when add new index
)

var contentMap = map[int]int{
	department: departmentD,
	game:       gameD,
	sponsor:    sponsorD,
	uid:        uidD,
	nickName:   nickNameD,
}

type Org struct {
	KolType string `csv:"kol_type"`
	Uid     string `csv:"uid"`
	AddDate string `csv:"add_date"`
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

func (s *Sheet) ReadSheetContent() error {
	sheetList := s.file.GetSheetList()
	for _, sheetName := range sheetList {
		// skip hidden sheet
		if !s.file.GetSheetVisible(sheetName) {
			continue
		}
		if strings.Contains(sheetName, s.name) && !strings.Contains(sheetName, "论坛") {
			startFound, err := s.file.SearchSheet(sheetName, s.start)
			if err != nil {
				return err
			} else if startFound == nil {
				continue
			}
			s.col, s.row, _ = excelize.CellNameToCoordinates(startFound[0])

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
					dynTypeFound := 0
					readCntFound := 0
					moneyFound := 0
					for id, colData := range colsData {
						if strings.Contains(colData, "动态类型") {
							s.indexs[dynType] = id
							dynTypeFound++
						} else if strings.Contains(colData, "阅读量") && !strings.Contains(colData, "求和") {
							s.indexs[readCnt] = id
							readCntFound++
						} else if strings.Contains(colData, "税前金额（自动计算)") && !strings.Contains(colData, "求和") {
							s.indexs[money] = id
							moneyFound++
						}
					}
					if dynTypeFound == 0 {
						fmt.Println("错误：“动态类型”列找不到", sheetName, s.fileName)
					} else if dynTypeFound > 1 {
						fmt.Println("错误：多于1个“动态类型”列", sheetName, s.fileName)
					}
					if readCntFound == 0 {
						fmt.Println("错误：“阅读量”列找不到", sheetName, s.fileName)
					} else if readCntFound > 1 {
						fmt.Println("错误：多于1个“阅读量”列", sheetName, s.fileName)
					}
					if moneyFound == 0 {
						fmt.Println("错误：“税前金额”列找不到", sheetName, s.fileName)
					} else if moneyFound > 1 {
						fmt.Println("错误：多于1个“税前金额”列", sheetName, s.fileName)
					}
					if dynTypeFound == 0 || readCntFound == 0 || moneyFound == 0 || dynTypeFound > 1 || readCntFound > 1 || moneyFound > 1 {
						break
					}

					continue
				} else if colsData == nil {
					break // absolutely data end
				} else if (len(colsData) > srcEnd) &&
					(len(colsData[uid]) == 0) && (len(colsData[nickName]) == 0) && (len(colsData[s.indexs[money]]) == 0) {
					break // maybe data end
				} else if (len(colsData) > srcEnd) &&
					((len(colsData[uid]) == 0) || (len(colsData[nickName]) == 0) || (len(colsData[s.indexs[money]-1]) == 0) || (len(colsData[s.indexs[money]]) == 0)) {
					continue // data not enough
				} else if len(colsData) <= srcEnd {
					break // maybe data end
				}
				s.data = append(s.data, colsData)
			}
		}
	}

	return nil
}

func (s *Sheet) WriteSheetContent(from *Sheet) error {
	s.fileMutex.Lock()
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
	s.fileMutex.Unlock()
	var dstAxis string
	var err error
	for _, colsData := range from.data {
		s.row++
		for col, colData := range colsData {
			if dstCol, ok := contentMap[col]; ok {
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
		if (from.indexs[dynType] == 0) || (from.indexs[dynType] >= len(colsData)) || (colsData[from.indexs[dynType]] == "") {
			dstAxis, _ := excelize.CoordinatesToCellName(unclsMoneyD+1, s.row)
			err = s.file.SetCellValue(s.name, dstAxis, colsData[from.indexs[money]])
		} else if strings.Contains(colsData[from.indexs[dynType]], "视频") {
			dstAxis, _ := excelize.CoordinatesToCellName(videoMoneyD+1, s.row)
			err = s.file.SetCellValue(s.name, dstAxis, colsData[from.indexs[money]])
		} else {
			dstAxis, _ := excelize.CoordinatesToCellName(textMoneyD+1, s.row)
			err = s.file.SetCellValue(s.name, dstAxis, colsData[from.indexs[money]])
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
		if from.indexs[readCnt] != 0 {
			dstAxis, _ = excelize.CoordinatesToCellName(readCntD+1, s.row)
			err = s.file.SetCellValue(s.name, dstAxis, colsData[from.indexs[readCnt]])
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

func (c *Collect) CollectForContent() error {
	orgsMap := make(map[string]string)
	if err := c.ReadCSV(orgsMap); err != nil {
		return err
	}

	sheets := make([]Sheet, 0)
	for fname, f := range c.srcFiles {
		monthReg1 := regexp.MustCompile(`[^\d]\d+月`)
		monthReg2 := regexp.MustCompile(`\d+`)
		monthRes := monthReg2.FindStringSubmatch(monthReg1.FindStringSubmatch(fname)[0])[0]
		sheets = append(sheets, Sheet{
			name:     "内容创作者",
			start:    "运营部门",
			file:     f,
			fileName: fname,
			month:    monthRes,
			indexs:   make([]int, ctIdEnd),
		})
		sheets = append(sheets, Sheet{
			name:     "内容采购",
			start:    "运营部门",
			file:     f,
			fileName: fname,
			month:    monthRes,
			indexs:   make([]int, ctIdEnd),
		})
	}

	targetSheet := &Sheet{
		name:      "大神内域作者费用明细",
		row:       1,
		col:       1,
		file:      c.dstFiles["项目立项及实际费用明细.xlsx"],
		fileName:  "项目立项及实际费用明细.xlsx",
		fileMutex: c.dstFilesMutex["项目立项及实际费用明细.xlsx"],
		org:       orgsMap,
	}

	for _, sheet := range sheets {
		if err := sheet.ReadSheetContent(); err != nil {
			return err
		}
		if err := targetSheet.WriteSheetContent(&sheet); err != nil {
			return err
		}
	}

	return nil
}
