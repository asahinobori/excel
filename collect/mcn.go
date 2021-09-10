// code for collect mcn , namely "MCN", sheet into "MCN外域作者费用明细" sheet
// parse month("月份") from file name

package collect

import (
	"github.com/xuri/excelize/v2"
	"regexp"
	"strings"
)

const (
	// col start from zero for src
	mcnGame     = 1  // 游戏产品
	mcnSponsor  = 2  // 出资方
	mcnType     = 4  // 结算类型
	mcnPlf      = 5  // 平台
	mcnUid      = 6  // 平台UID
	mcnNickName = 7  // 平台昵称
	mcnMoney    = 11 // 金额（橙列）
	mcnViewCnt  = 12 // 播放量

	// col start from zero for dst
	mcnMonthD = 0 // 月份
	// mcnOrgD        = 1  // 机构
	mcnGameD     = 2 // 游戏
	mcnPlfD      = 3 // 平台
	mcnUidD      = 4 // 平台UID
	mcnNickNameD = 5 // 平台昵称
	mcnMoneyD    = 6 // 作者费用
	mcnViewCntD  = 7 // 播放量
	mcnTypeD     = 8 // 类别
	mcnSponsorD  = 9 // 出资方

	// indexs in Sheet struct
	mcnTypeId = 0
	mcnIdEnd  = 1 // increase this number when add new index
)

var mcnMap = map[int]int{
	mcnGame:     mcnGameD,
	mcnSponsor:  mcnSponsorD,
	mcnType:     mcnTypeD,
	mcnPlf:      mcnPlfD,
	mcnUid:      mcnUidD,
	mcnNickName: mcnNickNameD,
	mcnMoney:    mcnMoneyD,
	mcnViewCnt:  mcnViewCntD,
}

func (s *Sheet) ReadSheetMcn() error {
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

			for _, found := range startFound {
				s.col, s.row, _ = excelize.CellNameToCoordinates(found)

				// traverse this sheet and get data from start coordinate
				curRow := 0
				rowsIt, err := s.file.Rows(sheetName)
				if err != nil {
					return err
				}
				for rowsIt.Next() {
					curRow++
					colsData, err := rowsIt.Columns()
					if colsData != nil {
						colsData = colsData[s.col-1:]
					}
					if err != nil {
						return err
					} else if curRow < s.row {
						continue
					} else if curRow == s.row {
						for id, colData := range colsData {
							if strings.Contains(colData, "结算类型") {
								s.indexs[mcnTypeId] = id
							}
						}
						continue
					} else if colsData == nil {
						break // absolutely data end
					} else if len(colsData) > mcnUid {
						needBreak := true // maybe data end
						for i := s.col - 1; i <= mcnUid; i++ {
							if len(colsData[i]) != 0 {
								needBreak = false
								break
							}
						}
						if needBreak {
							// (len(colsData[0]) == 0) && ... && (len(colsData[n]) == 0) is true
							break
						}

						if s.indexs[mcnTypeId] != 0 && !strings.Contains(colsData[s.indexs[mcnTypeId]], "自孵化") && !strings.Contains(colsData[s.indexs[mcnTypeId]], "签约作者") {
							break // skip this type
						}

						needContinue := false // true for data not enough
						for i := s.col - 1; i <= mcnUid; i++ {
							if len(colsData[i]) == 0 {
								needContinue = true
								break
							}
						}
						if needContinue {
							// (len(colsData[0]) == 0) || ... || (len(colsData[n]) == 0) is true
							continue
						}
					} else if len(colsData) <= mcnUid {
						break // maybe data end
					}

					colsData[mcnType] = colsData[s.indexs[mcnTypeId]]

					s.data = append(s.data, colsData)
				}
			}
		}
	}
	return nil
}

func (s *Sheet) WriteSheetMcn(from *Sheet) error {
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
	exp := "yyyy\"年\"m\"月\""
	monthStyle, err := s.file.NewStyle(&excelize.Style{CustomNumFmt: &exp})
	for _, colsData := range from.data {
		s.row++
		for col, colData := range colsData {
			if dstCol, ok := mcnMap[col]; ok {
				dstAxis, _ = excelize.CoordinatesToCellName(dstCol+1, s.row)
				err = s.file.SetCellValue(s.name, dstAxis, colData)
				if err != nil {
					return err
				}
			}
		}

		// deal with month
		dstAxis, _ = excelize.CoordinatesToCellName(mcnMonthD+1, s.row)
		err = s.file.SetCellValue(s.name, dstAxis, "2021/"+from.month+"/1")
		if err != nil {
			return err
		}

		err = s.file.SetCellStyle(s.name, dstAxis, dstAxis, monthStyle)
		if err != nil {
			return err
		}
	}

	if err = s.file.Save(); err != nil {
		return err
	}

	return nil
}

func (c *Collect) CollectForMcn() error {
	sheets := make([]Sheet, 0)
	for fname, f := range c.srcFiles {
		monthReg1 := regexp.MustCompile(`[^\d]\d+月`)
		monthReg2 := regexp.MustCompile(`\d+`)
		monthRes := monthReg2.FindStringSubmatch(monthReg1.FindStringSubmatch(fname)[0])[0]
		sheets = append(sheets, Sheet{
			name:   "MCN",
			start:  "运营部门",
			file:   f,
			month:  monthRes,
			indexs: make([]int, mcnIdEnd),
		})
	}

	targetSheet := &Sheet{
		name: "MCN外域作者费用明细",
		row:  1,
		col:  1,
		file: c.dstFiles["项目立项及实际费用明细.xlsx"],
		fileMutex: c.dstFilesMutex["项目立项及实际费用明细.xlsx"],
	}

	for _, sheet := range sheets {
		if err := sheet.ReadSheetMcn(); err != nil {
			return err
		}
		if err := targetSheet.WriteSheetMcn(&sheet); err != nil {
			return err
		}
	}

	return nil
}
