// code for collect common sheets into specified sheet
// now, common sheets include "活动，CPS分发，新游预约"

package collect

import (
	"github.com/xuri/excelize/v2"
	"strconv"
	"strings"
	"time"
)

const (
	// indexs in Sheet struct
	startDate = 0
	endDate   = 1
	cmIdEnd   = 2 // increase this number when add new index
)

var commonMap = map[string]int{
	// key col start from zero for each common sheet
	"活动":    4, // 入库活动名列
	"CPS分发": 5, // 项目名称列
	"新游预约":  4, // 项目名称列
}

func timeToExcelTime(t time.Time) (float64, error) {
	const (
		dayNanoseconds = 24 * time.Hour
		maxDuration    = 290 * 364 * dayNanoseconds
	)
	excelMinTime1900 := time.Date(1899, time.December, 31, 0, 0, 0, 0, time.UTC)
	excelBuggyPeriodStart := time.Date(1900, time.March, 1, 0, 0, 0, 0, time.UTC).Add(-time.Nanosecond)

	if t.Before(excelMinTime1900) {
		return 0.0, nil
	}

	tt := t
	diff := t.Sub(excelMinTime1900)
	result := float64(0)

	for diff >= maxDuration {
		result += float64(maxDuration / dayNanoseconds)
		tt = tt.Add(-maxDuration)
		diff = tt.Sub(excelMinTime1900)
	}

	rem := diff % dayNanoseconds
	result += float64(diff-rem)/float64(dayNanoseconds) + float64(rem)/float64(dayNanoseconds)

	if t.After(excelBuggyPeriodStart) {
		result += 1.0
	}
	return result, nil
}

func convertIfDate(f *excelize.File, sheetName string, col int, row int, data string) (string, bool) {
	coords, _ := excelize.CoordinatesToCellName(col, row)
	style, err := f.GetCellStyle(sheetName, coords)
	if style == 0 {
		return "", false
	}
	styleSheet := f.Styles
	if style >= len(styleSheet.CellXfs.Xf) {
		return "", false
	}
	var numFmtID int
	if styleSheet.CellXfs.Xf[style].NumFmtID != nil {
		numFmtID = *styleSheet.CellXfs.Xf[style].NumFmtID
	}
	var timeFormat string
	switch numFmtID {
	case 14:
		timeFormat = "01-02-06" // "mm-dd-yy"
	case 15:
		timeFormat = "02-Jan-06" // "d-mmm-yy"
	case 16:
		timeFormat = "02-Jan" // "d-mmm"
	case 17:
		timeFormat = "Jan-06" // "mmm-yy"
	case 22:
		timeFormat = "1/2/06 15:04" // "m/d/yy h:mm"
	default:
		return "", false
	}
	t, err := time.Parse(timeFormat, data)
	if err != nil {
		return "", false
	}
	dateExcel, err := timeToExcelTime(t)
	if err != nil {
		return "", false
	}
	return strconv.FormatInt(int64(dateExcel), 10), true
}

func parseDate(dateRaw string) (string, error) {
	// support "2021.09.12", or "2021.9.12"
	// support "2021/09/12", or "2021/9/12"
	date := ""
	dateSlice := strings.FieldsFunc(dateRaw, func(r rune) bool {
		return r == '.' || r == '/'
	})
	if len(dateSlice) == 1 {
		// no split delimiter found, go on and support "20210912"
		if len(dateSlice[0]) == 8 {
			date = dateSlice[0]
		} else {
			return dateRaw, nil
		}
	} else {
		if len(dateSlice) != 3 && len(dateSlice[0]) != 4 {
			return dateRaw, nil
		}
		if len(dateSlice[1]) < 1 || len(dateSlice[1]) > 2 {
			return dateRaw, nil
		} else if len(dateSlice[1]) == 1 {
			dateSlice[1] = "0" + dateSlice[1]
		}
		if len(dateSlice[2]) < 1 || len(dateSlice[2]) > 2 {
			return dateRaw, nil
		} else if len(dateSlice[2]) == 1 {
			dateSlice[2] = "0" + dateSlice[2]
		}
		date = dateSlice[0] + dateSlice[1] + dateSlice[2]
	}

	dateTime, err := time.Parse("20060102", date)
	if err != nil {
		return dateRaw, err
	}
	dateExcel, err := timeToExcelTime(dateTime)
	if err != nil {
		return dateRaw, err
	}
	return strconv.FormatInt(int64(dateExcel), 10), nil
}

func (s *Sheet) ReadSheetAll() error {
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
					for id, colData := range colsData {
						if strings.Contains(colData, "开始日期") {
							s.indexs[startDate] = id
						} else if strings.Contains(colData, "结束日期") {
							s.indexs[endDate] = id
						}
					}
				} else if colsData == nil {
					break // absolutely data end
				} else if len(colsData) > commonMap[s.name] {
					needBreak := true // maybe data end
					for i := 0; i <= commonMap[s.name]; i++ {
						if len(colsData[i]) != 0 {
							needBreak = false
							break
						}
					}
					if needBreak {
						// (len(colsData[0]) == 0) && ... && (len(colsData[n]) == 0) is true
						break
					}

					needContinue := false // true for data not enough
					if strings.Contains(colsData[commonMap[s.name]], "辅助列") {
						continue // skip this row
					}
					for i := 0; i <= commonMap[s.name]; i++ {
						if len(colsData[i]) == 0 {
							needContinue = true
							break
						}
					}
					if needContinue {
						// (len(colsData[0]) == 0) || ... || (len(colsData[n]) == 0) is true
						continue
					}
				} else if len(colsData) <= commonMap[s.name] {
					break // maybe data end
				}

				if curRow != s.row {
					// deal with date formatted col
					if dateExcel, isCv := convertIfDate(s.file, sheetName, s.indexs[startDate]+1, curRow, colsData[s.indexs[startDate]]); isCv {
						colsData[s.indexs[startDate]] = dateExcel
					}
					if dateExcel, isCv := convertIfDate(s.file, sheetName, s.indexs[endDate]+1, curRow, colsData[s.indexs[endDate]]); isCv {
						colsData[s.indexs[endDate]] = dateExcel
					}
				}
				s.data = append(s.data, colsData)
			}
		}
	}
	return nil
}

func (s *Sheet) WriteSheetAll(from *Sheet) error {
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
	dateStyle, _ := s.file.NewStyle(`{"number_format": 14}`)
	for row, colsData := range from.data {
		if row == 0 && s.row != 1 {
			continue
		}
		for col, colData := range colsData {
			dstAxis, _ = excelize.CoordinatesToCellName(col+1, s.row)
			// deal with date
			if row != 0 && (col == from.indexs[startDate] || col == from.indexs[endDate]) {
				if err := s.file.SetCellStyle(s.name, dstAxis, dstAxis, dateStyle); err != nil {
					return err
				}
				// maybe invalid date in text, such as "20210912", or "2021.09.12", or "2021.9.12"
				colData, err = parseDate(colData)
				if err != nil {
					return err
				}
			}

			err = s.file.SetCellValue(s.name, dstAxis, colData)
			if err != nil {
				return err
			}

		}
		s.row++
	}

	if err = s.file.Save(); err != nil {
		return err
	}

	return nil
}

// CollectForAll collect for 活动，CPS分发，新游预约
func (c *Collect) CollectForAll(keyword string) error {
	sheets := make([]Sheet, 0)
	for _, f := range c.srcFiles {
		sheets = append(sheets, Sheet{
			name:   keyword,
			start:  "运营部门",
			file:   f,
			indexs: make([]int, cmIdEnd),
		})
	}

	targetSheet := &Sheet{
		name: keyword,
		row:  1,
		col:  1,
		file: c.dstFiles["项目立项及实际费用明细.xlsx"],
	}

	for _, sheet := range sheets {
		if err := sheet.ReadSheetAll(); err != nil {
			return err
		}
		if err := targetSheet.WriteSheetAll(&sheet); err != nil {
			return err
		}
	}

	return nil
}
