package xlsxrange

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

var reg1 *regexp.Regexp = regexp.MustCompile(`^(\$?[A-Z]+|\$?[1-9][0-9]*|\$?[A-Z]+\$?[1-9][0-9]*)(:(\$?[A-Z]+|\$?[1-9][0-9]*|\$?[A-Z]+\$?[1-9][0-9]*))?$`)
var reg2 *regexp.Regexp = regexp.MustCompile(`\$?([A-Z]+|\$?[1-9][0-9]*)`)
var reg3 *regexp.Regexp = regexp.MustCompile(`\$?([A-Z]+)`)
var reg4 *regexp.Regexp = regexp.MustCompile(`\$?([1-9][0-9]*)`)

// ParseA1Notation parses A1 notation and return sheet name and range.
//
// It can parse regular Excel style notations:
//  ParseA1Notation("D3:E8")
//  // Output: "", 4, 3, 2, 6, nil
//  ParseA1Notation("Sheet 1!D3:E8")
//  // Output: "Sheet 1", 4, 3, 2, 6, nil
func ParseA1Notation(notation string) (string, []int, error) {
	sheetName, rangeNotation := divideA1Notation(notation)
	rangeNotation = strings.ToUpper(rangeNotation)
	if reg1.MatchString(rangeNotation) {
		res := strings.Split(rangeNotation, ":")
		newRow := 0
		newColumn := 0
		newNumRows := 0
		newNumColumns := 0

		for i, subStr := range res {
			for _, value := range reg2.FindAllString(subStr, -1) {
				if reg3.MatchString(value) {
					if i == 0 {
						newColumn = ColumnStrToNumber(value)
					} else {
						col := ColumnStrToNumber(value)
						if newColumn < col {
							newNumColumns = col - newColumn + 1
						} else {
							newNumColumns = newColumn - col + 1
							newColumn = col
						}
					}
				} else if reg4.MatchString(value) {
					if strings.HasPrefix(value, "$") {
						value = value[1:]
					}
					if i == 0 {
						intValue, err := strconv.ParseInt(value, 10, 64)
						if err != nil {
							newRow = 0
						} else {
							newRow = int(intValue)
						}
					} else {
						row := 0
						intValue, err := strconv.ParseInt(value, 10, 64)
						if err == nil {
							row = int(intValue)
						}
						if newRow < row {
							newNumRows = row - newRow + 1
						} else {
							newNumRows = newRow - row + 1
							newRow = row
						}
					}
				}
			}
		}

		if newRow == 0 {
			if newNumRows == 0 {
				newRow = AllRows
				newNumRows = AllRows
			} else {
				newRow = newNumRows - 1
				newNumRows = AllRows
			}
		} else if newNumRows == 0 {
			if len(res) > 1 {
				newNumRows = AllRows
			} else {
				newNumRows = 1
			}
		}

		if newColumn == 0 {
			if newNumColumns == 0 {
				newColumn = AllColumns
				newNumColumns = AllColumns
			} else {
				newColumn = newNumColumns - 1
				newNumColumns = AllColumns
			}
		} else if newNumColumns == 0 {
			if len(res) > 1 {
				newNumColumns = AllColumns
			} else {
				newNumColumns = 1
			}
		}
		ranges := []int{newRow, newColumn, newNumRows, newNumColumns}
		return sheetName, ranges, nil
	}
	return sheetName, nil, fmt.Errorf(`'%s' is invalid A1Notation`, notation)
}

func divideA1Notation(notation string) (string, string) {
	index := strings.LastIndex(notation, "!")
	if index == -1 {
		// Sheet name is not included
		return "", notation
	} else {
		//  Sheet name is included
		sheetNamePart := notation[0:index]
		rangePart := notation[index+1:]

		pattern := regexp.MustCompile("^'(.*)'$")
		strippedNamePart := sheetNamePart
		match := pattern.FindStringSubmatch(sheetNamePart)
		if len(match) > 0 {
			strippedNamePart = match[1]
		}
		if len(strippedNamePart) != len(sheetNamePart) {
			strippedNamePart = strings.Replace(strippedNamePart, "''", "'", -1)
		}
		return strippedNamePart, rangePart
	}
}
