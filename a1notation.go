package xlsxrange

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

var allRowsPattern *regexp.Regexp = regexp.MustCompile(`^\$?([A-Z]+):\$?([A-Z]+)$`)
var allColsPattern *regexp.Regexp = regexp.MustCompile(`^\$?([1-9][0-9]*):\$?([1-9][0-9]*)$`)
var otherPattern *regexp.Regexp = regexp.MustCompile(`^\$?([A-Z]+)\$?([1-9][0-9]*)(:\$?([A-Z]+)\$?([1-9][0-9]*))?$`)

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

	match := false

	newRow := 0
	newColumn := 0
	newNumRows := 0
	newNumColumns := 0

	subStrings1 := allRowsPattern.FindStringSubmatch(rangeNotation)
	if len(subStrings1) > 0 {
		match = true
		// A:B pattern
		newRow = 1
		newColumn = ColumnStrToNumber(subStrings1[1])
		newNumRows = AllRows
		newNumColumns = ColumnStrToNumber(subStrings1[2]) - newColumn + 1
	} else {
		subStrings2 := allColsPattern.FindStringSubmatch(rangeNotation)
		if len(subStrings2) > 0 {
			match = true
			// 1:2 pattern
			newRow, _ = strconv.Atoi(subStrings2[1])
			newColumn = 1
			newNumRows, _ = strconv.Atoi(subStrings2[2])
			newNumRows = newNumRows - newRow + 1
			newNumColumns = AllColumns
		} else {
			subStrings3 := otherPattern.FindStringSubmatch(rangeNotation)
			if len(subStrings3) > 0 {
				match = true
				newRow, _ = strconv.Atoi(subStrings3[2])
				newColumn = ColumnStrToNumber(subStrings3[1])
				if subStrings3[3] != "" {
					// A2:B4 pattern
					newNumRows, _ = strconv.Atoi(subStrings3[5])
					newNumRows = newNumRows - newRow + 1
					newNumColumns = ColumnStrToNumber(subStrings3[4]) - newColumn + 1
				} else {
					// A2 pattern
					newNumRows = 1
					newNumColumns = 1
				}
			}
		}
	}

	if match {
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
