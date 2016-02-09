package xlsxrange

import (
	"bytes"
	"fmt"
	"math"
	"strings"
)

// NumberToColumnStr() converts column number to column label string like A, B, C.. AE, EF...
func NumberToColumnStr(c int) string {
	if c < 1 {
		panic(fmt.Sprintf("NumberToColumnStr: column should be bigger than 1, but %d\n", c))
	}
	// 逆順
	var chars []byte
	c = c - 1
	for {
		code := c % 26
		chars = append(chars, byte(code))
		c = int(math.Floor(float64(c) / 26.0))
		if c == 0 {
			break
		}
	}
	var buffer bytes.Buffer
	for i := len(chars) - 1; i != -1; i-- {
		if i != 0 && len(chars) > 1 {
			buffer.WriteByte(byte("A"[0]) + chars[i] - 1)
		} else {
			buffer.WriteByte(byte("A"[0]) + chars[i])
		}
	}
	return buffer.String()
}

// ColumnStrToNumber converts column label string (A, B, C.. AE, EF...) to column number.
func ColumnStrToNumber(s string) int {
	s = strings.ToUpper(s)
	if !reg3.MatchString(s) {
		return -1
	}
	result := 0
	var base = int("A"[0])
	var last = int("Z"[0])
	for i := 0; i < len(s); i++ {
		code := int(s[i])
		if base <= code && code <= last {
			result = result*26 + (code - base + 1)
		}
	}
	return result
}
