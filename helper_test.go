package xlsxrange

import "testing"

func TestNumberToColumnStr(t *testing.T) {
	if NumberToColumnStr(1) != "A" {
		t.Errorf("NumberToColumnStr(1) should be 'A', but '%s'", NumberToColumnStr(1))
	}
	if NumberToColumnStr(2) != "B" {
		t.Errorf("NumberToColumnStr(2) should be 'B', but '%s'", NumberToColumnStr(2))
	}
	if NumberToColumnStr(3) != "C" {
		t.Errorf("NumberToColumnStr(3) should be 'C', but '%s'", NumberToColumnStr(3))
	}
	if NumberToColumnStr(26) != "Z" {
		t.Errorf("NumberToColumnStr(26) should be 'Z', but '%s'", NumberToColumnStr(26))
	}
	if NumberToColumnStr(27) != "AA" {
		t.Errorf("NumberToColumnStr(27) should be 'AA', but '%s'", NumberToColumnStr(27))
	}
	if NumberToColumnStr(52) != "AZ" {
		t.Errorf("NumberToColumnStr(52) should be 'AZ', but '%s'", NumberToColumnStr(52))
	}
	if NumberToColumnStr(53) != "BA" {
		t.Errorf("NumberToColumnStr(53) should be 'BA', but '%s'", NumberToColumnStr(53))
	}
	if NumberToColumnStr(78) != "BZ" {
		t.Errorf("NumberToColumnStr(78) should be 'BZ', but '%s'", NumberToColumnStr(78))
	}
	if NumberToColumnStr(79) != "CA" {
		t.Errorf("NumberToColumnStr(79) should be 'CA', but '%s'", NumberToColumnStr(79))
	}
}

func TestColumnStrToNumber(t *testing.T) {
	if ColumnStrToNumber("A") != 1 {
		t.Errorf(`ColumnStrToNumber("A") should be 1, but '%d'`, ColumnStrToNumber("A"))
	}
	if ColumnStrToNumber("B") != 2 {
		t.Errorf(`ColumnStrToNumber("B") should be 2, but '%d'`, ColumnStrToNumber("B"))
	}
	if ColumnStrToNumber("C") != 3 {
		t.Errorf(`ColumnStrToNumber("C") should be 3, but '%d'`, ColumnStrToNumber("C"))
	}
	if ColumnStrToNumber("Z") != 26 {
		t.Errorf(`ColumnStrToNumber("Z") should be 26, but '%d'`, ColumnStrToNumber("Z"))
	}
	if ColumnStrToNumber("AA") != 27 {
		t.Errorf(`ColumnStrToNumber("AA") should be 27, but '%d'`, ColumnStrToNumber("AA"))
	}
	if ColumnStrToNumber("AZ") != 52 {
		t.Errorf(`ColumnStrToNumber("AZ") should be 52, but '%d'`, ColumnStrToNumber("AZ"))
	}
	if ColumnStrToNumber("BA") != 53 {
		t.Errorf(`ColumnStrToNumber("BA") should be 53, but '%d'`, ColumnStrToNumber("BA"))
	}
	if ColumnStrToNumber("BZ") != 78 {
		t.Errorf(`ColumnStrToNumber("BZ") should be 78, but '%d'`, ColumnStrToNumber("BZ"))
	}
	if ColumnStrToNumber("CA") != 79 {
		t.Errorf(`ColumnStrToNumber("CA") should be 79, but '%d'`, ColumnStrToNumber("CA"))
	}
}
