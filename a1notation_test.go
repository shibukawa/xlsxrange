package xlsxrange

import "testing"

func TestParseA1Notation_1(t *testing.T) {
	sheetName, rangeValues, err := ParseA1Notation("B2:D5")

	if sheetName != "" {
		t.Errorf("Sheet name should be empty but %s", sheetName)
	}
	if rangeValues[0] != 2 || rangeValues[1] != 2 || rangeValues[2] != 4 || rangeValues[3] != 3 {
		t.Errorf("Range should be [2, 2, 4, 3] but [%d, %d, %d, %d]", rangeValues[0], rangeValues[1], rangeValues[2], rangeValues[3])
	}
	if err != nil {
		t.Errorf("Error should be nil but %v", err)
	}
}

func TestParseA1Notation_2(t *testing.T) {
	sheetName, rangeValues, err := ParseA1Notation("B2:B")

	if sheetName != "" {
		t.Errorf("Sheet name should be empty but %s", sheetName)
	}
	if rangeValues[0] != 2 || rangeValues[1] != 2 || rangeValues[2] != AllRows || rangeValues[3] != 1 {
		t.Errorf("Range should be [2, 2, AllRows, 1] but [%d, %d, %d, %d]", rangeValues[0], rangeValues[1], rangeValues[2], rangeValues[3])
	}
	if err != nil {
		t.Errorf("Error should be nil but %v", err)
	}
}

func TestParseA1Notation_3(t *testing.T) {
	sheetName, rangeValues, err := ParseA1Notation("B")

	if sheetName != "" {
		t.Errorf("Sheet name should be empty but %s", sheetName)
	}
	if rangeValues[0] != AllRows || rangeValues[1] != 2 || rangeValues[2] != AllRows || rangeValues[3] != 1 {
		t.Errorf("Range should be [AllRows, 2, AllRows, 1] but [%d, %d, %d, %d]", rangeValues[0], rangeValues[1], rangeValues[2], rangeValues[3])
	}
	if err != nil {
		t.Errorf("Error should be nil but %v", err)
	}
}

func TestParseA1Notation_4(t *testing.T) {
	sheetName, rangeValues, err := ParseA1Notation("B2:2")

	if sheetName != "" {
		t.Errorf("Sheet name should be empty but %s", sheetName)
	}
	if rangeValues[0] != 2 || rangeValues[1] != 2 || rangeValues[2] != 1 || rangeValues[3] != AllColumns {
		t.Errorf("Range should be [2, 2, 1, AllColumns] but [%d, %d, %d, %d]", rangeValues[0], rangeValues[1], rangeValues[2], rangeValues[3])
	}
	if err != nil {
		t.Errorf("Error should be nil but %v", err)
	}
}

func TestParseA1Notation_5(t *testing.T) {
	sheetName, rangeValues, err := ParseA1Notation("2")

	if sheetName != "" {
		t.Errorf("Sheet name should be empty but %s", sheetName)
	}
	if rangeValues[0] != 2 || rangeValues[1] != AllColumns || rangeValues[2] != 1 || rangeValues[3] != AllColumns {
		t.Errorf("Range should be [2, AllColumns, 1, AllColumns] but [%d, %d, %d, %d]", rangeValues[0], rangeValues[1], rangeValues[2], rangeValues[3])
	}
	if err != nil {
		t.Errorf("Error should be nil but %v", err)
	}
}

func TestParseA1Notation_6(t *testing.T) {
	sheetName, rangeValues, err := ParseA1Notation("b2:d5")

	if sheetName != "" {
		t.Errorf("Sheet name should be empty but %s", sheetName)
	}
	if rangeValues[0] != 2 || rangeValues[1] != 2 || rangeValues[2] != 4 || rangeValues[3] != 3 {
		t.Errorf("Range should be [2, 2, 4, 3] but [%d, %d, %d, %d]", rangeValues[0], rangeValues[1], rangeValues[2], rangeValues[3])
	}
	if err != nil {
		t.Errorf("Error should be nil but %v", err)
	}
}

func TestParseA1Notation_7(t *testing.T) {
	sheetName, rangeValues, err := ParseA1Notation("$B2:D$5")
	if err != nil {
		t.Errorf("Error should be nil but %v", err)
		return
	}
	if sheetName != "" {
		t.Errorf("Sheet name should be empty but %s", sheetName)
	}
	if rangeValues[0] != 2 || rangeValues[1] != 2 || rangeValues[2] != 4 || rangeValues[3] != 3 {
		t.Errorf("Range should be [2, 2, 4, 3] but [%d, %d, %d, %d]", rangeValues[0], rangeValues[1], rangeValues[2], rangeValues[3])
	}
}

func TestParseA1Notation_8(t *testing.T) {
	sheetName, rangeValues, err := ParseA1Notation("$B$2:$D$5")

	if err != nil {
		t.Errorf("Error should be nil but %v", err)
		return
	}
	if sheetName != "" {
		t.Errorf("Sheet name should be empty but %s", sheetName)
	}
	if rangeValues[0] != 2 || rangeValues[1] != 2 || rangeValues[2] != 4 || rangeValues[3] != 3 {
		t.Errorf("Range should be [2, 2, 4, 3] but [%d, %d, %d, %d]", rangeValues[0], rangeValues[1], rangeValues[2], rangeValues[3])
	}
}

func TestDivideA1Notation_1(t *testing.T) {
	sheet, notation := divideA1Notation("filename!A1:B5")
	if sheet != "filename" || notation != "A1:B5" {
		t.Errorf(` DivideA1Notation("filename!A1:B5") should return 'filename' and 'A1:B5', but '%s' and '%s'.`, sheet, notation)
	}
}

func TestDivideA1Notation_2(t *testing.T) {
	sheet, notation := divideA1Notation("A1:B5")
	if sheet != "" || notation != "A1:B5" {
		t.Errorf(` DivideA1Notation("A1:B5") should return '' and 'A1:B5', but '%s' and '%s'.`, sheet, notation)
	}
}

func TestDivideA1Notation_3(t *testing.T) {
	sheet, notation := divideA1Notation("'filename'!A1:B5")
	if sheet != "filename" || notation != "A1:B5" {
		t.Errorf(` DivideA1Notation("'filename'!A1:B5") should return 'filename' and 'A1:B5', but '%s' and '%s'.`, sheet, notation)
	}
}

func TestDivideA1Notation_4(t *testing.T) {
	sheet, notation := divideA1Notation("'filename'''!A1:B5")
	if sheet != "filename'" || notation != "A1:B5" {
		t.Errorf(` DivideA1Notation("'filename''!A1:B5") should return 'filename'' and 'A1:B5', but '%s' and '%s'.`, sheet, notation)
	}
}
