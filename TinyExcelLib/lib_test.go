package tinyexcellib

import (
	"fmt"
	"testing"
)

func Test_Indexing1(t *testing.T) {

	excel := []string{"A1", "B2", "N3", "X4", "Z5", "AA6", "AF7", "AM8", "AZ9", "BF10", "ANFO11", "FENK12"}
	zeroBased := [][]int{
		{0, 0}, {1, 1}, {2, 13}, {3, 23}, {4, 25}, {5, 26},
		{6, 31}, {7, 38}, {8, 51}, {9, 57}, {10, 27210}, {11, 109210},
	}

	var i, j int
	var res string
	for idx, val := range excel {
		if i, j = excelToZeroBasedIndeces(val); i != zeroBased[idx][0] || j != zeroBased[idx][1] {
			fmt.Printf("%3d) excelToZeroBasedIndeces: %s, expecting %d %d, got %d %d\n", idx, val, zeroBased[idx][0], zeroBased[idx][1], i, j)
			t.FailNow()
		}

		if res = zeroBasedIndecesToExcel(zeroBased[idx][0], zeroBased[idx][1]); res != val {
			fmt.Printf("%3d) zeroBasedIndecesToExcel: %d %d, expecting %s got %s\n", idx, zeroBased[idx][0], zeroBased[idx][1], val, res)
			t.FailNow()
		}
	}
}

func Test_Indexing2(t *testing.T) {
	var a, b int
	for i := 0; i < 1000; i++ {
		a, b = excelToZeroBasedIndeces(zeroBasedIndecesToExcel(i+i, i))
		if a != i+i || b != i {
			fmt.Printf("%3d) expecting %d %d, got %d %d\n", i, i+i, i, a, b)
			t.FailNow()
		}
	}
}

func Test_MaxIndex1(t *testing.T) {
	triplets := [][]string{{"F7", "C12", "F12"}, {"C12", "F8", "F12"}}

	for _, triplet := range triplets {
		if triplet[2] != getMaxExcelIndex(triplet[0], triplet[1]) {
			t.FailNow()
		}
	}
}
