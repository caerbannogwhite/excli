package tinyexcellib

import (
	"fmt"
	"strconv"
)

const ALPHA_LEN int = 26

func divMod(n, d int) (q, r int) {
	q = n / d
	r = n % d

	// the numerator or the denominator is negarive (but not both)
	if r != 0 && n*d < 0 {
		q--
		r += d
	}
	return
}

func excelToZeroBasedIndeces(index string) (int, int) {
	i := 0
	pi := 1
	j := 0
	pj := 1

	for idx := len(index) - 1; idx >= 0; idx-- {
		if d, err := strconv.Atoi(string(index[idx])); err == nil {
			i += d * pi
			pi *= 10
		} else {
			j += int(index[idx]-64) * pj
			pj *= ALPHA_LEN
		}
	}

	return i - 1, j - 1
}

func zeroBasedIndecesToExcel(i int, j int) string {
	res := ""
	var r int
	for j >= 0 {
		j, r = divMod(j, ALPHA_LEN)
		j--
		res = string(65+byte(r)) + res
	}
	return fmt.Sprintf("%s%d", res, i+1)
}

// Given 2 Excel indeces, return an index
// with the max row and column
// Ex.: if a is 'F7' and b is 'C12',
// return 'F12'
func getMaxExcelIndex(a, b string) string {
	iMax, jMax := excelToZeroBasedIndeces(a)
	i, j := excelToZeroBasedIndeces(b)

	if i > iMax {
		iMax = i
	}

	if j > jMax {
		jMax = j
	}

	return zeroBasedIndecesToExcel(iMax, jMax)
}
