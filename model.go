package excel

import (
	"github.com/xuri/excelize/v2"
	"reflect"
	"time"
)

var defaultTimeLayouts = []string{
	time.RFC3339, "2006-01-02 15:04:05", "2006-01-02", "02/01/2006",
}

type (
	Excel[T any] struct {
		file   *excelize.File
		IsNext bool
		rt     reflect.Type
		rows   *excelize.Rows
		rules  []fieldRule
		opt    Opt
	}

	Opt struct {
		HeaderRow    uint8
		DataStartRow uint8
		Limit        uint
	}

	fieldRule struct {
		fieldIdx int
		colIdx   int    // 0-based
		header   string // header name (for error)
		required bool
		layout   string // for time.Time
	}
)
