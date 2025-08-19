package excel

import (
	"github.com/xuri/excelize/v2"
	"reflect"
	"sync"
	"time"
)

var defaultTimeLayouts = []string{
	time.RFC3339, "2006-01-02 15:04:05", "2006-01-02", "02/01/2006",
}

var (
	// Pool for reusing reflect.Value allocations
	reflectValuePool = sync.Pool{
		New: func() interface{} {
			return make([]reflect.Value, 0, 10)
		},
	}

	// Pool for reusing string slices
	stringSlicePool = sync.Pool{
		New: func() interface{} {
			return make([]string, 0, 50)
		},
	}
)

type (
	Excel[T any] struct {
		file         *excelize.File
		IsNext       bool
		rt           reflect.Type
		rows         *excelize.Rows
		rules        []fieldRule
		opt          Opt
		instancePool *sync.Pool // Instance pool for reflect.Value reuse
		batchSize    int        // processing batch size
	}

	Opt struct {
		HeaderRow    uint8
		DataStartRow uint8
		Limit        uint
		BatchSize    int // new field for batch processing
	}

	fieldRule struct {
		fieldIdx  int
		colIdx    int    // 0-based
		header    string // header name (for error)
		required  bool
		layout    string       // for time.Time
		fieldType reflect.Type // cached field type
		isPointer bool         // whether field is pointer type
		kindCache reflect.Kind // cached kind for performance
	}

	fieldType int

	fieldSetter func(rv reflect.Value, cell string, rule fieldRule) error
)

const (
	fieldTypeUnknown fieldType = iota
	fieldTypeString
	fieldTypeBool
	fieldTypeInt
	fieldTypeUint
	fieldTypeFloat
	fieldTypeTime
)
