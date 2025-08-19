package excel

import (
	"github.com/xuri/excelize/v2"
	"reflect"
	"strings"
	"sync"
	"time"
)

var defaultTimeLayouts = []string{
	time.RFC3339, "2006-01-02 15:04:05", "2006-01-02", "02/01/2006",
}

// Performance-optimized pools
var (
	stringBuilderPool = sync.Pool{
		New: func() interface{} {
			return new(strings.Builder)
		},
	}

	bufferPool = sync.Pool{
		New: func() interface{} {
			return make([]byte, 0, 64)
		},
	}

	fieldCache = sync.Map{}
)

type (
	Excel[T any] struct {
		file     *excelize.File
		IsNext   bool
		rt       reflect.Type
		rows     *excelize.Rows
		rules    []fieldRule
		opt      Opt
		workers  int
		fieldMap map[string]int   // Cache for field lookups
		typeInfo *typeInformation // Cache for type information
	}

	Opt struct {
		HeaderRow    uint8
		DataStartRow uint8
		Limit        uint
		Workers      int  // Number of worker goroutines
		UseCache     bool // Enable field caching
	}

	fieldRule struct {
		fieldIdx  int
		colIdx    int
		header    string
		required  bool
		layout    string
		fastPath  bool      // Indicates if fast path processing is available
		converter converter // Fast path conversion function
	}

	typeInformation struct {
		fields     []reflect.StructField
		converters []converter
		fieldMap   map[string]int
	}

	converter func(string) (interface{}, error)
)

// Cache type information
func cacheTypeInfo(t reflect.Type) *typeInformation {
	info := &typeInformation{
		fieldMap: make(map[string]int),
	}

	for i := 0; i < t.NumField(); i++ {
		field := t.Field(i)
		info.fields = append(info.fields, field)
		info.fieldMap[field.Name] = i

		// Create optimized converter for the field
		conv := getOptimizedConverter(field.Type)
		info.converters = append(info.converters, conv)
	}

	return info
}

// Get cached type information
func getTypeInfo(t reflect.Type) *typeInformation {
	if cached, ok := fieldCache.Load(t); ok {
		return cached.(*typeInformation)
	}

	info := cacheTypeInfo(t)
	fieldCache.Store(t, info)
	return info
}
