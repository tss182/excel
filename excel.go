package excel

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"mime/multipart"
	"reflect"
	"strconv"
	"strings"
	"sync"
	"time"
	"unicode/utf8"
)

func New[T any]() *Excel[T] {
	var zero T
	rt := reflect.TypeOf(zero)
	if rt.Kind() == reflect.Pointer {
		rt = rt.Elem()
	}

	excel := &Excel[T]{
		file: excelize.NewFile(),
		rt:   rt,
	}

	// Initialize instance pool for better memory management
	excel.instancePool = &sync.Pool{
		New: func() interface{} {
			return reflect.New(rt).Elem()
		},
	}

	return excel
}

func Open[T any](filename string) (*Excel[T], error) {
	file, err := excelize.OpenFile(filename)
	if err != nil {
		return nil, fmt.Errorf("failed to open file %s: %w", filename, err)
	}

	var zero T
	rt := reflect.TypeOf(zero)
	if rt.Kind() == reflect.Pointer {
		rt = rt.Elem()
	}

	excel := &Excel[T]{
		file: file,
		rt:   rt,
	}

	// Initialize instance pool for better memory management
	excel.instancePool = &sync.Pool{
		New: func() interface{} {
			return reflect.New(rt).Elem()
		},
	}

	return excel, nil
}

func OpenReader[T any](file multipart.File) (*Excel[T], error) {
	f, err := excelize.OpenReader(file)
	if err != nil {
		return nil, fmt.Errorf("failed to open reader: %w", err)
	}

	var zero T
	rt := reflect.TypeOf(zero)
	if rt.Kind() == reflect.Pointer {
		rt = rt.Elem()
	}

	excel := &Excel[T]{
		file: f,
		rt:   rt,
	}

	// Initialize instance pool for better memory management
	excel.instancePool = &sync.Pool{
		New: func() interface{} {
			return reflect.New(rt).Elem()
		},
	}

	return excel, nil
}

func (e *Excel[T]) Close() error {
	return e.file.Close()
}

func (e *Excel[T]) Read(out *[]T, sheetName string, opts ...Opt) error {
	var headerRow, dataStartRow uint8 = 1, 2
	var limit uint = 0
	var batchSize int = 1000 // default batch size
	if opts != nil {
		for _, opt := range opts {
			if opt.HeaderRow > 0 {
				headerRow = opt.HeaderRow
			}
			if opt.DataStartRow > 0 {
				dataStartRow = opt.DataStartRow
			}
			if opt.Limit > 0 {
				limit = opt.Limit
			}
			if opt.BatchSize > 0 {
				batchSize = opt.BatchSize
			}
		}
	}
	e.opt = Opt{
		HeaderRow:    headerRow,
		DataStartRow: dataStartRow,
		Limit:        limit,
		BatchSize:    batchSize,
	}
	e.batchSize = batchSize

	if e.file == nil {
		return errors.New("file didn't set")
	}
	//get a sheet
	_, err := e.file.GetSheetIndex(sheetName)
	if err != nil {
		return fmt.Errorf("sheet %s in file excel: %w", sheetName, errors.New("not found"))
	}

	e.rows, err = e.file.Rows(sheetName)
	if err != nil {
		return fmt.Errorf("failed to read rows: %w", err)
	}

	for i := uint8(0); i < headerRow; i++ {
		if !e.rows.Next() {
			return fmt.Errorf("sheet %s in file excel: %w", sheetName, errors.New("empty sheet"))
		}
	}

	header, err := e.rows.Columns()
	if err != nil {
		return fmt.Errorf("failed to read header: %w", err)
	}

	colIdxHeader := make(map[string]int, len(header))
	for i, h := range header {
		colIdxHeader[strings.TrimSpace(h)] = i
	}

	// Use cached type info if available, otherwise initialize
	if e.rt == nil {
		var zero T
		rt := reflect.TypeOf(zero)
		if rt.Kind() == reflect.Pointer {
			rt = rt.Elem()
		}
		if rt.Kind() != reflect.Struct {
			return errors.New("out must be a struct or *struct")
		}
		e.rt = rt

		// Initialize instance pool if not already done
		if e.instancePool == nil {
			e.instancePool = &sync.Pool{
				New: func() interface{} {
					return reflect.New(rt).Elem()
				},
			}
		}
	}

	e.rules, err = buildRules(e.rt, colIdxHeader)
	if err != nil {
		return err
	}

	err = e.getRows(out)
	if err != nil {
		return err
	}

	return e.rows.Error()
}

func (e *Excel[T]) CloseRow() error {
	return e.rows.Close()
}

func (e *Excel[T]) Next(out *[]T) error {
	if e.IsNext {
		return e.getRows(out)
	} else {
		return errors.New("no next row")
	}
}

func (e *Excel[T]) getRows(out *[]T) error {
	// Pre-allocate slice capacity if limit is known
	if e.opt.Limit > 0 && cap(*out) < int(e.opt.Limit) {
		newSlice := make([]T, len(*out), e.opt.Limit)
		copy(newSlice, *out)
		*out = newSlice
	}

	var numberData uint = 0
	for e.rows.Next() {
		numberData++
		rowNum := numberData + uint(e.opt.DataStartRow) - 1
		cols, err := e.rows.Columns()
		if err != nil {
			return fmt.Errorf("read row %d: %w", rowNum, err)
		}

		// Get instance from pool for better memory reuse
		rv := e.instancePool.Get().(reflect.Value)

		// Reset all fields to zero values for reuse
		rv.Set(reflect.Zero(e.rt))

		// Process each field rule with optimized access
		for i := range e.rules {
			rule := &e.rules[i] // Use pointer to avoid copying
			var cell string
			if rule.colIdx < len(cols) {
				cell = strings.TrimSpace(cols[rule.colIdx])
			}

			// Early validation for required fields
			if rule.required && cell == "" {
				// Return instance to pool before error
				e.instancePool.Put(rv)
				return fmt.Errorf("row %d col %s (%s) is required", rowNum, idxToCol(rule.colIdx), rule.header)
			}
			if cell == "" {
				continue
			}

			fv := rv.Field(rule.fieldIdx)
			if !fv.CanSet() {
				continue
			}

			// Use optimized field setting with cached type info
			if err := setFieldValueOptimized(fv, cell, rule); err != nil {
				// Return instance to pool before error
				e.instancePool.Put(rv)
				return fmt.Errorf("row %d col %s (%s): %w", rowNum, idxToCol(rule.colIdx), rule.header, err)
			}
		}

		*out = append(*out, rv.Interface().(T))

		// Return instance to pool for reuse
		e.instancePool.Put(rv)

		if e.opt.Limit > 0 && numberData >= e.opt.Limit {
			break
		}
	}

	e.IsNext = e.rows.Next()

	return e.rows.Error()
}

func buildRules(rt reflect.Type, colIndexByHeader map[string]int) ([]fieldRule, error) {
	// Pre-allocate rules slice with estimated capacity
	rules := make([]fieldRule, 0, rt.NumField())

	for i := 0; i < rt.NumField(); i++ {
		sf := rt.Field(i)
		if sf.Anonymous {
			// flatten embedded struct
			sub, err := buildRules(sf.Type, colIndexByHeader)
			if err != nil {
				return nil, err
			}
			for _, r := range sub {
				r.fieldIdx = i // not correct for embedded; do deep set instead if needed
			}
			rules = append(rules, sub...)
			continue
		}

		tag := sf.Tag.Get("excel")
		if tag == "" || tag == "-" {
			continue
		}
		spec := parseTag(tag)
		var colIdx int
		var ok bool
		switch {
		case spec.fixedCol != "":
			colIdx, ok = colToIdx(spec.fixedCol)
			if !ok {
				return nil, fmt.Errorf("invalid column letter %q on field %s", spec.fixedCol, sf.Name)
			}
		case spec.header != "":
			colIdx, ok = colIndexByHeader[spec.header]
			if !ok {
				return nil, fmt.Errorf("header %q not found for field %s", spec.header, sf.Name)
			}
		default:
			return nil, fmt.Errorf("excel tag on field %s must specify header or col", sf.Name)
		}

		// Cache type information for faster field setting
		fieldType := sf.Type
		isPointer := fieldType.Kind() == reflect.Pointer
		if isPointer {
			fieldType = fieldType.Elem()
		}

		rules = append(rules, fieldRule{
			fieldIdx:  i,
			colIdx:    colIdx,
			header:    spec.header,
			required:  spec.required,
			layout:    spec.layout,
			fieldType: fieldType,
			isPointer: isPointer,
			kindCache: fieldType.Kind(),
		})
	}
	return rules, nil
}

type tagSpec struct {
	header   string // match header text
	fixedCol string // e.g. "A", "BC"
	required bool
	layout   string
}

func parseTag(s string) tagSpec {
	parts := strings.Split(s, ",")
	ts := tagSpec{}
	if len(parts) > 0 {
		p := strings.TrimSpace(parts[0])
		if strings.HasPrefix(strings.ToLower(p), "col=") {
			ts.fixedCol = strings.TrimSpace(p[4:])
		} else if p != "" {
			ts.header = p
		}
	}
	for _, opt := range parts[1:] {
		opt = strings.TrimSpace(opt)
		switch {
		case opt == "required":
			ts.required = true
		case strings.HasPrefix(opt, "layout="):
			ts.layout = strings.TrimPrefix(opt, "layout=")
		case strings.HasPrefix(strings.ToLower(opt), "col="):
			ts.fixedCol = strings.TrimSpace(opt[4:])
		}
	}
	return ts
}

// setFieldValueOptimized uses cached type information for faster field setting
func setFieldValueOptimized(fv reflect.Value, raw string, rule *fieldRule) (err error) {
	if rule == nil {
		return fmt.Errorf("field rule cannot be nil")
	}

	// Protect against panics in reflection operations
	defer func() {
		if r := recover(); r != nil {
			err = fmt.Errorf("panic in field value setting: %v", r)
		}
	}()

	if !fv.IsValid() {
		return fmt.Errorf("invalid reflect value")
	}

	if rule.isPointer {
		// Handle pointer types
		if rule.fieldType == nil {
			return fmt.Errorf("field type cannot be nil for pointer type")
		}
		elem := reflect.New(rule.fieldType)
		if err := setFieldValueOptimizedDirect(elem.Elem(), raw, rule.kindCache, rule.layout); err != nil {
			return fmt.Errorf("failed to set pointer field value: %w", err)
		}
		if !fv.CanSet() {
			return fmt.Errorf("cannot set field value")
		}
		fv.Set(elem)
		return nil
	}

	return setFieldValueOptimizedDirect(fv, raw, rule.kindCache, rule.layout)
}

// setFieldValueOptimizedDirect with improved error handling
//func setFieldValueOptimizedDirect(fv reflect.Value, raw string, kind reflect.Kind, layout string) (err error) {
//	// Protect against panics
//	defer func() {
//		if r := recover(); r != nil {
//			err = fmt.Errorf("panic in direct field value setting: %v", r)
//		}
//	}()
//
//	if !fv.IsValid() || !fv.CanSet() {
//		return fmt.Errorf("invalid or unsettable field")
//	}
//
//	// Validate kind
//	if fv.Kind() != kind {
//		return fmt.Errorf("kind mismatch: expected %v, got %v", kind, fv.Kind())
//	}
//
//	// Clean input string
//	raw = strings.TrimSpace(raw)
//	if raw == "" && kind != reflect.String {
//		return nil // Skip empty values for non-string fields
//	}
//
//	switch kind {
//	case reflect.String:
//		fv.SetString(raw)
//	case reflect.Bool:
//		// Optimized boolean parsing
//		switch strings.ToLower(raw) {
//		case "1", "true", "yes":
//			fv.SetBool(true)
//		case "0", "false", "no":
//			fv.SetBool(false)
//		default:
//			return fmt.Errorf("invalid boolean value: %s", raw)
//		}
//	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
//		i, err := strconv.ParseInt(cleanNumOptimized(raw), 10, 64)
//		if err != nil {
//			return fmt.Errorf("failed to parse int: %w", err)
//		}
//		if !fv.OverflowInt(i) {
//			fv.SetInt(i)
//		} else {
//			return fmt.Errorf("integer overflow for value: %d", i)
//		}
//	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
//		u, err := strconv.ParseUint(cleanNumOptimized(raw), 10, 64)
//		if err != nil {
//			return fmt.Errorf("failed to parse uint: %w", err)
//		}
//		if !fv.OverflowUint(u) {
//			fv.SetUint(u)
//		} else {
//			return fmt.Errorf("unsigned integer overflow for value: %d", u)
//		}
//	case reflect.Float32, reflect.Float64:
//		fl, err := strconv.ParseFloat(strings.ReplaceAll(raw, ",", "."), 64)
//		if err != nil {
//			return fmt.Errorf("failed to parse float: %w", err)
//		}
//		if !fv.OverflowFloat(fl) {
//			fv.SetFloat(fl)
//		} else {
//			return fmt.Errorf("float overflow for value: %f", fl)
//		}
//	case reflect.Struct:
//		if fv.Type() == reflect.TypeOf(time.Time{}) {
//			t, err := parseAnyTimeOptimized(raw, layout)
//			if err != nil {
//				return fmt.Errorf("failed to parse time: %w", err)
//			}
//			fv.Set(reflect.ValueOf(t))
//			return nil
//		}
//		return fmt.Errorf("unsupported struct type: %s", fv.Type())
//	default:
//		return fmt.Errorf("unsupported kind: %s", kind)
//	}
//
//	return nil
//}

// setFieldValueOptimizedDirect sets field value directly using cached kind
func setFieldValueOptimizedDirect(fv reflect.Value, raw string, kind reflect.Kind, layout string) error {
	switch kind {
	case reflect.String:
		fv.SetString(raw)
	case reflect.Bool:
		// Optimized boolean parsing
		switch raw {
		case "1", "true", "TRUE", "True", "yes", "YES", "Yes":
			fv.SetBool(true)
		case "0", "false", "FALSE", "False", "no", "NO", "No":
			fv.SetBool(false)
		default:
			b, err := strconv.ParseBool(strings.ToLower(raw))
			if err != nil {
				return err
			}
			fv.SetBool(b)
		}
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		i, err := strconv.ParseInt(cleanNumOptimized(raw), 10, 64)
		if err != nil {
			return err
		}
		fv.SetInt(i)
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		u, err := strconv.ParseUint(cleanNumOptimized(raw), 10, 64)
		if err != nil {
			return err
		}
		fv.SetUint(u)
	case reflect.Float32, reflect.Float64:
		// Optimized float parsing
		cleanRaw := raw
		if strings.Contains(raw, ",") {
			cleanRaw = strings.ReplaceAll(raw, ",", ".")
		}
		fl, err := strconv.ParseFloat(cleanRaw, 64)
		if err != nil {
			return err
		}
		fv.SetFloat(fl)
	case reflect.Struct:
		if fv.Type() == reflect.TypeOf(time.Time{}) {
			t, err := parseAnyTimeOptimized(raw, layout)
			if err != nil {
				return err
			}
			fv.Set(reflect.ValueOf(t))
			return nil
		}
		return fmt.Errorf("unsupported struct type: %s", fv.Type())
	default:
		return fmt.Errorf("unsupported kind: %s", kind)
	}
	return nil
}

// Keep original function for backward compatibility
func setFieldValue(fv reflect.Value, raw string, layout string) error {
	switch fv.Kind() {
	case reflect.String:
		fv.SetString(raw)
	case reflect.Bool:
		if raw == "1" {
			fv.SetBool(true)
			return nil
		}
		if raw == "0" {
			fv.SetBool(false)
			return nil
		}
		b, err := strconv.ParseBool(strings.ToLower(raw))
		if err != nil {
			return err
		}
		fv.SetBool(b)
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		i, err := strconv.ParseInt(cleanNum(raw), 10, 64)
		if err != nil {
			return err
		}
		fv.SetInt(i)
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		u, err := strconv.ParseUint(cleanNum(raw), 10, 64)
		if err != nil {
			return err
		}
		fv.SetUint(u)
	case reflect.Float32, reflect.Float64:
		fl, err := strconv.ParseFloat(strings.ReplaceAll(raw, ",", "."), 64)
		if err != nil {
			return err
		}
		fv.SetFloat(fl)
	case reflect.Struct:
		if fv.Type() == reflect.TypeOf(time.Time{}) {
			t, err := parseAnyTime(raw, layout)
			if err != nil {
				return err
			}
			fv.Set(reflect.ValueOf(t))
			return nil
		}
		return fmt.Errorf("unsupported struct type: %s", fv.Type())
	case reflect.Pointer:
		elem := reflect.New(fv.Type().Elem())
		if err := setFieldValue(elem.Elem(), raw, layout); err != nil {
			return err
		}
		fv.Set(elem)
	default:
		return fmt.Errorf("unsupported kind: %s", fv.Kind())
	}
	return nil
}

func parseAnyTime(s, layout string) (time.Time, error) {
	if layout != "" {
		if t, err := time.Parse(layout, s); err == nil {
			return t, nil
		}
	}
	for _, l := range defaultTimeLayouts {
		if t, err := time.Parse(l, s); err == nil {
			return t, nil
		}
	}
	// try excel serial number (e.g., "45123.0")
	if f, err := strconv.ParseFloat(s, 64); err == nil {
		// Excel serial date: days since 1899-12-30 (Excel's epoch, accounting for the 1900 leap year bug).
		base := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
		sec := int64(f * 86400.0)
		return base.Add(time.Duration(sec) * time.Second), nil
	}
	return time.Time{}, fmt.Errorf("cannot parse time %q", s)
}

// cleanNumOptimized provides faster number cleaning with fewer allocations
func cleanNumOptimized(s string) string {
	if s == "" {
		return s
	}

	// Quick check if cleaning is needed
	needsCleaning := false
	for i := 0; i < len(s); i++ {
		c := s[i]
		if c == ' ' || c == '\t' || c == '\n' || c == '\r' || c == ',' {
			needsCleaning = true
			break
		}
	}

	if !needsCleaning {
		return s
	}

	// Use builder for efficient string construction
	var builder strings.Builder
	builder.Grow(len(s)) // Pre-allocate capacity

	for i := 0; i < len(s); i++ {
		c := s[i]
		if c != ' ' && c != '\t' && c != '\n' && c != '\r' && c != ',' {
			builder.WriteByte(c)
		}
	}

	return builder.String()
}

// Keep original function for backward compatibility
func cleanNum(s string) string {
	s = strings.TrimSpace(s)
	s = strings.ReplaceAll(s, ",", "") // 1,234 -> 1234
	return s
}

// parseAnyTimeOptimized provides faster time parsing with caching
func parseAnyTimeOptimized(s, layout string) (time.Time, error) {
	// Try custom layout first if provided
	if layout != "" {
		if t, err := time.Parse(layout, s); err == nil {
			return t, nil
		}
	}

	// Try common formats first for better performance
	commonLayouts := []string{
		"2006-01-02", "02/01/2006", "2006-01-02 15:04:05",
	}

	for _, l := range commonLayouts {
		if t, err := time.Parse(l, s); err == nil {
			return t, nil
		}
	}

	// Fallback to all default layouts
	for _, l := range defaultTimeLayouts {
		if t, err := time.Parse(l, s); err == nil {
			return t, nil
		}
	}

	// Try excel serial number (e.g., "45123.0")
	if f, err := strconv.ParseFloat(s, 64); err == nil {
		// Excel serial date: days since 1899-12-30
		base := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
		sec := int64(f * 86400.0)
		return base.Add(time.Duration(sec) * time.Second), nil
	}

	return time.Time{}, fmt.Errorf("cannot parse time %q", s)
}

func colToIdx(col string) (int, bool) {
	col = strings.ToUpper(strings.TrimSpace(col))
	if col == "" {
		return 0, false
	}
	sum := 0
	for _, r := range col {
		if r < 'A' || r > 'Z' {
			return 0, false
		}
		sum = sum*26 + int(r-'A'+1)
	}
	return sum - 1, true
}

func idxToCol(idx int) string {
	if idx < 0 {
		return ""
	}
	var out []rune
	for idx >= 0 {
		r := rune(idx%26) + 'A'
		out = append([]rune{r}, out...)
		idx = idx/26 - 1
	}
	return string(out)
}

// helper: validate header rune-safety (optional)
func isPrintableASCII(s string) bool {
	for len(s) > 0 {
		r, size := utf8.DecodeRuneInString(s)
		if r == utf8.RuneError && size == 1 {
			return false
		}
		if r < 0x20 && r != '\t' {
			return false
		}
		s = s[size:]
	}
	return true
}
