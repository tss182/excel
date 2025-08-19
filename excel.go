package excel

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"mime/multipart"
	"reflect"
	"strconv"
	"strings"
	"time"
	"unicode/utf8"
)

func defaultOptions() Opt {
	return Opt{
		HeaderRow:    1,
		DataStartRow: 2,
		Limit:        100,
	}
}

func New[T any]() *Excel[T] {
	return &Excel[T]{
		file: excelize.NewFile(),
	}
}

func Open[T any](filename string) (*Excel[T], error) {
	file, err := excelize.OpenFile(filename)
	if err != nil {
		return nil, fmt.Errorf("failed to open file %s: %w", filename, err)
	}
	return &Excel[T]{file: file}, nil
}

func OpenReader[T any](file multipart.File) (*Excel[T], error) {
	f, err := excelize.OpenReader(file)
	if err != nil {
		return nil, fmt.Errorf("failed to open reader: %w", err)
	}
	return &Excel[T]{file: f}, nil
}

func (e *Excel[T]) Close() error {
	return e.file.Close()
}

func (e *Excel[T]) Read(out *[]T, sheetName string, opts ...Opt) error {
	e.opt = defaultOptions()
	if opts != nil {
		for _, opt := range opts {
			if opt.HeaderRow > 0 {
				e.opt.HeaderRow = opt.HeaderRow
			}
			if opt.DataStartRow > 0 {
				e.opt.DataStartRow = opt.DataStartRow
			}
			if opt.Limit > 0 {
				e.opt.Limit = opt.Limit
			}
		}
	}

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

	for i := uint8(0); i < e.opt.HeaderRow; i++ {
		if !e.rows.Next() {
			return fmt.Errorf("sheet %s in file excel: %w", sheetName, errors.New("empty sheet"))
		}
		e.activeRow = uint(i) + 1
	}

	fmt.Println("activeRow header:", e.activeRow)
	header, err := e.rows.Columns()
	if err != nil {
		return fmt.Errorf("failed to read header: %w", err)
	}

	fmt.Println("header:", header)

	colIdxHeader := make(map[string]int, len(header))
	for i, h := range header {
		colIdxHeader[strings.TrimSpace(h)] = i
	}

	var zero T
	rt := reflect.TypeOf(zero)
	if rt.Kind() == reflect.Pointer {
		rt = rt.Elem()
	}
	if rt.Kind() != reflect.Struct {
		return errors.New("out must be a struct or *struct")
	}

	e.rt = rt

	e.rules, err = buildRules(rt, colIdxHeader)
	if err != nil {
		return err
	}

	e.rows.Next()

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
	var numberData uint = 0
	for {
		e.activeRow++
		if e.activeRow < uint(e.opt.DataStartRow) {
			e.rows.Next()
			continue
		}
		numberData++
		cols, err := e.rows.Columns()
		if err != nil {
			return fmt.Errorf("read row %d: %w", e.activeRow, err)
		}

		// new instance of T
		rv := reflect.New(e.rt).Elem()

		for _, rule := range e.rules {
			var cell string
			if rule.colIdx < len(cols) {
				cell = strings.TrimSpace(cols[rule.colIdx])
			}

			if rule.required && cell == "" {
				return fmt.Errorf("row %d col %s (%s) is required", e.activeRow, idxToCol(rule.colIdx), rule.header)
			}
			if cell == "" {
				continue
			}

			fv := rv.Field(rule.fieldIdx)
			if !fv.CanSet() {
				continue
			}
			if err := setFieldValue(fv, cell, rule.layout); err != nil {
				return fmt.Errorf("row %d col %s (%s): %w", e.activeRow, idxToCol(rule.colIdx), rule.header, err)
			}
		}
		*out = append(*out, rv.Interface().(T))
		if (e.opt.Limit > 0 && numberData >= e.opt.Limit) || !e.rows.Next() {
			break
		}
	}

	e.IsNext = e.rows.Next()

	return e.rows.Error()
}

func buildRules(rt reflect.Type, colIndexByHeader map[string]int) ([]fieldRule, error) {
	var rules []fieldRule
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
		rules = append(rules, fieldRule{
			fieldIdx: i,
			colIdx:   colIdx,
			header:   spec.header,
			required: spec.required,
			layout:   spec.layout,
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
		// int64 also covers time.Duration if needed (could be extended).
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
		// allocate and set
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

func cleanNum(s string) string {
	s = strings.TrimSpace(s)
	s = strings.ReplaceAll(s, ",", "") // 1,234 -> 1234
	return s
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
