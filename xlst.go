package xlst

import (
	"bytes"
	"errors"
	"fmt"
	"html"
	"io"
	"reflect"
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/aymerick/raymond"
	xls "github.com/xuri/excelize/v2"
)

func init() {
	raymond.RegisterHelper("inc", func(value string) interface{} {
		intValue, err := strconv.Atoi(value)
		if err != nil {
			fmt.Printf("failed to inc %q: %v\n", value, err)
		}
		return strconv.Itoa(intValue + 1)
	})
}

var (
	listRgx = regexp.MustCompile(`\{\{\s*(\w+)\.\w+\s*\}\}`)
)

// Xlst Represents template struct
type Xlst struct {
	file *xls.File
}

// New creates new Xlst struct and returns pointer to it
func New() *Xlst {
	return &Xlst{}
}

// NewFromBinary creates new Xlst struct puts binary tempate into and returns pointer to it
func NewFromBinary(content []byte) (*Xlst, error) {
	file, err := xls.OpenReader(bytes.NewReader(content))
	if err != nil {
		return nil, err
	}

	res := &Xlst{file: file}
	return res, nil
}

// Render renders report and stores it in a struct
func (m *Xlst) Render(in interface{}) error {
	return m.RenderWithOptions(in)
}

func (m *Xlst) RenderWithOptions(in interface{}, opts ...Option) error {
	ctx := getCtx(in)
	for _, sheet := range m.file.GetSheetList() {
		if err := m.renderRows(ctx, sheet, opts...); err != nil {
			return fmt.Errorf("renderRows(%s): %w", sheet, err)
		}
	}
	return nil
}

// ReadTemplate reads template from disk and stores it in a struct
func (m *Xlst) ReadTemplate(path string) error {
	file, err := xls.OpenFile(path)
	if err != nil {
		return fmt.Errorf("xls.OpenFile: %w", err)
	}
	m.file = file
	return nil
}

// Save saves generated report to disk
func (m *Xlst) Save(path string) error {
	if m.file == nil {
		return errors.New("report was not generated")
	}
	return m.file.SaveAs(path)
}

// Write writes generated report to provided writer
func (m *Xlst) Write(writer io.Writer) error {
	if m.file == nil {
		return errors.New("report was not generated")
	}
	return m.file.Write(writer)
}

func (m *Xlst) renderRows(ctx map[string]interface{}, sheet string, opts ...Option) error {
	rows, err := m.file.Rows(sheet)
	if err != nil {
		return fmt.Errorf("file.Rows: %w", err)
	}

	modifications, err := collectModifications(ctx, rows, opts...)
	if err != nil {
		_ = rows.Close()
		return fmt.Errorf("collectModifications: %w", err)
	}
	_ = rows.Close()

	if err = m.modifyCells(sheet, modifications.CellModifications); err != nil {
		return fmt.Errorf("modifyCells: %w", err)
	}

	if err = m.insertRows(sheet, modifications.RowInsertions); err != nil {
		return fmt.Errorf("insertRows: %w", err)
	}

	return nil
}

func (m *Xlst) modifyCells(sheet string, cellModifications []*CellModification) error {
	for _, cm := range cellModifications {
		if err := m.modifyCell(sheet, cm.Column, cm.Row, cm.Value); err != nil {
			return fmt.Errorf("modifyCell: %w", err)
		}
	}
	return nil
}

func (m *Xlst) insertRows(sheet string, rowInsertions map[int][]*RowInsertion) error {
	if len(rowInsertions) == 0 {
		return nil
	}

	keys := make([]int, 0, len(rowInsertions))
	for key := range rowInsertions {
		keys = append(keys, key)
	}

	if len(keys) > 1 {
		sort.Slice(keys, func(i, j int) bool {
			return keys[i] < keys[j]
		})
	}

	for _, row := range keys {
		list := rowInsertions[row]

		if len(list) == 0 {
			if err := m.file.RemoveRow(sheet, row); err != nil {
				return fmt.Errorf("file.RemoveRow(%d): %w", row, err)
			}
			continue
		}
		for i := 1; i < len(list); i++ {
			if err := m.file.DuplicateRow(sheet, row); err != nil {
				return fmt.Errorf("file.DuplicateRow(%d): %w", row, err)
			}
		}

		for i, ri := range list {
			for _, cm := range ri.Columns {
				if err := m.modifyCell(sheet, cm.Column, row+i, cm.Value); err != nil {
					return fmt.Errorf("modifyCell: %w", err)
				}
			}
		}
	}

	return nil
}

func (m *Xlst) modifyCell(sheet string, column, row int, value string) error {
	cellName, err := xls.CoordinatesToCellName(column, row)
	if err != nil {
		return fmt.Errorf("xls.CoordinatesToCellName(%d, %d): %w", column, row, err)
	}
	if err = m.file.SetCellStr(sheet, cellName, value); err != nil {
		return fmt.Errorf("SetCellStr(%s, %s): %w", cellName, value, err)
	}
	return nil
}

func collectModifications(ctx map[string]interface{}, rows *xls.Rows, opts ...Option) (*Modifications, error) {
	row := 0
	result := NewModifications()

	var options options
	for _, opt := range opts {
		opt(&options)
	}

	for rows.Next() {
		row++

		columns, err := rows.Columns()
		if err != nil {
			return nil, fmt.Errorf("rows.Columns: %w", err)
		}

		listProperty := getListProp(columns)
		if listProperty == "" || !isArray(ctx, listProperty) {
			for i, column := range columns {
				if column == "" {
					continue
				}
				if !strings.Contains(column, "{{") || !strings.Contains(column, "}}") {
					continue
				}

				value, err := cellModification(column, ctx, &options)
				if err != nil {
					return nil, fmt.Errorf("cellModification (%d, %d): %w", row, i+1, err)
				}
				result.AddCellModification(&CellModification{
					Row: row,
					ColumnModification: ColumnModification{
						Column: i + 1,
						Value:  value,
					},
				})
			}
			continue
		}

		modifiedRow := row + result.Offset

		arr := reflect.ValueOf(ctx[listProperty])
		arrBackup := ctx[listProperty]
		idxBackup := ctx["Index"]
		for i := 0; i < arr.Len(); i++ {
			ctx[listProperty] = arr.Index(i).Interface()
			if idxBackup == nil {
				ctx["Index"] = strconv.Itoa(i)
			}

			rowInsertion := &RowInsertion{}

			for columnIndex, column := range columns {
				if column == "" {
					continue
				}
				if !strings.Contains(column, "{{") || !strings.Contains(column, "}}") {
					continue
				}

				value, err := cellModification(column, ctx, &options)
				if err != nil {
					return nil, fmt.Errorf("cellModification (%d, %d): %w", modifiedRow, columnIndex+1, err)
				}
				rowInsertion.Columns = append(rowInsertion.Columns, &ColumnModification{
					Column: columnIndex + 1,
					Value:  value,
				})
			}
			result.AddRowInsertion(modifiedRow, rowInsertion)
		}
		if arr.Len() == 0 {
			result.AddEmptyRowInsertion(modifiedRow)
		}
		result.Offset += arr.Len() - 1

		ctx[listProperty] = arrBackup
		ctx["Index"] = idxBackup
	}

	return result, nil
}

func cellModification(value string, ctx interface{}, opts *options) (string, error) {
	if value == "" {
		return value, nil
	}
	if !strings.Contains(value, "{{") || !strings.Contains(value, "}}") {
		return value, nil
	}

	template, err := raymond.Parse(value)
	if err != nil {
		return "", fmt.Errorf("raymond.Parse: %w", err)
	}
	out, err := template.Exec(ctx)
	if err != nil {
		return "", fmt.Errorf("template.Exec: %w", err)
	}
	if opts.unescapeHTML {
		out = html.UnescapeString(out)
	}
	return out, nil
}

func getCtx(in interface{}) map[string]interface{} {
	if ctx, ok := in.(map[string]interface{}); ok {
		return ctx
	}
	return nil
}

func isArray(in map[string]interface{}, prop string) bool {
	val, ok := in[prop]
	if !ok {
		return false
	}
	switch reflect.TypeOf(val).Kind() {
	case reflect.Array, reflect.Slice:
		return true
	}
	return false
}

func getListProp(columns []string) string {
	for _, column := range columns {
		if match := listRgx.FindAllStringSubmatch(column, -1); match != nil {
			return match[0][1]
		}
	}
	return ""
}
