package xlst

type ColumnModification struct {
	Column int
	Value  string
}

type CellModification struct {
	Row int
	ColumnModification
}

type RowInsertion struct {
	Columns []*ColumnModification
}

type Modifications struct {
	CellModifications []*CellModification
	RowInsertions     map[int][]*RowInsertion
	Offset            int
}

func NewModifications() *Modifications {
	return &Modifications{RowInsertions: make(map[int][]*RowInsertion)}
}

func (m *Modifications) AddCellModification(cm *CellModification) {
	m.CellModifications = append(m.CellModifications, cm)
}

func (m *Modifications) AddRowInsertion(row int, ri *RowInsertion) {
	m.RowInsertions[row] = append(m.RowInsertions[row], ri)
}

func (m *Modifications) AddEmptyRowInsertion(row int) {
	m.RowInsertions[row] = []*RowInsertion{}
}
