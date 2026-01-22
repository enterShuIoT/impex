package importer

type ImportResult[T any] struct {
	RowIndex int
	Data     T
	Error    error
}

type DataImporter[T any] interface {
	Import(path string) ([]T, error)
	ImportStream(path string) <-chan ImportResult[T]
}
