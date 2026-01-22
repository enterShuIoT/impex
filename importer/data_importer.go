package importer

type DataImporter[T any] interface {
	Import(path string) ([]T, error)
}
