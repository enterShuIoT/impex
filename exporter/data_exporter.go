package exporter

type DownloadResponse struct {
	FileName    string
	FileSize    int64
	ContentType string
	Content     []byte
}

type DataExporter interface {
	Export(data any) (*DownloadResponse, error)
}
