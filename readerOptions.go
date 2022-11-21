package eb

type ReaderOpt func(option *ReaderOption)

type ReaderOption struct {
	TagName           string // TagName: tag name
	SheetIndex        int    // SheetIndex: read sheet index
	HeaderRowIndex    int    // HeaderRowIndex: sheet header row index
	DataStartRowIndex int    // DataStartRowIndex: sheet data start row index
	TrimSpace         bool   // TrimSpace: trim space left and right only on `string` type
}

func WithReaderTag(tagName string) ReaderOpt {
	return func(option *ReaderOption) {
		option.TagName = tagName
	}
}

func WithSheetIndex(sheetIndex int) ReaderOpt {
	return func(option *ReaderOption) {
		option.SheetIndex = sheetIndex
	}
}

func WithHeaderRowIndex(headerRowIndex int) ReaderOpt {
	return func(option *ReaderOption) {
		option.HeaderRowIndex = headerRowIndex
	}
}

func WithDataStartRowIndex(dataStartRowIndex int) ReaderOpt {
	return func(option *ReaderOption) {
		option.DataStartRowIndex = dataStartRowIndex
	}
}

func WithTrimSpace(trimSpace bool) ReaderOpt {
	return func(option *ReaderOption) {
		option.TrimSpace = trimSpace
	}
}

func NewReaderOption(options ...ReaderOpt) *ReaderOption {
	readerOpt := &ReaderOption{}
	if len(options) == 0 {
		return defaultReaderOptions()
	}
	for _, apply := range options {
		apply(readerOpt)
	}
	return readerOpt
}

func defaultReaderOptions() *ReaderOption {
	return NewReaderOption(WithReaderTag("excel"), WithDataStartRowIndex(1))
}
