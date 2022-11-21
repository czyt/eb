package eb

type WriterOpt func(option *WriterOption)

type WriterOption struct {
	SheetName string // SheetName: default sheet name
	TagName   string // TagName: tag name
}

func WithWriterSheet(sheetName string) WriterOpt {
	return func(option *WriterOption) {
		option.SheetName = sheetName
	}
}

func WithWriterTag(tagName string) WriterOpt {
	return func(option *WriterOption) {
		option.TagName = tagName
	}
}

func NewWriterOption(options ...WriterOpt) *WriterOption {
	writerOption := &WriterOption{}
	if len(options) == 0 {
		return defaultWriterOptions()
	}
	for _, apply := range options {
		apply(writerOption)
	}
	return writerOption
}

func defaultWriterOptions() *WriterOption {
	return NewWriterOption(WithWriterSheet("sheet1"), WithWriterTag("excel"))
}
