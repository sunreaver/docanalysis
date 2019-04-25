package docanalysis

// Options 可选参数
type Options struct {
	// MaxImageCount 最多返回的图片量
	MaxImageCount int
	// ImageMinSize 返回图片的大小底限
	ImageMinSize int
	// ReadTextWithImageProportion 包含图片的页在总页数中的占比超过此值，则不返回文字
	ReadTextWithImageProportion float64
	// SkipPDFWithNumPages 跳过页数以下的pdf文件处理
	SkipPDFWithNumPages int

	// ExcelMaxRow excel max row
	ExcelMaxRow       int
	ExcelMaxCellInRow int
	ExcelMaxSheet     int
}

var defaultOption = Options{
	MaxImageCount:               10,
	ImageMinSize:                32 * 1024,
	ReadTextWithImageProportion: 0.7,
	SkipPDFWithNumPages:         0x01 << 16,
	ExcelMaxRow:                 1000,
	ExcelMaxCellInRow:           1000,
	ExcelMaxSheet:               20,
}

// Valid Valid
func (o *Options) Valid() {
	if o == nil {
		return
	}
	if o.ImageMinSize <= 0 {
		o.ImageMinSize = defaultOption.ImageMinSize
	}
	if o.ReadTextWithImageProportion < 0.0 {
		o.ReadTextWithImageProportion = 0.0
	}
	if o.MaxImageCount < 0 {
		o.MaxImageCount = defaultOption.MaxImageCount
	}
	if o.SkipPDFWithNumPages < 0 {
		o.SkipPDFWithNumPages = defaultOption.SkipPDFWithNumPages
	}
	if o.ExcelMaxRow <= 0 {
		o.ExcelMaxRow = defaultOption.ExcelMaxRow
	}
	if o.ExcelMaxCellInRow <= 0 {
		o.ExcelMaxCellInRow = defaultOption.ExcelMaxCellInRow
	}
	if o.ExcelMaxSheet <= 0 {
		o.ExcelMaxSheet = defaultOption.ExcelMaxSheet
	}
}
