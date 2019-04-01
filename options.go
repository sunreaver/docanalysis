package docanalysis

// Options 可选参数
type Options struct {
	// MaxImageCount 最多返回的图片量
	MaxImageCount int
	// ImageMinSize 返回图片的大小底限
	ImageMinSize int64
	// ReadTextWithImageProportion 包含图片的页在总页数中的占比超过此值，则不返回文字
	ReadTextWithImageProportion float64
	// SkipPDFWithNumPages 跳过页数以下的pdf文件处理
	SkipPDFWithNumPages int
}

var defaultOption = Options{
	MaxImageCount:               10,
	ImageMinSize:                32 * 1024,
	ReadTextWithImageProportion: 0.7,
	SkipPDFWithNumPages:         0x01 << 16,
}

// Valid Valid
func (o *Options) Valid() {
	if o == nil {
		return
	}
	if o.ImageMinSize < 0 {
		o.ImageMinSize = 0
	}
	if o.ReadTextWithImageProportion < 0.0 {
		o.ReadTextWithImageProportion = 0.0
	}
	if o.MaxImageCount <= 0 {
		o.MaxImageCount = 0
	}
	if o.SkipPDFWithNumPages < 0 {
		o.SkipPDFWithNumPages = defaultOption.SkipPDFWithNumPages
	}
}
