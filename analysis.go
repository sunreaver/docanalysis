package docanalysis

import (
	"bytes"
	"fmt"
	"image/jpeg"
	"io"
	"path"
	"strings"

	"baliance.com/gooxml/common"
	"baliance.com/gooxml/spreadsheet"

	"baliance.com/gooxml/document"
	"baliance.com/gooxml/presentation"

	"github.com/unidoc/unidoc/pdf/extractor"
	pdf "github.com/unidoc/unidoc/pdf/model"
)

// Document 文档
type Document struct {
	File io.ReaderAt
	Name string
	Size int64
}

// Image Image
type Image struct {
	Path string
	Ex   string
	Body []byte
}

func (i *Image) String() string {
	return fmt.Sprintf("%s.%s", i.Path, i.Ex)
}

func paragraphs(pg []document.Paragraph) (text string) {
	for _, para := range pg {
		for _, run := range para.Runs() {
			text += run.Text()
		}
	}
	return
}

func getImages(imgs []common.ImageRef) []*Image {
	tmpImage := make([]*Image, len(imgs))
	for index, img := range imgs {
		tmpImage[index] = &Image{
			Path: img.Path(),
			Ex:   img.Format(),
		}
	}
	return tmpImage
}

func (d *Document) docx() (images []*Image, text string, err error) {
	doc, e := document.Read(d.File, d.Size)
	if e != nil {
		return nil, "", e
	}

	// 页脚
	for _, footer := range doc.Footers() {
		text += paragraphs(footer.Paragraphs())
	}

	// 页眉
	for _, header := range doc.Headers() {
		text += paragraphs(header.Paragraphs())
	}

	// 段落、正文
	text += paragraphs(doc.Paragraphs())

	// 图像
	images = getImages(doc.Images)

	// 图表
	for _, table := range doc.Tables() {
		for _, row := range table.Rows() {
			for _, cell := range row.Cells() {
				text += paragraphs(cell.Paragraphs())
			}
		}
	}
	return
}

func (d *Document) xlsx() (images []*Image, text string, err error) {
	xls, e := spreadsheet.Read(d.File, d.Size)
	if e != nil {
		return nil, "", e
	}

	// 图像
	for _, ef := range xls.ExtraFiles {
		ex := strings.ToLower(path.Ext(ef.ZipPath))
		if len(ex) > 1 {
			ex = ex[1:]
		}
		if ex != "jpg" && ex != "jpeg" && ex != "png" {
			continue
		}
		images = append(images, &Image{
			Path: ef.DiskPath,
			Ex:   ex,
		})
	}

	for _, sheet := range xls.Sheets() {
		for _, row := range sheet.Rows() {
			for _, cell := range row.Cells() {
				if cell.IsNumber() || cell.IsBool() {
					continue
				}
				if value := cell.GetString(); len(value) != 0 {
					text += value
				}
			}
		}
	}

	return
}

func (d *Document) pptx() (images []*Image, text string, err error) {
	ppt, e := presentation.Read(d.File, d.Size)
	if e != nil {
		return nil, "", e
	}

	// 图像
	images = getImages(ppt.Images)
	return
}

func (d *Document) pdf() (images []*Image, text string, err error) {
	readSeeker := io.NewSectionReader(d.File, 0, d.Size)
	pdfReader, err := pdf.NewPdfReader(readSeeker)
	if err != nil {
		return nil, "", err
	}

	// pdf 是否加密
	isEncrypted, err := pdfReader.IsEncrypted()
	if err != nil {
		return nil, "", fmt.Errorf("pdf is encrypted err: %v", err)
	} else if isEncrypted {
		return nil, "", fmt.Errorf("pdf is encrypted")
	}

	pages, err := pdfReader.GetNumPages()
	if err != nil {
		return nil, "", fmt.Errorf("pdf get page nums err: %v", err)
	}

	for pageNum := 1; pageNum <= pages; pageNum++ {
		page, e := pdfReader.GetPage(pageNum)
		if e != nil {
			return nil, "", fmt.Errorf("pdf get page err: %v", e)
		}

		ex, err := extractor.New(page)
		if err != nil {
			return nil, "", fmt.Errorf("pdf new ext err: %v", err)
		}

		tmpText, err := ex.ExtractText()
		if err != nil {
			return nil, "", fmt.Errorf("pdf ext text err: %v", err)
		}
		text += tmpText

		// 检测本页图片
		rgbImages, err := extractImagesOnPage(page)
		if err != nil {
			return nil, "", fmt.Errorf("pdf ext image err: %v", err)
		}

		for _, img := range rgbImages {
			gimg, err := img.ToGoImage()
			if err != nil {
				return nil, "", fmt.Errorf("pdf img2goimg err: %v", err)
			}
			buffer := bytes.NewBuffer([]byte{})
			opt := jpeg.Options{Quality: 100}
			err = jpeg.Encode(buffer, gimg, &opt)
			if err != nil {
				return nil, "", fmt.Errorf("pdf img to jpeg err: %v", err)
			}

			images = append(images, &Image{
				Path: "",
				Ex:   "jpeg",
				Body: buffer.Bytes(),
			})
		}
	}
	return
}

// Analysis 执行解析
func (d *Document) Analysis() (images []*Image, text string, err error) {
	if d == nil || d.File == nil || len(d.Name) == 0 {
		return nil, "", ErrNoFile
	}

	if strings.HasSuffix(d.Name, ".doc") ||
		strings.HasSuffix(d.Name, ".docx") {
		return d.docx()
	} else if strings.HasSuffix(d.Name, ".ppt") ||
		strings.HasSuffix(d.Name, ".pptx") {
		return d.pptx()
	} else if strings.HasSuffix(d.Name, "xls") ||
		strings.HasSuffix(d.Name, "xlsx") {
		return d.xlsx()
	} else if strings.HasSuffix(d.Name, "pdf") {
		return d.pdf()
	}

	return nil, "", ErrNoSupport
}
