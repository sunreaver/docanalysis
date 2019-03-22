package docanalysis

import (
	"fmt"
	"path"
	"strings"

	"baliance.com/gooxml/common"
	"baliance.com/gooxml/spreadsheet"

	"baliance.com/gooxml/document"
	"baliance.com/gooxml/presentation"
)

// Document 文档
type Document struct {
	File string
}

// Image Image
type Image struct {
	Path string
	Ex   string
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
	doc, e := document.Open(d.File)
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
	xls, e := spreadsheet.Open(d.File)
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
	ppt, e := presentation.Open(d.File)
	if e != nil {
		return nil, "", e
	}

	// 图像
	images = getImages(ppt.Images)

	// Slides
	// fmt.Println("Slides", len(ppt.Slides()))
	// for _, slide := range ppt.Slides() {
	// 	for _, ph := range slide.PlaceHolders() {
	// 		for _, para := range ph.Paragraphs() {
	// 			fmt.Println(para.Properties())
	// 		}
	// 	}
	// }

	// // SlideLayouts
	// fmt.Println("SlideLayouts", len(ppt.SlideLayouts()))
	// for _, slide := range ppt.SlideLayouts() {
	// 	fmt.Println(slide.Name())
	// }

	return
}

// Analysis 执行解析
func (d *Document) Analysis() (images []*Image, text string, err error) {
	if d == nil || len(d.File) == 0 {
		return nil, "", ErrNoFile
	}
	if strings.HasSuffix(d.File, ".doc") ||
		strings.HasSuffix(d.File, ".docx") {
		return d.docx()
	} else if strings.HasSuffix(d.File, ".ppt") ||
		strings.HasSuffix(d.File, ".pptx") {
		return d.pptx()
	} else if strings.HasSuffix(d.File, "xls") ||
		strings.HasSuffix(d.File, "xlsx") {
		return d.xlsx()
	}

	return nil, "", ErrNoSupport
}
