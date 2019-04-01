package docanalysis

import (
	"bytes"
	"fmt"
	"io"
	"io/ioutil"
	"os"
	"path"
	"strings"

	"code.sajari.com/docconv"

	"baliance.com/gooxml/common"
	"baliance.com/gooxml/spreadsheet"

	"baliance.com/gooxml/document"
	"baliance.com/gooxml/presentation"

	"github.com/extrame/xls"
	"github.com/pixiv/go-libjpeg/jpeg"
	pdf "github.com/unidoc/unidoc/pdf/model"
)

// Document 文档
type Document struct {
	File *bytes.Reader
	Name string
	Size int64
}

// Image Image
type Image struct {
	Path string
	Ex   string
	Body *bytes.Buffer
}

func (i *Image) String() string {
	return fmt.Sprintf("%s.%s", i.Path, i.Ex)
}

func getImages(imgs []common.ImageRef, opt *Options) []*Image {
	var tmpImage []*Image
	for _, img := range imgs {
		if len(tmpImage) >= opt.MaxImageCount {
			break
		}

		stat, e := os.Stat(img.Path())
		if e != nil {
			continue
		}
		if stat.Size() < int64(opt.ImageMinSize) {
			continue
		}

		tmpImage = append(tmpImage, &Image{
			Path: img.Path(),
			Ex:   img.Format(),
		})
	}
	return tmpImage
}

func (d *Document) text() (string, error) {
	mime := docconv.MimeTypeByExtension(d.Name)
	resp, e := docconv.Convert(d.File, mime, true)
	if e != nil {
		return "", fmt.Errorf("convert %v error: %v", mime, e)
	}
	return resp.Body, nil
}

func (d *Document) doc(_ *Options) (images []*Image, text string, err error) {
	text, err = d.text()
	if err != nil {
		return nil, "", err
	}

	return nil, text, err
}

func (d *Document) docx(opt *Options) (images []*Image, text string, err error) {
	text, err = d.text()
	if err != nil {
		return nil, "", err
	}

	doc, e := document.Read(d.File, d.Size)
	if e == nil {
		// 图像
		images = getImages(doc.Images, opt)
	}
	return
}

func (d *Document) xls(_ *Options) (images []*Image, text string, err error) {
	xlFile, e := xls.OpenReader(d.File, "utf-8")
	if e != nil {
		return nil, "", fmt.Errorf("read xls err: %v", e)
	}

	var body strings.Builder
	for i := 0; i < xlFile.NumSheets(); i++ {
		sheet := xlFile.GetSheet(i)
		body.WriteString(sheet.Name)

		for j := 0; j < int(sheet.MaxRow); j++ {
			if sheet.Row(j) == nil {
				continue
			}
			row := sheet.Row(j)
			for index := row.FirstCol(); index < row.LastCol(); index++ {
				if row.Col(index) != "" {
					body.WriteString(row.Col(index))
				}
			}
		}
	}
	return nil, body.String(), nil
}

func (d *Document) xlsx(_ *Options) (images []*Image, text string, err error) {
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

func (d *Document) pptx(opt *Options) (images []*Image, text string, err error) {
	ppt, e := presentation.Read(d.File, d.Size)
	if e != nil {
		return nil, "", e
	}

	// 图像
	images = getImages(ppt.Images, opt)
	return
}

func (d *Document) pdf(opt *Options) (images []*Image, text string, err error) {
	// get img
	readSeeker := io.NewSectionReader(d.File, 0, d.Size)
	pdfReader, e := pdf.NewPdfReader(readSeeker)
	if e != nil {
		return nil, "", fmt.Errorf("read pdf err: %v", e)
	}

	// pdf 是否加密
	isEncrypted, err := pdfReader.IsEncrypted()
	if err != nil {
		return nil, "", fmt.Errorf("check pdf is_encrypted err: %v", e)
	} else if isEncrypted {
		return nil, "", fmt.Errorf("pdf is_encrypted")
	}

	pages, err := pdfReader.GetNumPages()
	if err != nil {
		return nil, "", fmt.Errorf("get pdf page_num err: %v", pages)
	} else if pages <= opt.SkipPDFWithNumPages {
		return nil, "", fmt.Errorf("pdf pages too less: %v/%v", pages, opt.SkipPDFWithNumPages)
	}

	imgPagesNum := 0
	maxImagePageCount := opt.ReadTextWithImageProportion * float64(pages)
	needReadText := true

	for pageNum := int(opt.SkipPDFWithNumPages) + 1; pageNum <= pages; pageNum++ {
		if !needReadText && len(images) >= opt.MaxImageCount {
			// 图片已满，且已经检测出不需要text处理
			// 则直接跳出图片处理逻辑
			// 因为接下来的所有图片都是无用图片，既不会输出，也不会影响text输出
			break
		}

		page, e := pdfReader.GetPage(pageNum)
		if e != nil {
			continue
		}

		// 检测本页图片
		rgbImages, err := extractImagesOnPage(page)
		if err != nil {
			// 如果图片解析错误，直接跳过
			continue
		}
		hadImg := false

		for _, img := range rgbImages {
			gimg, err := img.ToGoImage()
			if err != nil {
				continue
			}
			buffer := bytes.NewBuffer([]byte{})
			err = jpeg.Encode(buffer, gimg, &jpeg.EncoderOptions{
				Quality:         50,
				OptimizeCoding:  true,
				ProgressiveMode: true,
			})
			if err != nil {
				continue
			} else if buffer.Len() < int(opt.ImageMinSize) {
				// 图片大小过滤
				buffer = nil //gc
				continue
			}

			hadImg = true

			if len(images) < opt.MaxImageCount {
				images = append(images, &Image{
					Ex:   "jpg",
					Body: buffer,
				})
			} else {
				buffer = nil // gc
			}
		}
		if needReadText && hadImg {
			imgPagesNum++
			if float64(imgPagesNum) > maxImagePageCount {
				needReadText = false
			}
		}
	}

	if needReadText {
		text, err = d.text()
		if err != nil {
			return nil, "", err
		}
	}
	return
}

// Analysis 执行解析
func (d *Document) Analysis(opt *Options) (images []*Image, text string, err error) {
	if d == nil || d.File == nil || len(d.Name) == 0 {
		return nil, "", ErrNoFile
	}
	if opt == nil {
		opt = &defaultOption
	}

	switch path.Ext(d.Name) {
	case ".txt":
		body, e := ioutil.ReadAll(d.File)
		if e != nil {
			return nil, "", fmt.Errorf("read file err: %v", e)
		}
		return nil, string(body), nil
	case ".xml", ".htm", ".html", ".doc":
		return d.doc(opt)
	case ".docx":
		return d.docx(opt)
	case ".pptx":
		return d.pptx(opt)
	case ".xls":
		return d.xls(opt)
	case ".xlsx":
		return d.xlsx(opt)
	case ".pdf":
		return d.pdf(opt)
	default:

	}

	return nil, "", ErrNoSupport
}
