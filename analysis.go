package docanalysis

import (
	"bytes"
	"fmt"
	"image/jpeg"
	"io"
	"io/ioutil"
	"path"
	"strings"

	"code.sajari.com/docconv"

	"baliance.com/gooxml/common"
	"baliance.com/gooxml/spreadsheet"

	"baliance.com/gooxml/document"
	"baliance.com/gooxml/presentation"

	"github.com/extrame/xls"
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
	Body []byte
}

func (i *Image) String() string {
	return fmt.Sprintf("%s.%s", i.Path, i.Ex)
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

func (d *Document) text() (string, error) {
	mime := docconv.MimeTypeByExtension(d.Name)
	resp, e := docconv.Convert(d.File, mime, true)
	if e != nil {
		return "", fmt.Errorf("convert %v error: %v", mime, e)
	}
	return resp.Body, nil
}

func (d *Document) doc() (images []*Image, text string, err error) {
	text, err = d.text()
	if err != nil {
		return nil, "", err
	}

	return nil, text, err
}

func (d *Document) docx() (images []*Image, text string, err error) {
	text, err = d.text()
	if err != nil {
		return nil, "", err
	}

	doc, e := document.Read(d.File, d.Size)
	if e == nil {
		// 图像
		images = getImages(doc.Images)
	}
	return
}

func (d *Document) xls() (images []*Image, text string, err error) {
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
	text, err = d.text()
	if err != nil {
		return nil, "", err
	}

	// get img
	readSeeker := io.NewSectionReader(d.File, 0, d.Size)
	pdfReader, e := pdf.NewPdfReader(readSeeker)
	if e != nil {
		return nil, text, nil
	}

	// pdf 是否加密
	isEncrypted, err := pdfReader.IsEncrypted()
	if err != nil {
		return nil, text, nil
	} else if isEncrypted {
		return nil, text, nil
	}

	pages, err := pdfReader.GetNumPages()
	if err != nil {
		return nil, text, nil
	}

	// pagenum 从第二页开始
	// 第一页数据会在ocr中处理
	// 所以跳过第一页
	for pageNum := 2; pageNum <= pages; pageNum++ {
		page, e := pdfReader.GetPage(pageNum)
		if e != nil {
			continue
		}

		// 检测本页图片
		rgbImages, err := extractImagesOnPage(page)
		if err != nil {
			// 如果图片解析错误，直接跳过
			continue
			// return nil, "", fmt.Errorf("pdf ext image err: %v", err)
		}

		for _, img := range rgbImages {
			gimg, err := img.ToGoImage()
			if err != nil {
				continue
				// return nil, "", fmt.Errorf("pdf img2goimg err: %v", err)
			}
			buffer := bytes.NewBuffer([]byte{})
			opt := jpeg.Options{Quality: 100}
			err = jpeg.Encode(buffer, gimg, &opt)
			if err != nil {
				continue
				// return nil, "", fmt.Errorf("pdf img to jpeg err: %v", err)
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

	switch path.Ext(d.Name) {
	case ".txt":
		body, e := ioutil.ReadAll(d.File)
		if e != nil {
			return nil, "", fmt.Errorf("read file err: %v", e)
		}
		return nil, string(body), nil
	case ".xml", ".htm", ".html", ".doc":
		return d.doc()
	case ".docx":
		return d.docx()
	case ".pptx":
		return d.pptx()
	case ".xls":
		return d.xls()
	case ".xlsx":
		return d.xlsx()
	case ".pdf":
		return d.pdf()
	default:

	}

	return nil, "", ErrNoSupport
}
