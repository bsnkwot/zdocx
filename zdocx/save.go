package zdocx

import (
	"archive/zip"
	"bytes"
	"os"
	"time"

	"github.com/pkg/errors"
)

const (
	mainPageFooter         = "mainPageFooter"
	mainPageHeader         = "mainPageHeader"
	defaultFooter          = "defaultFooter"
	defaultHeader          = "defaultHeader"
	mainPageHeaderFileName = "header2"
	mainPageFooterFileName = "footer2"
	defaultHeaderFileName  = "header1"
	defaultFooterFileName  = "footer1"
)

type writeContentFileArgs struct {
	document *Document
	writer   *zip.Writer
}

func writeContentFile(args writeContentFileArgs) error {
	contentFile, err := args.writer.Create("word/document.xml")
	if err != nil {
		return errors.Wrap(err, "writer.Create")
	}

	args.document.writeBodyClose()

	_, err = contentFile.Write(args.document.Buf.Bytes())
	if err != nil {
		return errors.Wrap(err, "contentFile.Write")
	}

	return nil
}

type writeHeaderAndFooterFileArgs struct {
	document *Document
	writer   *zip.Writer
}

func writeHeaderAndFooterFile(args writeHeaderAndFooterFileArgs) error {
	if args.document.Header != nil {
		if err := writeHeaderOrFooter(writeHeaderOrFooterArgs{
			document:    args.document,
			p:           args.document.Header,
			writer:      args.writer,
			tag:         "hdr",
			fileName:    defaultHeaderFileName,
			sectionType: defaultHeader,
		}); err != nil {
			return errors.Wrap(err, "writeHeaderOrFooter")
		}
	}

	if args.document.Footer != nil {
		if err := writeHeaderOrFooter(writeHeaderOrFooterArgs{
			document:    args.document,
			p:           args.document.Footer,
			writer:      args.writer,
			tag:         "ftr",
			fileName:    defaultFooterFileName,
			sectionType: defaultFooter,
		}); err != nil {
			return errors.Wrap(err, "writeHeaderOrFooter")
		}
	}

	if args.document.MainPageFooter != nil {
		if err := writeHeaderOrFooter(writeHeaderOrFooterArgs{
			document:    args.document,
			p:           args.document.MainPageFooter,
			writer:      args.writer,
			tag:         "ftr",
			fileName:    mainPageFooterFileName,
			sectionType: mainPageFooter,
		}); err != nil {
			return errors.Wrap(err, "writeHeaderOrFooter")
		}
	}

	if args.document.MainPageHeader != nil {
		if err := writeHeaderOrFooter(writeHeaderOrFooterArgs{
			document:    args.document,
			p:           args.document.MainPageHeader,
			writer:      args.writer,
			tag:         "hdr",
			fileName:    mainPageHeaderFileName,
			sectionType: mainPageHeader,
		}); err != nil {
			return errors.Wrap(err, "writeHeaderOrFooter")
		}
	}

	return nil
}

type writeHeaderOrFooterArgs struct {
	document    *Document
	p           []*Paragraph
	writer      *zip.Writer
	tag         string
	fileName    string
	sectionType string
}

func (args *writeHeaderOrFooterArgs) Error() error {
	if args.sectionType == "" {
		return errors.New("no args.sectionType")
	}

	if args.tag == "" {
		return errors.New("no args.Tag")
	}

	if args.fileName == "" {
		return errors.New("no args.FileName")
	}

	return nil
}

func writeHeaderOrFooter(args writeHeaderOrFooterArgs) error {
	if err := args.Error(); err != nil {
		return err
	}

	if len(args.p) == 0 {
		return nil
	}

	var buf bytes.Buffer

	buf.WriteString(getDocumentStartTags(args.tag))

	content := []interface{}{}

	for _, p := range args.p {
		for _, i := range p.Texts {
			if i.Image != nil {
				switch args.sectionType {
				case mainPageHeader:
					i.Image.isMainPageHeader = true
				case mainPageFooter:
					i.Image.isMainPageFooter = true
				case defaultHeader:
					i.Image.isHeader = true
				case defaultFooter:
					i.Image.isFooter = true
				}
			}
		}

		p.Style.Margins = Margins{
			Top:    &Margin{Value: 0},
			Bottom: &Margin{Value: 0},
		}

		if args.sectionType == defaultFooter {
			content = append(content, p)
		} else {
			pString, err := p.string(args.document)
			if err != nil {
				return errors.Wrap(err, "p.Bytes")
			}

			buf.WriteString(pString)
		}
	}

	if args.sectionType == defaultFooter {
		pagination := Paragraph{
			isPagination: true,
		}

		style := TDStyle{
			Margins: Margins{
				Top:    &Margin{Value: 0},
				Left:   &Margin{Value: 0},
				Bottom: &Margin{Value: 0},
				Right:  &Margin{Value: 0},
			},
		}

		documentWidth := args.document.GetInnerWidth()
		table := Table{
			Type: "fixed",
			Grid: []int{
				int(float32(documentWidth) * 0.5),
				int(float32(documentWidth) * 0.5),
			},
			TR: []*TR{
				{
					TD: []*TD{
						{
							Style:   style,
							Content: content,
						},
						{
							Style: style,
							Content: []interface{}{
								&pagination,
							},
						},
					},
				},
			},
		}

		tableString, err := table.string(args.document)
		if err != nil {
			return errors.Wrap(err, "table.bytes")
		}

		buf.WriteString(tableString)
	}

	buf.WriteString("</w:" + args.tag + ">")

	contentFile, err := args.writer.Create("word/" + args.fileName + ".xml")
	if err != nil {
		return errors.Wrap(err, "writer.Create")
	}

	_, err = contentFile.Write(buf.Bytes())
	if err != nil {
		return errors.Wrap(err, "contentFile.Write")
	}

	if err := writeHeaderOrFooterRels(writeHeaderOrFooterRelsArgs{
		writer:      args.writer,
		document:    args.document,
		sectionType: args.sectionType,
		fileName:    args.fileName,
	}); err != nil {
		return errors.Wrap(err, "writeHeaderOrFooterRels")
	}

	return nil
}

type writeHeaderOrFooterRelsArgs struct {
	writer      *zip.Writer
	document    *Document
	fileName    string
	sectionType string
}

func (args *writeHeaderOrFooterRelsArgs) error() error {
	if args.sectionType == "" {
		return errors.New("no args.sectionType")
	}

	if args.fileName == "" {
		return errors.New("no args.fileName")
	}

	if args.document == nil {
		return errors.New("no args.document")
	}

	if args.writer == nil {
		return errors.New("no args.writer")
	}

	return nil
}

func writeHeaderOrFooterRels(args writeHeaderOrFooterRelsArgs) error {
	if err := args.error(); err != nil {
		return err
	}

	var images []*Image

	switch args.sectionType {
	case mainPageHeader:
		images = args.document.images.mainPageHeader
	case mainPageFooter:
		images = args.document.images.mainPageFooter
	case defaultHeader:
		images = args.document.images.header
	case defaultFooter:
		images = args.document.images.footer
	}

	if err := writeHeaderOrFooterRelsFile(writeHeaderOrFooterRelsFileArgs{
		fileName: args.fileName,
		images:   images,
		writer:   args.writer,
	}); err != nil {
		return errors.Wrap(err, "writeHeaderAndFooterRelsFile")
	}

	return nil
}

type writeHeaderOrFooterRelsFileArgs struct {
	fileName string
	images   []*Image
	writer   *zip.Writer
}

func writeHeaderOrFooterRelsFile(args writeHeaderOrFooterRelsFileArgs) error {
	if len(args.images) == 0 {
		return nil
	}

	file, err := args.writer.Create("word/_rels/" + args.fileName + ".xml.rels")
	if err != nil {
		return errors.Wrap(err, "writer.Create")
	}

	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8"?>`)
	buf.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`)

	for _, i := range args.images {
		buf.WriteString(`<Relationship Id="` + i.RelsID + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/` + i.FileName + `"/>`)
	}

	buf.WriteString(`</Relationships>`)

	_, err = file.Write(buf.Bytes())
	if err != nil {
		return errors.Wrap(err, "file.Write")
	}

	return nil
}

type writeMediaFilesArgs struct {
	images []*Image
	writer *zip.Writer
}

func isContentTypeValid(contentType string) bool {
	if contentType == "" {
		return false
	}

	if contentType == "image/jpeg" {
		return true
	}

	if contentType == "image/png" {
		return true
	}

	return false
}

func writeMediaFiles(args writeMediaFilesArgs) error {
	for _, i := range args.images {
		if !isContentTypeValid(i.ContentType) {
			continue
		}

		mediaFile, err := args.writer.Create("word/media/" + i.FileName)
		if err != nil {
			return errors.Wrap(err, "writer.Create")
		}

		_, err = mediaFile.Write(i.Bytes)
		if err != nil {
			return errors.Wrap(err, "mediaFile.Write")
		}
	}

	return nil

}

type zipFilesArgs struct {
	fileName string
	document *Document
}

func (args *zipFilesArgs) Error() error {
	if args.document == nil {
		return errors.New("no args.Documnet")
	}

	if args.fileName == "" {
		return errors.New("no args.FileName")
	}

	return nil
}

func (doc *Document) WriteToBuffer() (*bytes.Buffer, error) {
	b := new(bytes.Buffer)
	writer := zip.NewWriter(b)
	defer writer.Close()

	if err := zipWrite(zipWriteArgs{
		writer:   writer,
		document: doc,
	}); err != nil {
		return nil, errors.Wrap(err, "zipWrite")
	}

	return b, nil
}

func zipFiles(args zipFilesArgs) error {
	if err := args.Error(); err != nil {
		return err
	}

	newZip, err := os.Create(args.fileName)
	if err != nil {
		return err
	}

	defer newZip.Close()

	writer := zip.NewWriter(newZip)
	defer writer.Close()

	if err := zipWrite(zipWriteArgs{
		writer:   writer,
		document: args.document,
	}); err != nil {
		return errors.Wrap(err, "zipWrite")
	}

	return nil
}

type zipWriteArgs struct {
	writer   *zip.Writer
	document *Document
}

func zipWrite(args zipWriteArgs) error {
	if err := writeContentFile(writeContentFileArgs{
		document: args.document,
		writer:   args.writer,
	}); err != nil {
		return errors.Wrap(err, "setContent")
	}

	if err := writeHeaderAndFooterFile(writeHeaderAndFooterFileArgs{
		document: args.document,
		writer:   args.writer,
	}); err != nil {
		return errors.Wrap(err, "writeHeaderAndFooterFile")
	}

	if err := writeCorePropertiesFile(writeCorePropertiesFileArgs{
		writer: args.writer,
		lang:   args.document.Lang,
	}); err != nil {
		return errors.Wrap(err, "writeCorePropertiesFile")
	}

	if err := writeSettingsFile(writeSettingsFileArgs{
		writer: args.writer,
		lang:   args.document.Lang,
	}); err != nil {
		return errors.Wrap(err, "writeSettingsFile")
	}

	if err := writeContentTypesFile(writeContentTypesFileArgs{
		writer:   args.writer,
		document: args.document,
	}); err != nil {
		return errors.Wrap(err, "writeContentTypesFile")
	}

	if err := writeMediaFiles(writeMediaFilesArgs{
		images: args.document.images.content,
		writer: args.writer,
	}); err != nil {
		return errors.Wrap(err, "writeMediaFiles")
	}

	if err := writeMediaFiles(writeMediaFilesArgs{
		images: args.document.images.footer,
		writer: args.writer,
	}); err != nil {
		return errors.Wrap(err, "writeMediaFiles")
	}

	if err := writeMediaFiles(writeMediaFilesArgs{
		images: args.document.images.header,
		writer: args.writer,
	}); err != nil {
		return errors.Wrap(err, "writeMediaFiles")
	}

	if err := writeWordRelsFile(writeWordRelsFileArgs{
		document: args.document,
		writer:   args.writer,
	}); err != nil {
		return errors.Wrap(err, "writeWordRelsFile")
	}

	// if err := writeStylesFile(writeStylesFileArgs{
	// 	document: args.document,
	// 	writer:   args.writer,
	// }); err != nil {
	// 	return errors.Wrap(err, "writeStylesFile")
	// }

	if err := writeTemplatesFiles(writeTemplatesFilesArgs{
		writer: args.writer,
	}); err != nil {
		return errors.Wrap(err, "writeTemplatesFiles")
	}

	return nil
}

type writeTemplatesFilesArgs struct {
	writer *zip.Writer
}

type writeStylesFileArgs struct {
	writer   *zip.Writer
	document *Document
}

// func writeStylesFile(args writeStylesFileArgs) error {
// 	file, err := os.Open("temp/styles.xml")
// 	if err != nil {
// 		return errors.Wrap(err, "os.Open")
// 	}

// 	defer file.Close()

// 	info, err := file.Stat()
// 	if err != nil {
// 		return errors.Wrap(err, "file.Stat")
// 	}

// 	header, err := zip.FileInfoHeader(info)
// 	if err != nil {
// 		return errors.Wrap(err, "zip.FileInfoHeader")
// 	}

// 	header.Name = "word/styles.xml"
// 	header.Method = zip.Deflate

// 	writer, err := args.writer.CreateHeader(header)
// 	if err != nil {
// 		return errors.Wrap(err, "zip.Writer.CreateHeader")
// 	}

// 	_, err = io.Copy(writer, file)
// 	if err != nil {
// 		return errors.Wrap(err, "io.Copy")
// 	}

// 	return nil
// }

func writeTemplatesFiles(args writeTemplatesFilesArgs) error {
	for _, file := range templatesFilesList() {
		newFile, err := args.writer.Create(file.FullName())
		if err != nil {
			return errors.Wrap(err, "writer.Create")
		}

		_, err = newFile.Write(file.bytes)
		if err != nil {
			return errors.Wrap(err, "contentFile.Write")
		}
	}

	return nil
}

type writeWordRelsFileArgs struct {
	document *Document
	writer   *zip.Writer
}

func writeWordRelsFile(args writeWordRelsFileArgs) error {
	file, err := args.writer.Create("word/_rels/document.xml.rels")
	if err != nil {
		return errors.Wrap(err, "writer.Create")
	}

	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8"?>`)
	buf.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`)
	buf.WriteString(`<Relationship Id="rId` + stylesID + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`)
	buf.WriteString(`<Relationship Id="rId` + numberingID + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>`)
	buf.WriteString(`<Relationship Id="rId` + fontTableID + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>`)
	buf.WriteString(`<Relationship Id="rId` + settingsID + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>`)
	buf.WriteString(`<Relationship Id="rId` + themeID + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>`)

	if len(args.document.images.content) != 0 {
		for _, i := range args.document.images.content {
			buf.WriteString(`<Relationship Id="` + i.RelsID + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/` + i.FileName + `"/>`)
		}
	}

	if args.document.Header != nil {
		buf.WriteString(`<Relationship Id="rId` + defaultHeaderID + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="` + defaultHeaderFileName + `.xml"/>`)
	}

	if args.document.Footer != nil {
		buf.WriteString(`<Relationship Id="rId` + defaultFooterID + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="` + defaultFooterFileName + `.xml"/>`)
	}

	if args.document.MainPageHeader != nil {
		buf.WriteString(`<Relationship Id="rId` + mainPageHeaderID + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="` + defaultFooterFileName + `.xml"/>`)
	}

	if args.document.MainPageFooter != nil {
		buf.WriteString(`<Relationship Id="rId` + mainPageFooterID + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="` + mainPageFooterFileName + `.xml"/>`)
	}

	for _, i := range args.document.Links {
		buf.WriteString(`<Relationship Id="` + i.ID + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="` + i.URL + `" TargetMode="External"/>`)
	}

	buf.WriteString(`</Relationships>`)

	_, err = file.Write(buf.Bytes())
	if err != nil {
		return errors.Wrap(err, "file.Write")
	}

	return nil
}

type writeCorePropertiesFileArgs struct {
	lang   string
	writer *zip.Writer
}

func writeCorePropertiesFile(args writeCorePropertiesFileArgs) error {
	createdAt := time.Now()
	lang := "ru-RU"

	if args.lang == "en" {
		lang = "en-Us"
	}

	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">`)
	buf.WriteString(`<dcterms:created xsi:type="dcterms:W3CDTF">` + createdAt.Format("2006-01-02T15:04:05Z") + `</dcterms:created>`)
	buf.WriteString(`<dc:creator>Labrika</dc:creator>`)
	buf.WriteString(`<dc:description></dc:description>`)
	buf.WriteString(`<dc:language>` + lang + `</dc:language>`)
	buf.WriteString(`<cp:lastModifiedBy></cp:lastModifiedBy>`)
	buf.WriteString(`<dcterms:modified xsi:type="dcterms:W3CDTF">` + createdAt.Format("2006-01-02T15:04:05Z") + `</dcterms:modified>`)
	buf.WriteString(`<cp:revision>4</cp:revision>`)
	buf.WriteString(`<dc:subject></dc:subject>`)
	buf.WriteString(`<dc:title></dc:title>`)
	buf.WriteString(`</cp:coreProperties>`)

	file, err := args.writer.Create("docProps/core.xml")
	if err != nil {
		return errors.Wrap(err, "writer.Create")
	}

	_, err = file.Write(buf.Bytes())
	if err != nil {
		return errors.Wrap(err, "file.Write")
	}

	return nil
}

type writeContentTypesFileArgs struct {
	lang     string
	document *Document
	writer   *zip.Writer
}

func writeContentTypesFile(args writeContentTypesFileArgs) error {
	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8"?>`)
	buf.WriteString(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`)
	buf.WriteString(`<Default Extension="xml" ContentType="application/xml"/>`)
	buf.WriteString(`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`)
	buf.WriteString(`<Default Extension="png" ContentType="image/png"/>`)
	buf.WriteString(`<Default Extension="jpeg" ContentType="image/jpeg"/>`)
	buf.WriteString(`<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`)
	buf.WriteString(`<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`)
	buf.WriteString(`<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`)
	buf.WriteString(`<Override PartName="/word/_rels/document.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`)
	buf.WriteString(`<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>`)
	buf.WriteString(`<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>`)

	for _, i := range args.document.images.content {
		buf.WriteString(`<Override PartName="/word/media/` + i.FileName + `" ContentType="` + i.ContentType + `"/>`)
	}

	for _, i := range args.document.images.header {
		buf.WriteString(`<Override PartName="/word/media/` + i.FileName + `" ContentType="` + i.ContentType + `"/>`)
	}

	for _, i := range args.document.images.footer {
		buf.WriteString(`<Override PartName="/word/media/` + i.FileName + `" ContentType="` + i.ContentType + `"/>`)
	}

	if args.document.Footer != nil {
		buf.WriteString(`<Override PartName="/word/` + defaultFooterFileName + `.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>`)
	}

	if args.document.MainPageFooter != nil {
		buf.WriteString(`<Override PartName="/word/` + mainPageFooterFileName + `.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>`)
	}

	if args.document.Header != nil {
		buf.WriteString(`<Override PartName="/word/` + defaultHeaderFileName + `.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`)
	}

	if args.document.MainPageHeader != nil {
		buf.WriteString(`<Override PartName="/word/` + mainPageHeaderFileName + `.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`)
	}

	buf.WriteString(`<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>`)
	buf.WriteString(`<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>`)
	buf.WriteString(`<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>`)
	buf.WriteString(`<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`)
	buf.WriteString(`</Types>`)

	file, err := args.writer.Create("[Content_Types].xml")
	if err != nil {
		return errors.Wrap(err, "writer.Create")
	}

	_, err = file.Write(buf.Bytes())
	if err != nil {
		return errors.Wrap(err, "file.Write")
	}

	return nil
}

type writeSettingsFileArgs struct {
	lang   string
	writer *zip.Writer
}

func writeSettingsFile(args writeSettingsFileArgs) error {
	lang := "ru-RU"

	if args.lang == "en" {
		lang = "en-Us"
	}

	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString(`<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`)
	buf.WriteString(`<w:zoom w:percent="100"/>`)
	buf.WriteString(`<w:defaultTabStop w:val="708"/>`)
	buf.WriteString(`<w:autoHyphenation w:val="true"/>`)
	buf.WriteString(`<w:compat>`)
	buf.WriteString(`<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/>`)
	buf.WriteString(`<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>`)
	buf.WriteString(`<w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>`)
	buf.WriteString(`<w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>`)
	buf.WriteString(`</w:compat>`)
	buf.WriteString(`<w:themeFontLang w:val="` + lang + `" w:eastAsia="" w:bidi=""/>`)
	buf.WriteString(`</w:settings>`)

	file, err := args.writer.Create("word/settings.xml")
	if err != nil {
		return errors.Wrap(err, "writer.Create")
	}

	_, err = file.Write(buf.Bytes())
	if err != nil {
		return errors.Wrap(err, "file.Write")
	}

	return nil
}

type templateFile struct {
	name     string
	savePath string
	bytes    []byte
}

func (i *templateFile) FullName() string {
	if i.savePath == "" {
		return i.name
	}

	return i.savePath + "/" + i.name
}

func templatesFilesList() []*templateFile {
	return []*templateFile{
		{
			name:     ".rels",
			savePath: "_rels",
			bytes:    []byte(templateRelsRels),
		},
		{
			name:     "app.xml",
			savePath: "docProps",
			bytes:    []byte(templateDocPropsApp),
		},
		{
			name:     "styles.xml",
			savePath: "word",
			bytes:    []byte(templateWordStyles),
		},
		{
			name:     "numbering.xml",
			savePath: "word",
			bytes:    []byte(templateWordNumbering),
		},
		{
			name:     "fontTable.xml",
			savePath: "word",
			bytes:    []byte(templateWordFontTable),
		},
		{
			name:     "theme1.xml",
			savePath: "word/theme",
			bytes:    []byte(templateWordTheme),
		},
	}
}
