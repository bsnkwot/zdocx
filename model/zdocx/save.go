package zdocx

import (
	"archive/zip"
	"bytes"
	"os"
	"time"

	"github.com/pkg/errors"
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
	if len(args.document.Header) > 0 {
		if err := writeHeaderOrFooter(writeHeaderOrFooterArgs{
			document: args.document,
			p:        args.document.Header,
			writer:   args.writer,
			tag:      "hdr",
			fileName: "header",
		}); err != nil {
			return errors.Wrap(err, "writeHeaderOrFooter")
		}
	}

	if len(args.document.Footer) > 0 {
		if err := writeHeaderOrFooter(writeHeaderOrFooterArgs{
			document: args.document,
			p:        args.document.Footer,
			writer:   args.writer,
			tag:      "ftr",
			fileName: "footer",
		}); err != nil {
			return errors.Wrap(err, "writeHeaderOrFooter")
		}
	}

	return nil
}

type writeHeaderOrFooterArgs struct {
	document *Document
	p        []*Paragraph
	writer   *zip.Writer
	tag      string
	fileName string
}

func (args *writeHeaderOrFooterArgs) Error() error {
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

	for _, p := range args.p {
		p.StyleClass = args.fileName + "Class"

		pString, err := p.String(args.document)
		if err != nil {
			return errors.Wrap(err, "Paragraph.String")
		}

		buf.WriteString(pString)
	}

	buf.WriteString("</w:" + args.tag + ">")

	contentFile, err := args.writer.Create("word/" + args.fileName + "1.xml")
	if err != nil {
		return errors.Wrap(err, "writer.Create")
	}

	_, err = contentFile.Write(buf.Bytes())
	if err != nil {
		return errors.Wrap(err, "contentFile.Write")
	}

	return nil
}

type writeMediaFilesArgs struct {
	document *Document
	writer   *zip.Writer
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
	for _, i := range args.document.Images {
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

	if err := writeContentFile(writeContentFileArgs{
		document: args.document,
		writer:   writer,
	}); err != nil {
		return errors.Wrap(err, "setContent")
	}

	if err := writeHeaderAndFooterFile(writeHeaderAndFooterFileArgs{
		document: args.document,
		writer:   writer,
	}); err != nil {
		return errors.Wrap(err, "writeHeaderAndFooterFile")
	}

	if err := writeCorePropertiesFile(writeCorePropertiesFileArgs{
		writer: writer,
		lang:   args.document.Lang,
	}); err != nil {
		return errors.Wrap(err, "writeCorePropertiesFile")
	}

	if err := writeSettingsFile(writeSettingsFileArgs{
		writer: writer,
		lang:   args.document.Lang,
	}); err != nil {
		return errors.Wrap(err, "writeSettingsFile")
	}

	if err := writeContentTypesFile(writeContentTypesFileArgs{
		writer:   writer,
		document: args.document,
	}); err != nil {
		return errors.Wrap(err, "writeContentTypesFile")
	}

	if err := writeMediaFiles(writeMediaFilesArgs{
		document: args.document,
		writer:   writer,
	}); err != nil {
		return errors.Wrap(err, "writeMediaFiles")
	}

	if err := writeWordRelsFile(writeWordRelsFileArgs{
		document: args.document,
		writer:   writer,
	}); err != nil {
		return errors.Wrap(err, "writeWordRelsFile")
	}

	// if err := writeStylesFile(writeStylesFileArgs{
	// 	document: args.document,
	// 	writer:   writer,
	// }); err != nil {
	// 	return errors.Wrap(err, "writeStylesFile")
	// }

	if err := writeTemplatesFiles(writeTemplatesFilesArgs{
		writer: writer,
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
	buf.WriteString(`<Relationship Id="rId` + StylesID + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`)
	buf.WriteString(`<Relationship Id="rId` + NumberingID + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>`)
	buf.WriteString(`<Relationship Id="rId` + FontTableID + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>`)
	buf.WriteString(`<Relationship Id="rId` + SettingsID + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>`)
	buf.WriteString(`<Relationship Id="rId` + ThemeID + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>`)

	if len(args.document.Images) != 0 {
		for _, i := range args.document.Images {
			buf.WriteString(`<Relationship Id="` + i.RelsID + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/` + i.FileName + `"/>`)
		}
	}

	if len(args.document.Header) > 0 {
		buf.WriteString(`<Relationship Id="rId` + HeaderID + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>`)
	}

	if len(args.document.Footer) > 0 {
		buf.WriteString(`<Relationship Id="rId` + FooterID + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>`)
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

	for _, i := range args.document.Images {
		buf.WriteString(`<Override PartName="/word/media/` + i.FileName + `" ContentType="` + i.ContentType + `"/>`)
	}

	if len(args.document.Footer) > 0 {
		buf.WriteString(`<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>`)
	}

	if len(args.document.Header) > 0 {
		buf.WriteString(`<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`)
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

// func addTemplateFileToZip(zipWriter *zip.Writer, fileData *templateFile) error {
// 	file, err := os.Open(fileData. + "/" + fileData.Name)
// 	if err != nil {
// 		return errors.Wrap(err, "os.Opern")
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

// 	header.Name = fileData.FullName()
// 	header.Method = zip.Deflate

// 	writer, err := zipWriter.CreateHeader(header)
// 	if err != nil {
// 		return errors.Wrap(err, "zipWriter.CreateHeader")
// 	}

// 	_, err = io.Copy(writer, file)
// 	if err != nil {
// 		return errors.Wrap(err, "io.Copy")
// 	}

// 	return nil
// }
