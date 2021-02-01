package zdocx

import (
	"archive/zip"
	"bytes"
	"os"
	"strconv"
	"time"

	"github.com/pkg/errors"
)

type setContentArgs struct {
	Document *Document
	Writer   *zip.Writer
}

func setContent(args setContentArgs) error {
	contentFile, err := args.Writer.Create("word/document.xml")
	if err != nil {
		return errors.Wrap(err, "writer.Create")
	}

	_, err = contentFile.Write(args.Document.Buf.Bytes())
	if err != nil {
		return errors.Wrap(err, "contentFile.Write")
	}

	return nil
}

type setHeaderAndFooterArgs struct {
	Document *Document
	Writer   *zip.Writer
}

func setHeaderAndFooter(args setHeaderAndFooterArgs) error {
	if len(args.Document.Header) > 0 {
		if err := setHeaderOrFooter(setHeaderOrFooterArgs{
			Document: args.Document,
			Writer:   args.Writer,
			Tag:      "hdr",
			FileName: "header",
		}); err != nil {
			return errors.Wrap(err, "setHeaderOrFooter")
		}
	}

	if len(args.Document.Footer) > 0 {
		if err := setHeaderOrFooter(setHeaderOrFooterArgs{
			Document: args.Document,
			Writer:   args.Writer,
			Tag:      "ftr",
			FileName: "footer",
		}); err != nil {
			return errors.Wrap(err, "setHeaderOrFooter")
		}
	}

	return nil
}

type setHeaderOrFooterArgs struct {
	Document *Document
	Writer   *zip.Writer
	Tag      string
	FileName string
}

func (args *setHeaderOrFooterArgs) Error() error {
	if args.Tag == "" {
		return errors.New("no args.Tag")
	}

	if args.FileName == "" {
		return errors.New("no args.FileName")
	}

	return nil
}

func setHeaderOrFooter(args setHeaderOrFooterArgs) error {
	if err := args.Error(); err != nil {
		return err
	}

	var buf bytes.Buffer

	buf.WriteString(getDocumentStartTags("hdr"))

	for _, p := range args.Document.Header {
		p.StyleClass = args.FileName + "Class"
		buf.WriteString(p.String())
	}

	buf.WriteString("</w:hdr>")

	contentFile, err := args.Writer.Create("word/" + args.FileName + "1.xml")
	if err != nil {
		return errors.Wrap(err, "writer.Create")
	}

	_, err = contentFile.Write(buf.Bytes())
	if err != nil {
		return errors.Wrap(err, "contentFile.Write")
	}

	return nil
}

type zipFilesArgs struct {
	FileName       string
	TemplatesFiles []*templateFile
	Document       *Document
}

func (args *zipFilesArgs) Error() error {
	if args.Document == nil {
		return errors.New("no args.Documnet")
	}

	if args.FileName == "" {
		return errors.New("no args.FileName")
	}

	return nil
}

func zipFiles(args zipFilesArgs) error {
	if err := args.Error(); err != nil {
		return err
	}

	newZip, err := os.Create(args.FileName)
	if err != nil {
		return err
	}

	defer newZip.Close()

	writer := zip.NewWriter(newZip)
	defer writer.Close()

	if err := setContent(setContentArgs{
		Document: args.Document,
		Writer:   writer,
	}); err != nil {
		return errors.Wrap(err, "setContent")
	}

	if err := setHeaderAndFooter(setHeaderAndFooterArgs{
		Document: args.Document,
		Writer:   writer,
	}); err != nil {
		return errors.Wrap(err, "setHeaderAndFooter")
	}

	if err := setCoreProperies(setCoreProperiesArgs{
		Writer: writer,
		Lang:   args.Document.Lang,
	}); err != nil {
		return errors.Wrap(err, "setCoreProperies")
	}

	if err := setSettings(setSettingsArgs{
		Writer: writer,
		Lang:   args.Document.Lang,
	}); err != nil {
		return errors.Wrap(err, "setSettings")
	}

	if err := setContentTypes(setContentTypesArgs{
		Writer:   writer,
		Document: args.Document,
	}); err != nil {
		return errors.Wrap(err, "setContentTypes")
	}

	if err := setWordRels(setWordRelsArgs{
		Document: args.Document,
		Writer:   writer,
	}); err != nil {
		return errors.Wrap(err, "setWrodRels")
	}

	for _, file := range args.TemplatesFiles {
		newFile, err := writer.Create(file.FullName())
		if err != nil {
			return errors.Wrap(err, "writer.Create")
		}

		_, err = newFile.Write(file.Bytes)
		if err != nil {
			return errors.Wrap(err, "contentFile.Write")
		}
	}

	return nil
}

type setWordRelsArgs struct {
	Document *Document
	Writer   *zip.Writer
}

func setWordRels(args setWordRelsArgs) error {
	file, err := args.Writer.Create("word/_rels/document.xml.rels")
	if err != nil {
		return errors.Wrap(err, "writer.Create")
	}

	var buf bytes.Buffer
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8"?>`)
	buf.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`)
	buf.WriteString(`<Relationship Id="rId` + strconv.Itoa(StylesID) + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`)
	buf.WriteString(`<Relationship Id="rId` + strconv.Itoa(NumberingID) + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>`)
	buf.WriteString(`<Relationship Id="rId` + strconv.Itoa(FontTableID) + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>`)
	buf.WriteString(`<Relationship Id="rId` + strconv.Itoa(SettingsID) + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>`)
	buf.WriteString(`<Relationship Id="rId` + strconv.Itoa(ThemeID) + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>`)

	if len(args.Document.Images) != 0 {
		buf.WriteString(`<Relationship Id="rId` + strconv.Itoa(ImagesID) + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>`)
	}

	if len(args.Document.Header) > 0 {
		buf.WriteString(`<Relationship Id="rId` + strconv.Itoa(HeaderID) + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>`)
	}

	if len(args.Document.Footer) > 0 {
		buf.WriteString(`<Relationship Id="rId` + strconv.Itoa(FooterID) + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>`)
	}

	for index, i := range args.Document.Links {
		buf.WriteString(`<Relationship Id="` + (LinkIDPrefix + strconv.Itoa(index)) + `" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="` + i.URL + `" TargetMode="External"/>`)
	}

	buf.WriteString(`</Relationships>`)

	_, err = file.Write(buf.Bytes())
	if err != nil {
		return errors.Wrap(err, "file.Write")
	}

	return nil
}

type setCoreProperiesArgs struct {
	Lang   string
	Writer *zip.Writer
}

func setCoreProperies(args setCoreProperiesArgs) error {
	createdAt := time.Now()
	lang := "ru-RU"

	if args.Lang == "en" {
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

	file, err := args.Writer.Create("docProps/core.xml")
	if err != nil {
		return errors.Wrap(err, "writer.Create")
	}

	_, err = file.Write(buf.Bytes())
	if err != nil {
		return errors.Wrap(err, "file.Write")
	}

	return nil
}

type setContentTypesArgs struct {
	Lang     string
	Document *Document
	Writer   *zip.Writer
}

func setContentTypes(args setContentTypesArgs) error {
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

	for _, i := range args.Document.Images {
		buf.WriteString(`<Override PartName="/word/media/` + i.FileName + `" ContentType="` + i.ContentType + `"/>`)
	}

	if len(args.Document.Footer) > 0 {
		buf.WriteString(`<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>`)
	}

	if len(args.Document.Header) > 0 {
		buf.WriteString(`<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`)
	}

	buf.WriteString(`<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>`)
	buf.WriteString(`<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>`)
	buf.WriteString(`<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>`)
	buf.WriteString(`<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`)
	buf.WriteString(`</Types>`)

	file, err := args.Writer.Create("[Content_Types].xml")
	if err != nil {
		return errors.Wrap(err, "writer.Create")
	}

	_, err = file.Write(buf.Bytes())
	if err != nil {
		return errors.Wrap(err, "file.Write")
	}

	return nil
}

type setSettingsArgs struct {
	Lang   string
	Writer *zip.Writer
}

func setSettings(args setSettingsArgs) error {
	lang := "ru-RU"

	if args.Lang == "en" {
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

	file, err := args.Writer.Create("word/settings.xml")
	if err != nil {
		return errors.Wrap(err, "writer.Create")
	}

	_, err = file.Write(buf.Bytes())
	if err != nil {
		return errors.Wrap(err, "file.Write")
	}

	return nil
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
