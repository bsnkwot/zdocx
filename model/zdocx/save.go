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

	if err := setCoreProperies(setCoreProperiesArgs{
		Writer: writer,
		Lang:   args.Document.Lang,
	}); err != nil {
		return errors.Wrap(err, "setContent")
	}

	if err := setHeaderAndFooter(setHeaderAndFooterArgs{
		Document: args.Document,
		Writer:   writer,
	}); err != nil {
		return errors.Wrap(err, "setHeaderAndFooter")
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
