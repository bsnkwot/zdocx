package zdocx

import (
	"bytes"
	"image"
	_ "image/jpeg"
	"net/http"
	"path/filepath"
	"strconv"

	"github.com/pkg/errors"
)

const (
	ListDecimalID               = 1
	ListBulletID                = 2
	ListNoneID                  = 3
	ListDecimalType             = "decimal"
	ListBulletType              = "bullet"
	ListNoneType                = "none"
	TableCellDefaultMargin      = 55
	DocumentDefaultMargin       = 1440
	PageOrientationAlbum        = "album"
	PageOrientationBook         = "book"
	StylesID                    = "fileStylesID"
	ImagesID                    = "fileImagesID"
	NumberingID                 = "fileNumberingID"
	FontTableID                 = "fileFontTableID"
	SettingsID                  = "fileSettingsID"
	ThemeID                     = "fileThemeID"
	HeaderID                    = "fileHeaderID"
	FooterID                    = "fileFooterID"
	LinkIDPrefix                = "fileLinkId"
	DefaultImageVerticalAlign   = "top"
	DefaultImageHorisontalAlign = "center"
)

type Document struct {
	Buf             bytes.Buffer
	Header          []*Paragraph
	Footer          []*Paragraph
	PageOrientation string
	Lang            string
	Margin          *Margin
	FontSize        int
	Images          []*Image
	Links           []*Link
}

type Link struct {
	URL  string
	Text string
	ID   string
}

type Image struct {
	FileName        string
	Extension       string
	ContentType     string
	Description     string
	RelsID          string
	HorisontalAlign string
	VerticalAlign   string
	Width           int
	Height          int
	ZIndex          int
	IsRelative      bool
	WrapText        bool
	Margin          *Margin
	ID              int
	Bytes           []byte
}

type SectionProperties struct {
}

type ListParams struct {
	Level int
	Type  string
}

type Style struct {
	FontSize   int
	IsBold     bool
	IsItalic   bool
	FontFamily string
	Color      string
	MarginLeft int
}

type Text struct {
	Text       string
	Link       *Link
	Image      *Image
	StyleClass string
	Style
}

type Paragraph struct {
	Texts      []*Text
	ListParams *ListParams
	StyleClass string
	Style
}

type List struct {
	LI         []*LI
	Type       string
	StyleClass string
	Style
}

type LI struct {
	Items []interface{}
}

type TR struct {
	TD []*TD
}

type TD struct {
	Content []interface{}
}

type Margin struct {
	Top    int
	Left   int
	Right  int
	Bottom int
}

type Table struct {
	TR         []*TR
	Grid       []int
	StyleClass string
	Width      int
	CellMargin *Margin
}

func NewDocument() *Document {
	doc := Document{}
	doc.writeStartTags()
	doc.writeBody()
	return &doc
}

type SaveArgs struct {
	FileName string
}

func (args *SaveArgs) Error() error {
	if args.FileName == "" {
		return errors.New("no args.FileName")
	}

	return nil
}

func (d *Document) Save(args SaveArgs) error {
	if err := args.Error(); err != nil {
		return err
	}

	if err := zipFiles(zipFilesArgs{
		fileName: "test.docx",
		document: d,
	}); err != nil {
		return errors.Wrap(err, "ZipFiles")
	}

	return nil
}

func (d *Document) writeStartTags() {
	d.Buf.WriteString(getDocumentStartTags("document"))
}

func getDocumentStartTags(tag string) string {
	return `<?xml version="1.0" encoding="utf-8" standalone="yes"?><w:` + tag + ` xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">`
}

func (d *Document) writeBody() {
	d.Buf.WriteString("<w:body>")
}

func (d *Document) writeBodyClose() {
	d.writeSectionProperties()
	d.Buf.WriteString("</w:body>")
	d.Buf.WriteString("</w:document>")
}

func (d *Document) String() string {
	return d.Buf.String()
}

func (d *Document) SetSpace() {
	d.writeSpace()
}

func (d *Document) writeSpace() {
	d.Buf.WriteString(getSpace())
	// return setSpace()
}

func getSpace() string {
	return `<w:r><w:t xml:space="preserve"> </w:t></w:r>`
}

func (d *Document) SetP(p *Paragraph) error {
	if err := d.writeP(p); err != nil {
		return errors.Wrap(err, "Document.SetP")
	}

	return nil
}

func (d *Document) writeP(p *Paragraph) error {
	pString, err := p.String(d)
	if err != nil {
		return errors.Wrap(err, "Paragraph.String")
	}

	d.Buf.WriteString(pString)

	return nil
}

func (p *Paragraph) String(d *Document) (string, error) {
	var buf bytes.Buffer

	buf.WriteString("<w:p>")
	buf.WriteString(p.GetProperties())

	for index, t := range p.Texts {
		if index != 0 {
			buf.WriteString(getSpace())
		}

		textString, err := t.String(d)
		if err != nil {
			return "", errors.Wrap(err, "Text.String")
		}

		buf.WriteString(textString)
	}

	buf.WriteString("</w:p>")

	return buf.String(), nil
}

func (p *Paragraph) GetListParams() string {
	if p.ListParams == nil {
		return ""
	}

	var buf bytes.Buffer
	buf.WriteString(`<w:pStyle w:val="ListParagraph" />`)
	buf.WriteString("<w:numPr>")
	buf.WriteString(`<w:ilvl w:val="` + strconv.Itoa(p.ListParams.Level) + `" />`)

	var id int
	switch p.ListParams.Type {
	case ListBulletType:
		id = ListBulletID
	case ListDecimalType:
		id = ListDecimalID
	case ListNoneType:
		id = ListNoneID
	}

	buf.WriteString(`<w:numId w:val="` + strconv.Itoa(id) + `" />`)
	buf.WriteString("</w:numPr>")

	return buf.String()
}

func (p *Paragraph) HasProperties() bool {
	if p.StyleClass != "" {
		return true
	}

	if !p.Style.IsEmpty() {
		return true
	}

	if p.ListParams != nil {
		return true
	}

	return false
}

func (p *Paragraph) GetProperties() string {
	if !p.HasProperties() {
		return ""
	}

	var buf bytes.Buffer

	buf.WriteString("<w:pPr>")
	buf.WriteString(p.GetListParams())
	buf.WriteString(getCommonStyleClass(p.StyleClass))
	buf.WriteString(getCommonStyle(p.Style))
	buf.WriteString("</w:pPr>")

	return buf.String()
}

func (t *Text) String(d *Document) (string, error) {
	if t == nil {
		return "", nil
	}

	if t.Text == "" && t.Image == nil {
		return "", nil
	}

	var buf bytes.Buffer
	if t.Link != nil {
		t.Link.ID = LinkIDPrefix + strconv.Itoa(len(d.Links))
		d.Links = append(d.Links, t.Link)
		buf.WriteString(`<w:hyperlink r:id="` + t.Link.ID + `">`)
	}

	if t.Image != nil {
		imageString, err := t.Image.String(d)
		if err != nil {
			return "", errors.Wrap(err, "Iamge.String")
		}
		buf.WriteString(imageString)
	}

	buf.WriteString("<w:r>")
	buf.WriteString(t.GetProperties())
	buf.WriteString("<w:t>" + t.Text + "</w:t>")
	buf.WriteString("</w:r>")

	if t.Link != nil {
		buf.WriteString("</w:hyperlink>")
	}

	return buf.String(), nil
}

func (t *Text) GetProperties() string {
	if t.Style.IsEmpty() && t.StyleClass == "" && t.Link == nil {
		return ""
	}

	var buf bytes.Buffer

	buf.WriteString("<w:rPr>")

	if t.Link != nil {
		buf.WriteString(`<w:rStyle w:val="hyperlink" />`)
	} else {
		buf.WriteString(getCommonStyleClass(t.StyleClass))
	}

	buf.WriteString(getCommonStyle(t.Style))
	buf.WriteString("</w:rPr>")

	return buf.String()
}

func getCommonStyle(style Style) string {
	var buf bytes.Buffer

	if style.MarginLeft != 0 {
		buf.WriteString(`<w:pStyle w:val="Normal" />`)
		buf.WriteString(`<w:ind w:left="` + strconv.Itoa(style.MarginLeft) + `" />`)
	}

	if style.FontFamily != "" {
		buf.WriteString("<w:rFonts w:ascii=\"" + style.FontFamily + "\" w:hAnsi=\"" + style.FontFamily + "\" />")
	}

	if style.FontSize != 0 {
		buf.WriteString("<w:sz w:val=\"" + strconv.Itoa(style.FontSize) + "\"/>")
	}

	if style.IsBold {
		buf.WriteString("<w:b />")
	}

	if style.IsItalic {
		buf.WriteString("<w:i />")
	}

	return buf.String()
}

func getCommonStyleClass(styleClass string) string {
	switch styleClass {
	case "h1":
		return `<w:pStyle w:val="h1" />`
	case "h2":
		return `<w:pStyle w:val="h2" />`
	case "h3":
		return `<w:pStyle w:val="h3" />`
	}

	return ""
}

func (s *Style) IsEmpty() bool {
	if s.FontFamily != "" {
		return false
	}

	if s.FontSize != 0 {
		return false
	}

	if s.IsBold {
		return false
	}

	if s.IsItalic {
		return false
	}

	if s.MarginLeft != 0 {
		return false
	}

	return true
}

func (d *Document) writeSectionProperties() {
	d.Buf.WriteString("<w:sectPr>")

	if len(d.Header) != 0 {
		d.Buf.WriteString(`<w:headerReference w:type="default" r:id="rId` + HeaderID + `"/>`)
	}

	if len(d.Footer) != 0 {
		d.Buf.WriteString(`<w:footerReference w:type="default" r:id="rId` + FooterID + `"/>`)
	}

	d.Buf.WriteString(`<w:type w:val="nextPage"/>`)
	d.writePageSizes()
	d.writeMargins()
	d.Buf.WriteString(`<w:pgNumType w:fmt="decimal"/>`)
	d.Buf.WriteString(`<w:formProt w:val="false"/>`)
	d.Buf.WriteString(`<w:textDirection w:val="lrTb"/>`)
	d.Buf.WriteString(`<w:docGrid w:type="default" w:linePitch="100" w:charSpace="0"/>`)
	d.Buf.WriteString("</w:sectPr>")
}

func (d *Document) writePageSizes() {
	width := 12240
	height := 15840

	if d.PageOrientation == PageOrientationAlbum {
		height, width = width, height
	}

	d.Buf.WriteString(`<w:pgSz w:w="` + strconv.Itoa(width) + `" w:h="` + strconv.Itoa(height) + `"/>`)
}

func (d *Document) writeMargins() {
	if d.Margin == nil {
		d.Margin = &Margin{
			Top:    DocumentDefaultMargin,
			Left:   DocumentDefaultMargin,
			Right:  DocumentDefaultMargin,
			Bottom: DocumentDefaultMargin,
		}
	}

	d.Buf.WriteString(`<w:pgMar w:left="` + strconv.Itoa(d.Margin.Left) + `" w:right="` + strconv.Itoa(d.Margin.Right) + `" w:header="` + strconv.Itoa(d.Margin.Top) + `" w:top="2229" w:footer="` + strconv.Itoa(d.Margin.Bottom) + `" w:bottom="2229" w:gutter="0"/>`)
}

func (d *Document) SetList(list *List) error {
	var recursionDepth int

	if err := d.writeList(writeListArgs{
		list:           list,
		level:          0,
		recursionDepth: &recursionDepth,
	}); err != nil {
		return errors.Wrap(err, "d.writeList")
	}

	return nil
}

type writeListArgs struct {
	list           *List
	level          int
	recursionDepth *int
}

func (args *writeListArgs) Error() error {
	if args.recursionDepth == nil {
		return errors.New("no args.recursionDepth")
	}

	return nil
}

func (d *Document) writeList(args writeListArgs) error {
	if err := args.Error(); err != nil {
		return err
	}

	if *args.recursionDepth >= 1000 {
		return errors.New("infinity loop")
	}

	*args.recursionDepth++

	if args.list.LI == nil {
		return nil
	}

	for _, li := range args.list.LI {
		for index, i := range li.Items {
			switch i.(type) {
			case *Paragraph:
				if err := d.writeListP(writeListPArgs{
					index:    index,
					listType: ListBulletType,
					level:    args.level,
					item:     i,
				}); err != nil {
					return errors.Wrap(err, "d.writeListP")
				}
			case *List:
				if err := d.writeListInList(writeListInListArgs{
					item:           i,
					level:          args.level + 1,
					recursionDepth: args.recursionDepth,
					listType:       ListBulletType,
				}); err != nil {
					return errors.Wrap(err, "setListInList")
				}
			default:
				return errors.New("undefined item type")
			}
		}
	}

	return nil
}

type writeListInListArgs struct {
	item           interface{}
	level          int
	recursionDepth *int
	listType       string
}

func (d *Document) writeListInList(args writeListInListArgs) error {
	list, ok := args.item.(*List)
	if !ok {
		return errors.New("can't convert to List")
	}

	list.Type = args.listType

	err := d.writeList(writeListArgs{
		list:           list,
		level:          args.level,
		recursionDepth: args.recursionDepth,
	})
	if err != nil {
		return errors.Wrap(err, "writeList")
	}

	return nil
}

type writeListPArgs struct {
	item     interface{}
	index    int
	listType string
	level    int
}

func (d *Document) writeListP(args writeListPArgs) error {
	item, ok := args.item.(*Paragraph)
	if !ok {
		return errors.New("can't convert to Paragraph")
	}

	if args.index == 0 {
		if args.listType == "" {
			args.listType = ListBulletType
		}

		item.ListParams = &ListParams{
			Level: args.level,
			Type:  args.listType,
		}
	} else {
		item.Style.MarginLeft = 720 * (args.level + 1)
	}

	d.writeP(item)

	return nil
}

func (d *Document) writeTd(td *TD, width int) error {
	d.Buf.WriteString("<w:tc>")
	d.Buf.WriteString("<w:tcPr>")
	d.Buf.WriteString(`<w:tcW w:w="` + strconv.Itoa(width) + `" w:type="dxa" />`)
	d.Buf.WriteString("<w:tcBorders>")
	d.Buf.WriteString(`<w:top w:val="single" w:sz="2" w:space="0" w:color="000000"/>`)
	d.Buf.WriteString(`<w:left w:val="single" w:sz="2" w:space="0" w:color="000000"/>`)
	d.Buf.WriteString(`<w:right w:val="single" w:sz="2" w:space="0" w:color="000000"/>`)
	d.Buf.WriteString(`<w:bottom w:val="single" w:sz="2" w:space="0" w:color="000000"/>`)
	d.Buf.WriteString("</w:tcBorders>")
	d.Buf.WriteString("</w:tcPr>")

	for _, i := range td.Content {
		if err := d.writeContentFromInterface(i); err != nil {
			return errors.Wrap(err, "d.writeContentFromInterface")
		}
	}

	d.Buf.WriteString("</w:tc>")

	return nil
}

func (d *Document) writeContentFromInterface(content interface{}) error {
	switch content.(type) {
	case *Paragraph:
		p, ok := content.(*Paragraph)
		if !ok {
			return errors.New("can't convert to Paragraph")
		}

		d.writeP(p)
	case *List:
		list, ok := content.(*List)
		if !ok {
			return errors.New("can't convert to List")
		}

		var recursionDepth int
		if err := d.writeList(writeListArgs{
			list:           list,
			level:          0,
			recursionDepth: &recursionDepth,
		}); err != nil {
			return errors.Wrap(err, "d.writeList")
		}
	default:
		return errors.New("undefined item type")
	}

	return nil
}

func (d *Document) writeTr(tr *TR, grid []int) error {
	if len(grid) < len(tr.TD) {
		return errors.New("len of Grim less then len of TD")
	}

	if tr.TD == nil {
		return nil
	}

	d.Buf.WriteString("<w:tr>")

	for index, td := range tr.TD {
		if err := d.writeTd(td, grid[index]); err != nil {
			return nil
		}
	}

	d.Buf.WriteString("</w:tr>")

	return nil
}

func (d *Document) writeTable(t *Table) error {
	if t.TR == nil {
		return nil
	}

	d.Buf.WriteString("<w:tbl>")
	d.Buf.WriteString(getCommonStyleClass(t.StyleClass))
	d.Buf.WriteString(t.GetPropperties())
	d.Buf.WriteString(t.GetGrid())

	if err := d.writeRows(t); err != nil {
		return errors.Wrap(err, "d.writeRowsString")
	}

	d.Buf.WriteString("</w:tbl>")

	return nil
}

func (d *Document) writeRows(t *Table) error {
	if t.TR == nil {
		return nil
	}

	for _, tr := range t.TR {
		if err := d.writeTr(tr, t.Grid); err != nil {
			return errors.Wrap(err, "d.writeTr")
		}
	}

	return nil
}

func (t *Table) GetGrid() string {
	var buf bytes.Buffer

	buf.WriteString("<w:tblGrid>")

	for _, i := range t.Grid {
		buf.WriteString(`<w:gridCol w:w="` + strconv.Itoa(i) + `"/>`)
	}

	buf.WriteString("</w:tblGrid>")

	return buf.String()
}

func (t *Table) GetPropperties() string {
	var buf bytes.Buffer

	t.setCellMargin()

	buf.WriteString("<w:tblPr>")
	buf.WriteString(`<w:tblW w:w="0" type="auto" />`)
	buf.WriteString(`<w:jc w:val="left" />`)
	buf.WriteString(`<w:tblInd w:w="55" w:type="dxa" />`)
	buf.WriteString(`<w:tblLayout w:type="fixed" />`)
	buf.WriteString(`<w:tblCellMar>`)
	buf.WriteString(`<w:top w:w="` + strconv.Itoa(t.CellMargin.Top) + `" w:type="dxa" />`)
	buf.WriteString(`<w:left w:w="` + strconv.Itoa(t.CellMargin.Left) + `" w:type="dxa" />`)
	buf.WriteString(`<w:right w:w="` + strconv.Itoa(t.CellMargin.Right) + `" w:type="dxa" />`)
	buf.WriteString(`<w:bottom w:w="` + strconv.Itoa(t.CellMargin.Bottom) + `" w:type="dxa" />`)
	buf.WriteString(`</w:tblCellMar>`)
	buf.WriteString("</w:tblPr>")

	return buf.String()
}

func (t *Table) setCellMargin() {
	if t.CellMargin == nil {
		t.CellMargin = &Margin{
			Top:    TableCellDefaultMargin,
			Bottom: TableCellDefaultMargin,
			Left:   TableCellDefaultMargin,
			Right:  TableCellDefaultMargin,
		}
	}
}

func (t *Table) Error() error {
	if len(t.TR) > len(t.Grid) {
		return errors.New("len of TRs more then len of Grid")
	}

	return nil
}

func (d *Document) SetTable(table *Table) error {
	if err := table.Error(); err != nil {
		return err
	}

	if table.TR == nil {
		return nil
	}

	if err := d.writeTable(table); err != nil {
		return errors.Wrap(err, "table.String")
	}

	return nil
}

func (img *Image) Error() error {
	if img.Width == 0 {
		return errors.New("no img.Width")
	}

	return nil
}

func (img *Image) String(d *Document) (string, error) {
	if err := img.Error(); err != nil {
		return "", err
	}

	if err := img.populateSizes(); err != nil {
		return "", errors.Wrap(err, "Image.populateSizes")
	}

	img.ID = len(d.Images)
	img.RelsID = ImagesID + strconv.Itoa(img.ID)
	img.ContentType = http.DetectContentType(img.Bytes)
	img.Extension = filepath.Ext(img.FileName)
	img.FileName = filepath.Base(img.FileName)

	nameWithoutExt := img.FileName[0 : len(img.FileName)-len(img.Extension)]

	var buf bytes.Buffer
	buf.WriteString("<w:r>")
	buf.WriteString(`<w:drawing>`)
	buf.WriteString(img.getAnchor())
	buf.WriteString(img.getAlign())
	buf.WriteString(img.getMargin())
	buf.WriteString(`<wp:docPr id="` + strconv.Itoa(img.ID) + `" name="` + nameWithoutExt + `" descr=""></wp:docPr>`)
	buf.WriteString(`<wp:extent cx="` + strconv.Itoa(mmToEMU(img.Width)) + `" cy="` + strconv.Itoa(mmToEMU(img.Height)) + `"/>`)
	buf.WriteString(`<wp:cNvGraphicFramePr>`)
	buf.WriteString(`<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>`)
	buf.WriteString(`</wp:cNvGraphicFramePr>`)
	buf.WriteString(`<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">`)
	buf.WriteString(`<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">`)
	buf.WriteString(`<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">`)
	buf.WriteString(`<pic:nvPicPr>`)
	buf.WriteString(`<pic:cNvPr id="` + strconv.Itoa(img.ID) + `" name="` + nameWithoutExt + `"/>`)
	buf.WriteString(`<pic:cNvPicPr/>`)
	buf.WriteString(`</pic:nvPicPr>`)
	buf.WriteString(`<pic:blipFill>`)
	buf.WriteString(`<a:blip r:embed="` + img.RelsID + `"></a:blip>`)
	buf.WriteString(`<a:stretch>`)
	buf.WriteString(`<a:fillRect/>`)
	buf.WriteString(`</a:stretch>`)
	buf.WriteString(`</pic:blipFill>`)
	buf.WriteString(`<pic:spPr>`)
	buf.WriteString(`<a:xfrm>`)
	buf.WriteString(`<a:off x="0" y="0"/>`)
	buf.WriteString(`<a:ext cx="` + strconv.Itoa(mmToEMU(img.Width)) + `" cy="` + strconv.Itoa(mmToEMU(img.Height)) + `"/>`)
	buf.WriteString(`</a:xfrm>`)
	buf.WriteString(`<a:prstGeom rst="rect">`)
	buf.WriteString(`<a:avLst/>`)
	buf.WriteString(`</a:prstGeom>`)
	buf.WriteString(`</pic:spPr>`)
	buf.WriteString(`</pic:pic>`)
	buf.WriteString(`</a:graphicData>`)
	buf.WriteString(`</a:graphic>`)
	buf.WriteString(`</wp:anchor>`)
	buf.WriteString(`</w:drawing>`)
	buf.WriteString("</w:r>")

	d.Images = append(d.Images, img)

	return buf.String(), nil
}

func (img *Image) populateSizes() error {
	reader := bytes.NewReader(img.Bytes)
	config, _, err := image.DecodeConfig(reader)
	if err != nil {
		return errors.Wrap(err, "image.DecodeConfig")
	}

	if img.Width == 0 {
		img.Width = int(float32(img.Height) * float32(config.Width) / float32(config.Height))
	}

	if img.Height == 0 {
		img.Height = int(float32(img.Width) * float32(config.Height) / float32(config.Width))
	}

	return nil
}

func (img *Image) getAlign() string {
	if img.HorisontalAlign == "" && img.VerticalAlign == "" {
		return ""
	}

	var buf bytes.Buffer

	buf.WriteString(`<wp:positionH relativeFrom="margin">`)
	buf.WriteString(`<wp:align>` + img.HorisontalAlign + `</wp:align>`)
	buf.WriteString(`</wp:positionH>`)
	buf.WriteString(`<wp:positionV relativeFrom="margin">`)
	buf.WriteString(`<wp:align>` + img.VerticalAlign + `</wp:align>`)
	buf.WriteString(`</wp:positionV>`)

	return buf.String()
}

func (img *Image) getAnchor() string {
	isRelative := "0"
	if img.IsRelative {
		isRelative = "1"
	}

	if img.Margin == nil {
		img.Margin = &Margin{}
	}

	return `<wp:anchor behindDoc="` + isRelative + `" distT="` + strconv.Itoa(mmToEMU(img.Margin.Top)) + `" distB="` + strconv.Itoa(mmToEMU(img.Margin.Bottom)) + `" distL="` + strconv.Itoa(mmToEMU(img.Margin.Left)) + `" distR="` + strconv.Itoa(mmToEMU(img.Margin.Right)) + `" simplePos="0" locked="0" layoutInCell="0" allowOverlap="1" relativeHeight="` + strconv.Itoa(img.ZIndex) + `">`
}

func (img *Image) getMargin() string {
	if !img.WrapText {
		return ""
	}

	return `<wp:wrapSquare wrapText="largest" distT="0" distB="0" distL="0" distR="0" />`
}

func mmToEMU(mm int) int {
	return mm * 36000
	// return val
}
