package zdocx

import (
	"bytes"
	"image"
	_ "image/jpeg"
	_ "image/png"
	"net/http"
	"path/filepath"
	"strconv"

	"github.com/pkg/errors"
)

const (
	ListDecimalID          = 1
	ListBulletID           = 2
	ListNoneID             = 3
	ListDecimalType        = "decimal"
	ListBulletType         = "bullet"
	ListNoneType           = "none"
	TableCellDefaultMargin = 100
	DocumentDefaultMargin  = 1440
	PageOrientationAlbum   = "album"
	PageOrientationBook    = "book"
	StylesID               = "fileStylesID"
	ImagesID               = "fileImagesID"
	NumberingID            = "fileNumberingID"
	FontTableID            = "fileFontTableID"
	SettingsID             = "fileSettingsID"
	ThemeID                = "fileThemeID"
	HeaderID               = "fileHeaderID"
	FooterID               = "fileFooterID"
	LinkIDPrefix           = "fileLinkId"
	PageWidth              = 12240
	PageHeight             = 15840
	ImageDisplayFloat      = "float"
	ImageDisplayInline     = "inline"
)

type Document struct {
	Buf             bytes.Buffer
	Header          []*Paragraph
	Footer          []*Paragraph
	PageOrientation string
	Lang            string
	MarginTop       *Margin
	MarginLeft      *Margin
	MarginRight     *Margin
	MarginBottom    *Margin
	FontSize        int
	Images          []*Image
	Links           []*Link
	alertImage      *Image
}

type Link struct {
	URL  string
	Text string
	ID   string
}

type Image struct {
	FileName         string
	Extension        string
	ContentType      string
	Description      string
	RelsID           string
	HorisontalAnchor string
	HorisontalAlign  string
	VerticalAnchor   string
	VerticalAlign    string
	Display          string
	Width            int
	Height           int
	ZIndex           int
	IsRelative       bool
	MarginTop        *Margin
	MarginLeft       *Margin
	MarginRight      *Margin
	MarginBottom     *Margin
	ID               int
	Bytes            []byte
}

type SectionProperties struct {
}

type ListParams struct {
	Level int
	Type  string
}

type Style struct {
	FontSize        int
	IsBold          bool
	IsItalic        bool
	PageBreakBefore bool
	FontFamily      string
	Color           string
	HorisontalAlign string
	MarginTop       *Margin
	MarginLeft      *Margin
	MarginRight     *Margin
	MarginBottom    *Margin
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
	StyleClass string
	Content    []interface{}
}

type Margin struct {
	Value int
}

func (i *Margin) String() string {
	return strconv.Itoa(i.Value)
}

func (i *Margin) Int() int {
	return i.Value
}

type Table struct {
	TR          []*TR
	Grid        []int
	Type        string
	StyleClass  string
	Width       int
	CellMargin  *CellMargin
	BorderColor string
}

type CellMargin struct {
	Top    *Margin
	Left   *Margin
	Right  *Margin
	Bottom *Margin
}

func NewDocument() *Document {
	doc := Document{}
	doc.writeStartTags()
	doc.writeBody()
	doc.setMarginMaybe()
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

func (d *Document) GetInnerHeight() int {
	pageHeight := PageHeight

	if d.PageOrientation == PageOrientationAlbum {
		pageHeight = PageWidth
	}

	return pageHeight - d.MarginTop.Int() - d.MarginBottom.Int()
}

func (d *Document) GetInnerWidth() int {
	pageWidth := PageWidth

	if d.PageOrientation == PageOrientationAlbum {
		pageWidth = PageHeight
	}

	return pageWidth - d.MarginLeft.Int() - d.MarginRight.Int()
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
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:` + tag + ` xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14 wp14">`
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
	buf.WriteString(p.getStyleClass())
	buf.WriteString(p.GetListParams())
	buf.WriteString(getCommonStyle(p.Style))
	buf.WriteString("</w:pPr>")

	return buf.String()
}

func (p *Paragraph) getStyleClass() string {
	if p.ListParams != nil {
		return `<w:pStyle w:val="ListParagraph" />`
	}

	if p.StyleClass == "" {
		return `<w:pStyle w:val="Normal" />`
	}

	return `<w:pStyle w:val="` + p.StyleClass + `" />`
}

func (t *Text) getStyleClass() string {
	if t.Link != nil {
		return `<w:rStyle w:val="hyperlink" />`
	}

	if t.StyleClass == "" {
		return ""
	}

	return `<w:rStyle w:val="` + t.StyleClass + `" />`
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

	if t.Text != "" {
		buf.WriteString("<w:r>")
		buf.WriteString(t.GetProperties())
		buf.WriteString("<w:t>" + t.Text + "</w:t>")
		buf.WriteString("</w:r>")
	}

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
	buf.WriteString(t.getStyleClass())
	buf.WriteString(getCommonStyle(t.Style))
	buf.WriteString("</w:rPr>")

	return buf.String()
}

func getCommonStyle(style Style) string {
	var buf bytes.Buffer

	if style.MarginLeft != nil {
		buf.WriteString(`<w:ind w:left="` + strconv.Itoa(style.MarginLeft.Int()) + `" />`)
	}

	if style.MarginBottom != nil {
		buf.WriteString(`<w:spacing w:lineRule="auto" w:line="276" w:before="0" w:after="` + strconv.Itoa(style.MarginBottom.Int()) + `"/>`)
	}

	if style.FontFamily != "" {
		buf.WriteString(`<w:rFonts w:ascii="` + style.FontFamily + `" w:hAnsi="` + style.FontFamily + `" />`)
	}

	if style.Color != "" {
		buf.WriteString(`<w:color w:val="` + style.Color + `"/>`)
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

	if style.HorisontalAlign != "" {
		buf.WriteString(`<w:jc w:val="` + style.HorisontalAlign + `"/>`)
	}

	if style.PageBreakBefore {
		buf.WriteString(`<w:pageBreakBefore/>`)
	}

	return buf.String()
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

	if s.MarginLeft != nil {
		return false
	}

	if s.MarginBottom != nil {
		return false
	}

	if s.PageBreakBefore {
		return false
	}

	if s.HorisontalAlign != "" {
		return false
	}

	if s.Color != "" {
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
	width := PageWidth
	height := PageHeight

	if d.PageOrientation == PageOrientationAlbum {
		height, width = width, height
	}

	d.Buf.WriteString(`<w:pgSz w:w="` + strconv.Itoa(width) + `" w:h="` + strconv.Itoa(height) + `"`)
	if d.PageOrientation == PageOrientationAlbum {
		d.Buf.WriteString(` w:orient="landscape"`)
	}
	d.Buf.WriteString(` />`)
}

func (d *Document) setMarginMaybe() {
	if d.MarginTop == nil {
		d.MarginTop = &Margin{Value: DocumentDefaultMargin}
	}

	if d.MarginLeft == nil {
		d.MarginLeft = &Margin{Value: DocumentDefaultMargin}
	}

	if d.MarginRight == nil {
		d.MarginRight = &Margin{Value: DocumentDefaultMargin}
	}

	if d.MarginBottom == nil {
		d.MarginBottom = &Margin{Value: DocumentDefaultMargin}
	}
}

func (d *Document) writeMargins() {
	d.Buf.WriteString(`<w:pgMar w:left="` + d.MarginLeft.String() + `" w:right="` + d.MarginRight.String() + `" w:header="` + d.MarginTop.String() + `" w:top="2229" w:footer="` + d.MarginBottom.String() + `" w:bottom="2229" w:gutter="0"/>`)
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
		if item.Style.MarginLeft == nil {
			item.Style.MarginLeft = &Margin{}
		}

		item.Style.MarginLeft.Value = 720 * (args.level + 1)
	}

	if err := d.writeP(item); err != nil {
		return errors.Wrap(err, "Document.writeP")
	}

	return nil
}

func (d *Document) writeTd(td *TD, table *Table) error {
	if table.BorderColor == "" {
		table.BorderColor = "C0C0C0"
	}

	d.Buf.WriteString("<w:tc>")
	d.Buf.WriteString("<w:tcPr>")
	d.Buf.WriteString("<w:tcBorders>")
	d.Buf.WriteString(`<w:top w:val="single" w:sz="4" w:space="0" w:color="` + table.BorderColor + `"/>`)
	d.Buf.WriteString(`<w:left w:val="single" w:sz="4" w:space="0" w:color="` + table.BorderColor + `"/>`)
	d.Buf.WriteString(`<w:bottom w:val="single" w:sz="4" w:space="0" w:color="` + table.BorderColor + `"/>`)
	d.Buf.WriteString(`<w:right w:val="single" w:sz="4" w:space="0" w:color="` + table.BorderColor + `"/>`)
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

func (d *Document) writeContextualSpacing() {
	d.Buf.WriteString(`<w:p>`)
	d.Buf.WriteString(`<w:pPr>`)
	d.Buf.WriteString(`<w:spacing w:before="0" w:after="200"/>`)
	d.Buf.WriteString(`<w:ind w:hanging="0"/>`)
	d.Buf.WriteString(`<w:contextualSpacing/>`)
	d.Buf.WriteString(`</w:pPr>`)
	d.Buf.WriteString(`</w:p>`)
}

func (d *Document) writeContentFromInterface(content interface{}) error {
	switch content.(type) {
	case *Paragraph:
		p, ok := content.(*Paragraph)
		if !ok {
			return errors.New("can't convert to Paragraph")
		}

		if err := d.writeP(p); err != nil {
			return errors.Wrap(err, "Document.writeP")
		}
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

func (d *Document) writeTr(tr *TR, table *Table) error {
	if tr.TD == nil {
		return nil
	}

	d.Buf.WriteString("<w:tr>")

	for _, td := range tr.TD {
		if err := d.writeTd(td, table); err != nil {
			return err
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
		if err := d.writeTr(tr, t); err != nil {
			return errors.Wrap(err, "d.writeTr")
		}
	}

	return nil
}

func (t *Table) GetGrid() string {
	if t.getType() == "autofit" {
		return ""
	}

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

	t.setCellMarginMaybe()

	buf.WriteString("<w:tblPr>")
	buf.WriteString(t.getStyleClass())
	buf.WriteString(t.getWidth())
	buf.WriteString(`<w:jc w:val="center" />`)
	buf.WriteString(`<w:tblInd w:type="dxa" w:w="0" />`)
	buf.WriteString(`<w:tblLayout w:type="` + t.getType() + `" />`)
	buf.WriteString(`<w:tblCellMar>`)
	buf.WriteString(`<w:top w:w="` + strconv.Itoa(t.CellMargin.Top.Int()) + `" w:type="dxa" />`)
	buf.WriteString(`<w:left w:w="` + strconv.Itoa(t.CellMargin.Left.Int()) + `" w:type="dxa" />`)
	buf.WriteString(`<w:bottom w:w="` + strconv.Itoa(t.CellMargin.Bottom.Int()) + `" w:type="dxa" />`)
	buf.WriteString(`<w:right w:w="` + strconv.Itoa(t.CellMargin.Right.Int()) + `" w:type="dxa" />`)
	buf.WriteString(`</w:tblCellMar>`)
	buf.WriteString("</w:tblPr>")

	return buf.String()
}

func (t *Table) getWidth() string {
	if t.Width == 0 {
		return ""
	}

	return `<w:tblW w:type="dxa" w:w="` + strconv.Itoa(t.Width) + `"/>`
}

func (t *Table) getStyleClass() string {
	if t.StyleClass == "" {
		return `<w:tblStyle w:val="NormalTable"/>`
	}

	return `<w:tblStyle w:val="` + t.StyleClass + `"/>`
}

func (t *Table) getType() string {
	switch t.Type {
	case "fixed":
		return "fixed"
	case "autofit":
		return "autofit"
	default:
		return "autofit"
	}
}

func (t *Table) setCellMarginMaybe() {
	if t.CellMargin == nil {
		t.CellMargin = &CellMargin{
			Top:    &Margin{Value: TableCellDefaultMargin},
			Bottom: &Margin{Value: TableCellDefaultMargin},
			Left:   &Margin{Value: TableCellDefaultMargin},
			Right:  &Margin{Value: TableCellDefaultMargin},
		}
	}
}

func (t *Table) Error() error {
	if t.getType() == "autofit" {
		return nil
	}

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

	d.writeContextualSpacing()

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

	img.ID = len(d.Images) + 1
	img.RelsID = ImagesID + strconv.Itoa(img.ID)
	img.ContentType = http.DetectContentType(img.Bytes)
	img.FileName = filepath.Base(img.FileName)
	img.Extension = filepath.Ext(img.FileName)

	nameWithoutExt := img.FileName[0 : len(img.FileName)-len(img.Extension)]

	var buf bytes.Buffer
	buf.WriteString("<w:r>")
	buf.WriteString(`<w:drawing>`)
	buf.WriteString(img.getDisplayTag())
	buf.WriteString(img.getAlign())
	buf.WriteString(img.getWrap())
	buf.WriteString(`<wp:extent cx="` + strconv.Itoa(mmToEMU(img.Width)) + `" cy="` + strconv.Itoa(mmToEMU(img.Height)) + `"/>`)
	buf.WriteString(`<wp:effectExtent l="0" t="0" r="0" b="0"/>`)
	buf.WriteString(`<wp:docPr id="` + strconv.Itoa(img.ID) + `" name="` + nameWithoutExt + `" descr=""></wp:docPr>`)
	buf.WriteString(`<wp:cNvGraphicFramePr>`)
	buf.WriteString(`<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>`)
	buf.WriteString(`</wp:cNvGraphicFramePr>`)
	buf.WriteString(`<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">`)
	buf.WriteString(`<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">`)
	buf.WriteString(`<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">`)
	buf.WriteString(`<pic:nvPicPr>`)
	buf.WriteString(`<pic:cNvPr id="` + strconv.Itoa(img.ID) + `" name="` + nameWithoutExt + `" descr=""></pic:cNvPr>`)
	buf.WriteString(`<pic:cNvPicPr>`)
	buf.WriteString(`<a:picLocks noChangeAspect="1" noChangeArrowheads="1"/>`)
	buf.WriteString(`</pic:cNvPicPr>`)
	buf.WriteString(`</pic:nvPicPr>`)
	buf.WriteString(`<pic:blipFill>`)
	buf.WriteString(`<a:blip r:embed="` + img.RelsID + `"/>`)
	buf.WriteString(`<a:stretch>`)
	buf.WriteString(`<a:fillRect />`)
	buf.WriteString(`</a:stretch>`)
	buf.WriteString(`</pic:blipFill>`)
	buf.WriteString(`<pic:spPr bwMode="auto">`)
	buf.WriteString(`<a:xfrm>`)
	buf.WriteString(`<a:off x="0" y="0" />`)
	buf.WriteString(`<a:ext cx="` + strconv.Itoa(mmToEMU(img.Width)) + `" cy="` + strconv.Itoa(mmToEMU(img.Height)) + `" />`)
	buf.WriteString(`</a:xfrm>`)
	buf.WriteString(`<a:prstGeom prst="rect">`)
	buf.WriteString(`<a:avLst/>`)
	buf.WriteString(`</a:prstGeom>`)
	buf.WriteString(`</pic:spPr>`)
	buf.WriteString(`</pic:pic>`)
	buf.WriteString(`</a:graphicData>`)
	buf.WriteString(`</a:graphic>`)
	buf.WriteString(img.getDislayCloseTag())
	buf.WriteString(`</w:drawing>`)
	buf.WriteString("</w:r>")

	d.Images = append(d.Images, img)

	return buf.String(), nil
}

func (img *Image) getDislayCloseTag() string {
	if img.Display == ImageDisplayFloat {
		return `</wp:anchor>`
	}

	return `</wp:inline>`
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

	if img.HorisontalAlign != "" {
		buf.WriteString(`<wp:positionH relativeFrom="` + img.getHorisontalAnchor() + `">`)
		buf.WriteString(`<wp:align>` + img.HorisontalAlign + `</wp:align>`)
		buf.WriteString(`</wp:positionH>`)
	}

	if img.VerticalAlign != "" {
		buf.WriteString(`<wp:positionV relativeFrom="` + img.getVerticalAnchor() + `">`)
		buf.WriteString(`<wp:align>` + img.VerticalAlign + `</wp:align>`)
		buf.WriteString(`</wp:positionV>`)
	}

	return buf.String()
}

func (img *Image) getVerticalAnchor() string {
	switch img.VerticalAnchor {
	case "":
		return "paragraph"
	case "bottomMargin":
		return "bottomMargin"
	case "insideMargin":
		return "insideMargin"
	case "line":
		return "line"
	case "margin":
		return "margin"
	case "ousideMargin":
		return "ousideMargin"
	case "page":
		return "page"
	case "paragraph":
		return "paragraph"
	case "topMargin":
		return "topMargin"
	default:
		return "paragraph"
	}
}

func (img *Image) getHorisontalAnchor() string {
	switch img.HorisontalAnchor {
	case "":
		return "column"
	case "character":
		return "character"
	case "column":
		return "column"
	case "insideMargin":
		return "insideMargin"
	case "margin":
		return "margin"
	case "outsideMargin":
		return "outsideMargin"
	case "page":
		return "page"
	case "rightMargin":
		return "rightMargin"
	case "leftMargin":
		return "leftMargin"
	default:
		return "column"
	}
}

func (img *Image) setMarginMaybe() {
	if img.MarginTop == nil {
		img.MarginTop = &Margin{Value: 0}
	}

	if img.MarginLeft == nil {
		img.MarginLeft = &Margin{Value: 0}
	}

	if img.MarginRight == nil {
		img.MarginRight = &Margin{Value: 0}
	}

	if img.MarginBottom == nil {
		img.MarginBottom = &Margin{Value: 0}
	}
}

func (img *Image) getDisplayTag() string {
	img.setMarginMaybe()

	if img.Display == ImageDisplayFloat {
		isRelative := "0"
		if img.IsRelative {
			isRelative = "1"
		}

		return `<wp:anchor behindDoc="` + isRelative + `" distT="` + strconv.Itoa(mmToEMU(img.MarginTop.Int())) + `" distB="` + strconv.Itoa(mmToEMU(img.MarginBottom.Int())) + `" distL="` + strconv.Itoa(mmToEMU(img.MarginLeft.Int())) + `" distR="` + strconv.Itoa(mmToEMU(img.MarginRight.Int())) + `" simplePos="0" locked="1" layoutInCell="0" allowOverlap="1" relativeHeight="` + strconv.Itoa(img.ZIndex) + `">`
	}

	return `<wp:inline distT="` + strconv.Itoa(mmToEMU(img.MarginTop.Int())) + `" distB="` + strconv.Itoa(mmToEMU(img.MarginBottom.Int())) + `" distL="` + strconv.Itoa(mmToEMU(img.MarginLeft.Int())) + `" distR="` + strconv.Itoa(mmToEMU(img.MarginRight.Int())) + `">`
}

func (img *Image) getWrap() string {
	if img.Display == ImageDisplayFloat {
		return `<wp:wrapSquare wrapText="largest" distT="0" distB="0" distL="0" distR="0" />`
	}

	return ""
}

func mmToEMU(mm int) int {
	return mm * 36000
}

func (d *Document) SetPageBreak() {
	d.Buf.WriteString(`<w:p><w:r><w:br w:type="page"/></w:r></w:p>`)
}
