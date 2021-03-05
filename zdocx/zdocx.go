package zdocx

import (
	"bytes"
	"encoding/xml"
	"fmt"
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
	stylesID               = "fileStylesID"
	imagesID               = "fileImagesID"
	numberingID            = "fileNumberingID"
	fontTableID            = "fileFontTableID"
	settingsID             = "fileSettingsID"
	themeID                = "fileThemeID"
	defaultHeaderID        = "defaultHeaderID"
	mainPageHeaderID       = "mainPageHeaderID"
	defaultFooterID        = "defaultFooterID"
	mainPageFooterID       = "mainPageFooterID"
	linkIDPrefix           = "fileLinkId"
	PageWidth              = 12240
	PageHeight             = 15840
	ImageDisplayFloat      = "float"
	ImageDisplayInline     = "inline"
	HorisontalAlignLeft    = "left"
	HorisontalAlignRight   = "right"
	HorisontalAlignCenter  = "center"
	BorderSingleLine       = "single"
	BorderDotted           = "dotted"
	BorderDashed           = "dashed"
	BorderDashSmallGap     = "dashSmallGap"
	SectionTypeContinious  = "continuous"
	SectionTypeEvenPage    = "evenPage"
	SectionTypeNextColumn  = "nextColumn"
	SectionTypeNextPage    = "nextPage"
	SectionTypeOddPage     = "oddPage"
)

type Document struct {
	Buf             bytes.Buffer
	Header          []*Paragraph
	MainPageHeader  []*Paragraph
	Footer          []*Paragraph
	MainPageFooter  []*Paragraph
	PageOrientation string
	Lang            string
	Margins         Margins
	FontSize        int
	images          images
	Links           []*Link
	alertImage      *Image
}

type images struct {
	content        []*Image
	mainPageHeader []*Image
	mainPageFooter []*Image
	header         []*Image
	footer         []*Image
}

type Link struct {
	URL string
	ID  string
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
	Width            int64
	Height           int64
	ZIndex           int
	IsRelative       bool
	IsBackground     bool
	MarginTop        *Margin
	MarginLeft       *Margin
	MarginRight      *Margin
	MarginBottom     *Margin
	ID               int
	Bytes            []byte
	isMainPageHeader bool
	isMainPageFooter bool
	isHeader         bool
	isFooter         bool
}

type ListParams struct {
	Level int
	Type  string
}

type TableStyle struct {
	Margins         Margins
	Borders         Borders
	Background      string
	HorisontalAlign string
	Color           string
	FontSize        int
}

type Margins struct {
	Top    *Margin
	Left   *Margin
	Bottom *Margin
	Right  *Margin
}

type Borders struct {
	Top    Border
	Left   Border
	Right  Border
	Bottom Border
}

type Border struct {
	Width int
	Color string
	Type  string
}

type Text struct {
	Text       string
	Link       *Link
	Image      *Image
	StyleClass string
	Style      TextStyle
}

type Paragraph struct {
	Texts      []*Text
	ListParams *ListParams
	StyleClass string
	Style      PStyle

	isPagination bool
}

type PStyle struct {
	PageBreakBefore bool
	HorisontalAlign string
	Margins         Margins
	Borders         Borders
	Background      string
	Color           string
	FontSize        int
	LineHeight      int
}

type List struct {
	LI         []*LI
	Type       string
	StyleClass string
	Style      PStyle
}

type LI struct {
	Items []interface{}
}

type TR struct {
	TD        []*TD
	IsHeader  bool
	CantSplit bool
	Height    int
}

type TD struct {
	GridSpan   int
	StyleClass string
	Style      TDStyle
	Content    []interface{}
}

type TDStyle struct {
	Margins    Margins
	Borders    Borders
	Background string
	Color      string
	HideMark   bool
	FontSize   int
	Width      int
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

func (i *Margin) emuString() string {
	return strconv.FormatInt(dxaToEMU(int64(i.Value)), 10)
}

func (m *Margins) IsEmpty() bool {
	if m.Top != nil {
		return false
	}

	if m.Left != nil {
		return false
	}

	if m.Bottom != nil {
		return false
	}

	if m.Right != nil {
		return false
	}

	return true
}

func (m *Margins) SetValueByDefault(val int) {
	if m.Top == nil {
		m.Top = &Margin{Value: val}
	}

	if m.Left == nil {
		m.Left = &Margin{Value: val}
	}

	if m.Bottom == nil {
		m.Bottom = &Margin{Value: val}
	}

	if m.Right == nil {
		m.Right = &Margin{Value: val}
	}
}

type Table struct {
	TR             []*TR
	Grid           []int
	Type           string
	StyleClass     string
	Width          int
	CellMargin     *CellMargin
	Style          TableStyle
	NoMarginBottom bool
}

type CellMargin struct {
	Top    *Margin
	Left   *Margin
	Right  *Margin
	Bottom *Margin
}

type NewDocumentArgs struct {
	Margins *Margins
}

func NewDocument(args NewDocumentArgs) *Document {
	doc := Document{}

	if args.Margins != nil {
		doc.SetMargins(args.Margins)
	}

	doc.writeStartTags()
	doc.writeBody()
	doc.setMarginMaybe()
	return &doc
}

func (doc *Document) SetMargins(margins *Margins) {
	if margins == nil {
		return
	}

	if margins.Top != nil {
		doc.Margins.Top = margins.Top
	}

	if margins.Left != nil {
		doc.Margins.Left = margins.Left
	}

	if margins.Bottom != nil {
		doc.Margins.Bottom = margins.Bottom
	}

	if margins.Right != nil {
		doc.Margins.Right = margins.Right
	}
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

	return pageHeight - d.Margins.Top.Int() - d.Margins.Bottom.Int()
}

func (d *Document) GetInnerWidth() int {
	pageWidth := PageWidth

	if d.PageOrientation == PageOrientationAlbum {
		pageWidth = PageHeight
	}

	return pageWidth - d.Margins.Left.Int() - d.Margins.Right.Int()
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
	pString, err := p.string(d)
	if err != nil {
		return errors.Wrap(err, "p.string")
	}

	d.Buf.WriteString(pString)

	return nil
}

func pagination() string {
	var buf bytes.Buffer
	buf.WriteString(`<w:p>`)
	buf.WriteString(`<w:pPr>`)
	buf.WriteString(`<w:widowControl w:val="false"/>`)
	buf.WriteString(`<w:suppressLineNumbers/>`)
	buf.WriteString(`<w:bidi w:val="0"/>`)
	buf.WriteString(`<w:spacing w:before="0" w:after="0"/>`)
	buf.WriteString(`<w:jc w:val="right"/>`)
	buf.WriteString(`</w:pPr>`)
	buf.WriteString(`<w:r>`)
	buf.WriteString(`<w:fldChar w:fldCharType="begin"></w:fldChar>`)
	buf.WriteString(`</w:r>`)
	buf.WriteString(`<w:r>`)
	buf.WriteString(`<w:instrText> PAGE </w:instrText>`)
	buf.WriteString(`</w:r>`)
	buf.WriteString(`<w:r>`)
	buf.WriteString(`<w:fldChar w:fldCharType="separate"/>`)
	buf.WriteString(`</w:r>`)
	buf.WriteString(`<w:r>`)
	buf.WriteString(`<w:fldChar w:fldCharType="end"/>`)
	buf.WriteString(`</w:r>`)
	buf.WriteString(`</w:p>`)

	return buf.String()
}

func (p *Paragraph) string(d *Document) (string, error) {
	if p.isPagination {
		return pagination(), nil
	}

	var buf bytes.Buffer

	buf.WriteString("<w:p>")
	buf.WriteString(p.properties())

	for index, t := range p.Texts {
		if index != 0 {
			buf.WriteString(getSpace())
		}

		if t.Style.Color == "" {
			t.Style.Color = p.Style.Color
		}

		if t.Style.FontSize == 0 {
			t.Style.FontSize = p.Style.FontSize
		}

		textString, err := t.string(d)
		if err != nil {
			return "", errors.Wrap(err, "Text.string")
		}

		buf.WriteString(textString)
	}

	buf.WriteString("</w:p>")

	return buf.String(), nil
}

func (p *Paragraph) getListParams() string {
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

func (p *Paragraph) properties() string {
	var buf bytes.Buffer

	buf.WriteString("<w:pPr>")
	buf.WriteString(p.getStyleClass())
	buf.WriteString(p.getListParams())
	buf.WriteString(p.getStyles())
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

func (p *Paragraph) getStyles() string {
	var buf bytes.Buffer

	if !p.Style.Borders.isEmpty() {
		buf.WriteString(`<w:pBdr>`)
		buf.WriteString(p.border("top", &p.Style.Borders.Top))
		buf.WriteString(p.border("left", &p.Style.Borders.Left))
		buf.WriteString(p.border("bottom", &p.Style.Borders.Bottom))
		buf.WriteString(p.border("right", &p.Style.Borders.Right))
		buf.WriteString(`</w:pBdr>`)
	}

	if !p.Style.Margins.IsEmpty() {
		if p.Style.Margins.Top != nil || p.Style.Margins.Bottom != nil {
			if p.Style.Margins.Top == nil {
				p.Style.Margins.Top = &Margin{}
			}

			if p.Style.Margins.Bottom == nil {
				p.Style.Margins.Bottom = &Margin{}
			}

			line := 240

			if p.Style.LineHeight != 0 {
				line = p.Style.LineHeight
			}

			buf.WriteString(`<w:spacing w:before="` + p.Style.Margins.Top.String() + `" w:after="` + p.Style.Margins.Bottom.String() + `" w:lineRule="auto" w:line="` + strconv.Itoa(line) + `"/>`)

		}

		if p.ListParams == nil {
			p.Style.Margins.SetValueByDefault(0)

			buf.WriteString(`<w:ind w:left="` + p.Style.Margins.Left.String() + `" w:right="` + p.Style.Margins.Right.String() + `"/>`)
		}
	}

	if p.Style.HorisontalAlign != "" {
		buf.WriteString(`<w:jc w:val="` + p.Style.HorisontalAlign + `"/>`)
	}

	if p.Style.PageBreakBefore {
		buf.WriteString(`<w:pageBreakBefore/>`)
	}

	if p.Style.Background != "" {
		buf.WriteString(`<w:shd w:val="clear" w:color="auto" w:fill="` + p.Style.Background + `"/>`)
	}

	return buf.String()
}

func (t *Text) styleClass() string {
	if t.Link != nil {
		return `<w:rStyle w:val="hyperlink" />`
	}

	if t.StyleClass == "" {
		return ""
	}

	return `<w:rStyle w:val="` + t.StyleClass + `" />`
}

func (p *Paragraph) border(tagName string, border *Border) string {
	if border.isEmpty() {
		return ""
	}

	if border.Type == "" {
		border.Type = BorderSingleLine
	}

	if border.Width == 0 {
		border.Type = "none"
	}

	if border.Color == "" {
		border.Color = "C0C0C0"
	}

	return `<w:` + tagName + ` w:val="` + border.Type + `" w:sz="` + strconv.Itoa(border.Width) + `" w:space="0" w:color="` + border.Color + `"/>`
}

func (t *Text) string(d *Document) (string, error) {
	if t == nil {
		return "", nil
	}

	if t.Text == "" && t.Image == nil {
		return "", nil
	}

	var buf bytes.Buffer
	if t.Link != nil {
		t.Link.ID = linkIDPrefix + strconv.Itoa(len(d.Links))

		var linkBuf bytes.Buffer

		if err := xml.EscapeText(&linkBuf, []byte(t.Link.URL)); err != nil {
			return "", errors.Wrap(err, "xml.EscapeText")
		}

		t.Link.URL = linkBuf.String()

		d.Links = append(d.Links, t.Link)
		buf.WriteString(`<w:hyperlink r:id="` + t.Link.ID + `">`)
	}

	if t.Image != nil {
		imageString, err := t.Image.string(d)
		if err != nil {
			return "", errors.Wrap(err, "t.Image.string")
		}

		buf.WriteString(imageString)
	}

	if t.Text != "" {
		buf.WriteString("<w:r>")
		buf.WriteString(t.properties())
		buf.WriteString("<w:t")

		if t.Style.SpacePreserve {
			buf.WriteString(` xml:space="preserve"`)
		}

		buf.WriteString(">")

		if err := xml.EscapeText(&buf, []byte(t.Text)); err != nil {
			return "", errors.Wrap(err, "xml.EscapeText")
		}

		buf.WriteString("</w:t>")
		buf.WriteString("</w:r>")
	}

	if t.Link != nil {
		buf.WriteString("</w:hyperlink>")
	}

	return buf.String(), nil
}

func (t *Text) properties() string {
	var buf bytes.Buffer

	buf.WriteString("<w:rPr>")
	buf.WriteString(t.styleClass())
	buf.WriteString(t.styles())
	buf.WriteString("</w:rPr>")

	return buf.String()
}

func (t *Text) styles() string {
	var buf bytes.Buffer

	if t.Style.SuppressLineNumbers {
		buf.WriteString(`<w:widowControl w:val="false"/>`)
		buf.WriteString(`<w:suppressLineNumbers/>`)
	}

	if t.Style.IsBold {
		buf.WriteString("<w:b/>")
	}

	if t.Style.IsItalic {
		buf.WriteString("<w:i/>")
	}

	if t.Style.Color != "" {
		buf.WriteString(`<w:color w:val="` + t.Style.Color + `"/>`)
	}

	if t.Style.FontFamily != "" {
		buf.WriteString(`<w:rFonts w:ascii="` + t.Style.FontFamily + `" w:hAnsi="` + t.Style.FontFamily + `" />`)
	}

	if t.Style.FontSize != 0 {
		buf.WriteString(`<w:sz w:val="` + strconv.Itoa(t.Style.FontSize) + `"/>`)
	}

	if t.Style.Border != nil {
		borderType := "single"

		if t.Style.Border.Type != "" {
			borderType = t.Style.Border.Type
		}

		buf.WriteString(`<w:bdr w:val="` + borderType + `" w:sz="` + strconv.Itoa(t.Style.Border.Width) + `" w:space="0" w:color="` + t.Style.Border.Color + `" />`)
	}

	return buf.String()
}

type TextStyle struct {
	IsBold              bool
	IsItalic            bool
	SuppressLineNumbers bool
	SpacePreserve       bool
	Color               string
	FontFamily          string
	FontSize            int
	Border              *Border
}

func (d *Document) writeSectionProperties() {
	d.Buf.WriteString("<w:sectPr>")

	if d.Header != nil {
		d.Buf.WriteString(`<w:headerReference w:type="default" r:id="rId` + defaultHeaderID + `"/>`)
	}

	if d.MainPageHeader != nil {
		d.Buf.WriteString(`<w:headerReference w:type="first" r:id="rId` + mainPageHeaderID + `"/>`)
	}

	if d.Footer != nil {
		d.Buf.WriteString(`<w:footerReference w:type="default" r:id="rId` + defaultFooterID + `"/>`)
	}

	if d.MainPageFooter != nil {
		d.Buf.WriteString(`<w:footerReference w:type="first" r:id="rId` + mainPageFooterID + `"/>`)
	}

	d.Buf.WriteString(`<w:type w:val="nextPage"/>`)
	d.writePageSizes()
	d.writeMargins()
	d.Buf.WriteString(`<w:pgNumType w:fmt="decimal"/>`)
	d.Buf.WriteString(`<w:formProt w:val="false"/>`)

	if d.MainPageFooter != nil || d.MainPageHeader != nil {
		d.Buf.WriteString(`<w:titlePg/>`)
	}

	d.Buf.WriteString(`<w:textDirection w:val="lrTb"/>`)
	d.Buf.WriteString(`<w:docGrid w:type="default" w:linePitch="100" w:charSpace="0"/>`)

	d.Buf.WriteString("</w:sectPr>")
}

func (d *Document) writePageSizes() {
	d.Buf.WriteString(sectionSizes(d.PageOrientation))
}

func sectionSizes(orientation string) string {
	width := PageWidth
	height := PageHeight

	if orientation == PageOrientationAlbum {
		height, width = width, height
	}

	var buf bytes.Buffer
	buf.WriteString(`<w:pgSz w:w="` + strconv.Itoa(width) + `" w:h="` + strconv.Itoa(height) + `"`)
	if orientation == PageOrientationAlbum {
		buf.WriteString(` w:orient="landscape"`)
	}
	buf.WriteString(` />`)

	return buf.String()
}

func (d *Document) setMarginMaybe() {
	if d.Margins.Top == nil {
		d.Margins.Top = &Margin{Value: DocumentDefaultMargin}
	}

	if d.Margins.Left == nil {
		d.Margins.Left = &Margin{Value: DocumentDefaultMargin}
	}

	if d.Margins.Right == nil {
		d.Margins.Right = &Margin{Value: DocumentDefaultMargin}
	}

	if d.Margins.Bottom == nil {
		d.Margins.Bottom = &Margin{Value: DocumentDefaultMargin}
	}
}

func (d *Document) writeMargins() {
	d.Buf.WriteString(sectionMargins(d.Margins))
}

func sectionMargins(margins Margins) string {
	return `<w:pgMar w:left="` + margins.Left.String() + `" w:right="` + margins.Right.String() + `" w:header="` + margins.Top.String() + `" w:top="` + margins.Top.String() + `" w:footer="` + margins.Bottom.String() + `" w:bottom="` + margins.Bottom.String() + `" w:gutter="0"/>`
}

func (d *Document) SetList(list *List) error {
	var recursionDepth int

	listString, err := list.string(listStringArgs{
		level:          0,
		documnet:       d,
		recursionDepth: &recursionDepth,
	})
	if err != nil {
		return errors.Wrap(err, "list.string")
	}

	d.Buf.WriteString(listString)

	return nil
}

type listStringArgs struct {
	level          int
	recursionDepth *int
	documnet       *Document
}

func (args *listStringArgs) error() error {
	if args.recursionDepth == nil {
		return errors.New("no args.recursionDepth")
	}

	if args.documnet == nil {
		return errors.New("no args.document")
	}

	return nil
}

func (list *List) string(args listStringArgs) (string, error) {
	if err := args.error(); err != nil {
		return "", err
	}

	if *args.recursionDepth >= 1000 {
		return "", errors.New("infinity loop")
	}

	*args.recursionDepth++

	if list.LI == nil {
		return "", nil
	}

	var buf bytes.Buffer

	for _, li := range list.LI {
		for index, i := range li.Items {
			switch i.(type) {
			case *Paragraph:
				pString, err := listPString(listPStringArgs{
					index:    index,
					listType: ListBulletType,
					level:    args.level,
					item:     i,
					style:    list.Style,
					document: args.documnet,
				})
				if err != nil {
					return "", errors.Wrap(err, "listPString")
				}

				buf.WriteString(pString)

			case *List:
				listString, err := listInListString(listInListStringArgs{
					item:           i,
					level:          args.level + 1,
					recursionDepth: args.recursionDepth,
					listType:       ListBulletType,
					document:       args.documnet,
				})
				if err != nil {
					return "", errors.Wrap(err, "setListInList")
				}

				buf.WriteString(listString)

			default:
				return "", errors.New("undefined item type")
			}
		}
	}

	return buf.String(), nil
}

type listInListStringArgs struct {
	item           interface{}
	level          int
	recursionDepth *int
	listType       string
	document       *Document
}

func (args *listInListStringArgs) error() error {
	if args.document == nil {
		return errors.New("no args.document")
	}

	return nil
}

func listInListString(args listInListStringArgs) (string, error) {
	if err := args.error(); err != nil {
		return "", err
	}

	list, ok := args.item.(*List)
	if !ok {
		return "", errors.New("can't convert to List")
	}

	list.Type = args.listType

	listString, err := list.string(listStringArgs{
		level:          args.level,
		recursionDepth: args.recursionDepth,
		documnet:       args.document,
	})
	if err != nil {
		return "", errors.Wrap(err, "list.string")
	}

	return listString, nil
}

type listPStringArgs struct {
	item     interface{}
	index    int
	listType string
	level    int
	style    PStyle
	document *Document
}

func (args *listPStringArgs) error() error {
	if args.document == nil {
		return errors.New("no args.document")
	}

	return nil
}

func listPString(args listPStringArgs) (string, error) {
	if err := args.error(); err != nil {
		return "", err
	}

	item, ok := args.item.(*Paragraph)
	if !ok {
		return "", errors.New("can't convert to Paragraph")
	}

	item.Style.Color = args.style.Color

	if args.index == 0 {
		if args.listType == "" {
			args.listType = ListBulletType
		}

		item.ListParams = &ListParams{
			Level: args.level,
			Type:  args.listType,
		}
	} else {
		if item.Style.Margins.Left == nil {
			item.Style.Margins.Left = &Margin{}
		}

		item.Style.Margins.Left.Value = 720 * (args.level + 1)
	}

	pString, err := item.string(args.document)
	if err != nil {
		return "", errors.Wrap(err, "Document.writeP")
	}

	return pString, nil
}

func (td *TD) border(tagName string, border Border) string {
	if border.Type == "" {
		border.Type = BorderSingleLine
	}

	if border.Width == 0 {
		return ""
	}

	if border.Color == "" {
		border.Color = "C0C0C0"
	}

	return `<w:` + tagName + ` w:val="` + border.Type + `" w:sz="` + strconv.Itoa(border.Width) + `" w:space="0" w:color="` + border.Color + `"/>`
}

type tdBytesArgs struct {
	document *Document
}

func (args *tdBytesArgs) error() error {
	if args.document == nil {
		return errors.New("no args.document")
	}

	return nil
}

func (td *TD) string(args tdBytesArgs) (string, error) {
	if err := args.error(); err != nil {
		return "", err
	}

	var buf bytes.Buffer
	buf.WriteString("<w:tc>")
	buf.WriteString("<w:tcPr>")

	if td.GridSpan > 0 {
		buf.WriteString(`<w:gridSpan w:val="` + strconv.Itoa(td.GridSpan) + `"/>`)
	}

	if td.Style.HideMark {
		buf.WriteString(`<w:hideMark/>`)
	}

	if td.Style.Width != 0 {
		buf.WriteString(`<w:tcW w:type="dxa" w:w="` + strconv.Itoa(td.Style.Width) + `"/>`)
	}

	buf.WriteString("<w:tcBorders>")
	buf.WriteString(td.border("top", td.Style.Borders.Top))
	buf.WriteString(td.border("left", td.Style.Borders.Left))
	buf.WriteString(td.border("bottom", td.Style.Borders.Bottom))
	buf.WriteString(td.border("right", td.Style.Borders.Right))
	buf.WriteString("</w:tcBorders>")

	if td.Style.Background != "" {
		buf.WriteString(`<w:shd w:val="clear" w:color="auto" w:fill="` + td.Style.Background + `"/>`)
	}

	if !td.Style.Margins.IsEmpty() {
		buf.WriteString(`<w:tcMar>`)

		if td.Style.Margins.Top != nil {
			buf.WriteString(`<w:top w:w="` + td.Style.Margins.Top.String() + `" w:type="dxa"/>`)
		}

		if td.Style.Margins.Left != nil {
			buf.WriteString(`<w:left w:w="` + td.Style.Margins.Left.String() + `" w:type="dxa"/>`)
		}

		if td.Style.Margins.Bottom != nil {
			buf.WriteString(`<w:bottom w:w="` + td.Style.Margins.Bottom.String() + `" w:type="dxa"/>`)
		}

		if td.Style.Margins.Right != nil {
			buf.WriteString(`<w:right w:w="` + td.Style.Margins.Right.String() + `" w:type="dxa"/>`)
		}

		buf.WriteString(`</w:tcMar>`)
	}

	buf.WriteString("</w:tcPr>")

	for _, content := range td.Content {
		content, err := contentFromInterface(contentFromInterfaceArgs{
			content:  content,
			document: args.document,
			color:    td.Style.Color,
			fontSize: td.Style.FontSize,
		})
		if err != nil {
			return "", errors.Wrap(err, "contentFromInterface")
		}

		buf.WriteString(content)
	}

	buf.WriteString("</w:tc>")

	return buf.String(), nil
}

func contextualSpacing(hidden bool) string {
	var buf bytes.Buffer
	buf.WriteString(`<w:p>`)
	buf.WriteString(`<w:pPr>`)
	if hidden {
		buf.WriteString(`<w:widowControl w:val="false"/>`)
		buf.WriteString(`<w:suppressLineNumbers/>`)
		buf.WriteString(`<w:shd w:val="clear" w:color="auto" w:fill="ffffff"/>`)
		buf.WriteString(`<w:bidi w:val="0"/>`)
		buf.WriteString(`<w:spacing w:before="0" w:after="0" w:line="6" w:lineRule="auto" w:beforeAutospacing="0" w:afterAutospacing="0"/>`)
		buf.WriteString(`<w:rPr>`)
		buf.WriteString(`<w:sz w:val="12"/>`)
		buf.WriteString(`<w:szCs w:val="12"/>`)
		buf.WriteString(`</w:rPr>`)
	} else {
		buf.WriteString(`<w:spacing w:before="0" w:after="100"/>`)
		buf.WriteString(`<w:ind w:hanging="0"/>`)
		buf.WriteString(`<w:contextualSpacing/>`)
	}
	buf.WriteString(`</w:pPr>`)
	buf.WriteString(`<w:r>`)
	buf.WriteString(`<w:rPr></w:rPr>`)
	buf.WriteString(`</w:r>`)
	buf.WriteString(`</w:p>`)

	return buf.String()
}

func (d *Document) writeContextualSpacing(hidden bool) {
	d.Buf.WriteString(contextualSpacing(hidden))
}

type contentFromInterfaceArgs struct {
	content  interface{}
	document *Document
	color    string
	fontSize int
}

func (args *contentFromInterfaceArgs) error() error {
	if args.document == nil {
		return errors.New("no args.document")
	}

	return nil
}

func contentFromInterface(args contentFromInterfaceArgs) (string, error) {
	if err := args.error(); err != nil {
		return "", err
	}
	switch args.content.(type) {
	case *Paragraph:
		p, ok := args.content.(*Paragraph)
		if !ok {
			return "", errors.New("can't convert to Paragraph")
		}

		if p.Style.Color == "" {
			p.Style.Color = args.color
		}

		if p.Style.FontSize == 0 {
			p.Style.FontSize = args.fontSize
		}

		pString, err := p.string(args.document)
		if err != nil {
			return "", errors.Wrap(err, "p.string")
		}

		return pString, nil

	case *List:
		list, ok := args.content.(*List)
		if !ok {
			return "", errors.New("can't convert to List")
		}

		if list.Style.Color == "" {
			list.Style.Color = args.color
		}

		if list.Style.FontSize == 0 {
			list.Style.FontSize = args.fontSize
		}

		var recursionDepth int
		listString, err := list.string(listStringArgs{
			level:          0,
			documnet:       args.document,
			recursionDepth: &recursionDepth,
		})
		if err != nil {
			return "", errors.Wrap(err, "list.string")
		}

		return listString, nil

	case *Table:
		table, ok := args.content.(*Table)
		if !ok {
			return "", errors.New("can't convert to table")
		}

		if table == nil {
			return "", nil
		}

		if table.Style.Color == "" {
			table.Style.Color = args.color
		}

		if table.Style.FontSize == 0 {
			table.Style.FontSize = args.fontSize
		}

		tableString, err := table.string(args.document)
		if err != nil {
			return "", errors.Wrap(err, "table.string")
		}

		return tableString, nil

	default:
		println(fmt.Sprintf("%T", args.content))
		return "", errors.New("undefined item type")
	}
}

type trStringArgs struct {
	table    *Table
	index    int
	document *Document
}

func (args *trStringArgs) error() error {
	if args.document == nil {
		return errors.New("no args.document")
	}

	return nil
}

func (tr *TR) string(args trStringArgs) (string, error) {
	if err := args.error(); err != nil {
		return "", err
	}

	if tr.TD == nil {
		return "", nil
	}

	var buf bytes.Buffer

	buf.WriteString("<w:tr>")
	buf.WriteString(tr.properties())

	for index, td := range tr.TD {
		// td.prepareBorders(args.table.Style.Borders)
		td.setBorderMaybe(setBorderMaybeArgs{
			table:      args.table,
			trIndex:    args.index,
			tdIndex:    index,
			tdTotalCnt: len(tr.TD),
		})

		td, err := td.string(tdBytesArgs{
			document: args.document,
		})
		if err != nil {
			return "", errors.Wrap(err, "td.string")
		}

		buf.WriteString(td)
	}

	buf.WriteString("</w:tr>")

	return buf.String(), nil
}

func (tr *TR) properties() string {
	var buf bytes.Buffer

	buf.WriteString(`<w:trPr>`)

	if tr.CantSplit {
		buf.WriteString(`<w:cantSplit/>`)
	}

	if tr.IsHeader {
		buf.WriteString(`<w:tblHeader/>`)
	}

	if tr.Height > 0 {
		buf.WriteString(`<w:trHeight w:hRule="exact" w:val="` + strconv.Itoa(tr.Height) + `" />`)
	}

	buf.WriteString(`</w:trPr>`)

	return buf.String()
}

type setBorderMaybeArgs struct {
	table      *Table
	trIndex    int
	tdTotalCnt int
	tdIndex    int
}

func (td *TD) setBorderMaybe(args setBorderMaybeArgs) {
	if args.trIndex == 0 && !args.table.Style.Borders.Top.isEmpty() {
		td.Style.Borders.Top = args.table.Style.Borders.Top
	}

	if args.tdIndex == 0 && !args.table.Style.Borders.Left.isEmpty() {
		td.Style.Borders.Left = args.table.Style.Borders.Left
	}

	if args.tdIndex == args.tdTotalCnt-1 && !args.table.Style.Borders.Right.isEmpty() {
		td.Style.Borders.Right = args.table.Style.Borders.Right
	}

	if args.trIndex == len(args.table.TR)-1 && !args.table.Style.Borders.Bottom.isEmpty() {
		td.Style.Borders.Bottom = args.table.Style.Borders.Bottom
	}
}

func (t *Table) string(d *Document) (string, error) {
	var buf bytes.Buffer

	if t.TR == nil {
		return "", nil
	}

	buf.WriteString("<w:tbl>")
	buf.WriteString(t.properties())
	buf.WriteString(t.GetGrid())

	rows, err := t.rowsString(d)
	if err != nil {
		return "", errors.Wrap(err, "t.rowsBytes")
	}

	buf.WriteString(rows)
	buf.WriteString("</w:tbl>")
	buf.WriteString(contextualSpacing(t.NoMarginBottom))

	return buf.String(), nil
}

func (t *Table) rowsString(d *Document) (string, error) {
	if t.TR == nil {
		return "", nil
	}

	var buf bytes.Buffer

	for index, tr := range t.TR {
		trString, err := tr.string(trStringArgs{
			index:    index,
			table:    t,
			document: d,
		})
		if err != nil {
			return "", errors.Wrap(err, "tr.string")
		}

		buf.WriteString(trString)
	}

	return buf.String(), nil
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

func (t *Table) properties() string {
	var buf bytes.Buffer

	buf.WriteString("<w:tblPr>")
	buf.WriteString(t.getStyleClass())
	buf.WriteString(t.getWidth())

	horisontalAlign := "center"
	if t.Style.HorisontalAlign != "" {
		horisontalAlign = t.Style.HorisontalAlign
	}
	buf.WriteString(`<w:jc w:val="` + horisontalAlign + `" />`)
	buf.WriteString(`<w:tblInd w:type="dxa" w:w="0" />`)
	buf.WriteString(`<w:tblLayout w:type="` + t.getType() + `" />`)

	if t.Style.Background != "" {
		buf.WriteString(`<w:shd w:val="clear" w:color="` + t.Style.Background + `" w:fill="` + t.Style.Background + `"/>`)
	}

	if t.CellMargin != nil {
		buf.WriteString(`<w:tblCellMar>`)
		buf.WriteString(`<w:top w:w="` + strconv.Itoa(t.CellMargin.Top.Int()) + `" w:type="dxa" />`)
		buf.WriteString(`<w:left w:w="` + strconv.Itoa(t.CellMargin.Left.Int()) + `" w:type="dxa" />`)
		buf.WriteString(`<w:bottom w:w="` + strconv.Itoa(t.CellMargin.Bottom.Int()) + `" w:type="dxa" />`)
		buf.WriteString(`<w:right w:w="` + strconv.Itoa(t.CellMargin.Right.Int()) + `" w:type="dxa" />`)
		buf.WriteString(`</w:tblCellMar>`)
	}
	buf.WriteString("</w:tblPr>")

	return buf.String()
}

func (t *Table) getWidth() string {
	var width int

	for _, i := range t.Grid {
		width += i
	}

	return `<w:tblW w:type="dxa" w:w="` + strconv.Itoa(width) + `"/>`
}

func (t *Table) getStyleClass() string {
	if t.StyleClass == "" {
		return `<w:tblStyle w:val="normalTable"/>`
	}

	return `<w:tblStyle w:val="` + t.StyleClass + `"/>`
}

func (t *Table) getType() string {
	if len(t.Grid) > 0 {
		return "fixed"
	}

	switch t.Type {
	case "fixed":
		return "fixed"
	case "autofit":
		return "autofit"
	default:
		return "autofit"
	}
}

func (b *Borders) isEmpty() bool {
	if !b.Top.isEmpty() {
		return false
	}

	if !b.Left.isEmpty() {
		return false
	}

	if !b.Right.isEmpty() {
		return false
	}

	if !b.Bottom.isEmpty() {
		return false
	}

	return true
}

func (b *Border) isEmpty() bool {
	if b.Width != 0 {
		return false
	}

	if b.Color != "" {
		return false
	}

	return true
}

func (d *Document) SetTable(table *Table) error {
	if table.TR == nil {
		return nil
	}

	tableString, err := table.string(d)
	if err != nil {
		return errors.Wrap(err, "table.string")
	}

	d.Buf.WriteString(tableString)

	return nil
}

func (img *Image) Error() error {
	if img.Width == 0 {
		return errors.New("no img.Width")
	}

	return nil
}

func (img *Image) string(d *Document) (string, error) {
	if err := img.Error(); err != nil {
		return "", err
	}

	if err := img.populateSizes(); err != nil {
		return "", errors.Wrap(err, "Image.populateSizes")
	}

	id := len(d.images.content) + 1
	relsIdPrefix := imagesID

	if img.isMainPageHeader {
		id = len(d.images.mainPageHeader) + 1
		relsIdPrefix = "mainPageHeaderImageID"
	} else if img.isMainPageFooter {
		id = len(d.images.mainPageFooter) + 1
		relsIdPrefix = "mainPageFooterImageID"
	} else if img.isHeader {
		id = len(d.images.header) + 1
		relsIdPrefix = "fileHeaderImageID"
	} else if img.isFooter {
		id = len(d.images.footer) + 1
		relsIdPrefix = "fileFooterImageID"
	}

	img.ID = id
	img.RelsID = relsIdPrefix + strconv.Itoa(img.ID)
	img.ContentType = http.DetectContentType(img.Bytes)
	img.FileName = filepath.Base(img.FileName)
	img.Extension = filepath.Ext(img.FileName)

	nameWithoutExt := img.FileName[0 : len(img.FileName)-len(img.Extension)]

	widthInEMU := dxaToEMU(int64(img.Width))
	heightInEMU := dxaToEMU(int64(img.Height))

	var buf bytes.Buffer
	buf.WriteString("<w:r>")
	buf.WriteString(`<w:drawing>`)
	buf.WriteString(img.getDisplayTag())
	buf.WriteString(img.getAlign())
	buf.WriteString(img.getWrap())
	buf.WriteString(`<wp:extent cx="` + strconv.FormatInt(widthInEMU, 10) + `" cy="` + strconv.FormatInt(heightInEMU, 10) + `"/>`)
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
	buf.WriteString(`<a:ext cx="` + strconv.FormatInt(widthInEMU, 10) + `" cy="` + strconv.FormatInt(heightInEMU, 10) + `" />`)
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

	if img.isHeader {
		d.images.header = append(d.images.header, img)
	} else if img.isFooter {
		d.images.footer = append(d.images.footer, img)
	} else {
		d.images.content = append(d.images.content, img)
	}

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
		img.Width = int64(float64(img.Height) * float64(config.Width) / float64(config.Height))
	}

	if img.Height == 0 {
		img.Height = int64(float64(img.Width) * float64(config.Height) / float64(config.Width))
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

		return `<wp:anchor behindDoc="` + isRelative + `" distT="` + img.MarginTop.emuString() + `" distB="` + img.MarginBottom.emuString() + `" distL="` + img.MarginLeft.emuString() + `" distR="` + img.MarginRight.emuString() + `" simplePos="0" locked="1" layoutInCell="0" allowOverlap="1" relativeHeight="` + strconv.Itoa(img.ZIndex) + `">`
	}

	return `<wp:inline distT="` + img.MarginTop.emuString() + `" distB="` + img.MarginBottom.emuString() + `" distL="` + img.MarginLeft.emuString() + `" distR="` + img.MarginRight.emuString() + `">`
}

func (img *Image) getWrap() string {
	if img.IsBackground {
		return "<wp:wrapNone/>"
	}

	if img.Display == ImageDisplayFloat {
		return `<wp:wrapSquare wrapText="largest" distT="0" distB="0" distL="0" distR="0" />`
	}

	return ""
}

func dxaToEMU(value int64) int64 {
	// 1440 DXA per inch. 1 inch
	// 914400 EMUs is 1 inch
	return value * 635
}

func mmToEMU(mm int) int {
	return mm * 36000
}

func (d *Document) SetPageBreak() {
	d.Buf.WriteString(`<w:p>`)
	d.Buf.WriteString(`<w:pPr>`)
	d.Buf.WriteString(`<w:pStyle w:val="Normal"/>`)
	d.Buf.WriteString(`<w:rPr></w:rPr>`)
	d.Buf.WriteString(`</w:pPr>`)
	d.Buf.WriteString(`<w:r>`)
	d.Buf.WriteString(`<w:rPr></w:rPr>`)
	d.Buf.WriteString(`</w:r>`)
	d.Buf.WriteString(`<w:r>`)
	d.Buf.WriteString(`<w:br w:type="page"/>`)
	d.Buf.WriteString(`</w:r>`)
	d.Buf.WriteString(`</w:p>`)
}

type Section struct {
	Type            string
	PageOrientation string
	Margins         *Margins
}

func (d *Document) SetSection(section *Section) error {
	d.Buf.WriteString(section.string(d))

	return nil
}

func (section *Section) string(d *Document) string {
	if section.Type == "" {
		section.Type = SectionTypeContinious
	}

	if section.PageOrientation == "" {
		section.PageOrientation = d.PageOrientation
	}

	if section.Margins == nil {
		section.Margins = &d.Margins
	}

	var buf bytes.Buffer
	buf.WriteString(`<w:p>`)
	buf.WriteString(`<w:pPr>`)
	buf.WriteString(`<w:sectPr>`)
	buf.WriteString(`<w:type w:val="` + section.Type + `"/>`)
	buf.WriteString(sectionSizes(section.PageOrientation))
	buf.WriteString(sectionMargins(*section.Margins))
	buf.WriteString(`</w:sectPr>`)
	buf.WriteString(`</w:pPr>`)
	buf.WriteString(`</w:p>`)

	return buf.String()
}
