package zdocx

import (

	// "io/ioutil"
	"bytes"
	"math/rand"
	"strconv"
	"time"

	"github.com/pkg/errors"
)

const (
	ListDecimalID          = 1
	ListBulletID           = 2
	ListNoneID             = 3
	ListDecimalType        = "decimal"
	ListBulletType         = "bullet"
	ListNoneType           = "none"
	TableCellDefaultMargin = 55
	DocumentDefaultMargin  = 1440
	PageOrientationAlbum   = "album"
	PageOrientationBook    = "book"
	StylesID               = 1
	ImagesID               = 2
	NumberingID            = 3
	FontTableID            = 4
	SettingsID             = 5
	ThemeID                = 6
	HeaderID               = 7
	FooterID               = 8
	LinkIDPrefix           = "linkId"
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
}

type Image struct {
	FileName    string
	ContentType string
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

func NewDocument() *Document {
	doc := Document{}
	doc.setStartTags()
	return &doc
}

type SaveArgs struct {
	FileName  string
	Directory string
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
		FileName:       "test.docx",
		TemplatesFiles: templatesFilesList(),
		Document:       d,
	}); err != nil {
		return errors.Wrap(err, "ZipFiles")
	}

	return nil
}

type templateFile struct {
	Name     string
	SavePath string
	Bytes    []byte
}

func (i *templateFile) FullName() string {
	if i.SavePath == "" {
		return i.Name
	}

	return i.SavePath + "/" + i.Name
}

func templatesFilesList() []*templateFile {
	return []*templateFile{
		{
			Name:     ".rels",
			SavePath: "_rels",
			Bytes:    []byte(templateRelsRels),
		},
		{
			Name:     "app.xml",
			SavePath: "docProps",
			Bytes:    []byte(templateDocPropsApp),
		},
		{
			Name:     "styles.xml",
			SavePath: "word",
			Bytes:    []byte(templateWordStyles),
		},
		{
			Name:     "numbering.xml",
			SavePath: "word",
			Bytes:    []byte(templateWordNumbering),
		},
		{
			Name:     "fontTable.xml",
			SavePath: "word",
			Bytes:    []byte(templateWordFontTable),
		},
		{
			Name:     "theme1.xml",
			SavePath: "word/theme",
			Bytes:    []byte(templateWordTheme),
		},
	}
}

func (d *Document) setStartTags() {
	d.Buf.WriteString(getDocumentStartTags("document"))
}

func getDocumentStartTags(tag string) string {
	return `<?xml version="1.0" encoding="utf-8" standalone="yes"?><w:` + tag + ` xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">`
}

func (d *Document) SetBody() {
	d.Buf.WriteString("<w:body>")
}

func (d *Document) SetBodyClose() {
	d.SetSectionProperties()
	d.Buf.WriteString("</w:body>")
	d.Buf.WriteString("</w:document>")
}

func (d *Document) String() string {
	return d.Buf.String()
}

func (d *Document) SetSpace() string {
	return setSpace()
}

func setSpace() string {
	return `<w:r><w:t xml:space="preserve"> </w:t></w:r>`
}

func (d *Document) SetP(p *Paragraph) {
	d.Buf.WriteString(p.String())
}

func (p *Paragraph) String() string {
	var buf bytes.Buffer

	buf.WriteString("<w:p>")
	buf.WriteString(p.GetProperties())

	for index, t := range p.Texts {
		if index != 0 {
			buf.WriteString(setSpace())
		}

		buf.WriteString(t.String())
	}

	buf.WriteString("</w:p>")

	return buf.String()
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

func uniqueID(prefix string) string {
	rand.Seed(time.Now().UnixNano())
	return prefix + strconv.Itoa(rand.Intn(100000))
}

func (t *Text) String() string {
	if t.Text == "" {
		return ""
	}

	var buf bytes.Buffer

	if t.Link != nil {
		id := uniqueID(LinkIDPrefix)
		buf.WriteString(`<w:hyperlink r:id="` + id + `">`)
	}

	buf.WriteString("<w:r>")
	buf.WriteString(t.GetProperties())
	buf.WriteString("<w:t>" + t.Text + "</w:t>")
	buf.WriteString("</w:r>")

	if t.Link != nil {
		buf.WriteString("</w:hyperlink>")
	}

	return buf.String()
}

func (t *Text) GetProperties() string {
	if t.Style.IsEmpty() && t.StyleClass == "" {
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

func (d *Document) SetSectionProperties() {
	d.Buf.WriteString("<w:sectPr>")

	if len(d.Header) != 0 {
		d.Buf.WriteString(`<w:headerReference w:type="default" r:id="rId` + strconv.Itoa(HeaderID) + `"/>`)
	}

	if len(d.Footer) != 0 {
		d.Buf.WriteString(`<w:footerReference w:type="default" r:id="rId` + strconv.Itoa(FooterID) + `"/>`)
	}

	d.Buf.WriteString(`<w:type w:val="nextPage"/>`)
	d.SetPageSizes()
	d.SetMargins()
	d.Buf.WriteString(`<w:pgNumType w:fmt="decimal"/>`)
	d.Buf.WriteString(`<w:formProt w:val="false"/>`)
	d.Buf.WriteString(`<w:textDirection w:val="lrTb"/>`)
	d.Buf.WriteString(`<w:docGrid w:type="default" w:linePitch="100" w:charSpace="0"/>`)
	d.Buf.WriteString("</w:sectPr>")
}

func (d *Document) SetPageSizes() {
	width := 12240
	height := 15840

	if d.PageOrientation == PageOrientationAlbum {
		height, width = width, height
	}

	d.Buf.WriteString(`<w:pgSz w:w="` + strconv.Itoa(width) + `" w:h="` + strconv.Itoa(height) + `"/>`)
}

func (d *Document) SetMargins() {
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
	var infinityLoopCnt int

	listString, err := getList(getListArgs{
		List:            list,
		Level:           0,
		InfinityLoopCnt: &infinityLoopCnt,
	})
	if err != nil {
		return errors.Wrap(err, "getList")
	}

	d.Buf.WriteString(listString)

	return nil
}

type getListArgs struct {
	List            *List
	Level           int
	InfinityLoopCnt *int
}

func getList(args getListArgs) (string, error) {
	if *args.InfinityLoopCnt >= 1000 {
		return "", errors.New("infinity loop")
	}

	*args.InfinityLoopCnt++

	if args.List.LI == nil {
		return "", nil
	}

	var buf bytes.Buffer

	for _, li := range args.List.LI {
		if li.Items == nil {
			continue
		}

		for index, i := range li.Items {
			switch i.(type) {
			case *Paragraph:
				pString, err := getListP(getListPArgs{
					Index:    index,
					ListType: ListBulletType,
					Level:    args.Level,
					Item:     i,
				})
				if err != nil {
					return "", errors.Wrap(err, "getListP")
				}

				buf.WriteString(pString)
			case *List:
				listInListString, err := getListInList(getListInListArgs{
					Item:            i,
					Level:           args.Level + 1,
					InfinityLoopCnt: args.InfinityLoopCnt,
					Type:            ListBulletType,
				})
				if err != nil {
					return "", errors.Wrap(err, "setListInList")
				}

				buf.WriteString(listInListString)
			default:
				return "", errors.New("undefined item type")
			}
		}
	}

	return buf.String(), nil
}

type getListInListArgs struct {
	Item            interface{}
	Level           int
	InfinityLoopCnt *int
	Type            string
}

func getListInList(args getListInListArgs) (string, error) {
	list, ok := args.Item.(*List)
	if !ok {
		return "", errors.New("can't convert to List")
	}

	list.Type = args.Type

	listString, err := getList(getListArgs{
		List:            list,
		Level:           args.Level,
		InfinityLoopCnt: args.InfinityLoopCnt,
	})
	if err != nil {
		return "", errors.Wrap(err, "getList")
	}

	return listString, nil
}

type getListPArgs struct {
	Item     interface{}
	Index    int
	ListType string
	Level    int
}

func getListP(args getListPArgs) (string, error) {
	item, ok := args.Item.(*Paragraph)
	if !ok {
		return "", errors.New("can't convert to Paragraph")
	}

	if args.Index == 0 {
		if args.ListType == "" {
			args.ListType = ListBulletType
		}

		item.ListParams = &ListParams{
			Level: args.Level,
			Type:  args.ListType,
		}
	} else {
		item.Style.MarginLeft = 720 * (args.Level + 1)
	}

	return item.String(), nil
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
	CellMargin Margin
}

func (td *TD) GetProperties(width int) string {
	var buf bytes.Buffer

	buf.WriteString("<w:tcPr>")
	buf.WriteString(`<w:tcW w:w="` + strconv.Itoa(width) + `" w:type="dxa" />`)
	buf.WriteString("<w:tcBorders>")
	buf.WriteString(`<w:top w:val="single" w:sz="2" w:space="0" w:color="000000"/>`)
	buf.WriteString(`<w:left w:val="single" w:sz="2" w:space="0" w:color="000000"/>`)
	buf.WriteString(`<w:right w:val="single" w:sz="2" w:space="0" w:color="000000"/>`)
	buf.WriteString(`<w:bottom w:val="single" w:sz="2" w:space="0" w:color="000000"/>`)
	buf.WriteString("</w:tcBorders>")
	buf.WriteString("</w:tcPr>")

	return buf.String()
}

func (td *TD) String(width int) (string, error) {
	content, err := td.GetContent()
	if err != nil {
		return "", errors.Wrap(err, "td.GetContent")
	}

	var buf bytes.Buffer

	buf.WriteString("<w:tc>")
	buf.WriteString(td.GetProperties(width))
	buf.WriteString(content)
	buf.WriteString("</w:tc>")

	return buf.String(), nil
}

func getContentFromInterface(content interface{}) (string, error) {
	switch content.(type) {
	case *Paragraph:
		p, ok := content.(*Paragraph)
		if !ok {
			return "", errors.New("can't convert to Paragraph")
		}

		return p.String(), nil
	case *List:
		list, ok := content.(*List)
		if !ok {
			return "", errors.New("can't convert to List")
		}

		var infinityLoopCnt int
		listString, err := getList(getListArgs{
			List:            list,
			Level:           0,
			InfinityLoopCnt: &infinityLoopCnt,
		})
		if err != nil {
			return "", errors.Wrap(err, "setList")
		}

		return listString, nil
	default:
		return "", errors.New("undefined item type")
	}
}

func (td *TD) GetContent() (string, error) {
	for _, i := range td.Content {
		contentString, err := getContentFromInterface(i)
		if err != nil {
			return "", errors.Wrap(err, "getContentFromInterface")
		}

		return contentString, nil
	}

	return "", nil
}

func (tr *TR) String(table *Table) (string, error) {
	if len(table.Grid) < len(tr.TD) {
		return "", errors.New("len of Grim less then len of TD")
	}

	if tr.TD == nil {
		return "", nil
	}

	var buf bytes.Buffer
	buf.WriteString("<w:tr>")

	for index, td := range tr.TD {
		tdString, err := td.String(table.Grid[index])
		if err != nil {
			return "", nil
		}

		buf.WriteString(tdString)
	}

	buf.WriteString("</w:tr>")

	return buf.String(), nil
}

func (t *Table) String() (string, error) {
	if t.TR == nil {
		return "", nil
	}

	rowsString, err := t.GetRowsString()
	if err != nil {
		return "", errors.Wrap(err, "t.GetRowsString")
	}

	var buf bytes.Buffer

	buf.WriteString("<w:tbl>")
	buf.WriteString(getCommonStyleClass(t.StyleClass))
	buf.WriteString(t.GetPropperties())
	buf.WriteString(t.GetGrid())
	buf.WriteString(rowsString)
	buf.WriteString("</w:tbl>")

	return buf.String(), nil
}

func (t *Table) GetRowsString() (string, error) {
	if t.TR == nil {
		return "", nil
	}

	var buf bytes.Buffer

	for _, tr := range t.TR {
		trString, err := tr.String(t)
		if err != nil {
			return "", errors.Wrap(err, "tr.String")
		}

		buf.WriteString(trString)
	}

	return buf.String(), nil
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
	margin := Margin{
		Top:    t.CellMargin.Top,
		Bottom: t.CellMargin.Bottom,
		Left:   t.CellMargin.Left,
		Right:  t.CellMargin.Right,
	}

	if margin.Top == 0 {
		margin.Top = TableCellDefaultMargin
	}

	if margin.Bottom == 0 {
		margin.Bottom = TableCellDefaultMargin
	}

	if margin.Left == 0 {
		margin.Left = TableCellDefaultMargin
	}

	if margin.Right == 0 {
		margin.Right = TableCellDefaultMargin
	}

	t.CellMargin = margin
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

	tableString, err := table.String()
	if err != nil {
		return errors.Wrap(err, "table.String")
	}

	d.Buf.WriteString(tableString)

	return nil
}
