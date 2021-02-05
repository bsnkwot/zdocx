package zdocx

import (
	"github.com/pkg/errors"
)

type Alert struct {
	P     []*Paragraph
	Width int
}

func (alert *Alert) setWidth(d *Document) {
	if alert.Width != 0 {
		return
	}

	pageWidth := PageWidth
	if d.PageOrientation == PageOrientationAlbum {
		pageWidth = PageHeight
	}

	if d.Margin == nil {
		d.Margin = &Margin{
			Top:    DocumentDefaultMargin,
			Left:   DocumentDefaultMargin,
			Right:  DocumentDefaultMargin,
			Bottom: DocumentDefaultMargin,
		}
	}

	alert.Width = pageWidth - d.Margin.Left - d.Margin.Right
}

func (d *Document) SetAlert(alert *Alert) error {
	if len(alert.P) == 0 {
		return nil
	}

	d.writeContextualSpacing()

	alert.setWidth(d)

	firstParagraph := &Paragraph{
		Style: Style{
			PageBreakBefore: true,
		},
		Texts: []*Text{
			{
				Text:       "Внимание!",
				StyleClass: "alertTitle",
			},
		},
	}

	paragraphs := []interface{}{}

	for index, i := range alert.P {
		if index == 0 {
			firstParagraph.Texts = append(firstParagraph.Texts, alert.P[0].Texts[0:]...)
			paragraphs = append(paragraphs, firstParagraph)
		} else {
			i.Style.PageBreakBefore = true
			paragraphs = append(paragraphs, i)
		}
	}

	if err := d.SetTable(&Table{
		CellMargin: &Margin{
			Top:    100,
			Left:   300,
			Right:  200,
			Bottom: 300,
		},
		BorderColor: "DB4912",
		Type:        "fixed",
		Grid: []int{
			alert.Width,
		},
		TR: []*TR{
			{
				TD: []*TD{
					{
						Content: paragraphs,
					},
				},
			},
		},
	}); err != nil {
		return errors.Wrap(err, "Document.SetTable")
	}

	d.writeContextualSpacing()

	return nil
}

// func (d *Document) setAlertImageMaybe() error {
// 	if d.alertImage != nil {
// 		return nil
// 	}

// 	file, err := os.Open("temp/img.jpg")
// 	if err != nil {
// 		panic(err)
// 	}
// 	defer file.Close()

// 	imageBytes, err := ioutil.ReadAll(file)

// 	d.alertImage = &Image{
// 		FileName: file.Name(),
// 		Bytes:    imageBytes,
// 		Width:    50,
// 	}

// 	return nil
// }
