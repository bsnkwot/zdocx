package zdocx

import (
	"github.com/pkg/errors"
)

const (
	ErrorColor   = "DB5200"
	AccentColor  = "7B9CE6"
	WarningColor = "E8990C"
	SuccessColor = "7CAB2B"
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

	d.setMarginMaybe()

	alert.Width = pageWidth - int(d.MarginLeft.Value) - int(d.MarginRight.Value)
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
		Style: Style{
			Borders: Borders{
				Top: Border{
					Color: "DB4912",
					Width: 4,
				},
				Left: Border{
					Color: "DB4912",
					Width: 4,
				},
				Right: Border{
					Color: "DB4912",
					Width: 4,
				},
				Bottom: Border{
					Color: "DB4912",
					Width: 4,
				},
			},
		},
		Type: "fixed",
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
