package main

import (
	"bufio"
	"bytes"
	"fmt"
	"io/ioutil"
	"os"
	"time"
	"zdocx/zdocx"

	"github.com/wcharczuk/go-chart"
	"github.com/wcharczuk/go-chart/drawing"
)

func main() {
	doc := zdocx.NewDocument(zdocx.NewDocumentArgs{})
	doc.PageOrientation = zdocx.PageOrientationAlbum

	doc.Header = []*zdocx.Paragraph{
		{
			Texts: []*zdocx.Text{
				{
					Text: "header text",
				},
			},
		},
	}

	doc.Footer = []*zdocx.Paragraph{
		{
			Texts: []*zdocx.Text{
				{
					Text: "footer text",
				},
			},
		},
	}

	if err := doc.SetP(&zdocx.Paragraph{
		StyleClass: "h1",
		Texts: []*zdocx.Text{
			{
				Text: "Title",
			},
		},
	}); err != nil {
		panic(err)
	}

	if err := doc.SetP(&zdocx.Paragraph{
		StyleClass: "h2",
		Texts: []*zdocx.Text{
			{
				Text: "subtitle",
			},
		},
	}); err != nil {
		panic(err)
	}

	if err := doc.SetP(&zdocx.Paragraph{
		StyleClass: "h3",
		Texts: []*zdocx.Text{
			{
				Text: "medium title",
			},
		},
	}); err != nil {
		panic(err)
	}

	if err := doc.SetP(&zdocx.Paragraph{
		Texts: []*zdocx.Text{
			{
				Text: "medium title",
				Style: zdocx.TextStyle{
					IsBold:   true,
					IsItalic: true,
				},
			},
		},
	}); err != nil {
		panic(err)
	}

	if err := doc.SetP(&zdocx.Paragraph{
		Texts: []*zdocx.Text{
			{
				Text: "lorem ipsum?",
				Style: zdocx.TextStyle{
					FontSize: 40,
				},
			},
		},
	}); err != nil {
		panic(err)
	}

	if err := doc.SetList(&zdocx.List{
		Type: zdocx.ListDecimalType,
		LI: []*zdocx.LI{
			{
				Items: []interface{}{
					&zdocx.Paragraph{
						Texts: []*zdocx.Text{
							{
								Text: "Lorem ipsum dolor sit amet",
							},
						},
					},
				},
			},
		},
	}); err != nil {
		panic(err)
	}

	if err := doc.SetList(&zdocx.List{
		LI: []*zdocx.LI{
			{
				Items: []interface{}{
					&zdocx.Paragraph{
						Texts: []*zdocx.Text{
							{
								Text: "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum",
							},
						},
					},
					&zdocx.Paragraph{
						Texts: []*zdocx.Text{
							{
								Text: "second simple text",
							},
						},
					},
				},
			},
			{
				Items: []interface{}{
					&zdocx.Paragraph{
						Texts: []*zdocx.Text{
							{
								Text: "fird text",
							},
							{
								Text: "Bold text",
								Style: zdocx.TextStyle{
									IsBold: true,
								},
							},
						},
					},
				},
			},
			{
				Items: []interface{}{
					&zdocx.List{
						Type: zdocx.ListDecimalType,
						LI: []*zdocx.LI{
							{
								Items: []interface{}{
									&zdocx.Paragraph{
										Texts: []*zdocx.Text{
											{
												Text: "Lorem ipsum dolor sit amet",
											},
										},
									},
								},
							},
							{
								Items: []interface{}{
									&zdocx.Paragraph{
										Texts: []*zdocx.Text{
											{
												Text: "test test test",
											},
										},
									},
								},
							},
						},
					},
				},
			},
		},
	}); err != nil {
		panic(err)
	}

	doc.SetPageBreak()

	if err := doc.SetTable(&zdocx.Table{
		Type:  "fixed",
		Width: doc.GetInnerWidth(),
		Grid: []int{
			6480,
			6480,
		},
		TR: []*zdocx.TR{
			{
				TD: []*zdocx.TD{
					{
						Content: []interface{}{
							&zdocx.Paragraph{
								Texts: []*zdocx.Text{
									{
										Text: "cell 1",
									},
								},
							},
						},
					},
					{
						Content: []interface{}{
							&zdocx.Paragraph{
								Texts: []*zdocx.Text{
									{
										Text: "cell 2",
									},
								},
							},
						},
					},
				},
			},
			{
				TD: []*zdocx.TD{
					{
						Content: []interface{}{
							&zdocx.Paragraph{
								Texts: []*zdocx.Text{
									{
										Text: "cell 3",
									},
								},
							},
						},
					},
					{
						Content: []interface{}{
							&zdocx.List{
								LI: []*zdocx.LI{
									{
										Items: []interface{}{
											&zdocx.Paragraph{
												Texts: []*zdocx.Text{
													{
														Text: "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum",
													},
												},
											},
											&zdocx.Paragraph{
												Texts: []*zdocx.Text{
													{
														Text: "second simple text",
													},
												},
											},
										},
									},
									{
										Items: []interface{}{
											&zdocx.Paragraph{
												Texts: []*zdocx.Text{
													{
														Text: "fird text",
													},
													{
														Text: "Bold text",
														Style: zdocx.TextStyle{
															IsBold: true,
														},
													},
												},
											},
										},
									},
									{
										Items: []interface{}{
											&zdocx.List{
												Type: zdocx.ListDecimalType,
												LI: []*zdocx.LI{
													{
														Items: []interface{}{
															&zdocx.Paragraph{
																Texts: []*zdocx.Text{
																	{
																		Text: "Lorem ipsum dolor sit amet",
																	},
																},
															},
														},
													},
													{
														Items: []interface{}{
															&zdocx.Paragraph{
																Texts: []*zdocx.Text{
																	{
																		Text: "test test test",
																	},
																},
															},
														},
													},
												},
											},
										},
									},
								},
							},
						},
					},
				},
			},
		},
	}); err != nil {
		panic(err)
	}

	file, err := os.Open("temp/img.jpg")
	if err != nil {
		panic(err)
	}
	defer file.Close()

	imageBytes, err := ioutil.ReadAll(file)

	if err := doc.SetP(&zdocx.Paragraph{
		Style: zdocx.PStyle{
			HorisontalAlign: "center",
		},
		Texts: []*zdocx.Text{
			{
				Image: &zdocx.Image{
					FileName: file.Name(),
					Bytes:    imageBytes,
					Width:    50,
				},
			},
		},
	}); err != nil {
		panic(err)
	}

	gridColor := drawing.Color{R: 123, G: 156, B: 230, A: 1}.WithAlpha(80)
	gridWidth := 0.5
	graph := chart.Chart{
		Background: chart.Style{
			Padding: chart.Box{
				Top:    10,
				Left:   20,
				Right:  20,
				Bottom: 10,
			},
		},
		Height: 300,
		Width:  1200,
		YAxisSecondary: chart.YAxis{
			Style: chart.Style{
				Hidden: true,
			},
		},
		XAxis: chart.XAxis{
			Style: chart.Style{
				StrokeWidth: 0,
				StrokeColor: gridColor,
				FontColor:   drawing.ColorFromHex("9e9e9e"),
				DotWidth:    0,
			},
		},
		YAxis: chart.YAxis{
			AxisType: chart.YAxisSecondary,
			Style: chart.Style{
				StrokeColor: drawing.ColorFromHex("ffffff"),
				FontColor:   drawing.ColorFromHex("9e9e9e"),
			},
			GridMinorStyle: chart.Style{
				StrokeColor: gridColor,
				StrokeWidth: gridWidth,
			},
			Range: &chart.ContinuousRange{
				Min: 0.0,
				Max: 55.0,
			},
			ValueFormatter: func(v interface{}) string {
				if i, ok := v.(float64); ok {
					return fmt.Sprintf("%.0f", i)
				}
				return ""
			},
		},
		Series: []chart.Series{
			chart.TimeSeries{
				Style: chart.Style{
					StrokeColor: drawing.ColorFromHex("4692e8"),
					StrokeWidth: 2,

					DotColor: drawing.ColorFromHex("4692e8"),
					DotWidth: 3,
				},
				XValues: []time.Time{
					time.Now().AddDate(0, 0, -9),
					time.Now().AddDate(0, 0, -8),
					time.Now().AddDate(0, 0, -7),
					time.Now().AddDate(0, 0, -6),
					time.Now().AddDate(0, 0, -5),
					time.Now().AddDate(0, 0, -4),
					time.Now(),
				},
				YValues: []float64{50.0, 40.0, 47.0, 45.0, 22.0, 35.0, 33.0},
			},
			chart.TimeSeries{
				YValues: []float64{0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0},
			},
		},
	}

	var imageBuf bytes.Buffer
	writer := bufio.NewWriter(&imageBuf)

	graph.Render(chart.PNG, writer)

	if err := doc.SetP(&zdocx.Paragraph{
		Style: zdocx.PStyle{
			HorisontalAlign: "center",
		},
		Texts: []*zdocx.Text{
			{
				Image: &zdocx.Image{
					FileName: "temp_image.png",
					Bytes:    imageBuf.Bytes(),
					Width:    165,
				},
			},
		},
	}); err != nil {
		panic(err)
	}

	if err := doc.Save(zdocx.SaveArgs{
		FileName: "document",
	}); err != nil {
		panic(err)
	}

	println("done!")
}
