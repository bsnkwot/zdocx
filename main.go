package main

import (
	"io/ioutil"
	"os"
	"zdocx/model/zdocx"
)

func main() {
	doc := zdocx.NewDocument()
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
				Style: zdocx.Style{
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
				Style: zdocx.Style{
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
								Style: zdocx.Style{
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

	if err := doc.SetTable(&zdocx.Table{
		Type:        "fixed",
		Width:       doc.GetInnerWidth(),
		BorderColor: "000000",
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
														Style: zdocx.Style{
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

	if err := doc.SetAlert(&zdocx.Alert{
		P: []*zdocx.Paragraph{
			{
				Texts: []*zdocx.Text{
					{
						Text: "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.",
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
		Style: zdocx.Style{
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

	docInnerWidth := doc.GetInnerWidth()

	doc.SetTable(&zdocx.Table{
		Type: "fixed",
		Grid: []int{
			int(float32(docInnerWidth) * 0.7),
			int(float32(docInnerWidth) * 0.1),
			int(float32(docInnerWidth) * 0.1),
			int(float32(docInnerWidth) * 0.1),
		},
		TR: []*zdocx.TR{
			{
				TD: []*zdocx.TD{
					{
						Content: []interface{}{
							&zdocx.Paragraph{
								Style: zdocx.Style{
									MarginBottom: &zdocx.Margin{Value: 0},
								},
								Texts: []*zdocx.Text{
									{
										Style: zdocx.Style{
											IsBold: true,
										},
										Text: "Keywords",
									},
								},
							},
						},
					},
					{
						Content: []interface{}{
							&zdocx.Paragraph{
								Style: zdocx.Style{
									HorisontalAlign: "center",
									MarginBottom:    &zdocx.Margin{Value: 0},
								},
								Texts: []*zdocx.Text{
									{
										Style: zdocx.Style{
											IsBold: true,
										},
										Text: "11.01.2020",
									},
								},
							},
						},
					},
					{
						Content: []interface{}{
							&zdocx.Paragraph{
								Style: zdocx.Style{
									HorisontalAlign: "center",
									MarginBottom:    &zdocx.Margin{Value: 0},
								},
								Texts: []*zdocx.Text{
									{
										Style: zdocx.Style{
											IsBold: true,
										},
										Text: "21.01.2020",
									},
								},
							},
						},
					},
					{
						Content: []interface{}{
							&zdocx.Paragraph{
								Style: zdocx.Style{
									HorisontalAlign: "center",
									MarginBottom:    &zdocx.Margin{Value: 0},
								},
								Texts: []*zdocx.Text{
									{
										Style: zdocx.Style{
											IsBold: true,
										},
										Text: "11.02.2020",
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
								Style: zdocx.Style{
									MarginBottom: &zdocx.Margin{Value: 0},
								},
								Texts: []*zdocx.Text{
									{
										Text: "почта россии карточка сведений о контрагенте",
									},
								},
							},
						},
					},
					{
						Content: []interface{}{
							&zdocx.Paragraph{
								Style: zdocx.Style{
									HorisontalAlign: "center",
									MarginBottom:    &zdocx.Margin{Value: 0},
								},
								Texts: []*zdocx.Text{
									{
										Text: "40",
									},
								},
							},
						},
					},
					{
						Content: []interface{}{
							&zdocx.Paragraph{
								Style: zdocx.Style{
									HorisontalAlign: "center",
									MarginBottom:    &zdocx.Margin{Value: 0},
								},
								Texts: []*zdocx.Text{
									{
										Text: "18",
									},
								},
							},
						},
					},
					{
						Content: []interface{}{
							&zdocx.Paragraph{
								Style: zdocx.Style{
									HorisontalAlign: "center",
									MarginBottom:    &zdocx.Margin{Value: 0},
								},
								Texts: []*zdocx.Text{
									{
										Text: "2",
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
								Style: zdocx.Style{
									MarginBottom: &zdocx.Margin{Value: 0},
								},
								Texts: []*zdocx.Text{
									{
										Text: "бгт тюмень",
									},
								},
							},
						},
					},
					{
						Content: []interface{}{
							&zdocx.Paragraph{
								Style: zdocx.Style{
									HorisontalAlign: "center",
									MarginBottom:    &zdocx.Margin{Value: 0},
								},
								Texts: []*zdocx.Text{
									{
										Style: zdocx.Style{
											Color: zdocx.ErrorColor,
										},
										Text: ">50",
									},
								},
							},
						},
					},
					{
						Content: []interface{}{
							&zdocx.Paragraph{
								Style: zdocx.Style{
									HorisontalAlign: "center",
									MarginBottom:    &zdocx.Margin{Value: 0},
								},
								Texts: []*zdocx.Text{
									{
										Text: "38",
									},
								},
							},
						},
					},
					{
						Content: []interface{}{
							&zdocx.Paragraph{
								Style: zdocx.Style{
									MarginBottom:    &zdocx.Margin{Value: 0},
									HorisontalAlign: "center",
								},
								Texts: []*zdocx.Text{
									{
										Text: "10",
									},
								},
							},
						},
					},
				},
			},
		},
	})

	if err := doc.Save(zdocx.SaveArgs{
		FileName: "document",
	}); err != nil {
		panic(err)
	}

	println("done!")
}
