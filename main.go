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

	buf, err := doc.WriteToBuffer()
	if err != nil {
		panic(err)
	}

	f, err := os.Create("temp_file.docx")
	if err != nil {
		panic(err)
	}
	defer f.Close()

	if _, err := f.Write(buf.Bytes()); err != nil {
		panic(err)
	}

	// if err := doc.Save(zdocx.SaveArgs{
	// 	FileName: "document",
	// }); err != nil {
	// 	panic(err)
	// }

	println("done!")
}
