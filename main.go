package main

import (
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

	// if err := doc.SetP(&zdocx.Paragraph{
	// 	Texts: []*zdocx.Text{
	// 		{
	// 			Text: "But I must explain to you how all this mistaken idea of denouncing pleasure and praising pain was born and I will give you a complete account of the system, and expound the actual teachings of the great explorer of the truth, the master-builder of human happiness. No one rejects, dislikes, or avoids pleasure itself, because it is pleasure, but because those who do not know how to pursue pleasure rationally encounter consequences that are extremely painful. Nor again is there anyone who loves or pursues or desires to obtain pain of itself, because it is pain, but because occasionally circumstances occur in which toil and pain can procure him some great pleasure. To take a trivial example, which of us ever undertakes laborious physical exercise, except to obtain some advantage from it? But who has any right to find fault with a man who chooses to enjoy a pleasure that has no annoying consequences, or one who avoids a pain that produces no resultant pleasure?",
	// 		},
	// 	},
	// }); err != nil {
	// 	panic(err)
	// }

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

	if err := doc.Save(zdocx.SaveArgs{
		FileName: "document",
	}); err != nil {
		panic(err)
	}

	println("done!")
}
