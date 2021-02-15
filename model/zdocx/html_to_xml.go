package zdocx

import (
	"strings"

	"github.com/pkg/errors"
	"golang.org/x/net/html"
)

type Node struct {
	Tag      string
	Text     string
	Children []*Node
}

type ParseHTMLArgs struct {
	Text string
}

func ParseHTML(args ParseHTMLArgs) (*Node, error) {
	r := strings.NewReader(args.Text)

	doc, err := html.Parse(r)
	if err != nil {
		return nil, errors.Wrap(err, "html.Parse")
	}

	root := Node{}

	var f func(*html.Node, *Node)

	f = func(n *html.Node, parent *Node) {
		theParent := parent

		if n.Type == html.ElementNode {
			theParent = &Node{
				Tag: n.Data,
			}

			parent.Children = append(parent.Children, theParent)
		} else if n.Type == html.TextNode {
			trimedText := strings.TrimSpace(n.Data)

			if parent.Tag == "b" {
				parent.Text = n.Data
			} else if parent.Tag == "i" {
				parent.Text = n.Data
			} else if parent.Tag != "p" && trimedText != "" {
				parent.Children = append(parent.Children, &Node{
					Tag: "p",
					Children: []*Node{
						{
							Tag:  "span",
							Text: n.Data,
						},
					},
					Text: n.Data,
				})
			} else if trimedText != "" {
				parent.Children = append(parent.Children, &Node{
					Tag:  "span",
					Text: n.Data,
				})
			}
		}

		for c := n.FirstChild; c != nil; c = c.NextSibling {
			f(c, theParent)
		}
	}

	f(doc, &root)

	return &root, nil
}

func (d *Document) HTMLToXML(node *Node) error {
	if err := d.setTagsFromNode(node); err != nil {
		return errors.Wrap(err, "d.setTagsFromNode")
	}

	return nil
}

type ItemsFromHTMLArgs struct {
	Text string
}

func ItemsFromHTML(args ItemsFromHTMLArgs) ([]interface{}, error) {
	if args.Text == "" {
		return nil, nil
	}

	node, err := ParseHTML(ParseHTMLArgs{
		Text: args.Text,
	})
	if err != nil {
		return nil, errors.Wrap(err, "ParseHTML")
	}

	items, err := HTMLToXMLItems(node, []interface{}{})
	if err != nil {
		return nil, errors.Wrap(err, "HTMLToXMLItems")
	}

	return items, nil
}

func HTMLToXMLItems(node *Node, items []interface{}) ([]interface{}, error) {
	if node.Tag == "p" {
		item, err := node.xmlStruct()
		if err != nil {
			return nil, errors.Wrap(err, "node.xmlStruct")
		}

		p, ok := item.(*Paragraph)
		if !ok {
			return nil, errors.New("can't convert top Paragraph")
		}

		items = append(items, p)

		return items, nil
	}

	if node.Tag == "ul" || node.Tag == "ol" {
		item, err := node.xmlStruct()
		if err != nil {
			return nil, errors.Wrap(err, "node.xmlStruct")
		}

		list, ok := item.(*List)
		if !ok {
			return nil, errors.New("can't convert top List")
		}

		items = append(items, list)

		return items, nil
	}

	for _, n := range node.Children {
		theItems, err := HTMLToXMLItems(n, []interface{}{})
		if err != nil {
			return nil, errors.Wrap(err, "d.setTagsFromNode")
		}

		items = append(items, theItems...)
	}

	return items, nil
}

func (d *Document) setTagsFromNode(node *Node) error {
	if node.Tag == "p" {
		item, err := node.xmlStruct()
		if err != nil {
			return errors.Wrap(err, "node.xmlStruct")
		}

		p, ok := item.(*Paragraph)
		if !ok {
			return errors.New("can't convert top Paragraph")
		}

		if err := d.SetP(p); err != nil {
			return errors.Wrap(err, "d.SetP")
		}

		return nil
	}

	if node.Tag == "ul" || node.Tag == "ol" {
		item, err := node.xmlStruct()
		if err != nil {
			return errors.Wrap(err, "node.xmlStruct")
		}

		list, ok := item.(*List)
		if !ok {
			return errors.New("can't convert top List")
		}

		if err := d.SetList(list); err != nil {
			return errors.Wrap(err, "d.SetList")
		}

		return nil
	}

	for _, n := range node.Children {
		if err := d.setTagsFromNode(n); err != nil {
			return errors.Wrap(err, "d.setTagsFromNode")
		}
	}

	return nil
}

func (n *Node) xmlListStruct() (*List, error) {
	list := List{LI: []*LI{}}

	for _, i := range n.Children {
		item, err := i.xmlStruct()
		if err != nil {
			return nil, errors.Wrap(err, "i.xmlStruct")
		}

		li, ok := item.(*LI)
		if !ok {
			return nil, errors.Wrap(err, "item.(LI)")
		}

		list.LI = append(list.LI, li)
	}

	return &list, nil
}

func (n *Node) xmlLiStruct() (*LI, error) {
	li := &LI{Items: []interface{}{}}

	for _, i := range n.Children {
		node := i

		if i.Tag == "span" {
			node = &Node{
				Tag:  "p",
				Text: i.Text,
			}
		} else {
			item, err := node.xmlStruct()
			if err != nil {
				return nil, errors.Wrap(err, "i.xmlStruct")
			}

			li.Items = append(li.Items, item)
		}
	}

	return li, nil
}

func (n *Node) xmlPStruct() (*Paragraph, error) {
	p := &Paragraph{}

	for _, i := range n.Children {
		p.Texts = append(p.Texts, &Text{
			Text: i.Text,
			Style: Style{
				IsBold:   i.Tag == "b",
				IsItalic: i.Tag == "i",
			},
		})
	}

	return p, nil
}

func (n *Node) xmlStruct() (interface{}, error) {
	if n.Tag == "ul" || n.Tag == "ol" {
		list, err := n.xmlListStruct()
		if err != nil {
			return nil, errors.Wrap(err, "n.xmlListStruct")
		}

		return list, nil
	}

	if n.Tag == "p" {
		p, err := n.xmlPStruct()
		if err != nil {
			return nil, errors.Wrap(err, "n.xmlPStruct")
		}

		return p, nil
	}

	if n.Tag == "li" {
		li, err := n.xmlLiStruct()
		if err != nil {
			return nil, errors.Wrap(err, "n.xmlLiStruct")
		}

		return li, nil
	}

	return nil, nil
}
