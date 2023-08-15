package main

import (
	"github.com/krestkrest/go-xlsx-templater"
)

func main() {
	doc := xlst.New()
	if err := doc.ReadTemplate("./template.xlsx"); err != nil {
		panic(err)
	}
	ctx := map[string]interface{}{
		"name":           "Github User",
		"nameHeader":     "Item name",
		"quantityHeader": "Quantity",
		"items": []map[string]interface{}{
			{
				"name":     "Pen",
				"quantity": 2,
			},
			{
				"name":     "Pencil",
				"quantity": 1,
			},
			{
				"name":     "Condom",
				"quantity": 12,
			},
			{
				"name":     "Beer",
				"quantity": 24,
			},
		},
	}

	if err := doc.Render(ctx); err != nil {
		panic(err)
	}
	if err := doc.Save("./report.xlsx"); err != nil {
		panic(err)
	}
}
