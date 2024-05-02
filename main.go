package main

import (
	"fmt"
	"strings"

	"github.com/gocolly/colly"
	"github.com/xuri/excelize/v2"
)

type Word struct {
	enWord string `json:"en"`
	ruWord string `json:"ru"`
}

var wordCollection = []Word{}
var cnt = 0

func main() {
	scrapPage("http://en365.ru/top1000.htm")
	scrapPage("http://en365.ru/top1000a.htm")
	fmt.Printf("%s \n", wordCollection)
	writeResultXls()
}

func scrapPage(url string) {
	c := colly.NewCollector()

	// Find and visit all links
	c.OnHTML("td > table td", func(e *colly.HTMLElement) {
		enWord := e.DOM.Find("td:nth-child(2)").Text()
		ruWord := e.DOM.Find("td:nth-child(3)").Text()
		if !strings.Contains(enWord, "Английское") {
			wordCollection = append(wordCollection, Word{enWord, ruWord})
			cnt++
		}
	})

	c.Visit(url)
}

func writeResultXls() {
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// Create a new sheet.
	index, err := f.NewSheet("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}
	// Set value of a cell.
	for i, word := range wordCollection {
		f.SetCellValue("Sheet1", fmt.Sprintf("A%v", i+1), word.enWord)
		f.SetCellValue("Sheet1", fmt.Sprintf("B%v", i+1), word.ruWord)
	}

	// Set active sheet of the workbook.
	f.SetActiveSheet(index)
	// Save spreadsheet by the given path.
	if err := f.SaveAs("./Vocabluary.xlsx"); err != nil {
		fmt.Println(err)
	}
}
