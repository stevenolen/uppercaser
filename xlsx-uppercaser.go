package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"gopkg.in/urfave/cli.v1"
	"log"
	"os"
	"strconv"
	"strings"
)

func saveFile(file *excelize.File, currentFile string, overwrite bool) {
	var filename string
	if overwrite {
		filename = currentFile
	} else {
		filename = strings.TrimSuffix(currentFile, ".xlsx") + "-upper.xlsx"
	}
	file.SaveAs(filename)
}

func uppercaseFile(file string, overwrite bool) {
	xlsx, err := excelize.OpenFile(file)
	if err != nil {
		log.Fatal("Not a valid excel file: %s", file)
		return
	}

	// each sheet in spreadsheet
	for i := 1; i <= xlsx.SheetCount; i++ {
		sheetName := xlsx.GetSheetName(i)
		log.Printf("Working on File:%s,Sheet:%s", file, sheetName)

		// each row in sheet
		rows := xlsx.GetRows(sheetName)
		for r, row := range rows {
			// each cell in row
			for c, cellContents := range row {
				// columns are alphabetish, Column 1 == Column A, 27 == AA
				col := excelize.ToAlphaString(c)
				// compute the current cell's axis so we can write back to it
				// (not sure why this isn't feature in the lib)
				axis := fmt.Sprintf("%s%d", col, r+1)
				//log.Printf("Working on axis:%s, contents:%s\n",axis,cellContents)
				xlsx.SetCellStr(sheetName, axis, strings.ToUpper(cellContents))
			}
		}
	}
	saveFile(xlsx, file, overwrite)
}

func uppercase(args []string, overwrite bool) {
	for _, currentFile := range args {
		uppercaseFile(currentFile, overwrite)
	}
}

func main() {
	var overwrite string

	app := cli.NewApp()
	app.Name = "xlsx-uppercaser"
	app.Usage = "Given any number of excel files, make every cell uppercase. Don't ask why."

	flags := []cli.Flag{
		cli.StringFlag{
			Name:        "overwrite, o",
			Value:       "false",
			Usage:       "Overwrite existing file",
			Destination: &overwrite,
		},
	}

	app.Action = func(c *cli.Context) error {
		overwrite_bool, _ := strconv.ParseBool(overwrite)
		uppercase(c.Args(), overwrite_bool)
		return nil
	}

	app.Flags = flags

	err := app.Run(os.Args)
	if err != nil {
		log.Fatal(err)
	}
}
