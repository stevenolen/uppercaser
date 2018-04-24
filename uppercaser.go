package main

import (
	"fmt"
	"path/filepath"
	"bufio"
	"github.com/360EntSecGroup-Skylar/excelize"
	"gopkg.in/urfave/cli.v1"
	"log"
	"os"
	"strconv"
	"strings"
)

func replaceFile(originalFile string, upperFile string) {
	err := os.Rename(upperFile, originalFile)
	if err != nil {
        log.Fatal(err)
    }
}

func upperFilename(currentFile string) string {
	var outputFilename string
	ext := filepath.Ext(currentFile)
	outputFilename = strings.TrimSuffix(currentFile, ext) + "-upper" + ext
	return outputFilename
}

func uppercaseFile(fileLoc string, upperFileLoc string) {
	file, err := os.Open(fileLoc)

    if err != nil {
        log.Fatal(err)
    }
    defer file.Close()

	ofile, err := os.Create(upperFileLoc)
	if err != nil {
        log.Fatal(err)
    }
    defer ofile.Close()
	w := bufio.NewWriter(ofile)

    scanner := bufio.NewScanner(file)

    for scanner.Scan() {
		w.WriteString(strings.ToUpper(scanner.Text()+"\n"))
    }
	w.Flush()
}

func uppercaseXlsx(xlsx *excelize.File, upperFile string) {
	// each sheet in spreadsheet
	for i := 1; i <= xlsx.SheetCount; i++ {
		sheetName := xlsx.GetSheetName(i)
		log.Printf("Working on Sheet:%s", sheetName)

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
	xlsx.SaveAs(upperFile)
}

func uppercase(args []string, overwrite bool) {
	for _, currentFile := range args {
		u := upperFilename(currentFile)
		//uppercaseCsv(currentFile, overwrite)
		xlsx, err := excelize.OpenFile(currentFile)
		if err != nil {
			log.Printf("Not a valid excel file: %s\n", currentFile)
			uppercaseFile(currentFile, u)
		} else {
			uppercaseXlsx(xlsx, u)
		}

		if overwrite {
			replaceFile(currentFile, u)
		}

	}
}

func main() {
	var overwrite string

	app := cli.NewApp()
	app.Name = "uppercaser"
	app.Usage = "Given any number of files, make everything uppercase. Don't ask why."

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
