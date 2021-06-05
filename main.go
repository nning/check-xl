package main

import (
	"flag"
	"fmt"
	"io/ioutil"
	"os"
	"path"

	log "github.com/sirupsen/logrus"
	"github.com/tealeg/xlsx"
)

var removeUnparsable = flag.Bool("r", false, "Remove file if it could not be parsed")
var listSheets = flag.Bool("s", false, "List sheets contained in files")

func printHelp() {
	fmt.Fprintf(os.Stderr, `%s [options] <dir>

Options
  -s  List sheets contained in files

`, os.Args[0])

	os.Exit(1)
}

func main() {
	flag.Parse()

	if len(flag.Args()) < 1 || len(flag.Args()) > 2 {
		printHelp()
	}

	dir := flag.Args()[0]

	files, err := ioutil.ReadDir(dir)
	if err != nil {
		log.Fatal(err)
	}

	for _, f := range files {
		file := f.Name()
		filePath := path.Join(dir, file)

		xlFile, err := xlsx.OpenFile(filePath)

		if err != nil {
			log.Error("Could not parse ", filePath)

			if *removeUnparsable {
				os.Remove(filePath)
			}
		} else {
			if *listSheets {
				fmt.Println(filePath)

				for _, sheet := range xlFile.Sheets {
					fmt.Println("\t" + sheet.Name)
				}
			}
		}
	}
}
