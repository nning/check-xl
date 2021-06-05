package main

import (
	"flag"
	"fmt"
	"io/ioutil"
	"os"
	"path"
	"sync"

	log "github.com/sirupsen/logrus"
	"github.com/tealeg/xlsx"
)

var removeUnparsable = flag.Bool("r", false, "Remove file if it could not be parsed")
var listSheets = flag.Bool("s", false, "List sheets contained in files")

var ch = make(chan fileSummary, 1000)

type fileSummary struct {
	Path   string
	Sheets []*xlsx.Sheet
}

func (f fileSummary) String() string {
	s := f.Path + "\n"
	for _, sheet := range f.Sheets {
		s += "\t" + sheet.Name + "\n"
	}

	return s
}

func printHelp() {
	fmt.Fprintf(os.Stderr, `%s [options] <dir>

Options
  -r  Remove file if it could not be parsed
  -s  List sheets contained in files

`, os.Args[0])

	os.Exit(1)
}

func process(wg *sync.WaitGroup, filePath string) {
	defer wg.Done()

	xlFile, err := xlsx.OpenFile(filePath)

	if err != nil {
		log.Error("Could not parse ", filePath)

		if *removeUnparsable {
			os.Remove(filePath)
		}
	} else {
		if *listSheets {
			ch <- fileSummary{Path: filePath, Sheets: xlFile.Sheets}
		}
	}
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

	var wg sync.WaitGroup

	for _, f := range files {
		file := f.Name()
		filePath := path.Join(dir, file)

		wg.Add(1)
		go process(&wg, filePath)
	}

	if *listSheets {
		go func() {
			for c := range ch {
				fmt.Println(c)
			}
		}()
	}

	wg.Wait()
	close(ch)
}
