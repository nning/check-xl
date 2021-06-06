package main

import (
	"flag"
	"fmt"
	"io/ioutil"
	"os"
	"path"
	"runtime/debug"
	"strings"
	"sync"

	"github.com/extrame/xls"
	"github.com/tealeg/xlsx"

	log "github.com/sirupsen/logrus"
)

var removeUnparsable = flag.Bool("r", false, "Remove file if it could not be parsed")
var listSheets = flag.Bool("s", false, "List sheets contained in files")

var ch = make(chan fileSummary, 1000)

type fileSummary struct {
	Path   string
	Sheets []string
}

func (f fileSummary) String() string {
	s := f.Path + "\n"
	for _, sheet := range f.Sheets {
		s += "\t" + sheet + "\n"
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

func handleParseError(filePath string) {
	log.Error("Could not parse ", filePath)

	if *removeUnparsable {
		os.Remove(filePath)
	}
}

func xlsGetSheets(filePath string) []string {
	xlsFile, err := xls.Open(filePath, "utf-8")

	if err != nil {
		handleParseError(filePath)
		return []string{}
	}

	sheets := make([]string, 0)
	for i := 0; i < xlsFile.NumSheets(); i++ {
		sheets = append(sheets, xlsFile.GetSheet(i).Name)
	}

	return sheets
}

func xlsxGetSheets(filePath string) []string {
	xlsxFile, err := xlsx.OpenFile(filePath)

	if err != nil {
		handleParseError(filePath)
		return []string{}
	}

	sheets := make([]string, 0)
	for _, sheet := range xlsxFile.Sheets {
		sheets = append(sheets, sheet.Name)
	}

	return sheets
}

func process(wg *sync.WaitGroup, filePath string) {
	defer wg.Done()

	debug.SetPanicOnFault(true)
	defer func() {
		if p := recover(); p != nil {
			handleParseError(filePath)
			return
		}
	}()

	extension := strings.Split(filePath, ".")[1]
	var sheets []string

	if extension == "xls" {
		sheets = xlsGetSheets(filePath)
	}

	if extension == "xlsx" {
		sheets = xlsxGetSheets(filePath)
	}

	if *listSheets {
		ch <- fileSummary{filePath, sheets}
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
