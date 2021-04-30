package main

import (
	"fmt"
	"log"
	"os"
	"path"
	"path/filepath"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

func index(slice []string, item string) int {
	for i, _ := range slice {
		if slice[i] == item {
			return i
		}
	}
	return -1
}

func getFileName(URL string) string {
	s := strings.Split(URL, "/")

	return s[len(s)-1]
}

func getFileExtension(filename string) string {
	s := strings.Split(filename, ".")
	return s[len(s)-1]
}

func main() {
	const URL_HEADER = "Ссылки"
	const FILE_NAME_FIELD = "Артикул"

	exectable, _ := os.Executable()
	_path := filepath.Dir(exectable)

	xlsx_file_path := ""

	if _path == "MacOS" {
		xlsx_file_path = filepath.Dir(filepath.Dir(filepath.Dir(_path)))
	}

	xlsx_file_name := path.Join(xlsx_file_path, "таблица соответствия.xlsx")
	f, err := excelize.OpenFile(xlsx_file_name)
	if err != nil {
		fmt.Println(err)
		return
	}

	var headers []string
	firstSheet := f.WorkBook.Sheets.Sheet[0].Name
	rows, err := f.GetRows(firstSheet)
	for _, colCell := range rows[0] {
		headers = append(headers, colCell)
	}
	links_column_index := index(headers, URL_HEADER)
	filename_column_index := index(headers, FILE_NAME_FIELD)

	mapping := make(map[string]string)

	for i := 1; i <= len(rows)-1; i++ {
		link := rows[i][links_column_index]
		original_filename := getFileName(link)

		filename := rows[i][filename_column_index]
		fmt.Println(original_filename, "\t")
		fmt.Printf(filename)
		mapping[original_filename] = filename
	}

	for filename, new_filename := range mapping {

		ext := getFileExtension(filename)
		new_filename = new_filename + "." + ext
		err := os.Rename(filename, new_filename)
		if err != nil {
			log.Fatal(err)
		}
	}

}
