package main

import (
	"fmt"
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

	xlsx_file_path := _path

	if _path == "MacOS" {
		xlsx_file_path = filepath.Dir(filepath.Dir(filepath.Dir(_path)))
	}

	xlsx_file_name := path.Join(xlsx_file_path, "ссылки.xlsx")
	fmt.Println("Excel file: " + xlsx_file_name)
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

	var errors []string
	existing_files := make(map[string]int)

	for filename, new_filename := range mapping {

		if val, ok := existing_files[new_filename]; ok {

			existing_files[new_filename] = val + 1
			new_filename = fmt.Sprintf("%s_%d", new_filename, val)

		} else {
			existing_files[new_filename] = 1
		}

		ext := getFileExtension(filename)
		fmt.Printf("Переименование файла: %s -> %s\n", filename, new_filename)

		new_filename = path.Join(xlsx_file_path, new_filename+"."+ext)
		err := os.Rename(path.Join(xlsx_file_path, filename), new_filename)
		if err != nil {
			errors = append(errors, filename+"\t"+err.Error())
		}
	}
	if len(errors) > 0 {
		fmt.Println("Ошибки ниже")
		for _, err := range errors {
			fmt.Println(err)
		}
	} else {
		fmt.Println("Ошибок нет")
	}
}
