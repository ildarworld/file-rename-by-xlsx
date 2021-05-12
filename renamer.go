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
	const XLSX_FILE_NAME = "ссылки.xlsx"

	is_wb := false
	fmt.Print("Если нужно переименовать файлы для  Wildberries введите 'W'.\n Или просто нажмите на Enter: ")
	var input string
	fmt.Scanln(&input)
	fmt.Print(input)

	if strings.ToUpper(input) == "W" {
		is_wb = true
		fmt.Println("\nПереименовываем файлы под требования Wildberries:")
	}

	exectable, _ := os.Executable()
	_path := filepath.Dir(exectable)

	xlsx_file_path := _path

	if _path == "MacOS" {
		xlsx_file_path = filepath.Dir(filepath.Dir(filepath.Dir(_path)))
	}

	xlsx_file_name := path.Join(xlsx_file_path, XLSX_FILE_NAME)
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

		ext := getFileExtension(filename)
		old_filename := path.Join(xlsx_file_path, filename)
		new_filename_path := ""
		wb_file_path := ""

		if is_wb {
			wb_file_path = path.Join(new_filename, "photo")
			os.MkdirAll(path.Join(xlsx_file_path, wb_file_path), 0755)
		}

		if val, ok := existing_files[new_filename]; ok {
			existing_files[new_filename] = val + 1
			new_filename_path = path.Join(xlsx_file_path, wb_file_path, fmt.Sprintf("%s_%d.%s", new_filename, val, ext))
		} else {
			existing_files[new_filename] = 1
			new_filename_path = path.Join(xlsx_file_path, wb_file_path, fmt.Sprintf("%s.%s", new_filename, ext))
		}

		fmt.Printf("Переименование файла: %s -> %s\n", filename, new_filename)

		err := os.Rename(old_filename, new_filename_path)
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
