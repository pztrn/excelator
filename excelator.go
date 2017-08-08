// Excelator - xlsx file filtrator.
//
// Copyright (c) 2017, Stanislav N. aka pztrn.
// All rights reserved.
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.

package main

import (
    // stdlib
    "flag"
    "log"
    "strconv"
    "strings"
    "time"

    // other
    "github.com/360EntSecGroup-Skylar/excelize"
)

var (
    filename string
    keyword string
)

func main() {
    flag.StringVar(&filename, "filename", "", "File to parse.")
    flag.StringVar(&keyword, "keyword", "", "Keyword to search for.")

    flag.Parse()

    if keyword == "" {
        log.Fatal("Please specify keyword to search with '-keyword'! See help for more information.")
    }

    log.Println("Excelator is starting...")
    log.Println("Will try to parse file", filename)

    startTime := time.Now().UTC()
    // Open file and read.
    log.Println("Reading file...")
    xlsx, err := excelize.OpenFile(filename)
    if err != nil {
        panic(err)
    }

    var valid_rows [][]string
    rows := xlsx.GetRows("sheet1")
    log.Println("Got", len(rows), "rows for XLSX file. Starting searching...")
    for _, row := range rows {
        if isRowContainsString(row, keyword) {
            valid_rows = append(valid_rows, row)
        }
    }

    log.Println("Found", len(valid_rows), "rows")

    log.Println("Writing to 'filtered.xlsx' near binary...")

    filtered := excelize.NewFile()

    // Row number.
    row_idx := 1
    for _, row := range valid_rows {
        letters_iteration := 0

        cell_idx := 1
        for _, cell := range row {
            // Get letter
            letter_code := 64 + cell_idx
            if letter_code > 90 {
                letters_iteration += 1
                letter_code = 65
                cell_idx = 1
            }

            var first_letter int = 0
            if letters_iteration != 0 {
                first_letter = 64 + letters_iteration
            }

            if first_letter == 0 {
                filtered.SetCellValue("sheet1", string(letter_code) + strconv.Itoa(row_idx), cell)
            } else {
                filtered.SetCellValue("sheet1", string(first_letter) + string(letter_code) + strconv.Itoa(row_idx), cell)
            }

            cell_idx += 1
        }

        row_idx += 1
    }

    err1 := filtered.SaveAs("./filtered.xlsx")
    if err1 != nil {
        panic(err1)
    }

    log.Println("Done in", time.Since(startTime))
}

func isRowContainsString(row []string, str string) bool {
    for _, cell := range row {
        if strings.Contains(cell, str) {
            return true
        }
    }

    return false
}
