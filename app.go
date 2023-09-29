package main

import (
	"fmt"
	"github.com/skip2/go-qrcode"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"strings"
)

func main() {
	// get command line arguments
	args := os.Args[1:]
	if len(args) != 2 {
		fmt.Println("Usage: ./app <in excel file> <out excel file>")
		return
	}

	inExcelFile := args[0]
	outExcelFile := args[1]

	tempFile := "./qr_temp.png"

	// Open the Excel file
	f, err := excelize.OpenFile(inExcelFile)
	if err != nil {
		log.Fatal(err)
	}

	// Get the sheet name
	sheetName := f.GetSheetName(0)

	// Get the rows in the sheet
	rows, err := f.GetRows(sheetName)
	if err != nil {
		log.Fatal(err)
	}

	// Iterate through the rows and print them
	for index, row := range rows {
		if index == 0 {
			continue
		}
		// concat columns B to E, trim value then join with comma
		var concat string
		temp := make([]string, 4)
		for i := 1; i <= 4; i++ {
			temp[i-1] = strings.TrimSpace(row[i])
		}
		concat = strings.Join(temp, ",")
		fmt.Printf("%s \n", concat)

		// make a QR code image from the value of concat
		png, err := qrcode.Encode(concat, qrcode.Medium, 33)
		if err != nil {
			return
		}

		// write png to temp file
		err = os.WriteFile(tempFile, png, 0644)
		if err != nil {
			return
		}

		// save the QR code image to column F of the same row
		err = f.AddPicture(sheetName, fmt.Sprintf("F%d", index+1), tempFile, nil)
	}

	// delete the temp file
	err = os.Remove(tempFile)

	// save the file
	err = f.SaveAs(outExcelFile)
	if err != nil {
		log.Fatal(err)
	}
}
