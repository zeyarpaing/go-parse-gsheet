package main

import (
	"context"
	"fmt"
	"log"

	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

func main() {
	sheet, err := NewSpreadsheetService("service-account.json")
	if err != nil {
		log.Printf("Unable to read service account key file  %v", err)
	}
	val := sheet.service.Spreadsheets.Values.Get("15bX0S72f3EFJuZz1sGScxp4_AREdRCaBq8YgYN2LJWE", "Sheet1!A1:D1")

	fmt.Println("Values", val)
	// data := SpreadsheetPushRequest{
	// 	SpreadsheetId: "$your_sheet_id",
	// 	Range:         "Sheet1!A1:D1",
	// 	Values:        []interface{}{"lek", 18, 150, 53},
	// }
	// err = sheet.WriteToSpreadsheet(&data)
	// if err != nil {
	// log.Printf("WriteToSpreadsheet error: %v", err)
	// }
}

type SpreadsheetService struct {
	service *sheets.Service
}

func NewSpreadsheetService(path string) (*SpreadsheetService, error) {
	// Service account based oauth2 two legged integration
	sheet := &SpreadsheetService{}
	ctx := context.Background()
	srv, err := sheets.NewService(ctx, option.WithCredentialsFile(path), option.WithScopes(sheets.SpreadsheetsScope))
	if err != nil {
		return sheet, fmt.Errorf("Unable to retrieve Sheets Client %v", err)
	}
	sheet.service = srv
	return sheet, nil
}

func (s *SpreadsheetService) WriteToSpreadsheet(object *SpreadsheetPushRequest) error {

	var vr sheets.ValueRange
	vr.Values = append(vr.Values, object.Values)

	_, err := s.service.Spreadsheets.Values.Append(object.SpreadsheetId, object.Range, &vr).ValueInputOption("RAW").Do()
	if err != nil {
		return fmt.Errorf("Unable to update data to sheet  ", err)
	}
	return nil
}

type SpreadsheetPushRequest struct {
	SpreadsheetId string        `json:"spreadsheet_id"`
	Range         string        `json:"range"`
	Values        []interface{} `json:"values"`
}
