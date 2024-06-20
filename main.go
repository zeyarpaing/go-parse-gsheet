package main

import (
	"context"
	"encoding/json"
	"fmt"
	"log"
	"net/http"
	"strconv"

	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

func columnLetter(col int64) string {
	result := ""
	for col > 0 {
		col-- // Adjust because column index is 1-based
		result = string('A'+(col%26)) + result
		col /= 26
	}
	return result
}

type Response struct {
	Message string `json:"message"`
	Status  string `json:"status"`
	Data    any    `json:"data"`
}

func sendResponse(w http.ResponseWriter, response Response) {
	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(response)
}

func helloHandler(w http.ResponseWriter, r *http.Request) {
	if r.Method == http.MethodGet {
		queryParams := r.URL.Query()
		spreadSheetId := queryParams.Get("spreadsheet_id")
		sheetId := queryParams.Get("sheet_id")

		// parse string to int
		sheetIdInt, err := strconv.Atoi(sheetId)

		if err != nil {
			log.Printf("Unable to convert sheet_id to int %v", err)
			sendResponse(w, Response{
				Message: "Invalid sheet_id",
				Status:  "error",
				Data:    nil,
			})
			return
		}

		data, errorMessage := readGoogleSheet(spreadSheetId, sheetIdInt)
		if errorMessage != nil {
			sendResponse(w, Response{
				Message: errorMessage.Error(),
				Status:  "error",
				Data:    data,
			})
			return
		}
		sendResponse(w, Response{
			Message: "",
			Status:  "success",
			Data:    data,
		})

	} else {
		http.Error(w, "Invalid request method", http.StatusMethodNotAllowed)
	}
}

func main() {
	http.HandleFunc("/sheet-data", helloHandler)
	http.ListenAndServe(":8080", nil)
	log.Println("Server started on port 8080")
}

func readGoogleSheet(spreadsheetID string, sheetId int) ([][]interface{}, error) {
	sheetService, err := NewSpreadsheetService("service-account.json")
	if err != nil {
		log.Printf("Unable to read service account key file  %v", err)
	}
	// spreadsheetID := "15bX0S72f3EFJuZz1sGScxp4_AREdRCaBq8YgYN2LJWE"
	// sheetIdx := 0

	doc, err := sheetService.service.Spreadsheets.Get(spreadsheetID).Do()

	if err != nil {
		return [][]interface{}{}, fmt.Errorf("unable to retrieve spreadsheet")
	}

	sheet := &sheets.Sheet{}
	hasFound := false
	for _, s := range doc.Sheets {
		if s.Properties.SheetId == int64(sheetId) {
			sheet = s
			hasFound = true
			break
		}
	}

	if !hasFound {
		return [][]interface{}{}, fmt.Errorf("sheet not found")
	}

	sheetTitle := sheet.Properties.Title
	readRange := fmt.Sprintf("%s!A1:%s%d", sheetTitle, columnLetter(sheet.Properties.GridProperties.ColumnCount), sheet.Properties.GridProperties.RowCount)

	resp, err := sheetService.service.Spreadsheets.Values.Get(spreadsheetID, readRange).Do()
	if err != nil {
		return [][]interface{}{}, fmt.Errorf("unable to retrieve data from sheet")
	}

	if len(resp.Values) == 0 {
		return [][]interface{}{}, fmt.Errorf("no data found in sheet")
	}

	vals := resp.Values
	return vals, nil
}

type SpreadsheetService struct {
	service *sheets.Service
}

func NewSpreadsheetService(path string) (*SpreadsheetService, error) {
	// Service account based oauth2 two legged integration
	sheet := &SpreadsheetService{}
	ctx := context.Background()
	srv, err := sheets.NewService(ctx, option.WithCredentialsFile(path), option.WithScopes(sheets.SpreadsheetsReadonlyScope))
	if err != nil {
		return sheet, fmt.Errorf("unable to retrieve sheets client %v", err)
	}
	sheet.service = srv
	return sheet, nil
}

func (s *SpreadsheetService) WriteToSpreadsheet(object *SpreadsheetPushRequest) error {

	var vr sheets.ValueRange
	vr.Values = append(vr.Values, object.Values)

	_, err := s.service.Spreadsheets.Values.Append(object.SpreadsheetId, object.Range, &vr).ValueInputOption("RAW").Do()
	if err != nil {
		return fmt.Errorf("unable to update data to sheet %v", err)
	}
	return nil
}

type SpreadsheetPushRequest struct {
	SpreadsheetId string        `json:"spreadsheet_id"`
	Range         string        `json:"range"`
	Values        []interface{} `json:"values"`
}
