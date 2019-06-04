package drive

// NewSheetService creates a new google sheets service
import (
	"fmt"
	"io/ioutil"
	"log"

	"github.com/spf13/viper"

	"golang.org/x/oauth2/google"
	sheets "google.golang.org/api/sheets/v4"
)

// SheetService is a wrapper around the google sheets service so we can attach nice funcs to it and maybe hold other
// metadata in it.
type SheetService struct {
	service *sheets.Service
	sheetID string
}

// NewSheetService creates a new service for google sheets. It also handles dealing with OAuth shenanigans.
func NewSheetService(sheetID string) (srv SheetService, err error) {
	secretFileName := viper.GetString("drive.clientSecretFileName")

	b, err := ioutil.ReadFile(secretFileName)
	if err != nil {
		return srv, fmt.Errorf("Unable to read client secret file: %v", err)
	}

	// If modifying these scopes, delete your previously saved client_secret.json.
	// append .readonly to the url below to request readonly to a sheet
	googleHost := viper.GetString("drive.host")
	config, err := google.ConfigFromJSON(b, googleHost)
	if err != nil {
		return srv, fmt.Errorf("Unable to parse client secret file to config: %v", err)
	}

	client := getClient(config)

	srv.service, err = sheets.New(client)
	srv.sheetID = sheetID

	return
}

// SendDataTable sends a 2d string array table to the specified google drive
func (srv SheetService) SendDataTable(upperLeftCell string, table [][]string) error {
	writeCell := upperLeftCell
	var vr sheets.ValueRange

	interfaceVals := make([][]interface{}, len(table))
	for i, row := range table {
		interfaceRow := make([]interface{}, len(row))
		for j, cell := range row {
			interfaceRow[j] = cell
		}
		interfaceVals[i] = interfaceRow
	}
	vr.Values = interfaceVals

	inputOption := viper.GetString("drive.inputValueOption")
	_, err := srv.service.Spreadsheets.Values.Update(srv.sheetID, writeCell, &vr).ValueInputOption(inputOption).Do()
	if err != nil {
		log.Fatalf("Unable to update data in sheet. %v", err)
	}

	return nil
}

// ClearTableArea clear data values from a range in a spreadsheet
// The range notation uses A1 notation
func (srv SheetService) ClearTableArea(rangeNotation string) error {
	// Genuinely don't know what this does, we need to research it
	rb := &sheets.ClearValuesRequest{}

	//_, err := srv.service.Spreadsheets.Values.C(srv.sheetID, writeCell, &vr).ValueInputOption(inputOption).Do()
	_, err := srv.service.Spreadsheets.Values.Clear(srv.sheetID, rangeNotation, rb).Do()
	if err != nil {
		log.Fatalf("Unable to update data in sheet. %v", err)
	}

	return nil
}
