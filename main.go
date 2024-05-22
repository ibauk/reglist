package main

/*
 * This is a validator/transformer to create a "Registration list" spreadsheet ready
 * for the RBLR1000 and IBA scatter rallies.
 *
 * It will be run several times before a "final" version shortly before the ride date.
 *
 * It must be kept in sync with the Wufoo form used to capture entrant records.
 *
 */

import (
	"database/sql"
	"encoding/csv"
	"flag"
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"time"

	_ "github.com/mattn/go-sqlite3"
	"github.com/xuri/excelize/v2"
)

var rally *string = flag.String("cfg", "", "Which rally is this (yml file)")
var csvName *string = flag.String("csv", "", "Path to CSV downloaded from Wufoo")
var csvReport *bool = flag.Bool("rpt", true, "CSV is downloaded from Wufoo report")
var csvAdmin *bool = flag.Bool("adm", false, "CSV is downloaded from Wufoo administrator page")
var sqlName *string = flag.String("sql", "entrantdata.db", "Path to SQLite database")
var xlsName *string = flag.String("xls", "", "Path to output XLSX, defaults to cfg name+year")
var noCSV *bool = flag.Bool("nocsv", false, "Don't load a CSV file, just use the SQL database")
var safemode *bool = flag.Bool("safe", true, "Safe mode avoid formulas, no live updating")
var livemode *bool = flag.Bool("live", false, "Self-updating, live mode")
var expReport *string = flag.String("exp", "", "Path to output standard format CSV")
var expGmail *string = flag.String("gmail", "", "Path to CSV output for Gmail")
var ridesdb *string = flag.String("rd", "", "Path of rides database for lookup")
var noLookup *bool = flag.Bool("nolookup", false, "Don't lookup unidentified IBA members")
var summaryOnly *bool = flag.Bool("summary", true, "Produce Summary/overview tabs only")
var allTabs *bool = flag.Bool("full", false, "Generate all tabs")
var showusage *bool = flag.Bool("?", false, "Show this help")
var verbose *bool = flag.Bool("v", false, "Verbose mode, debugging")

const apptitle = "IBAUK Reglist v1.27\nCopyright (c) 2024 Bob Stammers\n\n"
const progdesc = `I parse and enhance rally entrant records in CSV format downloaded from Wufoo forms either 
using the admin interface or one of the reports. I output a spreadsheet in XLSX format of
the records presented in various useful ways and, optionally, a CSV containing the enhanced
data in a format suitable for input to a ScoreMaster database and, optionally, a CSV suitable for import to
a Gmail account.
`

var rblr_routes = [...]string{" A-NC", " B-NAC", " C-SC", " D-SAC", " E-500C", " F-500AC"}
var rblr_routes_ridden = [...]int{0, 0, 0, 0, 0, 0}

// cancelsLoseOut determines whether entrants with Paid=Cancelled lose T-shirts, camping and patches
// If so, they aren't counted and moneys paid are added to sponsorship
const cancelsLoseOut = false

const max_tshirt_sizes int = 10

var tshirt_sizes [max_tshirt_sizes]string

var overview_patch_column, shop_patch_column string

// dbFields must be kept in sync with the downloaded CSV from Wufoo
// Fieldnames don't matter but the order and number both do

var dbfieldsx string

var overviewsheet string = "Overview"
var noksheet string = "Contacts"
var paysheet string = "Money"
var subssheet string = "Sponsorship"

// The Stats sheet (totsheet) needs to be first as otherwise Google Sheets
// doesn't show the chart. Much diagnostic phaffery has led me to this
// workaround rather than actually diagnosing the fault so shoot me.
var totsheet string = "Sheet1" // Renamed on init
var chksheet string = "Carpark"
var regsheet string = "Registration"
var shopsheet string = "Shop"

const sqlx_rblr = `ifnull(RiderName,''),ifnull(RiderLast,''),ifnull(RiderIBANumber,''),ifnull(RiderRBL,''),
ifnull(PillionName,''),ifnull(PillionLast,''),ifnull(PillionIBANumber,''),ifnull(PillionRBL,''),
ifnull(BikeMakeModel,''),round(ifnull(MilesTravelledToSquires,'0')),
ifnull(FreeCamping,''),ifnull(WhichRoute,'A'),
ifnull(Tshirt1,''),ifnull(Tshirt2,''),ifnull(Patches,'0'),ifnull(Cash,'0'),
ifnull(Mobilephone,''),
ifnull(NOKName,''),ifnull(NOKNumber,''),ifnull(NOKRelation,''),
FinalRiderNumber,ifnull(PaymentTotal,''),ifnull(Sponsorshipmoney,''),ifnull(PaymentStatus,''),
ifnull(NoviceRider,''),ifnull(NovicePillion,''),ifnull(odometer_counts,''),ifnull(Registration,''),
ifnull(MilestravelledToSquires,''),ifnull(FreeCamping,''),
ifnull(Address1,''),ifnull(Address2,''),ifnull(Town,''),ifnull(County,''),
ifnull(Postcode,''),ifnull(Country,''),ifnull(Email,''),ifnull(Mobilephone,''),
ifnull(Date_Created,''),ifnull(Withdrawn,''),ifnull(HasPillion,'')`

const sqlx_rally = `ifnull(RiderName,''),ifnull(RiderLast,''),ifnull(RiderIBANumber,''),
ifnull(PillionName,''),ifnull(PillionLast,''),ifnull(PillionIBANumber,''),
ifnull(BikeMakeModel,''),
ifnull(Tshirt1,''),ifnull(Tshirt2,''),
ifnull(Mobilephone,''),
ifnull(NOKName,''),ifnull(NOKNumber,''),ifnull(NOKRelation,''),
FinalRiderNumber,ifnull(PaymentTotal,''),ifnull(PaymentStatus,''),
ifnull(NoviceRider,''),ifnull(NovicePillion,''),ifnull(odometer_counts,''),ifnull(Registration,''),
ifnull(Address1,''),ifnull(Address2,''),ifnull(Town,''),ifnull(County,''),
ifnull(Postcode,''),ifnull(Country,''),ifnull(Email,''),ifnull(Mobilephone,''),
ifnull(Date_Created,''),ifnull(Withdrawn,''),ifnull(HasPillion,'')`

var sqlx string

var styleH, styleH2, styleH2L, styleT, styleV, styleV2, styleV2L, styleV2LBig, styleV3, styleW, styleCancel, styleRJ, styleRJSmall int

var cfg *Config
var words *Words

var db *sql.DB
var includeShopTab bool
var xl *excelize.File
var exportingCSV bool
var exportingGmail bool
var csvF *os.File
var csvW *csv.Writer
var csvFGmail *os.File
var csvGmail *csv.Writer
var num_tshirt_sizes int
var totTShirts [max_tshirt_sizes]int = [max_tshirt_sizes]int{0}

var tot *Totals

var totx struct {
	srow  int
	srowx string
}

var lookupOnline bool

func main() {

	if !*noCSV {
		if *csvName != "" {
			loadCSVFile()
			fixRiderNumbers()
		} else if cfg.CsvUrl != "" {
			downloadCSVFile()
			fixRiderNumbers()
		} else {
			fmt.Println("No CSV input available")
			return
		}
	}

	if *verbose {
		fmt.Println("dbg: Initialising spreadsheet")
	}
	initSpreadsheet()
	if *verbose {
		fmt.Println("dbg: Spreadsheet initialised")
	}

	if exportingCSV {
		initExportCSV()
		defer csvF.Close()
	}
	if exportingGmail {
		initExportGmail()
		defer csvFGmail.Close()
	}

	mainloop()

	if exportingCSV {
		csvW.Flush()
	}
	if exportingGmail {
		csvGmail.Flush()
	}
	if tot.NumWithdrawn > 0 {
		fmt.Printf("%v entries withdrawn\n", tot.NumWithdrawn)
	}
	fmt.Printf("%v entrants written\n", tot.NumRiders)

	writeTotals()

	setTabFormats()

	markSpreadsheet()

	// Save spreadsheet by the given path.
	if err := xl.SaveAs(*xlsName); err != nil {
		fmt.Println(err)
	}

	reportDuplicates()
}

// Alphabetic from here on down ==========================================================

func downloadCSVFile() {

	if *verbose {
		fmt.Printf("Downloading from %v\n", cfg.CsvUrl)
	}
	resp, err := http.Get(cfg.CsvUrl)
	if err != nil {
		fmt.Printf("Error downloading %v\n", err)
		return
	}

	defer resp.Body.Close()

	reader := csv.NewReader(resp.Body)

	makeSQLTable(db)

	hdrSkipped := false

	debugCount := 0

	for {
		record, err := reader.Read()

		// if we hit end of file (EOF) or another unexpected error
		if err == io.EOF {
			break
		} else if err != nil {
			fmt.Printf("\nDownloading CSV - Record - %v - Error: %v\n", record, err)
			if !hdrSkipped {
				fmt.Printf("Is %v a valid URL?\nHas the Wufoo report been flagged as Public?\n\n", cfg.CsvUrl)
				os.Exit(-4)
			}
			return
		} else if *verbose {
			fmt.Printf("CSV == (%v) %v\n", len(record), record)
		}

		if !hdrSkipped {
			hdrSkipped = true
			continue
		}

		debugCount++

		if *verbose {
			fmt.Printf("dbg: Loading %v\n", debugCount)
		}

		sqlx := "INSERT INTO entrants ("
		sqlx += dbfieldsx
		sqlx += ") VALUES("

		fx := strings.Split(dbfieldsx, ",")
		rl := len(fx)
		if len(record) < rl {
			rl = len(record)
		}

		for i := 0; i < rl; i++ {
			if i > 0 {
				sqlx += ","
			}
			if len(record[i]) == 0 || record[i] == "NULL" {
				sqlx += "null"
			} else {
				//sqlx += "\"" + record[i] + "\"" // Use " rather than ' as the data might contain single quotes anyway
				sqlx += "'" + strings.ReplaceAll(record[i], "'", "''") + "'"
			}
		}
		sqlx += ");"
		_, err = db.Exec(sqlx)
		if err != nil {
			db.Exec("COMMIT")
			fmt.Println(sqlx)
			log.Fatal(err)
		}
	}

	db.Exec("COMMIT")
	if *verbose {
		fmt.Println("dbg: Load complete")
	}

}

func fieldlistFromConfig(cols []string) string {

	var res string = ""

	for i := 0; i < len(cols); i++ {
		if i > 0 {
			res += ","
		}
		res += "\"" + cols[i] + "\""
	}

	return res
}

// fixRiderNumbers creates the field FinalRiderNumber in the database. The original EntryID is
// adjusted by cfg.Add2entrantid or overridden by RiderNumber if non-zero.
func fixRiderNumbers() {

	var old string
	var new, newseq int
	var mannum string
	var withdrawn string

	oldnew := make(map[string]int, 250) // More than enough

	if *verbose {
		fmt.Println("dbg: Reading rider numbers")
	}
	sqlx := "SELECT EntryId,ifnull(RiderNumber,''),ifnull(withdrawn,'') FROM entrants"
	rows, err := db.Query(sqlx) // There is scope for renumber alphabetically if desired.
	if err != nil {
		log.Fatal(err)
	}
	for rows.Next() {

		rows.Scan(&old, &mannum, &withdrawn)
		if withdrawn == "Withdrawn" {
			oldnew[old] = intval(old)
			continue
		}
		if cfg.RenumberCSV {
			newseq++
			new = newseq + cfg.Add2entrantid
		} else {
			new = intval(old) + cfg.Add2entrantid
		}
		if intval(mannum) > 0 {
			new = intval(mannum)
		}
		oldnew[old] = new
	}

	if *verbose {
		fmt.Println("dbg: Writing rider numbers")
	}
	tx, _ := db.Begin()
	for old, new := range oldnew {
		sqlx := "UPDATE entrants SET FinalRiderNumber=" + strconv.Itoa(new) + " WHERE EntryId='" + old + "'"
		if *verbose {
			fmt.Println(sqlx)
		}
		_, err := tx.Exec(sqlx)
		if err != nil {
			log.Fatal(err)
		}
	}
	err = tx.Commit()
	if err != nil {
		log.Fatal()
	}

	n := len(oldnew)
	fmt.Printf("%v entries loaded\n", n)
	rows.Close()
}

// formatSheet sets printed properties include page orientation and margins
func formatSheet(sheetName string, portrait bool) {

	var sz int = 9
	var ft int = 2
	var om string

	if portrait {
		om = "portrait"
	} else {
		om = "landscape"
	}

	xl.SetPageLayout(
		sheetName, &excelize.PageLayoutOptions{
			Orientation: &om,
			Size:        &sz, /* xlPaperSizeA4 (10 = xlPaperSizeA4Small!) */
			FitToHeight: &ft,
			FitToWidth:  &ft,
		})

	var marg float64 = 0.2
	xl.SetPageMargins(sheetName, &excelize.PageLayoutMarginsOptions{
		Bottom: &marg,
		Footer: &marg,
		Header: &marg,
		Left:   &marg,
		Right:  &marg,
		Top:    &marg,
	})

}

func initExportCSV() {
	csvF = makeFile(*expReport)
	csvW = makeCSVFile(csvF, false)
	fmt.Printf("Exporting CSV to %v\n", *expReport)
}

func initExportGmail() {
	csvFGmail = makeFile(*expGmail)
	csvGmail = makeCSVFile(csvFGmail, true)
	fmt.Printf("Exporting Gmail CSV to %v\n", *expGmail)
}

func initSpreadsheet() {

	xl = excelize.NewFile()

	if *xlsName == "" {
		*xlsName = cfg.Rally + cfg.Year
	}
	if filepath.Ext(*xlsName) == "" {
		*xlsName = *xlsName + ".xlsx"
	}

	fmt.Printf("Creating %v\n", *xlsName)

	initStyles()
	// First sheet is called Sheet1
	formatSheet(totsheet, false)
	xl.NewSheet(overviewsheet)
	formatSheet(overviewsheet, false)
	if !*summaryOnly {
		xl.NewSheet(regsheet)
		formatSheet(regsheet, false)
		xl.NewSheet(noksheet)
		formatSheet(noksheet, false)
		if includeShopTab {
			xl.NewSheet(shopsheet)
			formatSheet(shopsheet, false)
		}
		xl.NewSheet(paysheet)
		formatSheet(paysheet, false)
		if cfg.Rally == "rblr" {
			xl.NewSheet(subssheet)
			formatSheet(subssheet, false)
		}
		xl.NewSheet(chksheet)
		formatSheet(chksheet, true)
	}
	renameSheet(&totsheet, "Stats")

	// Set heading styles
	xl.SetCellStyle(overviewsheet, "A1", overview_patch_column+"1", styleH2L)
	if cfg.Rally == "rblr" {
		xl.SetCellStyle(overviewsheet, "K1", "R1", styleH)
		xl.SetCellStyle(overviewsheet, "E1", "E1", styleH)
		xl.SetCellStyle(overviewsheet, "H1", "H1", styleH)
		xl.SetColVisible(overviewsheet, "E", false)
		xl.SetColVisible(overviewsheet, "G:H", false)
		if !*summaryOnly {
			xl.SetColVisible(chksheet, "A", false)
		}
	} else {
		xl.SetColWidth(overviewsheet, "K", "R", 1)
		xl.SetColVisible(overviewsheet, "K:R", false)
	}
	if len(cfg.Tshirts) > 0 {
		n, _ := excelize.ColumnNameToNumber("S")
		x, _ := excelize.ColumnNumberToName(n + len(cfg.Tshirts) - 1)
		xl.SetCellStyle(overviewsheet, "S1", x+"1", styleH)
	}
	if cfg.Patchavail {
		xl.SetCellStyle(overviewsheet, overview_patch_column+"1", overview_patch_column+"1", styleH)
	}

	if !*summaryOnly {
		xl.SetCellStyle(regsheet, "A1", "I1", styleH2)
		if cfg.Rally == "rblr" {
			xl.SetCellStyle(regsheet, "J1", "L1", styleH2)
		}
		xl.SetCellStyle(noksheet, "A1", "H1", styleH2L)

		xl.SetCellStyle(paysheet, "A1", "K1", styleH2)
		if cfg.Rally == "rblr" {
			xl.SetCellStyle(subssheet, "A1", "I1", styleH2)
		}

		xl.SetCellStyle(chksheet, "A1", "C1", styleH2L)
		xl.SetCellStyle(chksheet, "D1", "E1", styleH2)

		if includeShopTab {
			xl.SetCellStyle(shopsheet, "A1", shop_patch_column+"1", styleH2)
		}
	}

}

func intval(x string) int {

	re := regexp.MustCompile(`(\d+)`)
	sm := re.FindSubmatch([]byte(x))
	if len(sm) < 2 {
		return 0
	}
	n, _ := strconv.Atoi(string(sm[1]))
	if strings.Contains(x, "-") {
		n = 0 - n
	}
	return n

}

func loadCSVFile() {

	if *verbose {
		fmt.Printf("dbg: loadCSVFile = %v\n", *csvName)
	}
	file, err := os.Open(*csvName)
	// error - if we have one exit as CSV file not right
	if err != nil {
		fmt.Printf("ERROR: %s\n", err)
		os.Exit(-3)
	}
	// now file is open - defer the close of CSV file handle until we return
	defer file.Close()
	// connect a CSV reader to the file handle - which is the actual opened
	// CSV file
	// TODO : is there an error from this to check?
	reader := csv.NewReader(file)

	makeSQLTable(db)

	hdrSkipped := false

	debugCount := 0
	for {
		record, err := reader.Read()

		// if we hit end of file (EOF) or another unexpected error
		if err == io.EOF {
			break
		} else if err != nil {
			fmt.Println("Error:", err)
			return
		}

		if !hdrSkipped {
			hdrSkipped = true
			continue
		}

		debugCount++

		if *verbose {
			fmt.Printf("dbg: Loading %v\n", debugCount)
		}

		sqlx := "INSERT INTO entrants ("
		sqlx += dbfieldsx
		sqlx += ") VALUES("

		for i := 0; i < len(record); i++ {
			if i > 0 {
				sqlx += ","
			}
			if len(record[i]) == 0 || record[i] == "NULL" {
				sqlx += "null"
			} else {
				sqlx += "\"" + record[i] + "\"" // Use " rather than ' as the data might contain single quotes anyway
			}
		}
		sqlx += ");"
		_, err = db.Exec(sqlx)
		if err != nil {
			db.Exec("COMMIT")
			fmt.Println(sqlx)
			log.Fatal(err)
		}
	}

	db.Exec("COMMIT")
	if *verbose {
		fmt.Println("dbg: Load complete")
	}

}

func makeCSVFile(f *os.File, gmail bool) *csv.Writer {

	writer := csv.NewWriter(f)
	if gmail {
		writer.Write(EntrantHeadersGmail())
	} else {
		writer.Write(EntrantHeaders())
	}
	return writer
}

func makeFile(csvname string) *os.File {

	file, err := os.Create(csvname)
	if err != nil {
		panic(err)
	}
	return file

}

func makeSQLTable(db *sql.DB) {

	var x string = ""
	re := regexp.MustCompile(`\bRiderNumber\b`)
	if !re.Match([]byte(dbfieldsx)) {
		x = ",RiderNumber"
	}
	x += ",FinalRiderNumber"

	if *verbose {
		fmt.Println("dbg: Initialising database")
	}
	db.Exec("PRAGMA foreign_keys=OFF")
	db.Exec("BEGIN TRANSACTION")
	_, err := db.Exec("DROP TABLE IF EXISTS entrants")
	if err != nil {
		log.Fatal(err)
	}

	if *verbose {
		fmt.Printf("Making entrants => %v\n", dbfieldsx)
	}
	_, err = db.Exec("CREATE TABLE entrants (" + dbfieldsx + x + " INTEGER)")
	if err != nil {
		log.Fatal(err)
	}
	_, err = db.Exec("DROP TABLE IF EXISTS rally")
	if err != nil {
		log.Fatal(err)
	}
	_, err = db.Exec(`CREATE TABLE "rally" (
		"name"	TEXT,
		"year"	TEXT,
		"extracted"	TEXT,
		"csv" TEXT
	)`)
	if err != nil {
		log.Fatal(err)
	}
	_, err = db.Exec("INSERT INTO rally (name,Year,extracted,csv) VALUES(?,?,?,?)",
		cfg.Rally,
		cfg.Year,
		time.Now().Format("Mon Jan 2 15:04:05 MST 2006"),
		filepath.Base(*csvName))
	if err != nil {
		log.Fatal(err)
	}
	if *verbose {
		fmt.Println("dbg: Database initialised")
	}

}

func markCancelledEntrants() {
	for _, r := range tot.CancelledRows {
		rx := strconv.Itoa(r)
		xl.SetCellStyle(overviewsheet, "A"+rx, "J"+rx, styleCancel)

		if !*summaryOnly {
			xl.SetCellStyle(regsheet, "A"+rx, "I"+rx, styleCancel)
			if cfg.Rally == "rblr" {
				xl.SetCellStyle(regsheet, "J"+rx, "K"+rx, styleCancel)
				xl.SetCellStyle(subssheet, "A"+rx, "I"+rx, styleCancel)
			}
			xl.SetCellStyle(noksheet, "A"+rx, "H"+rx, styleCancel)
			if includeShopTab {
				xl.SetCellStyle(shopsheet, "A"+rx, "I"+rx, styleCancel)
			}
			xl.SetCellStyle(paysheet, "A"+rx, "J"+rx, styleCancel)
			xl.SetCellStyle(chksheet, "A"+rx, "H"+rx, styleCancel)
		}
	}

}

func markSpreadsheet() {

	var creator []string = strings.Split(apptitle, "\n")

	var dp excelize.DocProperties
	dp.Created = time.Now().Format(time.RFC3339)
	dp.Modified = time.Now().Format(time.RFC3339)
	dp.Creator = creator[0]
	dp.LastModifiedBy = creator[0]
	dp.Subject = cfg.Rally
	dp.Description = "This reflects the status of " + cfg.Rally + " as at " + time.Now().UTC().Format(time.UnixDate)
	if *safemode {
		dp.Description += "\n\nThis spreadsheet holds static values only and will not reflect changed data everywhere."
	} else {
		dp.Description += "\n\nThis spreadsheet is active and will reflect changed data everywhere."
	}
	dp.Title = "Rally management spreadsheet"
	err := xl.SetDocProps(&dp)
	if err != nil {
		fmt.Printf("%v\n", err)
	}

}

func renameSheet(oldname *string, newname string) {

	xl.SetSheetName(*oldname, newname)
	*oldname = newname

}

func reportDuplicates() {

	var name, last string
	var rex int

	dupes, err := db.Query("SELECT RiderName,RiderLast,Count(*) FROM entrants WHERE Withdrawn IS NULL GROUP BY Upper(Trim(RiderLast)), Upper(Trim(RiderName)) HAVING Count(EntryID) > 1;")
	if err != nil {
		panic(err)
	}
	for dupes.Next() {
		dupes.Scan(&name, &last, &rex)
		fmt.Printf("*** Rider %v %v is entered more than once (%v times!)\n", name, last, rex)
	}

}

func reportEntriesByPeriod() {

	var reportingperiod string
	if cfg.ReportWeekly {
		reportingperiod = "week"
	} else {
		reportingperiod = "month"
	}

	sort.Slice(tot.EntriesByPeriod, func(i, j int) bool { return tot.EntriesByPeriod[i].Month > tot.EntriesByPeriod[j].Month })

	xl.SetColWidth(totsheet, "B", "B", 7)
	xl.SetColWidth(totsheet, "C", "C", 2)
	xl.SetColVisible(totsheet, "D", false)
	xl.SetColWidth(totsheet, "F", "F", 5)
	xl.SetColWidth(totsheet, "G", "G", 2)
	xl.SetColWidth(totsheet, "I", "K", 5)

	xl.SetCellValue(totsheet, "K2", "Novices")
	xl.SetCellValue(totsheet, "J2", "IBA members")
	xl.SetCellValue(totsheet, "L2", "All entries")
	xl.SetCellValue(totsheet, "I2", "British Legion")

	row := 3
	for _, p := range tot.EntriesByPeriod {
		srow := strconv.Itoa(row)
		md := strings.Split(p.Month, "-")
		mth := []string{"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}[intval(md[0])-1]
		xl.SetCellValue(totsheet, "H"+srow, mth+" "+md[1])
		xl.SetCellValue(totsheet, "L"+srow, p.Total)
		xl.SetCellValue(totsheet, "J"+srow, p.NumIBA)
		xl.SetCellValue(totsheet, "K"+srow, p.NumNovice)
		xl.SetCellValue(totsheet, "I"+srow, p.NumRBLRiders+p.NumRBLBranch)
		row++
	}

	var bTrue bool = true
	var bFalse bool = false

	chartseries := make([]excelize.ChartSeries, 0)
	xrow := strconv.Itoa(row - 1)
	row = 3
	var cols string
	if cfg.Rally == "rblr" {
		cols = "IJKL"
	} else {
		cols = "JKL"
	}

	for i := 0; i < len(cols); i++ {
		ll := cols[i : i+1]
		cs := excelize.ChartSeries{
			Name:       totsheet + `!$` + ll + `$2`,
			Categories: totsheet + `!$H$3:$H$` + xrow,
			Values:     totsheet + `!` + ll + `3:` + ll + xrow,
		}
		chartseries = append(chartseries, cs)
		row++
	}

	fmtx := excelize.Chart{
		Type: excelize.Bar,
		Format: excelize.GraphicOptions{
			ScaleX:          1.0,
			ScaleY:          1.2,
			OffsetX:         15,
			OffsetY:         10,
			PrintObject:     &bTrue,
			LockAspectRatio: true,
			Locked:          &bFalse,
		},
		Legend: excelize.ChartLegend{
			Position:      "right",
			ShowLegendKey: false,
		},
		Title: []excelize.RichTextRun{{Text: "New signups by " + reportingperiod}},
		PlotArea: excelize.ChartPlotArea{
			ShowBubbleSize:  true,
			ShowCatName:     false,
			ShowLeaderLines: false,
			ShowPercent:     true,
			ShowSerName:     true,
			ShowVal:         true,
		},
		ShowBlanksAs: "zero",
		Series:       chartseries,
	}

	err := xl.AddChart(totsheet, "N2", &fmtx)
	if err != nil {
		fmt.Printf("OMG: %v\n%v\n", err, fmtx)
	}

	xl.SetColVisible(totsheet, "H:M", false)
}

func setPagePane(sheet string) {
	xl.SetPanes(sheet, &excelize.Panes{
		Freeze:      true,
		Split:       false,
		XSplit:      0,
		YSplit:      1,
		TopLeftCell: "A2",
		ActivePane:  "bottomLeft",
		Selection:   []excelize.Selection{{SQRef: "A2:X2", ActiveCell: "A2", Pane: "bottomLeft"}},
	})
}

// setPageTitle sets each sheet, except Stats, to repeat its
// top line on each printed page
func setPageTitle(sheet string) {

	var dn excelize.DefinedName

	dn.Name = "_xlnm.Print_Titles"
	dn.RefersTo = sheet + "!$1:$1"
	dn.Scope = sheet
	xl.SetDefinedName(&dn)
}

// setTabFormats sets the page headers to repeat when printed and
// sets the appropriate print area
func setTabFormats() {

	setPageTitle(totsheet)
	setPageTitle(overviewsheet)

	if !*summaryOnly {
		setPageTitle(noksheet)
		setPageTitle(paysheet)
		setPageTitle(chksheet)
		if cfg.Rally == "rblr" {
			setPageTitle(subssheet)
		}
		setPageTitle(regsheet)
	}
	setPagePane(overviewsheet)
	if !*summaryOnly {
		setPagePane(noksheet)
		setPagePane(paysheet)
		if cfg.Rally == "rblr" {
			setPagePane(subssheet)
		}
		setPagePane(chksheet)
		setPagePane(regsheet)
	}

	if includeShopTab && !*summaryOnly {
		setPageTitle(shopsheet)
		setPagePane(shopsheet)
	}

	markCancelledEntrants()
}

func writeTotals() {

	reportEntriesByPeriod()

	// Write out totals
	xl.SetColWidth(totsheet, "A", "A", 30)
	xl.SetColWidth(totsheet, "E", "E", 15)

	xl.SetCellStyle(totsheet, "A3", "A18", styleRJ)
	xl.SetCellStyle(totsheet, "E3", "E19", styleRJ)
	for i := 3; i <= 19; i++ {
		xl.SetRowHeight(totsheet, i, 30)
	}
	xl.SetCellValue(totsheet, "A3", "Number of riders")
	xl.SetCellValue(totsheet, "A4", "Number of pillions")
	xl.SetCellValue(totsheet, "A5", "Number of "+cfg.Novice+"s")
	xl.SetCellValue(totsheet, "A6", "Number of IBA members")
	if cfg.Rally == "rblr" {
		xl.SetCellValue(totsheet, "A7", "Number of Legion members")
		xl.SetCellValue(totsheet, "A8", "of which, RBL Riders")
		xl.SetCellValue(totsheet, "A9", "Nearest to Squires")
		xl.SetCellValue(totsheet, "A10", "Furthest from Squires")
		xl.SetCellValue(totsheet, "A11", "Camping at Squires")
		xl.SetCellValue(totsheet, "A12", "Funds raised for Poppy Appeal")
		xl.SetCellValue(totsheet, "A13", "A - North clockwise")
		xl.SetCellValue(totsheet, "A14", "B - North anti-clockwise")
		xl.SetCellValue(totsheet, "A15", "C - South clockwise")
		xl.SetCellValue(totsheet, "A16", "D - South anti-clockwise")
		xl.SetCellValue(totsheet, "A17", "E - 500 clockwise")
		xl.SetCellValue(totsheet, "A18", "F - 500 anti-clockwise")
	}

	xl.SetCellInt(totsheet, "B3", tot.NumRiders)
	xl.SetCellInt(totsheet, "B4", tot.NumPillions)
	xl.SetCellInt(totsheet, "B5", tot.NumNovices)
	xl.SetCellInt(totsheet, "B6", tot.NumIBAMembers)

	if cfg.Rally == "rblr" {
		xl.SetCellInt(totsheet, "B7", tot.NumRBLBranch+tot.NumRBLRiders)
		xl.SetCellInt(totsheet, "B8", tot.NumRBLRiders)
		xl.SetCellInt(totsheet, "B9", tot.LoMiles2Squires)
		xl.SetCellInt(totsheet, "B10", tot.HiMiles2Squires)
		xl.SetCellInt(totsheet, "B11", tot.NumCamping)
		if *safemode {

			xl.SetCellInt(totsheet, "B12", tot.TotMoneySponsor)
			r := 13
			for i := 0; i < len(rblr_routes_ridden); i++ {
				if rblr_routes_ridden[i] > 0 {
					xl.SetCellInt(totsheet, "B"+strconv.Itoa(r), rblr_routes_ridden[i])
				}
				r++
			}
		} else {
			xl.SetCellFormula(totsheet, "B12", paysheet+"!I"+strconv.Itoa(totx.srow+1))
			r := 13
			c := "MNOPQR"
			for i := 0; i < len(rblr_routes_ridden); i++ {
				xl.SetCellFormula(totsheet, "B"+strconv.Itoa(r), overviewsheet+"!"+string(c[i])+strconv.Itoa(totx.srow+1))
				r++
			}
		}

	}

	xl.SetCellStyle(overviewsheet, "A2", "A"+totx.srowx, styleV2)
	xl.SetCellStyle(overviewsheet, "B2", "J"+totx.srowx, styleV2L)
	xl.SetCellStyle(overviewsheet, "E2", "E"+totx.srowx, styleV2)
	xl.SetCellStyle(overviewsheet, "H2", "H"+totx.srowx, styleV2)

	if !*summaryOnly {
		xl.SetCellStyle(chksheet, "A2", "A"+totx.srowx, styleV2LBig)
		xl.SetCellStyle(chksheet, "B2", "C"+totx.srowx, styleV2LBig)
		xl.SetCellStyle(chksheet, "D2", "E"+totx.srowx, styleRJSmall)
		//xl.SetCellStyle(chksheet, "H2", "H"+totx.srowx, styleV2)

		if includeShopTab {
			xl.SetCellStyle(shopsheet, "A2", "A"+totx.srowx, styleV2)
			xl.SetCellStyle(shopsheet, "B2", "C"+totx.srowx, styleV2L)
			xl.SetCellStyle(shopsheet, "D2", shop_patch_column+totx.srowx, styleV2)
		}

		xl.SetCellStyle(regsheet, "A2", "A"+totx.srowx, styleV2)
		xl.SetCellStyle(regsheet, "B2", "C"+totx.srowx, styleV2L)
		xl.SetCellStyle(regsheet, "D2", "D"+totx.srowx, styleV)
		xl.SetCellStyle(regsheet, "E2", "E"+totx.srowx, styleV2L)
		xl.SetCellStyle(regsheet, "F2", "F"+totx.srowx, styleV)
		xl.SetCellStyle(regsheet, "G2", "H"+totx.srowx, styleV2L)
		xl.SetCellStyle(regsheet, "I2", "I"+totx.srowx, styleV)

		xl.SetCellStyle(noksheet, "A2", "A"+totx.srowx, styleV3)

	}

	if cfg.Rally == "rblr" {
		xl.SetCellStyle(overviewsheet, "L2", "R"+totx.srowx, styleV)
		if !*summaryOnly {
			xl.SetCellStyle(regsheet, "J2", "J"+totx.srowx, styleV2L)
			xl.SetCellStyle(regsheet, "L2", "L"+totx.srowx, styleV)
		}
	}
	if len(cfg.Tshirts) > 0 {
		n, _ := excelize.ColumnNameToNumber("S")
		x, _ := excelize.ColumnNumberToName(n + len(cfg.Tshirts) - 1)
		xl.SetCellStyle(overviewsheet, "S2", x+totx.srowx, styleV)
	}
	if cfg.Patchavail {
		xl.SetCellStyle(overviewsheet, overview_patch_column+"2", overview_patch_column+totx.srowx, styleV)
	}

	//xl.SetCellStyle(overviewsheet, "G2", "J"+totx.srowx, styleV2)

	if !*summaryOnly {
		xl.SetCellStyle(paysheet, "A2", "A"+totx.srowx, styleV3)
		xl.SetCellStyle(paysheet, "D2", "J"+totx.srowx, styleV)
		xl.SetCellStyle(paysheet, "K2", "K"+totx.srowx, styleV)
		if cfg.Rally == "rblr" {
			xl.SetCellStyle(subssheet, "A2", "A"+totx.srowx, styleV3)
			xl.SetCellStyle(subssheet, "D2", "I"+totx.srowx, styleV)
		}
	}

	totx.srow++ // Leave a gap before totals

	ncol, _ := excelize.ColumnNameToNumber("L")
	xcol := ""
	srowt := strconv.Itoa(totx.srow)
	if *safemode {
		xcol, _ = excelize.ColumnNumberToName(ncol)
		xl.SetCellStyle(overviewsheet, xcol+srowt, xcol+srowt, styleT)
		if cfg.Rally == "rblr" {
			xl.SetCellInt(overviewsheet, xcol+srowt, tot.NumCamping)
		}
		ncol++
		if cfg.Rally == "rblr" {
			for i := 0; i < len(rblr_routes_ridden); i++ {
				xcol, _ = excelize.ColumnNumberToName(ncol)
				xl.SetCellStyle(overviewsheet, xcol+srowt, xcol+srowt, styleT)
				if rblr_routes_ridden[i] > 0 {
					xl.SetCellInt(overviewsheet, xcol+srowt, rblr_routes_ridden[i])
				}
				ncol++
			}
		}
		for i := 0; i < num_tshirt_sizes; i++ {
			xcol, _ = excelize.ColumnNumberToName(ncol)
			xl.SetCellStyle(overviewsheet, xcol+srowt, xcol+srowt, styleT)
			if totTShirts[i] > 0 {
				xl.SetCellInt(overviewsheet, xcol+srowt, totTShirts[i])
			}
			ncol++
		}
		if cfg.Patchavail {
			xcol, _ = excelize.ColumnNumberToName(ncol)
			xl.SetCellStyle(overviewsheet, xcol+srowt, xcol+srowt, styleT)
			if tot.NumPatches > 0 {
				xl.SetCellInt(overviewsheet, xcol+srowt, tot.NumPatches)
			}
			ncol++
		}
	} else {
		for _, c := range "NOPQRSTUVWXYZ" {
			ff := "sum(" + string(c) + "2:" + string(c) + totx.srowx + ")"
			xl.SetCellFormula(overviewsheet, string(c)+strconv.Itoa(totx.srow), "if("+ff+"=0,\"\","+ff+")")
			xl.SetCellStyle(overviewsheet, string(c)+strconv.Itoa(totx.srow), string(c)+strconv.Itoa(totx.srow), styleT)
		}
	}

	// Shop totals
	if includeShopTab && !*summaryOnly {
		ncol, _ = excelize.ColumnNameToNumber("D")

		if *safemode {
			for i := 0; i < num_tshirt_sizes; i++ {
				xcol, _ = excelize.ColumnNumberToName(ncol)
				xl.SetCellStyle(shopsheet, xcol+srowt, xcol+srowt, styleT)
				if totTShirts[i] > 0 {
					xl.SetCellInt(shopsheet, xcol+srowt, totTShirts[i])
				}
				ncol++
			}
			if cfg.Patchavail {
				xcol, _ = excelize.ColumnNumberToName(ncol)
				xl.SetCellStyle(shopsheet, xcol+srowt, xcol+srowt, styleT)
				if tot.NumPatches > 0 {
					xl.SetCellInt(shopsheet, xcol+srowt, tot.NumPatches)
				}
				ncol++
			}
		} else {
			for _, c := range "DEFGHI" {
				ff := "sum(" + string(c) + "2:" + string(c) + totx.srowx + ")"
				xl.SetCellFormula(shopsheet, string(c)+strconv.Itoa(totx.srow), "if("+ff+"=0,\"\","+ff+")")
				xl.SetCellStyle(shopsheet, string(c)+strconv.Itoa(totx.srow), string(c)+strconv.Itoa(totx.srow), styleT)
			}
		}
	}

	if *safemode {
		// paysheet totals
		ncol, _ = excelize.ColumnNameToNumber("D")
		var moneytot int = 0

		// Riders
		xcol, _ = excelize.ColumnNumberToName(ncol)
		moneytot = tot.NumRiders * cfg.Riderfee
		if !*summaryOnly {
			xl.SetCellStyle(paysheet, xcol+srowt, xcol+srowt, styleT)
			xl.SetCellInt(paysheet, xcol+srowt, moneytot)
		}
		ncol++

		// Pillions
		xcol, _ = excelize.ColumnNumberToName(ncol)
		moneytot = tot.NumPillions * cfg.Pillionfee
		if !*summaryOnly {
			xl.SetCellStyle(paysheet, xcol+srowt, xcol+srowt, styleT)
			xl.SetCellInt(paysheet, xcol+srowt, moneytot)
		}
		ncol++

		// T-shirts
		xcol, _ = excelize.ColumnNumberToName(ncol)
		moneytot = tot.NumTshirts * cfg.Tshirtcost
		if num_tshirt_sizes > 0 && !*summaryOnly {
			xl.SetCellStyle(paysheet, xcol+srowt, xcol+srowt, styleT)
			xl.SetCellInt(paysheet, xcol+srowt, moneytot)
		}
		ncol++

		// Patches
		xcol, _ = excelize.ColumnNumberToName(ncol)
		moneytot = tot.NumPatches * cfg.Patchcost
		if cfg.Patchavail && !*summaryOnly {
			xl.SetCellStyle(paysheet, xcol+srowt, xcol+srowt, styleT)
			xl.SetCellInt(paysheet, xcol+srowt, moneytot)
		}
		ncol++

		ncol++ // Skip cheque @ Squires

		// Sponsorship
		xcol, _ = excelize.ColumnNumberToName(ncol)
		moneytot = tot.TotMoneySponsor
		if cfg.Sponsorship && !*summaryOnly {
			xl.SetCellStyle(paysheet, xcol+srowt, xcol+srowt, styleT)
			xl.SetCellInt(paysheet, xcol+srowt, moneytot)
		}
		ncol++

		// Total received
		xcol, _ = excelize.ColumnNumberToName(ncol)
		moneytot = tot.TotMoneyMainPaypal + tot.TotMoneyCashPaypal
		if !*summaryOnly {
			xl.SetCellStyle(paysheet, xcol+srowt, xcol+srowt, styleT)
			xl.SetCellInt(paysheet, xcol+srowt, moneytot)
		}
		ncol++

	} else {
		for _, c := range "DEFGHIJKL" {
			ff := "sum(" + string(c) + "2:" + string(c) + totx.srowx + ")"
			if !*summaryOnly {
				xl.SetCellFormula(paysheet, string(c)+strconv.Itoa(totx.srow), "if("+ff+"=0,\"\","+ff+")")
				xl.SetCellStyle(paysheet, string(c)+strconv.Itoa(totx.srow), string(c)+strconv.Itoa(totx.srow), styleT)
			}
		}
	}
	xl.SetActiveSheet(0)
	if cfg.Rally == "rblr" {
		xl.SetCellValue(overviewsheet, "A1", "BL")
	} else {
		xl.SetCellValue(overviewsheet, "A1", "No.")
	}
	if !*summaryOnly {
		xl.SetCellValue(noksheet, "A1", "No.")
		xl.SetCellValue(paysheet, "A1", "No.")
		xl.SetCellValue(chksheet, "A1", "No.")
		xl.SetCellValue(regsheet, "A1", "No.")
	}
	xl.SetColWidth(overviewsheet, "A", "A", 5)
	if !*summaryOnly {
		xl.SetColWidth(noksheet, "A", "A", 5)
		xl.SetColWidth(paysheet, "A", "A", 5)
		xl.SetColWidth(regsheet, "A", "A", 5)

		if includeShopTab {
			xl.SetCellValue(shopsheet, "A1", "No.")
			xl.SetColWidth(shopsheet, "A", "A", 5)
			xl.SetCellValue(shopsheet, "B1", "Rider(first)")
			xl.SetCellValue(shopsheet, "C1", "Rider(last)")
			xl.SetColWidth(shopsheet, "B", "I", 12)
			xl.SetColWidth(shopsheet, "C", "C", 18)
		}
	}

	xl.SetColWidth(overviewsheet, "B", "D", 1)

	if !*summaryOnly {
		xl.SetColWidth(regsheet, "B", "B", 12)
		xl.SetColWidth(regsheet, "C", "C", 18)
		xl.SetColWidth(regsheet, "D", "D", 5)
		xl.SetColWidth(regsheet, "E", "E", 16)
		xl.SetColWidth(regsheet, "F", "F", 5)
		xl.SetColWidth(regsheet, "G", "G", 30)
		xl.SetColWidth(regsheet, "H", "H", 16)
		xl.SetColWidth(regsheet, "I", "I", 5)
		xl.SetColWidth(regsheet, "J", "J", 10)
		xl.SetColWidth(regsheet, "K", "K", 3)
		xl.SetColWidth(regsheet, "L", "L", 8)

		xl.SetCellValue(regsheet, "B1", "Rider(first)")
		xl.SetCellValue(regsheet, "C1", "Rider(last)")
		xl.SetCellValue(regsheet, "D1", "✓")
		xl.SetCellValue(paysheet, "B1", "Rider(first)")
		xl.SetCellValue(paysheet, "C1", "Rider(last)")
		if cfg.Rally == "rblr" {
			xl.SetCellValue(subssheet, "B1", "Rider(first)")
			xl.SetCellValue(subssheet, "C1", "Rider(last)")
		}
		xl.SetCellValue(chksheet, "B1", "Rider(first)")
		xl.SetCellValue(chksheet, "C1", "Rider(last)")
		//xl.SetCellValue(chksheet, "D1", "Bike")
		xl.SetCellValue(regsheet, "E1", "Pillion")
		xl.SetCellValue(regsheet, "F1", "✓")
		xl.SetCellValue(chksheet, "D1", "Odo")
		xl.SetCellValue(chksheet, "E1", "Time")

	}

	if cfg.Rally == "rblr" && !*summaryOnly {
		//xl.SetCellValue(chksheet, "E1", "Route")
		xl.SetCellValue(regsheet, "J1", "Route")
		xl.SetCellValue(regsheet, "K1", "✓")
		xl.SetCellValue(regsheet, "L1", "Inits")
	}

	//xl.SetCellValue(chksheet, "H1", "Notes")

	if !*summaryOnly {

		xl.SetCellValue(paysheet, "D1", "Entry")
		xl.SetCellValue(paysheet, "E1", "Pillion")
		xl.SetCellValue(regsheet, "G1", "Bike")
		xl.SetCellValue(regsheet, "H1", "Reg")
		xl.SetCellValue(regsheet, "I1", "✓")
		if len(cfg.Tshirts) > 0 {
			xl.SetCellValue(paysheet, "F1", "T-shirts")
		}
		if cfg.Patchavail {
			xl.SetCellValue(paysheet, "G1", "Patches")
		}
		if cfg.Sponsorship {
			xl.SetCellValue(paysheet, "H1", cfg.Fundsonday)
			xl.SetCellValue(paysheet, "I1", "Total Sponsorship")
			if cfg.Rally == "rblr" {
				xl.SetCellValue(subssheet,"A1","No.")
				xl.SetCellValue(subssheet, "D1", "Via Wufoo")
				xl.SetCellValue(subssheet, "E1", "Squires cheque")
				xl.SetCellValue(subssheet, "F1", "Squires cash")
				xl.SetCellValue(subssheet, "G1", "Bank transfer")
				xl.SetCellValue(subssheet, "H1", "JustGiving amount")
				xl.SetCellValue(subssheet, "I1", "JustGiving link")
			}
		}
		//xl.SetCellValue(paysheet, "K1", "+Cash")
		xl.SetCellValue(paysheet, "J1", "Total received")
		xl.SetCellValue(paysheet, "K1", "JustGiving")
		xl.SetColWidth(paysheet, "B", "B", 12)
		xl.SetColWidth(paysheet, "C", "C", 12)
		xl.SetColWidth(paysheet, "D", "G", 8)
		xl.SetColWidth(paysheet, "H", "J", 12)
		xl.SetColWidth(paysheet, "J", "J", 15)
		xl.SetColWidth(paysheet, "K", "K", 30)
		if cfg.Rally == "rblr" {
			xl.SetColWidth(subssheet, "B", "B", 12)
			xl.SetColWidth(subssheet, "C", "C", 18)
			xl.SetColWidth(subssheet, "D", "H", 10)
			xl.SetColWidth(subssheet, "I", "I", 40)

		}
	}

	xl.SetCellValue(overviewsheet, "B1", "Rider(first)")
	xl.SetCellValue(overviewsheet, "C1", "Rider(last)")
	xl.SetColWidth(overviewsheet, "A", "A", 4)
	xl.SetColWidth(overviewsheet, "B", "B", 12)
	xl.SetColWidth(overviewsheet, "C", "C", 18)

	if !*summaryOnly {
		xl.SetColWidth(chksheet, "B", "B", 15)
		xl.SetColWidth(chksheet, "C", "C", 18)
		xl.SetColWidth(chksheet, "D", "E", 20)
		//xl.SetColWidth(chksheet, "F", "G", 10)
		//xl.SetColWidth(chksheet, "H", "H", 40)
	}

	xl.SetColWidth(overviewsheet, "D", "D", 6) // Rider IBA

	if !*summaryOnly {
		xl.SetCellValue(noksheet, "B1", "Rider(first)")
		xl.SetCellValue(noksheet, "C1", "Rider(last)")
		xl.SetColWidth(noksheet, "B", "B", 10)
		xl.SetColWidth(noksheet, "C", "C", 14)
		xl.SetColWidth(noksheet, "D", "D", 15)

		xl.SetCellValue(noksheet, "D1", "Mobile")
		xl.SetCellValue(noksheet, "E1", "Contact name")
		xl.SetCellValue(noksheet, "F1", "Relationship")
		xl.SetCellValue(noksheet, "G1", "Contact number")
		xl.SetCellValue(noksheet, "H1", "Rider email")

		xl.SetColWidth(noksheet, "E", "E", 20)
		xl.SetColWidth(noksheet, "F", "F", 12)
		xl.SetColWidth(noksheet, "G", "G", 24)
		xl.SetColWidth(noksheet, "H", "H", 33)

	}

	xl.SetCellValue(overviewsheet, "D1", "IBA #")
	xl.SetCellValue(overviewsheet, "E1", stringsTitle(cfg.Novice))
	xl.SetCellValue(overviewsheet, "F1", "Pillion")
	xl.SetColWidth(overviewsheet, "F", "F", 16)
	xl.SetColWidth(overviewsheet, "G", "G", 6)
	xl.SetCellValue(overviewsheet, "G1", "IBA #")
	xl.SetCellValue(overviewsheet, "H1", stringsTitle(cfg.Novice))

	//xl.SetColVisible(overviewsheet, "B:D", false)

	xl.SetCellValue(overviewsheet, "I1", "Make")
	xl.SetColWidth(overviewsheet, "I", "I", 15)
	xl.SetCellValue(overviewsheet, "J1", "Model")
	xl.SetColWidth(overviewsheet, "J", "J", 20)

	if cfg.Rally == "rblr" {
		xl.SetCellValue(overviewsheet, "K1", " To Squires")
		xl.SetColWidth(overviewsheet, "K", "K", 4)

		xl.SetCellValue(overviewsheet, "L1", " Camping")
		xl.SetColWidth(overviewsheet, "L", "L", 2)
		xl.SetColWidth(overviewsheet, "M", "R", 3)

		xl.SetCellValue(overviewsheet, "M1", rblr_routes[0])
		xl.SetCellValue(overviewsheet, "N1", rblr_routes[1])
		xl.SetCellValue(overviewsheet, "O1", rblr_routes[2])
		xl.SetCellValue(overviewsheet, "P1", rblr_routes[3])
		xl.SetCellValue(overviewsheet, "Q1", rblr_routes[4])
		xl.SetCellValue(overviewsheet, "R1", rblr_routes[5])

	}

	if len(cfg.Tshirts) > 0 {
		n, _ := excelize.ColumnNameToNumber("S")
		for i := 0; i < len(cfg.Tshirts); i++ {
			x, _ := excelize.ColumnNumberToName(n + i)
			xl.SetColWidth(overviewsheet, x, x, 3)
			xl.SetCellValue(overviewsheet, x+"1", tshirt_sizes[i])
		}
	}
	if cfg.Patchavail {
		xl.SetColWidth(overviewsheet, overview_patch_column, overview_patch_column, 3)
		xl.SetCellValue(overviewsheet, overview_patch_column+"1", " Patches")
	}
	if includeShopTab && !*summaryOnly {
		if len(cfg.Tshirts) > 0 {
			n, _ := excelize.ColumnNameToNumber("D")
			for i := 0; i < len(cfg.Tshirts); i++ {
				x, _ := excelize.ColumnNumberToName(n + i)
				xl.SetCellValue(shopsheet, x+"1", tshirt_sizes[i])
			}
		}
		if cfg.Patchavail {
			xl.SetCellValue(shopsheet, shop_patch_column+"1", " Patches")
		}
	}

	xl.SetRowHeight(overviewsheet, 1, 70)
	if !*summaryOnly {
		xl.SetRowHeight(noksheet, 1, 20)
		xl.SetRowHeight(paysheet, 1, 70)
		if cfg.Rally == "rblr" {
			xl.SetRowHeight(subssheet, 1, 70)
		}
	}
	sort.Slice(tot.Bikes, func(i, j int) bool {
		if tot.Bikes[i].Num == tot.Bikes[j].Num {
			return strings.Compare(tot.Bikes[i].Make, tot.Bikes[j].Make) < 0
		}
		return tot.Bikes[i].Num > tot.Bikes[j].Num
	})
	//fmt.Printf("%v\n", bikes)
	totx.srow = 2
	ntot := 0
	for i := 0; i < len(tot.Bikes); i++ {

		xl.SetCellValue(totsheet, "E"+strconv.Itoa(totx.srow+1), tot.Bikes[i].Make)
		xl.SetCellInt(totsheet, "F"+strconv.Itoa(totx.srow+1), tot.Bikes[i].Num)
		xl.SetCellStyle(totsheet, "F"+strconv.Itoa(totx.srow+1), "F"+strconv.Itoa(totx.srow), styleRJ)

		ntot += tot.Bikes[i].Num
		totx.srow++
	}

	totx.srow++

}
