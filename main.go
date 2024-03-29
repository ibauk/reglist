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
	"golang.org/x/text/cases"
	"golang.org/x/text/language"
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

const apptitle = "IBAUK Reglist v1.25\nCopyright (c) 2024 Bob Stammers\n\n"
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

// properBike attempts to properly capitalise the various parts of a
// bike description. Mostly but not always that means uppercasing it.
func properBike(x string) string {

	var specials = words.Bikewords
	for _, e := range specials {
		re := regexp.MustCompile(`(?i)(.*)\b(` + e + `)\b(.*)`) // a word on its own
		if re.MatchString(x) {
			res := re.FindStringSubmatch(x)
			x = res[1] + e + res[3]
		} else {
			re := regexp.MustCompile(`(?i)(.*)(0` + e + `)\b(.*)`) // or right after an engine size
			if re.MatchString(x) {
				res := re.FindStringSubmatch(x)
				x = res[1] + "0" + e + res[3]
			}
		}
	}
	return x
}

// properMake2 fixes two word Makes such as 'Royal Enfield' and 'Moto Guzzi'
// by replacing the intervening space with an underscore, replaced later in
// processing
func properMake2(x string) string {

	var specials = words.Bikewords
	var xwords = strings.Fields(x)

	if len(xwords) < 2 {
		return x
	}
	var make2 = strings.Join(xwords[0:2], " ")
	var model2 = strings.Join(xwords[2:], " ")

	for _, e := range specials {
		if strings.EqualFold(strings.Replace(e, "-", " ", 1), make2) {
			var res strings.Builder
			res.WriteString(strings.Replace(e, " ", "_", 1))
			res.WriteString(" ")
			res.WriteString(model2)
			return res.String()
		}
	}
	return x
}

func stringsTitle(x string) string {

	caser := cases.Title(language.English)
	return caser.String(x)

}
func properName(x string) string {

	var specials = words.Specialnames
	var xx = strings.TrimSpace(x)
	if strings.ToUpper(xx) == xx || strings.ToLower(xx) == xx {
		// Now need to special names like McCrea, McCreanor, etc
		// This might be one word or more than one so
		w := strings.Split(xx, " ")
		for i := 0; i < len(w); i++ {
			var wx = w[i]
			if words.Propernames {
				wx = strings.ToLower(w[i])
				w[i] = stringsTitle(wx)
			}
			for _, wy := range specials {
				if strings.EqualFold(wx, wy) {
					w[i] = wy
				}
			}
		}
		return strings.Join(w, " ")
	}
	return xx

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
	fmt.Printf("%v entrants loaded\n", n)
	rows.Close()
}

// formatSheet sets printed properties include page orientation and margins
func formatSheet(sheetName string, portrait bool) {

	var om excelize.PageLayoutOrientation

	if portrait {
		om = excelize.OrientationPortrait
	} else {
		om = excelize.OrientationLandscape
	}

	xl.SetPageLayout(
		sheetName,
		excelize.PageLayoutOrientation(om),
		excelize.PageLayoutPaperSize(9), /* xlPaperSizeA4 (10 = xlPaperSizeA4Small!) */
		excelize.FitToHeight(2),
		excelize.FitToWidth(2),
	)
	xl.SetPageMargins(sheetName,
		excelize.PageMarginBottom(0.2),
		excelize.PageMarginFooter(0.2),
		excelize.PageMarginHeader(0.2),
		excelize.PageMarginLeft(0.2),
		excelize.PageMarginRight(0.2),
		excelize.PageMarginTop(0.2),
	)

}

func init() {

	flag.Usage = func() {
		w := flag.CommandLine.Output()
		fmt.Fprintf(w, "%v\n", apptitle)
		fmt.Fprintf(w, "%v\n", progdesc)
		flag.PrintDefaults()
	}
	flag.Parse()
	if *showusage {
		flag.Usage()
		os.Exit(1)
	}

	fmt.Print(apptitle)

	var cfgerr error

	words, cfgerr = NewWords()
	if cfgerr != nil {
		log.Fatal(cfgerr)
	}
	//fmt.Printf("%v\n\n", words)

	if *rally == "" {
		log.Fatal("You must specify the configuration file to use: -cfg rblr")
	}
	cfg, cfgerr = NewConfig(*rally + ".yml")
	if cfgerr != nil {
		log.Fatal(cfgerr)
	}
	if *csvAdmin {
		*csvReport = false
	}
	if *csvReport {
		dbfieldsx = fieldlistFromConfig(cfg.Rfields)
		fmt.Printf("CSV downloaded from Wufoo report\n")
	} else {
		dbfieldsx = fieldlistFromConfig(cfg.Afields)
		fmt.Printf("CSV downloaded from Wufoo Administrator page\n")
	}

	if *allTabs {
		*summaryOnly = false
	}

	var sm string = "live"
	if *livemode {
		*safemode = false
	}
	if *safemode {
		sm = "safe spreadsheet format"
	}

	if cfg.Rally == "rblr" {
		fmt.Printf("Running in RBLR mode, %v\n", sm)
		sqlx = "SELECT " + sqlx_rblr + " FROM entrants ORDER BY " + cfg.EntrantOrder
	} else {
		fmt.Printf("Running in rally mode, %v\n", sm)
		sqlx = "SELECT " + sqlx_rally + " FROM entrants ORDER BY " + cfg.EntrantOrder
	}

	exportingCSV = *expReport != ""
	exportingGmail = *expGmail != ""

	// This needs to be at least as big as the number of sizes declared
	num_tshirt_sizes = len(cfg.Tshirts)
	if num_tshirt_sizes > max_tshirt_sizes {
		num_tshirt_sizes = max_tshirt_sizes
	}

	// 6 is the number of RBLR routes - should be more generalised class taken from config, slapped wrist
	tot = NewTotals(6, max_tshirt_sizes, 0)

	var err error
	db, err = sql.Open("sqlite3", *sqlName)
	if err != nil {
		log.Fatal(err)
	}

	lookupOnline = lookupOnlineAvail()

	*noLookup = *noLookup || (*ridesdb == "" && !lookupOnline)

	if !*noLookup && *ridesdb != "" {
		if _, err = os.Stat(*ridesdb); os.IsNotExist(err) {
			*noLookup = true
		} else {
			_, err = db.Exec("ATTACH '" + *ridesdb + "' As rd")
			if err != nil {
				log.Fatal(err)
			}
			lookupOnline = false
		}
	}

	if *noLookup {
		fmt.Printf("Automatic IBA member identification not running\n")
	} else if *ridesdb == "" {
		fmt.Print("IBA member details being checked online")
		if *verbose {
			fmt.Printf(" %v", words.LiveDBURL)
		}
		fmt.Println()
	} else {
		fmt.Printf("Unidentified IBA members looked up using %v\n", *ridesdb)
	}

	includeShopTab = len(cfg.Tshirts) > 0 || cfg.Patchavail
	if includeShopTab {
		fmt.Printf("Including shop tab\n")
		for i := 0; i < len(cfg.Tshirts); i++ { // Let's just have an uncontrolled panic if someone specifies too many sizes
			tshirt_sizes[i] = " T-shirt " + cfg.Tshirts[i] // The leading space just makes sense
		}
	}

	// Fix columns for patches
	numsizes := len(cfg.Tshirts)
	n, _ := excelize.ColumnNameToNumber("S")
	overview_patch_column, _ = excelize.ColumnNumberToName(n + numsizes)
	n, _ = excelize.ColumnNameToNumber("D")
	shop_patch_column, _ = excelize.ColumnNumberToName(n + numsizes)
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

		xl.SetCellStyle(chksheet, "A1", "C1", styleH2L)
		xl.SetCellStyle(chksheet, "D1", "E1", styleH2)

		if includeShopTab {
			xl.SetCellStyle(shopsheet, "A1", shop_patch_column+"1", styleH2)
		}
	}

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

func setPagePane(sheet string) {
	xl.SetPanes(sheet, `{
		"freeze": true,
		"split": false,
		"x_split": 0,
		"y_split": 1,
		"top_left_cell": "A2",
		"active_pane": "bottomLeft",
		"panes": [
		{
			"sqref": "A2:X2",
			"active_cell": "A2",
			"pane": "bottomLeft"
		}]
	}`)
}

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

func mainloop() {
	rows1, err1 := db.Query(sqlx)
	if err1 != nil {
		log.Fatal(err1)
	}
	totx.srow = 2 // First spreadsheet row to populate

	var tshirts [max_tshirt_sizes]int

	for rows1.Next() {
		var RiderFirst string
		var RiderLast string
		var RiderIBA string
		var RiderRBL, PillionRBL string
		var PillionFirst, PillionLast, PillionIBA string
		var Bike, Make, Model string
		var Miles string
		var Camp, Route, T1, T2, Patches string
		var Mobile, NokName, NokNumber, NokRelation string
		var PayTot string
		var Sponsor, Paid, Cash string
		var novicerider, novicepillion string
		var miles2squires, freecamping string
		var entrantid int
		var feesdue int = 0
		var odocounts string
		var isFOC bool = false
		var withdrawn string
		var isWithdrawn bool = false
		var isCancelled bool = false
		var hasPillionVal string
		var hasPillion bool = false
		var NokRiderClash bool = false
		var NokPillionClash bool = false
		var NokMobileClash bool = false

		// Entrant record for export
		var e Entrant

		for i := 0; i < num_tshirt_sizes; i++ {
			tshirts[i] = 0
		}

		var err2 error
		if cfg.Rally == "rblr" {
			err2 = rows1.Scan(&RiderFirst, &RiderLast, &RiderIBA, &RiderRBL, &PillionFirst, &PillionLast, &PillionIBA, &PillionRBL,
				&Bike, &Miles, &Camp, &Route, &T1, &T2, &Patches, &Cash,
				&Mobile, &NokName, &NokNumber, &NokRelation, &entrantid, &PayTot, &Sponsor, &Paid, &novicerider, &novicepillion,
				&odocounts, &e.BikeReg, &miles2squires, &freecamping,
				&e.Address1, &e.Address2, &e.Town, &e.County, &e.Postcode, &e.Country,
				&e.Email, &e.Phone, &e.EnteredDate, &withdrawn, &hasPillionVal)
		} else {
			err2 = rows1.Scan(&RiderFirst, &RiderLast, &RiderIBA, &PillionFirst, &PillionLast, &PillionIBA,
				&Bike, &T1, &T2,
				&Mobile, &NokName, &NokNumber, &NokRelation, &entrantid, &PayTot, &Paid, &novicerider, &novicepillion, &odocounts,
				&e.BikeReg, &e.Address1, &e.Address2, &e.Town, &e.County, &e.Postcode, &e.Country,
				&e.Email, &e.Phone, &e.EnteredDate, &withdrawn, &hasPillionVal)
		}
		if err2 != nil {
			log.Fatalf("mainloop/err2 %v\n", err2)
		}

		isFOC = Paid == "Refunded"
		isCancelled = Paid == "Cancelled"
		isWithdrawn = withdrawn == "Withdrawn"
		hasPillion = strings.ToLower(hasPillionVal) != "no pillion" && hasPillionVal != ""

		//fmt.Printf("[ %v ] = %v \n", hasPillionVal, hasPillion)

		Bike = properMake2(Bike)
		Bike = properBike(Bike)
		if words.DefaultRE != "" {
			re := regexp.MustCompile(words.DefaultRE)
			if re.MatchString(Bike) {
				Make = words.DefaultBike
				Model = ""
			} else {
				Make, Model = extractMakeModel(Bike)
			}
		} else {
			Make, Model = extractMakeModel(Bike)
		}
		if Make != words.DefaultBike && Model == "" {
			Model = words.DefaultBike
		}

		e.Entrantid = strconv.Itoa(entrantid) // All adjustments already applied
		e.RiderFirst = properName(RiderFirst)
		e.RiderLast = properName(RiderLast)
		if isWithdrawn {
			fmt.Printf("    Rider %v %v [#%v] is withdrawn\n", e.RiderFirst, e.RiderLast, e.Entrantid)
			e.RiderLast += " (PROV)"
			continue
		} else if *verbose && Paid != "Completed" {
			fmt.Printf("    Rider %v %v [#%v] has payment status = %v\n", e.RiderFirst, e.RiderLast, e.Entrantid, Paid)
		}
		e.RiderIBA = fmtIBA(RiderIBA)
		e.RiderRBL = fmtRBL(RiderRBL)
		e.RiderNovice = bNoviceYN(novicerider, e.RiderIBA) //fmtNoviceYN(novicerider)
		e.PillionFirst = properName(PillionFirst)
		if hasPillion && PillionLast == "" {
			e.PillionLast = properName(RiderLast)
		} else {
			e.PillionLast = properName(PillionLast)
		}
		e.PillionIBA = fmtIBA(PillionIBA)
		e.PillionRBL = fmtRBL(PillionRBL)
		e.PillionNovice = bNoviceYN(novicepillion, e.PillionIBA) // fmtNoviceYN(novicepillion)
		e.BikeMake = Make
		e.BikeModel = Model
		e.OdoKms = fmtOdoKM(odocounts)

		e.BikeReg = strings.ToUpper(e.BikeReg)
		// e.Email = ""
		// e.Phone = ""
		// e.Address1 = ""
		// e.Address2 = ""
		// e.Town = ""
		// e.County = ""
		e.Postcode = strings.ToUpper(e.Postcode)
		// e.Country = ""

		e.NokName = properName(NokName)
		e.NokPhone = NokNumber
		e.NokRelation = properName(NokRelation)

		// e.BonusClaimMethod = ""
		e.RouteClass = Route
		e.Tshirt1 = T1
		e.Tshirt2 = T2
		e.Patches = Patches
		e.Camping = fmtCampingYN(freecamping)
		e.Miles2Squires = strconv.Itoa(intval(miles2squires))
		e.Bike = Bike

		if !*noLookup {
			LookupIBANumbers(&e)
		}

		RiderFirst = properName(e.RiderFirst)
		RiderLast = properName(e.RiderLast)
		PillionFirst = properName(e.PillionFirst)
		PillionLast = properName(e.PillionLast)

		//fmt.Printf("%v (%v) %v (%v)\n", RiderFirst, T1, RiderLast, T2)
		if isFOC {
			fmt.Printf("Rider %v %v [#%v] has Paid=%v and is therefore FOC\n", e.RiderFirst, e.RiderLast, e.Entrantid, Paid)
		}
		if isCancelled {
			fmt.Printf("Rider %v %v [#%v] has Paid=%v\n", e.RiderFirst, e.RiderLast, e.Entrantid, Paid)
		}

		if e.RiderFirst+" "+e.RiderLast == e.NokName {
			fmt.Printf("*** Rider %v [#%v] is the emergency contact (%v)\n", e.NokName, e.Entrantid, e.NokRelation)
			NokRiderClash = true
		} else if e.PillionFirst+" "+e.PillionLast == e.NokName {
			fmt.Printf("*** Pillion %v %v [#%v] is the emergency contact (%v)\n", e.PillionFirst, e.PillionLast, e.Entrantid, e.NokRelation)
			NokPillionClash = true
		}

		if strings.ReplaceAll(e.Phone, " ", "") == strings.ReplaceAll(e.NokPhone, " ", "") {
			fmt.Printf("*** Rider %v %v [#%v] has the same mobile as emergency contact %v\n", e.RiderFirst, e.RiderLast, e.Entrantid, e.Phone)
			NokMobileClash = true
		}

		npatches := intval(Patches)
		totx.srowx = strconv.Itoa(totx.srow)

		ebym := Entrystats{ReportingPeriod(e.EnteredDate), 1, 0, 0, 0, 0}

		if !isCancelled || !cancelsLoseOut {
			for i := 0; i < num_tshirt_sizes; i++ {
				if cfg.Tshirts[i] == T1 {
					tshirts[i]++
					totTShirts[i]++
					tot.NumTshirtsBySize[i]++
					tot.NumTshirts++
				}
				if cfg.Tshirts[i] == T2 {
					tshirts[i]++
					totTShirts[i]++
					tot.NumTshirtsBySize[i]++
					tot.NumTshirts++
				}
			}
		}
		if !isCancelled {
			// Count the bikes by Make
			var ok bool = true
			for i := 0; i < len(tot.Bikes); i++ {
				if tot.Bikes[i].Make == Make {
					tot.Bikes[i].Num++
					ok = false
				}
			}
			if ok { // Add a new make tothe list
				bmt := Bikemake{Make, 1}
				tot.Bikes = append(tot.Bikes, bmt)
			}

			tot.NumRiders++

			if strings.Contains(novicerider, cfg.Novice) {
				tot.NumNovices++
				ebym.NumNovice++
			}
			if strings.Contains(novicepillion, cfg.Novice) {
				tot.NumNovices++
			}
			if e.RiderIBA != "" {
				tot.NumIBAMembers++
				ebym.NumIBA++
			}
			if e.PillionIBA != "" {
				tot.NumIBAMembers++
			}

			if e.RiderRBL == "R" {
				tot.NumRBLRiders++
				ebym.NumRBLRiders++
			}
			if e.RiderRBL == "L" {
				tot.NumRBLBranch++
				ebym.NumRBLBranch++
			}
			if e.PillionRBL == "R" {
				tot.NumRBLRiders++
				ebym.NumRBLRiders++
			}
			if e.PillionRBL == "L" {
				tot.NumRBLBranch++
				ebym.NumRBLBranch++
			}

			ok = false
			for i := 0; i < len(tot.EntriesByPeriod); i++ {
				if tot.EntriesByPeriod[i].Month == ebym.Month {
					ok = true
					tot.EntriesByPeriod[i].Total += ebym.Total
					tot.EntriesByPeriod[i].NumIBA += ebym.NumIBA
					tot.EntriesByPeriod[i].NumNovice += ebym.NumNovice
					tot.EntriesByPeriod[i].NumRBLBranch += ebym.NumRBLBranch
					tot.EntriesByPeriod[i].NumRBLRiders += ebym.NumRBLRiders
				}
			}
			if !ok {
				tot.EntriesByPeriod = append(tot.EntriesByPeriod, ebym)
			}

		} // !isCancelled

		if !isCancelled || !cancelsLoseOut {

			if cfg.Rally == "rblr" {
				if intval(miles2squires) < tot.LoMiles2Squires {
					tot.LoMiles2Squires = intval(miles2squires)
				}
				if intval(miles2squires) > tot.HiMiles2Squires {
					tot.HiMiles2Squires = intval(miles2squires)
				}
				if fmtCampingYN(freecamping) == "Y" {
					tot.NumCamping++
				}
			}

			tot.NumPatches += npatches

		}
		if isCancelled {
			tot.CancelledRows = append(tot.CancelledRows, totx.srow)
		}

		if !*summaryOnly {
			if isCancelled {
				xl.SetRowVisible(chksheet, totx.srow, false)
				xl.SetRowVisible(regsheet, totx.srow, false)
				xl.SetRowVisible(noksheet, totx.srow, false)
			} else {
				xl.SetRowHeight(chksheet, totx.srow, 25)
			}
		}

		// Entrant IDs
		if cfg.Rally == "rblr" {
			xl.SetCellValue(overviewsheet, "A"+totx.srowx, e.RiderRBL)
		} else {
			xl.SetCellInt(overviewsheet, "A"+totx.srowx, entrantid)
		}
		if !*summaryOnly {
			xl.SetCellInt(regsheet, "A"+totx.srowx, entrantid)
			xl.SetCellInt(noksheet, "A"+totx.srowx, entrantid)
			xl.SetCellInt(paysheet, "A"+totx.srowx, entrantid)
			if includeShopTab {
				xl.SetCellInt(shopsheet, "A"+totx.srowx, entrantid)
			}
			xl.SetCellInt(chksheet, "A"+totx.srowx, entrantid)

		}
		// Rider names
		xl.SetCellValue(overviewsheet, "B"+totx.srowx, RiderFirst)
		xl.SetCellValue(overviewsheet, "C"+totx.srowx, RiderLast)

		if !*summaryOnly {
			xl.SetCellValue(regsheet, "B"+totx.srowx, RiderFirst)
			xl.SetCellValue(regsheet, "C"+totx.srowx, RiderLast)
			xl.SetCellValue(noksheet, "B"+totx.srowx, RiderFirst)
			xl.SetCellValue(noksheet, "C"+totx.srowx, RiderLast)
			xl.SetCellValue(paysheet, "B"+totx.srowx, RiderFirst)
			xl.SetCellValue(paysheet, "C"+totx.srowx, RiderLast)
			if includeShopTab {
				xl.SetCellValue(shopsheet, "B"+totx.srowx, RiderFirst)
				xl.SetCellValue(shopsheet, "C"+totx.srowx, RiderLast)
			}
			xl.SetCellValue(chksheet, "B"+totx.srowx, RiderFirst)
			xl.SetCellValue(chksheet, "C"+totx.srowx, RiderLast)
			//if !isCancelled {
			//	xl.SetCellValue(chksheet, "D"+totx.srowx, Bike)
			//}
			if len(odocounts) > 0 && odocounts[0] == 'K' {
				xl.SetCellValue(chksheet, "D"+totx.srowx, "kms")
			}

		}

		cancelledFees := 0

		if !isCancelled {
			if !*summaryOnly {
				// Fees on Money tab
				xl.SetCellInt(paysheet, "D"+totx.srowx, cfg.Riderfee) // Basic entry fee
			}
			feesdue += cfg.Riderfee
		} else {
			cancelledFees += cfg.Riderfee
		}

		if PillionFirst != "" && PillionLast != "" {
			if !isCancelled {
				if !*summaryOnly {
					xl.SetCellInt(paysheet, "E"+totx.srowx, cfg.Pillionfee)
				}
				tot.NumPillions++
				feesdue += cfg.Pillionfee
			} else {
				cancelledFees += cfg.Pillionfee
			}
		}
		var nt int = 0
		for i := 0; i < len(tshirts); i++ {
			nt += tshirts[i]
		}
		if nt > 0 {
			if !isCancelled || !cancelsLoseOut {
				if !*summaryOnly {
					xl.SetCellInt(paysheet, "F"+totx.srowx, cfg.Tshirtcost*nt)
				}
				feesdue += nt * cfg.Tshirtcost
			} else {
				cancelledFees += nt * cfg.Tshirtcost
			}
		}

		if cfg.Patchavail && npatches > 0 {
			if !isCancelled || !cancelsLoseOut {
				xl.SetCellInt(overviewsheet, "X"+totx.srowx, npatches) // Overview tab

				if !*summaryOnly {
					xl.SetCellInt(paysheet, "G"+totx.srowx, npatches*cfg.Patchcost)
					xl.SetCellInt(shopsheet, shop_patch_column+totx.srowx, npatches) // Shop tab
				}
				feesdue += npatches * cfg.Patchcost

			} else {
				cancelledFees += npatches * cfg.Patchcost
			}
		}

		intCash := intval(Cash)

		tot.TotMoneyCashPaypal += intCash

		if isFOC {
			PayTot = strconv.Itoa(feesdue - intCash)
		}

		Sponsorship := cancelledFees

		tot.TotMoneyMainPaypal += intval(PayTot)

		due := (intval(PayTot) + intCash) - feesdue

		if cfg.Sponsorship {
			// This extracts a number if present from either "Include ..." or "I'll bring ..."
			Sponsorship += intval(Sponsor) // "50"

			due -= Sponsorship
			if due > 0 {
				Sponsorship += due
				due = 0
			}

			tot.TotMoneySponsor += Sponsorship

			if !*summaryOnly {
				if *safemode {
					if Sponsorship != 0 {
						xl.SetCellInt(paysheet, "I"+totx.srowx, Sponsorship)
					}
					xl.SetCellInt(paysheet, "J"+totx.srowx, intCash+intval(PayTot))
				} else {
					sf := "H" + totx.srowx + "+" + strconv.Itoa(Sponsorship)
					xl.SetCellFormula(paysheet, "I"+totx.srowx, "if("+sf+"=0,\"0\","+sf+")")
					xl.SetCellFormula(paysheet, "J"+totx.srowx, "H"+totx.srowx+"+"+strconv.Itoa(intCash)+"+"+strconv.Itoa(intval(PayTot)))
				}

			} else {
				xl.SetCellInt(paysheet, "J"+totx.srowx, intval(PayTot))
			}

		}
		if !*summaryOnly {
			if Paid == "Unpaid" && false {
				xl.SetCellValue(paysheet, "K"+totx.srowx, " UNPAID")
				xl.SetCellStyle(paysheet, "K"+totx.srowx, "K"+totx.srowx, styleW)
			} else if !*safemode {
				ff := "J" + totx.srowx + "-(sum(D" + totx.srowx + ":G" + totx.srowx + ")+I" + totx.srowx + ")"
				xl.SetCellFormula(paysheet, "K"+totx.srowx, "if("+ff+"=0,\"\","+ff+")")
			} else {
				//due := (intval(PayTot) + intCash) - (feesdue + Sponsorship)
				if due != 0 {
					xl.SetCellInt(paysheet, "K"+totx.srowx, due)
				}
			}
		}

		if !*summaryOnly {
			// NOK List
			xl.SetCellValue(noksheet, "D"+totx.srowx, trimPhone(Mobile))
			xl.SetCellStyle(noksheet, "B"+totx.srowx, "H"+totx.srowx, styleV2L)

			if !isCancelled {
				xl.SetCellValue(noksheet, "E"+totx.srowx, properName(NokName))
				xl.SetCellValue(noksheet, "F"+totx.srowx, properName(NokRelation))
				xl.SetCellValue(noksheet, "G"+totx.srowx, trimPhone(NokNumber))
				if NokMobileClash {
					xl.SetCellStyle(noksheet, "G"+totx.srowx, "G"+totx.srowx, styleCancel)
				}
				if NokRiderClash || NokPillionClash {
					xl.SetCellStyle(noksheet, "E"+totx.srowx, "E"+totx.srowx, styleCancel)
				}
			}
			xl.SetCellValue(noksheet, "H"+totx.srowx, e.Email)
		}

		if !*summaryOnly {
			// Registration log
			xl.SetCellValue(regsheet, "E"+totx.srowx, properName(PillionFirst)+" "+properName(PillionLast))
			if !isCancelled {
				xl.SetCellValue(regsheet, "G"+totx.srowx, Make+" "+Model)
				xl.SetCellValue(regsheet, "H"+totx.srowx, e.BikeReg)
			}
		}
		// Overview
		xl.SetCellValue(overviewsheet, "D"+totx.srowx, fmtIBA(e.RiderIBA))

		xl.SetCellValue(overviewsheet, "F"+totx.srowx, PillionFirst+" "+PillionLast)
		if cfg.Rally != "rblr" {
			//			xl.SetCellValue(overviewsheet, "E"+totx.srowx, fmtNovice(novicerider))
			xl.SetCellValue(overviewsheet, "E"+totx.srowx, fmtNoviceYb(e.RiderNovice))
			xl.SetCellValue(overviewsheet, "G"+totx.srowx, fmtIBA(e.PillionIBA))
			//			xl.SetCellValue(overviewsheet, "H"+totx.srowx, fmtNovice(novicepillion))
			xl.SetCellValue(overviewsheet, "H"+totx.srowx, fmtNoviceYb(e.PillionNovice))
		}
		if !isCancelled {
			xl.SetCellValue(overviewsheet, "I"+totx.srowx, ShortMaker(Make))
			xl.SetCellValue(overviewsheet, "J"+totx.srowx, Model)
		}

		xl.SetCellValue(overviewsheet, "K"+totx.srowx, Miles)

		if Camp == "Yes" && cfg.Rally == "rblr" && (!isCancelled || !cancelsLoseOut) {
			xl.SetCellValue(overviewsheet, "L"+totx.srowx, "Y")
		}
		var cols string = "MNOPQR"
		var col int = 0
		if cfg.Rally == "rblr" && !isCancelled {
			col = strings.Index("ABCDEF", string(Route[0])) // Which route is being ridden. Compare the A -, B -, ...
			xl.SetCellInt(overviewsheet, string(cols[col])+totx.srowx, 1)

			if !*summaryOnly {
				//xl.SetCellValue(chksheet, "E"+totx.srowx, rblr_routes[col]) // Carpark
				xl.SetCellValue(regsheet, "J"+totx.srowx, rblr_routes[col]) // Registration
			}

			rblr_routes_ridden[col]++
		}

		if includeShopTab && !*summaryOnly {
			//cols = "DEFGH"
			n, _ := excelize.ColumnNameToNumber("D")
			for col = 0; col < len(tshirts); col++ {
				if tshirts[col] > 0 {
					x, _ := excelize.ColumnNumberToName(n + col)
					xl.SetCellInt(shopsheet, x+totx.srowx, tshirts[col])
				}
			}
		}

		//cols = "STUVW"
		n, _ := excelize.ColumnNameToNumber("S")
		for col = 0; col < len(tshirts); col++ {
			if tshirts[col] > 0 {
				x, _ := excelize.ColumnNumberToName(n + col)
				xl.SetCellInt(overviewsheet, x+totx.srowx, tshirts[col])
			}
		}

		totx.srow++

		//fmt.Printf("%v\n", Entrant2Strings(e))

		if exportingCSV && !isWithdrawn && !isCancelled {
			csvW.Write(Entrant2Strings(e))
		}
		if exportingGmail && !isWithdrawn && !isCancelled {
			csvGmail.Write(Entrant2Gmail(e))
		}

	} // End reading loop

}

func markCancelledEntrants() {
	for _, r := range tot.CancelledRows {
		rx := strconv.Itoa(r)
		xl.SetCellStyle(overviewsheet, "A"+rx, "J"+rx, styleCancel)

		if !*summaryOnly {
			xl.SetCellStyle(regsheet, "A"+rx, "I"+rx, styleCancel)
			if cfg.Rally == "rblr" {
				xl.SetCellStyle(regsheet, "J"+rx, "K"+rx, styleCancel)
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

func trimPhone(tel string) string {

	var res string

	telx := strings.ReplaceAll(tel, " ", "")
	if telx[0:2] == "00" {
		telx = strings.Replace(telx, "00", "+", 1)
	}

	if len(telx) > words.MaxPhone && words.MaxPhone > 0 {
		res = telx[:words.MaxPhone]
	} else {
		res = telx
	}
	return res
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
		setPageTitle(regsheet)
	}
	setPagePane(overviewsheet)
	if !*summaryOnly {
		setPagePane(noksheet)
		setPagePane(paysheet)
		setPagePane(chksheet)
		setPagePane(regsheet)
	}

	if includeShopTab && !*summaryOnly {
		setPageTitle(shopsheet)
		setPagePane(shopsheet)
	}

	markCancelledEntrants()
}

func reportEntriesByPeriod() {

	//xl.SetCellValue(totsheet, "A16", " ") // Mark the bottom row just in case

	sort.Slice(tot.EntriesByPeriod, func(i, j int) bool { return tot.EntriesByPeriod[i].Month > tot.EntriesByPeriod[j].Month })

	//fmt.Printf("%v\n", tot.EntriesByPeriod)

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
	//xl.SetCellValue(totsheet, "M2", "RBL others")
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

	fmtx := `{"type":"bar","series": [`

	xrow := strconv.Itoa(row - 1)
	row = 3
	var cols string
	if cfg.Rally == "rblr" {
		cols = "IJKL"
	} else {
		cols = "JKL"
	}

	for i := 0; i < len(cols); i++ {
		if i > 0 {
			fmtx += `,
			`
		}
		ll := cols[i : i+1]
		fmtx += `{"name":"` + totsheet + `!$` + ll + `$2",	"categories":"` + totsheet + `!$H$3:$H$` + xrow + `","values":"` + totsheet + `!` + ll + `3:` + ll + xrow + `"}`
		row++
	}

	var reportingperiod string
	if cfg.ReportWeekly {
		reportingperiod = "week"
	} else {
		reportingperiod = "month"
	}

	fmtx += `],
	"format":
	{
		"x_scale": 1.0,
		"y_scale": 1.2,
		"x_offset": 15,
		"y_offset": 10,
		"print_obj": true,
		"lock_aspect_ratio": true,
		"locked": false
	},
	"legend":
	{
		"position": "right",
		"show_legend_key": false
	},
	"title":
	{
		"name": "New signups by ` + reportingperiod + `"
	},
	"plotarea":
	{
		"show_bubble_size": true,
		"show_cat_name": false,
		"show_leader_lines": false,
		"show_percent": true,
		"show_series_name": true,
		"show_val": true
	},
	"show_blanks_as": "zero"

	}`

	err := xl.AddChart(totsheet, "N2", fmtx)
	if err != nil {
		fmt.Printf("OMG: %v\n%v\n", err, fmtx)
	}

	xl.SetColVisible(totsheet, "H:M", false)
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

func renameSheet(oldname *string, newname string) {

	xl.SetSheetName(*oldname, newname)
	*oldname = newname

}

func makeFile(csvname string) *os.File {

	file, err := os.Create(csvname)
	if err != nil {
		panic(err)
	}
	return file

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

		for i := 0; i < len(record); i++ {
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

func initStyles() {

	// styleCancel for highlighting cancelled entrants
	styleCancel, _ = xl.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Vertical: "center", Horizontal: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"edeb57"}, Pattern: 1},
	})

	// Totals
	styleT, _ = xl.NewStyle(`{	
			"alignment":
			{
				"horizontal": "center",
				"ident": 1,
				"justify_last_line": true,
				"reading_order": 0,
				"relative_indent": 1,
				"shrink_to_fit": true,
				"text_rotation": 0,
				"vertical": "",
				"wrap_text": true
			},
			"font":
			{
				"bold": true,
				"italic": false,
				"family": "Arial",
				"size": 12,
				"color": "000000"
			}
		}`)

	// Header, vertical
	styleH, _ = xl.NewStyle(`{
			"alignment":
			{
				"horizontal": "center",
				"ident": 1,
				"justify_last_line": true,
				"reading_order": 0,
				"relative_indent": 1,
				"shrink_to_fit": true,
				"text_rotation": 90,
				"vertical": "",
				"wrap_text": true
			},
			"fill":{"type":"pattern","color":["#dddddd"],"pattern":1}	}`)

	// Header, horizontal
	styleH2, _ = xl.NewStyle(`{
				"alignment":
				{
					"horizontal": "center",
					"ident": 1,
					"justify_last_line": true,
					"reading_order": 0,
					"relative_indent": 1,
					"shrink_to_fit": true,
					"text_rotation": 0,
					"vertical": "center",
					"wrap_text": true
				},
				"fill":{"type":"pattern","color":["#dddddd"],"pattern":1}	}`)

	styleH2L, _ = xl.NewStyle(`{
					"alignment":
					{
						"horizontal": "left",
						"ident": 1,
						"justify_last_line": true,
						"reading_order": 0,
						"relative_indent": 1,
						"shrink_to_fit": true,
						"text_rotation": 0,
						"vertical": "center",
						"wrap_text": true
					},
					"fill":{"type":"pattern","color":["#dddddd"],"pattern":1}	}`)

	// Data
	styleV, _ = xl.NewStyle(`{
			"alignment":
			{
				"horizontal": "center",
				"ident": 1,
				"justify_last_line": true,
				"reading_order": 0,
				"relative_indent": 1,
				"shrink_to_fit": true,
				"text_rotation": 0,
				"vertical": "",
				"wrap_text": false
			},
			"border": [
				{
					"type": "left",
					"color": "000000",
					"style": 1
				},
				{
					"type": "bottom",
					"color": "000000",
					"style": 1
				},
				{
					"type": "right",
					"color": "000000",
					"style": 1
				}]		
		}`)

	// Open data
	styleV2, _ = xl.NewStyle(`{
			"alignment":
			{
				"horizontal": "center",
				"ident": 1,
				"justify_last_line": true,
				"reading_order": 0,
				"relative_indent": 1,
				"shrink_to_fit": true,
				"text_rotation": 0,
				"vertical": "",
				"wrap_text": false
			},
			"border": [
				{
					"type": "bottom",
					"color": "000000",
					"style": 1
				}]		
		}`)

	styleV2L, _ = xl.NewStyle(`{
			"alignment":
			{
				"horizontal": "left",
				"ident": 1,
				"justify_last_line": true,
				"reading_order": 0,
				"relative_indent": 1,
				"shrink_to_fit": true,
				"text_rotation": 0,
				"vertical": "",
				"wrap_text": false
			},
			"border": [
				{
					"type": "bottom",
					"color": "000000",
					"style": 1
				}]		
		}`)

	styleV2LBig, _ = xl.NewStyle(`{
			"alignment":
			{
				"horizontal": "left",
				"ident": 1,
				"justify_last_line": true,
				"reading_order": 0,
				"relative_indent": 1,
				"shrink_to_fit": true,
				"text_rotation": 0,
				"vertical": "",
				"wrap_text": false
			},
			"border": [
				{
					"type": "bottom",
					"color": "000000",
					"style": 1
				}],
			"font":
				{
					"size": 16
				}
		
		}`)

	styleV3, _ = xl.NewStyle(`{
			"alignment":
			{
				"horizontal": "center",
				"ident": 1,
				"justify_last_line": true,
				"reading_order": 0,
				"relative_indent": 1,
				"shrink_to_fit": true,
				"text_rotation": 0,
				"vertical": "",
				"wrap_text": false
			}
		}`)

	// styleW for highlighting, particularly errorneous, cells
	styleW, _ = xl.NewStyle(`{ 
			"alignment":
			{
				"horizontal": "center",
				"ident": 1,
				"justify_last_line": true,
				"reading_order": 0,
				"relative_indent": 1,
				"shrink_to_fit": true,
				"text_rotation": 0,
				"vertical": "",
				"wrap_text": true
			},
			"fill":{"type":"pattern","color":["#ffff00"],"pattern":1}	}`)

	styleRJ, _ = xl.NewStyle(`{ 
				"alignment":
				{
					"horizontal": "right",
					"ident": 1,
					"justify_last_line": true,
					"reading_order": 0,
					"relative_indent": 1,
					"shrink_to_fit": true,
					"text_rotation": 0,
					"vertical": "",
					"wrap_text": true
				}
	}`)

	styleRJSmall, _ = xl.NewStyle(`{ 
		"alignment":
		{
			"horizontal": "right",
			"ident": 1,
			"justify_last_line": true,
			"reading_order": 0,
			"relative_indent": 1,
			"shrink_to_fit": true,
			"text_rotation": 0,
			"vertical": "",
			"wrap_text": true
		},
		"border": [
			{
				"type": "bottom",
				"color": "000000",
				"style": 1
			}],

		"font":
		{
			"size": 8
		}
}`)

	xl.SetDefaultFont("Arial")

}

func extractMakeModel(bike string) (string, string) {

	if strings.TrimSpace(bike) == "" {
		return "", ""
	}
	re := regexp.MustCompile(`'*\d*\s*([A-Za-z\-\_]*)\s*(.*)`)
	sm := re.FindSubmatch([]byte(bike))
	if len(sm) < 3 {
		return strings.ReplaceAll(string(sm[1]), "_", " "), ""
	}
	return strings.ReplaceAll(string(sm[1]), "_", " "), strings.ReplaceAll(string(sm[2]), "_", " ")

}

func fmtIBA(x string) string {

	if x == "-1" {
		return "n/a"
	}
	return strings.ReplaceAll(x, ".0", "")

}

func fmtRBL(x string) string {

	if x == cfg.LegionMember && cfg.LegionMember != "" {
		return "L"
	} else if x == cfg.LegionRider && cfg.LegionRider != "" {
		return "R"
	}
	return ""
}

/*
*
func fmtNovice(x string) string {

		if strings.Contains(x, cfg.Novice) {
			return "Yes"
		}
		return ""
	}

*
*/
func fmtNoviceYb(noviceYN string) string {

	if noviceYN == "Y" {
		return "Yes"
	}
	return ""
}

// boolean (Y/N) novice or not
func bNoviceYN(x string, iba string) string {

	res := "N"
	if strings.Contains(x, "IBA") { // "Check for IBA number" for example
		if iba == "" {
			res = "Y" // No IBA number means I'm a novice
		}
	} else if strings.Contains(x, cfg.Novice) {
		res = "Y" // "I'm a novice" for example
	}
	return res
}

/**
func fmtNoviceYN(x string) string {
	if fmtNovice(x) != "" && x[0] != 'N' && x[0] != 'n' {
		return "Y"
	} else {
		return "N"
	}
}
**/

func fmtOdoKM(x string) string {

	y := strings.ToUpper(x)
	if len(y) > 0 && y[0] == 'K' {
		return "K"
	}
	return "M"

}

func fmtCampingYN(x string) string {

	if x == cfg.FreeCamping && cfg.FreeCamping != "" {
		return "Y"
	}
	return ""
}

func ShortMaker(x string) string {

	p := strings.Index(x, "-")
	if p < 0 {
		return x
	}
	return x[0:p]
}
