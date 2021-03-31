package main

/*
 * This is a quick and dirty, yes really, transformer to create a "Registration list"
 * spreadsheet ready for the RBLR1000.
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
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	_ "github.com/mattn/go-sqlite3"
)

var rally *string = flag.String("cfg", "rblr", "Which rally is this (yml file)")
var csvName *string = flag.String("csv", "entrants.csv", "Path to CSV downloaded from Wufoo")
var csvReport *bool = flag.Bool("rpt", false, "CSV downloaded from Wufoo report")
var sqlName *string = flag.String("sql", "entrantdata.db", "Path to SQLite database")
var xlsName *string = flag.String("xls", "reglist.xlsx", "Path to output XLSX")
var noCSV *bool = flag.Bool("nocsv", false, "Don't load a CSV file, just use the SQL database")
var safemode *bool = flag.Bool("safe", false, "Safe mode avoid formulas, no live updating")
var expReport *string = flag.String("exp", "", "Path to output standard format CSV")

const apptitle = "IBAUK Reglist v1.0.0-d\nCopyright (c) 2021 Bob Stammers\n\n"

var rblr_routes = [...]string{" A-NC", " B-NAC", " C-SC", " D-SAC", " E-500C", " F-500AC"}
var rblr_routes_ridden = [...]int{0, 0, 0, 0, 0, 0}

const max_tshirt_sizes int = 10

var tshirt_sizes [max_tshirt_sizes]string

var overview_patch_column, shop_patch_column string

// dbFields must be kept in sync with the downloaded CSV from Wufoo
// Fieldnames don't matter but the order and number both do

var dbfieldsx string

var overviewsheet string = "Sheet1" // Renamed on init
var noksheet string = "NOK list"
var paysheet string = "Money"
var totsheet string = "Stats"
var chksheet string = "Carpark"
var regsheet string = "Registration"
var shopsheet string = "Shop"

const sqlx_rblr = `ifnull(RiderName,''),ifnull(RiderLast,''),ifnull(RiderIBANumber,''),
ifnull(PillionName,''),ifnull(PillionLast,''),ifnull(PillionIBANumber,''),
ifnull(BikeMakeModel,''),round(ifnull(MilesTravelledToSquires,'0')),
ifnull(FreeCamping,''),ifnull(WhichRoute,'A'),
ifnull(Tshirt1,''),ifnull(Tshirt2,''),ifnull(Patches,'0'),ifnull(Cash,'0'),
ifnull(Mobilephone,''),
ifnull(NOKName,''),ifnull(NOKNumber,''),ifnull(NOKRelation,''),
FinalRiderNumber,ifnull(PaymentTotal,''),ifnull(Sponsorshipmoney,''),ifnull(PaymentStatus,''),
ifnull(NoviceRider,''),ifnull(NovicePillion,''),ifnull(odometer_counts,''),ifnull(Registration,''),
ifnull(MilestravelledToSquires,''),ifnull(FreeCamping,''),
ifnull(Address1,''),ifnull(Address2,''),ifnull(Town,''),ifnull(County,''),
ifnull(Postcode,''),ifnull(Country,''),ifnull(Email,''),ifnull(Mobilephone,''),ifnull(ao_BCM,'')`

const sqlx_rally = `ifnull(RiderName,''),ifnull(RiderLast,''),ifnull(RiderIBANumber,''),
ifnull(PillionName,''),ifnull(PillionLast,''),ifnull(PillionIBANumber,''),
ifnull(BikeMakeModel,''),
ifnull(Tshirt1,''),ifnull(Tshirt2,''),
ifnull(Mobilephone,''),
ifnull(NOKName,''),ifnull(NOKNumber,''),ifnull(NOKRelation,''),
FinalRiderNumber,ifnull(PaymentTotal,''),ifnull(PaymentStatus,''),
ifnull(NoviceRider,''),ifnull(NovicePillion,''),ifnull(odometer_counts,''),ifnull(Registration,''),
ifnull(Address1,''),ifnull(Address2,''),ifnull(Town,''),ifnull(County,''),
ifnull(Postcode,''),ifnull(Country,''),ifnull(Email,''),ifnull(Mobilephone,''),ifnull(ao_BCM,'')`

var sqlx string

var styleH, styleH2, styleT, styleV, styleV2, styleV2L, styleV3, styleW, styleRJ, styleRJSmall int

var cfg *Config
var words *Words

var db *sql.DB
var includeShopTab bool

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
				w[i] = strings.Title(wx)
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

func FixRiderNumbers(db *sql.DB) {

	var old string
	var new int
	var mannum string

	oldnew := make(map[string]int, 250) // More than enough

	rows, err := db.Query("SELECT EntryId,ifnull(RiderNumber,'') FROM entrants;") // There is scope for renumber alphabetically if desired.
	if err != nil {
		log.Fatal(err)
	}
	for rows.Next() {

		rows.Scan(&old, &mannum)
		new = intval(old) + cfg.Add2entrantid
		if intval(mannum) > 0 {
			new = intval(mannum)
		}
		oldnew[old] = new
	}

	tx, _ := db.Begin()
	for old, new := range oldnew {
		sqlx := "UPDATE entrants SET FinalRiderNumber=" + strconv.Itoa(new) + " WHERE EntryId='" + old + "'"
		//fmt.Println(sqlx)
		_, err := db.Exec(sqlx)
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
func formatSheet(f *excelize.File, sheetName string, portrait bool) {

	var om excelize.PageLayoutOrientation

	if portrait {
		om = excelize.OrientationPortrait
	} else {
		om = excelize.OrientationLandscape
	}

	f.SetPageLayout(
		sheetName,
		excelize.PageLayoutOrientation(om),
		excelize.PageLayoutPaperSize(10),
		excelize.FitToHeight(2),
		excelize.FitToWidth(2),
	)
	f.SetPageMargins(sheetName,
		excelize.PageMarginBottom(0.2),
		excelize.PageMarginFooter(0.2),
		excelize.PageMarginHeader(0.2),
		excelize.PageMarginLeft(0.2),
		excelize.PageMarginRight(0.2),
		excelize.PageMarginTop(0.2),
	)

}

func init() {

	flag.Parse()

	fmt.Print(apptitle)

	var cfgerr error

	words, cfgerr = NewWords()
	if cfgerr != nil {
		log.Fatal(cfgerr)
	}
	//fmt.Printf("%v\n\n", words)

	cfg, cfgerr = NewConfig(*rally + ".yml")
	if cfgerr != nil {
		log.Fatal(cfgerr)
	}
	if *csvReport {
		dbfieldsx = fieldlistFromConfig(cfg.Rfields)
	} else {
		dbfieldsx = fieldlistFromConfig(cfg.Afields)
	}

	var sm string = "live"
	if *safemode {
		sm = "safe"
	}

	if cfg.Rally == "rblr" {
		fmt.Printf("Running in RBLR mode - %v\n", sm)
		sqlx = "SELECT " + sqlx_rblr + " FROM entrants ORDER BY " + cfg.EntrantOrder
	} else {
		fmt.Printf("Running in rally mode - %v\n", sm)
		sqlx = "SELECT " + sqlx_rally + " FROM entrants ORDER BY " + cfg.EntrantOrder
	}
	var err error
	db, err = sql.Open("sqlite3", *sqlName)
	if err != nil {
		log.Fatal(err)
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

func initSpreadsheet() *excelize.File {

	f := excelize.NewFile()
	initStyles(f)
	// First sheet is called Sheet1
	formatSheet(f, overviewsheet, false)
	f.NewSheet(regsheet)
	formatSheet(f, regsheet, false)
	f.NewSheet(noksheet)
	formatSheet(f, noksheet, false)
	if includeShopTab {
		f.NewSheet(shopsheet)
		formatSheet(f, shopsheet, false)
	}
	f.NewSheet(paysheet)
	formatSheet(f, paysheet, false)
	f.NewSheet(totsheet)
	formatSheet(f, totsheet, false)
	f.NewSheet(chksheet)
	formatSheet(f, chksheet, false)

	renameSheet(f, &overviewsheet, "Overview")

	// Set heading styles
	f.SetCellStyle(overviewsheet, "A1", overview_patch_column+"1", styleH2)
	if cfg.Rally == "rblr" {
		f.SetCellStyle(overviewsheet, "K1", "R1", styleH)
		f.SetCellStyle(overviewsheet, "E1", "E1", styleH)
		f.SetCellStyle(overviewsheet, "H1", "H1", styleH)
		f.SetColVisible(overviewsheet, "E", false)
		f.SetColVisible(overviewsheet, "G:H", false)
	} else {
		f.SetColWidth(overviewsheet, "K", "R", 1)
		f.SetColVisible(overviewsheet, "K:R", false)
	}
	if len(cfg.Tshirts) > 0 {
		n, _ := excelize.ColumnNameToNumber("S")
		x, _ := excelize.ColumnNumberToName(n + len(cfg.Tshirts) - 1)
		f.SetCellStyle(overviewsheet, "S1", x+"1", styleH)
	}
	if cfg.Patchavail {
		f.SetCellStyle(overviewsheet, overview_patch_column+"1", overview_patch_column+"1", styleH)
	}

	f.SetCellStyle(regsheet, "A1", "H1", styleH2)
	if cfg.Rally == "rblr" {
		f.SetCellStyle(regsheet, "I1", "J1", styleH2)
	}
	f.SetCellStyle(noksheet, "A1", "G1", styleH2)

	f.SetCellStyle(paysheet, "A1", "K1", styleH2)

	f.SetCellStyle(chksheet, "A1", "A1", styleH2)
	f.SetCellStyle(chksheet, "A1", "H1", styleH2)

	if includeShopTab {
		f.SetCellStyle(shopsheet, "A1", shop_patch_column+"1", styleH2)
	}

	return f

}

// setPageTitle sets each sheet, except Stats, to repeat its
// top line on each printed page
func setPageTitle(f *excelize.File, sheet string) {

	var dn excelize.DefinedName

	dn.Name = "_xlnm.Print_Titles"
	dn.RefersTo = sheet + "!$1:$1"
	dn.Scope = sheet
	f.SetDefinedName(&dn)
}

func setPagePane(f *excelize.File, sheet string) {
	f.SetPanes(sheet, `{
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
		loadCSVFile(db)
	}
	FixRiderNumbers(db)

	f := initSpreadsheet()

	exportingCSV := *expReport != ""
	var csvF *os.File
	var csvW *csv.Writer

	if exportingCSV {
		csvF = makeFile(*expReport)
		defer csvF.Close()
		csvW = makeCSVFile(csvF)
		fmt.Printf("Exporting CSV to %v\n", *expReport)
	}

	rows1, err1 := db.Query(sqlx)
	if err1 != nil {
		log.Fatal(err1)
	}
	var srow int = 2 // First spreadsheet row to populate
	var srowx string
	type bikemake struct {
		make string
		num  int
	}
	var bikes []bikemake

	var numRiders int = 0
	var numPillions int = 0
	var numNovices int = 0
	var shortestSquires int = 9999
	var longestSquires int = 0
	var numCamping int = 0
	var numIBAMembers int = 0
	var totPatches int = 0
	var grandTotalTShirts int = 0
	var totSponsorship int = 0
	var totCash int = 0
	var totPayment int = 0

	// This needs to be at least as big as the number of sizes declared
	num_tshirt_sizes := len(cfg.Tshirts)
	if num_tshirt_sizes > max_tshirt_sizes {
		num_tshirt_sizes = max_tshirt_sizes
	}
	var tshirts [max_tshirt_sizes]int
	var totTShirts [max_tshirt_sizes]int = [max_tshirt_sizes]int{0}

	// 6 is the number of RBLR routes - should be more generalised class taken from config, slapped wrist
	tot := NewTotals(6, max_tshirt_sizes, 0)

	for rows1.Next() {
		var RiderFirst string
		var RiderLast string
		var RiderIBA string
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

		// Entrant record for export
		var e Entrant

		for i := 0; i < num_tshirt_sizes; i++ {
			tshirts[i] = 0
		}

		var err2 error
		if cfg.Rally == "rblr" {
			err2 = rows1.Scan(&RiderFirst, &RiderLast, &RiderIBA, &PillionFirst, &PillionLast, &PillionIBA,
				&Bike, &Miles, &Camp, &Route, &T1, &T2, &Patches, &Cash,
				&Mobile, &NokName, &NokNumber, &NokRelation, &entrantid, &PayTot, &Sponsor, &Paid, &novicerider, &novicepillion,
				&odocounts, &e.BikeReg, &miles2squires, &freecamping,
				&e.Address1, &e.Address2, &e.Town, &e.County, &e.Postcode, &e.Country,
				&e.Email, &e.Phone, &e.BonusClaimMethod)
		} else {
			err2 = rows1.Scan(&RiderFirst, &RiderLast, &RiderIBA, &PillionFirst, &PillionLast, &PillionIBA,
				&Bike, &T1, &T2,
				&Mobile, &NokName, &NokNumber, &NokRelation, &entrantid, &PayTot, &Paid, &novicerider, &novicepillion, &odocounts,
				&e.BikeReg, &e.Address1, &e.Address2, &e.Town, &e.County, &e.Postcode, &e.Country,
				&e.Email, &e.Phone, &e.BonusClaimMethod)
		}
		if err2 != nil {
			log.Fatal(err2)
		}

		Bike = properBike(Bike)
		Make, Model = extractMakeModel(Bike)

		e.Entrantid = strconv.Itoa(entrantid) // All adjustments already applied
		e.RiderFirst = properName(RiderFirst)
		e.RiderLast = properName(RiderLast)
		e.RiderIBA = fmtIBA(RiderIBA)
		e.RiderNovice = fmtNoviceYN(novicerider)
		e.PillionFirst = properName(PillionFirst)
		e.PillionLast = properName(PillionLast)
		e.PillionIBA = fmtIBA(PillionIBA)
		e.PillionNovice = fmtNoviceYN(novicepillion)
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

		RiderFirst = properName(RiderFirst)
		RiderLast = properName(RiderLast)
		PillionFirst = properName(PillionFirst)
		PillionLast = properName(PillionLast)

		//fmt.Printf("%v (%v) %v (%v)\n", RiderFirst, T1, RiderLast, T2)

		for i := 0; i < num_tshirt_sizes; i++ {
			if cfg.Tshirts[i] == T1 {
				tshirts[i]++
				totTShirts[i]++
				grandTotalTShirts++
				tot.NumTshirtsBySize[i]++
				tot.NumTshirts++
			}
			if cfg.Tshirts[i] == T2 {
				tshirts[i]++
				totTShirts[i]++
				grandTotalTShirts++
				tot.NumTshirtsBySize[i]++
				tot.NumTshirts++
			}
		}
		srowx = strconv.Itoa(srow)

		// Count the bikes by Make
		var ok bool = true
		for i := 0; i < len(bikes); i++ {
			if bikes[i].make == Make {
				bikes[i].num++
				ok = false
				tot.Bikes[i].Num++
			}
		}
		if ok { // Add a new make tothe list
			bm := bikemake{Make, 1}
			bikes = append(bikes, bm)
			bmt := Bikemake{Make, 1}
			tot.Bikes = append(tot.Bikes, bmt)
		}

		numRiders++
		tot.NumRiders++

		if strings.Contains(novicerider, cfg.Novice) {
			numNovices++
			tot.NumNovices++
		}
		if strings.Contains(novicepillion, cfg.Novice) {
			numNovices++
			tot.NumNovices++
		}
		if RiderIBA != "" {
			numIBAMembers++
			tot.NumIBAMembers++
		}
		if PillionIBA != "" {
			numIBAMembers++
			tot.NumIBAMembers++
		}

		if cfg.Rally == "rblr" {
			if intval(miles2squires) < shortestSquires {
				shortestSquires = intval(miles2squires)
				tot.LoMiles2Squires = intval(miles2squires)
			}
			if intval(miles2squires) > longestSquires {
				longestSquires = intval(miles2squires)
				tot.HiMiles2Squires = intval(miles2squires)
			}
			if freecamping == "Yes" {
				numCamping++
				tot.NumCamping++
			}
		}

		npatches := intval(Patches)
		tot.NumPatches += npatches

		// Entrant IDs
		f.SetCellInt(overviewsheet, "A"+srowx, entrantid)
		f.SetCellInt(regsheet, "A"+srowx, entrantid)
		f.SetCellInt(noksheet, "A"+srowx, entrantid)
		f.SetCellInt(paysheet, "A"+srowx, entrantid)
		if includeShopTab {
			f.SetCellInt(shopsheet, "A"+srowx, entrantid)
		}
		f.SetCellInt(chksheet, "A"+srowx, entrantid)

		// Rider names
		f.SetCellValue(overviewsheet, "B"+srowx, RiderFirst)
		f.SetCellValue(overviewsheet, "C"+srowx, RiderLast)
		f.SetCellValue(regsheet, "B"+srowx, RiderFirst)
		f.SetCellValue(regsheet, "C"+srowx, RiderLast)
		f.SetCellValue(noksheet, "B"+srowx, RiderFirst)
		f.SetCellValue(noksheet, "C"+srowx, RiderLast)
		f.SetCellValue(paysheet, "B"+srowx, RiderFirst)
		f.SetCellValue(paysheet, "C"+srowx, RiderLast)
		if includeShopTab {
			f.SetCellValue(shopsheet, "B"+srowx, RiderFirst)
			f.SetCellValue(shopsheet, "C"+srowx, RiderLast)
		}
		f.SetCellValue(chksheet, "B"+srowx, RiderFirst)
		f.SetCellValue(chksheet, "C"+srowx, RiderLast)
		f.SetCellValue(chksheet, "D"+srowx, Bike)
		if len(odocounts) > 0 && odocounts[0] == 'K' {
			f.SetCellValue(chksheet, "F"+srowx, "kms")
		}

		// Fees on Money tab
		f.SetCellInt(paysheet, "D"+srowx, cfg.Riderfee) // Basic entry fee
		feesdue += cfg.Riderfee

		if PillionFirst != "" && PillionLast != "" {
			f.SetCellInt(paysheet, "E"+srowx, cfg.Pillionfee)
			numPillions++
			tot.NumPillions++
			feesdue += cfg.Pillionfee
		}
		var nt int = 0
		for i := 0; i < len(tshirts); i++ {
			nt += tshirts[i]
		}
		if nt > 0 {
			f.SetCellInt(paysheet, "F"+srowx, cfg.Tshirtcost*nt)
			feesdue += nt * cfg.Tshirtcost
		}

		totPatches += npatches
		if cfg.Patchavail && npatches > 0 {
			f.SetCellInt(overviewsheet, "X"+srowx, npatches) // Overview tab
			f.SetCellInt(paysheet, "G"+srowx, npatches*cfg.Patchcost)
			f.SetCellInt(shopsheet, shop_patch_column+srowx, npatches) // Shop tab
			feesdue += npatches * cfg.Patchcost
		}

		intCash := intval(Cash)
		totCash += intCash

		tot.TotMoneyCashPaypal += intCash

		Sponsorship := 0
		totPayment += intval(PayTot)

		tot.TotMoneyMainPaypal += intval(PayTot)

		if cfg.Sponsorship {
			// This extracts a number if present from either "Include ..." or "I'll bring ..."
			Sponsorship = intval(Sponsor) // "50"
			totSponsorship += Sponsorship

			tot.TotMoneySponsor += Sponsorship

			if *safemode {
				if Sponsorship != 0 {
					f.SetCellInt(paysheet, "I"+srowx, Sponsorship)
				}
				f.SetCellInt(paysheet, "J"+srowx, intCash+intval(PayTot))
			} else {
				sf := "H" + srowx + "+" + strconv.Itoa(Sponsorship)
				f.SetCellFormula(paysheet, "I"+srowx, "if("+sf+"=0,\"0\","+sf+")")
				f.SetCellFormula(paysheet, "J"+srowx, "H"+srowx+"+"+strconv.Itoa(intCash)+"+"+strconv.Itoa(intval(PayTot)))
			}

		} else {
			f.SetCellInt(paysheet, "J"+srowx, intval(PayTot))
		}

		if Paid == "Unpaid" {
			f.SetCellValue(paysheet, "K"+srowx, " UNPAID")
			f.SetCellStyle(paysheet, "K"+srowx, "K"+srowx, styleW)
		} else if !*safemode {
			ff := "J" + srowx + "-(sum(D" + srowx + ":G" + srowx + ")+I" + srowx + ")"
			f.SetCellFormula(paysheet, "K"+srowx, "if("+ff+"=0,\"\","+ff+")")
		} else {
			due := (intval(PayTot) + intCash) - (feesdue + Sponsorship)
			if due != 0 {
				f.SetCellInt(paysheet, "K"+srowx, due)
			}
		}

		// NOK List
		f.SetCellValue(noksheet, "D"+srowx, Mobile)
		f.SetCellValue(noksheet, "E"+srowx, properName(NokName))
		f.SetCellValue(noksheet, "F"+srowx, properName(NokRelation))
		f.SetCellValue(noksheet, "G"+srowx, NokNumber)

		// Registration log
		f.SetCellValue(regsheet, "E"+srowx, properName(PillionFirst)+" "+properName(PillionLast))
		f.SetCellValue(regsheet, "G"+srowx, Make+" "+Model)

		// Overview
		f.SetCellValue(overviewsheet, "D"+srowx, fmtIBA(RiderIBA))

		f.SetCellValue(overviewsheet, "F"+srowx, PillionFirst+" "+PillionLast)
		if cfg.Rally != "rblr" {
			f.SetCellValue(overviewsheet, "E"+srowx, fmtNovice(novicerider))
			f.SetCellValue(overviewsheet, "G"+srowx, fmtIBA(PillionIBA))
			f.SetCellValue(overviewsheet, "H"+srowx, fmtNovice(novicepillion))
		}
		f.SetCellValue(overviewsheet, "I"+srowx, ShortMaker(Make))
		f.SetCellValue(overviewsheet, "J"+srowx, Model)

		f.SetCellValue(overviewsheet, "K"+srowx, Miles)

		if Camp == "Yes" && cfg.Rally == "rblr" {
			f.SetCellInt(overviewsheet, "L"+srowx, 1)
		}
		var cols string = "MNOPQR"
		var col int = 0
		if cfg.Rally == "rblr" {
			col = strings.Index("ABCDEF", string(Route[0])) // Which route is being ridden. Compare the A -, B -, ...
			f.SetCellInt(overviewsheet, string(cols[col])+srowx, 1)

			f.SetCellValue(chksheet, "E"+srowx, rblr_routes[col]) // Carpark
			f.SetCellValue(regsheet, "I"+srowx, rblr_routes[col]) // Registration

			rblr_routes_ridden[col]++
		}

		if includeShopTab {
			//cols = "DEFGH"
			n, _ := excelize.ColumnNameToNumber("D")
			for col = 0; col < len(tshirts); col++ {
				if tshirts[col] > 0 {
					x, _ := excelize.ColumnNumberToName(n + col)
					f.SetCellInt(shopsheet, x+srowx, tshirts[col])
				}
			}
		}

		//cols = "STUVW"
		n, _ := excelize.ColumnNameToNumber("S")
		for col = 0; col < len(tshirts); col++ {
			if tshirts[col] > 0 {
				x, _ := excelize.ColumnNumberToName(n + col)
				f.SetCellInt(overviewsheet, x+srowx, tshirts[col])
			}
		}

		srow++

		//fmt.Printf("%v\n", Entrant2Strings(e))

		if exportingCSV {
			csvW.Write(Entrant2Strings(e))
		}

	} // End reading loop

	if exportingCSV {
		csvW.Flush()
	}

	fmt.Printf("%v entrants written\n", numRiders)

	// Write out totals
	f.SetColWidth(totsheet, "A", "A", 30)
	f.SetColWidth(totsheet, "E", "E", 15)
	f.SetCellStyle(totsheet, "A3", "A16", styleRJ)
	f.SetCellStyle(totsheet, "E3", "E16", styleRJ)
	for i := 3; i <= 16; i++ {
		f.SetRowHeight(totsheet, i, 30)
	}
	f.SetCellValue(totsheet, "A3", "Number of riders")
	f.SetCellValue(totsheet, "A4", "Number of pillions")
	f.SetCellValue(totsheet, "A5", "Number of "+cfg.Novice+"s")
	f.SetCellValue(totsheet, "A6", "Number of IBA members")
	if cfg.Rally == "rblr" {
		f.SetCellValue(totsheet, "A7", "Nearest to Squires")
		f.SetCellValue(totsheet, "A8", "Furthest from Squires")
		f.SetCellValue(totsheet, "A9", "Camping at Squires")
		f.SetCellValue(totsheet, "A10", "Funds raised for Poppy Appeal")
		f.SetCellValue(totsheet, "A11", "A - North clockwise")
		f.SetCellValue(totsheet, "A12", "B - North anti-clockwise")
		f.SetCellValue(totsheet, "A13", "C - South clockwise")
		f.SetCellValue(totsheet, "A14", "D - South anti-clockwise")
		f.SetCellValue(totsheet, "A15", "E - 500 clockwise")
		f.SetCellValue(totsheet, "A16", "F - 500 anti-clockwise")
	}

	f.SetCellInt(totsheet, "B3", numRiders)
	f.SetCellInt(totsheet, "B4", numPillions)
	f.SetCellInt(totsheet, "B5", numNovices)
	f.SetCellInt(totsheet, "B6", numIBAMembers)

	if cfg.Rally == "rblr" {
		f.SetCellInt(totsheet, "B7", shortestSquires)
		f.SetCellInt(totsheet, "B8", longestSquires)
		f.SetCellInt(totsheet, "B9", numCamping)
		if *safemode {

			f.SetCellInt(totsheet, "B10", totSponsorship)
			r := 11
			for i := 0; i < len(rblr_routes_ridden); i++ {
				if rblr_routes_ridden[i] > 0 {
					f.SetCellInt(totsheet, "B"+strconv.Itoa(r), rblr_routes_ridden[i])
				}
				r++
			}
		} else {
			f.SetCellFormula(totsheet, "B10", paysheet+"!I"+strconv.Itoa(srow+1))
			r := 11
			c := "MNOPQR"
			for i := 0; i < len(rblr_routes_ridden); i++ {
				f.SetCellFormula(totsheet, "B"+strconv.Itoa(r), overviewsheet+"!"+string(c[i])+strconv.Itoa(srow+1))
				r++
			}
		}

	}
	f.SetCellStyle(overviewsheet, "A2", "A"+srowx, styleV2)
	f.SetCellStyle(overviewsheet, "B2", "J"+srowx, styleV2L)
	f.SetCellStyle(overviewsheet, "E2", "E"+srowx, styleV2)
	f.SetCellStyle(overviewsheet, "H2", "H"+srowx, styleV2)
	f.SetCellStyle(chksheet, "A2", "A"+srowx, styleV2)
	f.SetCellStyle(chksheet, "B2", "E"+srowx, styleV2L)
	f.SetCellStyle(chksheet, "F2", "G"+srowx, styleRJSmall)
	f.SetCellStyle(chksheet, "H2", "H"+srowx, styleV2)

	if includeShopTab {
		f.SetCellStyle(shopsheet, "A2", "A"+srowx, styleV2)
		f.SetCellStyle(shopsheet, "B2", "C"+srowx, styleV2L)
		f.SetCellStyle(shopsheet, "D2", shop_patch_column+srowx, styleV2)
	}

	f.SetCellStyle(regsheet, "A2", "A"+srowx, styleV2)
	f.SetCellStyle(regsheet, "B2", "C"+srowx, styleV2L)
	f.SetCellStyle(regsheet, "D2", "D"+srowx, styleV)
	f.SetCellStyle(regsheet, "E2", "E"+srowx, styleV2L)
	f.SetCellStyle(regsheet, "F2", "F"+srowx, styleV)
	f.SetCellStyle(regsheet, "G2", "G"+srowx, styleV2L)
	f.SetCellStyle(regsheet, "H2", "H"+srowx, styleV)

	f.SetCellStyle(noksheet, "A2", "A"+srowx, styleV3)

	if cfg.Rally == "rblr" {
		f.SetCellStyle(overviewsheet, "K2", "R"+srowx, styleV)
		f.SetCellStyle(regsheet, "I2", "I"+srowx, styleV2)
		f.SetCellStyle(regsheet, "J2", "J"+srowx, styleV)
	}
	if len(cfg.Tshirts) > 0 {
		n, _ := excelize.ColumnNameToNumber("S")
		x, _ := excelize.ColumnNumberToName(n + len(cfg.Tshirts) - 1)
		f.SetCellStyle(overviewsheet, "S2", x+srowx, styleV)
	}
	if cfg.Patchavail {
		f.SetCellStyle(overviewsheet, overview_patch_column+"2", overview_patch_column+srowx, styleV)
	}

	//f.SetCellStyle(overviewsheet, "G2", "J"+srowx, styleV2)

	f.SetCellStyle(paysheet, "A2", "A"+srowx, styleV3)
	f.SetCellStyle(paysheet, "D2", "J"+srowx, styleV)
	f.SetCellStyle(paysheet, "K2", "K"+srowx, styleT)

	srow++ // Leave a gap before totals

	// L to X
	ncol, _ := excelize.ColumnNameToNumber("L")
	xcol := ""
	srowt := strconv.Itoa(srow)
	if *safemode {
		xcol, _ = excelize.ColumnNumberToName(ncol)
		f.SetCellStyle(overviewsheet, xcol+srowt, xcol+srowt, styleT)
		if cfg.Rally == "rblr" {
			f.SetCellInt(overviewsheet, xcol+srowt, numCamping)
		}
		ncol++
		if cfg.Rally == "rblr" {
			for i := 0; i < len(rblr_routes_ridden); i++ {
				xcol, _ = excelize.ColumnNumberToName(ncol)
				f.SetCellStyle(overviewsheet, xcol+srowt, xcol+srowt, styleT)
				if rblr_routes_ridden[i] > 0 {
					f.SetCellInt(overviewsheet, xcol+srowt, rblr_routes_ridden[i])
				}
				ncol++
			}
		}
		for i := 0; i < num_tshirt_sizes; i++ {
			xcol, _ = excelize.ColumnNumberToName(ncol)
			f.SetCellStyle(overviewsheet, xcol+srowt, xcol+srowt, styleT)
			if totTShirts[i] > 0 {
				f.SetCellInt(overviewsheet, xcol+srowt, totTShirts[i])
			}
			ncol++
		}
		if cfg.Patchavail {
			xcol, _ = excelize.ColumnNumberToName(ncol)
			f.SetCellStyle(overviewsheet, xcol+srowt, xcol+srowt, styleT)
			if totPatches > 0 {
				f.SetCellInt(overviewsheet, xcol+srowt, totPatches)
			}
			ncol++
		}
	} else {
		for _, c := range "LMNOPQRSTUVWX" {
			ff := "sum(" + string(c) + "2:" + string(c) + srowx + ")"
			f.SetCellFormula(overviewsheet, string(c)+strconv.Itoa(srow), "if("+ff+"=0,\"\","+ff+")")
			f.SetCellStyle(overviewsheet, string(c)+strconv.Itoa(srow), string(c)+strconv.Itoa(srow), styleT)
		}
	}

	// Shop totals
	if includeShopTab {
		ncol, _ = excelize.ColumnNameToNumber("D")

		if *safemode {
			for i := 0; i < num_tshirt_sizes; i++ {
				xcol, _ = excelize.ColumnNumberToName(ncol)
				f.SetCellStyle(shopsheet, xcol+srowt, xcol+srowt, styleT)
				if totTShirts[i] > 0 {
					f.SetCellInt(shopsheet, xcol+srowt, totTShirts[i])
				}
				ncol++
			}
			if cfg.Patchavail {
				xcol, _ = excelize.ColumnNumberToName(ncol)
				f.SetCellStyle(shopsheet, xcol+srowt, xcol+srowt, styleT)
				if totPatches > 0 {
					f.SetCellInt(shopsheet, xcol+srowt, totPatches)
				}
				ncol++
			}
		} else {
			for _, c := range "DEFGHI" {
				ff := "sum(" + string(c) + "2:" + string(c) + srowx + ")"
				f.SetCellFormula(shopsheet, string(c)+strconv.Itoa(srow), "if("+ff+"=0,\"\","+ff+")")
				f.SetCellStyle(shopsheet, string(c)+strconv.Itoa(srow), string(c)+strconv.Itoa(srow), styleT)
			}
		}
	}

	if *safemode {
		// paysheet totals
		ncol, _ = excelize.ColumnNameToNumber("D")
		var moneytot int = 0

		// Riders
		xcol, _ = excelize.ColumnNumberToName(ncol)
		f.SetCellStyle(paysheet, xcol+srowt, xcol+srowt, styleT)
		moneytot = numRiders * cfg.Riderfee
		f.SetCellInt(paysheet, xcol+srowt, moneytot)
		ncol++

		// Pillions
		xcol, _ = excelize.ColumnNumberToName(ncol)
		f.SetCellStyle(paysheet, xcol+srowt, xcol+srowt, styleT)
		moneytot = numPillions * cfg.Pillionfee
		f.SetCellInt(paysheet, xcol+srowt, moneytot)
		ncol++

		// T-shirts
		xcol, _ = excelize.ColumnNumberToName(ncol)
		moneytot = grandTotalTShirts * cfg.Tshirtcost
		if num_tshirt_sizes > 0 {
			f.SetCellStyle(paysheet, xcol+srowt, xcol+srowt, styleT)
			f.SetCellInt(paysheet, xcol+srowt, moneytot)
		}
		ncol++

		// Patches
		xcol, _ = excelize.ColumnNumberToName(ncol)
		moneytot = totPatches * cfg.Patchcost
		if cfg.Patchavail {
			f.SetCellStyle(paysheet, xcol+srowt, xcol+srowt, styleT)
			f.SetCellInt(paysheet, xcol+srowt, moneytot)
		}
		ncol++

		ncol++ // Skip cheque @ Squires

		// Sponsorship
		xcol, _ = excelize.ColumnNumberToName(ncol)
		moneytot = totSponsorship
		if cfg.Sponsorship {
			f.SetCellStyle(paysheet, xcol+srowt, xcol+srowt, styleT)
			f.SetCellInt(paysheet, xcol+srowt, moneytot)
		}
		ncol++

		// Total received
		xcol, _ = excelize.ColumnNumberToName(ncol)
		f.SetCellStyle(paysheet, xcol+srowt, xcol+srowt, styleT)
		moneytot = totPayment + totCash
		f.SetCellInt(paysheet, xcol+srowt, moneytot)
		ncol++

	} else {
		for _, c := range "DEFGHIJKL" {
			ff := "sum(" + string(c) + "2:" + string(c) + srowx + ")"
			f.SetCellFormula(paysheet, string(c)+strconv.Itoa(srow), "if("+ff+"=0,\"\","+ff+")")
			f.SetCellStyle(paysheet, string(c)+strconv.Itoa(srow), string(c)+strconv.Itoa(srow), styleT)
		}
	}
	f.SetActiveSheet(0)
	f.SetCellValue(overviewsheet, "A1", "No.")
	f.SetCellValue(noksheet, "A1", "No.")
	f.SetCellValue(paysheet, "A1", "No.")
	f.SetCellValue(chksheet, "A1", "No.")
	f.SetCellValue(regsheet, "A1", "No.")
	f.SetColWidth(overviewsheet, "A", "A", 5)
	f.SetColWidth(noksheet, "A", "A", 5)
	f.SetColWidth(paysheet, "A", "A", 5)
	f.SetColWidth(regsheet, "A", "A", 5)

	if includeShopTab {
		f.SetCellValue(shopsheet, "A1", "No.")
		f.SetColWidth(shopsheet, "A", "A", 5)
		f.SetCellValue(shopsheet, "B1", "Rider(first)")
		f.SetCellValue(shopsheet, "C1", "Rider(last)")
		f.SetColWidth(shopsheet, "B", "I", 12)
	}

	f.SetColWidth(overviewsheet, "B", "D", 1)
	f.SetColWidth(regsheet, "B", "C", 12)

	f.SetColWidth(regsheet, "D", "D", 5)
	f.SetColWidth(regsheet, "E", "E", 20)
	f.SetColWidth(regsheet, "F", "F", 5)
	f.SetColWidth(regsheet, "G", "G", 30)
	f.SetColWidth(regsheet, "H", "H", 5)
	f.SetColWidth(regsheet, "I", "I", 10)
	f.SetColWidth(regsheet, "J", "J", 5)

	f.SetCellValue(regsheet, "B1", "Rider(first)")
	f.SetCellValue(regsheet, "C1", "Rider(last)")
	f.SetCellValue(regsheet, "D1", "✓")
	f.SetCellValue(paysheet, "B1", "Rider(first)")
	f.SetCellValue(paysheet, "C1", "Rider(last)")
	f.SetCellValue(chksheet, "B1", "Rider(first)")
	f.SetCellValue(chksheet, "C1", "Rider(last)")
	f.SetCellValue(chksheet, "D1", "Bike")
	f.SetCellValue(regsheet, "E1", "Pillion")
	f.SetCellValue(regsheet, "F1", "✓")
	f.SetCellValue(chksheet, "F1", "Odo")
	f.SetCellValue(chksheet, "G1", "Time")

	if cfg.Rally == "rblr" {
		f.SetCellValue(chksheet, "E1", "Route")
		f.SetCellValue(regsheet, "I1", "Route")
		f.SetCellValue(regsheet, "J1", "✓")
	}

	f.SetCellValue(chksheet, "H1", "Notes")

	f.SetCellValue(paysheet, "D1", "Entry")
	f.SetCellValue(paysheet, "E1", "Pillion")
	f.SetCellValue(regsheet, "G1", "Bike")
	f.SetCellValue(regsheet, "H1", "✓")
	if len(cfg.Tshirts) > 0 {
		f.SetCellValue(paysheet, "F1", "T-shirts")
	}
	if cfg.Patchavail {
		f.SetCellValue(paysheet, "G1", "Patches")
	}
	if cfg.Sponsorship {
		f.SetCellValue(paysheet, "H1", cfg.Fundsonday)
		f.SetCellValue(paysheet, "I1", "Total Sponsorship")
	}
	//f.SetCellValue(paysheet, "K1", "+Cash")
	f.SetCellValue(paysheet, "J1", "Total received")
	f.SetCellValue(paysheet, "K1", " !!!")
	f.SetColWidth(paysheet, "B", "B", 12)
	f.SetColWidth(paysheet, "C", "C", 12)
	f.SetColWidth(paysheet, "D", "G", 8)
	f.SetColWidth(paysheet, "H", "J", 15)
	f.SetColWidth(paysheet, "J", "J", 15)
	f.SetColWidth(paysheet, "K", "K", 15)

	f.SetCellValue(overviewsheet, "B1", "Rider(first)")
	f.SetCellValue(overviewsheet, "C1", "Rider(last)")
	f.SetColWidth(overviewsheet, "B", "C", 12)

	f.SetColWidth(chksheet, "B", "C", 12)
	f.SetColWidth(chksheet, "D", "D", 30)
	f.SetColWidth(chksheet, "F", "G", 10)
	f.SetColWidth(chksheet, "H", "H", 40)

	f.SetColWidth(overviewsheet, "D", "D", 8) // Rider IBA

	f.SetCellValue(noksheet, "B1", "Rider(first)")
	f.SetCellValue(noksheet, "C1", "Rider(last)")
	f.SetColWidth(noksheet, "B", "C", 15)

	f.SetCellValue(noksheet, "D1", "Mobile")
	f.SetCellValue(noksheet, "E1", "NOK name")
	f.SetCellValue(noksheet, "F1", "Relationship")
	f.SetCellValue(noksheet, "G1", "Contact number")
	f.SetColWidth(noksheet, "D", "G", 20)

	f.SetCellValue(overviewsheet, "D1", "IBA #")
	f.SetCellValue(overviewsheet, "E1", strings.Title(cfg.Novice))
	f.SetCellValue(overviewsheet, "F1", "Pillion")
	f.SetColWidth(overviewsheet, "F", "F", 19)
	f.SetCellValue(overviewsheet, "G1", "IBA #")
	f.SetCellValue(overviewsheet, "H1", strings.Title(cfg.Novice))

	//f.SetColVisible(overviewsheet, "B:D", false)

	f.SetCellValue(overviewsheet, "I1", "Make")
	f.SetColWidth(overviewsheet, "I", "I", 10)
	f.SetCellValue(overviewsheet, "J1", "Model")
	f.SetColWidth(overviewsheet, "J", "J", 20)

	if cfg.Rally == "rblr" {
		f.SetCellValue(overviewsheet, "K1", " Miles to Squires")
		f.SetColWidth(overviewsheet, "K", "K", 5)

		f.SetCellValue(overviewsheet, "L1", " Camping")
		f.SetColWidth(overviewsheet, "L", "R", 3)

		f.SetCellValue(overviewsheet, "M1", rblr_routes[0])
		f.SetCellValue(overviewsheet, "N1", rblr_routes[1])
		f.SetCellValue(overviewsheet, "O1", rblr_routes[2])
		f.SetCellValue(overviewsheet, "P1", rblr_routes[3])
		f.SetCellValue(overviewsheet, "Q1", rblr_routes[4])
		f.SetCellValue(overviewsheet, "R1", rblr_routes[5])

	}

	if len(cfg.Tshirts) > 0 {
		n, _ := excelize.ColumnNameToNumber("S")
		for i := 0; i < len(cfg.Tshirts); i++ {
			x, _ := excelize.ColumnNumberToName(n + i)
			f.SetColWidth(overviewsheet, x, x, 3)
			f.SetCellValue(overviewsheet, x+"1", tshirt_sizes[i])
		}
	}
	if cfg.Patchavail {
		f.SetColWidth(overviewsheet, overview_patch_column, overview_patch_column, 3)
		f.SetCellValue(overviewsheet, overview_patch_column+"1", " Patches")
	}
	if includeShopTab {
		if len(cfg.Tshirts) > 0 {
			n, _ := excelize.ColumnNameToNumber("D")
			for i := 0; i < len(cfg.Tshirts); i++ {
				x, _ := excelize.ColumnNumberToName(n + i)
				f.SetCellValue(shopsheet, x+"1", tshirt_sizes[i])
			}
		}
		if cfg.Patchavail {
			f.SetCellValue(shopsheet, shop_patch_column+"1", " Patches")
		}
	}

	f.SetRowHeight(overviewsheet, 1, 70)
	f.SetRowHeight(noksheet, 1, 20)
	f.SetRowHeight(paysheet, 1, 70)

	sort.Slice(bikes, func(i, j int) bool { return bikes[i].make < bikes[j].make })
	//fmt.Printf("%v\n", bikes)
	srow = 2
	ntot := 0
	for i := 0; i < len(bikes); i++ {

		f.SetCellValue(totsheet, "E"+strconv.Itoa(srow+1), bikes[i].make)
		f.SetCellInt(totsheet, "F"+strconv.Itoa(srow+1), bikes[i].num)
		f.SetCellStyle(totsheet, "F"+strconv.Itoa(srow+1), "F"+strconv.Itoa(srow), styleRJ)

		ntot += bikes[i].num
		srow++
	}

	srow++

	setPageTitle(f, overviewsheet)
	setPageTitle(f, noksheet)
	setPageTitle(f, paysheet)
	setPageTitle(f, totsheet)
	setPageTitle(f, chksheet)
	setPageTitle(f, regsheet)

	setPagePane(f, overviewsheet)
	setPagePane(f, noksheet)
	setPagePane(f, paysheet)
	setPagePane(f, chksheet)
	setPagePane(f, regsheet)

	if includeShopTab {
		setPageTitle(f, shopsheet)
		setPagePane(f, shopsheet)
	}

	markSpreadsheet(f, cfg)

	// Save spreadsheet by the given path.
	if err := f.SaveAs(*xlsName); err != nil {
		fmt.Println(err)
	}
}

func renameSheet(f *excelize.File, oldname *string, newname string) {

	f.SetSheetName(*oldname, newname)
	*oldname = newname

}

func makeFile(csvname string) *os.File {

	file, err := os.Create(csvname)
	if err != nil {
		panic(err)
	}
	return file

}

func makeCSVFile(f *os.File) *csv.Writer {

	writer := csv.NewWriter(f)
	writer.Write(EntrantHeaders())
	return writer
}

func loadCSVFile(db *sql.DB) {

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

}

func makeSQLTable(db *sql.DB) {

	var x string = ""
	re := regexp.MustCompile(`\bRiderNumber\b`)
	if !re.Match([]byte(dbfieldsx)) {
		x = ",RiderNumber"
	}
	x += ",FinalRiderNumber"
	db.Exec("PRAGMA foreign_keys=OFF")
	db.Exec("BEGIN TRANSACTION")
	_, err := db.Exec("DROP TABLE IF EXISTS entrants")
	if err != nil {
		log.Fatal(err)
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

func markSpreadsheet(f *excelize.File, cfg *Config) {

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
	err := f.SetDocProps(&dp)
	if err != nil {
		fmt.Printf("%v\n", err)
	}

}

func initStyles(f *excelize.File) {

	// Totals
	styleT, _ = f.NewStyle(`{	
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
	styleH, _ = f.NewStyle(`{
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
	styleH2, _ = f.NewStyle(`{
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

	// Data
	styleV, _ = f.NewStyle(`{
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
	styleV2, _ = f.NewStyle(`{
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
			"border": [
				{
					"type": "bottom",
					"color": "000000",
					"style": 1
				}]		
		}`)

	styleV2L, _ = f.NewStyle(`{
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
				"wrap_text": true
			},
			"border": [
				{
					"type": "bottom",
					"color": "000000",
					"style": 1
				}]		
		}`)

	styleV3, _ = f.NewStyle(`{
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
			}
		}`)

	// styleW for highlighting, particularly errorneous, cells
	styleW, _ = f.NewStyle(`{ 
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

	styleRJ, _ = f.NewStyle(`{ 
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

	styleRJSmall, _ = f.NewStyle(`{ 
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

	f.SetDefaultFont("Arial")

}

func extractMakeModel(bike string) (string, string) {

	if strings.TrimSpace(bike) == "" {
		return "", ""
	}
	re := regexp.MustCompile(`\d*\s*([A-Za-z\-]*)\s*(.*)`)
	sm := re.FindSubmatch([]byte(bike))
	if len(sm) < 3 {
		return string(sm[1]), ""
	}
	return string(sm[1]), string(sm[2])

}

func fmtIBA(x string) string {

	if x == "-1" {
		return "n/a"
	}
	return strings.ReplaceAll(x, ".0", "")

}

func fmtNovice(x string) string {

	if strings.Contains(x, cfg.Novice) {
		return "Yes"
	}
	return ""
}

func fmtNoviceYN(x string) string {
	if fmtNovice(x) != "" && x[0] != 'N' && x[0] != 'n' {
		return "Y"
	} else {
		return "N"
	}
}

func fmtOdoKM(x string) string {

	y := strings.ToUpper(x)
	if len(y) > 0 && y[0] == 'K' {
		return "K"
	}
	return "M"

}

func fmtCampingYN(x string) string {

	y := strings.ToUpper(x)
	if len(y) > 0 && y[0] == 'Y' {
		return "Y"
	}
	return "N"

}

func ShortMaker(x string) string {

	p := strings.Index(x, "-")
	if p < 0 {
		return x
	}
	return x[0:p]
}
