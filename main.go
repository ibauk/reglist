package main

/*
 * This is a quick and dirty, yes really, transformer to create a "Registration list"
 * spreadsheet ready for the RBLR1000.
 *
 * It will be run several times before a "final" version shortly before the ride date.
 *
 * It must be kept in sync with the Wufoo form used to capture entrant records. The
 * process is download CSV from Wufoo; create and load rblrdata.db SQLite database;
 * remove column headers record then run this program to produce an XLSX file.
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
	"sort"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	_ "github.com/mattn/go-sqlite3"
)

var csvName *string = flag.String("csv", "rblrentrants.csv", "Path to CSV downloaded from Wufoo")
var sqlName *string = flag.String("sql", "rblrdata.db", "Path to SQLite database")
var xlsName *string = flag.String("xls", "reglist.xlsx", "Path to output XLSX")
var noCSV *bool = flag.Bool("nocsv", false, "Don't load a CSV file, just use the SQL database")

const dbFields = `"EntryId","Date_Created","Created_By","Date_Updated","Updated_By",
					"IP_Address","Last_Page_Accessed","Completion_Status","RiderName","RiderLast","RiderIBANumber",
					"Is_this_your_first_RBLR1000",
					"Are_you_riding_with_a_pillion","PillionName","PillionLast","PillionIBANumber",
					"Pillion_first_RBLR1000",
					"Rider_Address","Address_Line_2","City","State_Province_Region","Postal_Zip_Code","Country",
					"Mobilephone","Email",
					"BikeMakeModel","Registration","Odometer_counts",
					"Emergencycontactname","Emergencycontactnumber","Emergencycontactrelationship",
					"ao_BCM","Detailed_Instructions",
					"RBLR1000Tshirt1","RBLR1000Tshirt2",
					"WhichRoute",
					"FreeCamping","MilestravelledToSquires",
					"Admin_markers","Sponsorshipmoney",
					"Patches","Cash",
					"PaymentStatus","PaymentTotal","Payment_Currency","Payment_Confirmation","Payment_Merchant"`

const regsheet = "Sheet1"
const noksheet = "Sheet2"
const bikesheet = "Sheet3"
const paysheet = "Sheet4"

const sqlx = `SELECT RiderName,RiderLast,ifnull(RiderIBANumber,''),
ifnull(PillionName,''),ifnull(PillionLast,''),ifnull(PillionIBANumber,''),
BikeMakeModel,round(MilesTravelledToSquires),
FreeCamping,WhichRoute,RBLR1000Tshirt1,RBLR1000Tshirt2,ifnull(Patches,'0'),ifnull(Cash,'0'),
Mobilephone,Emergencycontactname,Emergencycontactnumber,Emergencycontactrelationship,
EntryId,PaymentTotal,Sponsorshipmoney,PaymentStatus
FROM entrants ORDER BY upper(RiderLast),upper(RiderName)`

var styleH, styleH2, styleT, styleV, styleV2, styleW int

func proper(x string) string {
	var xx = strings.TrimSpace(x)
	if strings.ToLower(xx) == xx {
		return strings.Title(xx)
	}
	return xx

}

func showRecordCount(db *sql.DB) {

	rows, err := db.Query("SELECT count(*) FROM entrants;")
	if err != nil {
		log.Fatal(err)
	}
	var n int64
	rows.Next()
	err = rows.Scan(&n)
	if err != nil {
		log.Fatal(err)
	}
	fmt.Printf("%v entrants loaded\n", n)
	rows.Close()
}

func formatSheet(f *excelize.File, sheetName string) {

	f.SetPageLayout(
		sheetName,
		excelize.PageLayoutOrientation(excelize.OrientationLandscape),
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

func main() {
	flag.Parse()

	fmt.Printf("IBAUK Reglist v0.0.3\nCopyright (c) 2021 Bob Stammers\n\n")

	db, err := sql.Open("sqlite3", *sqlName)
	if err != nil {
		log.Fatal(err)
	}
	if !*noCSV {
		loadCSVFile(db)
	}
	showRecordCount(db)

	f := excelize.NewFile()

	initStyles(f)
	f.SetDefaultFont("Arial")
	// First sheet is called Sheet1
	formatSheet(f, regsheet)
	f.NewSheet(noksheet)
	formatSheet(f, noksheet)
	f.NewSheet(bikesheet)
	formatSheet(f, bikesheet)
	f.NewSheet(paysheet)
	formatSheet(f, paysheet)

	tshirtsizes := [...]string{"S", "M", "L", "XL", "XXL"}

	f.SetCellStyle(regsheet, "A1", "A1", styleH2)
	f.SetCellStyle(regsheet, "E1", "J1", styleH2)
	f.SetCellStyle(noksheet, "A1", "G1", styleH2)

	f.SetCellStyle(bikesheet, "A1", "B1", styleH2)
	f.SetCellStyle(paysheet, "A1", "K1", styleH2)

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

	for rows1.Next() {
		var RiderFirst string
		var RiderLast string
		var RiderIBA string
		var PillionFirst, PillionLast, PillionIBA string
		var Bike, Make, Model string
		var Miles, EntryID string
		var Camp, Route, T1, T2, Patches string
		var Mobile, NokName, NokNumber, NokRelation string
		var PayTot string
		var Sponsor, Paid, Cash string
		tshirts := [...]int{0, 0, 0, 0, 0}

		err2 := rows1.Scan(&RiderFirst, &RiderLast, &RiderIBA, &PillionFirst, &PillionLast, &PillionIBA,
			&Bike, &Miles, &Camp, &Route, &T1, &T2, &Patches, &Cash,
			&Mobile, &NokName, &NokNumber, &NokRelation, &EntryID, &PayTot, &Sponsor, &Paid)
		if err2 != nil {
			log.Fatal(err2)
		}
		//fmt.Printf("%v %v\n", RiderFirst, RiderLast)
		var tottshirts int = 0
		for i := 0; i < len(tshirtsizes); i++ {
			if tshirtsizes[i] == T1 {
				tshirts[i]++
				tottshirts++
			}
			if tshirtsizes[i] == T2 {
				tshirts[i]++
				tottshirts++
			}
		}
		srowx = strconv.Itoa(srow)
		var tmp []string = strings.Split(Bike, " ")
		Make = proper(tmp[0])

		var i int
		var ok bool = true
		for i = 0; i < len(bikes); i++ {
			if bikes[i].make == Make {
				bikes[i].num++
				ok = false
			}
		}

		if ok {
			bm := bikemake{Make, 1}
			bikes = append(bikes, bm)
		}
		Model = strings.Join(tmp[1:], " ")
		f.SetCellInt(regsheet, "A"+srowx, intval(EntryID))
		f.SetCellInt(noksheet, "A"+srowx, intval(EntryID))
		f.SetCellInt(paysheet, "A"+srowx, intval(EntryID))
		f.SetCellValue(regsheet, "E"+srowx, strings.Title(RiderFirst))
		f.SetCellValue(regsheet, "F"+srowx, strings.Title(RiderLast))
		f.SetCellValue(noksheet, "B"+srowx, strings.Title(RiderFirst))
		f.SetCellValue(noksheet, "C"+srowx, strings.Title(RiderLast))
		f.SetCellValue(paysheet, "B"+srowx, strings.Title(RiderFirst))
		f.SetCellValue(paysheet, "C"+srowx, strings.Title(RiderLast))

		f.SetCellInt(paysheet, "D"+srowx, 20) // Basic entry fee

		if PillionFirst != "" && PillionLast != "" {
			f.SetCellInt(paysheet, "E"+srowx, 10)
		}
		if tottshirts > 0 {
			f.SetCellInt(paysheet, "F"+srowx, 10*tottshirts)
		}

		f.SetCellValue(noksheet, "D"+srowx, Mobile)
		f.SetCellValue(noksheet, "E"+srowx, strings.Title(NokName))
		f.SetCellValue(noksheet, "F"+srowx, strings.Title(NokRelation))
		f.SetCellValue(noksheet, "G"+srowx, NokNumber)

		f.SetCellValue(regsheet, "G"+srowx, strings.ReplaceAll(RiderIBA, ".0", ""))
		f.SetCellValue(regsheet, "H"+srowx, strings.Title(PillionFirst)+" "+strings.Title(PillionLast))
		f.SetCellValue(regsheet, "I"+srowx, proper(Make))
		f.SetCellValue(regsheet, "J"+srowx, proper(Model))
		f.SetCellValue(regsheet, "K"+srowx, Miles)
		if Camp == "Yes" {
			f.SetCellInt(regsheet, "L"+srowx, 1)
		}
		var cols string = "MNOPQR"
		var col int = strings.Index("ABCDEF", string(Route[0]))
		f.SetCellInt(regsheet, string(cols[col])+srowx, 1)
		cols = "STUVW"
		for col = 0; col < len(tshirts); col++ {
			if tshirts[col] > 0 {
				f.SetCellInt(regsheet, string(cols[col])+srowx, tshirts[col])
			}
		}
		if Patches[0] == '1' {
			f.SetCellInt(regsheet, "X"+srowx, 1)
			f.SetCellInt(paysheet, "G"+srowx, 5)
		} else if Patches[0] == '2' {
			f.SetCellInt(noksheet, "X"+srowx, 2)
			f.SetCellInt(paysheet, "G"+srowx, 10)
		}

		Sponsorship := "0"
		if strings.Contains(Sponsor, "Include") {
			Sponsorship = "50"
			//f.SetCellInt(paysheet, "I"+srowx, 50)
		}
		sf := "H" + srowx + "+" + Sponsorship
		f.SetCellFormula(paysheet, "I"+srowx, "if("+sf+"=0,\"0\","+sf+")")

		intCash := intval(Cash)
		if false && intCash != 0 {
			f.SetCellInt(paysheet, "K"+srowx, intval(Cash))
		}
		//f.SetCellInt(paysheet, "J"+srowx, intval(PayTot))
		f.SetCellFormula(paysheet, "J"+srowx, "H"+srowx+"+"+strconv.Itoa(intCash)+"+"+strconv.Itoa(intval(PayTot)))

		if Paid == "Unpaid" {
			f.SetCellValue(paysheet, "K"+srowx, " UNPAID")
			f.SetCellStyle(paysheet, "K"+srowx, "K"+srowx, styleW)
		} else {
			ff := "J" + srowx + "-(sum(D" + srowx + ":G" + srowx + ")+I" + srowx + ")"
			f.SetCellFormula(paysheet, "K"+srowx, "if("+ff+"=0,\"\","+ff+")")
		}
		// =IF(A2-A3=0,””,A2-A3)
		// =IF(J11-(SUM(D11:G11)+ISBLANK(I11),0,I11))=0,"",J11-(SUM(D11:G11)+ISBLANK(I11),0,I11)))
		srow++
	}

	f.SetCellStyle(regsheet, "B1", "D1", styleH)
	f.SetCellStyle(regsheet, "K1", "X1", styleH)
	f.SetCellStyle(regsheet, "A2", "D"+srowx, styleV)
	f.SetCellStyle(noksheet, "A2", "A"+srowx, styleV)
	f.SetCellStyle(regsheet, "K2", "X"+srowx, styleV)
	f.SetCellStyle(regsheet, "E2", "J"+srowx, styleV2)

	f.SetCellStyle(paysheet, "A2", "A"+srowx, styleV)
	f.SetCellStyle(paysheet, "D2", "J"+srowx, styleV)
	f.SetCellStyle(paysheet, "K2", "K"+srowx, styleT)

	srow++ // Leave a gap before totals

	// L to X
	for _, c := range "LMNOPQRSTUVWX" {
		ff := "sum(" + string(c) + "2:" + string(c) + srowx + ")"
		f.SetCellFormula(regsheet, string(c)+strconv.Itoa(srow), "if("+ff+"=0,\"\","+ff+")")
		f.SetCellStyle(regsheet, string(c)+strconv.Itoa(srow), string(c)+strconv.Itoa(srow), styleT)
	}

	for _, c := range "DEFGHIJKL" {
		ff := "sum(" + string(c) + "2:" + string(c) + srowx + ")"
		f.SetCellFormula(paysheet, string(c)+strconv.Itoa(srow), "if("+ff+"=0,\"\","+ff+")")
		f.SetCellStyle(paysheet, string(c)+strconv.Itoa(srow), string(c)+strconv.Itoa(srow), styleT)
	}

	f.SetActiveSheet(0)
	f.SetCellValue(regsheet, "A1", "No.")
	f.SetCellValue(noksheet, "A1", "No.")
	f.SetCellValue(paysheet, "A1", "No.")
	f.SetColWidth(regsheet, "A", "A", 5)
	f.SetColWidth(noksheet, "A", "A", 5)
	f.SetColWidth(paysheet, "A", "A", 5)

	f.SetCellValue(regsheet, "B1", " Registered")
	f.SetCellValue(regsheet, "C1", " Started")
	f.SetCellValue(regsheet, "D1", " Finished")
	f.SetColWidth(regsheet, "B", "D", 3)

	f.SetCellValue(paysheet, "B1", "Rider(first)")
	f.SetCellValue(paysheet, "C1", "Rider(last)")
	f.SetCellValue(paysheet, "D1", "Entry")
	f.SetCellValue(paysheet, "E1", "Pillion")
	f.SetCellValue(paysheet, "F1", "T-shirts")
	f.SetCellValue(paysheet, "G1", "Patches")
	f.SetCellValue(paysheet, "H1", "Cheque @ Squires")
	f.SetCellValue(paysheet, "I1", "Total Sponsorship")
	//f.SetCellValue(paysheet, "K1", "+Cash")
	f.SetCellValue(paysheet, "J1", "Total cash")
	f.SetCellValue(paysheet, "K1", " !!!")
	f.SetColWidth(paysheet, "B", "B", 12)
	f.SetColWidth(paysheet, "C", "C", 12)
	f.SetColWidth(paysheet, "D", "G", 8)
	f.SetColWidth(paysheet, "H", "J", 15)
	f.SetColWidth(paysheet, "J", "J", 15)
	f.SetColWidth(paysheet, "K", "K", 15)

	f.SetCellValue(regsheet, "E1", "Rider(first)")
	f.SetCellValue(regsheet, "F1", "Rider(last)")
	f.SetColWidth(regsheet, "E", "E", 12)
	f.SetColWidth(regsheet, "F", "F", 12)
	f.SetColWidth(regsheet, "G", "G", 8)

	f.SetCellValue(noksheet, "B1", "Rider(first)")
	f.SetCellValue(noksheet, "C1", "Rider(last)")
	f.SetColWidth(noksheet, "B", "C", 15)

	f.SetCellValue(noksheet, "D1", "Mobile")
	f.SetCellValue(noksheet, "E1", "NOK name")
	f.SetCellValue(noksheet, "F1", "Relationship")
	f.SetCellValue(noksheet, "G1", "Contact number")
	f.SetColWidth(noksheet, "D", "G", 20)

	f.SetCellValue(bikesheet, "A1", "Make")
	f.SetCellValue(bikesheet, "B1", "Number")
	f.SetColWidth(bikesheet, "A", "A", 10)

	f.SetCellValue(regsheet, "G1", "IBA #")
	f.SetCellValue(regsheet, "H1", "Pillion")
	f.SetColWidth(regsheet, "H", "H", 12)

	f.SetCellValue(regsheet, "I1", "Make")
	f.SetColWidth(regsheet, "I", "I", 10)
	f.SetCellValue(regsheet, "J1", "Model")
	f.SetColWidth(regsheet, "J", "J", 20)
	f.SetCellValue(regsheet, "K1", " Miles to Squires")
	f.SetColWidth(regsheet, "K", "K", 5)

	f.SetCellValue(regsheet, "L1", " Camping")
	f.SetColWidth(regsheet, "L", "X", 3)

	f.SetCellValue(regsheet, "M1", " A-NC")
	f.SetCellValue(regsheet, "N1", " B-NAC")
	f.SetCellValue(regsheet, "O1", " C-SC")
	f.SetCellValue(regsheet, "P1", " D-SAC")
	f.SetCellValue(regsheet, "Q1", " E-500C")
	f.SetCellValue(regsheet, "R1", " F-500AC")

	f.SetCellValue(regsheet, "S1", " T-shirt S")
	f.SetCellValue(regsheet, "T1", " T-shirt M")
	f.SetCellValue(regsheet, "U1", " T-shirt L")
	f.SetCellValue(regsheet, "V1", " T-shirt XL")
	f.SetCellValue(regsheet, "W1", " T-shirt XXL")

	f.SetCellValue(regsheet, "X1", " Patches")

	f.SetRowHeight(regsheet, 1, 70)
	f.SetRowHeight(noksheet, 1, 20)
	f.SetRowHeight(bikesheet, 1, 20)
	f.SetRowHeight(paysheet, 1, 70)

	sort.Slice(bikes, func(i, j int) bool { return bikes[i].make < bikes[j].make })
	//fmt.Printf("%v\n", bikes)
	srow = 2
	ntot := 0
	for i := 0; i < len(bikes); i++ {
		f.SetCellValue(bikesheet, "A"+strconv.Itoa(srow), bikes[i].make)
		f.SetCellInt(bikesheet, "B"+strconv.Itoa(srow), bikes[i].num)
		f.SetCellStyle(bikesheet, "B"+strconv.Itoa(srow), "B"+strconv.Itoa(srow), styleV)
		ntot += bikes[i].num
		srow++
	}

	srow++
	f.SetCellInt(bikesheet, "B"+strconv.Itoa(srow), ntot)
	err = f.SetCellStyle(bikesheet, "B"+strconv.Itoa(srow), "B"+strconv.Itoa(srow), styleT)
	if err != nil {
		log.Fatal(err)
	}

	f.SetSheetName(regsheet, "Registration")
	f.SetSheetName(noksheet, "NOK list")
	f.SetSheetName(bikesheet, "Bikes")
	f.SetSheetName(paysheet, "Money")
	// Save spreadsheet by the given path.
	if err := f.SaveAs(*xlsName); err != nil {
		fmt.Println(err)
	}
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

		sqlx := "INSERT INTO entrants (" + dbFields + ") VALUES("
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

	_, err := db.Exec("DROP TABLE IF EXISTS entrants")
	if err != nil {
		log.Fatal(err)
	}
	db.Exec("PRAGMA foreign_keys=OFF")
	db.Exec("BEGIN TRANSACTION")
	_, err = db.Exec("CREATE TABLE entrants (" + dbFields + ")")
	if err != nil {
		log.Fatal(err)
	}

}

func intval(x string) int {

	n, _ := strconv.Atoi(strings.Replace(x, ".0", "", 1)) // There shouldn't be any decimals on any of the input so ...
	return n

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

}
