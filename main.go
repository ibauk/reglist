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
	"flag"
	"fmt"
	"log"
	"sort"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	_ "github.com/mattn/go-sqlite3"
)

var sqlName *string = flag.String("sql", "rblrdata.db", "Path to SQLite database")
var xlsName *string = flag.String("xls", "2021-RBLR-Registration-List.xlsx", "Path to output XLSX")

func proper(x string) string {
	var xx = strings.TrimSpace(x)
	if strings.ToLower(xx) == xx {
		return strings.Title(xx)
	}
	return xx

}

func main() {
	flag.Parse()

	fmt.Printf("IBAUK Reglist v0.0.1\nCopyright (c) 2021 Bob Stammers\n\n")

	db, err := sql.Open("sqlite3", *sqlName)
	if err != nil {
		log.Fatal(err)
	}
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
	f := excelize.NewFile()
	f.SetDefaultFont("Arial")
	// First sheet is called Sheet1

	// Create a new sheet.
	f.NewSheet("Sheet2")
	f.NewSheet("Sheet3")
	tshirtsizes := [...]string{"S", "M", "L", "XL", "XXL"}

	styleT, _ := f.NewStyle(`{
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
			"color": "black"
		}
	}`)

	styleH, _ := f.NewStyle(`{
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
		"fill":{"type":"pattern","color":["#E0EBF5"],"pattern":1}	}`)

	styleH2, _ := f.NewStyle(`{
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
			"fill":{"type":"pattern","color":["#E0EBF5"],"pattern":1}	}`)

	styleV, _ := f.NewStyle(`{
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

	err = f.SetCellStyle("Sheet1", "A1", "A1", styleH2)
	err = f.SetCellStyle("Sheet1", "E1", "J1", styleH2)
	err = f.SetCellStyle("Sheet2", "A1", "G1", styleH2)

	err = f.SetCellStyle("Sheet3", "A1", "B1", styleH2)

	sqlx := "SELECT RiderName,RiderLast,ifnull(RiderIBANumber,''),ifnull(PillionName,''),ifnull(PillionLast,''),ifnull(PillionIBANumber,''),BikeMakeModel,round(MilesTravelledToSquires)"
	sqlx += ",FreeCamping,WhichRoute,RBLR1000Tshirt1,RBLR1000Tshirt2,ifnull(Patches,'0')"
	sqlx += ",Mobilephone,Emergencycontactname,Emergencycontactnumber,Emergencycontactrelationship"
	sqlx += ",EntryId"
	sqlx += " FROM entrants ORDER BY upper(RiderLast),upper(RiderName)"

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
		var Miles, EntryID int
		var Camp, Route, T1, T2, Patches string
		var Mobile, NokName, NokNumber, NokRelation string

		tshirts := [...]int{0, 0, 0, 0, 0}

		err2 := rows1.Scan(&RiderFirst, &RiderLast, &RiderIBA, &PillionFirst, &PillionLast, &PillionIBA,
			&Bike, &Miles, &Camp, &Route, &T1, &T2, &Patches,
			&Mobile, &NokName, &NokNumber, &NokRelation, &EntryID)
		if err2 != nil {
			log.Fatal(err2)
		}
		//fmt.Printf("%v %v\n", RiderFirst, RiderLast)
		for i := 0; i < len(tshirtsizes); i++ {
			if tshirtsizes[i] == T1 {
				tshirts[i]++
			}
			if tshirtsizes[i] == T2 {
				tshirts[i]++
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
		f.SetCellInt("Sheet1", "A"+srowx, EntryID)
		f.SetCellInt("Sheet2", "A"+srowx, EntryID)
		f.SetCellValue("Sheet1", "E"+srowx, strings.Title(RiderFirst))
		f.SetCellValue("Sheet1", "F"+srowx, strings.Title(RiderLast))
		f.SetCellValue("Sheet2", "B"+srowx, strings.Title(RiderFirst))
		f.SetCellValue("Sheet2", "C"+srowx, strings.Title(RiderLast))
		f.SetCellValue("Sheet2", "D"+srowx, Mobile)
		f.SetCellValue("Sheet2", "E"+srowx, strings.Title(NokName))
		f.SetCellValue("Sheet2", "F"+srowx, strings.Title(NokRelation))
		f.SetCellValue("Sheet2", "G"+srowx, NokNumber)

		f.SetCellValue("Sheet1", "G"+srowx, strings.ReplaceAll(RiderIBA, ".0", ""))
		f.SetCellValue("Sheet1", "H"+srowx, strings.Title(PillionFirst)+" "+strings.Title(PillionLast))
		f.SetCellValue("Sheet1", "I"+srowx, proper(Make))
		f.SetCellValue("Sheet1", "J"+srowx, proper(Model))
		f.SetCellValue("Sheet1", "K"+srowx, Miles)
		if Camp == "Yes" {
			f.SetCellInt("Sheet1", "L"+srowx, 1)
		}
		var cols string = "MNOPQR"
		var col int = strings.Index("ABCDEF", string(Route[0]))
		f.SetCellInt("Sheet1", string(cols[col])+srowx, 1)
		cols = "STUVW"
		for col = 0; col < len(tshirts); col++ {
			if tshirts[col] > 0 {
				f.SetCellInt("Sheet1", string(cols[col])+srowx, tshirts[col])
			}
		}
		if Patches[0] == '1' {
			f.SetCellInt("Sheet1", "X"+srowx, 1)
		} else if Patches[0] == '2' {
			f.SetCellInt("Sheet2", "X"+srowx, 2)
		}

		srow++
	}

	err = f.SetCellStyle("Sheet1", "B1", "D1", styleH)
	err = f.SetCellStyle("Sheet1", "K1", "X1", styleH)
	err = f.SetCellStyle("Sheet1", "A2", "D"+srowx, styleV)
	err = f.SetCellStyle("Sheet2", "A2", "A"+srowx, styleV)
	err = f.SetCellStyle("Sheet1", "K2", "X"+srowx, styleV)

	srow++ // Leave a gap before totals

	// L to X
	for _, c := range "LMNOPQRSTUVWX" {
		f.SetCellFormula("Sheet1", string(c)+strconv.Itoa(srow), "=sum("+string(c)+"2:"+string(c)+srowx)
		f.SetCellStyle("Sheet1", string(c)+strconv.Itoa(srow), string(c)+strconv.Itoa(srow), styleT)
	}
	f.SetConditionalFormat("Sheet1", "L"+strconv.Itoa(srow)+"X"+strconv.Itoa(srow), `{"no_blanks":}`)

	f.SetActiveSheet(0)
	f.SetCellValue("Sheet1", "A1", "No.")
	f.SetCellValue("Sheet2", "A1", "No.")
	f.SetColWidth("Sheet1", "A", "A", 5)
	f.SetColWidth("Sheet2", "A", "A", 5)

	f.SetCellValue("Sheet1", "B1", "Registered")
	f.SetCellValue("Sheet1", "C1", "Started")
	f.SetCellValue("Sheet1", "D1", "Finished")
	f.SetColWidth("Sheet1", "B", "D", 3)

	f.SetCellValue("Sheet1", "E1", "Rider(first)")
	f.SetCellValue("Sheet1", "F1", "Rider(last)")
	f.SetColWidth("Sheet1", "E", "F", 15)

	f.SetCellValue("Sheet2", "B1", "Rider(first)")
	f.SetCellValue("Sheet2", "C1", "Rider(last)")
	f.SetColWidth("Sheet2", "B", "C", 15)

	f.SetCellValue("Sheet2", "D1", "Mobile")
	f.SetCellValue("Sheet2", "E1", "NOK name")
	f.SetCellValue("Sheet2", "F1", "Relationship")
	f.SetCellValue("Sheet2", "G1", "Contact number")
	f.SetColWidth("Sheet2", "D", "G", 20)

	f.SetCellValue("Sheet3", "A1", "Make")
	f.SetCellValue("Sheet3", "B1", "Number")
	f.SetColWidth("Sheet3", "A", "A", 12)

	f.SetCellValue("Sheet1", "G1", "IBA #")
	f.SetCellValue("Sheet1", "H1", "Pillion")
	f.SetColWidth("Sheet1", "H", "H", 20)

	f.SetCellValue("Sheet1", "I1", "Make")
	f.SetColWidth("Sheet1", "I", "I", 12)
	f.SetCellValue("Sheet1", "J1", "Model")
	f.SetColWidth("Sheet1", "J", "J", 25)
	f.SetCellValue("Sheet1", "K1", "Miles to Squires")
	f.SetColWidth("Sheet1", "K", "K", 5)

	f.SetCellValue("Sheet1", "L1", "Camping")
	f.SetColWidth("Sheet1", "L", "X", 4)

	f.SetCellValue("Sheet1", "M1", "A-NC")
	f.SetCellValue("Sheet1", "N1", "B-NAC")
	f.SetCellValue("Sheet1", "O1", "C-SC")
	f.SetCellValue("Sheet1", "P1", "D-SAC")
	f.SetCellValue("Sheet1", "Q1", "E-500C")
	f.SetCellValue("Sheet1", "R1", "F-500AC")

	f.SetCellValue("Sheet1", "S1", "T-shirt S")
	f.SetCellValue("Sheet1", "T1", "T-shirt M")
	f.SetCellValue("Sheet1", "U1", "T-shirt L")
	f.SetCellValue("Sheet1", "V1", "T-shirt XL")
	f.SetCellValue("Sheet1", "W1", "T-shirt XXL")

	f.SetCellValue("Sheet1", "X1", "Patches")

	f.SetRowHeight("Sheet1", 1, 100)
	f.SetRowHeight("Sheet2", 1, 20)
	f.SetRowHeight("Sheet3", 1, 20)

	sort.Slice(bikes, func(i, j int) bool { return bikes[i].make < bikes[j].make })
	//fmt.Printf("%v\n", bikes)
	srow = 2
	ntot := 0
	for i := 0; i < len(bikes); i++ {
		f.SetCellValue("Sheet3", "A"+strconv.Itoa(srow), bikes[i].make)
		f.SetCellInt("Sheet3", "B"+strconv.Itoa(srow), bikes[i].num)
		f.SetCellStyle("Sheet3", "B"+strconv.Itoa(srow), "B"+strconv.Itoa(srow), styleV)
		ntot += bikes[i].num
		srow++
	}

	srow++
	f.SetCellInt("Sheet3", "B"+strconv.Itoa(srow), ntot)
	err = f.SetCellStyle("Sheet3", "B"+strconv.Itoa(srow), "B"+strconv.Itoa(srow), styleT)
	if err != nil {
		log.Fatal(err)
	}

	f.SetSheetName("Sheet1", "Registration sheets")
	f.SetSheetName("Sheet2", "NOK sheets")
	f.SetSheetName("Sheet3", "Bikes")
	// Save spreadsheet by the given path.
	if err := f.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
	}
}
