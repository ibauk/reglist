package main

import (
	"database/sql"
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

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
		sqlx = sqlx_rblr
	} else {
		fmt.Printf("Running in rally mode, %v\n", sm)
		sqlx = sqlx_rally
	}
	sqlx = "SELECT " + sqlx + " FROM entrants"
	if len(cfg.PaymentStatus) != 0 {
		sqlx += " WHERE PaymentStatus IN ('" + strings.Join(cfg.PaymentStatus, "','") + "')"
	}
	sqlx += " ORDER BY " + cfg.EntrantOrder

	if *expReport == "" {
		*expReport = cfg.Rally + cfg.Year
	}
	if filepath.Ext(*expReport) == "" {
		*expReport = *expReport + ".csv"
	}

	//fmt.Printf("Creating %v\n", *xlsName)

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

	if cfg.RBLRDB != "" {
		rblrdb, err = sql.Open("sqlite3", cfg.RBLRDB)
		checkerr(err)
		fmt.Println("RBLR database " + cfg.RBLRDB + " is opened")
		sqlx := "DELETE FROM entrants"
		rblrdb.Exec(sqlx)
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

func initStyles() {

	// styleCancel for highlighting cancelled entrants
	styleCancel, _ = xl.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{Vertical: "center", Horizontal: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"edeb57"}, Pattern: 1},
	})

	// Totals
	styleT, _ = xl.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal:      "center",
			Indent:          1,
			JustifyLastLine: true,
			ReadingOrder:    0,
			RelativeIndent:  1,
			ShrinkToFit:     true,
			TextRotation:    0,
			Vertical:        "",
			WrapText:        true,
		},
		Font: &excelize.Font{
			Bold:   true,
			Italic: false,
			Family: "Arial",
			Size:   12,
			Color:  "000000",
		},
	})

	// Header, vertical
	styleH, _ = xl.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal:      "center",
			Indent:          1,
			JustifyLastLine: true,
			ReadingOrder:    0,
			RelativeIndent:  1,
			ShrinkToFit:     true,
			TextRotation:    90,
			Vertical:        "",
			WrapText:        true,
		},
		Fill: excelize.Fill{Type: "pattern", Color: []string{"#dddddd"}, Pattern: 1},
	})

	// Header, horizontal
	styleH2, _ = xl.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal:      "center",
			Indent:          1,
			JustifyLastLine: true,
			ReadingOrder:    0,
			RelativeIndent:  1,
			ShrinkToFit:     true,
			TextRotation:    0,
			Vertical:        "center",
			WrapText:        true,
		},
		Fill: excelize.Fill{Type: "pattern", Color: []string{"#dddddd"}, Pattern: 1},
	})

	styleH2L, _ = xl.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal:      "left",
			Indent:          1,
			JustifyLastLine: true,
			ReadingOrder:    0,
			RelativeIndent:  1,
			ShrinkToFit:     true,
			TextRotation:    0,
			Vertical:        "center",
			WrapText:        true,
		},
		Fill: excelize.Fill{Type: "pattern", Color: []string{"#dddddd"}, Pattern: 1},
	})

	// Data
	styleV, _ = xl.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal:      "center",
			Indent:          1,
			JustifyLastLine: true,
			ReadingOrder:    0,
			RelativeIndent:  1,
			ShrinkToFit:     true,
			TextRotation:    0,
			Vertical:        "",
			WrapText:        true,
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})

	// Open data
	styleV2, _ = xl.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal:      "center",
			Indent:          1,
			JustifyLastLine: true,
			ReadingOrder:    0,
			RelativeIndent:  1,
			ShrinkToFit:     true,
			TextRotation:    0,
			Vertical:        "",
			WrapText:        true,
		},
		Border: []excelize.Border{
			{
				Type:  "bottom",
				Color: "000000",
				Style: 1,
			}},
	})

	styleV2L, _ = xl.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal:      "left",
			Indent:          1,
			JustifyLastLine: true,
			ReadingOrder:    0,
			RelativeIndent:  1,
			ShrinkToFit:     true,
			TextRotation:    0,
			Vertical:        "",
			WrapText:        false,
		},
		Border: []excelize.Border{
			{
				Type:  "bottom",
				Color: "000000",
				Style: 1,
			}},
	})

	styleV2LBig, _ = xl.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal:      "left",
			Indent:          1,
			JustifyLastLine: true,
			ReadingOrder:    0,
			RelativeIndent:  1,
			ShrinkToFit:     true,
			TextRotation:    0,
			Vertical:        "",
			WrapText:        false,
		},
		Border: []excelize.Border{
			{
				Type:  "bottom",
				Color: "000000",
				Style: 1,
			}},
		Font: &excelize.Font{
			Size: 16,
		},
	})

	styleV3, _ = xl.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal:      "center",
			Indent:          1,
			JustifyLastLine: true,
			ReadingOrder:    0,
			RelativeIndent:  1,
			ShrinkToFit:     true,
			TextRotation:    0,
			Vertical:        "",
			WrapText:        false,
		},
	})

	// styleW for highlighting, particularly errorneous, cells
	styleW, _ = xl.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal:      "center",
			Indent:          1,
			JustifyLastLine: true,
			ReadingOrder:    0,
			RelativeIndent:  1,
			ShrinkToFit:     true,
			TextRotation:    0,
			Vertical:        "",
			WrapText:        false,
		},
		Fill: excelize.Fill{Type: "pattern", Color: []string{"#ffff00"}, Pattern: 1}})

	styleRJ, _ = xl.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal:      "right",
			Indent:          1,
			JustifyLastLine: true,
			ReadingOrder:    0,
			RelativeIndent:  1,
			ShrinkToFit:     true,
			TextRotation:    0,
			Vertical:        "",
			WrapText:        false,
		},
	})

	styleRJSmall, _ = xl.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal:      "center",
			Indent:          1,
			JustifyLastLine: true,
			ReadingOrder:    0,
			RelativeIndent:  1,
			ShrinkToFit:     true,
			TextRotation:    0,
			Vertical:        "",
			WrapText:        false,
		},
		Border: []excelize.Border{
			{
				Type:  "bottom",
				Color: "000000",
				Style: 1,
			}},

		Font: &excelize.Font{
			Size: 8,
		},
	})

	xl.SetDefaultFont("Arial")

}
