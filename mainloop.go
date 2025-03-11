package main

import (
	"fmt"
	"log"
	"regexp"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func mainloop() {

	//fmt.Println(sqlx)
	rows1, err1 := db.Query(sqlx)
	if err1 != nil {
		log.Fatal(err1)
	}
	totx.srow = 2 // First spreadsheet row to populate

	var tshirts [max_tshirt_sizes]int

	if rblrdb != nil {
		rblrdb.Exec("BEGIN")
	}

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
			tot.NumWithdrawn++
			if *verbose {
				fmt.Printf("    Rider %v %v [#%v] is withdrawn\n", e.RiderFirst, e.RiderLast, e.Entrantid)
			}
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
		e.Postcode = strings.ToUpper(e.Postcode)

		e.NokName = properName(NokName)
		e.NokPhone = NokNumber
		e.NokRelation = properName(NokRelation)

		e.RouteClass = Route
		e.Tshirt1 = T1
		e.Tshirt2 = T2
		e.Patches = Patches
		e.Camping = fmtCampingYN(freecamping)
		e.Miles2Squires = strconv.Itoa(intval(miles2squires))
		e.Bike = Bike

		ss := intval(Sponsor)
		if ss > 0 {
			e.Sponsorship = strconv.Itoa(ss)
		}

		if !*noLookup {
			LookupIBANumbers(&e)
		}

		if *rally == "rblr" {
			rblre := BuildRBLR(e)
			writeRBLR(&rblre)
		}

		RiderFirst = properName(e.RiderFirst)
		RiderLast = properName(e.RiderLast)
		PillionFirst = properName(e.PillionFirst)
		PillionLast = properName(e.PillionLast)

		if isFOC && *verbose {
			fmt.Printf("    Rider %v %v [#%v] has Paid=%v and is therefore FOC\n", e.RiderFirst, e.RiderLast, e.Entrantid, Paid)
		}
		if isCancelled && *verbose {
			fmt.Printf("    Rider %v %v [#%v] has Paid=%v\n", e.RiderFirst, e.RiderLast, e.Entrantid, Paid)
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
			if cfg.Rally == "rblr" {
				xl.SetCellInt(subssheet, "A"+totx.srowx, entrantid)
			}
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
			if cfg.Rally == "rblr" {
				xl.SetCellValue(subssheet, "B"+totx.srowx, RiderFirst)
				xl.SetCellValue(subssheet, "C"+totx.srowx, RiderLast)
			}
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

		Sponsorship := 0 /* cancelledFees - PW 2024-07-18 */

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
						if cfg.Rally == "rblr" {
							xl.SetCellInt(subssheet, "D"+totx.srowx, Sponsorship)
						}
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
			if Paid == "Unpaid" && true {
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

	if rblrdb != nil {
		rblrdb.Exec("COMMIT")
	}

}
