package main

import (
	"fmt"
	"log"
	"reflect"
	"time"
)

type Bikemake struct {
	Make string
	Num  int
}

type Entrystats struct {
	Month     string
	Total     int
	NumIBA    int
	NumNovice int
}

type Totals struct {
	NumRiders          int
	NumPillions        int
	NumNovices         int
	NumIBAMembers      int
	NumCamping         int
	NumPatches         int
	NumTshirts         int
	NumTshirtsBySize   []int
	NumRidersByRoute   []int
	NumMiles2Squires   int
	LoMiles2Squires    int
	HiMiles2Squires    int
	TotMoneyOnDay      int // Sponsor money received on day
	TotMoneySponsor    int // Sponsor money paid up front
	TotMoneyMainPaypal int // Original Paypal payment
	TotMoneyCashPaypal int // Subsequent Paypal payments
	Bikes              []Bikemake
	EntriesByPeriod    []Entrystats
	CancelledRows      []int
}

func NewTotals(numRoutes, numSizes, numBikes int) *Totals {

	var t Totals
	t.NumTshirtsBySize = make([]int, numSizes)
	t.NumRidersByRoute = make([]int, numRoutes)
	t.Bikes = make([]Bikemake, numBikes)
	t.EntriesByPeriod = make([]Entrystats, 0)
	t.LoMiles2Squires = 9999
	t.CancelledRows = make([]int, 0)
	return &t
}

type Entrant struct {
	Entrantid        string
	RiderFirst       string
	RiderLast        string
	RiderIBA         string
	RiderNovice      string
	PillionFirst     string
	PillionLast      string
	PillionIBA       string
	PillionNovice    string
	Bike             string
	BikeMake         string
	BikeModel        string
	BikeReg          string
	OdoKms           string
	Email            string
	Phone            string
	Address1         string
	Address2         string
	Town             string
	County           string
	Postcode         string
	Country          string
	NokName          string
	NokPhone         string
	NokRelation      string
	BonusClaimMethod string
	RouteClass       string
	Tshirt1          string
	Tshirt2          string
	Patches          string
	Camping          string
	Miles2Squires    string
	EnteredDate      string
}

func EntrantHeaders() []string {

	var e Entrant
	te := reflect.TypeOf(e)
	var res []string
	for i := 0; i < te.NumField(); i++ {
		res = append(res, te.Field(i).Name)
	}
	return res
}

func Entrant2Strings(e Entrant) []string {

	te := reflect.ValueOf(e)
	var res []string
	for i := 0; i < te.NumField(); i++ {
		res = append(res, te.Field(i).String())
	}
	return res

}

func ReportingPeriod(isodate string) string {
	t, _ := time.Parse("2006-01-02 15:04:05", isodate)
	if cfg.ReportWeekly {
		for t.Weekday() != time.Monday && t.Day() > 1 {
			t = t.AddDate(0, 0, -1)
		}
		return t.Format("01-02")
	}
	return t.Format("01-")
}

func lookupIBA(first, last string) (string, string) {

	var num, email string

	sqlx := "SELECT IBA_Number, Email FROM rd.riders WHERE Rider_Name = '" + first + "' || ' ' || '" + last + "' AND IBA_Number <>'' COLLATE NOCASE"
	//fmt.Printf("%v\n", sqlx)
	rows, err := db.Query(sqlx)
	if err != nil {
		log.Fatal(err)
	}
	defer rows.Close()
	if rows.Next() {
		rows.Scan(&num, &email)
		return num, email
	}
	return "", ""
}

func LookupIBANumbers(e *Entrant) {

	if e.RiderIBA == "" {
		riba, remail := lookupIBA(e.RiderFirst, e.RiderLast)
		if riba != "" {
			fmt.Printf("Rider %v %v (%v) is IBA %v (%v)\n", e.RiderFirst, e.RiderLast, e.Email, riba, remail)
			e.RiderIBA = riba
		}
	}
	if e.PillionFirst != "" && e.PillionLast != "" && e.PillionIBA == "" {
		piba, pemail := lookupIBA(e.PillionFirst, e.PillionLast)
		if piba != "" {
			fmt.Printf("Pillion %v %v (%v) is IBA %v (%v)\n", e.PillionFirst, e.PillionLast, "", piba, pemail)
			e.PillionIBA = piba
		}
	}
}
