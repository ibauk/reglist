package main

import (
	"reflect"
)

type Bikemake struct {
	Make string
	Num  int
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
}

func NewTotals(numRoutes, numSizes, numBikes int) *Totals {

	var t Totals
	t.NumTshirtsBySize = make([]int, numSizes)
	t.NumRidersByRoute = make([]int, numRoutes)
	t.Bikes = make([]Bikemake, numBikes)
	t.LoMiles2Squires = 9999
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
