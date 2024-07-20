package main

import (
	"strconv"
	"strings"
)

type Person = struct {
	First        string
	Last         string
	IBA          string
	HasIBANumber bool
	RBL          string
	Email        string
	Phone        string
	Address      string
	Address1     string
	Address2     string
	Town         string
	County       string
	Postcode     string
	Country      string
}

type Money = struct {
	EntryDonation string
	SquiresCheque string
	SquiresCash   string
	RBLRAccount   string
	JustGivingAmt string
	JustGivingURL string
}

type EntrantRBLR = struct {
	EntrantID            int
	EntrantStatus        int
	Rider                Person
	Pillion              Person
	NokName              string
	NokRelation          string
	NokPhone             string
	Bike                 string
	BikeReg              string
	Route                string
	OdoStart             string
	OdoFinish            string
	OdoCounts            string
	StartTime            string
	FinishTime           string
	FundsRaised          Money
	FreeCamping          string
	CertificateAvailable string
	CertificateDelivered string
	Tshirt1              string
	Tshirt2              string
	Patches              int
	EditMode             string
}

func BuildRBLR(e Entrant) EntrantRBLR {

	const DNS = 0
	var Routes = map[byte]string{'A': "NCW", 'B': "NAC", 'C': "SCW", 'D': "SAC", 'E': "500CW", 'F': "500AC"}
	var E EntrantRBLR

	E.EntrantID, _ = strconv.Atoi(e.Entrantid)
	E.EntrantStatus = DNS
	E.Rider.First = e.RiderFirst
	E.Rider.Last = e.RiderLast
	E.Rider.IBA = e.RiderIBA
	E.Rider.HasIBANumber = E.Rider.IBA != ""
	E.Rider.RBL = e.RiderRBL
	E.Rider.Email = e.Email
	E.Rider.Phone = e.Phone
	E.Rider.Address1 = e.Address1
	E.Rider.Address2 = e.Address2
	E.Rider.Town = e.Town
	E.Rider.County = e.County
	E.Rider.Postcode = e.Postcode
	E.Rider.Country = e.Country

	E.Pillion.First = e.PillionFirst
	E.Pillion.Last = e.PillionLast
	E.Pillion.IBA = e.PillionIBA
	E.Pillion.HasIBANumber = E.Pillion.IBA != ""
	E.Pillion.RBL = e.PillionRBL
	E.Pillion.Email = e.PEmail
	E.Pillion.Phone = e.PPhone
	E.Pillion.Address1 = e.PAddress1
	E.Pillion.Address2 = e.PAddress2
	E.Pillion.Town = e.PTown
	E.Pillion.County = e.PCounty
	E.Pillion.Postcode = e.PPostcode
	E.Pillion.Country = e.PCountry

	E.NokName = e.NokName
	E.NokRelation = e.NokRelation
	E.NokPhone = e.NokPhone

	E.Bike = e.Bike
	E.BikeReg = e.BikeReg

	E.Route = Routes[e.RouteClass[0]]
	E.FundsRaised.EntryDonation = e.Sponsorship

	E.OdoCounts = e.OdoKms

	E.FreeCamping = e.Camping
	E.CertificateAvailable = "Y"
	E.CertificateDelivered = "N"
	E.Tshirt1 = e.Tshirt1
	E.Tshirt2 = e.Tshirt2
	E.Patches, _ = strconv.Atoi(e.Patches)

	//E.FundsRaised.EntryDonation=e
	return E
}

func rblrPersonFieldNames(P string, Flds []string) string {

	var res []string

	for _, x := range Flds {
		res = append(res, P+x)
	}
	return strings.Join(res, ",")
}

func q(x string) string {
	return `'` + strings.ReplaceAll(x, `'`, `''`) + `'`
}
func writeRBLR(e *EntrantRBLR) {

	if rblrdb == nil {
		return
	}
	var PersonFields = []string{`First`, `Last`, `Address1`, `Address2`, `Town`, `County`, `Postcode`, `Country`, `IBA`, `RBL`, `Phone`, `Email`}
	var Fieldnames = `EntrantID,Bike,BikeReg,` + rblrPersonFieldNames("Rider", PersonFields) + `,` + rblrPersonFieldNames("Pillion", PersonFields)
	Fieldnames += `,NokName,NokRelation,NokPhone,Route,OdoCounts,EntryDonation,FreeCamping,CertificateAvailable,Tshirt1,Tshirt2,Patches`

	sqlx := "INSERT INTO entrants(" + Fieldnames + ") VALUES("
	sqlx += strconv.Itoa(e.EntrantID)
	sqlx += `,` + q(e.Bike) + `,` + q(e.BikeReg)
	sqlx += `,` + q(e.Rider.First) + `,` + q(e.Rider.Last)
	sqlx += `,` + q(e.Rider.Address1) + `,` + q(e.Rider.Address2)
	sqlx += `,` + q(e.Rider.Town) + `,` + q(e.Rider.County)
	sqlx += `,` + q(e.Rider.Postcode) + `,` + q(e.Rider.Country)
	sqlx += `,` + q(e.Rider.IBA) + `,` + q(e.Rider.RBL)
	sqlx += `,` + q(e.Rider.Phone) + `,` + q(e.Rider.Email)
	sqlx += `,` + q(e.Pillion.First) + `,` + q(e.Pillion.Last)
	sqlx += `,` + q(e.Pillion.Address1) + `,` + q(e.Pillion.Address2)
	sqlx += `,` + q(e.Pillion.Town) + `,` + q(e.Pillion.County)
	sqlx += `,` + q(e.Pillion.Postcode) + `,` + q(e.Pillion.Country)
	sqlx += `,` + q(e.Pillion.IBA) + `,` + q(e.Pillion.RBL)
	sqlx += `,` + q(e.Pillion.Phone) + `,` + q(e.Rider.Email)

	sqlx += `,` + q(e.NokName) + `,` + q(e.NokRelation) + `,` + q(e.NokPhone)
	sqlx += `,` + q(e.Route)
	sqlx += `,` + q(e.OdoCounts)
	sqlx += `,` + q(e.FundsRaised.EntryDonation)
	sqlx += `,` + q(e.FreeCamping) + `,` + q(e.CertificateAvailable)
	sqlx += `,` + q(e.Tshirt1) + `,` + q(e.Tshirt2) + `,` + q(strconv.Itoa(e.Patches))
	sqlx += `)
	`

	//	fmt.Println(sqlx)

	_, err := rblrdb.Exec(sqlx)
	checkerr(err)

}
