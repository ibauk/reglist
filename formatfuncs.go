package main

import (
	"regexp"
	"strings"

	"golang.org/x/text/cases"
	"golang.org/x/text/language"
)

// boolean (Y/N) novice or not. Used to set flag in record. Contrast with fmtNoviceYb below
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

func fmtCampingYN(x string) string {

	if x == cfg.FreeCamping && cfg.FreeCamping != "" {
		return "Y"
	}
	return ""
}

func fmtIBA(x string) string {

	if x == "-1" {
		return "n/a"
	}
	return strings.ReplaceAll(x, ".0", "")

}

func fmtNoviceYb(noviceYN string) string {

	if noviceYN == "Y" {
		return "Yes"
	}
	return ""
}

func fmtOdoKM(x string) string {

	y := strings.ToUpper(x)
	if len(y) > 0 && y[0] == 'K' {
		return "K"
	}
	return "M"

}

func fmtRBL(x string) string {

	if x == cfg.LegionMember && cfg.LegionMember != "" {
		return "L"
	} else if x == cfg.LegionRider && cfg.LegionRider != "" {
		return "R"
	}
	return ""
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

func ShortMaker(x string) string {

	p := strings.Index(x, "-")
	if p < 0 {
		return x
	}
	return x[0:p]
}

func stringsTitle(x string) string {

	caser := cases.Title(language.English)
	return caser.String(x)

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
