package main

import (
	"testing"
)

// This handles any testing flags that would otherwise be passed
// through to the main program and cause problems
var _ = func() bool {
	testing.Init()
	return true
}()

func TestIntval(t *testing.T) {
	tables := []struct {
		x string
		n int
	}{
		{"30", 30},
		{"aa30", 30},
		{"30bbb", 30},
		{"'30", 30},
		{"£30", 30},
		{"-30", -30},
		{"30-", -30},
		{"-30-", -30},
		{"30.5", 30},
	}
	for _, table := range tables {
		n := intval(table.x)
		if n != table.n {
			t.Errorf("%v gives %v", table.x, n)
		}
	}
}
func TestMakeModel(t *testing.T) {
	tables := []struct {
		bk string
		mk string
		md string
	}{
		{"Triumph Herald", "Triumph", "Herald"},
		{"'98 Harley", "Harley", ""},
		{"2005 Suzuki Bandit", "Suzuki", "Bandit"},
	}

	for _, table := range tables {
		mk, md := extractMakeModel(table.bk)
		if mk != table.mk || md != table.md {
			t.Errorf("%v doesn't split right", table.bk)
		}
	}

}

func TestProperBike(t *testing.T) {
	tables := []struct {
		ip string
		op string
	}{
		{"triumph", "Triumph"},
		{"TRIUMPH", "Triumph"},
		{"tRIUMPH", "Triumph"},
		{"tRiUmPH", "Triumph"},
	}
	for _, table := range tables {
		x := properBike(table.ip)
		if x != table.op {
			t.Errorf("%v yields %v", table.ip, x)
		}
	}
}

func TestProperName(t *testing.T) {
	tables := []struct {
		ip string
		op string
	}{
		{"bob stammers", "Bob Stammers"},
		{"colin mccrea", "Colin McCrea"},
		{"john o'keefe", "John O'Keefe"},
	}
	for _, table := range tables {
		x := properName(table.ip)
		if x != table.op {
			t.Errorf("%v yields %v", table.ip, x)
		}
	}
}
