package main

import (
	"os"

	yaml "gopkg.in/yaml.v2"
)

// Words holds the lists of words used to fix capitalisation
type Words struct {
	Propernames  bool     `yaml:"propercasenames"`
	Specialnames []string `yaml:"specialnames"`
	Bikewords    []string `yaml:"bikewords"`
}

// NewWords returns the word lists
func NewWords() (*Words, error) {

	configPath := "reglist.yml"
	words := &Words{}

	if _, err := os.Stat(configPath); os.IsNotExist(err) {
		return words, err // Empty so no cleansing will happen
	}

	file, err := os.Open(configPath)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	// Init new YAML decode
	d := yaml.NewDecoder(file)

	// Start YAML decoding from file
	if err := d.Decode(&words); err != nil {
		return nil, err
	}

	return words, nil
}

// Config holds the contents of the configuration file
type Config struct {
	Rally         string   `yaml:"name"`
	Year          string   `yaml:"year"`
	Afields       []string `yaml:"afields"`
	Rfields       []string `yaml:"rfields"`
	Tshirts       []string `yaml:"tshirtsizes"`
	Tshirtcost    int      `yaml:"tshirtcost"`
	Riderfee      int      `yaml:"riderfee"`
	Pillionfee    int      `yaml:"pillionfee"`
	Patchavail    bool     `yaml:"patchavail"`
	Patchcost     int      `yaml:"patchcost"`
	Sponsorship   bool     `yaml:"sponsorship"`
	Fundsonday    string   `yaml:"fundsonday"`
	Novice        string   `yaml:"novice"`
	Add2entrantid int      `yaml:"add2entrantid"`
	EntrantOrder  string   `yaml:"entrantorder"`
}

// NewConfig returns a new decoded Config struct
func NewConfig(configPath string) (*Config, error) {
	// Create config structure
	config := &Config{}
	config.Rally = "test"
	config.Novice = "novice"
	config.EntrantOrder = "upper(RiderLast),upper(RiderName)"

	// Open config file
	file, err := os.Open(configPath)
	if err != nil {
		return config, err
	}
	defer file.Close()

	// Init new YAML decode
	d := yaml.NewDecoder(file)

	// Start YAML decoding from file
	if err := d.Decode(&config); err != nil {
		return config, err
	}

	return config, nil
}
