# Reglist - Rally info spreadsheet

This creates a spreadsheet workbook (.xlsx) of entrant data ,derived from the original Wufoo form data, used to manage the RBLR1000 and other IBAUK rallies.

The book is used to keep track of registrations, T-shirt and patch allocations, sponsorship monies, etc.

Entrant data all begins at the point when they sign up for the event using the relevant Wufoo form. All those entries are checked and corrected where necessary by either Phil or Bob so that the Wufoo dataset always holds the latest and most accurate information.  This spreadsheet extracts that data automagically directly from the Wufoo dataset and reformats it ready for use either as a source of statistics or as forms suitable for registration, check-out, check-in and merchandising processes. A separate automatic process produces fully populated registration/disclaimer forms as well as finisher certificates.

## Safe / live versions
The workbook can be generated as either a "safe" version, containing values only with no formulas, or as a live version using formulas to keep track of any changes made to the data. The advantage of the safe version is that it can be viewed in a variety of environments without fear of tripping local security measures.

## Workbook pages

### Overview tab
This front page is intended as a quick check page, with access to most information in one place.

### Registration tab
May be used as a physical registration log with columns to tick off key pieces of information.

### NOK list
Holds contact details for entrants and who should be contacted in the event of accidents etc.

### Shop tab
This is only present if merchandise such as T-shirts and patches is offered, whether or not such items are chargeable.

### Money tab
The "Money" tab has live cells (not in the 'safe' version) for input of amounts received at registration reflected in totals and on the Stats tab. The '!!!' column contains code to self-check the sheet ensuring that it adds up correctly and highlights unpaid amounts.

### Stats tab
Presents simple statistics relating to various aspects of the event.

### Carpark tab
Intended for "carpark check-out, check-in" use while the *Registration* and *NOK list* tabs provide a more comprehensive checklist.

## Commandline arguments
Reglist is run from a shell (terminal or cmd) prompt (commandline) and its operation is controlled by several arguments or parameters as below:-

**-cfg** *cfgname*
>The default is "rblr" which uses the file **rblr.yml** in the current folder. The ".yml" is appended to *cfgname* so specify "bbr", "bbl", etc

**-csv** *filename*
>Full path of the input .CSV file containing entrant data. The default is **entrants.csv** in the current folder.

**-exp** *filename*
>Full path of a .CSV file to be created as input to, *inter alia*, the ScoreMaster rally administration software. This file is in a format standard across all IBAUK events and reflecting any renumbering or data cleansing carried out by Reglist.

**-nocsv**
>Don't import a .CSV file, just reuse the existing contents of the intermediate SQLite database

**-rpt**
>The .CSV file was produced by a Wufoo report as opposed to the default format exported when logged in as administrator. This switch actually chooses the **rfields** entry in the configuration rather than the **afields**. For some reason in their infinite wisdom Wufoo see fit to export the metadata fields at the end of each record in report extracts rather than at the beginning for admin downloads.

**-safe**
>Produce a spreadsheet with values only, no formulas. The default is for totals, etc to be live formulas so the sheet can be updated by hand

**-sql** *filename*
>The full path to the SQLite database file used by the process. The default is **entrantdata.db** in the current folder.

**-xls** *filename*
>The full path for the resultant spreadsheet. The default is **reglist.xlsx** in the current folder.


## Configuration files
Further fine control over the output is achieved by the use of configuration files, one for each rally covered. The files are in standard [YAML](https://yaml.org/) format with contents as below:-

**name:** *rallyname*
>The short name of the rally, used only for internal purposes. The reserved name **rblr** (lowercase) triggers the distinction between the RBLR1000 and other events. Any name can be used for other rallies but it's generally a good idea to use a meaningful name.

**year:** *year*
>Used purely for identification purposes.

**afields:** / **rfields**
>These hold arrays of fieldnames reflecting the order of input fields in the incoming .CSV. The fieldnames are the names used within Reglist and may differ from those used in the .CSV, only the order matters, not the names. **afields** refers to the files downloaded from the form administration facility of Wufoo and **rfields** refers to the files exported via the corresponding Wufoo report. The two files contain the same information but in their infinite wisdom Wufoo have seen fit to place the metadata before the data in one and after in the other.

**tshirtsizes:** 
>An array of sizes available.

**tshirtcost:** *integer*
>The cost of a single shirt, pounds only, we can't be doing with penny pinchers.

**riderfee:** / **pillionfee:** *integer*
>Entry fee.

**patchavail:** true/false
>Is there a patch available for this event?

**patchcost:** *integer*
>The cost of a single patch, pounds only, we can't be doing with penny pinchers.

**sponsorship:** true/false
>Whether we're collecting sponsorship monies through Wufoo in addition to entry fees.

**fundsonday:** *title*
>If we're accounting for sponsorship this sets the heading for the live column collecting funds on the day. For the RBLR1000 this would be "Cheque @ Squires" or similar.

**novice:** *label*
>Sometimes newbies are called 'novice', sometimes 'rookie'. This provides endless possibilities for amusement.

**add2entrantid:** *integer*
>This provides a crude facility to adjust entrant numbers from Wufoo in cases where, for example, test entries were entered on a live form so that the numbers don't start from 1. If the first record is actually 3 but we want to start at 1, seeting this value to -2 will achieve that.
>The reserved field *RiderNumber* in the incoming .CSV can also be used to override entrant numbers.

**entrantorder:** *fieldlist*
>Entrants will be listed on the spreadsheet in this order. *fieldlist* may contain SQL functions including **upper**, **lower**, etc and may also specify **DESC** to reverse the order.

