# Rally info spreadsheet

Entrant data all begins at the point when they sign up for the event using the relevant Wufoo form. All those entries are checked and corrected where necessary by either Phil or Bob so that the Wufoo dataset always holds the latest and most accurate information.  This spreadsheet extracts that data automagically directly from the Wufoo dataset and reformats it ready for use either as a source of statistics or as forms suitable for registration, check-out, check-in and merchandising processes. A separate automatic process produces fully populated registration/disclaimer forms as well as finisher certificates.

The front "Overview" tab is intended as a quick check page, with access to most information in one place.

The "Money" tab has live cells (not in the 'safe' version) for input of amounts received at registration reflected in totals and on the Stats tab. The '!!!' column contains code to self-check the sheet ensuring that it adds up correctly and highlights unpaid amounts.

The "Checkouts" tab is intended for "carpark check-out, check-in" use while the "Registration" tab provides a more comprehensive checklist.

## Commandline arguments

-cfg *cfgname*
>The default is "rblr" which uses the file **rblr.yml**. The ".yml" is appended to *cfgname* so specify "bbr", "bbl", etc

-csv *filename*
>Full path of the input .CSV file containing entrant data. The default is **entrants.csv**

-nocsv
>Don't import a .CSV file, just use the existing contents of the SQLite database

-rpt
>The .CSV file was produced by a Wufoo report as opposed to the default format exported when logged in as administrator. This switch actually chooses the **rfields** entry in the configuration rather than the **afields**. For some reason in their infinite wisdom Wufoo see fit to export the metadata fields at the end of each record in report extracts rather than at the beginning for admin downloads.

-safe
>Produce a spreadsheet with values only, no formulas. The default is for totals, etc to be live formulas so the sheet can be updated by hand

-sql *filename*
>The full path to the SQLite database file used by the process. The default is **entrantdata.db**

-xls *filename*
>The full path for the resultant spreadsheet. The default is **reglist.xlsx**

