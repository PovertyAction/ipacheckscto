# CHECK SurveyCTO XLS Forms

## Overview

ipacheckscto is a Stata Module for checking SurveyCTO XLS forms for possible errors and best practices. Please note that this program does not check for issues that are already included in the SurveyCTO server debug program. Please ensure that have succesfully uploaded your form onto the SurveyCTO server before running this program.


## Installation

```stata
* ipacheckscto can be installed from github

net install ipacheckscto, all replace ///
	from("https://raw.githubusercontent.com/PovertyAction/ipacheckscto/master/ado")
```

## Syntax
```stata
ipacheckscto using filename[, outfile(string) other(integer) dontknow(integer) refuse(integer) replace]

options
	outfile 		- Output results to excel filename
	other			- Specify integer value for other specify
	dontknow		- Specify integer value for don't know
	refuse			- Specify integer value for refuse to answer
	replace			- Overwrite existing outfile

```

## Example Syntax
```stata
* Check XLS form without exporting issues
ipacheckscto using "Baseline Survey.xlsx"

* Check XLS form and export issues to excel
ipacheckscto using "Baseline Survey.xlsx", outfile("basline_survey_issues.xlsx") replace

```

Please report all bugs/feature request to the [github issues page](https://github.com/iabaako/ipacheckscto/issues)
