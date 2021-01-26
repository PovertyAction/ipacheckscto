*! version 1.1
*! Ishmail Azindoo Baako (IPA) February, 2018

* Stata program to review SurveyCTO xls form and export issues to an excel sheet
* This program checks for the following type of issues:
	** checks for errors that cannot be detected by scto server or 
	** for compliance with IPA standards

	******////
		* check for situations of multiple Languages
		* check for situations of double

version 14.0
program define ipacheckscto
	syntax using/ [, OUTfile(str) OTHer(numlist max = 1) replace]
	
	qui {

		tempname  summ ch0
		tempfile  survey choices repeat repeat_long grouppairs
		tempfile  summary check0 check1 check2 check3 check4 check5 check6 check7

		* add dummies for exporting data

			if "`outfile'" ~= "" 	loc export 1
			else 					loc export 0 

		* add file extension if needed
		if !regexm("`using'", "\.xlsx|\.xls") loc using "`using'.xlsx"

		* set file
		if `export' {
			* include .xlsx extension by default if not included in filename
			if !regexm("`outfile'", "\.xlsx|\.xls") loc outfile "`outfile'.xlsx"
			* check that use didnt specify using as outfile
			if "`using'" == "`outfile'" {
				disp as err "options using and outfile cannot have the same {help filename:filename}"
				ex 602
			} 
			* check if file exist and warn user
			cap confirm file "`outfile'"
			if !_rc & "`replace'" == "" {
				noi disp as err "file `outfile' already exist. Specify replace to replace file"  
				ex 602
			}
			 * Remove old file 
			else if !_rc & "`replace'" ~= "" cap rm "`outfile'"
		}

		* save filenames in local
		if `export' loc filename = substr("`outfile'", -strpos(reverse(subinstr("`outfile'", "\", "/", .)), "/") + 1, .)

		* import information about xls form
		if `export' {
			import	excel using "`using'", sh("settings") first allstr clear
		
			postfile `summ' str23 field str100 (value comment) using "`summary'"

			* populate information from settings sheet into summary dataset

			post `summ' ("") 				("") 					("")
			post `summ' ("Form Details") 	("") 					("")
			post `summ' ("") 				("") 					("")
			post `summ' ("") 				("") 					("")
			post `summ' ("filename") 		("`filename'") 			("")
			post `summ' ("Form Title") 		("`=form_title[1]'") 	("")
			
			post `summ' ("Form ID") 				("`=form_id[1]'") 			("")
			post `summ' ("Form Definition Version") ("`=version[1]'") 			("")
			post `summ' ("Number of Languages") 	("`=default_language[1]'") 	("")

			if "`=public_key[1]'" ~= "" loc encrypted "Yes"
			else 						loc encrypted "No"

			post `summ' ("Form Encrypted") ("`encrypted'") ("")

			if "`=submission_url[1]'" ~= "" loc suburl 	"`=submission_url[1]'"
			else 							loc suburl 	"None"

			post `summ' ("Submission URL") ("`suburl'") ("")

			* populate names and extra text for other inputs
			post `summ' ("") 				("") ("")
			post `summ' ("Check Summary") 	("") ("")
			post `summ' ("") 				("") ("")

			post `summ' ("check") 					("description")															("result")
			post `summ' ("0. recommended vars") 	("Check if recommeded fields not included in survey") 					("")	
			post `summ' ("1. var length") 			("Check if length of fields is > 22 chars") 							("")
			post `summ' ("2. disabled, read only") 	("Checks for disabled or read only fields") 							("")
			post `summ' ("3. field requirements") 	("Checks for non-required fields and required note fields") 			("")
			post `summ' ("4. constraint") 			("Check that numeric fields are constraint") 							("")
			post `summ' ("5. other specify") 		("Check that or_other is not used and manual osp fields are specified") ("")
			post `summ' ("6. group names") 			("Check that group and repeat gropu names match in begin & end repeat") ("")
			post `summ' ("7. repeat vars") 			("Check that fields in repeat group are properly suffixed") 			("")

		}
		
		* Import Survey sheet
		import 	excel using "`using'", sheet("survey") first allstr clear 

		* Prepare import data by removing unneeded variables and observations 
		* and trimming vars
		prep_data

		* save survey data
		save "`survey'"
		
		* count number of label variables
		unab labels: label*
		loc labcount = wordcount("`labels'")
		
		* display titles
		noi disp "{hline}"
		noi disp "**" _column(5) "XLS REVIEW OF FILE: `using'"
		noi disp "{hline}"
				
		* Check 0: Check if recommeded fields are not not included in survey
		* ----------------------------------------------------------------
		* Auto generated starttime, endtime and duration are useful for IPA HFC Templates

		count if inlist(type, "start", "end") | inlist(calculation, "duration()")
		if `r(N)' < 3 {
			
			noi header, checknum(0) checkmessage("SOME RECOMMENDED FIELDS ARE MISSING") 
			
			noi disp  	"{p}The following fields are missing from your form. Please note that these field are " ///
						"required for IPA Data Quality Checks{p_end}"
			noi disp
			
			postfile 	`ch0' str50 variable str100 comment using `check0'
			post 		`ch0'	  	("fieldname") ("comment")
			
			loc i 3
			foreach name in starttime endtime duration {
				cap assert !regexm(name, "`name'")
				if !_rc {
					noi disp "`name'"
					loc ++i

					if `export' {
						post `ch0' ("`name'") ("`name' is needed for data quality checks")
					}
				}
				
			}

			noi disp 
			noi disp
			
		}

		* Check 1: Check that variable lengths do not exceed 22 chars
		* -----------------------------------------------------------

		use "`survey'", clear
		keep if length(name) > 22
		
		if `=_N' > 0 {
			
			noi header, checknum(1) checkmessage("SOME FIELD NAMES ARE TOO LONG")

			noi disp 			"{p}The following fields have names with character length greater than 22.{p_end}"
			noi disp

			gen char_length = length(name)
			keep row type name char_length 

			save "`check1'"

			noi list row type name char_length, noobs abbrev(32) sep(0) table
			noi disp
			
		}
			
		* Check 2: Check for disabled field and readonly field
		* ---------------------------------
		
		use "`survey'", clear
		keep if regexm(lower(disabled), "yes") | regexm(lower(readonly), "yes") 
		
		if `=_N' > 0 {
			noi header, checknum(2) checkmessage("DISABLED & READ ONLY FIELDS")
		
			noi disp 			"{p}The following fields have been disabled.{p_end}"
			noi disp

			keep row type name disabled readonly

			noi list row type name disabled readonly, noobs abbrev(32) sep(0)
			noi disp
			
			save "`check2'"
		}
		
		* Check 3: Check requirement settings. Flag the following:
			* field is not required & is ("integer|text|date|time|select")
			* field is required & is a note 
			* field is required & is readonly
			* field is required & has appearance type label
		* ----------------------------------
		
		use "`survey'", clear
		keep if (lower(required) ~= "yes" & regexm(type, "integer|(text)$|date|time|select_") & lower(appearance) ~= "label") | ///
				(lower(required) == "yes" & (type == "note" | lower(readonly) == "yes"))

		if `=_N' > 0 {
			noi header, checknum(3) checkmessage("FIELD REQUIREMENTS")

			noi disp 			"{p}The following fields have issues with requirement.{p_end}"

			keep row type name label appearance readonly required
			
			* label each issue type
			gen comment = 	cond(lower(required) ~= "yes" & regexm(type, "integer|(text)$|date|time|select_"), "field is not required", ///
							cond(lower(required) == "yes" & type == "note", "required note field", ///
								"required & read only"))
			
			noi list row type name appearance readonly required comment, noobs abbrev(32) sep(0)
			noi disp

			save "`check3'"

		}

		* Check 4: Constraint Messages
			* check that numeric fields are constraint
			* check that text fields are constraint if using appearance type numbers, numbers_phone
			* check that constrained fields have constraint messages
		* ----------------------------

		use "`survey'", clear
		keep if (missing(constraint) & inlist(type, "integer", "decimal")) 				| ///
				(missing(constraint) & type == "text" & regexm(appearance, "numbers")) 	| ///
				(!missing(constraint) & missing(constraintmessage)) 

		if `=_N' > 0 {

			noi header, checknum(4) checkmessage("CONSTRAINT")

			noi disp 			"{p}The following fields missing constraint or constraint message.{p_end}"

			keep row type name label appearance constraint constraintmessage

			* label each issue type
			gen comment = cond(!missing(constraint) & missing(constraintmessage), "missing constraint message", "missing constraint")
			
			noi list row type name appearance constraint constraintmessage comment, noobs abbrev(32) sep(0)
			noi disp

			save "`check4'"

		}


		* ---------------------------------------------------------------
		* Imort and prepare choices
		* ---------------------------------------------------------------

		import	excel using "`using'", sh("choices") first allstr clear
		
		* prepare data
		prep_data
		
		foreach var of varlist _all {
			cap assert missing(`var')
			if !_rc {
				drop `var'
			}
		}

		* make a list of list_name (s) with other specify
		if "`other'" ~= "" {
			levelsof list_name if value == "`other'", loc (other_list) clean
		}

		save "`choices'"


		* check 5: or_other and other specify
			* check that or_other is not used with select_one | select_multiple
			* check that fields using choices with other specify have defined an osp field
		*-------------------------------------

		use "`survey'", clear
		if wordcount("`other_list'") > 0 {

			* mark all fields using choices with other specify
			gen choice_other = 0
			foreach item in `other_list' {
				replace choice_other = 1 if word(type, 2) == "`item'" & regexm(type, "^(select_one)|^(select_multiple)")
			}
		}

		* flag fields with other specify
		keep row type name label relevance choice_filter choice_other
		gen child_index = ""
		gen child_name 	= ""
		gen child_row 	= .

		
		getrow if choice_other, loc (indexes)
		
		foreach index of numlist `indexes' {
			
			loc parent = name[`index']
			getrow if regexm(relevance, "{`parent'}") & regexm(relevance, "`other'"), loc(child_index)
			if "`child_index'" ~= "" {	
				replace child_name = name[`child_index'] in `index'
				replace child_row = row[`child_index'] in `index'
			}
		}

		keep if (regexm(type, "or_other$") & wordcount(type) == 3) | ///
				(missing(child_name) & choice_other) | ///
				(!missing(child_name) & (child_row < row) & choice_other)

		if `=_N' > 0 {
			
			noi header, checknum(5) checkmessage("OTHER SPECIFY")

			noi disp 			"{p}The following fields have missing other specify fields or use the or_other syntax.{p_end}"

			* generate comments for each issue
			gen comment = cond(regexm(type, "or_other$") & wordcount(type) == 3, "or_other syntax used", ///
						  cond(missing(child_name) & choice_other, "missing other specify field", ///
						  "other specify field [" + child_name + "] on row " + string(child_row) + " comes before parent field"))
			

			noi list row type name choice_filter comment, noobs abbrev(32) sep(0)
			noi disp

			save "`check5'"
		}

		* ---------------------------------------------------------------
		* Check and pair up group names
		* ---------------------------------------------------------------

		use "`survey'", clear

		keep if regexm(type, "group|repeat")

		if `=_N' > 0 {

			* generate new variables (begin_row begin_fieldname end_row end_fieldname)
			gen begin_row 		= .
			gen begin_fieldname = ""
			gen end_row			= .
			gen end_fieldname 	= ""

			* get the indexes of all begin groups|repeat and check names match pair
			getrow if (regexm(type, "^(begin)") & regexm(type, "group|repeat")), loc (indexes)

			count if regexm(type, "^(begin)")
			loc b_cnt `r(N)'
			count if regexm(type, "^(end)")
			loc e_cnt `r(N)'
			
			if `b_cnt' ~= `e_cnt' noi disp as err ///
				"Invalid form: There are `b_cnt' begin group/repeat and `e_cnt' end group/repeat fields"

			foreach index in `indexes' {				
				loc b 1
				loc e 0
				loc curr_index `index'
				loc stop 0
				while `stop' == 0 {
					loc ++curr_index 
					cap assert regexm(type, "^(end)") & regexm(type, "group|repeat") in `curr_index'
					if !_rc {
						loc ++e
						if `b' == `e' {
							loc end `curr_index'
							loc stop 1
						}
					}
					else loc ++b
				}

				replace begin_row 		= 	row[`index']		in `index'
				replace begin_fieldname =	name[`index']		in `index'
				replace end_row 		= 	row[`end']			in `index'
				replace end_fieldname 	=	name[`end']			in `index'

			}

			keep if regexm(type, "^(begin)")
			keep row type name label begin_row begin_fieldname end_row end_fieldname
			save "`grouppairs'"
		}

		* check 6: check group names
		*----------------------------

		keep if begin_fieldname ~= end_fieldname

		if `=_N' > 0 {
			noi header, checknum(6) checkmessage("GROUP NAMES")

			noi disp 			"{p}The following following groups have different names and begin and end.{p_end}"			

			noi list type begin_row begin_fieldname end_row end_fieldname, noobs abbrev(32) sep(0)
			noi disp

			save "`check6'"
		}

		* check 7: Repeat group vars
		*---------------------------

		if "`check6'" ~= "" {
			use "`survey'", clear

			merge 1:1 row using "`grouppairs'", nogen keepusing(begin_row begin_fieldname end_row end_fieldname)

			* mark out all variables in 
			gen rpt_field	= 0
			gen rpt_group 	= ""
			getrow if  type == "begin repeat", loc (indexes)
			* set trace on
			foreach index in `indexes' {
				loc groupname 			= begin_fieldname[`index']

				getrow if row == `=end_row[`index']', loc(lastrow)
				replace rpt_field 	= 1 in `index'/`lastrow'
				replace rpt_group 	= cond(missing(rpt_group), "`groupname'", rpt_group + "/" + "`groupname'") ///
										  in  `index'/`lastrow'
			}


			* foreach repeat group var, check if it was used outside the repeat group
			levelsof name if rpt_field, loc (rvars) clean

			** first reove functions that allow repeat vars

			#d;
			loc funcs
				"
				"join" 				"join-if"
				"sum"				"sum-if"
				"min" 				"min-if"
				"max"				"max-if"
				"rank-index" 		"rank-index-if"
				"indexed-repeat"
				"
				;
			#d cr
			
			foreach var of varlist appearance constraint relevance calculation repeat_count {
				
				replace `var' = subinstr(`var', "$", "#", .)

				* check for syntax that are used in the program
				foreach func in `funcs' {
					cap assert !regexm(`var', "`func'\(")
					loc rc = _rc
					while `rc' == 9 {
						getsyntax `var', function("`func'") gen(function)
						replace `var' = subinstr(`var', function, "##", .)
						drop function

						cap assert !regexm(`var', "`func'\(")
						loc rc = _rc
					}
				}

			}

			* check for inappropraite use of syntax
			gen rpt_flag 	= 0
			gen rpt_flagvar = "/"
			gen rpt_flagcol	= "/"

			#d;
			unab cols:
				label* 
				mediaimage* 
				mediaaudio* 
				mediavideo* 
				appearance 
				constraint 
				relevance 
				calculation 
				repeat_count 
				;
			#d cr
			
			foreach var of varlist `cols' {
			
				foreach rvar in `rvars' {
					levelsof rpt_group if name == "`rvar'", loc(rvar_group) clean 
						replace rpt_flag 	= 1 if regexm(`var', "{`rvar'}") & rpt_group != "`rvar_group'" 
						replace rpt_flagvar = rpt_flagvar + "`rvar'/" if regexm(`var', "{`rvar'}") & rpt_group != "`rvar_group'"
						replace rpt_flagcol = rpt_flagcol + "`var'/" if regexm(`var', "{`rvar'}") & rpt_group != "`rvar_group'"
				}				

			}

			keep row type name label rpt_flag rpt_flagvar rpt_flagcol

			keep if rpt_flag

			gen sheet = "survey"

			save "`check7'"


			* import and check choices sheet
			use "`choices'", clear

			gen rpt_flag 	= 0
			gen rpt_flagvar = "/"
			gen rpt_flagcol	= "/"

			loc choice_flag_cnt 0
			foreach var of varlist label* {
				foreach rvar in `rvars' {
						replace rpt_flag 	= 1 if regexm(`var', "{`rvar'}") 
						replace rpt_flagvar = rpt_flagvar + "`rvar'/" if regexm(`var', "{`rvar'}")
						replace rpt_flagcol = rpt_flagcol + "`var'/" if regexm(`var', "{`rvar'}")
				}	
			}
			


			if `=_N' > 0 {
				replace rpt_flagvar = itrim(trim(subinstr(rpt_flagvar, "/", " ", .)))
				replace rpt_flagcol = itrim(trim(subinstr(rpt_flagcol, "/", " ", .)))

				split rpt_flagvar
				split rpt_flagcol

				drop rpt_flagcol rpt_flagvar

				ren rpt_flagcol column

				reshape long rpt_flagvar rpt_flagcol, i(row) j(instance)
				drop instance

				ren rpt_flagvar repeat_field

				noi header, checknum(7) checkmessage("GROUP NAMES")

				noi disp 			"{p}The following following fields contain repeat group fields that have been used illegally.{p_end}"			

				noi list row type name repeat_field column, noobs abbrev(32) sep(0)
				noi disp

				save "`check7'"


			}

		}

	}
		
		
end


* Program to remove unneeded variables from inported worksheet
program define prep_data
	syntax
		* loop through variables and remove additional vars that may have been imported from worksheet
		foreach var of varlist _all {
			* Check that variable name has at least 4 chars, else drop 
				*** This should be changed if scto introduces a 3 char name column***
			if length("`var'") < 	4 drop `var'

			* trim values
			if length("`var'") >= 	4 replace `var' = trim(itrim(`var')) 
		}

		* generate excel rows for variable
		gen row = _n + 1

		* drop all empty rows
		cap drop if missing(type)
		cap drop if missing(list_name)
		cap drop index
end

* program to display headers for each check
program define header
	syntax, checknum(integer) checkmessage(string) [main]
		* if main is specified use 2 lines of stars
		noi disp "{hline}"
		noi disp "CHECK #`checknum':" _column(15) " `checkmessage'"
		noi disp "{hline}"
		noi disp
end

* returns row numbers based on a condition
program define getrow, rclass

	syntax [if], LOCal(name)
	
	tempvar row
	
	gen `row' = _n
	levelsof `row' `if', clean loc (rows)
	c_local `local' "`rows'"

end 

program define getsyntax 

	syntax varname, FUNCTION(string) GENerate(name)
	* set trace on
	getrow if regexm(`varlist', "`function'\("), loc (rows)
	
	gen `generate' 		= ""
	loc len_function 	= length("`function'") + 1
	
	foreach row in `rows' {
	    
		loc text = `varlist'[`row']
		loc len_text = length("`text'")
		loc cont 1
		loc spos = strpos("`text'", "`function'(")
		loc start = `spos' + `len_function' + 1
		
		loc open 	1
		loc close 	0 
		while `cont' == 1 {
			loc char = substr("`text'", `start', 1)
			if 		"`char'" == "(" loc ++open
			else if "`char'" == ")" loc ++close
			
			if `open' == `close' {
				replace `generate' = substr("`text'", `spos', `start') in `row'
				loc cont 0
			}
			else if `start' > `len_text' {
				di as error "Invalid syntax: Unequal number of ( or ) on row `row'"
				loc cont 0
			}
			else loc ++start

		}
		
	}
	


end



* adjust outfile columns
mata:
mata clear
void adjust_column_width(string scalar filename, string scalar sheet, real rowvector columns, real rowvector widths)
{
class xl scalar b

	b = xl()
	b.load_book(filename)
	b.set_sheet(sheet)
	b.set_mode("open")

	for (i = 1;i <= length(columns); i++) {
		b.set_column_width(columns[i], columns[i], widths[i])
	}

	b.close_book()

}
end
