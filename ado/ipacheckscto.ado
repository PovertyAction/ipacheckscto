*! version 1.0 (ALPHA) 
*! Ishmail Azindoo Baako (IPA) February, 2018

* Stata program to review SurveyCTO xls form and export an excel sheet with issues
* This program only checks for errors that cannot be detected by scto server or 
* for compliance with IPA standards

version 14.0
program define ipacheckscto
	syntax using/, [OUTfile(string) replace]
	
	qui {

		tempfile 	_survey 	_choices 	_setting 	_repeat		_repeat_long

		* import survey sheet 
		if !regexm("`using'", "\.xlsx|\.xls") loc using "`using'.xlsx"

				* set file
		if "`outfile'" ~= "" {
			* include .xlsx extension by default if its not included
			if !regexm("`outfile'", "\.xlsx|\.xls") loc outfile "`outfile'.xlsx"
			* check that use didnt specify using as outfile
			if "`using'" == "`outfile'" {
				disp as err "options using and outfile cannot have the same {help filename:filename}"
				exit 602
			} 
			* check if file exist and warn user
			cap confirm file "`outfile'"
			if !_rc & "`replace'" == "" {
				noi disp as err "file `outfile' already exist. specify replace to replace file"  
				exit 602
			}
			 * Remove old file 
			else if !_rc & "`replace'" ~= "" cap rm "`outfile'"
		}


		* import information about xls form
		if "`outfile'" ~= "" {
			import	excel using "`using'", sh("settings") first allstr clear
			* export form information
			keep in 1
			keep form_title form_id version public_key submission_url default_language
		
			loc row 2
			loc col1_len 2
			loc col2_len 2
			putexcel set "`outfile'", sheet("Form Details") replace
			foreach var of varlist _all {
				putexcel A`row' = "`var'" B`row' = `var'[1]
				loc ++row

				* check for length of columns
				if length("`var'") > `col1_len' loc col1_len = length("`var'") + 2
				if length(`var'[1]) > `col2_len' loc col2_len = length(`var'[1]) + 2
			}

			
				putexcel A2:B2, border(top, thick)
				putexcel A2:A`=`row'-1', border(left, thick) bold
				putexcel B2:B`=`row'-1', border(right, thick)
				putexcel A`row':B`row', border(top, thick)
			
				mata: adjust_column_width("`outfile'", "Form Details", (1, 2), (`col1_len', `col2_len'))
		}
		
		import 	excel using "`using'", sh("survey") first allstr clear 

		* Prepare import data by removing unneeded variables and observations 
		* and trimming vars
		prep_data
		
		* count number of label variables
		unab labels: label*
		loc labcount = wordcount("`labels'")
		
		* display titles
		noi disp "{hline}"
		noi disp "**" _column(5) "XLS REVIEW OF FILE: `using'"
		noi disp "{hline}"
				
		* Check 0: Variables starttime and endtime are present in the form
		* ----------------------------------------------------------------
		* Auto generated starttime and endtime are useful for IPA HFC Templates

		count if regexm(name, "starttime|endtime")
		if `r(N)' < 2 {
		
			noi header, checknum(0) checkmessage("MISSING FIELDS") 
			
			noi disp 			"{p}The following fields are missing from your form. Please note that these field are " ///
								"important for IPA Data Quality Checks{p_end}"
			noi disp
			
			if "`outfile'" ~= "" {
				* output headers
				putexcel set "`outfile'", sheet("0. missing fields") modify
				putexcel A1 = "fieldname" B1 = "message", bold border(bottom)
			}

			noi disp 			"{ul:missing field}"
			
			loc i 2
			foreach name in starttime endtime {
				cap assert !regexm(name, "`name'")
				if !_rc {
					noi disp "`name'"
					if "`outfile'" ~= "" putexcel 	A`i' = "`name'" ///
								B`i' = "default field `name' is missing. This is required for for IPA Checks"
					loc ++i
				}
				
			}

			if "`outfile'" ~= "" mata: adjust_column_width("`outfile'", "0. missing fields", (1, 2), (10, 60))

			noi disp 
			noi disp
			
		}
		
		* Check 1: Check that variable lengths do not exceed 22 chars
		* -----------------------------------------------------------

		count if length(name) > 22
		if `r(N)' > 0 {
			noi header, checknum(1) checkmessage("LENGTH OF FIELD NAMES")

			noi disp 			"{p}The following `r(N)' fields have names with character length greater than 22. " ///
								"It is recommended to keep field names shorter than 22 characters to allow renaming during" ///
								"data import, HFCs and data cleaning{p_end}"
			noi disp

			gen char_length = length(name)

			noi disp 			"{ul:Variables with field names greater than 22 characters}"
			noi list row type name char_length if char_length > 22, noobs abbrev(32) sep(0) table
			noi disp
			
			* output results to excel
			if "`outfile'" ~= "" {
				export excel row type name label* char_length ///
					using "`outfile'" if char_length > 22, sheet("1. field names") sheetmodify first(var)
				
				* format headers
				loc colnum = 4 + `labcount'
				alphacol `colnum'
				loc colname "`r(alphacol)'"
				putexcel set "`outfile'", sheet("1. field names") modify
				putexcel A1:`colname'1, bold border(bottom)

				mata: adjust_column_width("`outfile'", "1. field names", (1, 2, 3, 4, 5), (5, 30, 40, 100, 12))
				putexcel D1:D`=_N', txtwrap
			}
		} 

		* Check 2: Check for disabled field
		* ---------------------------------
		count if regexm(disabled, "[Yy][Ee][Ss]")
		if `r(N)' > 0 {
			noi header, checknum(2) checkmessage("DISABLED FIELDS")

			noi disp 			"{p}The following `r(N)' fields have been disabled. " ///
								"Please note disabled fields will not appear in Survey{p_end}"
			noi disp

			noi disp 			"{ul:Disabled Fields}"
			noi list row type name disabled if regexm(disabled, "[Yy][Ee][Ss]"), noobs abbrev(32) sep(0)
			noi disp
			
			if "`outfile'" ~= "" {
				export excel row type name label* disabled ///
					using "`outfile'" if regexm(disabled, "[Yy][Ee][Ss]"), ///
					sheet("2. disabled fields") sheetmodify first(var)
					
				* format headers
				loc colnum = 4 + `labcount'
				alphacol `colnum'
				loc colname "`r(alphacol)'"
				putexcel set "`outfile'", sheet("2. disabled fields") modify
				putexcel A1:`colname'1, bold border(bottom)

				mata: adjust_column_width("`outfile'", "2. disabled fields", (1, 2, 3, 4, 5), (5, 30, 40, 100, 12))
				putexcel D1:D`=_N', txtwrap
			}
		}
		
		* Check 3: Check for readonly fields
		* ----------------------------------
		count if regexm(readonly, "[Yy][Ee][Ss]") 
		if `r(N)' > 0 {
			noi header, checknum(3) checkmessage("READ ONLY FIELDS")

			noi disp 			"{p}The following `r(N)' fields are readonly. " ///
								"Please note readonly fields do not accept inputs from users{p_end}"
			noi disp

			noi disp 			"{ul:Readonly Fields}"
			noi list row type name readonly if regexm(readonly, "[Yy][Ee][Ss]"), noobs abbrev(32) sep(0)
			noi disp
			
			if "`outfile'" ~= "" {
				export excel row type name label* readonly ///
					using "`outfile'" if regexm(readonly, "[Yy][Ee][Ss]"), ///
					sheet("3. readonly fields") sheetmodify first(var)
					
				* format headers
				loc colnum = 4 + `labcount'
				alphacol `colnum'
				loc colname "`r(alphacol)'"
				putexcel set "`outfile'", sheet("3. readonly fields") modify
				putexcel A1:`colname'1, bold border(bottom)

				mata: adjust_column_width("`outfile'", "3. readonly fields", (1, 2, 3, 4, 5), (5, 30, 40, 100, 12))
				putexcel D1:D`=_N', txtwrap
			}
		}
		
		* Check 4: Check for non required fields
		* ----------------------------------
		count if !regexm(required, "[Yy][Ee][Ss]") & regexm(type, "integer|text|date|time|select") & ///
			!regexm(type, "text audit") & appearance != "label"

		if `r(N)' > 0 {
			noi header, checknum(4) checkmessage("NON-REQUIRED FIELDS")

			noi disp 			"{p}The following `r(N)' fields are not required. " ///
								"It is recommended that these fields are required to avoid missing data{p_end}"

			count if regexm(type, "select") & appearance == "label" & !regexm(required, "[Yy][Ee][Ss]")
			if `r(N)' > 0 noi disp _column(10)	"NB: Excludes: `r(N)' non-required fields with appearance type -- label"

			noi disp

			noi disp 			"{ul:non-required fields}"
			noi list row type name required if !regexm(required, "[Yy][Ee][Ss]") & regexm(type, "integer|text|date|time|select") & ///
				!regexm(type, "text audit")  & appearance != "label", noobs abbrev(32) sep(0)
			noi disp
			
			if "`outfile'" ~= "" {
				export excel row type name label* required ///
					using "`outfile'" ///
					if !regexm(required, "[Yy][Ee][Ss]") & regexm(type, "integer|text|date|time|select") & ///
					!regexm(type, "text audit")  & appearance != "label", ///
					sheet("4. non-required fields") sheetmodify first(var)
					
				* format headers
				loc colnum = 4 + `labcount'
				alphacol `colnum'
				loc colname "`r(alphacol)'"
				putexcel set "`outfile'", sheet("4. non-required fields") modify
				putexcel A1:`colname'1, bold border(bottom)

				mata: adjust_column_width("`outfile'", "4. non-required fields", (1, 2, 3, 4, 5), (5, 30, 40, 100, 12))
				putexcel D1:D`=_N', txtwrap
			}

		}
	
		* Check 5: Required notes and geopoint and appearance(label)
		* ----------------------------------------------------------
		count if (regexm(type, "note|geopoint|geoshape|geotrace") | (regexm(type, "select") & appearance == "label")) & ///
			regexm(required, "[Yy][Ee][Ss]")

		if `r(N)' > 0 {
			noi header, checknum(5) checkmessage("REQUIRED NOTES, GPS TYPES OR APPEARANCE TYPE (label")
			
			* gen message var
			gen message = ""
			
			count if regexm(type, "note") & regexm(required, "[Yy][Ee][Ss]")
			if `r(N)' > 0 {
				noi disp 			"{p}The following `r(N)' note field(s) are are required. " ///
									"Note fields should only be required when used as checks within a form{p_end}"

				noi disp

				noi disp 			"{ul:Required Note Fields}"
				noi list row type name required if regexm(type, "note") ///
					& regexm(required, "[Yy][Ee][Ss]"), noobs abbrev(32) sep(0)
				
				replace message = "Required notes can prevent users from finalizing forms. Note fields should only be required when used as checks within a form" ///
					if regexm(type, "note") & regexm(required, "[Yy][Ee][Ss]")
			}

			count if regexm(type, "geopoint|geoshape|geotrace") & regexm(required, "[Yy][Ee][Ss]")
			if `r(N)' > 0 {
				noi disp 			"{p}The following `r(N)' gps type field(s) are are required." ///
									"Forms cannot be saved if the device fails to capture GPS{p_end}"

				noi disp

				noi disp 			"{ul:Required GPS Fields}"
				noi list row type name required if regexm(type, "geopoint|geoshape|geotrace") & regexm(required, "[Yy][Ee][Ss]"), ///
					noobs abbrev(32) sep(0)
					
				replace message = "Forms cannot be saved if the device fails to capture GPS data" ///
					if regexm(type, "geopoint|geoshape|geotrace") & regexm(required, "[Yy][Ee][Ss]")
			}
			
			* gen lowercase values for appearance
			gen appearance_l = lower(appearance)
			count if regexm(type, "select")  & appearance_l == "label" & regexm(required, "[Yy][Ee][Ss]")
			if `r(N)' > 0 {
				noi disp 			"The following `r(N)' field(s) with appearance (label) are required."
				noi disp _column(5)	"Forms cannot be finalized if these fields are required"

				noi disp

				noi disp 			"{ul:Required Label Fields}"
				noi list row type name appearance required ///
					if regexm(type, "select") & appearance_l == "label" & regexm(required, "[Yy][Ee][Ss]"), ///
					noobs abbrev(32) sep(0)
					
				replace message = `"Fields with appearance type "label" should not be required"' ///
					if regexm(type, "select") & appearance_l == "label" & regexm(required, "[Yy][Ee][Ss]")
			
			}
			
			if "`outfile'" ~= "" {
				export excel row type name label* appearance required message ///
					using "`outfile'" if !missing(message), ///						
					sheet("5. required fields") sheetmodify first(var)
					
				* format headers
				loc colnum = 6 + `labcount'
				alphacol `colnum'
				loc colname "`r(alphacol)'"
				putexcel set "`outfile'", sheet("5. required fields") modify
				putexcel A1:`colname'1, bold border(bottom)

				mata: adjust_column_width("`outfile'", "5. required fields", (1, 2, 3, 4, 5, 6, 7), (5, 30, 40, 100, 40, 10, 100))
				putexcel D1:D`=_N', txtwrap
				putexcel G1:G`=_N', txtwrap
			}
			
			* drop message var
			drop message
		}
		
		* Check 6: Constraint Messages
		* ----------------------------
		count if !missing(constraint) & missing(constraintmessage)

		if `r(N)' > 0 {
			noi header, checknum(6) checkmessage("NO CONTRAINT MESSAGES")

			noi disp 			"{p}The following `r(N)' fields with constraints do not have constraint messages " ///
								"It is recommended that you include a clear constraint message to make it easier for enumerators " ///
								"to understand the reason their inputs are invalid.{p_end}" ///
			
			noi disp

			noi disp 			"{ul:Fields with missing contraint messages}"
			noi list row type name constraint constraintmessage if !missing(constraint) & missing(constraintmessage), noobs abbrev(32) sep(0)
			noi disp
			
			if "`outfile'" ~= "" {
				gen message = "Missing constraint mesaage. Indicate a clear constraint message for this field" ///
					if !missing(constraint) & missing(constraintmessage)
			
				export excel row type name label* constraint constraintmessage message ///
					using "`outfile'" if !missing(constraint) & missing(constraintmessage), ///						
					sheet("6. constraint messages") sheetmodify first(var)
					
				* format headers
				loc colnum = 6 + `labcount'
				alphacol `colnum'
				loc colname "`r(alphacol)'"
				putexcel set "`outfile'", sheet("6. constraint messages") modify
				putexcel A1:`colname'1, bold border(bottom)

				mata: adjust_column_width("`outfile'", "6. constraint messages", (1, 2, 3, 4, 5, 6, 7), (5, 30, 40, 100, 50, 17, 65))
				putexcel D1:D`=_N', txtwrap
				
				* drop message
				drop message
			}

		}
		
		* Check 7: uncontrained numeric field
		* -----------------------------------
		count if inlist(type, "integer", "decimal") & missing(constraint)
		
		if `r(N)' > 0 {
				noi header, checknum(7) checkmessage("UNCONSTRAINED NUMERIC FIELDS")
				
				noi disp 			"{p}The following `r(N)' numeric field(s) are not constrained " ///
									"It is recommended that all numeric fields (integer, decimal) be constrained{p_end}"
				
				noi disp

				noi disp 			"{ul:numeric fields with no constraint}"
				noi list row type name constraint if inlist(type, "integer", "decimal") & missing(constraint), noobs abbrev(32) sep(0)
				noi disp
				
				if "`outfile'" ~= "" {
				gen message = "It is recommeded to constraint numeric fields to reduce errors" ///
					if inlist(type, "integer", "decimal") & missing(constraint)
			
				export excel row type name label* constraint ///
					using "`outfile'" if inlist(type, "integer", "decimal") & missing(constraint), ///						
					sheet("7. unconstrained numeric fields") sheetmodify first(var)
					
				* format headers
				loc colnum = 4 + `labcount'
				alphacol `colnum'
				loc colname "`r(alphacol)'"
				putexcel set "`outfile'", sheet("7. unconstrained numeric fields") modify
				putexcel A1:`colname'1, bold border(bottom)

				mata: adjust_column_width("`outfile'", "7. unconstrained numeric fields", (1, 2, 3, 4, 5), (5, 30, 40, 100, 40))
				putexcel D1:D`=_N', txtwrap
				
				* drop message
				drop message
			}


		}
		
		* Check 8: or_other
		* -----------------
		count if regexm(type, "select") & regexm(type, "or_other")
		
		if `r(N)' > 0 {
			noi header, checknum(8) checkmessage("OR_OTHER")

			noi disp 			"{p}The following `r(N)' field(s) use the field type or_other " ///
								"It is recommended that you define a text field for other specify{p_end}"

			noi disp

			noi disp 			"{ul:fields with or_other}"
			noi list row type name if regexm(type, "select") & regexm(type, "or_other"), noobs abbrev(32) sep(0)
			noi disp
			
			if "`outfile'" ~= "" {
				gen message = "It is recommended that you define a text field for other specify" ///
					if regexm(type, "select") & regexm(type, "or_other")
			
				export excel row type name label* ///
					using "`outfile'" if regexm(type, "select") & regexm(type, "or_other"), ///						
					sheet("8. or_other fields") sheetmodify first(var)
					
				* format headers
				loc colnum = 3 + `labcount'
				alphacol `colnum'
				loc colname "`r(alphacol)'"
				putexcel set "`outfile'", sheet("8. or_other fields") modify
				putexcel A1:`colname'1, bold border(bottom)

				mata: adjust_column_width("`outfile'", "8. or_other fields", (1, 2, 3, 4), (5, 30, 40, 100))
				putexcel D1:D`=_N', txtwrap
				
				* drop message
				drop message
			}

		}
		
		* Check 9: Group Naming
		* ---------------------
		* save data
		save `_survey' 

		* keep only group and repet field
		count if regexm(type, "group|repeat")

		if `r(N)' > 0 {

			* keep only groups and repeats
			keep if regexm(type, "group|repeat")
			
			* generate _n to mark groups
			gen _sn = _n
			
			* generate new variables (begin_row begin_fieldname end_row end_fieldname)
			gen begin_row 		= .
			gen begin_fieldname = ""
			gen end_row			= .
			gen end_fieldname 	= ""
			* cls
			* set trace on
			* get the name of all begin groups|repeat and check if the name if their pairs match
			levelsof _sn if (regexm(type, "^(begin)") & regexm(type, "group|repeat")), ///
				loc (_sns) clean
			
			count if regexm(type, "^(begin)")
			loc b_cnt `r(N)'
			count if regexm(type, "^(end)")
			loc e_cnt `r(N)'
			
			if `b_cnt' ~= `e_cnt' noi di as err "Invalid form: There are `b_cnt' begin types and `e_cnt' end types"
		
			foreach _sn in `_sns' {				
				loc b 1
				loc e 0
				loc curr_sn `_sn'
				loc stop 0
				while `stop' == 0 {
					loc ++curr_sn 
					cap assert regexm(type, "^(end)") & regexm(type, "group|repeat") in `curr_sn'
					if !_rc {
						loc ++e
						if `b' == `e' {
							loc end `curr_sn'
							loc stop 1
						}
					}
					else {
						loc ++b
					}
				}

				replace begin_row 		= 	row[`_sn']		in `_sn'
				replace begin_fieldname =	name[`_sn']		in `_sn'
				replace end_row 		= 	row[`end']		in `_sn'
				replace end_fieldname 	=	name[`end']		in `_sn'
			}
			
			* save data
			save `_repeat', replace
			
			* count violations
			count if begin_fieldname != end_fieldname
			
			if `r(N)' > 0 {
				noi header, checknum(9) checkmessage("UNMATCHED BEGIN GROUP AND REPEAT")

				noi disp 			"{p}The following `r(N)' begin field(s) do not have matching names for their pairs " ///
									"It is recommended that you begin and end groups with the same name{p_end}"

				noi disp

				noi disp 			"{ul:begin fields with fields with unmatching names}"
				noi list type begin_row begin_fieldname end_row end_fieldname ///
					if begin_fieldname != end_fieldname, noobs abbrev(32) sep(0)

				if "`outfile'" ~= "" {
					gen message = "It is recommended that you begin and end groups/repeats with the same name" ///
						if begin_fieldname != end_fieldname
			
					export excel type name label* begin_row begin_fieldname end_row end_fieldname message ///
						using "`outfile'" ///
						if begin_fieldname != end_fieldname, ///
						sheet("9. umatched group names") sheetmodify first(var)
					
					* format headers
					loc colnum = 7 + `labcount'
					alphacol `colnum'
					loc colname "`r(alphacol)'"
					putexcel set "`outfile'", sheet("9. umatched group names") modify
					putexcel A1:`colname'1, bold border(bottom)

					mata: adjust_column_width("`outfile'", "9. umatched group names", (1, 2, 3, 4, 5, 6, 7, 8), (5, 30, 40, 100, 10, 40, 10, 40, 70))
					putexcel D1:D`=_N', txtwrap
				}
			}
		}
		
		* Check if there is a need to do checks 10 and 11		
		count if type == "begin repeat"
		if `r(N)' > 0 {
			
			* Merge in survey sheet
			use `_survey', clear
			merge 1:1 row using `_repeat', ///
				keepusing(begin_row begin_fieldname end_row end_fieldname) nogen

			* get the names of all repeat groups (unsorted)
			loc rpts ""
			forval i = 1/`=_N' {
				if type[`i'] == "begin repeat" {
					loc name = name[`i']
					loc rpts = "`rpts' `name'"
				}
			}
		
			* generate variable to hold names of repeat groups in which variables are
			gen repeat 		= ""
			gen nrpt_count 	= 0

			* mark repeat variables
			foreach rpt in `rpts' {
				levelsof row if type == "begin repeat" 	& name == "`rpt'", loc (start) 	clean
				levelsof end_row if type == "begin repeat" 	& name == "`rpt'", loc (end) 	clean

				replace repeat 		= repeat + " " + "`rpt'" 	if inrange(row, `start', `end')
				loc ++start
				loc --end 

				count 							if inrange(row, `start', `end') & inlist(type, "begin repeat") 

				if `r(N)' > 0 replace nrpt_count = `r(N)'		if inrange(row, `start', `end') ///
																& !inlist(type, "begin group", "end group", "begin repeat", "end repeat")
			}

			* Check 10: Repeat Variables
			* -------------------------

			* generate suffixes for all repeat group vars
			* gen dummy to mark repeat variables
			gen repeat_var = !missing(repeat) & !inlist(type, "begin group", "end group", "begin repeat", "end repeat")
			gen suffix = substr(name, -(strpos(reverse(name), "_")), .) if !missing(repeat) & repeat_var
			
			* gen repeat check 
			gen repeat_check = "/" + subinstr(repeat, " ", "/", .) + "/"
			
			* foreah repeat, for the most common suffix used and save in variable
			gen main_suffix = ""
			foreach rpt in `rpts' {
				levelsof suffix if regexm(repeat_check, "/`rpt'/") & repeat_var, loc(suffixes) clean
				loc suffix_count = wordcount("`suffixes'")

				if `suffix_count' == 1 replace main_suffix = "`suffixes'" if regexm(repeat_check, "/`rpt'/") & repeat_var
				if `suffix_count' >	 1 {
					gen suffix_count = 0
					foreach suffix in `suffixes' {
						count 							if regexm(repeat_check, "/`rpt'/") & repeat_var & suffix == "`suffix'"
						replace suffix_count = `r(N)' 	if regexm(repeat_check, "/`rpt'/") & repeat_var & suffix == "`suffix'"
					}

					* check for the suffix with the highest count and replace main_suffix
					summ suffix_count							if regexm(repeat_check, "/`rpt'/") & repeat_var
					levelsof suffix if suffix_count == `r(max)' & 	regexm(repeat_check, "/`rpt'/") & repeat_var, loc (main_suffix) clean
					replace main_suffix = "`main_suffix'"		if regexm(repeat_check, "/`rpt'/") & repeat_var 

					drop suffix_count
				}
			}

			* check that some suffixes dont match and list fields with unmatching suffixes
			count if suffix != main_suffix & repeat_var
			if `r(N)' > 0 {
				noi header, checknum(10) checkmessage("UNMATCHING REPEAT VARIABLE SUFFIXES")

				noi disp 			"{p}The following `r(N)' field(s) do not having matching suffixes their repeat groups " ///
									"It is recommended that all variables in repeat groups end in the same suffix. eg. _r, _rr, _rpt etc{p_end}"

				noi disp

				noi disp 			"{ul:fields with unmatching variable suffixes}"
				noi list row type repeat name suffix main_suffix if suffix != main_suffix & repeat_var, noobs abbrev(32) sep(0)
			}
			
			if "`outfile'" ~= "" {
				gen message = "It is recommended to end repeat group vars with the same suffix. eg. _r _rr _rpt" ///
					if suffix != main_suffix & repeat_var
				
				count if suffix != main_suffix & repeat_var 
				if `r(N)' > 0 {
					export excel row type name label* repeat suffix main_suffix message ///
						using "`outfile'" ///
						if suffix != main_suffix & repeat_var, ///
						sheet("10. repeat var names") sheetmodify first(var)
						
					* format headers
					loc colnum = 7 + `labcount'
					alphacol `colnum'
					loc colname "`r(alphacol)'"
					putexcel set "`outfile'", sheet("10. repeat var names") modify
					putexcel A1:`colname'1, bold border(bottom)

					mata: adjust_column_width("`outfile'", "10. repeat var names", (1, 2, 3, 4, 5, 6, 7), (5, 30, 40, 100, 40, 30, 30, 70))
					putexcel D1:D`=_N', txtwrap
				}
			}

			* Check 11: Repeat Variables
			* -------------------------

			* get names of all repeat fieds 

			gen rpt_flag 	= 0
			gen rpt_flagvar = "/"
			gen rpt_var		= ""
			foreach rpt in `rpts' {
				levelsof name if repeat_var & regexm(repeat_check, "/`rpt'/"), loc (rpt_vars) clean
				foreach rpt_var in `rpt_vars' {
					foreach var of varlist label* hint* appearance constraint relevance repeat_count choice_filter calculation {
						if "`var'" == "calculation" {
							replace rpt_flag 		= rpt_flag 	  +  1 ///
								if regexm(`var', "{`rpt_var'}") & !regexm(repeat_check, "/`rpt'/") ///
								&  !(regexm(calculation, "sum\(|min\(|max\(|join\(|indexed-repeat"))
								
							replace rpt_flagvar 	= rpt_flagvar + "/`var'[`rpt_var']/"  ///
								 if regexm(`var', "{`rpt_var'}") & !regexm(repeat_check, "/`rpt'/") ///
								&  !(regexm(calculation, "sum\(|min\(|max\(|join\(|indexed-repeat")) 
						}
						else {
							replace rpt_flag 		= rpt_flag 	  +  1 ///
								if regexm(`var', "{`rpt_var'}") & !regexm(repeat_check, "/`rpt'/")
							replace rpt_flagvar 	= rpt_flagvar + "/`var'[`rpt_var']/"  ///
								 if regexm(`var', "{`rpt_var'}") & !regexm(repeat_check, "/`rpt'/")
						}
					}
				}
			}
			
			su rpt_flag
			loc max_flag `r(max)'
			
			if `max_flag' > 0 {
				* keep relevant variables 
				keep row type name label* hint* disabled appearance constraint relevance repeat_count ///
					calculation choice_filter rpt_flagvar
				
				* check for the max number of flags and split apt_flagvar
				replace rpt_flagvar = trim(itrim(subinstr(rpt_flagvar, "/", " ", .)))
				
				forval i = 1/`max_flag' {
					gen column_`i' = word(rpt_flag, `i')
				}
				
				* reshape data to long format
				reshape long column_, ///
					i(row) j(index)
				
				* clean up
				ren column_ column
				keep if !missing(column)
				keep row type name label* hint* disabled appearance constraint ///
					relevance calculation repeat_count choice_filter column
					
				* gen repeat_field to hold the name of the repeat field used
				replace column 		= subinstr(column, "[", " ", .)
				replace column 		= subinstr(column, "]", "", .)
				gen repeat_field	= word(column, 2)
				replace column 		= trim(itrim(subinstr(column, repeat_field, "", .)))
				
				noi disp 			"{ul:repeat fields used outside repeat group}"
				noi list row type name disabled column repeat_field, noobs abbrev(32) sep(0)
				
				if "`outfile'" ~= "" {
					gen message = "Field " + repeat_field + ///
						" is from a repeat group and should not be referenced outside the group unless is been used in a calulate field with special functions"
					
					* drop columns with all miss
					foreach var of varlist _all {
						cap assert missing(`var')
						if !_rc {
							drop `var'
						}
					}
					
					export excel using "`outfile'", sheet("11. repeat fields") sheetmodify first(var)
					
					* format headers
					d, s
					alphacol `r(k)'
					loc colname "`r(alphacol)'"
					putexcel set "`outfile'", sheet("11. repeat fields") modify
					putexcel A1:`colname'1, bold border(bottom)

					mata: adjust_column_width("`outfile'", "11. repeat fields", (1, 2, 3, 4, 5, 6, 7, 8), (5, 30, 40, 100, 100, 30, 30, 100))
					putexcel D1:D`=_N', txtwrap
					putexcel E1:E`=_N', txtwrap
					putexcel H1:H`=_N', txtwrap
				}
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

* alphacol by chris boyer
program alphacol, rclass
	syntax anything(name = num id = "number")

	local col = ""

	while `num' > 0 {
		local let = mod(`num'-1, 26)
		local col = char(`let' + 65) + "`col'"
		local num = floor((`num' - `let') / 26)
	}

	return local alphacol = "`col'"
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


