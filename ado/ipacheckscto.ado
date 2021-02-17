*! version 1.1
*! Ishmail Azindoo Baako (IPA) February, 2018

* Stata program to review SurveyCTO xls form and export issues to an excel sheet
* This program checks for the following type of issues:
	** checks for errors that cannot be detected by scto server or 
	** for compliance with IPA standards

version 14.0
program define ipacheckscto
	syntax using/ [, OUTfile(str) OTHer(numlist max = 1) replace]
	
	qui {

		tempname  summ ch0
		tempfile  survey choices repeat repeat_long grouppairs
		tempfile  summary check0 check1 check2 check3 check4 check5 check6 check7 check8

		* add dummies for exporting data

			if "`outfile'" ~= "" 	loc export 1
			else 					loc export 0 

		* add file extension if needed
		if !regexm("`using'", "\.xlsx|\.xls") loc using "`using'.xlsx"

		* set export file
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

		* save checks and descriptions in locals to allow for easy updates

		loc checkname0 "0. recommended vars"
		loc checkdesc0 "Check if recommeded fields not included in survey"
		loc checkname1 "1. var length"
		loc checkdesc1 "Check if length of fields is > 22 chars"
		loc checkname2 "2. disabled, read only"
		loc checkdesc2 "Checks for disabled or read only fields"
		loc checkname3 "3. field requirements"
		loc checkdesc3 "Checks for non-required fields and required note fields"
		loc checkname4 "4. constraint"
		loc checkdesc4 "Check that numeric fields are constraint"
		loc checkname5 "5. other specify"
		loc checkdesc5 "Check that or_other is not used and manual osp fields are specified"
		loc checkname6 "6. group names"
		loc checkdesc6 "Check that group and repeat gropu names match in begin & end repeat"
		loc checkname7 "7. repeat vars"
		loc checkdesc7 "Check that fields in repeat group are properly suffixed"
		loc checkname8 "8. choices"
		loc checkdesc8 "Check for duplicates in choices list"

		* import information about xls form
		if `export' {
			import	excel using "`using'", sh("settings") first allstr clear
		
			postfile `summ' str23 field str100 (value comment) using "`summary'"

			* populate information from settings sheet into summary dataset

			post `summ' ("") 				("") 					("")
			post `summ' ("Form Details") 	("") 					("")
			post `summ' ("") 				("") 					("")
			post `summ' ("filename") 		("`filename'") 			("")
			post `summ' ("Form Title") 		("`=form_title[1]'") 	("")
			
			post `summ' ("Form ID") 				("`=form_id[1]'") 			("")
			post `summ' ("Form Definition Version") ("`=version[1]'") 			("")

			post `summ' ("Number of Languages") 	("") 						("")
			post `summ' ("Default Language") 		("`=default_language[1]'") 	("")

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

			post `summ' ("check") 			("description")		("result")
			post `summ' ("`checkname0'") 	("`checkdesc0'") 	("")	
			post `summ' ("`checkname1'") 	("`checkdesc1'") 	("")
			post `summ' ("`checkname2'") 	("`checkdesc2'") 	("")
			post `summ' ("`checkname3'") 	("`checkdesc3'") 	("")
			post `summ' ("`checkname4'") 	("`checkdesc4'") 	("")
			post `summ' ("`checkname5'") 	("`checkdesc5'") 	("")
			post `summ' ("`checkname6'") 	("`checkdesc6'") 	("")
			post `summ' ("`checkname7'") 	("`checkdesc7'") 	("")
			post `summ' ("`checkname8'")	("`checkdesc8'")	("")

			postclose `summ'

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

		count if (inlist(type, "start", "end") | inlist(calculation, "duration()")) & (lower(disabled) ~= "yes")
		
		postfile 	`ch0' str50 (variable disabled) str100 comment using "`check0'"
		count if type == "start"
		if `r(N)' == 0 post `ch0' ("starttime") ("") ("starttime field is missing")
		else {
			count if type == "start" & lower(disabled) == "yes" 
			if `r(N)' > 0 post `ch0' ("starttime") ("yes") ("starttime field is disabled")
		}

		count if type == "end"
		if `r(N)' == 0 post `ch0' ("endtime") ("") ("endtime field is missing")
		else {
			count if type == "end" & lower(disabled) == "yes" 
			if `r(N)' > 0 post `ch0' ("endtime") ("yes") ("endtime field is disabled")
		}
	
		count if type == "calculate" & calculation == "duration()"
		if `r(N)' == 0 post `ch0' ("duration") ("") ("duration field is missing")
		else {
			count if type == "calculate" & calculation == "duration()" & lower(disabled) == "yes"
			if `r(N)' > 0 post `ch0' ("duration") ("yes") ("duration field is disabled")
		}
		
		postclose `ch0'

		noi header, checkname("`checkname0'") checkdesc("`checkdesc0'") 		

		use "`check0'", clear

		loc check0_cnt `=_N'
	
		
		if `check0_cnt' > 0 {
			
			noi disp  	"{p}The following fields are missing or disabled. Please note that these field are " ///
					"required for IPA Data Quality Checks{p_end}"
			noi disp

			noi list, noobs abbrev(32) sep(0) table

			noi disp 
			
		}
		else noi disp "no issues identified"


		* Check 1: Check that variable lengths do not exceed 22 chars
		* -----------------------------------------------------------
		
		noi header, checkname("`checkname1'") checkdesc("`checkdesc1'")

		use "`survey'", clear
		keep if length(name) > 22

		loc check1_cnt `=_N'
		
		if `check1_cnt' > 0 {

			noi disp 			"{p}The following fields have names with character length greater than 22.{p_end}"
			noi disp

			gen char_length = length(name)
			keep row type name char_length 

			order row type name char_length

			save "`check1'"

			noi list, noobs abbrev(32) sep(0) table
			noi disp
			
		}
		else noi disp "no issues identified"
			
		* Check 2: Check for disabled field and readonly field
		* ---------------------------------
		
		noi header, checkname("`checkname2'") checkdesc("`checkdesc2'")

		use "`survey'", clear
		keep if regexm(lower(disabled), "yes") | regexm(lower(readonly), "yes") 

		loc check2_cnt `=_N'
		
		if `check2_cnt' > 0 {
		
			noi disp 			"{p}The following fields have been disabled.{p_end}"
			noi disp

			keep row type name disabled readonly
			sort row type name disabled readonly

			noi list, noobs abbrev(32) sep(0)
			noi disp
			
			save "`check2'"
		}
		else noi disp "no issues identified"
		
		* Check 3: Check requirement settings. Flag the following:
			* field is not required & is ("integer|text|date|time|select")
			* field is required & is a note 
			* field is required & is readonly
			* field is required & has appearance type label
		* ----------------------------------
		
		noi header, checkname("`checkname3'") checkdesc("`checkdesc3'")

		use "`survey'", clear
		keep if (lower(required) ~= "yes" & ///
				(regexm(lower(type), "^(select_)|^(date)") | inlist(lower(type), "integer", "text", "time", "audio", "video", "file", "image")) ///
				& lower(appearance) ~= "label") | ///
				(lower(required) == "yes" & (type == "note" | lower(readonly) == "yes"))

		loc check3_cnt `=_N'

		if `check3_cnt' > 0 {

			noi disp 			"{p}The following fields have issues with requirement.{p_end}"

			keep row type name label* appearance readonly required
			order row type name label* appearance readonly required
			
			* label each issue type
			gen comment = 	cond(lower(required) ~= "yes" & regexm(type, "integer|(text)$|date|time|select_"), "field is not required", ///
							cond(lower(required) == "yes" & type == "note", "required note field", ///
								"required & read only"))
			
			noi list row type name appearance readonly required comment, noobs abbrev(32) sep(0)
			noi disp

			save "`check3'"

		}
		else noi disp "no issues identified"

		* Check 4: Constraint Messages
			* check that numeric fields are constraint
			* check that text fields are constraint if using appearance type numbers, numbers_phone
			* check that constrained fields have constraint messages
		* ----------------------------

		noi header, checkname("`checkname4'") checkdesc("`checkdesc4'")

		use "`survey'", clear

		unab mcm: constraintmessage*
		loc mcm_cnt = wordcount("`mcm'") 

		egen nmcm_cnt = rownonmiss(constraintmessage*), strok
		keep if (nmcm_cnt < `mcm_cnt' & inlist(type, "integer", "decimal")) 				| ///
				(nmcm_cnt < `mcm_cnt' & type == "text" & regexm(appearance, "numbers")) 	| ///
				(nmcm_cnt < `mcm_cnt' & !missing(constraint)) 

		loc check4_cnt `=_N'

		if `check4_cnt' > 0 {

			noi disp 			"{p}The following fields are missing constraint or constraint message.{p_end}"

			keep row type name label* appearance constraint constraintmessage* nmcm_cnt
			order row type name label* appearance constraint constraintmessage* nmcm_cnt

			* label each issue type
			gen comment = cond(!missing(constraint) & nmcm_cnt < `mcm_cnt', "missing constraint message", "missing constraint")
			
			noi list row type name appearance constraint constraintmessage* comment, noobs abbrev(32) sep(0)
			noi disp

			drop nmcm_cnt

			save "`check4'"

		}
		else noi disp "no issues identified"
		
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

		noi header, checkname("`checkname5'") checkdesc("`checkdesc5'")

		use "`survey'", clear
		gen choice_other = 0
		if wordcount("`other_list'") > 0 {

			* mark all fields using choices with other specify
			foreach item in `other_list' {
				replace choice_other = 1 if word(type, 2) == "`item'" & regexm(type, "^(select_one)|^(select_multiple)")
			}
		}


		* flag fields with other specify
		keep row type name label* relevance choice_filter choice_other
		gen child_index = ""
		gen child_name 	= ""
		gen child_row 	= .

		
		getrow if choice_other, loc (indexes)
		
		if "`indexes'" ~= "" {
			foreach index of numlist `indexes' {
				
				loc parent = name[`index']
				getrow if regexm(relevance, "{`parent'}") & regexm(relevance, "`other'"), loc(child_index)
				if "`child_index'" ~= "" {	
					replace child_name = name[`child_index'] in `index'
					replace child_row = row[`child_index'] in `index'
				}
			}
		}

		keep if (regexm(type, "or_other$") & wordcount(type) == 3) | ///
				(missing(child_name) & choice_other) | ///
				(!missing(child_name) & (child_row < row) & choice_other)

		if `=_N' > 0 {

			noi disp 			"{p}The following fields have missing other specify fields or use the or_other syntax.{p_end}"

			* generate comments for each issue
			gen comment = cond(regexm(type, "or_other$") & wordcount(type) == 3, "or_other syntax used", ///
						  cond(missing(child_name) & choice_other, "missing other specify field", ///
						  "other specify field [" + child_name + "] on row " + string(child_row) + " comes before parent field"))
			
			keep row type name label* choice_filter comment	
			order row type name label* choice_filter comment

			noi list row type name choice_filter comment, noobs abbrev(32) sep(0)
			noi disp

			loc check5_cnt `=_N'

			save "`check5'"
		}
		else {
			loc check5_cnt 0
			noi disp "no issues identified"
		}

		* ---------------------------------------------------------------
		* Check and pair up group names
		* ---------------------------------------------------------------

		use "`survey'", clear

		keep if inlist(type, "begin group", "end group", "begin repeat", "end repeat")

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
			keep row type name label* begin_row begin_fieldname end_row end_fieldname
			save "`grouppairs'"
		}

		* check 6: check group names
		*----------------------------

		noi header, checkname("`checkname6'") checkdesc("`checkdesc6'")

		count if inlist(type, "begin group", "end group", "begin repeat", "end repeat")

		if `r(N)' > 0 {
			keep if begin_fieldname ~= end_fieldname

			loc check6_cnt `=_N'

			if `check6_cnt' > 0 {

				noi disp 			"{p}The following following groups have different names and begin and end.{p_end}"			

				order type begin_row begin_fieldname end_row end_fieldname
				noi list type begin_row begin_fieldname end_row end_fieldname, noobs abbrev(32) sep(0)
				noi disp

				save "`check6'"
			}
			else noi disp "no issues identified"

		}
		
		* check 7: Repeat group vars
		*---------------------------

		noi header, checkname("`checkname7'") checkdesc("`checkdesc7'")

		use "`survey'", clear
		count if type == "begin repeat"

		loc rpt_cnt `r(N)'

		if `r(N)' > 0 {

			merge 1:1 row using "`grouppairs'", nogen keepusing(begin_row begin_fieldname end_row end_fieldname)

			* mark out all variables in 
			gen rpt_field	= 0
			gen rpt_group 	= ""
			getrow if  type == "begin repeat", loc (indexes)
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
			gen sheet 		= "/"

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
						replace sheet = sheet + "survey/" if regexm(`var', "{`rvar'}") & rpt_group != "`rvar_group'"
				}				

			}

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

			keep if rpt_flag
			loc ch_cnt `=_N'
			
			if `ch_cnt' > 0 {

				replace rpt_flagvar = itrim(trim(subinstr(rpt_flagvar, "/", " ", .)))
					replace rpt_flagcol = itrim(trim(subinstr(rpt_flagcol, "/", " ", .)))

				split rpt_flagvar
				split rpt_flagcol

				drop rpt_flagcol rpt_flagvar

				reshape long rpt_flagvar rpt_flagcol, i(row) j(instance)
				
				* save values in locals
				forval i = 1/`ch_cnt' {
					loc list_name`i' 	= list_name[`i']
					loc rpt_flagvar`i'	= rpt_flagvar[`i']
					loc rpt_flagcol`i' 		= rpt_flagcol[`i']
				}


			}

			* check and flag issues from choices sheet
			use "`check7'", clear

			forval i = 1/`ch_cnt' {

				loc list_name 	= "`list_name`i''"
				loc flagvar 	= "`rpt_flagvar`i''"
				loc flagcol 	= "`rpt_flagcol`i''"

				getrow if name == "`flagvar'", loc (index)
				loc rvar_group = rpt_group[`index']

				replace rpt_flag = 1 if (regexm(type, "select_one `list_name'") 	| ///
										regexm(type, "select_multiple `list_name'")) ///
										&  rpt_group != "`rvar_group'"

				replace rpt_flagvar = rpt_flagvar + "`flagvar'/" if (regexm(type, "select_one `list_name'") 	| ///
																	regexm(type, "select_multiple `list_name'"))  ///
																	&  rpt_group != "`rvar_group'"


				replace rpt_flagcol = rpt_flagcol + "`flagcol'/" if (regexm(type, "select_one `list_name'") 	| ///
																	regexm(type, "select_multiple `list_name'"))  ///
																	&  rpt_group != "`rvar_group'"

				replace sheet = sheet + "choices/" if (regexm(type, "select_one `list_name'") 	| ///
																	regexm(type, "select_multiple `list_name'"))  ///
																	&  rpt_group != "`rvar_group'"
			}

			keep row sheet type name label* rpt_group rpt_flag rpt_flagvar rpt_flagcol

			keep if rpt_flag

			save "`check7'", replace

		
			if `=_N' > 0 {
				
				replace rpt_flagvar = itrim(trim(subinstr(rpt_flagvar, "/", " ", .)))
				replace rpt_flagcol = itrim(trim(subinstr(rpt_flagcol, "/", " ", .)))
				replace sheet 		= itrim(trim(subinstr(sheet, "/", " ", .)))

				split rpt_flagvar
				split rpt_flagcol
				split sheet

				drop rpt_flagcol rpt_flagvar sheet

				reshape long rpt_flagvar rpt_flagcol sheet, i(row) j(instance)
				drop instance

				drop if missing(sheet)

				ren rpt_flagcol column
				ren rpt_flagvar repeat_field

				noi disp 			"{p}The following following fields contain repeat group fields that have been used illegally.{p_end}"			

				order sheet row type name label* repeat_field column

				noi list sheet row type name repeat_field column, noobs abbrev(32) sep(0)
				noi disp

				loc check7_cnt `=_N'

				save "`check7'", replace
			}
			else {
				loc check7_cnt 0
				noi disp "no issues identified"
			}
		}
		else {
			loc check7_cnt 0
			noi disp "no repeat groups"
		}

		* check 8: choices list: Check for duplicates in choice list
		*----------------------

		noi header, checkname("`checkname8'") checkdesc("`checkdesc8'")

		use "`choices'", clear

		sort list_name row

		unab cols: value label*

		loc i = 1
		foreach col of varlist `cols' {
			duplicates tag list_name `col', gen (dup_`i')
			loc ++i
		}

		egen keep = rowtotal(dup_*) 

		drop if !keep
		
		loc i = 1
		foreach col of varlist `cols' {
			count if dup_`i'
			loc col`i'_cnt `r(N)'
			if `col`i'_cnt' > 0 {
				getrow if dup_`i', loc (col`i')
				loc col`i' = subinstr(trim(itrim("`col`i''")), " ", ",", .)
			}
			loc ++i
		}

		if `=_N' > 0 {

			noi disp "{p}The following following choice list contain duplicates.{p_end}"

			keep row value list_name label*
			order row value list_name label*
			noi list row list_name value label*, noobs abbrev(32) sepby(list_name)

			loc check8_cnt `=_N'

			save "`check8'"
		}
		else {
			loc check8_cnt 0
			noi disp "no issues identified"
		} 

		* export data
		* -----------

		if "`outfile'" ~= "" {

			use "`summary'", clear

			* replace number of languages
			replace value = "`labcount'" in 8
			loc  summ_cols ""
			
			forval i = 0/8 {

				if `i' == 7 & `rpt_cnt' == 0 {
					replace comment = "no repeat groups" in 23
					loc summ_cols "`summ_cols' -1" 
				}

				else if `check`i'_cnt' == 0 {
					replace comment = "no issues identified" in `=16+`i'' 
					loc summ_cols "`summ_cols' 0"
				}
				else {
					replace comment = "`check`i'_cnt' issues identified" in `=16+`i''
					loc summ_cols "`summ_cols' 1"
				}

				loc exp_name`i' = field[`=16+`i'']

			}

			loc summ_cols = subinstr(trim(itrim("`summ_cols'")), " ", ",", .)
			
			* export and format summary sheet
				export excel using "`outfile'", sheet("summary") replace cell(B1)
				gen _a = "", before(field)
				mata: adjust_column_width("`outfile'", "summary")
				mata: format_summary("`outfile'", "summary", (`summ_cols'))
				

			* export remaining sheets
		
			forval i = 0/8 {
				if `check`i'_cnt' > 0 {
					use "`check`i''", clear
					export excel using "`outfile'", sheet("`exp_name`i''") first(var)
					mata: adjust_column_width("`outfile'", "`exp_name`i''")
					mata: add_borders("`outfile'", "`exp_name`i''")
				}
			}

			* add color flags to check 8
			if `check8_cnt' > 0 {
				loc cols_cnt = wordcount("`cols'")
				forval i = 1/`cols_cnt' {
					if "`col`i''" ~= "" mata: add_flags("`outfile'", "`exp_name8'", 1 + `i', (`col`i''))
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
		cap drop index
end

* program to display headers for each check
program define header
	syntax, checkname(string) checkdesc(string) [main]
		* if main is specified use 2 lines of stars
		noi disp "{hline}"
		loc checkname = upper("`checkname'")
		noi disp "CHECK #`checkname'"
		noi disp "`checkdesc'"
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
				replace `generate' = substr("`text'", `spos', `start' - `spos' + 1) in `row'
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
void adjust_column_width(string scalar filename, string scalar sheetname)
{

	class xl scalar b
	real scalar column_width, columns, ncols, nrows, i, colmaxval

	ncols = st_nvar()
	nrows = st_nobs()

	b = xl()

	b.load_book(filename)
	b.set_sheet(sheetname)
	b.set_mode("open")

	for (i = 1;i <= ncols;i ++) {
		
		namelen = strlen(st_varname(i))		
		collen = colmax(strlen(st_sdata(., i)))
		
		if (st_varname(i) == "_a") {
			column_width = 1
		} 
		else if (namelen > collen) {
			column_width = namelen + 3
		}
		else {
			column_width = collen + 3
		}
		
		if (column_width > 101) {
			column_width = 101
		}	
		
		b.set_column_width(i, i, column_width)

		if (column_width >= 100) {
			b.set_text_wrap((1, nrows), (i, i), "on")
		}

	}

		b.close_book()

}

void format_summary(string scalar filename, string scalar sheetname, numeric vector colors)
{

	class xl scalar b

	b = xl()

	b.load_book(filename)
	b.set_sheet(sheetname)
	b.set_mode("open")

	b.set_font_bold((2, 24), (2, 2), "on")
	b.set_font_bold((15, 15), (2, 4), "on")

	b.set_font_italic((4, 11), (2, 2), "on")
	b.set_font_italic((16, 24), (2, 2), "on")

	b.set_sheet_merge(sheetname, (2, 2), (2, 3))
	b.set_sheet_merge(sheetname, (13, 13), (2, 4))

	b.set_font((2, 2), (2, 2), "calibri", 14)
	b.set_font((13, 13), (2, 2), "calibri", 14)

	b.set_horizontal_align((2, 2), (2, 3), "center")
	b.set_horizontal_align((13, 13), (2, 3), "center")

	b.set_row_height(3, 3, 10)
	b.set_row_height(12, 12, 10)
	b.set_row_height(14, 14, 10)

	b.set_left_border((4, 11), (2, 4), "thin")
	b.set_left_border((15, 24), (2, 5), "thin")

	b.set_top_border((4, 4), (2, 3), "thin")
	b.set_top_border((15, 16), (2, 4), "thin")

	b.set_bottom_border((11, 11), (2, 3), "thin")
	b.set_bottom_border((24, 24), (2, 4), "thin")


	encrypted = st_sdata(10, "value")
	if (encrypted == "No") {
		b.set_fill_pattern(10, 3, "solid", "lightpink")
	} 

	for (i = 1;i <= 9;i ++) {
		if (colors[i] == -1) {
			b.set_fill_pattern(15 + i, 4, "solid", "lightyellow")
		}
		else if (colors[i] == 0) {
			b.set_fill_pattern(15 + i, 4, "solid", "lightgreen")
		}
		else {
			b.set_fill_pattern(15 + i, 4, "solid", "lightpink")
		}
	}

	b.close_book()

}

void add_borders(string scalar filename, string scalar sheetname)
{
	class xl scalar b
	real scalar ncols, nrows

	b = xl()

	b.load_book(filename)
	b.set_sheet(sheetname)
	b.set_mode("open")

	ncols = st_nvar()
	nrows = st_nobs() + 1

	for (i = 1;i <= ncols; i++) {
		b.set_font_bold((1, 1), (1, ncols), "on")
		b.set_bottom_border((1, 1), (1, ncols), "thin")
		b.set_bottom_border((nrows, nrows), (1, ncols), "thin")
		b.set_right_border((1, nrows), (ncols, ncols), "thin")
	}

	b.close_book()

}

void add_flags(string scalar filename, string scalar sheetname, numeric scalar column, numeric vector rows) 
{
	class xl scalar b
	real scalar ncols, nrows

	b = xl()

	b.load_book(filename)
	b.set_sheet(sheetname)
	b.set_mode("open")

	ncols = st_nvar()
	nrows = st_nobs() + 1

	for (i = 1;i <= length(rows);i ++) {
		b.set_fill_pattern(rows[i] + 1, column, "solid", "lightpink")
	}

	b.close_book()
}

end



