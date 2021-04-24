*! version 1.2
*! Ishmail Azindoo Baako (IPA) March, 2021 

* Stata program to review SurveyCTO xls form and export issues to an excel sheet
* This program checks for the following type of issues:
	** checks for errors that cannot be detected by scto server or 
	** for compliance with IPA standards

version 14.0
program define ipacheckscto
	#d;
	syntax using/ 
				[, 
				OUTfile(str) 
				OTHer(numlist max = 1) 
				DONTKnow(numlist max = 1) 
				REFuse(numlist max = 1)
				replace
				]
			;
	#d cr
	
	qui {

		tempname  summ_hand chk1_hand
		tempfile  survey choices repeat repeat_long grouppairs
		tempfile  summary chk1 chk2 chk3 chk4 chk5 chk6 chk7 chk8 chk9 chk10

		* set dummy for exporting to excel

			if "`outfile'" ~= "" 	loc export 1
			else 					loc export 0 

		* add file extension if needed
		if !regexm("`using'", "\.xlsx$|\.xls$") loc using "`using'.xlsx"

		* set export file
		if `export' {
			* include .xlsx extension by default if not included in filename
			if !regexm("`outfile'", "\.xlsx$|\.xls$") loc outfile "`outfile'.xlsx"
			* check that use didnt specify using as outfile
			if "`using'" == "`outfile'" {
				disp as err "options using and outfile cannot have the same {help filename:filename}"
				ex 602
			} 

			* check if file exist and warn user
			cap confirm file "`outfile'"
			if !_rc & "`replace'" == "" {
				disp as err "file `outfile' already exist. Specify replace to replace file"  
				ex 602
			}
			
			* Remove old outfile. This is to minimize errors when exporting
			else if !_rc & "`replace'" ~= "" cap rm "`outfile'"
		}

		* save filenames in local
		loc filename = substr("`using'", -strpos(reverse(subinstr("`using'", "\", "/", .)), "/") + 1, .)
	
		* save checks and descriptions in locals to allow for easy updates

		loc chkname1 	"1. recommended fields"
		loc chkdesc1 	"Display status of recommended/meta-data fields in XLS form"
		loc chkname2 	"2. field names"
		loc chkdesc2 	"Flag issues in field names"
		loc chkname3 	"3. disabled, read only"
		loc chkdesc3 	"Flag fields that are disabled or marked as read only"
		loc chkname4 	"4. field requirements"
		loc chkdesc4 	"Flag fields with various requirement issues"
		loc chkname5 	"5. constraint"
		loc chkdesc5 	"Flag fields with constraint issues"
		loc chkname6 	"6. other specify"
		loc chkdesc6 	"Flag for issues with other specify"
		loc chkname7 	"7. dont know, refuse"
		loc chkdesc7 	"Flag field that do not allow for don't know and refuse to answer"
		loc chkname8 	"8. group names"
		loc chkdesc8 	"Flag groups with mismatch in group names at begin & end"
		loc chkname9 	"9. repeat vars"
		loc chkdesc9 	"Flag illegal use of repeat fields"
		loc chkname10 	"10. choices"
		loc chkdesc10 	"Flag duplicates or missing values/labels in choices"

		* import & prep settings sheet
		import	excel using "`using'", sh("settings") first allstr clear
		prep_data
		
		postfile `summ_hand' str50 (field) str100 (value comment) using "`summary'"

		* populate information from settings sheet into summary dataset

		post `summ_hand' ("") 				("") 					("")
		post `summ_hand' ("Form Details") 	("") 					("")
		post `summ_hand' ("") 				("") 					("")
		post `summ_hand' ("filename") 		("`filename'") 			("")

		loc formtitle 	 "`=form_title[1]'"
		post `summ_hand' ("Form Title") 	("`formtitle'") 		("")
		
		loc formid       "`=form_id[1]'"
		post `summ_hand' ("Form ID") 				 ("`formid'") 	("")

		loc formdef      "`=version[1]'"
		post `summ_hand' ("Form Definition Version") ("`formdef'") 	("")

		post `summ_hand' ("Number of Languages") 	("") 				("")

		loc default_lang "`=default_language[1]'"
		post `summ_hand' ("Default Language") 		("`default_lang'") 	("")

		if "`=public_key[1]'" ~= "" loc encrypted "Yes"
		else 						loc encrypted "No"

		post `summ_hand' ("Form Encrypted") 					("`encrypted'") ("")
		post `summ_hand' ("Number of Publishable fields")	("")			("")

		if "`=submission_url[1]'" ~= "" loc suburl 	"`=submission_url[1]'"
		else 							loc suburl 	"None"

		post `summ_hand' ("Submission URL") ("`suburl'") ("")

		* populate names and extra text for other inputs
		post `summ_hand' ("") 				("") ("")
		post `summ_hand' ("Check Summary") 	("") ("")
		post `summ_hand' ("") 				("") ("")

		post `summ_hand' ("check") 			("description")		("result")
		post `summ_hand' ("`chkname1'") 	("`chkdesc1'") 		("")	
		post `summ_hand' ("`chkname2'") 	("`chkdesc2'") 		("")
		post `summ_hand' ("`chkname3'") 	("`chkdesc3'") 		("")
		post `summ_hand' ("`chkname4'") 	("`chkdesc4'") 		("")
		post `summ_hand' ("`chkname5'") 	("`chkdesc5'") 		("")
		post `summ_hand' ("`chkname6'") 	("`chkdesc6'") 		("")
		post `summ_hand' ("`chkname7'") 	("`chkdesc7'") 		("")
		post `summ_hand' ("`chkname8'") 	("`chkdesc8'") 		("")
		post `summ_hand' ("`chkname9'")		("`chkdesc9'")		("")
		post `summ_hand' ("`chkname10'")	("`chkdesc10'")		("")

		postclose `summ_hand'

		* Import & prep Survey sheet
		import 	excel using "`using'", sheet("survey") first allstr clear 
		prep_data

		* drop other labels that may have been included in survey which are not used by scto
		cap unab label_drop: label_*
		if !_rc {
			drop `label_drop'
		}

		* count number of label variables
		unab labels : label*
		loc lab_cnt = wordcount("`labels'")

		* save survey data
		save "`survey'"

		* Display main headers
		
		noi disp
		noi disp
		noi disp _dup(120) "="
		noi disp 
		noi disp "{ul:Form details}" 
		noi disp 
		noi disp "Filename" 					_column(30) ": `filename'"
		noi disp "Form Title" 					_column(30) ": `formtitle'"
		noi disp "Form ID" 						_column(30) ": `formid'"
		noi disp "Form Definition Version" 		_column(30) ": `formdef'"
		noi disp "Number of languages"			_column(30) ": `lab_cnt'"
		noi disp "Default language"				_column(30) ": `default_lang'"
		
		if "`encrypted'" == "Yes" {
			noi disp "Form Encrypted"	_column(30) ": Yes"
			count if lower(publishable) == "yes"
			loc pub_cnt `r(N)'
			if `pub_cnt' > 0 noi disp "Number of Publishable fields"	_column(30) ": `pub_cnt'"
			else 		  	 noi disp "Number of Publishable fields"	_column(30) ": {red:`pub_cnt'}"		
		}
		else noi disp "Form Encrypted"	_column(30) ": {red:No}"
	
		noi disp
		noi disp _dup(120) "="
		
		
		* Check 1: Display status of recommended/meta-data fields in XLS form
		* ----------------------------------------------------------------
		
		* Save results in new postfile

		postfile `chk1_hand' str32 (field) str300 (description) str10 (status) str300(rows names) using "`chk1'"

		* Required:

		post `chk1_hand'	("") 			("") ("") ("") ("")
		post `chk1_hand'	("Required:") 	("") ("") ("") ("")
		post `chk1_hand'	("") 			("") ("") ("") ("")
		
		postresult if lower(type) == "start", field(starttime) handle(`chk1_hand') ///
			desc("Auto-records the date & time the survey was started")

		postresult if lower(type) == "end", field(endtime) handle(`chk1_hand') ///
			desc("Auto-records the date & time the survey was ended")

		postresult if lower(calculation) == "duration()", field(duration) handle(`chk1_hand') ///
			desc("Auto-records the duration (in secs) of the entire survey")

		post `chk1_hand'	("") 				("") ("") ("") ("")
		post `chk1_hand'	("Recommended:") 	("") ("") ("") ("")
		post `chk1_hand'	("") 				("") ("") ("") ("")

		postresult if lower(type) == "comments", field(comments) handle(`chk1_hand') ///
			desc("Allows users to enter comments associated with any field(s)")

		postresult if lower(type) == "text audit", field(text audit) handle(`chk1_hand') ///
			desc("Auto-record meta-data about how the form was filled out")

		postresult if lower(type) == "audio audit", field(audio audit) handle(`chk1_hand') ///
			desc("Audio-record some or all survey administration (invisibly)")

		postresult if lower(type) == "geopoint", field(geopoint) handle(`chk1_hand') ///
			desc("Collects GPS coordinates using the device's built-in GPS support")

		postresult if regexm(lower(type), "^sensor_statistic"), field(sensor_statistic) handle(`chk1_hand') ///
			desc("Summarize device sensor meta-data")

		postresult if regexm(lower(type), "^sensor_stream"), field(sensor_stream) handle(`chk1_hand') ///
			desc("Capture a stream of sensor meta-data")

		postclose `chk1_hand'	

		noi header, checkname("`chkname1'") checkdesc("`chkdesc1'") 		

		use "`chk1'", clear
		compress

		count if status == "missing"	
		loc chk1_cnt `r(N)'	
		
		noi disp
		noi disp "{bf:Required}:"
		noi list in 4/6, noobs abbrev(32) sep(0) table
		noi disp
		noi disp "{bf:Recommended}:"
		noi list in 10/l, noobs abbrev(32) sep(0) table

		save "`chk1'", replace

		* Check 2: Flag issues in field names
		* -----------------------------------------------------------
		
		noi disp
		noi header, checkname("`chkname2'") checkdesc("`chkdesc2'")
		
		use "`survey'", clear
		gen char_length 	= length(name)
		gen invalid_name	= regexm(name, "\.|\-")
		
		keep if char_length > 22 | invalid_name

		loc chk2_cnt `=_N'
		
		if `chk2_cnt' > 0 {

			gen issue = cond(char_length > 22 & invalid_name, "long & invalid varname", ///
						  cond(invalid_name, "invalid varname", "long varname"))

			keep row type name char_length issue

			order row type name char_length issue

			save "`chk2'"

			noi list, noobs abbrev(32) sep(0) table
			noi disp
			
		}
		else noi disp "no issues identified"
		
		* Check 3: Check for disabled and readonly field
		* ---------------------------------
		
		noi disp
		noi header, checkname("`chkname3'") checkdesc("`chkdesc3'")

		use "`survey'", clear
		keep if lower(disabled) == "yes" | lower(readonly) == "yes" 

		loc chk3_cnt `=_N'
		
		if `chk3_cnt' > 0 {

			keep row type name disabled readonly

			gen issue = cond(lower(disabled) == "yes" & lower(readonly) == "yes", "disabled & readonly", ///
						cond(lower(disabled) == "yes", "disabled", "readonly"))

			sort row type name disabled readonly issue
			order row type name disabled readonly issue

			noi list, noobs abbrev(32) sep(0)
			noi disp
			
			save "`chk3'"
		}
		else noi disp "no issues identified"
		
		* Check 4: Flag fields with various requirement issues:
			* field is not required & is ("integer|text|date|time|select")
			* field is required & is a note 
			* field is required & is readonly
			* field is required & has appearance type label
		* ----------------------------------

		noi disp
		noi header, checkname("`chkname4'") checkdesc("`chkdesc4'")
		
		use "`survey'", clear
		keep if (lower(required) ~= "yes" & ///
				(regexm(lower(type), "^(select_)|^(date)|^(geo)") | ///
				inlist(lower(type), "text", "integer", "decimal", "barcode", "time", "image", "audio", "video", "file")) ///
				& lower(appearance) ~= "label") | ///
				(lower(required) == "yes" & (type == "note" | lower(readonly) == "yes"))

		* exclude geopoint fields with a background appearance
		drop if lower(type) == "geopoint" & regexm(lower(appearance), "background")

		loc chk4_cnt `=_N'

		if `chk4_cnt' > 0 {

			* label each issue type
			gen issue = cond(lower(required) == "yes" & type == "note", "required note field", ///
						cond(lower(required) == "yes" & lower(readonly) == "yes", "required & read only", ///
						"not required"))

			keep row type name appearance readonly required issue
			order row type name appearance readonly required issue
			
			noi list, noobs abbrev(32) sep(0)
			noi disp

			save "`chk4'"

		}
		else noi disp "no issues identified"
		
		* Check 5: Flag fields with constraint issues
			* check that numeric fields are constraint
			* check that text fields are constraint if using appearance type numbers, numbers_phone
			* check that constrained fields have constraint messages
		* ----------------------------

		noi header, checkname("`chkname5'") checkdesc("`chkdesc5'")

		use "`survey'", clear

		unab mcm: constraintmessage*
		loc mcm_cnt = wordcount("`mcm'") 

		egen nmcm_cnt = rownonmiss(constraintmessage*), strok
		keep if (nmcm_cnt < `mcm_cnt' & inlist(type, "integer", "decimal")) 				| ///
				(nmcm_cnt < `mcm_cnt' & type == "text" & regexm(appearance, "numbers")) 	| ///
				(nmcm_cnt < `mcm_cnt' & !missing(constraint)) 

		loc chk5_cnt `=_N'
		
		if `chk5_cnt' > 0 {

			keep row type name appearance constraint constraintmessage* nmcm_cnt
			order row type name appearance constraint constraintmessage* nmcm_cnt

			* label each issue type
			gen issue = cond(!missing(constraint) & nmcm_cnt < `mcm_cnt', "missing constraint message", "missing constraint")
			
			noi list row type name appearance constraint constraintmessage* issue, noobs abbrev(32) sep(0)
			noi disp

			drop nmcm_cnt

			save "`chk5'"

		}
		else noi disp "no issues identified"
		
		* ---------------------------------------------------------------
		* Imort and prepare choices
		* ---------------------------------------------------------------

		import	excel using "`using'", sh("choices") first allstr clear
		
		* prepare data
		prep_data

		* drop other labels that may have been included in survey which are not used by scto
		cap unab label_drop: label_*
		if !_rc {
			drop `label_drop'
		}
		
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

		* make a list of list_name (s) with don't know
		if "`dontknow'" ~= "" {
			levelsof list_name if value == "`dontknow'", loc (dontknow_list) clean
		}

		* make a list list of list_name (s) with refuse to answer
 		if "`refuse'" ~= "" {
			levelsof list_name if value == "`refuse'", loc (refuse_list) clean
		}
		
		save "`choices'"

		* check 6: Flag for issues with other specify
			* check that or_other is not used with select_one | select_multiple
			* check that fields using choices with other specify have defined an osp field
		*-------------------------------------

		noi header, checkname("`chkname6'") checkdesc("`chkdesc6'")

		use "`survey'", clear
		if "`other'" ~= "" {
			gen choice_other = 0

			* mark all fields using choices with other specify
			if wordcount("`other_list'") > 0 {
				foreach item in `other_list' {
					replace choice_other = 1 if word(type, 2) == "`item'" & regexm(type, "^(select_one)|^(select_multiple)")
				}
			}
		
			* flag fields with other specify
			keep row type name relevance choice_filter choice_other
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
		}
		else {

			keep if (regexm(type, "or_other$") & wordcount(type) == 3) 
		}

		if `=_N' > 0 {

			* generate comments for each issue
			if "`other'" ~= "" {
				gen issue = cond(regexm(type, "or_other$") & wordcount(type) == 3, "or_other syntax used", ///
						  cond(missing(child_name) & choice_other, "missing other specify field", ///
						  "other specify field [" + child_name + "] on row " + string(child_row) + " comes before parent field"))
			}
			else gen issue = "or_other syntax used"
			
			keep row type name choice_filter issue	
			order row type name choice_filter issue

			noi list, noobs abbrev(32) sep(0)
			noi disp

			loc chk6_cnt `=_N'

			save "`chk6'"
		}
		else {
			loc chk6_cnt 0
			noi disp "no issues identified"
		}

		* check 7: Flag field that do not allow for don't know and refuse to answer
			* check that choice list includes dk, ref for select_one | select_multiple
			* check that integer, decimal & text fields with constraint allow 
		* ----------------------------------------------------------------------------

		noi header, checkname("`chkname7'") checkdesc("`chkdesc7'")

		if "`dontknow'" ~= "" | "`refuse'" ~= "" {
			use "`survey'", clear

			* Mark select fields that include don't know
			gen dontknow = 0
			if wordcount("`dontknow_list'") > 0 {
				foreach item in `dontknow_list' {
					replace dontknow = 1 if word(type, 2) == "`item'" & regexm(type, "^(select_one)|^(select_multiple)")
				}
			}
			gen refuse = 0
			if wordcount("`refuse_list'") > 0 {
				foreach item in `refuse_list' {
					replace refuse = 1 if word(type, 2) == "`item'" & regexm(type, "^(select_one)|^(select_multiple)")
				}
			}

			if "`dontknow'" ~= "" {
				replace dontknow = 1 if inlist(type, "integer", "decimal", "text") & ///
				regexm(subinstr(trim(itrim(constraint)), " ", "", .), "^\.=`dontknow'|or.*\.=`dontknow'|regex\(.*\-`dontknow'.*\)")

			}
			
			if "`refuse'" ~= "" {
				replace refuse = 1 if inlist(type, "integer", "decimal", "text") & ///
				regexm(subinstr(trim(itrim(constraint)), " ", "", .), "^\.=`refuse'|or.*\.=`refuse'|regex\(.*\-`refuse'.*\)")
			}
			
			keep if inlist(type, "integer", "decimal", "text") | regexm(type, "^select_one|^select_multiple")
			
			if "`dontknow'" ~= "" & "`refuse'" ~= "" drop if dontknow & refuse
			else if "`dontknow'" ~= "" drop if dontknow
			else drop if refuse

			drop if inlist(type, "integer", "decimal", "text") & missing(constraint)

			loc chk7_cnt `=_N'

			if `chk7_cnt' > 0 {
			
				keep row type name constraint dontknow refuse

				if "`dontknow'" ~= "" & "`refuse'" ~= "" {
					gen issue = cond(!dontknow & !refuse, "don't know & refuse not allowed", ///
							cond(!dontknow, "don't know not allowed", "refuse not allowed"))
				}
				else if "`dontknow'" == "" 	gen issue = "refuse not allowed"
				else 						gen issue = "don't know not allowed"

				drop dontknow refuse
				order row type name constraint issue

				compress
				noi list, noobs abbrev(32) sep(0)
				noi disp

				save "`chk7'", replace
			}
			else noi disp "no issues identified"

		}
		else {
			noi disp "not checked"
			loc chk7_cnt 0
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
			
			if `b_cnt' ~= `e_cnt' {
				disp as err "Invalid form: There are `b_cnt' begin group/repeat and `e_cnt' end group/repeat fields"
				keep type name
				
				count if inlist(type, "begin group", "end group")
				if mod(`r(N)', 2) keep if inlist(type, "begin group", "end group")
				else keep if inlist(type, "begin repeat", "end repeat")

				loc cont 1
				while `cont' {
					gen tag = (regexm(type, "^begin") & regexm(type[_n+1], "^end")) | ///
						  	  (regexm(type, "^end") & regexm(type[_n-1], "^begin"))	
						  	  
					drop if tag
					drop tag

					count if regexm(type, "^begin")	
					loc b_cnt `r(N)'
					count if regexm(type, "^end")
					loc e_cnt `r(N)'

					if !`b_cnt' | !`e_cnt' {
						noi disp "The following fields do not have a matching pair"
						noi list, noobs abbrev(32) sep(0)
						exit 198
					}
				}
			}

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

		* check 8: check group names
		*----------------------------

		noi header, checkname("`chkname8'") checkdesc("`chkdesc8'")
		
		if `=_N' > 0 {

			keep if begin_fieldname ~= end_fieldname

			loc chk8_cnt `=_N'

			if `chk8_cnt' > 0 {

				gen issue = "group names don't match"

				order type begin_row begin_fieldname end_row end_fieldname
				keep type begin_row begin_fieldname end_row end_fieldname issue

				noi list, noobs abbrev(32) sep(0)
				noi disp

				save "`chk8'"
			}
			else noi disp "no issues identified"

		}
		else {
			noi disp "{red: Not checked; XLS form has no groups or repeats}"
			loc chk8_cnt 0
		}
		
		* check 9: Repeat group vars
		*---------------------------

		noi header, checkname("`chkname9'") checkdesc("`chkdesc9'")
		
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
				replace rpt_group 	= cond(missing(rpt_group), "/`groupname'", rpt_group + "/" + "`groupname'") ///
										  in  `index'/`lastrow'
			}

			* foreach repeat group var, check if it was used outside the repeat group
			levelsof name if rpt_field, loc (rvars) clean

			** first remove functions that allow repeat vars

			#d;
			loc funcs
				"
				"join" 				"join-if"
				"sum"				"sum-if"
				"min" 				"min-if"
				"max"				"max-if"
				"rank-index" 		"rank-index-if"
				"indexed-repeat"
				"count"				"count-if"
				"
				;
			#d cr

			foreach var of varlist appearance constraint relevance calculation repeat_count {

				replace `var' = subinstr(`var', "$", "#", .)
				replace `var' = subinstr(`var', char(34), char(39), .)

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
						replace rpt_flag 	= 1 if regexm(`var', "{`rvar'}") & !regexm(rpt_group, "`rvar_group'")
						replace rpt_flagvar = rpt_flagvar + "`rvar'/" if regexm(`var', "{`rvar'}") & !regexm(rpt_group, "`rvar_group'")
						replace rpt_flagcol = rpt_flagcol + "`var'/" if regexm(`var', "{`rvar'}") & !regexm(rpt_group, "`rvar_group'")
						replace sheet = sheet + "survey/" if regexm(`var', "{`rvar'}") & !regexm(rpt_group, "`rvar_group'")
				}				

			}
		
			save "`chk9'"
			
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
			use "`chk9'", clear
			
			forval i = 1/`ch_cnt' {

				loc list_name 	= "`list_name`i''"
				loc flagvar 	= "`rpt_flagvar`i''"
				loc flagcol 	= "`rpt_flagcol`i''"

				getrow if name == "`flagvar'", loc (index)
				loc rvar_group = rpt_group[`index']

				replace rpt_flag = 1 if (regexm(type, "select_one `list_name'") 	| ///
										regexm(type, "select_multiple `list_name'")) ///
										& !regexm(rpt_group, "`rvar_group'")

				replace rpt_flagvar = rpt_flagvar + "`flagvar'/" if (regexm(type, "select_one `list_name'") 	| ///
																	regexm(type, "select_multiple `list_name'"))  ///
																	&  !regexm(rpt_group, "`rvar_group'")


				replace rpt_flagcol = rpt_flagcol + "`flagcol'/" if (regexm(type, "select_one `list_name'") 	| ///
																	regexm(type, "select_multiple `list_name'"))  ///
																	&  !regexm(rpt_group, "`rvar_group'")

				replace sheet = sheet + "choices/" if (regexm(type, "select_one `list_name'") 	| ///
																	regexm(type, "select_multiple `list_name'"))  ///
																	&  !regexm(rpt_group, "`rvar_group'")
			}

			keep row sheet type name rpt_group rpt_flag rpt_flagvar rpt_flagcol

			keep if rpt_flag

			save "`chk9'", replace

		
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

				gen issue = "illegal use of repeat field"

				ren rpt_flagcol column
				ren rpt_flagvar repeat_field

				order sheet row type name column repeat_field issue
				keep sheet row type name column repeat_field issue

				noi list, noobs abbrev(32) sepby(row)
				noi disp

				loc chk9_cnt `=_N'

				save "`chk9'", replace
			}
			else {
				loc chk9_cnt 0
				noi disp "no issues identified"
			}
		}
		else {
			loc chk9_cnt 0
			noi disp "{red: Not checked; XLS form has no repeat groups}"
		}
		
		* check 10: choices list: Check for duplicates in choice list
		*----------------------

		noi header, checkname("`chkname10'") checkdesc("`chkdesc10'")

		use "`choices'", clear

		sort list_name row
		keep list_name value label* row

		unab labs: label*
		loc first_lab = word("`labs'", 1) 
		
		save "`choices'", replace


		* Check value column

		keep list_name value `first_lab' row
		duplicates tag list_name value, gen (dup)
		keep if dup
		drop dup
		gen column = "value"
		if "`first_lab'" ~= "" ren `first_lab' label
		save "`chk10'", replace emptyok

		* check each label column
		foreach lab in `labs' {
			use "`choices'", clear
			keep list_name value `lab' row
			duplicates tag list_name `lab', gen (dup)
			keep if dup
			gen column = "`lab'"
			drop dup
			ren `lab' label
			append using "`chk10'"
			save "`chk10'", replace
		}

		loc chk10_cnt `=_N'

		if `chk10_cnt' > 0 {
			sort column list_name row

			gen issue = cond(column == "value", "duplicate value", cond(missing(label), "missing label", "duplicate label"))
			order list_name value label row column
			gen seperator = list_name + "/" + column + "/" + issue 
			sort list_name issue column row
			noi list list_name value label row column issue, noobs abbrev(32) sepby(seperator)

			drop seperator
			save "`chk10'", replace
		}
		else {
			noi disp "no issues identified"
		}

		* export data
		* -----------

		if "`outfile'" ~= "" {

			use "`summary'", clear

			* replace number of languages
			replace value = "`lab_cnt'" if field == "Number of Languages"
			replace value = "`pub_cnt'" if field == "Number of Publishable fields"

			loc  summ_cols ""

			forval i = 1/10 {

				if `i' == 9 & `rpt_cnt' == 0 {
					replace comment = "no repeat groups" in 25
					loc summ_cols "`summ_cols' -1" 
				}
				else if `chk`i'_cnt' == 0 {
					replace comment = "no issues identified" in `=16+`i'' 
					loc summ_cols "`summ_cols' 0"
				}
				else {
					replace comment = "`chk`i'_cnt' issues identified" in `=16+`i''
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
		
			forval i = 1/10 {
				if `chk`i'_cnt' > 0 {
					use "`chk`i''", clear
					export excel using "`outfile'", sheet("`exp_name`i''") first(var)
					mata: adjust_column_width("`outfile'", "`exp_name`i''")
					mata: add_borders("`outfile'", "`exp_name`i''")
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
		noi disp _dup(120) "-"
		loc checkname = upper("`checkname'")
		noi disp "CHECK #`checkname'"
		noi disp "`checkdesc'"
		noi disp _dup(120) "-"
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


* program for check 1: Required fields
program define postresult

	syntax if, field(string) handle(string) desc(string)
	count `if'
	loc field_cnt `r(N)'
	loc status = cond(`field_cnt' == 0, "missing", "included")
	if `field_cnt' > 0 {
		getrow `if', loc (rows)
		loc rows = subinstr(trim(itrim("`rows'")), " ", ",", .)
	}
	levelsof name `if', loc (names) clean sep(",")

	post `handle' 	("`field'") ("`desc'") ("`status'") ("`rows'") ("`names'")

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

	b.set_font_bold((2, 26), (2, 2), "on")
	b.set_font_bold((16, 16), (2, 4), "on")

	b.set_font_italic((4, 12), (2, 2), "on")
	b.set_font_italic((17, 26), (2, 2), "on")

	b.set_sheet_merge(sheetname, (2, 2), (2, 3))
	b.set_sheet_merge(sheetname, (14, 14), (2, 4))

	b.set_font((2, 2), (2, 2), "calibri", 14)
	b.set_font((14, 14), (2, 2), "calibri", 14)

	b.set_horizontal_align((2, 2), (2, 3), "center")
	b.set_horizontal_align((14, 14), (2, 3), "center")

	b.set_row_height(3, 3, 10)
	b.set_row_height(13, 13, 10)
	b.set_row_height(15, 15, 10)

	b.set_left_border((4, 12), (2, 4), "thin")
	b.set_left_border((16, 26), (2, 5), "thin")

	b.set_top_border((4, 4), (2, 3), "thin")
	b.set_top_border((16, 17), (2, 4), "thin")

	b.set_bottom_border((12, 12), (2, 3), "thin")
	b.set_bottom_border((26, 26), (2, 4), "thin")

	b.set_fill_pattern((16, 16), (2, 4), "solid", "231 230 230")

	encrypted = st_sdata(10, "value")
	if (encrypted == "No") {
		b.set_fill_pattern(10, 3, "solid", "lightpink")
	} 

	for (i = 1;i <= 10;i ++) {
		if (colors[i] == -1) {
			b.set_fill_pattern(16 + i, 4, "solid", "lightyellow")
		}
		else if (colors[i] == 0) {
			b.set_fill_pattern(16 + i, 4, "solid", "lightgreen")
		}
		else {
			b.set_fill_pattern(16 + i, 4, "solid", "lightpink")
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

	b.set_fill_pattern((1, 1), (1, ncols), "solid", "231 230 230")

	b.close_book()

}

void add_flags(string scalar filename, string scalar sheetname, numeric scalar column, numeric vector rows, numeric scalar openxl) 
{
	class xl scalar b
	real scalar ncols, nrows

	b = xl()

	b.load_book(filename)
	b.set_sheet(sheetname)

	if (openxl == 1) { 
		b.set_mode("open")
	}

	ncols = st_nvar()
	nrows = st_nobs() + 1

	for (i = 1;i <= length(rows);i ++) {
		b.set_fill_pattern(rows[i] + 1, column, "solid", "lightpink")
	}

	if (openxl == 0) {
		b.close_book()
	}
}

end



