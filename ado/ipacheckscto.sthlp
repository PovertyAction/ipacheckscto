{smcl}
{* *! version 1.2 Innovations for Poverty Action Mar2021}{...}
{title:Title}

{phang}
{cmd:ipacheckscto} {hline 2} Check SurveyCTO XLS form for possible issues


{marker syntax}{...}
{title:Syntax}

{p 4}
{cmd:ipacheckscto} {cmd:using} {it:{help filename}} [{cmd:,} {it:{help ipacheckscto##options:options}} ]


{synoptset 37}{...}
{marker options}{...}
{synopthdr :options}
{synoptline}
{* Using -help ca postestimation- as a template.}{...}
{p2coldent: {opth out:file(filename)}}Output results to excel filename{p_end}
{p2coldent: {opth oth:er(integer)}}Specify integer value for other specify{p_end}
{p2coldent: {opth dontk:now(integer)}}Specify integer value for don't know{p_end}
{p2coldent: {opth ref:use(integer)}}Specify integer value for refuse to answer{p_end}
{p2coldent: {opt replace}}Overwrite existing outfile{p_end}


{marker description}{...}
{title:Description}

{pstd}
{cmd:ipacheckscto} checks your SurveyCTO xls form for issues that may not be flagged by the SurveyCTO server. 
These include checking for bugs as well as commonn programming protocols and errors. {cmd:ipacheckscto} is 
designed to complement the SurveyCTO server debug tool and should therefore be used after the XLS forms has 
already been accepted by the SurveyCTO server.

{marker optionsdesc}{...}
{title:Options}

{phang}
{opt "outfile(filename)"} Specifies the filename of the excel file where the results will be exported. The 
filename must include the extension {cmd:.xls} or {cmd:.xlsx}. The default is to display issues on the Stata 
result window. 

{phang}
{opt "other(integer)"} Specifies the integer value for other specify option. If this option is specified, 
{cmd:ipacheckscto} will flag {bf:select_one/select_multiple} fields which (1) include other specify option
in their choice list but are missing an other specify field;(2) come after the other specify field; and (3) use
the or_other syntax. The default is to only check for the or_other syntax. 

{phang}
{opt "dontknow(integer)"} Specifies the integer value for don't know option. If this option is specified, 
{cmd:ipacheckscto} will flag any field that use a choice list that does not include an option for don't know. 

{phang}
{opt "refuse(integer)"} Specifies the integer value for refuse option. If this option is specified, 
{cmd:ipacheckscto} will flag any field that use a choice list that does not include an option for refuses to answer.

{phang}
{cmd:replace} Overwrites the outfile if it already exist. Default is to report an error if outfile already exist. 


{marker examples}{...}
{title:Examples}

{pstd}
Check a SurveyCTO XLS form and display results on Stata window. 

	{cmd}. ipacheckscto using "Bontanga Baseline.xlsx"{text}

{pstd}
Check a SurveyCTO XLS form and display results on Stata window. Include checks for other specify with the value 
of -666 

	{cmd}. ipacheckscto using "Bontanga Baseline.xlsx", other(-666){text}

{pstd}
Check a SurveyCTO XLS form and display results on Stata window. Include checks for other specify (-666), dontknow(-999)
and refuse to answer (-888).  

	{cmd}. ipacheckscto using "Bontanga Baseline.xlsx", other(-666) dontknow(-999) refuse(-888){text}

{pstd}
Check a SurveyCTO XLS form and display results on Stata window. Include checks for other specify (-666), dontknow(-999){text}
and refuse to answer (-888). Export results to excel file "botanga_baseline_check.xlsx"

	{cmd}. ipacheckscto using "Bontanga Baseline.xlsx", other(-666) dontknow(-999) refuse(-888) outfile("botanga_baseline_check.xlsx"){text}


{marker author}{...}
{title:Author}

{pstd}Ishmail Azindoo Baako{p_end}

{pstd}For questions or suggestions, submit a
{browse "https://github.com/PovertyAction/ipacheckscto/issues":GitHub issue}
or e-mail researchsupport@poverty-action.org.{p_end}