{smcl}
{* *! version 1.1.0 Innovations for Poverty Action 17feb2020}{...}
{title:Title}

{phang}
{cmd:ipacheckscto} {hline 2} Check SurveyCTO XLS form


{marker syntax}{...}
{title:Syntax}


{col 5}ipacheckscto using {help filename:filename} [, {ul:out}file({help filename:filename}) {ul:oth}er({help int:integer}) replace]


{it:options}{col 25}{it:description}
{hline}

{bf:{ul:out}file}{col 25}Save output into excel file
{bf:{ul:oth}er}{col 25}Other specify value
{bf:{ul:replace}}{col 25}overwrite output file


{title:Options}
{p 5 10}outfile(filename) specifies the name of excel file to be saved. If an extension to the outfile is not specified .xlsx is assumed

{col 5}replace specifies that the outfile should be replaced if it already exist.



{title:Description}

{p 5 10}ipacheckscto is a user written stata command that checks your SurveyCTO xls form for issues and that may not be flagged by the SurveyCTO server. These ranges from small issues such as flagging non required fields to more serious problems like using a repeat field outside the repeat group.{p_end}



{title:Examples}

{p 5}Check surveycto form without saving an excel output{p_end}
{p 10} ipacheckscto using "baseline_survey.xlsx"

{p 5}Check surveycto form (including other specify fields){p_end}
{p 10} ipacheckscto using "baseline_survey.xlsx", other(-666) replace

{p 5}Check surveycto form and save excel output{p_end}
{p 10} ipacheckscto using "baseline_survey.xlsx", other(-666) outfile(baseline_survey_check.xlsx) replace


{title:Authors}

{p 5} Innovations for Poverty Action
{p 5} Ishmail Azindoo Baako (researchsupport@poverty-action.org)
{smcl}