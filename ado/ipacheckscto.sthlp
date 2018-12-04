{smcl}
{com}{sf}{ul off}{res}
{title:Title}


{col 5}ipacheckscto - Check SurveyCTO XLS form


{title:Syntax}

{col 5}ipacheckscto using {help filename:filename} [, {ul:out}file({help filename:filename}) replace]


{it:options}{col 25}{it:description}
{hline}

{bf:{ul:out}file}{col 25}Save output into excel file
{bf:{ul:replace}}{col 25}overwrite output file


{title:Options}
{p 5 10}outfile(filename) specifies the name of excel file to be saved. If an extension to the outfile is not specified .xlsx is assumed

{col 5}replace specifies that the outfile should be replaced if it already exist.



{title:Description}

{p 5 10}ipacheckscto is a user written stata command that checks your SurveyCTO xls form for issues and that may not be flagged by the SurveyCTO server. These ranges from small issues such as flagging non required fields to more serious problemslike using a repeat field outside the repeat group.{p_end}



{title:Examples}

{p 5}Check surveycto form without saving an excel output{p_end}
{p 10} ipacheckscto using "baseline_survey.xlsx

{p 5}Check surveycto form and save excel output{p_end}
{p 10} ipacheckscto using "baseline_survey.xlsx", outfile(baseline_survey_check.xlsx) replace


{title:Authors}

Ishmail Azindoo Baako (IPA) iabaako@poverty-action.org

Kindly email issues, suggestion and comments to researchsupport@poverty-action.org
{smcl}
{res}{sf}{ul off}