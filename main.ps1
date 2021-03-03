## Initialize planner
if (1) {
    $application = New-Object -ComObject powerpoint.application
    $ppt = $PSScriptRoot + "\planner.pptx"
    $presentation = $application.Presentations.Open($ppt)
    $templates = @{ templates = $presentation.Slides("templates_stickers"); }
}
else {
    $application = New-Object -ComObject powerpoint.application
    $presentation = $application.Presentations[$application.Presentations.Count]
    $templates = @{ templates = $presentation.Slides("templates_stickers"); }
}

## Import holidays.csv
$csv = $PSScriptRoot + "\holidays.csv"
$hols = Import-Csv $csv

## Default data
$cm = $presentation.SlideMaster.Shapes.Title.Left
$layouts = $presentation.SlideMaster.CustomLayouts
$days = "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"
$months = "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"
$dType = "Sermon", "Devotions"
$divs = "Yearly", "Monthly", "Weekly", "Daily", "Extras"
$prefixes = "yr_", "mth_", "wk_", "day_", "templates"
$del = New-Object System.Collections.Generic.List[Object]

## Change Y-M-D here:
[int]$year = 2021
[int]$mth = 3
[int]$day = 1

## Initialise miniCal templates
.\year.ps1 -year:$year -layers:1 -sName:"templates_miniCal_month" -sPos:($presentation.Slides.Count + 1) -type:"month"
$templates = $templates + @{ month = $presentation.Slides("templates_miniCal_month") }
.\year.ps1 -year:$year -layers:1 -sName:"templates_miniCal_week" -sPos:($presentation.Slides.Count + 1) -type:"week"
$templates = $templates + @{ week = $presentation.Slides("templates_miniCal_week") }

## Generate planner
.\year.ps1 -year:$year -sName:("yr_" + $year) -sPos:($presentation.Slides.Count - $templates.Count + 1) -type:"year" -div
.\year.ps1 -year:($year + 1) -sName:("yr_" + ($year + 1)) -sPos:($presentation.Slides.Count - $templates.Count + 1) -type:"year" -div
.\month.ps1 -year:$year -mth:$mth -sPos:($presentation.Slides.Count - $templates.Count + 1)
.\week.ps1 -year:$year -mth:$mth -day:$day -sPos:($presentation.Slides.Count - $templates.Count + 1)
.\days.ps1 -year:$year -mth:$mth -day:$day -sPos:($presentation.Slides.Count - $templates.Count + 1)



## Clear templates
$templates.month.delete()
$templates.Remove('month')
$templates.week.delete()
$templates.Remove('week')
while (1) {
    try { $templates.templates.Shapes[2].delete() }
    catch { break }
}
$labels = $divs[0], $divs[1], $divs[2], $divs[3]
Add-divs $template.templates, $presentation.SlideMaster.Shapes.Title, $labels

## Create hyperlinks between pages
.\link.ps1 -year:$year -mth:$mth -day:$day