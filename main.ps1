if (1) {
    $application = New-Object -ComObject powerpoint.application
    $ppt = $PSScriptRoot + "\planner.pptx"
    $presentation = $application.Presentations.Open($ppt)
    $templates = @{ templates = $presentation.Slides("templates_stickers"); }

    $csv = $PSScriptRoot + "\holidays.csv"
    $hols = Import-Csv $csv
}

$cm = $presentation.SlideMaster.Shapes.Title.Left
$layouts = $presentation.SlideMaster.CustomLayouts

$days = "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"
$months = "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"
$dType = "Sermon", "Devotions"
[int]$year = 2021
[int]$mth = 3
[int]$day = 1

$currHols = New-Object System.Collections.Generic.List[Object]
$holIdx = 0
$del = New-Object System.Collections.Generic.List[Object]
$divs = "Yearly", "Monthly", "Weekly", "Daily", "Extras"
$prefixes = "yr_", "mth_", "wk_", "day_", "templates"

.\year.ps1 -year:$year -layers:1 -sName:"templates_miniCal_month" -sPos:($presentation.Slides.Count + 1) -type:"month"
$templates = $templates + @{ month = $presentation.Slides("templates_miniCal_month") }
.\year.ps1 -year:$year -layers:1 -sName:"templates_miniCal_week" -sPos:($presentation.Slides.Count + 1) -type:"week"
$templates = $templates + @{ week = $presentation.Slides("templates_miniCal_week") }

.\year.ps1 -year:$year -sName:("yr_" + $year) -sPos:($presentation.Slides.Count - $templates.Count + 1) -type:"year" -div
.\year.ps1 -year:($year+1) -sName:("yr_" + ($year+1)) -sPos:($presentation.Slides.Count - $templates.Count + 1) -type:"year" -div
.\month.ps1 -year:$year -mth:$mth -sPos:($presentation.Slides.Count - $templates.Count + 1)
.\week.ps1 -year:$year -mth:$mth -day:$day -sPos:($presentation.Slides.Count - $templates.Count + 1)
.\days.ps1 -year:$year -mth:$mth -day:$day -sPos:($presentation.Slides.Count - $templates.Count + 1)

.\link.ps1 -year:$year -mth:$mth -day:$day

$templates.month.delete()
$templates.week.delete()

while (1) {
    try { $templates.templates.Shapes[2].delete() }
    catch { break }
}