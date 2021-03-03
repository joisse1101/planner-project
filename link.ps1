param (
    [int]$year = $(throw "-year is required."),
    [int]$mth = $(throw "-mth is required."),
    [int]$day = $(throw "-day is required.")
)

. ".\gen-functions.ps1"

$firsts = New-Object System.Collections.Generic.List[Object]
foreach ($p in $prefixes) {
    $idx = 1
    while ($idx -le $presentation.Slides.Count ) {
        $name = $presentation.Slides($idx).Name
        if ($name -match ("^" + $p) ) {
            $firsts.Add($presentation.Slides($idx))
            break
        }
        $idx++
    }
}


Write-Output "Working on years"
## Hyperlinking yr_$year slide
for ($i = 0; $i -le 1; $i++) {
    $slide = $presentation.Slides("yr_" + ($year+$i))
    $labels = 1,2,3,4
    ## hyperlinking months
    foreach ($m in $months) {
        try { $link = $presentation.Slides("mth_" + $m) }
        catch { Continue }
        $target = $slide.Shapes.Range( (("mini_" + $m + "_" + $year), ("mini_" + $m + "_" + $year + "_cal")) )
        Link $target $link
    }
    ## hyperlinking divs
    foreach ($l in $labels) {
        $target = $slide.Shapes("div_" + $divs[$l])
        Link $target $firsts[$l]
    }
}

Write-Output "Working on months"
## Hyperlinking mth_$months[$m] slides
$labels = 0,2,3,4

for ($mIdx = 1; $mIdx -le $months.Count; $mIdx++) {
    $m = $months[$mIdx - 1]
    try { $slide = $presentation.Slides("mth_" + $m) }
    catch { Continue }

    $numDays = [DateTime]::DaysInMonth($year, $mIdx)
    [int]$fDay, [int]$fMth, [int]$fYear = Get-fSunday 1 $mIdx $year

    ## hyperlink day_num
    for ($d = 1; $d -le $numDays; $d++) {
        if ( (Get-Date -Year $year -Month $mIdx -Day $d ).DayOfWeek.value__ -eq 0 ) {
            ## link to Sunday daily page
            $fDay, $fMth, $fYear = $d, $mIdx, $year
            $link = $presentation.Slides("day_" + $fDay + $months[$fMth - 1])
        }
        ## link to weekly of week
        else { $link = $presentation.Slides("wk_" + $fDay + $months[$fMth - 1]) }

        $target = $slide.Shapes("day_" + $d)
        Link $target $link
    }
    ## hyperlink mini_cals
    for ($i = 0; $i -lt 3; $i++) {
        try { $target = $slide.Shapes.Range( (("mini_" + $months[($mIdx + $i)] + "_" + $year), ("mini_" + $months[($mIdx + $i)] + "_" + $year + "_cal")) ) }
        catch { continue }
        try { $link = $presentation.Slides("mth_" + $months[($mIdx + $i)]) }
        catch { continue }        
        Link $target $link
    }
    ## hyperlinking divs
    foreach ($l in $labels) {
        $target = $slide.Shapes("div_" + $divs[$l])
        Link $target $firsts[$l]
    }
}

Write-Output "Working on weeks"
## Hyperlinking wk_$num$months[$m] slides
$d, $mIdx, $y = Get-fSunday $day $mth $year
$labels = 0, 1, 3, 4

while ($y -le $year) {
    try { $slide = $presentation.Slides("wk_" + $d + $months[$mIdx - 1]) }
    catch {
        $d, $mIdx, $y = AddDMY $d $mIdx $y 7
        continue
    }
    
    $numDays = [DateTime]::DaysInMonth($y, $mIdx)

    ## hyperlinking mini cal
    $x, $mCal, $yCal = AddDMY $d $mIdx $y 6
    if ($yCal -eq $year) {
        try { 
            $link = $presentation.Slides("mth_" + $months[$mCal - 1])
            $target = $slide.Shapes.Range( (("mini_" + $months[$mCal - 1] + "_" + $year), ("mini_" + $months[$mCal - 1] + "_" + $year + "_cal")) )
            Link $target $link
        }
        catch { continue }
    }
    ## hyperlinking days and dates
    for ($dayIdx = 0; $dayIdx -lt $days.Count; $dayIdx++) {
        $m = $months[$mIdx - 1]
        if ($y -eq $year) {
            if ($dayIdx -eq 0) {
                try { 
                    $link = $presentation.Slides("day_" + [string]($d) + $m)
                }
                catch { 
                    $link = $null
                }
                $target = $slide.Shapes.Range( (("day_" + $days[$dayIdx]), ("date_" + [string]$d + $m)) )
                Link $target $link
                $d, $mIdx, $y = AddDMY $d $mIdx $y 1
            }
            else {
                try { 
                    $link = $presentation.Slides("day_" + [string]($d) + $m)
                }
                catch { 
                    $link = $null
                }
                $target = $slide.Shapes.Range( (("day_" + $days[$dayIdx]), ("date_" + [string]$d + $m)) )
                Link $target $link
                $dayIdx++
                $d, $mIdx, $y = AddDMY $d $mIdx $y 1
                $target = $slide.Shapes.Range( (("day_" + $days[$dayIdx]), ("date_" + [string]$d + $months[$mIdx - 1])) )
                Link $target $link
                $d, $mIdx, $y = AddDMY $d $mIdx $y 1
            }          
        }
    }
    ## hyperlinking divs
    foreach ($l in $labels) {
        $target = $slide.Shapes("div_" + $divs[$l])
        Link $target $firsts[$l]
    }
}


Write-Output "Working on days"
$d, $mIdx, $y = Get-fDay $day $mth $year
$dayIdx = ( Get-Date -Year $y -Month $mIdx -Day $d ).DayOfWeek.value__
$labels = 0, 1, 2, 4

while ($y -le $year) {
    try { 
        $slide = $presentation.Slides("day_" + $d + $months[$mIdx - 1])
        
    }
    catch {
        if ($dayIdx -eq 0) {
            $d, $mIdx, $y = AddDMY $d $mIdx $y 1
            $dayIdx = 1
        }
        else {
            $d, $mIdx, $y = AddDMY $d $mIdx $y 2
            $dayIdx = $dayIdx + 2
        }
        continue
    }
    
    ## hyperlinking days and dates
    $num = $dayIdx % 2
    for ($i = 0; $i -le $num; $i++) {
        $lDay, $lMth, $lYear = Get-fSunday $d $mIdx $y
        $m = $months[$mIdx - 1]
        if ($lYear -eq $year) {
            try { $link = $presentation.Slides("wk_" + $lDay + $months[$lMth - 1]) }
            catch { $link = $null }
            $target = $slide.Shapes.Range( ( ("day_" + $d + "_" + $m), ("date_" + $d + "_" + $m), ("type_" + $d + "_" + $m) ) )
            Link $target $link
        }
        $d, $mIdx, $y = AddDMY $d $mIdx $y 1
        $dayIdx = ($dayIdx + 1) % 7
    }
    ## hyperlinking divs
    foreach ($l in $labels) {
        $target = $slide.Shapes("div_" + $divs[$l])
        Link $target $firsts[$l]
    }

}
