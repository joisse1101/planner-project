## [void] Link $targetSlide $linkSlide
function Link {
    param($target, $link)
    $target.ActionSettings(1).Action = 10
    $target.ActionSettings(1).Hyperlink.SubAddress = [string]$link.SlideID + ", " + [string]$link.SlideIndex + ", " + $link.Name
}

## [[int]Day, [int]Mth, [int]Year] AddDMY $d $m $y $numDays
function AddDMY {
    param([int]$d, [int]$m, [int]$y, [int]$num)
    $YMD = (Get-Date -Year $y -Month $m -Day $d).AddDays($num).ToString("yyyyMMdd")
    $y = [int]$YMD.Substring(0, 4)
    $m = [int]$YMD.Substring(4, 2)
    $d = [int]$YMD.Substring(6, 2)
    return [int]$d, [int]$m, [int]$y
}

## [[int]Day, [int]Mth, [int]Year] Get-fSunday $d $m $y
function Get-fSunday {
    param([int]$d, [int]$m, [int]$y)
    while ( (Get-Date -Year $y -Month $m -Day $d ).DayOfWeek.value__ -ne 0) {
        $d, $m, $y = AddDMY $d $m $y -1
    }
    return [int]$d, [int]$m, [int]$y
}

## [[int]Day, [int]Mth, [int]Year] Get-fDay $d $m $y
function Get-fDay {
    param([int]$d, [int]$m, [int]$y)
    $day_idx = ( Get-Date -Year $y -Month $m -Day $d ).DayOfWeek.value__
    if ((($day_idx % 2) -eq 0) -and ($day_idx -ne 0)) { 
        $d, $m, $y = AddDMY $d $m $y -1
    }
    return [int]$d, [int]$m, [int]$y
}

## [[int]holIdx [List<Object>]currHols] Get-holIdx $d $m $y $hols
function Get-holIdx {
    param([int]$d, [int]$m, [int]$y, $hols)
    [string]$date = [string]$y + "-" + [string]$m + "-" + [string]$d

    $currHols = New-Object System.Collections.Generic.List[Object]
    foreach ($h in $hols) {
        [string]$hol = $h.Y + "-" + $h.M + "-" + $h.D
        $diff = (New-TimeSpan -Start $hol -End $date).Days
        if ($diff -le 0) {
            ## $hol starts on or after day
            return $holIdx, $currHols
        }
        else {
            $hTime = $h.Days * $h.Repeat
            if ($diff -gt $hTime) {
                ## add to currHols
                $days = [Math]::Floor( ($hTime - $diff) / $h.Repeat)
                $timer = ($hTime - $diff) - ($days * $h.Repeat) + 1
                $currHols.Add( ($h.Name, $days, $timer, $h.Repeat, $holIdx) )
            }
        }
        $holIdx++
    }
    return $holIdx, $currHols
}
