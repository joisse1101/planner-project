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
    $holIdx = 0
    $currHols = New-Object System.Collections.Generic.List[Object]
    foreach ($h in $hols) {
        [string]$hol = $h.Y + "-" + $h.M + "-" + $h.D
        $diff = (New-TimeSpan -Start $hol -End $date).Days
        if ($diff -le 0) {
            ## $hol starts on or after day
            return $holIdx, $currHols
        }
        else {
            $hTime = [int]$h.Days * [int]$h.Repeat
            if ($diff -lt $hTime) {
                ## add to currHols
                $days = [Math]::Floor( ($hTime - $diff) / $h.Repeat)
                $timer = ($hTime - $diff) - ($days * $h.Repeat) + 1
                if ($h.Name -eq "Lent") {
                    $lDay, $lMth, $lYear = Get-fSunday $h.D $h.M $h.Y
                    [string]$lDate = [string]$lYear + "-" + [string]$lMth + "-" + [string]$lDay
                    $numSundays = [Math]::Floor( (New-TimeSpan -Start $lDate -End $date).Days / 7 )
                    $days = $days + $numSundays
                }
                $currHols.Add( ($h.Name, $days, $timer, $h.Days, $h.Repeat) )
            }
        }
        $holIdx++
    }
    return $holIdx, $currHols
}

function Add-divs {
    param($slide, $title, $labels)
    $nTop = 0.35 * $cm
    $nLeft = 20 * $cm
    foreach ($l in $labels) {
        $box = $slide.Shapes.AddTextbox(5, $nLeft, $nTop, 0.75 * $cm, 7.25 * $cm) # vertical textbox
        use-DivText $box $title    
        $box.TextFrame.TextRange.Text = $l
        $box.Name = "div_" + $l
        $nTop = $nTop + $box.Height
    }
}
